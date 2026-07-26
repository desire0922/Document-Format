# -*- coding: utf-8 -*-
"""语文试卷/假期作业一键排版工具（python-docx 版）
兼容目标：Windows 7/10/11 + Python 3.8+；不依赖 Microsoft Word，只处理 .docx。
"""
from __future__ import annotations

import json, os, re, sys, traceback
from dataclasses import dataclass, asdict
from datetime import date
from pathlib import Path
from typing import Optional, List, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

APP_NAME = "语文试卷一键排版"
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "word_exam_formatter_config.json"
PROFILE_DIR = APP_DIR / "填充配置方案"
OLD_CONFIG_FILE = Path.home() / ".word_exam_formatter.json"
OLD_PROFILE_DIR = Path.home() / ".word_exam_formatter_profiles"
DEFAULT_PROFILE_NAME = "自动保存"
ABOUT_TITLE = "关于语文试卷一键排版"
ABOUT_VERSION = "V8 python-docx"
ABOUT_AUTHOR = "校内教学排版工具"
ABOUT_TEXT = """语文试卷一键排版工具

项目地址：https://github.com/2SH33P/word_exam_formatter

众所周知每当(小)长假，老师们都会为假期作业而苦恼，区区出题尚且 easy，苦苦排版实属无聊，为了缓解老师的工作压力，我们成立了这个工作小组，用 20 天的超长暑假研发出可用 20 年的排版程序。它可以帮助老师包括但不限于创建文件，格式化输出内容，字体字号准确设置，页面页码精确排版，同时我们也会不断完善这一软件，不断优化老师们的使用体验，也期待更多小伙伴可以参加到程序完善中来。作为在校高中生，创作不易，希望满意的老师给予一些小小的支持。您的鼓励是我们不懈前进的动力。

""".strip()
NOTICE_TITLE = "使用前注意"
NOTICE_TEXT = """第一次使用一定要先配置好最终文件储存在哪里，一定不要忘记！！

当前 python-docx 版本只处理 .docx 文件；老式 .doc 文件请先用 Word 或 WPS 另存为 .docx。

批量处理前，请确认文件队列顺序正确。编号会严格按照界面中的最终顺序生成。

如果原文包含复杂公式、文本框、分栏、特殊页眉页脚或大量图片，处理后请人工检查一次。

默认会生成“原文件名_已排版.docx”，不会覆盖原稿。
""".strip()


def enable_windows_dpi_awareness():
    if os.name != "nt": return
    try:
        import ctypes
        for call in (
            lambda: ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)),
            lambda: ctypes.windll.shcore.SetProcessDpiAwareness(2),
            lambda: ctypes.windll.user32.SetProcessDPIAware(),
        ):
            try: call(); return
            except Exception: pass
    except Exception: pass

CHINESE_DIGITS = {"零":0,"〇":0,"一":1,"二":2,"两":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9}
CHINESE_UNITS = {"十":10,"百":100,"千":1000,"万":10000}

def chinese_number_to_int(value):
    value = value.strip()
    if value.startswith("第"): value = value[1:]
    if not value: return None
    if all(ch in CHINESE_DIGITS for ch in value):
        try: return int("".join(str(CHINESE_DIGITS[ch]) for ch in value))
        except ValueError: return None
    if any(ch not in CHINESE_DIGITS and ch not in CHINESE_UNITS for ch in value): return None
    total = section = number = 0
    for ch in value:
        if ch in CHINESE_DIGITS:
            number = CHINESE_DIGITS[ch]; continue
        unit = CHINESE_UNITS[ch]
        if unit == 10000:
            section = (section + number) * unit; total += section; section = number = 0
        else:
            if number == 0: number = 1
            section += number * unit; number = 0
    return total + section + number

def int_to_chinese_number(number):
    if number < 0: raise ValueError("中文序号暂不支持负数")
    if number == 0: return "零"
    if number > 99999999: raise ValueError("中文序号目前最多支持到 99999999")
    digits, units = "零一二三四五六七八九", ["","十","百","千"]
    def four(n):
        out, zero = [], False
        for pos in range(3, -1, -1):
            unit = 10 ** pos; d = n // unit; n %= unit
            if d == 0:
                if out and n > 0: zero = True
                continue
            if zero: out.append("零"); zero = False
            if not (d == 1 and pos == 1 and not out): out.append(digits[d])
            out.append(units[pos])
        return "".join(out)
    high, low = number // 10000, number % 10000
    parts = []
    if high:
        parts += [four(high), "万"]
        if 0 < low < 1000: parts.append("零")
    if low: parts.append(four(low))
    return "".join(parts)

def parse_sequence_start(value):
    raw = str(value).strip()
    if not raw: raise ValueError("请填写批量编号的首位数，例如 1、001 或 一")
    if re.fullmatch(r"\d+", raw): return int(raw), "arabic", len(raw)
    if re.fullmatch(r"[零〇一二两三四五六七八九十百千万]+", raw):
        parsed = chinese_number_to_int(raw)
        if parsed is None: raise ValueError("无法识别首位数“{}”。".format(raw))
        return parsed, "chinese", 1
    raise ValueError("首位数“{}”格式不支持。只支持阿拉伯数字（如 1、001）或汉字数字（如 一、十一）。".format(raw))

def extract_name_order(stem):
    cleaned = stem.strip()
    for pattern in [r"^\s*[（(【\[]?\s*(\d{1,6})\s*[）)】\]]?\s*[-—_、.．·：: ]+\s*(.+?)\s*$", r"^\s*[（(【\[]\s*(\d{1,6})\s*[）)】\]]\s*(.+?)\s*$", r"^\s*(\d{1,6})\s+(.+?)\s*$"]:
        m = re.match(pattern, cleaned)
        if m: return int(m.group(1)), m.group(2).strip(" -—_、.．·：:") or cleaned
    chinese_num = r"[零〇一二两三四五六七八九十百千万]+"
    for pattern in [rf"^\s*[（(【\[]?\s*({chinese_num})\s*[）)】\]]?\s*[-—_、.．·：: ]+\s*(.+?)\s*$", rf"^\s*第\s*({chinese_num})\s*[章节单元课天套卷]\s*[-—_、.．·：: ]*\s*(.+?)\s*$", rf"^\s*第\s*({chinese_num})\s*[章节单元课天套卷]\s*$"]:
        m = re.match(pattern, cleaned)
        if m:
            order = chinese_number_to_int(m.group(1))
            if order is None: continue
            if m.lastindex and m.lastindex >= 2: return order, m.group(2).strip(" -—_、.．·：:") or cleaned
            return order, cleaned
    m = re.match(rf"^\s*第\s*({chinese_num})\s*[章节单元课天套卷]\s*(.+?)\s*$", cleaned)
    if m:
        order = chinese_number_to_int(m.group(1))
        if order is not None: return order, m.group(2).strip(" -—_、.．·：:") or cleaned
    m = re.fullmatch(r"\s*(\d{1,6})\s*", cleaned)
    if m: return int(m.group(1)), cleaned
    if re.fullmatch(chinese_num, cleaned): return chinese_number_to_int(cleaned), cleaned
    return None, cleaned

def natural_sort_key(value):
    key = []
    for part in re.split(r"(\d+|[零〇一二两三四五六七八九十百千万]+)", value):
        if not part: continue
        if part.isdigit(): key.append((0, int(part)))
        else:
            cn = chinese_number_to_int(part)
            key.append((0, cn) if cn is not None else (1, part.casefold()))
    return key

def split_leading_order(stem): return extract_name_order(stem)

@dataclass
class BatchTask:
    source: Path
    detected_order: Optional[int] = None

@dataclass
class FormatSettings:
    title_line_small: str = "廉江市实验学校 2025—2026学年第二学期"
    title_line_big: str = "高一语文 暑假作业（{1}）卷"
    setter: str = "xxx"
    reviewer: str = "xxx"
    use_date: str = "2026年7月xx日"
    insert_mode: str = "start_and_markers"
    marker: str = "[[题头]]"
    page_break_before_repeated_title: bool = False
    sequence_rules: Optional[List[dict]] = None
    top_margin_cm: float = 2.3
    bottom_margin_cm: float = 2.3
    left_margin_cm: float = 2.1
    right_margin_cm: float = 2.1
    header_distance_cm: float = 1.2
    footer_distance_cm: float = 1.2
    body_font: str = "宋体"
    body_size: float = 12.0
    heading_font: str = "宋体"
    heading_size: float = 12.0
    title_font: str = "宋体"
    title_size: float = 16.0
    small_title_font: str = "宋体"
    small_title_size: float = 12.0
    footer_font: str = "宋体"
    footer_size: float = 9.0
    line_spacing_multiple: float = 1.5
    first_line_indent_chars: float = 0.0
    paragraph_after_pt: float = 0.0
    apply_heading_format: bool = True
    heading_patterns: str = r"^[一二三四五六七八九十]+、|^\d+[、.]|^第[一二三四五六七八九十\d]+[章节单元天]|^练习[一二三四五六七八九十\d]+"
    add_footer_page_number: bool = True
    footer_subject_text: str = "语文"
    overwrite: bool = False

class WordFormatter:
    def __init__(self, settings, log_callback):
        self.s, self.log, self.current_sequences = settings, log_callback, {}
    def format_sequence_value(self, rule, number):
        return int_to_chinese_number(number) if rule["style"] == "chinese" else ("{:0" + str(max(1, int(rule["width"]))) + "d}").format(number)
    def render_sequence_text(self, value):
        value = str(value); rules = self.s.sequence_rules or []
        for index, rule in enumerate(rules, 1):
            current = self.current_sequences.get(index, rule["start_value"]); rendered = self.format_sequence_value(rule, current)
            value = value.replace("{" + str(index) + "}", rendered).replace("{编号" + str(index) + "}", rendered)
        if rules and "{}" in value:
            first = rules[0]; current = self.current_sequences.get(1, first["start_value"])
            value = value.replace("{}", self.format_sequence_value(first, current))
        return value
    def title_lines(self):
        return (self.render_sequence_text(self.s.title_line_small), self.render_sequence_text(self.s.title_line_big), "出题人：{}    审题人：{}    使用日期：{}".format(self.render_sequence_text(self.s.setter), self.render_sequence_text(self.s.reviewer), self.render_sequence_text(self.s.use_date)))
    def set_run_font(self, run, name, size, bold=None):
        run.font.name = name; run.font.size = Pt(float(size))
        if bold is not None: run.bold = bool(bold)
        try: run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
        except Exception: pass
    def set_paragraph_text(self, p, text, font, size, bold, align=None):
        p.clear(); r = p.add_run(text); self.set_run_font(r, font, size, bold)
        if align is not None: p.alignment = align
        p.paragraph_format.first_line_indent = Pt(0); p.paragraph_format.left_indent = Pt(0); p.paragraph_format.right_indent = Pt(0)
        return p
    def configure_sections(self, doc):
        for sec in doc.sections:
            was_landscape = sec.page_width > sec.page_height
            sec.top_margin = Cm(self.s.top_margin_cm); sec.bottom_margin = Cm(self.s.bottom_margin_cm); sec.left_margin = Cm(self.s.left_margin_cm); sec.right_margin = Cm(self.s.right_margin_cm)
            sec.header_distance = Cm(self.s.header_distance_cm); sec.footer_distance = Cm(self.s.footer_distance_cm)
            if not was_landscape:
                sec.orientation = WD_ORIENT.PORTRAIT; sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    def format_body(self, doc):
        for p in doc.paragraphs:
            pf = p.paragraph_format; pf.line_spacing = float(self.s.line_spacing_multiple); pf.space_after = Pt(float(self.s.paragraph_after_pt)); pf.space_before = Pt(0); pf.first_line_indent = Pt(float(self.s.first_line_indent_chars) * float(self.s.body_size))
            for r in p.runs: self.set_run_font(r, self.s.body_font, self.s.body_size, False)
    def format_headings(self, doc):
        if not self.s.apply_heading_format: return
        try: regex = re.compile(self.s.heading_patterns)
        except re.error as exc: raise RuntimeError("大题标题识别表达式有误：{}".format(exc))
        for p in doc.paragraphs:
            if p.text.strip() and regex.search(p.text.strip()):
                for r in p.runs: self.set_run_font(r, self.s.heading_font, self.s.heading_size, True)
    def _insert_title_before_paragraph(self, p, add_page_break=False):
        if add_page_break: p.insert_paragraph_before("").add_run().add_break()
        line1, line2, line3 = self.title_lines(); p1 = p.insert_paragraph_before(line1); p2 = p.insert_paragraph_before(line2); p3 = p.insert_paragraph_before(line3)
        self.set_paragraph_text(p1, line1, self.s.title_font, self.s.title_size, True, WD_ALIGN_PARAGRAPH.CENTER)
        self.set_paragraph_text(p2, line2, self.s.title_font, self.s.title_size, True, WD_ALIGN_PARAGRAPH.CENTER)
        self.set_paragraph_text(p3, line3, self.s.small_title_font, self.s.small_title_size, False, WD_ALIGN_PARAGRAPH.CENTER)
        p1.paragraph_format.space_after = Pt(2); p2.paragraph_format.space_after = Pt(3); p3.paragraph_format.space_after = Pt(6)
    def _remove_paragraph(self, p):
        el = p._element; el.getparent().remove(el); p._p = p._element = None
    def insert_title_at_start(self, doc):
        line1, _, _ = self.title_lines(); preview = "\n".join(p.text for p in doc.paragraphs[:8])
        if line1 and line1 in preview: self.log("  文档开头已检测到相同题头，跳过重复插入。"); return
        if doc.paragraphs: self._insert_title_before_paragraph(doc.paragraphs[0], False)
        else:
            p = doc.add_paragraph(""); self._insert_title_before_paragraph(p, False); self._remove_paragraph(p)
    def replace_markers_with_titles(self, doc):
        marker = self.s.marker.strip(); count = 0
        if not marker: return 0
        for p in list(doc.paragraphs):
            if marker not in p.text: continue
            if p.text.strip() == marker:
                self._insert_title_before_paragraph(p, self.s.page_break_before_repeated_title); self._remove_paragraph(p)
            else:
                self.set_paragraph_text(p, p.text.replace(marker, "\n".join(self.title_lines())), self.s.body_font, self.s.body_size, False)
            count += 1
        return count
    def add_field(self, p, instruction):
        run = p.add_run(); begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin"); instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " {} ".format(instruction); sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate"); text = OxmlElement("w:t"); text.text = "1"; end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
        for el in (begin, instr, sep, text, end): run._r.append(el)
        self.set_run_font(run, self.s.footer_font, self.s.footer_size, False)
    def add_footer(self, doc):
        if not self.s.add_footer_page_number: return
        for sec in doc.sections:
            p = sec.footer.paragraphs[0] if sec.footer.paragraphs else sec.footer.add_paragraph(); p.clear(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for txt, field in [(self.s.footer_subject_text + " 第 ", None), (None, "PAGE"), (" 页 共 ", None), (None, "NUMPAGES"), (" 页", None)]:
                if field: self.add_field(p, field)
                else: self.set_run_font(p.add_run(txt), self.s.footer_font, self.s.footer_size, False)
    def remove_excess_blank_paragraphs(self, doc):
        blank = 0
        for p in list(doc.paragraphs):
            if p.text.strip(): blank = 0; continue
            blank += 1
            if blank >= 3: self._remove_paragraph(p)
    def output_path_for(self, src, out_dir):
        if self.s.overwrite: return src
        out = out_dir / (src.stem + "_已排版.docx"); n = 2
        while out.exists(): out = out_dir / (src.stem + "_已排版_{}.docx".format(n)); n += 1
        return out
    def process_one(self, task, out_dir, queue_index):
        src = task.source
        if src.suffix.lower() != ".docx": self.log("跳过：{} 不是 .docx，请先另存为 .docx。".format(src.name)); return False, "not docx"
        self.current_sequences, parts = {}, []
        for index, rule in enumerate(self.s.sequence_rules or [], 1):
            val = int(rule["start_value"]) + (queue_index - 1); self.current_sequences[index] = val; parts.append("编号{}={}".format(index, self.format_sequence_value(rule, val)))
        self.log("正在处理：{}  → {}".format(src.name, "；".join(parts) if parts else "未设置编号"))
        try:
            doc = Document(str(src.resolve())); self.configure_sections(doc); self.format_body(doc)
            if self.s.insert_mode in ("start", "start_and_markers"): self.insert_title_at_start(doc)
            if self.s.insert_mode in ("markers", "start_and_markers"):
                self.log("  已替换题头标记：{} 处".format(self.replace_markers_with_titles(doc)))
            self.format_headings(doc); self.add_footer(doc); self.remove_excess_blank_paragraphs(doc); out_dir.mkdir(parents=True, exist_ok=True); out = self.output_path_for(src, out_dir); doc.save(str(out.resolve())); self.log("完成：{}".format(out.name)); return True, str(out)
        except Exception as exc:
            self.log("失败：{}：{}".format(src.name, exc)); self.log(traceback.format_exc()); return False, str(exc)
    def process(self, tasks, out_dir):
        ok = failed = 0
        for i, task in enumerate(tasks, 1):
            success, _ = self.process_one(task, out_dir, i)
            if success: ok += 1
            else: failed += 1
        return ok, failed


class ToolTip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self.show); widget.bind("<Leave>", self.hide)
    def show(self, event=None):
        if self.tip or not self.text: return
        x, y = self.widget.winfo_rootx() + 18, self.widget.winfo_rooty() + 18
        self.tip = tk.Toplevel(self.widget); self.tip.wm_overrideredirect(True); self.tip.wm_geometry("+{}+{}".format(x, y))
        frame = tk.Frame(self.tip, background="#ffffff", borderwidth=1, relief="solid"); frame.pack(fill="both", expand=True)
        tk.Label(frame, text=self.text, justify="left", background="#ffffff", foreground="#333333", padx=8, pady=5, wraplength=300, font=("TkDefaultFont", 9)).pack()
    def hide(self, event=None):
        if self.tip: self.tip.destroy(); self.tip = None

FIELD_HELP = {
    "title_line_small":"题头第一行，通常写学校、学年学期等信息，居中、加粗、三号显示。支持 {1}、{2} 编号。",
    "title_line_big":"题头第二行，通常写年级学科、作业/试卷名称和卷号，字号较大并加粗。支持 {1}、{2} 编号。",
    "setter":"第三行小字中的出题人，可留 xxx，也支持编号占位符。",
    "reviewer":"第三行小字中的审题人，可留 xxx，也支持编号占位符。",
    "use_date":"第三行小字中的使用日期，例如 2026年7月{1}日。",
    "marker":"文档中单独放一行这个标记时，程序会把它替换成完整题头。",
    "insert_mode":"控制题头插入位置：文档开头、替换标记、两者都做，或不加题头。",
    "top_margin_cm":"页面上边距，单位厘米。", "bottom_margin_cm":"页面下边距，单位厘米。", "left_margin_cm":"页面左边距，单位厘米。", "right_margin_cm":"页面右边距，单位厘米。",
    "header_distance_cm":"页眉距离页面边界的距离，单位厘米。", "footer_distance_cm":"页脚距离页面边界的距离，单位厘米。", "footer_subject_text":"页脚开头文字，例如“语文 第 X 页 共 Y 页”中的“语文”。",
    "body_font":"正文中文字体。", "body_size":"正文字号，单位磅。小四通常是 12。", "heading_font":"自动识别到的大题标题字体。", "heading_size":"自动识别到的大题标题字号，单位磅。",
    "title_font":"第二行大字题头字体。", "title_size":"第二行大字题头字号，单位磅。", "small_title_font":"第一行和第三行小字题头字体。", "small_title_size":"第一行和第三行小字题头字号，单位磅。",
    "footer_font":"页脚字体。", "footer_size":"页脚字号，单位磅。", "line_spacing_multiple":"正文行距倍数，例如 1.5。", "first_line_indent_chars":"正文首行缩进字符数，0 表示不缩进。", "paragraph_after_pt":"正文段后间距，单位磅。",
}

class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title(APP_NAME); self.geometry("980x780"); self.minsize(640, 520)
        self.batch_tasks, self.vars, self.sequence_rule_vars = [], {}, []
        self.input_path, self.output_path = tk.StringVar(), tk.StringVar()
        self.active_insert_widget, self.last_notice_date = None, ""
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        self._autosave_job = None; self._autosave_ready = False
        self._build_ui(); self._load_config(); self.refresh_profile_list(); self.setup_autosave(); self._autosave_ready = True

    def migrate_old_config_files(self):
        try:
            if OLD_CONFIG_FILE.exists() and not CONFIG_FILE.exists():
                CONFIG_FILE.write_text(OLD_CONFIG_FILE.read_text(encoding="utf-8"), encoding="utf-8")
            if OLD_PROFILE_DIR.exists():
                PROFILE_DIR.mkdir(parents=True, exist_ok=True)
                for old in OLD_PROFILE_DIR.glob("*.json"):
                    new = PROFILE_DIR / old.name
                    if not new.exists():
                        new.write_text(old.read_text(encoding="utf-8"), encoding="utf-8")
        except Exception:
            pass
    def _var(self, name, value="", kind="str"):
        var = tk.BooleanVar(value=bool(value)) if kind == "bool" else tk.StringVar(value=str(value)); self.vars[name] = var; return var
    def _on_mousewheel(self, event):
        try:
            units = int(-1 * (event.delta / 120))
            if event.state & 0x0001:
                self.canvas.xview_scroll(units, "units")
            else:
                self.canvas.yview_scroll(units, "units")
        except Exception:
            pass
    def _update_scroll_region(self, event=None):
        try:
            req_width = self.scroll_content.winfo_reqwidth()
            canvas_width = self.canvas.winfo_width()
            self.canvas.itemconfigure(self.canvas_window, width=max(req_width, canvas_width))
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass
    def _entry(self, parent, textvariable, width=None):
        entry = ttk.Entry(parent, textvariable=textvariable, width=width)
        entry.bind("<FocusIn>", lambda e: self._remember_insert_widget(e.widget)); entry.bind("<ButtonRelease-1>", lambda e: self._remember_insert_widget(e.widget)); entry.bind("<KeyRelease>", lambda e: self._remember_insert_widget(e.widget))
        return entry
    def _remember_insert_widget(self, widget): self.active_insert_widget = widget
    def _labeled_entry(self, parent, row, label, name, default, width=30):
        q = tk.Label(parent, text="?", width=2, cursor="question_arrow", fg="#2563eb", font=("TkDefaultFont", 8, "bold")); q.grid(row=row, column=0, sticky="e", padx=(3,0), pady=4); ToolTip(q, FIELD_HELP.get(name, "填写“{}”对应内容。".format(label)))
        ttk.Label(parent, text=label).grid(row=row, column=1, sticky="e", padx=5, pady=4)
        entry = self._entry(parent, self._var(name, default), width=width); entry.grid(row=row, column=2, sticky="ew", padx=5, pady=4); return entry

    def _build_ui(self):
        outer = ttk.Frame(self); outer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(outer, highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        hscroll = ttk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        vscroll.pack(side="right", fill="y")
        hscroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        root = ttk.Frame(self.canvas, padding=10)
        self.scroll_content = root
        self.canvas_window = self.canvas.create_window((0,0), window=root, anchor="nw")
        root.bind("<Configure>", self._update_scroll_region)
        self.canvas.bind("<Configure>", self._update_scroll_region)
        self.bind_all("<MouseWheel>", self._on_mousewheel)
        top = ttk.Frame(root); top.pack(fill="x", pady=(0,2)); about = tk.Label(top, text="关于", cursor="hand2", fg="#666", font=("TkDefaultFont",8,"underline")); about.pack(side="left"); about.bind("<Button-1>", lambda e: self.show_about())

        file_box = ttk.LabelFrame(root, text="一、批量文件队列（表格顺序就是处理顺序）", padding=8); file_box.pack(fill="both")
        toolbar = ttk.Frame(file_box); toolbar.pack(fill="x")
        for text, cmd, padx in [("添加 DOCX 文件", self.pick_file, 0), ("添加整个文件夹", self.pick_folder, 4), ("按文件名重新排序", self.sort_tasks, 4)]: ttk.Button(toolbar, text=text, command=cmd).pack(side="left", padx=padx)
        ttk.Button(toolbar, text="上移", command=lambda:self.move_task(-1)).pack(side="left", padx=(12,4)); ttk.Button(toolbar, text="下移", command=lambda:self.move_task(1)).pack(side="left", padx=4); ttk.Button(toolbar, text="删除选中", command=self.remove_selected_tasks).pack(side="left", padx=4); ttk.Button(toolbar, text="清空", command=self.clear_tasks).pack(side="left", padx=4); ttk.Button(toolbar, text="▶ 开始批量处理", command=self.run_format).pack(side="right", padx=(12,0))
        qf = ttk.Frame(file_box); qf.pack(fill="both", expand=True, pady=6); self.queue_tree = ttk.Treeview(qf, columns=("order","file"), show="headings", height=6, selectmode="extended"); self.queue_tree.heading("order", text="处理顺序"); self.queue_tree.heading("file", text="原文件名"); self.queue_tree.column("order", width=75, anchor="center", stretch=False); self.queue_tree.column("file", width=520, stretch=True); self.queue_tree.pack(side="left", fill="both", expand=True); qs = ttk.Scrollbar(qf, orient="vertical", command=self.queue_tree.yview); qs.pack(side="right", fill="y"); self.queue_tree.configure(yscrollcommand=qs.set)
        out = ttk.Frame(file_box); out.pack(fill="x"); ttk.Label(out, text="输出目录：").pack(side="left"); self._entry(out, self.output_path).pack(side="left", fill="x", expand=True, padx=6); ttk.Button(out, text="选择输出目录", command=self.pick_output).pack(side="left")

        profile_box = ttk.LabelFrame(root, text="二、填充配置方案（多个 JSON，可一键切换）", padding=8); profile_box.pack(fill="x", pady=(8,0))
        ttk.Label(profile_box, text="当前方案：").pack(side="left"); self.profile_var = tk.StringVar(value=DEFAULT_PROFILE_NAME); self.profile_combo = ttk.Combobox(profile_box, textvariable=self.profile_var, state="readonly", width=18); self.profile_combo.pack(side="left", padx=4); self.profile_combo.bind("<<ComboboxSelected>>", lambda e: self.load_selected_profile())
        ttk.Button(profile_box, text="应用", command=self.load_selected_profile).pack(side="left", padx=2); ttk.Button(profile_box, text="保存到当前方案", command=self.save_selected_profile).pack(side="left", padx=2); ttk.Button(profile_box, text="另存为新方案", command=self.save_profile_as).pack(side="left", padx=2); ttk.Button(profile_box, text="删除方案", command=self.delete_selected_profile).pack(side="left", padx=2)

        notebook = ttk.Notebook(root); notebook.pack(fill="both", expand=True, pady=8)
        basic, page, style, advanced = ttk.Frame(notebook, padding=10), ttk.Frame(notebook, padding=10), ttk.Frame(notebook, padding=10), ttk.Frame(notebook, padding=10)
        notebook.add(basic, text="题头内容"); notebook.add(page, text="页面设置"); notebook.add(style, text="字体段落"); notebook.add(advanced, text="高级选项")
        self._build_basic_tab(basic); self._build_page_tab(page); self._build_style_tab(style); self._build_advanced_tab(advanced)
        btns = ttk.Frame(root); btns.pack(fill="x"); ttk.Button(btns, text="保存当前预设", command=self.save_config).pack(side="left"); ttk.Button(btns, text="恢复推荐值", command=self.reset_defaults).pack(side="left", padx=6); self.process_button = ttk.Button(btns, text="▶ 开始批量处理", command=self.run_format); self.process_button.pack(side="right"); ttk.Label(btns, text="快捷键：Ctrl+Enter").pack(side="right", padx=(0,10)); self.bind_all("<Control-Return>", lambda e:self.run_format())
        log_box = ttk.LabelFrame(root, text="处理日志", padding=6); log_box.pack(fill="both", expand=True, pady=(8,0)); self.log_text = tk.Text(log_box, height=6, wrap="word"); self.log_text.pack(fill="both", expand=True)

    def _build_basic_tab(self, f):
        self.title_entries = {}
        fields = [("第一行题头","title_line_small","廉江市实验学校 2025—2026学年第二学期"),("第二行大字","title_line_big","高一语文 暑假作业（{1}）卷"),("出题人","setter","xxx"),("审题人","reviewer","xxx"),("使用日期","use_date","2026年7月xx日"),("重复题头标记","marker","[[题头]]")]
        for row, item in enumerate(fields): self.title_entries[item[1]] = self._labeled_entry(f, row, *item)
        f.columnconfigure(2, weight=1); q = tk.Label(f, text="?", width=2, cursor="question_arrow", fg="#2563eb", font=("TkDefaultFont",8,"bold")); q.grid(row=6,column=0,sticky="e",padx=(3,0),pady=4); ToolTip(q, FIELD_HELP["insert_mode"]); ttk.Label(f,text="题头插入方式").grid(row=6,column=1,sticky="e",padx=5,pady=4); ttk.Combobox(f,textvariable=self._var("insert_mode","start_and_markers"),values=["start","markers","start_and_markers","none"],state="readonly").grid(row=6,column=2,sticky="w",padx=5,pady=4); ttk.Label(f,text="start=仅文档开头；markers=仅替换标记；start_and_markers=两者都做；none=不加题头").grid(row=7,column=0,columnspan=3,sticky="w",padx=5)
        seq_box = ttk.LabelFrame(f, text="多组独立编号", padding=6); seq_box.grid(row=8,column=0,columnspan=3,sticky="ew",padx=5,pady=(10,3)); seq_box.columnconfigure(0, weight=1); self.sequence_rules_frame = ttk.Frame(seq_box); self.sequence_rules_frame.grid(row=0,column=0,sticky="ew"); sb = ttk.Frame(seq_box); sb.grid(row=1,column=0,sticky="w",pady=(6,0)); ttk.Button(sb,text="添加编号",command=self.add_sequence_rule_row).pack(side="left"); ttk.Button(sb,text="删除最后一组",command=self.remove_sequence_rule_row).pack(side="left",padx=6); ttk.Label(sb,text="点击右侧 {1}/{2} 可插入到当前光标处").pack(side="left",padx=10); self.add_sequence_rule_row("1"); self.add_sequence_rule_row("001"); ttk.Label(f,text="在题头任意位置写 {1}、{2}、{3}……；每组按自己的首位数分别递增。{} 仍等同于 {1}。").grid(row=9,column=0,columnspan=3,sticky="w",padx=5,pady=(6,0))

    def add_sequence_rule_row(self, default_value="1"):
        index = len(self.sequence_rule_vars) + 1; row = ttk.Frame(self.sequence_rules_frame); row.pack(fill="x", pady=2); ttk.Label(row,text="编号{}".format(index),width=8).pack(side="left"); ttk.Label(row,text="首位数").pack(side="left",padx=(4,2)); var = tk.StringVar(value=str(default_value)); self._entry(row,var,width=14).pack(side="left"); ttk.Button(row,text="{{{}}}".format(index),width=6,command=lambda i=index:self.insert_placeholder("{{{}}}".format(i))).pack(side="left",padx=(8,2)); ttk.Label(row,text="支持 1、01、001 或 一、十、十一").pack(side="left",padx=(8,0)); self.sequence_rule_vars.append({"frame":row,"var":var})
        if getattr(self, "_autosave_ready", False):
            try: var.trace_add("write", lambda *args: self.schedule_auto_save())
            except Exception: pass
    def remove_sequence_rule_row(self):
        if len(self.sequence_rule_vars) <= 1: messagebox.showwarning("不能删除","至少需要保留一组编号。"); return
        self.sequence_rule_vars.pop()["frame"].destroy()
    def rebuild_sequence_rules(self, values):
        for item in self.sequence_rule_vars: item["frame"].destroy()
        self.sequence_rule_vars = []
        for value in (values or ["1"]): self.add_sequence_rule_row(value)
    def insert_placeholder(self, placeholder):
        widget = self.active_insert_widget or self.title_entries.get("title_line_big")
        try:
            pos = widget.index(tk.INSERT); widget.insert(pos, placeholder); widget.focus_set(); widget.icursor(pos + len(placeholder)); self.active_insert_widget = widget
        except Exception:
            entry = self.title_entries.get("title_line_big"); entry.insert(tk.END, placeholder); entry.focus_set(); entry.icursor(tk.END); self.active_insert_widget = entry

    def _build_page_tab(self, f):
        fields = [("上边距（厘米）","top_margin_cm","2.3"),("下边距（厘米）","bottom_margin_cm","2.3"),("左边距（厘米）","left_margin_cm","2.1"),("右边距（厘米）","right_margin_cm","2.1"),("页眉距边界（厘米）","header_distance_cm","1.2"),("页脚距边界（厘米）","footer_distance_cm","1.2"),("页脚学科文字","footer_subject_text","语文")]
        for row, item in enumerate(fields): self._labeled_entry(f, row, *item)
        f.columnconfigure(2, weight=1); ttk.Checkbutton(f,text="添加页脚页码：语文 第 X 页 共 Y 页",variable=self._var("add_footer_page_number",True,"bool")).grid(row=7,column=0,columnspan=3,sticky="w",padx=5,pady=6)
    def _build_style_tab(self, f):
        fields = [("正文字体","body_font","宋体"),("正文字号（磅）","body_size","12"),("大题标题字体","heading_font","宋体"),("大题标题字号（磅）","heading_size","12"),("第二行大字字体","title_font","宋体"),("第二行大字字号（磅）","title_size","16"),("小字题头字体","small_title_font","宋体"),("小字题头字号（磅）","small_title_size","12"),("页脚字体","footer_font","宋体"),("页脚字号（磅）","footer_size","9"),("行距倍数","line_spacing_multiple","1.5"),("首行缩进字符数","first_line_indent_chars","0"),("段后间距（磅）","paragraph_after_pt","0")]
        for row, item in enumerate(fields): self._labeled_entry(f, row, *item)
        f.columnconfigure(2, weight=1)
    def _build_advanced_tab(self, f):
        ttk.Checkbutton(f,text="自动识别并加粗大题标题",variable=self._var("apply_heading_format",True,"bool")).grid(row=0,column=0,sticky="w",pady=4)
        ttk.Label(f,text="大题标题识别表达式：").grid(row=1,column=0,sticky="nw",pady=4); self.pattern_text = tk.Text(f,height=5,wrap="word"); self.pattern_text.grid(row=2,column=0,sticky="nsew"); self.pattern_text.insert("1.0", FormatSettings.heading_patterns)
        ttk.Checkbutton(f,text="每个重复题头前插入分页符",variable=self._var("page_break_before_repeated_title",False,"bool")).grid(row=3,column=0,sticky="w",pady=4)
        ttk.Checkbutton(f,text="覆盖原文件（不推荐）",variable=self._var("overwrite",False,"bool")).grid(row=4,column=0,sticky="w",pady=4)
        ttk.Label(f,text="建议：在每一份假期作业开始位置单独放一行 [[题头]]，程序会将其替换成完整题头。纯 python-docx 版不导出 PDF。").grid(row=5,column=0,sticky="w",pady=10)
        f.rowconfigure(2, weight=1); f.columnconfigure(0, weight=1)

    def show_notice_dialog(self):
        dialog = tk.Toplevel(self); dialog.title(NOTICE_TITLE); dialog.resizable(False, False); dialog.transient(self); result = tk.BooleanVar(value=False)
        body = ttk.Frame(dialog, padding=18); body.pack(fill="both", expand=True); ttk.Label(body,text=NOTICE_TITLE,font=("TkDefaultFont",14,"bold")).pack(anchor="w")
        msg = tk.Text(body,width=62,height=15,wrap="word",relief="flat",borderwidth=0,background=dialog.cget("background")); msg.insert("1.0", NOTICE_TEXT); msg.configure(state="disabled"); msg.pack(fill="both",expand=True,pady=(10,12))
        btns = ttk.Frame(body); btns.pack(fill="x"); ttk.Button(btns,text="取消",command=lambda:(result.set(False),dialog.destroy())).pack(side="right"); ttk.Button(btns,text="我已了解，开始排版",command=lambda:(result.set(True),dialog.destroy())).pack(side="right",padx=(0,8))
        dialog.protocol("WM_DELETE_WINDOW", lambda:(result.set(False),dialog.destroy())); dialog.update_idletasks(); dialog.geometry("+{}+{}".format(self.winfo_rootx()+max(20,(self.winfo_width()-dialog.winfo_reqwidth())//2), self.winfo_rooty()+max(20,(self.winfo_height()-dialog.winfo_reqheight())//3))); dialog.grab_set(); dialog.focus_force(); self.wait_window(dialog); return bool(result.get())
    def maybe_show_daily_notice(self):
        today = date.today().isoformat()
        if self.last_notice_date == today: return True
        if not self.show_notice_dialog(): return False
        self.last_notice_date = today; self.save_config(show_message=False); return True
    def show_about(self):
        dialog = tk.Toplevel(self); dialog.title(ABOUT_TITLE); dialog.resizable(False,False); dialog.transient(self); body = ttk.Frame(dialog,padding=18); body.pack(fill="both",expand=True); ttk.Label(body,text=ABOUT_TITLE,font=("TkDefaultFont",13,"bold")).pack(anchor="w"); ttk.Label(body,text="版本：{}    {}".format(ABOUT_VERSION, ABOUT_AUTHOR)).pack(anchor="w",pady=(4,10)); msg = tk.Text(body,width=58,height=16,wrap="word",relief="flat",borderwidth=0,background=dialog.cget("background")); msg.insert("1.0", ABOUT_TEXT); msg.configure(state="disabled"); msg.pack(fill="both",expand=True); ttk.Button(body,text="关闭",command=dialog.destroy).pack(anchor="e",pady=(10,0)); dialog.update_idletasks(); dialog.geometry("+{}+{}".format(self.winfo_rootx()+40,self.winfo_rooty()+40)); dialog.grab_set()

    def ask_directory_stable(self, title):
        old_geometry, old_state = self.geometry(), self.state()
        try: old_scaling = float(self.tk.call("tk", "scaling"))
        except Exception: old_scaling = None
        path = filedialog.askdirectory(parent=self, title=title, mustexist=True); self.update_idletasks()
        if old_scaling is not None:
            try: self.tk.call("tk", "scaling", old_scaling)
            except Exception: pass
        try:
            if old_state == "normal": self.geometry(old_geometry)
            else: self.state(old_state)
        except Exception: pass
        return path
    def pick_file(self):
        paths = filedialog.askopenfilenames(title="选择一个或多个 DOCX 文件", filetypes=[("Word DOCX 文档","*.docx"),("旧版 DOC（需转换）","*.doc"),("所有文件","*.*")])
        if paths:
            self.add_paths([Path(x) for x in paths]); self.input_path.set(str(Path(paths[0]).parent))
            if not self.output_path.get(): self.output_path.set(str(Path(paths[0]).parent / "已排版"))
    def pick_folder(self):
        path = self.ask_directory_stable("选择包含 DOCX 文件的文件夹")
        if path:
            folder = Path(path); paths = [x for x in folder.iterdir() if x.is_file() and x.suffix.lower() in (".docx",".doc") and not x.name.startswith("~$") and "_已排版" not in x.stem]
            if not paths: messagebox.showwarning("没有文件", "该文件夹中没有 .docx 文件。"); return
            self.add_paths(paths); self.input_path.set(path)
            if not self.output_path.get(): self.output_path.set(str(folder / "已排版"))
    def add_paths(self, paths):
        existing = {str(t.source.resolve()).casefold() for t in self.batch_tasks}; added = skipped_doc = 0
        for path in paths:
            if path.suffix.lower() == ".doc": skipped_doc += 1; continue
            if path.suffix.lower() != ".docx": continue
            key = str(path.resolve()).casefold()
            if key in existing: continue
            order, _ = split_leading_order(path.stem); self.batch_tasks.append(BatchTask(source=path, detected_order=order)); existing.add(key); added += 1
        if skipped_doc: messagebox.showwarning("已跳过 DOC", "已跳过 {} 个 .doc 文件。当前版本只支持 .docx，请先另存为 .docx。".format(skipped_doc))
        if added: self.sort_tasks()
        self.refresh_task_tree()
    def sort_tasks(self):
        self.batch_tasks.sort(key=lambda t:(0 if t.detected_order is not None else 1, t.detected_order if t.detected_order is not None else 0, natural_sort_key(t.source.stem), t.source.name.casefold())); self.refresh_task_tree()
    def refresh_task_tree(self):
        if not hasattr(self,"queue_tree"): return
        self.queue_tree.delete(*self.queue_tree.get_children())
        for i, task in enumerate(self.batch_tasks, 1): self.queue_tree.insert("","end",iid=str(i-1),values=("{:03d}".format(i), task.source.name))
    def selected_task_indices(self): return sorted(int(iid) for iid in self.queue_tree.selection())
    def move_task(self, direction):
        selected = self.selected_task_indices()
        if len(selected) != 1: messagebox.showinfo("请选择一个文件", "上移或下移时，请只选择一个文件。"); return
        i, ni = selected[0], selected[0] + direction
        if ni < 0 or ni >= len(self.batch_tasks): return
        self.batch_tasks[i], self.batch_tasks[ni] = self.batch_tasks[ni], self.batch_tasks[i]; self.refresh_task_tree(); self.queue_tree.selection_set(str(ni)); self.queue_tree.see(str(ni))
    def remove_selected_tasks(self):
        for i in reversed(self.selected_task_indices()): del self.batch_tasks[i]
        self.refresh_task_tree()
    def clear_tasks(self): self.batch_tasks = []; self.refresh_task_tree()
    def pick_output(self):
        path = self.ask_directory_stable("选择输出目录")
        if path: self.output_path.set(path)
    def log(self, msg): self.log_text.insert("end", msg + "\n"); self.log_text.see("end"); self.update_idletasks()
    def _float(self, name):
        try: return float(self.vars[name].get().strip())
        except ValueError: raise ValueError("“{}”必须填写数字".format(name))

    def collect_settings(self):
        rules = []
        for index, item in enumerate(self.sequence_rule_vars, 1):
            start_text = item["var"].get().strip()
            try: start_value, style, width = parse_sequence_start(start_text)
            except ValueError as exc: raise ValueError("编号{}：{}".format(index, exc))
            rules.append({"name":"编号{}".format(index),"start_text":start_text,"start_value":start_value,"style":style,"width":width})
        return FormatSettings(
            title_line_small=self.vars["title_line_small"].get().strip(), title_line_big=self.vars["title_line_big"].get().strip(), setter=self.vars["setter"].get().strip(), reviewer=self.vars["reviewer"].get().strip(), use_date=self.vars["use_date"].get().strip(),
            insert_mode=self.vars["insert_mode"].get(), marker=self.vars["marker"].get(), page_break_before_repeated_title=self.vars["page_break_before_repeated_title"].get(), sequence_rules=rules,
            top_margin_cm=self._float("top_margin_cm"), bottom_margin_cm=self._float("bottom_margin_cm"), left_margin_cm=self._float("left_margin_cm"), right_margin_cm=self._float("right_margin_cm"), header_distance_cm=self._float("header_distance_cm"), footer_distance_cm=self._float("footer_distance_cm"),
            body_font=self.vars["body_font"].get().strip(), body_size=self._float("body_size"), heading_font=self.vars["heading_font"].get().strip(), heading_size=self._float("heading_size"), title_font=self.vars["title_font"].get().strip(), title_size=self._float("title_size"), small_title_font=self.vars["small_title_font"].get().strip(), small_title_size=self._float("small_title_size"), footer_font=self.vars["footer_font"].get().strip(), footer_size=self._float("footer_size"),
            line_spacing_multiple=self._float("line_spacing_multiple"), first_line_indent_chars=self._float("first_line_indent_chars"), paragraph_after_pt=self._float("paragraph_after_pt"), apply_heading_format=self.vars["apply_heading_format"].get(), heading_patterns=self.pattern_text.get("1.0", "end").strip(), add_footer_page_number=self.vars["add_footer_page_number"].get(), footer_subject_text=self.vars["footer_subject_text"].get().strip(), overwrite=self.vars["overwrite"].get())

    def get_tasks(self):
        if not self.batch_tasks: raise FileNotFoundError("请先添加一个或多个 .docx 文件。")
        missing = [str(t.source) for t in self.batch_tasks if not t.source.exists()]
        if missing: raise FileNotFoundError("以下文件已不存在：\n" + "\n".join(missing[:10]))
        bad = [str(t.source) for t in self.batch_tasks if t.source.suffix.lower() != ".docx"]
        if bad: raise ValueError("当前版本只支持 .docx，请移除或转换以下文件：\n" + "\n".join(bad[:10]))
        return list(self.batch_tasks)

    def run_format(self):
        try:
            if not self.maybe_show_daily_notice(): return
            self.auto_save_current_profile()
            if hasattr(self,"process_button"): self.process_button.configure(state="disabled", text="处理中，请稍候……"); self.update_idletasks()
            settings, tasks = self.collect_settings(), self.get_tasks(); out_dir = Path(self.output_path.get().strip() or (tasks[0].source.parent / "已排版"))
            if settings.overwrite and not messagebox.askyesno("确认覆盖", "已勾选覆盖原文件。建议先备份。是否继续？"): return
            self.log_text.delete("1.0", "end"); self.log("队列中共有 {} 个 .docx 文档，严格按表格顺序处理。".format(len(tasks)))
            for i, task in enumerate(tasks, 1):
                detected = str(task.detected_order) if task.detected_order is not None else "无"; previews = []
                for si, rule in enumerate(settings.sequence_rules or [], 1):
                    val = rule["start_value"] + (i - 1); show = int_to_chinese_number(val) if rule["style"] == "chinese" else ("{:0" + str(rule["width"]) + "d}").format(val); previews.append("编号{}={}".format(si, show))
                self.log("  {:03d}. {} （文件名前序号：{}；{}）".format(i, task.source.name, detected, "；".join(previews)))
            ok, failed = WordFormatter(settings, self.log).process(tasks, out_dir); self.save_config(show_message=False); messagebox.showinfo("处理完成", "成功：{} 个\n失败：{} 个\n输出目录：{}".format(ok, failed, out_dir))
        except Exception as exc:
            self.log(traceback.format_exc()); messagebox.showerror("无法处理", str(exc))
        finally:
            if hasattr(self,"process_button"): self.process_button.configure(state="normal", text="▶ 开始批量处理")

    def safe_profile_name(self, name):
        name = (name or DEFAULT_PROFILE_NAME).strip(); name = re.sub(r'[\\/:*?"<>|]+', "_", name); return name or DEFAULT_PROFILE_NAME
    def profile_path(self, name): return PROFILE_DIR / (self.safe_profile_name(name) + ".json")
    def current_profile_name(self): return self.safe_profile_name(self.profile_var.get() if hasattr(self,"profile_var") else DEFAULT_PROFILE_NAME)
    def profile_payload(self): return {"settings":asdict(self.collect_settings()), "saved_at":date.today().isoformat()}
    def apply_settings_dict(self, s):
        if "title_line_small" not in s:
            s["title_line_small"] = "{} {}".format(s.get("school","廉江市实验学校"), s.get("school_year","2025—2026学年第二学期")).strip()
        if "title_line_big" not in s:
            s["title_line_big"] = "{} {}（{}）卷".format(s.get("grade_subject","高一语文"), s.get("work_name","暑假作业"), s.get("paper_no","{1}")).strip()
        for name, var in self.vars.items():
            if name in s and name != "heading_patterns": var.set(s[name])
        if "heading_patterns" in s: self.pattern_text.delete("1.0","end"); self.pattern_text.insert("1.0", s["heading_patterns"])
        seq = s.get("sequence_rules")
        if seq: self.rebuild_sequence_rules([r.get("start_text","1") for r in seq])
    def refresh_profile_list(self):
        if not hasattr(self,"profile_combo"): return
        PROFILE_DIR.mkdir(parents=True, exist_ok=True); names = sorted([p.stem for p in PROFILE_DIR.glob("*.json")])
        if DEFAULT_PROFILE_NAME not in names: names.insert(0, DEFAULT_PROFILE_NAME)
        self.profile_combo["values"] = names
        if self.profile_var.get() not in names: self.profile_var.set(names[0])
    def save_profile_named(self, name, show_message=False):
        path = self.profile_path(name); path.write_text(json.dumps(self.profile_payload(), ensure_ascii=False, indent=2), encoding="utf-8"); self.refresh_profile_list()
        if show_message: messagebox.showinfo("已保存", "配置方案已保存：\n{}".format(path))
    def auto_save_current_profile(self):
        try: self.save_profile_named(self.current_profile_name(), show_message=False)
        except Exception: pass
    def save_selected_profile(self):
        try: self.save_profile_named(self.current_profile_name(), show_message=True); self.save_config(show_message=False)
        except Exception as exc: messagebox.showerror("保存失败", str(exc))
    def save_profile_as(self):
        name = simpledialog.askstring("另存为新方案", "请输入配置方案名称：", parent=self)
        if not name: return
        self.profile_var.set(self.safe_profile_name(name)); self.save_selected_profile()
    def load_selected_profile(self):
        path = self.profile_path(self.current_profile_name())
        if not path.exists(): self.auto_save_current_profile(); return
        try:
            data = json.loads(path.read_text(encoding="utf-8")); self.apply_settings_dict(data.get("settings", data)); self.save_config(show_message=False); messagebox.showinfo("已应用", "已应用配置方案：{}".format(self.current_profile_name()))
        except Exception as exc: messagebox.showerror("载入失败", str(exc))
    def delete_selected_profile(self):
        name = self.current_profile_name()
        if name == DEFAULT_PROFILE_NAME: messagebox.showinfo("不能删除", "默认自动保存方案不能删除。"); return
        path = self.profile_path(name)
        if path.exists() and messagebox.askyesno("确认删除", "删除配置方案“{}”？".format(name)): path.unlink(); self.profile_var.set(DEFAULT_PROFILE_NAME); self.refresh_profile_list()

    def setup_autosave(self):
        for var in self.vars.values():
            try: var.trace_add("write", lambda *args: self.schedule_auto_save())
            except Exception: pass
        for item in self.sequence_rule_vars:
            try: item["var"].trace_add("write", lambda *args: self.schedule_auto_save())
            except Exception: pass
        try:
            self.pattern_text.bind("<<Modified>>", self._on_pattern_modified)
        except Exception: pass
    def _on_pattern_modified(self, event=None):
        try:
            if self.pattern_text.edit_modified():
                self.pattern_text.edit_modified(False); self.schedule_auto_save()
        except Exception: pass
    def schedule_auto_save(self):
        if not getattr(self, "_autosave_ready", False): return
        if self._autosave_job:
            try: self.after_cancel(self._autosave_job)
            except Exception: pass
        self._autosave_job = self.after(1000, self.auto_save_current_profile)
    def save_config(self, show_message=True):
        try:
            data = {"settings":asdict(self.collect_settings()), "input_path":self.input_path.get(), "output_path":self.output_path.get(), "last_notice_date":self.last_notice_date, "active_profile":self.current_profile_name()}
            CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"); self.auto_save_current_profile()
            if show_message: messagebox.showinfo("已保存", "预设已保存到：\n{}\n\n当前配置方案：{}".format(CONFIG_FILE, self.profile_path(self.current_profile_name())))
        except Exception as exc: messagebox.showerror("保存失败", str(exc))
    def _load_config(self):
        if not CONFIG_FILE.exists(): return
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8")); s = data.get("settings", {}); self.apply_settings_dict(s)
            self.input_path.set(data.get("input_path", "")); self.output_path.set(data.get("output_path", "")); self.last_notice_date = data.get("last_notice_date", "")
            if hasattr(self,"profile_var"): self.profile_var.set(data.get("active_profile", DEFAULT_PROFILE_NAME) or DEFAULT_PROFILE_NAME)
        except Exception: pass
    def reset_defaults(self):
        d = asdict(FormatSettings())
        for name, var in self.vars.items():
            if name in d and name != "heading_patterns": var.set(d[name])
        self.pattern_text.delete("1.0","end"); self.pattern_text.insert("1.0", d["heading_patterns"]); self.rebuild_sequence_rules(["1","001"]); self.auto_save_current_profile()


def main():
    enable_windows_dpi_awareness(); app = App(); app.mainloop()

if __name__ == "__main__":
    main()





