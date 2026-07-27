# -*- coding: utf-8 -*-
"""
假期作业排版模块 - DocumentFormatter

读取同学GUI的配置变量，对组卷网原始DOCX进行：
   - 文件名解析（提取序号、日期）
   - 文件名生成（按规则组装输出文件名）
   - 日期分配（按做/休周期计算每套卷子的日期）
   - 文档结构分析（定位各大题、参考答案）
   - 输出文档构建（标题区 + 正文内容 + 页面格式）

用法：
   formatter = DocumentFormatter(gui_instance)
   output_path = formatter.process_file('源文件.docx', '输出目录')
"""

import re
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


class SourceParseError(ValueError):
    """源文件命名格式不匹配时抛出的异常"""
    pass


class DocumentFormatter:
    """主排版类"""

    CHINESE_NUMS = [
        "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
        "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
        "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十",
        "三十一", "三十二", "三十三", "三十四", "三十五", "三十六", "三十七", "三十八", "三十九", "四十",
        "四十一", "四十二", "四十三", "四十四", "四十五", "四十六", "四十七", "四十八", "四十九", "五十",
        "五十一", "五十二", "五十三", "五十四", "五十五", "五十六", "五十七", "五十八", "五十九", "六十",
        "六十一", "六十二", "六十三", "六十四", "六十五", "六十六", "六十七", "六十八", "六十九", "七十",
        "七十一", "七十二", "七十三", "七十四", "七十五", "七十六", "七十七", "七十八", "七十九", "八十",
        "八十一", "八十二", "八十三", "八十四", "八十五", "八十六", "八十七", "八十八", "八十九", "九十",
        "九十一", "九十二", "九十三", "九十四", "九十五", "九十六", "九十七", "九十八", "九十九",
    ]
    _CHINESE_TO_INT = {cn: i + 1 for i, cn in enumerate(CHINESE_NUMS)}

    def __init__(self, gui):
        self.g = gui

    # ========================== 辅助获取 GUI 值 ==========================
    def _get_text(self, obj):
        """获取 QLineEdit、QComboBox 等控件的文本值"""
        if obj is None:
            return ''
        if hasattr(obj, 'text'):
            return obj.text().strip()
        return str(obj).strip()

    def _get_current_text(self, combo):
        """获取 QComboBox 当前文本"""
        if combo is None:
            return ''
        if hasattr(combo, 'currentText'):
            return combo.currentText()
        return str(combo)

    def _get_int_value(self, obj, default=0):
        """获取 QSpinBox 或数值控件的值"""
        if obj is None:
            return default
        if hasattr(obj, 'value'):
            try:
                return int(obj.value())
            except:
                return default
        try:
            return int(str(obj).strip())
        except:
            return default

    def _get_bool(self, obj):
        """获取 QCheckBox 的选中状态"""
        if obj is None:
            return False
        if hasattr(obj, 'isChecked'):
            return obj.isChecked()
        return bool(obj)

    def _get_attr_text(self, obj):
        if obj is None:
            return ""
        if hasattr(obj, "currentText"):
            return obj.currentText().strip()
        if hasattr(obj, "text"):
            try:
                return obj.text().strip()
            except Exception:
                pass
        return str(obj).strip()

    def _get_qcolor(self, obj):
        """获取 QColor 对象（从 QPushButton 的 property）"""
        if obj is None:
            return None
        if hasattr(obj, 'property'):
            return obj.property('color')
        return None

    # ========================== 主入口 ==========================

    def process_file(self, source_path, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        seq_num, source_date = self._parse_source_filename(os.path.basename(source_path))
        assignment_date = self._calculate_assignment_date(seq_num)
        output_path = os.path.join(output_dir, self._build_output_filename(seq_num, source_date, assignment_date))

        source_doc = Document(source_path)
        output_doc = Document()
        self._apply_page_format(output_doc)
        seq_display = self._format_seq(seq_num)
        self._add_title_section(output_doc, seq_display, seq_num, assignment_date)
        self._copy_source_content(output_doc, source_doc)
        output_doc.save(output_path)
        return output_path

    # ====================== 文件名解析 ==========================

    def _parse_source_filename(self, filename):
        name = filename[:-5] if filename.lower().endswith('.docx') else filename
        source_date = None

        if self._get_bool(self.g.orig_date_check):
            date_pattern = self._date_format_to_regex(self._get_current_text(self.g.orig_date_format))
            if date_pattern:
                m = re.search(date_pattern, name)
                if m:
                    source_date = m.group(0)
                    name = name[:m.start()] + name[m.end():]

        prefix = self._get_text(self.g.orig_prefix_edit)
        suffix = self._get_text(self.g.orig_suffix_edit)

        # 校验前缀
        if prefix:
            if name.startswith(prefix):
                name = name[len(prefix):]
            else:
                raise SourceParseError(
                    f'文件名 "{filename}" 不匹配前缀 "{prefix}"'
                )

        # 校验后缀
        if suffix:
            if name.endswith(suffix):
                name = name[:-len(suffix)]
            else:
                raise SourceParseError(
                    f'文件名 "{filename}" 不匹配后缀 "{suffix}"'
                )

        seq_str = name.strip()
        if not seq_str:
            raise SourceParseError(
                f'文件名 "{filename}" 中未能提取到序号部分'
            )

        seq_num = self._parse_seq_str(seq_str, self._get_current_text(self.g.orig_seq_type_combo))
        return seq_num, source_date

    def _parse_seq_str(self, seq_str, fmt):
        # 试阿拉伯数字（含01格式）
        try:
            return int(seq_str)
        except ValueError:
            pass
        # 试中文数字
        if seq_str in self._CHINESE_TO_INT:
            return self._CHINESE_TO_INT[seq_str]
        raise SourceParseError(
            f'无法将 "{seq_str}" 解析为有效序号'
        )

    # ====================== 文件名生成 ==========================

    def _build_output_filename(self, seq_num, source_date, assignment_date):
        parts = []
        if self._get_bool(self.g.gen_date_check):
            if assignment_date is not None:
                # 使用 assignment_date 格式化，但 assignment_date 是 QDate 对象
                fmt = self._get_current_text(self.g.gen_date_format)
                parts.append(self._format_date_obj(assignment_date, fmt))
        parts.append(self._get_text(self.g.gen_prefix_edit))
        parts.append(self._format_seq(seq_num))
        parts.append(self._get_text(self.g.gen_suffix_edit))
        return ''.join(parts) + '.docx'

    def _format_seq_with_source(self, seq_num, combo):
        combo_text = self._get_current_text(combo)
        if combo_text in ["数字1、2、3", "数字"]:
            return str(seq_num)
        elif combo_text in ["数字01、02、03", "01格式"]:
            return f"{seq_num:02d}"
        elif combo_text in ["汉字一、二、三", "中文"]:
            if 1 <= seq_num <= len(self.CHINESE_NUMS):
                return self.CHINESE_NUMS[seq_num - 1]
            return str(seq_num)
        return str(seq_num)

    def _format_seq(self, seq_num):
        return self._format_seq_with_source(seq_num, self.g.seq_type_combo)

    # ====================== 日期计算 =============================

    def _calculate_assignment_date(self, seq_num):
        start_edit = self.g.start_date
        if start_edit is None:
            return None
        if hasattr(start_edit, 'date'):
            start = start_edit.date()
        else:
            start = start_edit

        # 检查日期计算方式
        cycle_mode = self._get_attr_text(getattr(self.g, 'cycle_type_combo', ''))
        if cycle_mode and "工作日" in cycle_mode:
            return self._calc_weekday_date(start, seq_num)

        # 自定义模式
        do_days = self._get_int_value(self.g.rest_do, 1)
        rest_days = self._get_int_value(self.g.rest_rest, 0)
        if do_days <= 0:
            do_days = 1
        cycle_length = do_days + rest_days
        full_cycles = (seq_num - 1) // do_days
        remainder = (seq_num - 1) % do_days
        total_offset = full_cycles * cycle_length + remainder
        return start.addDays(total_offset)

    def _calc_weekday_date(self, start, seq_num):
        """按工作日计算：周一至周五做，跳过周六日"""
        current = start
        count = 0
        while True:
            if hasattr(current, 'dayOfWeek'):
                dow = current.dayOfWeek()
            else:
                dow = current.isoweekday() if hasattr(current, 'isoweekday') else 0
            if dow <= 5:
                count += 1
                if count == seq_num:
                    return current
            current = current.addDays(1)

    def _find_answer_index(self, doc):
        for i, p in enumerate(doc.paragraphs):
            if '参考答案' in p.text:
                return i
        return -1

    def _find_first_content_index(self, doc):
        """找到第一段以题号开头的正文段落索引，跳过前面的卷头多余内容"""
        import re
        pattern = re.compile(r'^[一二三四五六七八九十]+[、．.]')
        for i, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            if text and pattern.match(text):
                return i
        return 0

    def _copy_source_content(self, output_doc, source_doc):
        first_content_idx = self._find_first_content_index(source_doc)
        body = source_doc.element.body
        para_idx = 0
        tables = source_doc.tables
        table_idx = 0

        for child in body:
            tag = child.tag
            if tag.endswith('}p'):
                if para_idx < len(source_doc.paragraphs):
                    if para_idx >= first_content_idx:
                        self._copy_paragraph(output_doc, source_doc.paragraphs[para_idx])
                    para_idx += 1
            elif tag.endswith('}tbl'):
                if table_idx < len(tables):
                    self._copy_table(output_doc, tables[table_idx])
                    table_idx += 1

    def _copy_paragraph(self, output_doc, src_para):
        text = src_para.text
        if not text and not src_para.runs:
            output_doc.add_paragraph('')
            return
        p = output_doc.add_paragraph()
        if src_para.alignment is not None:
            p.alignment = src_para.alignment
        for run in src_para.runs:
            new_run = p.add_run(run.text)
            if run.font.name:
                try:
                    new_run.font.name = run.font.name
                    r = new_run._element
                    rPr = r.get_or_add_rPr()
                    rFonts = rPr.find(qn('w:rFonts'))
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        rPr.insert(0, rFonts)
                    rFonts.set(qn('w:eastAsia'), run.font.name)
                except Exception:
                    pass
            if run.font.size:
                new_run.font.size = run.font.size
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            # 正文字体强制覆盖
            if self._get_bool(getattr(self.g, 'enforce_body_font', False)):
                try:
                    new_run.font.name = 'Times New Roman'
                    r = new_run._element
                    rPr = r.get_or_add_rPr()
                    rFonts = rPr.find(qn('w:rFonts'))
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        rPr.insert(0, rFonts)
                    rFonts.set(qn('w:ascii'), 'Times New Roman')
                    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                    rFonts.set(qn('w:eastAsia'), '宋体')
                except Exception:
                    pass
        if not src_para.runs and text:
            p.add_run(text)

    def _copy_table(self, output_doc, src_table):
        rows = len(src_table.rows)
        cols = len(src_table.columns)
        if rows == 0 or cols == 0:
            return
        new_table = output_doc.add_table(rows=rows, cols=cols)
        new_table.style = 'Table Grid'
        for r_idx, row in enumerate(src_table.rows):
            for c_idx, cell in enumerate(row.cells):
                new_table.cell(r_idx, c_idx).text = cell.text

    # ====================== 标题区 ===============================

    def _add_title_section(self, doc, seq_display, seq_num, assignment_date):
        if self._get_bool(getattr(self.g, 'header_enable_check', True)):
            self._add_header_line(doc, self.g)
        if self._get_bool(getattr(self.g, 'main_enable_check', True)):
            main_seq_display = self._format_seq_with_source(seq_num, self.g.main_seq_type)
            self._add_main_title_line(doc, self.g, main_seq_display)
        sub_enabled = getattr(self.g, 'enable_sub_cb', None)
        if sub_enabled is None:
            sub_enabled = getattr(self.g, 'sub_enable_check', True)
        if self._get_bool(sub_enabled):
            self._add_subtitle_line(doc, self.g, seq_num, assignment_date)

    def _add_header_line(self, doc, g):
        text = self._get_text(g.header_content_edit)
        p = doc.add_paragraph()
        self._apply_para_format(p, g, 'header')
        if text:
            run = p.add_run(text)
            self._apply_run_format(run, g, 'header')

    def _add_main_title_line(self, doc, g, seq_display):
        prefix = self._get_text(g.main_prefix_edit)
        suffix = self._get_text(g.main_suffix_edit)
        text = prefix + seq_display + suffix
        p = doc.add_paragraph()
        self._apply_para_format(p, g, 'main')
        if text:
            run = p.add_run(text)
            self._apply_run_format(run, g, 'main')

    def _add_subtitle_line(self, doc, g, seq_num, assignment_date):
        parts = []
        examiners = getattr(g, 'examiners', None) or []
        reviewers = getattr(g, 'reviewers', None) or []
        show_examiner = getattr(g, 'enable_examiner_cb', True)
        if self._get_bool(show_examiner):
            if examiners and 0 <= seq_num - 1 < len(examiners):
                parts.append('命题人：' + str(examiners[seq_num - 1]))
            if reviewers and 0 <= seq_num - 1 < len(reviewers):
                parts.append('审题人：' + str(reviewers[seq_num - 1]))
        if assignment_date is not None:
            date_str = self._format_date_obj(assignment_date, 'yyyy年M月d日')
            parts.append(f'使用日期：{date_str}')
        text = '    '.join(parts)
        p = doc.add_paragraph()
        self._apply_para_format(p, g, 'sub')
        if text:
            run = p.add_run(text)
            self._apply_run_format(run, g, 'sub')

    # ====================== 页面格式 =============================

    def _apply_page_format(self, doc):
        section = doc.sections[0]
        g = self.g

        # 纸张
        page_w, page_h = self._parse_paper_size(self._get_attr_text(g.paper_size))
        if page_w and page_h:
            section.page_width = page_w
            section.page_height = page_h

        # 页边距
        # 页边距 (GUI名 -> python-docx节属性名)
        margin_map = {'margin_top': 'top_margin', 'margin_bottom': 'bottom_margin',
                       'margin_left': 'left_margin', 'margin_right': 'right_margin'}
        for gui_name, default in [('margin_top', 2.54), ('margin_bottom', 2.54),
                                    ('margin_left', 3.18), ('margin_right', 3.18)]:
            try:
                val = float(getattr(g, gui_name, default) or default)
                section_attr = margin_map[gui_name]
                setattr(section, section_attr, Cm(val))
            except Exception:
                setattr(section, margin_map[gui_name], Cm(default))

        # 计算右对齐制表位(相对页边缘)
        _pw = section.page_width
        _mr = section.right_margin
        _emu_per_twip = 914400 / 1440
        _right_tab = int((_pw - _mr) / _emu_per_twip)

        # 页眉(左+右同行，右对齐制表位)
        header = section.header
        header.is_linked_to_previous = False
        left_text = self._get_text(g.header_left_edit)
        right_text = self._get_text(g.header_right_edit)
        hp = header.paragraphs[0]
        hp.clear()
        if left_text or right_text:
            from docx.oxml import OxmlElement
            pPr = hp._element.get_or_add_pPr()
            tabs_el = OxmlElement('w:tabs')
            tab_r = OxmlElement('w:tab')
            tab_r.set(qn('w:val'), 'right')
            tab_r.set(qn('w:pos'), str(_right_tab))
            tabs_el.append(tab_r)
            pPr.append(tabs_el)
            if left_text:
                r1 = hp.add_run(left_text)
                self._apply_run_format(r1, g, 'header_footer')
            if left_text and right_text:
                hp.add_run(chr(9))
            if right_text:
                r2 = hp.add_run(right_text)
                self._apply_run_format(r2, g, 'header_footer')
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # 页脚(左+右+页码共存)
        footer = section.footer
        footer.is_linked_to_previous = False
        fl_text = ""
        fr_text = ""
        fp = footer.paragraphs[0]
        fp.clear()
        if self._get_bool(g.insert_page_check):
            self._build_footer_with_page(fp, g, section, fl_text, fr_text)
        elif fl_text or fr_text:
            fp.clear()
            pPr = fp._element.get_or_add_pPr()
            tabs_el = OxmlElement('w:tabs')
            tab_r = OxmlElement('w:tab')
            tab_r.set(qn('w:val'), 'right')
            tab_r.set(qn('w:pos'), str(_right_tab))
            tabs_el.append(tab_r)
            pPr.append(tabs_el)
            if fl_text:
                r1 = fp.add_run(fl_text)
                self._apply_run_format(r1, g, 'header_footer')
            if fl_text and fr_text:
                fp.add_run(chr(9))
            if fr_text:
                r2 = fp.add_run(fr_text)
                self._apply_run_format(r2, g, 'header_footer')
            fp.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def _build_footer_with_page(self, fp, g, section, fl_text, fr_text):
        """构建包含左文本+页码+右文本的页脚段落（制表位法，pos相对页边缘）"""
        from docx.oxml import OxmlElement
        pw = section.page_width      # EMU
        ml = section.left_margin     # EMU
        mr = section.right_margin    # EMU
        emu_per_twip = 914400 / 1440

        # 制表位位置 = 相对于页面左边缘(twip)
        center_tab = int(pw / 2 / emu_per_twip)           # 页面中心
        _right_tab  = int((pw - mr) / emu_per_twip)        # 正文区右边界

        style = self._get_attr_text(getattr(g, 'page_style', ''))
        subject = self._get_text(getattr(g, 'subject_name_edit', ''))
        page_pos = self._get_text(g.page_position)
        need_numpages = '共y页' in style or '/y' in style

        def add_page_number():
            if style == "科目  第x页  共y页":
                if subject:  fp.add_run(subject + "  ")
                fp.add_run("第")
                self._add_page_field_run(fp, g)
                fp.add_run("页  共")
                self._add_numpages_field_run(fp, g)
                fp.add_run("页")
            elif style == "科目  x/y":
                if subject:  fp.add_run(subject + "  ")
                self._add_page_field_run(fp, g)
                fp.add_run("/")
                self._add_numpages_field_run(fp, g)
            elif style == "科目  x":
                if subject:  fp.add_run(subject + "  ")
                self._add_page_field_run(fp, g)
            elif style == "第x页  共y页":
                fp.add_run("第")
                self._add_page_field_run(fp, g)
                fp.add_run("页  共")
                self._add_numpages_field_run(fp, g)
                fp.add_run("页")
            elif style == "x/y":
                self._add_page_field_run(fp, g)
                fp.add_run("/")
                self._add_numpages_field_run(fp, g)
            elif style == "x":
                self._add_page_field_run(fp, g)
            else:
                prefix = self._get_text(getattr(g, 'page_prefix_edit', ''))
                suffix = self._get_text(getattr(g, 'page_suffix_edit', ''))
                if prefix:
                    r = fp.add_run(prefix)
                    self._apply_run_format(r, g, 'page')
                self._add_page_field_run(fp, g)
                if suffix:
                    r = fp.add_run(suffix)
                    self._apply_run_format(r, g, 'page')

        # 无左右文本 → 直接用段落对齐
        if not fl_text and not fr_text:
            add_page_number()
            if '中' in page_pos:
                fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif '右' in page_pos:
                fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            return

        # 有左右文本 → 用制表位精确定位
        pPr = fp._element.get_or_add_pPr()
        tabs_el = OxmlElement('w:tabs')

        if '中' in page_pos:
            # 左文本←左边缘 | TAB | 页码→居中 | TAB | 右文本→右边缘
            tab_c = OxmlElement('w:tab')
            tab_c.set(qn('w:val'), 'center')
            tab_c.set(qn('w:pos'), str(center_tab))
            tabs_el.append(tab_c)
            tab_r = OxmlElement('w:tab')
            tab_r.set(qn('w:val'), 'right')
            tab_r.set(qn('w:pos'), str(_right_tab))
            tabs_el.append(tab_r)
            pPr.append(tabs_el)
            if fl_text:
                r1 = fp.add_run(fl_text)
                self._apply_run_format(r1, g, 'header_footer')
            fp.add_run(chr(9))
            add_page_number()
            if fr_text:
                fp.add_run(chr(9))
                r2 = fp.add_run(fr_text)
                self._apply_run_format(r2, g, 'header_footer')

        elif '右' in page_pos:
            # 左文本←左边缘 | TAB | 页码+右文本→右边缘
            tab_r = OxmlElement('w:tab')
            tab_r.set(qn('w:val'), 'right')
            tab_r.set(qn('w:pos'), str(_right_tab))
            tabs_el.append(tab_r)
            pPr.append(tabs_el)
            if fl_text:
                r1 = fp.add_run(fl_text)
                self._apply_run_format(r1, g, 'header_footer')
            fp.add_run(chr(9))
            add_page_number()
            if fr_text:
                r2 = fp.add_run(fr_text)
                self._apply_run_format(r2, g, 'header_footer')

        else:  # 靠左
            # 左文本+页码←左边缘 | TAB | 右文本→右边缘
            tab_r = OxmlElement('w:tab')
            tab_r.set(qn('w:val'), 'right')
            tab_r.set(qn('w:pos'), str(_right_tab))
            tabs_el.append(tab_r)
            pPr.append(tabs_el)
            if fl_text:
                r1 = fp.add_run(fl_text)
                self._apply_run_format(r1, g, 'header_footer')
            add_page_number()
            if fr_text:
                fp.add_run(chr(9))
                r2 = fp.add_run(fr_text)
                self._apply_run_format(r2, g, 'header_footer')

        fp.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # 页码颜色
        color = self._get_qcolor(g.page_color_btn)
        if color is not None:
            rgb = self._qcolor_to_rgb(color)
            if rgb:
                for run in fp.runs:
                    self._set_run_color(run, rgb)


    def _build_page_runs(self, fp, g, style, subject, need_numpages):
        """根据 page_style 在段落中添加页码文本+字段代码"""
        from docx.oxml import OxmlElement

        if style == "\u79d1\u76ee  \u7b2cx\u9875  \u5171y\u9875":
            if subject:
                fp.add_run(subject + "  ")
            fp.add_run("\u7b2c")
            self._add_page_field_run(fp, g)
            fp.add_run("\u9875  \u5171")
            self._add_numpages_field_run(fp, g)
            fp.add_run("\u9875")
        elif style == "\u79d1\u76ee  x/y":
            if subject:
                fp.add_run(subject + "  ")
            self._add_page_field_run(fp, g)
            fp.add_run("/")
            self._add_numpages_field_run(fp, g)
        elif style == "\u79d1\u76ee  x":
            if subject:
                fp.add_run(subject + "  ")
            self._add_page_field_run(fp, g)
        elif style == "\u7b2cx\u9875  \u5171y\u9875":
            fp.add_run("\u7b2c")
            self._add_page_field_run(fp, g)
            fp.add_run("\u9875  \u5171")
            self._add_numpages_field_run(fp, g)
            fp.add_run("\u9875")
        elif style == "x/y":
            self._add_page_field_run(fp, g)
            fp.add_run("/")
            self._add_numpages_field_run(fp, g)
        elif style == "x":
            self._add_page_field_run(fp, g)
        else:
            # fallback: 旧版前缀+PAGE+后缀
            prefix = self._get_text(getattr(g, 'page_prefix_edit', ''))
            suffix = self._get_text(getattr(g, 'page_suffix_edit', ''))
            if prefix:
                r = fp.add_run(prefix)
                self._apply_run_format(r, g, 'page')
            self._add_page_field_run(fp, g)
            if suffix:
                r = fp.add_run(suffix)
                self._apply_run_format(r, g, 'page')

    def _add_page_field_run(self, fp, g):
        from docx.oxml import OxmlElement
        run1 = fp.add_run()
        fld1 = OxmlElement('w:fldChar')
        fld1.set(qn('w:fldCharType'), 'begin')
        run1._element.append(fld1)
        run2 = fp.add_run()
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = ' PAGE '
        run2._element.append(instr)
        self._apply_run_format(run2, g, 'page')
        run3 = fp.add_run()
        fld2 = OxmlElement('w:fldChar')
        fld2.set(qn('w:fldCharType'), 'end')
        run3._element.append(fld2)

    def _add_numpages_field_run(self, fp, g):
        from docx.oxml import OxmlElement
        run1 = fp.add_run()
        fld1 = OxmlElement('w:fldChar')
        fld1.set(qn('w:fldCharType'), 'begin')
        run1._element.append(fld1)
        run2 = fp.add_run()
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = ' NUMPAGES '
        run2._element.append(instr)
        self._apply_run_format(run2, g, 'page')
        run3 = fp.add_run()
        fld2 = OxmlElement('w:fldChar')
        fld2.set(qn('w:fldCharType'), 'end')
        run3._element.append(fld2)

    def _calc_space_padding(self, left_text, right_text, section):
        """计算左右文本之间的空格填充数量"""
        pw = section.page_width  # EMU
        ml = section.left_margin
        mr = section.right_margin
        text_w = (pw - ml - mr) / 914400 * 2.54  # EMU -> cm

        def char_width_cm(s):
            w = 0
            for c in s:
                w += 0.22 if ord(c) > 127 else 0.11
            return w

        lw = char_width_cm(left_text) if left_text else 0
        rw = char_width_cm(right_text) if right_text else 0
        space_w = 0.10  # 一个空格的宽度(cm)
        rest = text_w - lw - rw
        return max(0, int(rest / space_w))

    # ====================== 格式辅助 =============================

    def _apply_run_format(self, run, g, prefix):
        font_name = self._get_attr_text(getattr(g, f'{prefix}_font', ''))
        if font_name:
            try:
                run.font.name = font_name
                r = run._element
                rPr = r.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = OxmlElement('w:rFonts')
                    rPr.insert(0, rFonts)
                rFonts.set(qn('w:eastAsia'), font_name)
            except Exception:
                pass

        size_str = self._get_attr_text(getattr(g, f'{prefix}_font_size', ''))
        pt = self._parse_font_size(size_str)
        if pt:
            run.font.size = Pt(pt)

        run.bold = self._get_bool(getattr(g, f'{prefix}_bold', False))
        run.underline = self._get_bool(getattr(g, f'{prefix}_underline', False))

        color = self._get_qcolor(getattr(g, f'{prefix}_color_btn', None))
        if color is not None:
            rgb = self._qcolor_to_rgb(color)
            if rgb:
                self._set_run_color(run, rgb)

    def _apply_para_format(self, p, g, prefix):
        align_text = self._get_attr_text(getattr(g, f'{prefix}_align', ''))
        if '中' in align_text:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif '右' in align_text:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif '左' in align_text:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            # 主标题和副标题默认居中
            if prefix in ("main", "sub"):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        ls_type = self._get_attr_text(getattr(g, f'{prefix}_line_spacing_type', ''))
        ls_value = self._get_attr_text(getattr(g, f'{prefix}_line_spacing_value', ''))
        if ls_type and ls_value:
            try:
                val = float(ls_value)
                if '倍' in ls_type:
                    p.paragraph_format.line_spacing = val
                elif '固定' in ls_type:
                    p.paragraph_format.line_spacing = Pt(val)
                else:
                    p.paragraph_format.line_spacing = val
            except Exception:
                pass

    # ====================== 工具 =================================

    def _bool_val(self, v):
        # 保留兼容，但不再使用，使用 _get_bool
        return self._get_bool(v)

    def _int_val(self, v, default=0):
        # 保留兼容，但不再使用，使用 _get_int_value
        return self._get_int_value(v, default)

    def _parse_font_size(self, s):
        s = s.strip()
        if not s:
            return None
        SIZE_MAP = {
            '初号': 42, '小初': 36, '一号': 26, '小一': 24,
            '二号': 22, '小二': 18, '三号': 16, '小三': 15,
            '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
            '六号': 7.5, '小六': 6.5, '七号': 5.5, '八号': 5,
        }
        if s in SIZE_MAP:
            return SIZE_MAP[s]
        try:
            return float(s)
        except ValueError:
            return None

    def _qcolor_to_rgb(self, qcolor):
        try:
            return (qcolor.red(), qcolor.green(), qcolor.blue())
        except Exception:
            return None

    def _set_run_color(self, run, rgb):
        rPr = run._element.get_or_add_rPr()
        color_elem = rPr.find(qn('w:color'))
        if color_elem is None:
            color_elem = OxmlElement('w:color')
            rPr.append(color_elem)
        hex_color = '{:02X}{:02X}{:02X}'.format(*rgb)
        color_elem.set(qn('w:val'), hex_color)

    def _date_format_to_regex(self, fmt):
        if not fmt:
            return None
        pattern = re.escape(fmt)
        pattern = pattern.replace('yyyy', r'(\d{4})')
        pattern = pattern.replace('yy', r'(\d{2})')
        pattern = pattern.replace('MM', r'(\d{2})')
        pattern = pattern.replace('M', r'(\d{1,2})')
        pattern = pattern.replace('dd', r'(\d{2})')
        pattern = pattern.replace('d', r'(\d{1,2})')
        return pattern

    def _format_date_obj(self, date_obj, fmt):
        if date_obj is None:
            return ''
        try:
            qfmt = fmt.replace('yyyy', 'yyyy').replace('yy', 'yy')
            qfmt = qfmt.replace('MM', 'MM').replace('M', 'M')
            qfmt = qfmt.replace('dd', 'dd').replace('d', 'd')
            return date_obj.toString(qfmt)
        except AttributeError:
            pass
        try:
            import datetime as dt
            if isinstance(date_obj, dt.date):
                py_fmt = fmt.replace('yyyy', '%Y').replace('yy', '%y')
                py_fmt = py_fmt.replace('MM', '%m')
                py_fmt = py_fmt.replace('dd', '%d')
                py_fmt = py_fmt.replace('M', '%-m')
                py_fmt = py_fmt.replace('d', '%-d')
                return date_obj.strftime(py_fmt)
        except Exception:
            pass
        return str(date_obj)

    def _parse_paper_size(self, text):
        if not text:
            return None, None
        known = {
            'A3': (Cm(29.7), Cm(42.0)),
            'A4': (Cm(21.0), Cm(29.7)),
            'A5': (Cm(14.8), Cm(21.0)),
            'B4': (Cm(25.7), Cm(36.4)),
            'B5': (Cm(18.2), Cm(25.7)),
            '16开': (Cm(18.4), Cm(26.0)),
            '32开': (Cm(13.0), Cm(18.4)),
            'Letter': (Cm(21.59), Cm(27.94)),
            'Legal': (Cm(21.59), Cm(35.56)),
        }
        for key, dim in known.items():
            if key in text:
                return dim
        return None, None


if __name__ == '__main__':
    print("DocumentFormatter 模块加载成功")
