import sys
import os
import datetime

# 修复 Qt 平台插件路径（双击运行时不自动加载）
import PyQt5

os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(
    os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins', 'platforms'
)

import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QFileDialog, QMessageBox,
    QGroupBox, QCheckBox, QColorDialog, QDateEdit, QDialog,
    QTextEdit, QDialogButtonBox, QDoubleSpinBox, QSpinBox, QFrame,
    QListWidget, QListWidgetItem, QAbstractItemView, QScrollArea,
    QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt, QDate, QThread, pyqtSignal, QUrl
from PyQt5.QtGui import QFontDatabase, QColor, QIcon, QPixmap, QDesktopServices
import pythoncom
import win32com.client
import re
import zipfile
from docx_formatter import DocumentFormatter, SourceParseError


class CustomTabWidget(QTabWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent

    def setCurrentIndex(self, index):
        if self.currentIndex() == 1 and index != 1:
            if self.parent and not self.parent.validate_title_inputs():
                QMessageBox.warning(self.parent, "输入错误",
                                    "标题页存在非法输入，请修正后再切换。")
                return
        super().setCurrentIndex(index)


class FileInfoThread(QThread):
    finished = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def __init__(self, file_list):
        super().__init__()
        self.file_list = file_list

    def run(self):
        result = []
        pythoncom.CoInitialize()
        word_app = None
        try:
            try:
                word_app = win32com.client.Dispatch("Kwps.Application")
            except:
                try:
                    word_app = win32com.client.Dispatch("wps.Application")
                except:
                    word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0  # 禁用所有Word对话框
            word_app.ScreenUpdating = False
            for fullpath in self.file_list:
                try:
                    if not os.path.exists(fullpath):
                        raise FileNotFoundError(f"文件不存在: {fullpath}")
                    if fullpath.lower().endswith(".docx"):
                        try:
                            with zipfile.ZipFile(fullpath, "r") as zf:
                                zf.testzip()
                        except zipfile.BadZipFile:
                            raise ValueError("文件已损坏，无法作为有效的docx文件打开")
                        except Exception:
                            raise ValueError("文件已损坏，无法作为有效的docx文件打开")
                    doc = word_app.Documents.Open(fullpath)
                    doc.Repaginate()
                    pages = doc.ComputeStatistics(2)  # wdStatisticPages
                    doc.Close()
                    mtime = os.path.getmtime(fullpath)
                    size = os.path.getsize(fullpath)
                    result.append((fullpath, pages, mtime, size, None))
                except Exception as e:
                    # 任何异常（包括损坏文件）均标记错误
                    try:
                        mtime = os.path.getmtime(fullpath)
                        size = os.path.getsize(fullpath)
                    except:
                        mtime = 0
                        size = 0
                    result.append((fullpath, -1, mtime, size, str(e)))
        except Exception as e:
            self.error_signal.emit(f"初始化 Word/WPS 失败：{str(e)}")
            for fullpath in self.file_list:
                try:
                    mtime = os.path.getmtime(fullpath)
                    size = os.path.getsize(fullpath)
                except:
                    mtime = 0
                    size = 0
                result.append((fullpath, -1, mtime, size, "Word/WPS 初始化失败"))
        finally:
            if word_app:
                word_app.Quit()
            pythoncom.CoUninitialize()
        self.finished.emit(result)


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("试卷排版工具")
        if getattr(sys, "frozen", False):
            icon_path = os.path.join(sys._MEIPASS, "layout.ico")
        else:
            icon_path = "layout.ico"
        self.setWindowIcon(QIcon(icon_path))
        self.setGeometry(100, 100, 1200, 900)
        self.center()
        self.initUI()
        self.examiners = []
        self.reviewers = []
        self.preset_config_path = None
        self._preset_gen_file_pages = None
        self.process_file_list = []
        self.process_file_info = []

    def center(self):
        screen = QApplication.primaryScreen()
        if screen:
            geometry = screen.availableGeometry()
            frame = self.frameGeometry()
            fw, fh = frame.width(), frame.height()
            # 窗口未显示时 frameGeometry() 不包含标题栏和边框，用 API 估算
            if fh <= self.height():
                try:
                    import ctypes
                    user32 = ctypes.windll.user32
                    cap = user32.GetSystemMetrics(4)   # SM_CYCAPTION
                    border = user32.GetSystemMetrics(33)  # SM_CYSIZEFRAME
                    pad = user32.GetSystemMetrics(92)     # SM_CXPADDEDBORDER
                    extra = cap + border + pad * 2
                    fh = self.height() + extra
                except Exception:
                    fh = self.height() + 39  # 备用：约 31px 标题栏 + 8px 边框
            x = geometry.x() + (geometry.width() - fw) // 2
            y = geometry.y() + (geometry.height() - fh) // 2
            self.move(x, y)

    def initUI(self):
        self.tabs = CustomTabWidget(self)
        self.setCentralWidget(self.tabs)
        tab_file = QWidget()
        self.tabs.addTab(tab_file, "文件名")
        self.initFileTab(tab_file)
        self.tab_title = QWidget()
        self.tabs.addTab(self.tab_title, "标题")
        self.initTitleTab(self.tab_title)
        tab_page = QWidget()
        self.tabs.addTab(tab_page, "页面")
        self.initPageTab(tab_page)
        tab_process = QWidget()
        self.tabs.addTab(tab_process, "处理")
        self.initProcessTab(tab_process)
        tab_about = QWidget()
        self.tabs.addTab(tab_about, "关于我们")
        self.initAboutTab(tab_about)

    # ==================== 关于我们 ====================
    def initAboutTab(self, parent):
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(30, 20, 30, 20)

        # Title
        title = QLabel("关于我们")
        title.setStyleSheet("font-size: 32px; font-weight: bold; color: #2c3e50;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Subtitle
        subtitle = QLabel("一个简单的试卷文件排版工具，支持文件名整理、标题重排、页面设置等功能")
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("font-size: 20px; color: #555; margin-top: 8px;")
        subtitle.setAlignment(Qt.AlignCenter)
        layout.addWidget(subtitle)
        layout.addSpacing(10)

        # Operation instructions button
        self.about_show_help_btn = QPushButton("📖 查看操作说明")
        self.about_show_help_btn.setFixedSize(220, 50)
        self.about_show_help_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db; color: white; border-radius: 8px;
                font-size: 18px; font-weight: bold;
            }
            QPushButton:hover { background-color: #2980b9; }
        """)
        self.about_show_help_btn.clicked.connect(self.show_help_dialog)
        h_btn = QHBoxLayout()
        h_btn.addStretch()
        h_btn.addWidget(self.about_show_help_btn)
        h_btn.addStretch()
        layout.addLayout(h_btn)
        layout.addSpacing(15)

        # ========== 支持我们 ==========
        support_box = QFrame()
        support_box.setStyleSheet("""
            QFrame {
                background-color: #e3f2fd;
                border-radius: 12px;
            }
        """)
        support_box.setMaximumWidth(1160)
        support_layout = QHBoxLayout(support_box)
        support_layout.setContentsMargins(30, 25, 40, 25)
        support_layout.setSpacing(20)

        # Left side: text content
        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(0)

        # Title
        support_title = QLabel("请支持我们！")
        support_title.setStyleSheet("font-size: 28px; font-weight: bold; color: #1565c0; font-family: 'Microsoft YaHei', '微软雅黑';")
        support_title.setFixedHeight(34)
        text_layout.addWidget(support_title)
        text_layout.addSpacing(28)

        # Description text
        support_desc = QLabel(
            "<p style=\"line-height: 1.7; margin: 0;\">作为在校高中生，创作不易，</p>"
            "<p style=\"line-height: 1.7; margin: 0;\">希望满意的老师们给予一些小小的支持。</p>"
            "<p style=\"line-height: 1.7; margin: 0;\">您的鼓励是我们不懈前进的动力。</p>")
        support_desc.setWordWrap(True)
        support_desc.setStyleSheet("font-size: 22px; color: #2c3e50; font-family: '楷体', 'KaiTi', 'STKaiti';")
        support_desc.setFixedWidth(480)
        text_layout.addWidget(support_desc)
        text_layout.addSpacing(28)

        # Tip button
        tip_btn = QPushButton("❤爱发电")
        tip_btn.setFixedSize(220, 55)
        tip_btn.setStyleSheet("""
            QPushButton {
                background-color: #f06292; color: white; border-radius: 10px;
                font-size: 20px; font-weight: bold; font-family: 'Microsoft YaHei', '微软雅黑';
            }
            QPushButton:hover { background-color: #e91e63; }
        """)
        tip_btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://ifdian.net/a/desire_1?tab=home")))
        text_layout.addWidget(tip_btn)
        text_layout.addStretch()

        # ========== 版本与作者信息 ==========
        info_box = QFrame()
        info_box.setStyleSheet("""
            QFrame {
                background-color: #e8f5e9;
                border-radius: 10px;
            }
        """)
        info_layout = QVBoxLayout(info_box)
        info_layout.setContentsMargins(16, 12, 16, 12)
        info_layout.setSpacing(10)
        info_layout.setAlignment(Qt.AlignCenter)

        ib_style = "font-size: 20px; color: #444;"

        # Version row
        h_ver = QHBoxLayout()
        h_ver.setAlignment(Qt.AlignCenter)
        v_tag = QLabel("版本 V2.14")
        v_tag.setStyleSheet(ib_style + " font-weight: bold;")
        h_ver.addWidget(v_tag)
        h_ver.addSpacing(20)
        d_tag = QLabel("更新日期 20260719")
        d_tag.setStyleSheet(ib_style)
        h_ver.addWidget(d_tag)
        info_layout.addLayout(h_ver)

        # Developer - merged
        author_title = QLabel("软件作者（合作完成）：")
        author_title.setStyleSheet("font-size: 20px; color: #444;")
        author_title.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(author_title)

        author_names = QLabel(
            '<a href="https://space.bilibili.com/1610128267" style="color: #0000EE; text-decoration: underline;">desire（排版内核）</a>'
            "、"
            '<a href="https://space.bilibili.com/650793568" style="color: #0000EE; text-decoration: underline;">三春牛-创客（外观UI）</a>')
        author_names.setStyleSheet("font-size: 20px; color: #444;")
        author_names.setOpenExternalLinks(True)
        author_names.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(author_names)

        qq_tag = QLabel("如有疑问，请添加QQ:3158510381")
        qq_tag.setStyleSheet("font-size: 20px; color: #444; font-weight: bold;")
        qq_tag.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(qq_tag)

        # Copyright
        cp_tag = QLabel("<p style=\"line-height: 1.7; margin: 0;\">本软件仅供学习交流使用，请勿用于商业用途</p>")
        cp_tag.setStyleSheet("font-size: 16px; color: #888;")
        cp_tag.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(cp_tag)

        text_layout.addWidget(info_box)
        text_layout.addStretch()

        text_widget = QWidget()
        text_widget.setLayout(text_layout)
        support_layout.addWidget(text_widget)
        support_layout.addSpacing(15)

        # Right side: tip image (if exists) - supports frozen/packaged mode
        tip_loaded = False
        if getattr(sys, "frozen", False):
            # Packaged: try internal resource first
            try:
                tip_path = os.path.join(sys._MEIPASS, "tip.jpg")
                if os.path.exists(tip_path):
                    tip_pixmap = QPixmap(tip_path)
                    if not tip_pixmap.isNull():
                        tip_loaded = True
            except Exception:
                pass
        if not tip_loaded:
            # Not packaged or internal failed: try external file
            tip_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tip.jpg")
            if os.path.exists(tip_path):
                tip_pixmap = QPixmap(tip_path)
                if not tip_pixmap.isNull():
                    tip_loaded = True
        if tip_loaded:
            tip_img = QLabel()
            tip_pixmap = tip_pixmap.scaled(520, 520, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            tip_img.setPixmap(tip_pixmap)
            tip_img.setAlignment(Qt.AlignCenter)
            tip_img.setFixedSize(540, 540)
            support_layout.addWidget(tip_img)
        h_box = QHBoxLayout()
        h_box.addStretch()
        h_box.addWidget(support_box)
        h_box.addStretch()
        layout.addLayout(h_box)
        layout.addStretch()

    # ==================== 操作说明 ====================
    def show_help_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("操作说明")
        dialog.resize(750, 600)
        layout = QVBoxLayout(dialog)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setStyleSheet("""
            QTextEdit {
                font-size: 20px;
                line-height: 2.0;
                padding: 15px;
                background-color: #fafafa;
            }
        """)
        help_content = """
<h2 style="text-align:center; color:#2c3e50; margin-top:5px;">试卷排版工具</h2>
<p style="text-align:center; color:#555; font-size:22px;">本软件提供试卷批量插入题头、标题、副标题、页眉页脚、设置页面等功能，能大幅提高老师们排版多份试卷时的工作效率。</p>

<h3 style="color:#2980b9; margin-top:15px;">一、准备工作和文件名设置</h3>
<p>① 准备好所有的试卷（docx文档格式）并放在一个文件夹里，确保这些试卷文档文件名统一格式为“前缀（+日期）+试卷序号（支持三种序号）+后缀，前后缀可选。试卷文档里只有题目，没有标题；</p>
<p>② 在“文件名”标签页中浏览你的文件夹，软件会自动把输出文件夹放在你的文件夹里，你也可以手动设置输出文件夹；</p>
<p>③ 填写好你的原文件名格式、输出文件的生成文件名格式、你的学科名称；</p>
<p>④ 确认无误后点击“检查输入并录入软件”按钮，然后软件会自动检测符合要求的文件，然后录入其中。然后，请转到“标题”标签页。</p>

<h3 style="color:#2980b9; margin-top:15px;">二、标题设置</h3>
<p style="color:#888; font-size:18px;">注：此处设置的是文档内部开头的标题</p>
<p>① 请在此标签设置好题头（例如xxx学校xx年级暑假作业）（刚才设置的标题在“页面”中才会被读取，如果需要在标题中显示学科，请手动添加）、主标题、副标题（根据需求来填写，若不需要也可不填，但至少要填其中的一个框）。</p>
<p>② 关于“出题人、审题人录入”功能：点击按钮后会显示两个大输入框，请老师们在表格或文档中设置好每份试卷的出题人和审题人，一行一个，每行分别对应每份试卷（按照01、02、03……（或其他序号）的顺序排列，然后直接粘贴到输入框中。若不需要，也可以不录入。</p>
<p>③ 若需设置每份试卷的完成日期，请设置好开始日期、结束日期、休息设置，若无休息，将“休”的天数设置为0，“做”的天数设置为≤试卷份数即可。处理好后，请点击“检查日期”按钮，确认无误后切换到“页面”标签页。</p>

<h3 style="color:#2980b9; margin-top:15px;">三、页面设置</h3>
<p>请在此标签页设置文档的页面，包括页边距、页眉页脚、纸张大小、题目（注意是文档现有的内容）行距，设置好后请转到“处理”标签页。</p>

<h3 style="color:#2980b9; margin-top:15px;">四、处理文档</h3>
<p>① 上方的表格框显示的是待处理的文件，这些文件在录入软件时已经自动添加上去，你可以调整处理的顺序（关系到文件序号、出审题人、完成日期等信息，请务必检查好）、删除不需要处理的文件（如只是临时不想处理，可以在表格框中取消勾选），如果发现这些文件不是需要处理的，可以清空列表，然后回到“文件名”标签页重新设置并重新录入。</p>
<p>② 如果领导有文件页数的要求，软件添加标题后可能页数不符合要求，你可以勾选“设置并检查生成文件的期望页数”然后设置每份文档的期望页数，如果页数统一，可以设置第一项并点击“一键应用第一项页数数据”快速设置。下方的框里显示的是生成文件的列表。</p>
<p>③ 请根据需要选择输出的格式，支持.docx格式和.pdf格式，也可以两者都输出。</p>
<p>④ 确认一切无误后，就可以点击“一键排版”按钮开始排版了！</p>

<h3 style="color:#2980b9; margin-top:15px;">五、保存和打开预设</h3>
<p>如果你不想每次都手动输入大量的内容，可以把所有信息填好并检查后点击“处理”标签页中的“保存当前预设”按钮，然后把配置文件保存起来，下次打开的时候就可以直接在“文件名”标签页中点击“打开预设配置”按钮并浏览你保存的预设配置文件，打开后软件会自动根据保存在文件里的内容填写到对应的位置上，你只需检查并根据实际需要进行微调即可。记得在“文件名”标签页中点击“检查输入并录入软件”按钮以将现有文件录入软件。</p>

<h3 style="color:#e67e22; margin-top:15px;">关于错误文件</h3>
<p>如果你添加到了不可识别的文件，软件会给出提示，请将其删除，检查好后再添加。若处理失败，请检查文件权限和内容是否有误。</p>
"""
        text.setHtml(help_content)
        layout.addWidget(text)

        btn_close = QPushButton("关闭")
        btn_close.setFixedSize(100, 32)
        btn_close.clicked.connect(dialog.accept)
        hb = QHBoxLayout()
        hb.addStretch()
        hb.addWidget(btn_close)
        hb.addStretch()
        layout.addLayout(hb)

        dialog.exec_()

    def initFileTab(self, parent):
        layout = QVBoxLayout(parent)
        tip_top = QLabel(
            "<div style=\"font-size:19px; color:#2a5caa; line-height:1.1;\">"
            "💡操作提示<b>（用前必看）</b>：<br>"
            "①输入或浏览当前试卷<b>（确保只含有正文部分）</b>的目录，然后根据现有文件输入文件名前后缀并<b>选择序号类型</b>；<br>"
            "②如果当前文件名有不同的日期，请<b>勾选“日期”并选择日期格式</b>，程序将自动匹配文件名中符合该格式的任意日期；<br>"
            "③输入<b>生成文件名</b>的前后缀并选择序号类型（可选日期）；<br>"
            "④点击<b>“检查输入并录入软件”</b>按钮以让程序获取要处理的文件。<br>"
            "⑤若需另外添加文件或不处理部分文件，请<b>设置完前三个标签页</b>后转到“处理”标签页进行操作。<br>"
            "<span style=\"font-size:18px; font-weight:normal; font-style:italic; color:black; font-family:&#39;楷体&#39;, KaiTi, serif;\">"
            "注：现有文件名是原卷的文件名，供程序读取；<br>"
            "生成文件名是程序排版后输出文件的文件名，一般根据实际教导处安排来填写即可。<br></span>"
            "<span style=\"font-size:20px; font-weight:bold; color:red; font-family:&#39;楷体&#39;, KaiTi, serif;\">"
            "⭐ 完整操作说明请前往“关于我们”标签页中查看。</span>"
            "</div>"
        )
        tip_top.setStyleSheet("background-color: #f0f8ff; padding: 8px; border: 1px solid #c0d8e8;")
        tip_top.setWordWrap(True)
        layout.addWidget(tip_top)
        # ---- 预设配置 ----
        hbox_config = QHBoxLayout()
        self.config_path_label = QLabel("未加载预设配置")
        self.config_path_label.setStyleSheet("color: gray; font-style: italic;")
        self.config_path_label.setObjectName("config_path_label")
        hbox_config.addWidget(self.config_path_label)
        hbox_config.addStretch()
        load_preset_btn = QPushButton("打开预设配置")
        load_preset_btn.clicked.connect(self.loadPreset)
        load_preset_btn.setObjectName("load_preset_btn")
        hbox_config.addWidget(load_preset_btn)
        layout.addLayout(hbox_config)
        # 文件夹位置
        hbox1 = QHBoxLayout()
        hbox1.addWidget(QLabel("文件夹位置："))
        self.folder_edit = QLineEdit()
        self.folder_edit.setObjectName("folder_edit")
        self.folder_edit.textChanged.connect(self.on_folder_changed)
        hbox1.addWidget(self.folder_edit)
        browse_btn = QPushButton("浏览")
        browse_btn.clicked.connect(self.browseFolder)
        hbox1.addWidget(browse_btn)
        layout.addLayout(hbox1)
        # 生成文档输出目录
        hbox_out = QHBoxLayout()
        hbox_out.addWidget(QLabel("生成文档输出目录："))
        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setObjectName("output_folder_edit")
        hbox_out.addWidget(self.output_folder_edit)
        browse_out_btn = QPushButton("浏览")
        browse_out_btn.clicked.connect(self.browseOutputFolder)
        hbox_out.addWidget(browse_out_btn)
        layout.addLayout(hbox_out)
        # ---- 原文件名格式 ----
        hbox2 = QHBoxLayout()
        hbox2.addWidget(QLabel("原文件名格式"))
        self.orig_date_check = QCheckBox("日期")
        self.orig_date_check.setObjectName("orig_date_check")
        hbox2.addWidget(self.orig_date_check)
        # 前缀
        self.orig_prefix_edit = QLineEdit()
        self.orig_prefix_edit.setToolTip("输入前缀，如 '2025-2026廉实语文'")
        self.orig_prefix_edit.setObjectName("orig_prefix_edit")
        hbox2.addWidget(self.orig_prefix_edit)
        # 日期格式下拉框
        self.orig_date_format = QComboBox()
        self.orig_date_format.addItems([
            "yyyy年MM月dd日", "MM月dd日", "yyyy.MM.dd", "MM.dd",
            "[dd]th,MM(Eng),yyyy", "[dd]th,MM(Eng)", "yyyy-MM-dd", "MM-dd"
        ])
        self.orig_date_format.setCurrentText("mm月dd日")
        self.orig_date_format.setObjectName("orig_date_format")
        self.orig_date_format.setVisible(False)  # 默认隐藏
        hbox2.addWidget(self.orig_date_format)
        # 复选框控制显示
        self.orig_date_check.toggled.connect(self.orig_date_format.setVisible)
        # 序号标签
        label_seq = QLabel("+【文件序号】+")
        label_seq.setAlignment(Qt.AlignCenter)
        label_seq.setStyleSheet("font-weight: bold; color: #2a5caa;")
        hbox2.addWidget(label_seq)
        # 后缀
        self.orig_suffix_edit = QLineEdit()
        self.orig_suffix_edit.setToolTip("输入后缀，如 '暑假作业'（不含扩展名）")
        self.orig_suffix_edit.setObjectName("orig_suffix_edit")
        hbox2.addWidget(self.orig_suffix_edit)
        # 源文件序号类型
        self.orig_seq_type_combo = QComboBox()
        self.orig_seq_type_combo.addItems(["数字1、2、3", "数字01、02、03", "汉字一、二、三"])
        self.orig_seq_type_combo.setObjectName("orig_seq_type_combo")
        self.orig_seq_type_combo.setToolTip("原文件使用的序号格式")
        hbox2.addWidget(QLabel("原序号类型："))
        hbox2.addWidget(self.orig_seq_type_combo)
        # .docx
        label_docx = QLabel(".docx")
        label_docx.setStyleSheet("color: green; font-weight: bold;")
        hbox2.addWidget(label_docx)
        layout.addLayout(hbox2)

        # ---- 生成文件名格式 ----
        hbox4 = QHBoxLayout()
        hbox4.addWidget(QLabel("生成文件名格式"))
        # 日期（勾选后出现在最前面）
        self.gen_date_check = QCheckBox("日期")
        self.gen_date_check.setObjectName("gen_date_check")
        hbox4.addWidget(self.gen_date_check)
        # 前缀
        self.gen_prefix_edit = QLineEdit()
        self.gen_prefix_edit.setToolTip("输入生成文件的前缀")
        self.gen_prefix_edit.setObjectName("gen_prefix_edit")
        hbox4.addWidget(self.gen_prefix_edit)
        # 日期格式下拉框
        self.gen_date_format = QComboBox()
        self.gen_date_format.addItems([
            "yyyy年MM月dd日", "MM月dd日", "yyyy.MM.dd", "MM.dd",
            "[dd]th,MM(Eng),yyyy", "[dd]th,MM(Eng)", "yyyy-MM-dd", "MM-dd"
        ])
        self.gen_date_format.setCurrentText("mm月dd日")
        self.gen_date_format.setObjectName("gen_date_format")
        self.gen_date_format.setVisible(False)
        hbox4.addWidget(self.gen_date_format)
        self.gen_date_check.toggled.connect(self.gen_date_format.setVisible)
        # 序号标签
        label_seq2 = QLabel("+【文件序号】+")
        label_seq2.setAlignment(Qt.AlignCenter)
        label_seq2.setStyleSheet("font-weight: bold; color: #2a5caa;")
        hbox4.addWidget(label_seq2)
        # 后缀
        self.gen_suffix_edit = QLineEdit()
        self.gen_suffix_edit.setToolTip("输入生成文件的后缀（不含扩展名）")
        self.gen_suffix_edit.setObjectName("gen_suffix_edit")
        hbox4.addWidget(self.gen_suffix_edit)
        # 输出序号类型
        hbox4.addWidget(QLabel("输出序号类型："))
        self.seq_type_combo = QComboBox()
        self.seq_type_combo.addItems(["汉字一、二、三", "数字01、02、03", "数字1、2、3"])
        self.seq_type_combo.setObjectName("seq_type_combo")
        hbox4.addWidget(self.seq_type_combo)
        # .docx
        label_docx2 = QLabel(".docx")
        label_docx2.setStyleSheet("color: green; font-weight: bold;")
        hbox4.addWidget(label_docx2)
        layout.addLayout(hbox4)

        # 学科名称（带帮助图标）
        hbox_subject = QHBoxLayout()
        hbox_subject.addWidget(QLabel("学科名称："))
        help_label = QLabel("ⓘ")
        help_label.setToolTip(
            "这里输入的科目名用于后续设置页码，不会出现在文件名和标题（等下设置）中。\n"
            "如需在文件名和标题中显示学科名，请自行在前后缀输入框中添加。"
        )
        help_label.setStyleSheet("color: #2a5caa; font-size: 14px; font-weight: bold;")
        hbox_subject.addWidget(help_label)
        self.subject_name_edit = QLineEdit("语文")
        self.subject_name_edit.setToolTip("输入学科名称，仅用于页码设置，不参与文件名和标题生成。")
        self.subject_name_edit.setObjectName("subject_name_edit")
        hbox_subject.addWidget(self.subject_name_edit)
        hbox_subject.addStretch()
        layout.addLayout(hbox_subject)
        # ---- 检查按钮 + 状态图标（放在底部提示上方，加大） ----
        hbox_check = QHBoxLayout()
        self.check_btn = QPushButton("检查输入并录入软件")
        self.check_btn.setFixedSize(500, 60)  # 加大
        self.check_btn.clicked.connect(self.checkInput)
        hbox_check.addWidget(self.check_btn)
        self.status_icon = QLabel("×")
        self.status_icon.setFixedSize(50, 50)
        self.status_icon.setAlignment(Qt.AlignCenter)
        self.status_icon.setStyleSheet("color: red; font-size: 20px; font-weight: bold;")
        self.status_icon.setObjectName("status_icon")
        hbox_check.addWidget(self.status_icon)
        hbox_check.addStretch()
        layout.addLayout(hbox_check)
        # 底部提示
        tip = QLabel(
            "设置完此标签后请检查输入是否有效并将文件名设置录入软件，然后转到“标题”标签进行文档内标题的设置。\n"
            "若需另外添加文件或不处理部分文件，请设置完前三个标签页后转到“处理”标签页进行操作。"
        )
        tip.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(tip)
        layout.addStretch()

    def on_folder_changed(self):
        """文件夹路径改变时重置状态图标为红色×"""
        self.setStatusIcon(False)

    def on_sub_features_toggled(self, enabled):
        """控制副标题内部功能控件的显示/隐藏（仅影响内部按钮、日期等，不影响整个副标题组）"""
        self.sub_features_container.setVisible(enabled)

    def on_orig_date_toggled(self, checked):
        """控制原文件名日期相关控件的启用状态"""
        self.orig_date_edit.setEnabled(checked)
        self.orig_date_format.setEnabled(checked)
        if not checked:
            # 取消勾选时，重置状态图标为红色×
            self.setStatusIcon(False)

    def browseFolder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            # 统一使用正斜杠，避免转义问题
            folder = folder.replace('\\', '/')
            self.folder_edit.setText(folder)
            # 自动填充输出目录，使用正斜杠
            output_dir = os.path.join(folder, "已排版文件").replace('\\', '/')
            self.output_folder_edit.setText(output_dir)

    def browseOutputFolder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if folder:
            folder = folder.replace('\\', '/')
            self.output_folder_edit.setText(folder)

    def checkInput(self):
        folder = self.folder_edit.text().strip()
        if not folder:
            QMessageBox.warning(self, "警告", "请先选择文件夹")
            self.setStatusIcon(False)
            return
        if not os.path.isdir(folder):
            QMessageBox.warning(self, "警告", "文件夹不存在")
            self.setStatusIcon(False)
            return
        prefix = self.orig_prefix_edit.text().strip()
        suffix = self.orig_suffix_edit.text().strip()
        seq_type = self.orig_seq_type_combo.currentText()
        seq_strs = self.getSeqListFromFile(seq_type)
        if not seq_strs:
            QMessageBox.warning(self, "警告", "无法读取序号列表")
            self.setStatusIcon(False)
            return
        use_date = self.orig_date_check.isChecked()
        date_pattern = ""
        if use_date:
            fmt = self.orig_date_format.currentText()
            date_pattern = self.build_date_regex(fmt)
            if date_pattern is None:
                QMessageBox.warning(self, "警告", "日期格式错误")
                self.setStatusIcon(False)
                return
        try:
            all_files = os.listdir(folder)
        except Exception as e:
            QMessageBox.warning(self, "警告", f"无法读取文件夹内容：{e}")
            self.setStatusIcon(False)
            return
        found_files = []  # 存储 (fullpath, seq_str)
        for seq_str in seq_strs:
            escaped_prefix = re.escape(prefix)
            escaped_suffix = re.escape(suffix + ".docx")
            if use_date:
                full_pattern = f"^{escaped_prefix}{date_pattern}{re.escape(seq_str)}{escaped_suffix}$"
            else:
                full_pattern = f"^{escaped_prefix}{re.escape(seq_str)}{escaped_suffix}$"
            regex = re.compile(full_pattern)
            for fname in all_files:
                if regex.match(fname):
                    fullpath = os.path.join(folder, fname).replace('\\', '/')
                    found_files.append((fullpath, seq_str))
                    break
        if found_files:
            self.setStatusIcon(True)
            self.process_file_list = found_files  # 存储元组列表
            self.update_process_file_list()
            msg = "找到以下有效文件：\n" + "\n".join([os.path.basename(f[0]) for f in found_files])
            QMessageBox.information(self, "检查结果", msg)
        else:
            self.setStatusIcon(False)
            self.process_file_list = []
            self.update_process_file_list()
            QMessageBox.warning(self, "检查结果", "未找到任何匹配的文件")

    def build_date_regex(self, fmt):
        if fmt == "yyyy年MM月dd日":
            return r"\d{4}年\d{1,2}月\d{1,2}日"
        elif fmt == "MM月dd日":
            return r"\d{1,2}月\d{1,2}日"
        elif fmt == "yyyy.MM.dd":
            return r"\d{4}\.\d{1,2}\.\d{1,2}"
        elif fmt == "MM.dd":
            return r"\d{1,2}\.\d{1,2}"
        elif fmt == "[dd]th,MM(Eng),yyyy":
            months = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
            return rf"\d{{1,2}}(?:st|nd|rd|th),({months}),\d{{4}}"
        elif fmt == "[dd]th,MM(Eng)":
            months = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
            return rf"\d{{1,2}}(?:st|nd|rd|th),({months})"
        elif fmt == "yyyy-MM-dd":
            return r"\d{4}-\d{1,2}-\d{1,2}"
        elif fmt == "MM-dd":
            return r"\d{1,2}-\d{1,2}"
        else:
            return None

    def generate_date_variants(self, date, fmt):
        year = date.year()
        month = date.month()
        day = date.day()
        variants = []
        if fmt == "yyyy年MM月dd日":
            s1 = f"{year:04d}年{month:02d}月{day:02d}日"
            s2 = f"{year:04d}年{month}月{day}日"
            variants = [s1, s2]
        elif fmt == "MM月dd日":
            s1 = f"{month:02d}月{day:02d}日"
            s2 = f"{month}月{day}日"
            variants = [s1, s2]
        elif fmt == "yyyy.MM.dd":
            s1 = f"{year:04d}.{month:02d}.{day:02d}"
            s2 = f"{year:04d}.{month}.{day}"
            variants = [s1, s2]
        elif fmt == "MM.dd":
            s1 = f"{month:02d}.{day:02d}"
            s2 = f"{month}.{day}"
            variants = [s1, s2]
        elif fmt == "[dd]th,MM(Eng),yyyy":
            months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            mon_eng = months[month - 1]

            def day_suffix(d):
                if 4 <= d <= 20 or 24 <= d <= 30:
                    return "th"
                else:
                    return {1: "st", 2: "nd", 3: "rd"}.get(d % 10, "th")

            s1 = f"{day:02d}{day_suffix(day)},{mon_eng},{year:04d}"
            s2 = f"{day}{day_suffix(day)},{mon_eng},{year:04d}"
            variants = [s1, s2]
        elif fmt == "[dd]th,MM(Eng)":
            months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            mon_eng = months[month - 1]

            def day_suffix(d):
                if 4 <= d <= 20 or 24 <= d <= 30:
                    return "th"
                else:
                    return {1: "st", 2: "nd", 3: "rd"}.get(d % 10, "th")

            s1 = f"{day:02d}{day_suffix(day)},{mon_eng}"
            s2 = f"{day}{day_suffix(day)},{mon_eng}"
            variants = [s1, s2]
        elif fmt == "yyyy-MM-dd":
            s1 = f"{year:04d}-{month:02d}-{day:02d}"
            s2 = f"{year:04d}-{month}-{day}"
            variants = [s1, s2]
        elif fmt == "MM-dd":
            s1 = f"{month:02d}-{day:02d}"
            s2 = f"{month}-{day}"
            variants = [s1, s2]
        else:
            variants = [""]
        return list(set(variants))

    def getSeqListFromFile(self, seq_type):
        """根据序号类型从外部文件读取序号列表"""
        if seq_type == "汉字一、二、三":
            fname = "name_list_03"
        elif seq_type == "数字01、02、03":
            fname = "name_list_02"
        elif seq_type == "数字1、2、3":
            fname = "name_list_01"
        else:
            return []
        # 获取当前程序所在目录
        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        file_path = os.path.join(base_dir, fname)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                # 使用 eval 解析列表字符串
                seq_list = eval(content)
                if isinstance(seq_list, list):
                    return seq_list
                else:
                    return []
        except Exception as e:
            print(f"读取序号文件失败：{e}")
            return []

    def setStatusIcon(self, success):
        if success:
            self.status_icon.setText("√")
            self.status_icon.setStyleSheet("color: green; font-size: 20px; font-weight: bold;")
        else:
            self.status_icon.setText("×")
            self.status_icon.setStyleSheet("color: red; font-size: 20px; font-weight: bold;")

    # ==================== 标题标签页 ====================
    def initTitleTab(self, parent):
        layout = QVBoxLayout(parent)
        self.enable_header_cb = QCheckBox("启用题头功能")
        self.enable_header_cb.toggled.connect(self.on_header_enable_toggled)
        self.enable_header_cb.blockSignals(True)
        self.enable_header_cb.setChecked(True)
        self.enable_header_cb.blockSignals(False)
        layout.addWidget(self.enable_header_cb)

        group_header = QGroupBox("题头")
        group_header.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        header_layout = QVBoxLayout(group_header)
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("题头内容："))
        self.header_content_edit = QLineEdit()
        self.header_content_edit.setToolTip("输入题头内容，如'学校2025-2026暑假作业'")
        self.header_content_edit.setObjectName("header_content_edit")
        h1.addWidget(self.header_content_edit)
        header_layout.addLayout(h1)
        self.header_font = QComboBox()
        self.header_font.addItems(QFontDatabase().families())
        self.header_font.setCurrentText("宋体")
        self.header_font.setObjectName("header_font")
        self.header_font_size = QComboBox()
        self.header_font_size.addItems([
            "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
            "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
            "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24",
            "26", "28", "30", "36", "48", "72"
        ])
        self.header_font_size.setCurrentText("三号")
        self.header_font_size.setObjectName("header_font_size")
        self.header_bold = QCheckBox("加粗")
        self.header_bold.setChecked(True)
        self.header_bold.setObjectName("header_bold")
        self.header_underline = QCheckBox("下划线")
        self.header_underline.setObjectName("header_underline")
        self.header_color_btn = QPushButton()
        self.header_color_btn.setFixedSize(30, 20)
        self.header_color_btn.setStyleSheet("background-color: black;")
        self.header_color_btn.clicked.connect(lambda: self.pickColor(self.header_color_btn))
        self.header_color_btn.setObjectName("header_color_btn")
        self.header_align = QComboBox()
        self.header_align.addItems(["左对齐", "右对齐", "居中", "两端对齐", "分散对齐"])
        self.header_align.setCurrentText("居中")
        self.header_align.setObjectName("header_align")
        self.header_line_spacing_type = QComboBox()
        self.header_line_spacing_type.addItems(["n倍行距", "固定值"])
        self.header_line_spacing_type.setObjectName("header_line_spacing_type")
        self.header_line_spacing_value = QDoubleSpinBox()
        self.header_line_spacing_value.setRange(0.1, 999.0)
        self.header_line_spacing_value.setSingleStep(0.1)
        self.header_line_spacing_value.setValue(1.0)
        self.header_line_spacing_value.setObjectName("header_line_spacing_value")
        self.header_line_spacing_unit = QLabel("倍")
        self.header_line_spacing_unit.setObjectName("header_line_spacing_unit")
        self.header_line_spacing_type.currentTextChanged.connect(
            lambda txt: self.header_line_spacing_unit.setText("倍" if txt == "n倍行距" else "磅")
        )
        row1 = QHBoxLayout()
        h_font_group = QHBoxLayout()
        h_font_group.addWidget(QLabel("字体："))
        h_font_group.addWidget(self.header_font)
        row1.addLayout(h_font_group)
        row1.addSpacing(20)
        h_size_group = QHBoxLayout()
        h_size_group.addWidget(QLabel("字号："))
        h_size_group.addWidget(self.header_font_size)
        row1.addLayout(h_size_group)
        row1.addStretch()
        header_layout.addLayout(row1)
        row2 = QHBoxLayout()
        h_part1 = QHBoxLayout()
        h_part1.addWidget(self.header_bold)
        h_part1.addWidget(self.header_underline)
        h_part1.addWidget(QLabel("字体颜色："))
        h_part1.addWidget(self.header_color_btn)
        row2.addLayout(h_part1)
        row2.addSpacing(20)
        h_align_group = QHBoxLayout()
        h_align_group.addWidget(QLabel("对齐方式："))
        h_align_group.addWidget(self.header_align)
        row2.addLayout(h_align_group)
        row2.addSpacing(20)
        h_spacing_group = QHBoxLayout()
        h_spacing_group.addWidget(QLabel("行距："))
        h_spacing_group.addWidget(self.header_line_spacing_type)
        h_spacing_group.addWidget(self.header_line_spacing_value)
        h_spacing_group.addWidget(self.header_line_spacing_unit)
        row2.addLayout(h_spacing_group)
        row2.addStretch()
        header_layout.addLayout(row2)
        self.header_warning = QLabel("")
        self.header_warning.setStyleSheet("color: red;")
        self.header_warning.setVisible(False)
        header_layout.addWidget(self.header_warning)
        layout.addWidget(group_header)

        # 主标题功能开关
        self.enable_main_cb = QCheckBox("启用主标题功能")
        self.enable_main_cb.toggled.connect(self.on_main_enable_toggled)
        self.enable_main_cb.blockSignals(True)
        self.enable_main_cb.setChecked(True)
        self.enable_main_cb.blockSignals(False)
        layout.addWidget(self.enable_main_cb)

        group_main = QGroupBox("主标题")
        group_main.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        main_layout = QVBoxLayout(group_main)
        h_main = QHBoxLayout()
        h_main.addWidget(QLabel("主标题格式："))
        self.main_prefix_edit = QLineEdit()
        self.main_prefix_edit.setToolTip("输入主标题前缀，如'高一级语文'")
        self.main_prefix_edit.setObjectName("main_prefix_edit")
        h_main.addWidget(self.main_prefix_edit)
        label_seq_main = QLabel("+【文件序号】+")
        label_seq_main.setAlignment(Qt.AlignCenter)
        label_seq_main.setStyleSheet("font-weight: bold; color: #2a5caa;")
        h_main.addWidget(label_seq_main)
        self.main_suffix_edit = QLineEdit()
        self.main_suffix_edit.setToolTip("输入主标题后缀，如'暑假作业'")
        self.main_suffix_edit.setObjectName("main_suffix_edit")
        h_main.addWidget(self.main_suffix_edit)
        main_layout.addLayout(h_main)
        self.main_seq_type = QComboBox()
        self.main_seq_type.addItems(["汉字一、二、三", "数字01、02、03", "数字1、2、3"])
        self.main_seq_type.setObjectName("main_seq_type")
        self.main_font = QComboBox()
        self.main_font.addItems(QFontDatabase().families())
        self.main_font.setCurrentText("宋体")
        self.main_font.setObjectName("main_font")
        self.main_font_size = QComboBox()
        self.main_font_size.addItems([
            "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
            "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
            "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24",
            "26", "28", "30", "36", "48", "72"
        ])
        self.main_font_size.setCurrentText("三号")
        self.main_font_size.setObjectName("main_font_size")
        self.main_bold = QCheckBox("加粗")
        self.main_bold.setChecked(True)
        self.main_bold.setObjectName("main_bold")
        self.main_underline = QCheckBox("下划线")
        self.main_underline.setObjectName("main_underline")
        self.main_color_btn = QPushButton()
        self.main_color_btn.setFixedSize(30, 20)
        self.main_color_btn.setStyleSheet("background-color: black;")
        self.main_color_btn.clicked.connect(lambda: self.pickColor(self.main_color_btn))
        self.main_color_btn.setObjectName("main_color_btn")
        self.main_align = QComboBox()
        self.main_align.addItems(["左对齐", "右对齐", "居中", "两端对齐", "分散对齐"])
        self.main_align.setCurrentText("居中")
        self.main_align.setObjectName("main_align")
        self.main_line_spacing_type = QComboBox()
        self.main_line_spacing_type.addItems(["n倍行距", "固定值"])
        self.main_line_spacing_type.setObjectName("main_line_spacing_type")
        self.main_line_spacing_value = QDoubleSpinBox()
        self.main_line_spacing_value.setRange(0.1, 999.0)
        self.main_line_spacing_value.setSingleStep(0.1)
        self.main_line_spacing_value.setValue(1.0)
        self.main_line_spacing_value.setObjectName("main_line_spacing_value")
        self.main_line_spacing_unit = QLabel("倍")
        self.main_line_spacing_unit.setObjectName("main_line_spacing_unit")
        self.main_line_spacing_type.currentTextChanged.connect(
            lambda txt: self.main_line_spacing_unit.setText("倍" if txt == "n倍行距" else "磅")
        )
        row1_main = QHBoxLayout()
        h_seq_group = QHBoxLayout()
        h_seq_group.addWidget(QLabel("序号类型："))
        h_seq_group.addWidget(self.main_seq_type)
        row1_main.addLayout(h_seq_group)
        row1_main.addSpacing(20)
        h_font_group = QHBoxLayout()
        h_font_group.addWidget(QLabel("字体："))
        h_font_group.addWidget(self.main_font)
        row1_main.addLayout(h_font_group)
        row1_main.addSpacing(20)
        h_size_group = QHBoxLayout()
        h_size_group.addWidget(QLabel("字号："))
        h_size_group.addWidget(self.main_font_size)
        row1_main.addLayout(h_size_group)
        row1_main.addStretch()
        main_layout.addLayout(row1_main)
        row2_main = QHBoxLayout()
        h_part1 = QHBoxLayout()
        h_part1.addWidget(self.main_bold)
        h_part1.addWidget(self.main_underline)
        h_part1.addWidget(QLabel("字体颜色："))
        h_part1.addWidget(self.main_color_btn)
        row2_main.addLayout(h_part1)
        row2_main.addSpacing(20)
        h_align_group = QHBoxLayout()
        h_align_group.addWidget(QLabel("对齐方式："))
        h_align_group.addWidget(self.main_align)
        row2_main.addLayout(h_align_group)
        row2_main.addSpacing(20)
        h_spacing_group = QHBoxLayout()
        h_spacing_group.addWidget(QLabel("行距："))
        h_spacing_group.addWidget(self.main_line_spacing_type)
        h_spacing_group.addWidget(self.main_line_spacing_value)
        h_spacing_group.addWidget(self.main_line_spacing_unit)
        row2_main.addLayout(h_spacing_group)
        row2_main.addStretch()
        main_layout.addLayout(row2_main)
        self.main_warning = QLabel("")
        self.main_warning.setStyleSheet("color: red;")
        self.main_warning.setVisible(False)
        main_layout.addWidget(self.main_warning)
        layout.addWidget(group_main)

        # 副标题功能开关（控制整个副标题组显示/隐藏）
        self.enable_sub_cb = QCheckBox("启用副标题功能")
        self.enable_sub_cb.toggled.connect(self.on_sub_enable_toggled)
        self.enable_sub_cb.blockSignals(True)
        self.enable_sub_cb.setChecked(True)
        self.enable_sub_cb.blockSignals(False)
        layout.addWidget(self.enable_sub_cb)

        group_sub = QGroupBox("副标题")
        group_sub.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        sub_layout = QVBoxLayout(group_sub)

        # ---- 1. 出题人、审题人录入 ----
        self.enable_examiner_cb = QCheckBox("启用出题人、审题人录入")
        self.enable_examiner_cb.setChecked(True)
        self.enable_examiner_cb.toggled.connect(lambda checked: self.examiner_container.setVisible(checked))
        sub_layout.addWidget(self.enable_examiner_cb)

        self.examiner_container = QWidget()
        examiner_layout = QVBoxLayout(self.examiner_container)
        examiner_layout.setContentsMargins(20, 0, 0, 0)  # 缩进
        btn_examiner = QPushButton("出题人、审题人录入")
        btn_examiner.clicked.connect(self.openExaminerDialog)
        examiner_layout.addWidget(btn_examiner)
        sub_layout.addWidget(self.examiner_container)

        # ---- 2. 日期设置 ----
        self.enable_date_cb = QCheckBox("启用日期设置")
        self.enable_date_cb.setChecked(True)
        self.enable_date_cb.toggled.connect(lambda checked: self.date_container.setVisible(checked))
        sub_layout.addWidget(self.enable_date_cb)

        self.date_container = QWidget()
        date_layout = QVBoxLayout(self.date_container)
        date_layout.setContentsMargins(20, 0, 0, 0)

        h_date = QHBoxLayout()
        h_start_group = QHBoxLayout()
        h_start_group.addWidget(QLabel("开始日期："))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setObjectName("start_date")
        h_start_group.addWidget(self.start_date)
        h_date.addLayout(h_start_group)
        h_date.addSpacing(20)
        h_end_group = QHBoxLayout()
        h_end_group.addWidget(QLabel("结束日期："))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate().addDays(7))
        self.end_date.setObjectName("end_date")
        h_end_group.addWidget(self.end_date)
        h_date.addLayout(h_end_group)
        h_date.addStretch()
        date_layout.addLayout(h_date)

        h_date_check = QHBoxLayout()
        self.date_check_btn = QPushButton("检查日期")
        self.date_check_btn.setFixedSize(120, 38)
        self.date_check_btn.setStyleSheet("font-size: 16px; font-weight: bold; background-color: #4CAF50; color: white; border-radius: 6px;")
        self.date_check_btn.clicked.connect(self.checkDates)
        h_date_check.addWidget(self.date_check_btn)
        self.date_status_icon = QLabel("×")
        self.date_status_icon.setFixedSize(20, 20)
        self.date_status_icon.setAlignment(Qt.AlignCenter)
        self.date_status_icon.setStyleSheet("color: red; font-size: 16px; font-weight: bold;")
        self.date_status_icon.setObjectName("date_status_icon")
        h_date_check.addWidget(self.date_status_icon)
        h_date_check.addStretch()
        date_layout.addLayout(h_date_check)

        sub_layout.addWidget(self.date_container)

        # ---- 3. 日期计算方式 ----
        h_cycle = QHBoxLayout()
        h_cycle.addWidget(QLabel("日期计算方式："))
        self.cycle_type_combo = QComboBox()
        self.cycle_type_combo.addItems(["按正常工作日安排", "自定义"])
        self.cycle_type_combo.setObjectName("cycle_type_combo")
        h_cycle.addWidget(self.cycle_type_combo)
        h_cycle.addStretch()
        sub_layout.addLayout(h_cycle)
        # ---- 4. 休息设置 ----
        self.rest_container = QWidget()
        rest_layout = QVBoxLayout(self.rest_container)
        rest_layout.setContentsMargins(20, 0, 0, 0)
        h_rest = QHBoxLayout()
        h_rest.setSpacing(0)
        h_rest.addWidget(QLabel("休息设置：做"))
        self.rest_do = QSpinBox()
        self.rest_do.setRange(0, 999)
        self.rest_do.setValue(0)
        self.rest_do.setFixedWidth(45)
        self.rest_do.setObjectName("rest_do")
        h_rest.addWidget(self.rest_do)
        h_rest.addWidget(QLabel("天，休"))
        self.rest_rest = QSpinBox()
        self.rest_rest.setRange(0, 999)
        self.rest_rest.setValue(0)
        self.rest_rest.setFixedWidth(45)
        self.rest_rest.setObjectName("rest_rest")
        h_rest.addWidget(self.rest_rest)
        h_rest.addWidget(QLabel("天"))
        h_rest.addStretch()
        rest_layout.addLayout(h_rest)
        self.rest_warning = QLabel("")
        self.rest_warning.setStyleSheet("color: red;")
        self.rest_warning.setVisible(False)
        rest_layout.addWidget(self.rest_warning)
        sub_layout.addWidget(self.rest_container)
        # 根据日期计算方式显示/隐藏休息设置
        def on_cycle_changed(mode):
            show = (mode == "自定义")
            self.rest_container.setVisible(show)
            self.cycle_type_combo.currentTextChanged.connect(on_cycle_changed)
        on_cycle_changed(self.cycle_type_combo.currentText())
        # 对齐方式
        hbox_sub_align = QHBoxLayout()
        hbox_sub_align.setSpacing(0)
        hbox_sub_align.addWidget(QLabel("对齐方式："))
        self.sub_align = QComboBox()
        hbox_sub_align.addWidget(self.sub_align)
        self.sub_align.addItems(["左对齐", "右对齐", "居中", "两端对齐", "分散对齐"])
        self.sub_align.setCurrentText("居中")
        self.sub_align.setObjectName("sub_align")
        hbox_sub_align.addStretch()
        sub_layout.addLayout(hbox_sub_align)

        # 字体设置
        hbox_sub_font = QHBoxLayout()
        h_sub_font_group = QHBoxLayout()
        h_sub_font_group.addWidget(QLabel("字体："))
        self.sub_font = QComboBox()
        self.sub_font.addItems(QFontDatabase().families())
        self.sub_font.setCurrentText("\u5b8b\u4f53")
        self.sub_font.setObjectName("sub_font")
        h_sub_font_group.addWidget(self.sub_font)
        hbox_sub_font.addLayout(h_sub_font_group)
        hbox_sub_font.addSpacing(20)
        h_sub_size_group = QHBoxLayout()
        h_sub_size_group.addWidget(QLabel("字号："))
        self.sub_font_size = QComboBox()
        self.sub_font_size.addItems([
            "\u521d\u53f7", "\u5c0f\u521d", "\u4e00\u53f7", "\u5c0f\u4e00", "\u4e8c\u53f7", "\u5c0f\u4e8c",
            "\u4e09\u53f7", "\u5c0f\u4e09", "\u56db\u53f7", "\u5c0f\u56db", "\u4e94\u53f7", "\u5c0f\u4e94",
            "\u516d\u53f7", "\u5c0f\u516d", "\u4e03\u53f7", "\u516b\u53f7",
            "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24",
            "26", "28", "30", "36", "48", "72"
        ])
        self.sub_font_size.setCurrentText("\u56db\u53f7")
        self.sub_font_size.setObjectName("sub_font_size")
        h_sub_size_group.addWidget(self.sub_font_size)
        hbox_sub_font.addLayout(h_sub_size_group)
        hbox_sub_font.addSpacing(20)
        self.sub_bold = QCheckBox("\u52a0\u7c97")
        self.sub_bold.setObjectName("sub_bold")
        self.sub_underline = QCheckBox("\u4e0b\u5212\u7ebf")
        self.sub_underline.setObjectName("sub_underline")
        self.sub_color_btn = QPushButton()
        self.sub_color_btn.setFixedSize(30, 20)
        self.sub_color_btn.setStyleSheet("background-color: black;")
        self.sub_color_btn.clicked.connect(lambda: self.pickColor(self.sub_color_btn))
        self.sub_color_btn.setObjectName("sub_color_btn")
        h_sub_extra_group = QHBoxLayout()
        h_sub_extra_group.addWidget(self.sub_bold)
        h_sub_extra_group.addWidget(self.sub_underline)
        h_sub_extra_group.addWidget(QLabel("\u5b57\u4f53\u989c\u8272\uff1a"))
        h_sub_extra_group.addWidget(self.sub_color_btn)
        hbox_sub_font.addLayout(h_sub_extra_group)
        hbox_sub_font.addStretch()
        sub_layout.addLayout(hbox_sub_font)

        self.sub_line_spacing_container = QWidget()
        sub_line_spacing_layout = QVBoxLayout(self.sub_line_spacing_container)
        sub_line_spacing_layout.setContentsMargins(20, 0, 0, 0)

        sub_line_layout = QHBoxLayout()
        sub_line_layout.addWidget(QLabel("行距："))
        self.sub_line_spacing_type = QComboBox()
        self.sub_line_spacing_type.addItems(["n倍行距", "固定值"])
        self.sub_line_spacing_type.setObjectName("sub_line_spacing_type")
        sub_line_layout.addWidget(self.sub_line_spacing_type)
        self.sub_line_spacing_value = QDoubleSpinBox()
        self.sub_line_spacing_value.setRange(0.1, 999.0)
        self.sub_line_spacing_value.setSingleStep(0.1)
        self.sub_line_spacing_value.setValue(1.0)
        self.sub_line_spacing_value.setObjectName("sub_line_spacing_value")
        sub_line_layout.addWidget(self.sub_line_spacing_value)
        self.sub_line_spacing_unit = QLabel("倍")
        self.sub_line_spacing_unit.setObjectName("sub_line_spacing_unit")
        sub_line_layout.addWidget(self.sub_line_spacing_unit)
        self.sub_line_spacing_type.currentTextChanged.connect(
            lambda txt: self.sub_line_spacing_unit.setText("倍" if txt == "n倍行距" else "磅")
        )
        sub_line_layout.addStretch()
        sub_line_spacing_layout.addLayout(sub_line_layout)
        self.sub_warning = QLabel("")
        self.sub_warning.setStyleSheet("color: red;")
        self.sub_warning.setVisible(False)
        sub_line_spacing_layout.addWidget(self.sub_warning)

        sub_layout.addWidget(self.sub_line_spacing_container)

        layout.addWidget(group_sub)

        tip = QLabel(
            "设置完此标签后请检查输入是否有效，点击\u201c检查日期\u201d按钮，然后转到\u201c页面\u201d标签进行页面布局设置。")
        tip.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(tip)
        layout.addStretch()

    # ==================== 页面标签页 ====================
    def initPageTab(self, parent):
        layout = QVBoxLayout(parent)
        # 页边距
        group_margin = QGroupBox("页边距设置")
        group_margin.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        margin_layout = QHBoxLayout(group_margin)
        self.margin_top = QDoubleSpinBox()
        self.margin_top.setRange(0, 50)
        self.margin_top.setValue(2.3)
        self.margin_top.setSingleStep(0.1)
        self.margin_top.setObjectName("margin_top")
        self.margin_bottom = QDoubleSpinBox()
        self.margin_bottom.setRange(0, 50)
        self.margin_bottom.setValue(2.3)
        self.margin_bottom.setSingleStep(0.1)
        self.margin_bottom.setObjectName("margin_bottom")
        self.margin_left = QDoubleSpinBox()
        self.margin_left.setRange(0, 50)
        self.margin_left.setValue(2.1)
        self.margin_left.setSingleStep(0.1)
        self.margin_left.setObjectName("margin_left")
        self.margin_right = QDoubleSpinBox()
        self.margin_right.setRange(0, 50)
        self.margin_right.setValue(2.1)
        self.margin_right.setSingleStep(0.1)
        self.margin_right.setObjectName("margin_right")
        margin_layout.addWidget(QLabel("上"))
        margin_layout.addWidget(self.margin_top)
        margin_layout.addWidget(QLabel("cm"))
        margin_layout.addSpacing(20)
        margin_layout.addWidget(QLabel("下"))
        margin_layout.addWidget(self.margin_bottom)
        margin_layout.addWidget(QLabel("cm"))
        margin_layout.addSpacing(20)
        margin_layout.addWidget(QLabel("左"))
        margin_layout.addWidget(self.margin_left)
        margin_layout.addWidget(QLabel("cm"))
        margin_layout.addSpacing(20)
        margin_layout.addWidget(QLabel("右"))
        margin_layout.addWidget(self.margin_right)
        margin_layout.addWidget(QLabel("cm"))
        margin_layout.addStretch()
        layout.addWidget(group_margin)
        # 页眉页脚
        group_header_footer = QGroupBox("页眉设置")
        group_header_footer.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        hf_layout = QVBoxLayout(group_header_footer)
        h_header = QHBoxLayout()
        h_header.addWidget(QLabel("页眉位置："))
        self.header_mode = QComboBox()
        self.header_mode.addItems(["无页眉", "左页眉", "右页眉"])
        self.header_mode.setObjectName("header_mode")
        h_header.addWidget(self.header_mode)
        h_header.addSpacing(15)
        self.header_text_label = QLabel("页眉文字：")
        h_header.addWidget(self.header_text_label)
        self.header_text_edit = QLineEdit()
        self.header_text_edit.setObjectName("header_text_edit")
        h_header.addWidget(self.header_text_edit)
        h_header.addStretch()
        hf_layout.addLayout(h_header)
        self.header_mode.currentTextChanged.connect(self._on_header_mode_changed)
        self._on_header_mode_changed(self.header_mode.currentText())
        h_page = QHBoxLayout()
        self.insert_page_check = QCheckBox("插入页码")
        self.insert_page_check.setChecked(True)
        self.insert_page_check.setObjectName("insert_page_check")
        h_page.addWidget(self.insert_page_check)
        self.page_position = QComboBox()
        self.page_position.addItems(["居中", "靠左", "靠右"])
        self.page_position.setCurrentText("居中")
        self.page_position.setObjectName("page_position")
        h_page.addWidget(QLabel("页码位置："))
        h_page.addWidget(self.page_position)
        h_page.addStretch()
        hf_layout.addLayout(h_page)
        self.page_settings_group = QFrame()
        self.page_settings_group.setFrameShape(QFrame.StyledPanel)
        page_set_layout = QVBoxLayout(self.page_settings_group)
        h_style = QHBoxLayout()
        h_style.addWidget(QLabel("页码样式："))
        self.page_style = QComboBox()
        self.page_style.addItems([
            "科目  第x页  共y页", "科目  x/y", "科目  x",
            "第x页  共y页", "x/y", "x"
        ])
        self.page_style.setCurrentIndex(0)
        self.page_style.setObjectName("page_style")
        h_style.addWidget(self.page_style)
        h_style.addSpacing(30)
        self.page_font = QComboBox()
        self.page_font.addItems(QFontDatabase().families())
        self.page_font.setCurrentText("宋体")
        self.page_font.setObjectName("page_font")
        h_style.addWidget(QLabel("字体："))
        h_style.addWidget(self.page_font)
        self.page_font_size = QComboBox()
        self.page_font_size.addItems([
            "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
            "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
            "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24",
            "26", "28", "30", "36", "48", "72"
        ])
        self.page_font_size.setCurrentText("五号")
        self.page_font_size.setObjectName("page_font_size")
        h_style.addWidget(QLabel("字号："))
        h_style.addWidget(self.page_font_size)
        h_style.addStretch()
        page_set_layout.addLayout(h_style)
        h_page_format = QHBoxLayout()
        self.page_bold = QCheckBox("加粗")
        self.page_bold.setObjectName("page_bold")
        self.page_underline = QCheckBox("下划线")
        self.page_underline.setObjectName("page_underline")
        self.page_italic = QCheckBox("斜体")
        self.page_italic.setObjectName("page_italic")
        h_page_format.addWidget(self.page_bold)
        h_page_format.addWidget(self.page_underline)
        h_page_format.addWidget(self.page_italic)
        h_page_format.addWidget(QLabel("字体颜色："))
        self.page_color_btn = QPushButton()
        self.page_color_btn.setFixedSize(30, 20)
        self.page_color_btn.setStyleSheet("background-color: black;")
        self.page_color_btn.clicked.connect(lambda: self.pickColor(self.page_color_btn))
        self.page_color_btn.setObjectName("page_color_btn")
        h_page_format.addWidget(self.page_color_btn)
        h_page_format.addStretch()
        page_set_layout.addLayout(h_page_format)
        hf_layout.addWidget(self.page_settings_group)
        # 页眉页脚字体设置
        hf_font_layout = QHBoxLayout()
        hf_font_layout.addWidget(QLabel("页眉页脚字体："))
        self.header_footer_font = QComboBox()
        self.header_footer_font.addItems(QFontDatabase().families())
        self.header_footer_font.setCurrentText("宋体")
        self.header_footer_font.setObjectName("header_footer_font")
        hf_font_layout.addWidget(self.header_footer_font)
        hf_font_layout.addSpacing(15)
        hf_font_layout.addWidget(QLabel("字号："))
        self.header_footer_font_size = QComboBox()
        self.header_footer_font_size.addItems([
            "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
            "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
            "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24",
            "26", "28", "30", "36", "48", "72"
        ])
        self.header_footer_font_size.setCurrentText("五号")
        self.header_footer_font_size.setObjectName("header_footer_font_size")
        hf_font_layout.addWidget(self.header_footer_font_size)
        hf_font_layout.addSpacing(15)
        self.header_footer_bold = QCheckBox("加粗")
        self.header_footer_bold.setObjectName("header_footer_bold")
        hf_font_layout.addWidget(self.header_footer_bold)
        hf_font_layout.addStretch()
        hf_layout.addLayout(hf_font_layout)
        self.insert_page_check.toggled.connect(self.on_insert_page_toggled)
        self.page_position.currentTextChanged.connect(self.on_page_position_changed)
        self.on_insert_page_toggled(self.insert_page_check.isChecked())
        self.on_page_position_changed(self.page_position.currentText())
        layout.addWidget(group_header_footer)
        # 纸张
        group_paper = QGroupBox("纸张大小")
        group_paper.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        paper_layout = QHBoxLayout(group_paper)
        paper_layout.addWidget(QLabel("纸张大小："))
        self.paper_size = QComboBox()
        paper_sizes = [
            ("A4", "210×297 mm"),
            ("A3", "297×420 mm"),
            ("A5", "148×210 mm"),
            ("B4", "257×364 mm"),
            ("B5", "182×257 mm"),
            ("Letter", "216×279 mm"),
            ("Legal", "216×356 mm"),
            ("Executive", "184×267 mm"),
            ("16开", "195×270 mm"),
        ]
        for name, size in paper_sizes:
            self.paper_size.addItem(f"{name} ({size})")
        self.paper_size.setCurrentIndex(0)
        self.paper_size.setObjectName("paper_size")
        paper_layout.addWidget(self.paper_size)
        paper_layout.addStretch()
        layout.addWidget(group_paper)
        # 题目行距
        group_question_spacing = QGroupBox("题目行距")
        group_question_spacing.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        qs_layout = QHBoxLayout(group_question_spacing)
        qs_layout.addWidget(QLabel("行距："))
        self.question_line_spacing_type = QComboBox()
        self.question_line_spacing_type.addItems(["n倍行距", "固定值"])
        self.question_line_spacing_type.setObjectName("question_line_spacing_type")
        qs_layout.addWidget(self.question_line_spacing_type)
        self.question_line_spacing_value = QDoubleSpinBox()
        self.question_line_spacing_value.setRange(0.1, 999.0)
        self.question_line_spacing_value.setSingleStep(0.1)
        self.question_line_spacing_value.setValue(1.0)
        self.question_line_spacing_value.setObjectName("question_line_spacing_value")
        qs_layout.addWidget(self.question_line_spacing_value)
        self.question_line_spacing_unit = QLabel("倍")
        self.question_line_spacing_unit.setObjectName("question_line_spacing_unit")
        qs_layout.addWidget(self.question_line_spacing_unit)
        self.question_line_spacing_type.currentTextChanged.connect(
            lambda txt: self.question_line_spacing_unit.setText("倍" if txt == "n倍行距" else "磅")
        )
        qs_layout.addStretch()
        layout.addWidget(group_question_spacing)
        # 正文字体强制覆盖
        hbox_enforce = QHBoxLayout()
        self.enforce_body_font = QCheckBox("强制覆盖正文字体（中文宋体、英文Times New Roman）")
        self.enforce_body_font.setChecked(True)
        self.enforce_body_font.setObjectName("enforce_body_font")
        hbox_enforce.addWidget(self.enforce_body_font)
        hbox_enforce.addStretch()
        layout.addLayout(hbox_enforce)
        tip = QLabel("设置完此标签后请检查输入和选择是否有效和符合要求，然后转到“处理”标签进行文件处理。")
        tip.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(tip)
        layout.addStretch()

    def on_insert_page_toggled(self, checked):
        self.page_settings_group.setVisible(checked)
        self.page_position.setVisible(checked)
        if not checked:
            self.page_position.setCurrentText("居中")

    def on_page_position_changed(self, position):
        pass

    def _on_header_mode_changed(self, mode):
        has_header = mode != '无页眉'
        self.header_text_label.setVisible(has_header)
        self.header_text_edit.setVisible(has_header)
        if mode == '无页眉':
            self.header_text_edit.clear()

    # ==================== 处理标签页 ====================

    def initProcessTab(self, parent):
        layout = QVBoxLayout(parent)
        tip = QLabel(
            "请选择你要处理的文件，确保现有文件中的内容都只有题目，然后调整好参数，最后点击“一键排版”按钮以进行文件排版。")
        tip.setStyleSheet("color: blue;")
        layout.addWidget(tip)
        # 表格：列：文件名、页数、修改日期、文件大小、文件路径
        self.process_table = QTableWidget()
        self.process_table.setColumnCount(5)
        self.process_table.setHorizontalHeaderLabels(["文件名", "页数", "修改日期", "文件大小", "文件路径"])
        # 设置列宽：文件名和文件路径自适应，其他固定
        self.process_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.process_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.process_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.process_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.process_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)  # 文件路径
        # 禁用编辑
        self.process_table.setEditTriggers(QTableWidget.NoEditTriggers)
        # 选择行为：单选
        self.process_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.process_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.process_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.process_table.setObjectName("process_table")
        layout.addWidget(self.process_table)
        # 启用水平滚动条
        self.process_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.process_table.setObjectName("process_table")
        layout.addWidget(self.process_table)
        # 操作按钮
        hbox_buttons = QHBoxLayout()
        self.move_up_btn = QPushButton("上移")
        self.move_up_btn.clicked.connect(self.move_up)
        self.move_down_btn = QPushButton("下移")
        self.move_down_btn.clicked.connect(self.move_down)
        self.add_file_btn = QPushButton("添加文件")
        self.add_file_btn.clicked.connect(self.add_files)
        hbox_buttons.addWidget(self.add_file_btn)
        self.delete_btn = QPushButton("删除所选")
        self.delete_btn.clicked.connect(self.delete_selected)
        self.clear_btn = QPushButton("清空")
        self.clear_btn.clicked.connect(self.clear_list)
        hbox_buttons.addWidget(self.move_up_btn)
        hbox_buttons.addWidget(self.move_down_btn)
        hbox_buttons.addWidget(self.add_file_btn)
        hbox_buttons.addWidget(self.delete_btn)
        hbox_buttons.addWidget(self.clear_btn)
        hbox_buttons.addStretch()
        layout.addLayout(hbox_buttons)
        # ---- 第二个生成文件列表 ----
        # 复选框：设置并检查生成文件的期望页数
        hbox_gen_check = QHBoxLayout()
        self.gen_pages_check = QCheckBox("设置并检查生成文件的期望页数")
        self.gen_pages_check.setChecked(True)
        self.gen_pages_check.setObjectName("gen_pages_check")
        self.gen_pages_check.toggled.connect(self.on_gen_pages_check_toggled)
        hbox_gen_check.addWidget(self.gen_pages_check)
        hbox_gen_check.addStretch()
        layout.addLayout(hbox_gen_check)
        # 提示文字（动态变化）
        self.gen_label = QLabel("下方显示的是生成文件，请填写文件期望页数并检查文件名是否符合要求：")
        self.gen_label.setObjectName("gen_label")
        layout.addWidget(self.gen_label)
        # 一键应用按钮
        hbox_apply = QHBoxLayout()
        self.apply_btn = QPushButton("一键应用第一项页数数据")
        self.apply_btn.clicked.connect(self.apply_first_pages)
        hbox_apply.addWidget(self.apply_btn)
        hbox_apply.addStretch()
        layout.addLayout(hbox_apply)
        # 滚动区域用于生成文件列表
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        self.gen_file_container = QWidget()
        self.gen_file_layout = QVBoxLayout(self.gen_file_container)
        scroll_layout.addWidget(self.gen_file_container)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        # 排版日志
        log_group = QGroupBox("排版日志")
        log_group.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 10px; }")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("log_text")
        self.log_text.setMaximumHeight(160)
        self.log_text.setStyleSheet("background-color: #f5f5f5; font-family: Consolas, 'Courier New', monospace; font-size: 12px;")
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)
        # 输出格式
        hbox_format = QHBoxLayout()
        hbox_format.addWidget(QLabel("输出格式："))
        self.output_format = QComboBox()
        self.output_format.addItems(["Microsoft Word 2007 文档（.docx）", "Portable Document Format（.pdf）", "两者都输出"])
        self.output_format.setObjectName("output_format")
        hbox_format.addWidget(self.output_format)
        hbox_format.addStretch()
        layout.addLayout(hbox_format)
        # 保存预设按钮
        hbox_preset = QHBoxLayout()
        save_preset_btn = QPushButton("保存当前预设")
        save_preset_btn.clicked.connect(self.savePreset)
        save_preset_btn.setObjectName("save_preset_btn")
        hbox_preset.addStretch()
        hbox_preset.addWidget(save_preset_btn)
        hbox_preset.addStretch()
        layout.addLayout(hbox_preset)
        # 一键排版按钮
        self.process_btn = QPushButton("一键排版")
        self.process_btn.setFixedSize(200, 40)
        self.process_btn.clicked.connect(self.start_process)
        hbox_process = QHBoxLayout()
        hbox_process.addStretch()
        hbox_process.addWidget(self.process_btn)
        hbox_process.addStretch()
        layout.addLayout(hbox_process)
        layout.addStretch()
        self.move_up_btn.setEnabled(False)
        self.move_down_btn.setEnabled(False)
        self.delete_btn.setEnabled(False)
        self.update_button_state()
        self.process_table.itemSelectionChanged.connect(self.update_button_state)
        self.process_table.itemChanged.connect(self.on_table_item_changed)  # 连接 itemChanged 信号，监听复选框变化
        # 确保初始状态
        self.process_table.setColumnHidden(1, False)
        self.gen_pages_check.setEnabled(True)
        self.gen_pages_check.setChecked(True)

    def seq_int_to_str(self, num):
        """将整数转换为当前序号类型的字符串"""
        seq_type = self.seq_type_combo.currentText()
        if seq_type == "数字01、02、03":
            return f"{num:02d}"
        elif seq_type == "数字1、2、3":
            return str(num)
        elif seq_type == "汉字一、二、三":
            seq_list = self.getSeqListFromFile(seq_type)
            if num <= len(seq_list):
                return seq_list[num - 1]
            else:
                return f"_{num}"
        else:
            return f"_{num}"

    def add_files(self):
        """添加额外文件到列表"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择要添加的文件", "", "Word文档 (*.docx)"
        )
        if not files:
            return
        seq_type = self.orig_seq_type_combo.currentText()

        def extract_seq_from_filename(fname):
            base = os.path.splitext(os.path.basename(fname))[0]
            if seq_type == "数字01、02、03":
                match = re.search(r'(\d{2})$', base)
                if match:
                    return match.group(1)
            elif seq_type == "数字1、2、3":
                match = re.search(r'(\d+)$', base)
                if match:
                    return match.group(1)
            elif seq_type == "汉字一、二、三":
                seq_list = self.getSeqListFromFile(seq_type)
                for seq in seq_list:
                    if base.endswith(seq):
                        return seq
            return None

        seqs = []
        unknown = []
        for f in files:
            seq = extract_seq_from_filename(f)
            if seq is not None:
                seqs.append(seq)
            else:
                unknown.append(f)
        # 获取当前最大序号（用于接着）
        max_seq = self.get_max_seq_from_table()
        files = [f.replace('\\', '/') for f in files]
        # 处理未知文件
        if unknown:
            reply = QMessageBox.question(
                self, "序号选择",
                f"以下 {len(unknown)} 个文件未能自动提取序号：\n" + "\n".join([os.path.basename(f) for f in unknown]) +
                "\n\n请选择序号分配方式：",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                # 接着列表中的序号
                for i, f in enumerate(unknown):
                    seq_str = self.seq_int_to_str(max_seq + i + 1)
                    seqs.append(seq_str)
            else:
                # 自行输入序号
                self.show_manual_seq_dialog(unknown, seqs, max_seq)
                if len(seqs) != len(unknown):
                    return  # 取消或失败
        else:
            # 所有文件都有序号，直接添加
            pass
        if len(seqs) != len(files):
            QMessageBox.warning(self, "警告", "序号数量与文件数量不匹配，操作取消。")
            return
        # 将新文件添加到 process_file_list
        for f, seq in zip(files, seqs):
            self.process_file_list.append((f, seq))
        # 刷新表格
        self.update_process_file_list()

    def get_max_seq_from_table(self):
        """从当前表格中获取最大的序号（用于接着序号）"""
        max_val = 0
        for item in self.process_file_list:
            seq_str = item[1]  # 序号字符串
            # 转换为数字
            if seq_str.isdigit():
                val = int(seq_str)
                if val > max_val:
                    max_val = val
            else:
                # 如果是汉字，尝试从 name_list_03 中获取索引
                seq_list = self.getSeqListFromFile("汉字一、二、三")
                if seq_str in seq_list:
                    val = seq_list.index(seq_str) + 1
                    if val > max_val:
                        max_val = val
        return max_val

    def show_manual_seq_dialog(self, files, seqs):
        """显示手动输入序号的对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle("手动输入序号")
        dialog.resize(500, 400)
        layout = QVBoxLayout(dialog)
        label = QLabel("请为每个文件输入序号：")
        layout.addWidget(label)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        spinboxes = []
        for f in files:
            hbox = QHBoxLayout()
            lbl = QLabel(os.path.basename(f))
            hbox.addWidget(lbl)
            spin = QSpinBox()
            spin.setRange(1, 999)
            spin.setValue(1)  # 默认值
            spinboxes.append(spin)
            hbox.addWidget(spin)
            hbox.addStretch()
            scroll_layout.addLayout(hbox)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addWidget(btn_box)

    def show_manual_seq_dialog(self, files, seqs, max_seq):
        """显示手动输入序号的对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle("手动输入序号")
        dialog.resize(500, 400)
        layout = QVBoxLayout(dialog)
        label = QLabel("请为每个文件输入序号：")
        layout.addWidget(label)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        spinboxes = []
        for i, f in enumerate(files):
            hbox = QHBoxLayout()
            lbl = QLabel(os.path.basename(f))
            hbox.addWidget(lbl)
            hbox.addWidget(QLabel("序号："))  # 新增标签
            spin = QSpinBox()
            spin.setRange(1, 999)
            spin.setValue(max_seq + i + 1)  # 默认接着序号
            spinboxes.append(spin)
            hbox.addWidget(spin)
            hbox.addStretch()
            scroll_layout.addLayout(hbox)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addWidget(btn_box)
        if dialog.exec_() == QDialog.Accepted:
            seq_type = self.seq_type_combo.currentText()
            for spin in spinboxes:
                val = spin.value()
                seq_str = self.seq_int_to_str(val)
                seqs.append(seq_str)
        else:
            # 取消，清空 seqs 以便上层判断
            seqs.clear()

    def on_add_files_ready(self, info_list):
        # 将新文件信息追加到 process_file_list 和 process_file_info
        if not hasattr(self, 'process_file_list') or self.process_file_list is None:
            self.process_file_list = []
            self.process_file_info = []
        # 构造新的序号（从当前最大序号+1，但需要保持格式）
        # 从 self.process_file_list 中获取已有的序号（如果有）
        existing_seqs = [item[1] for item in self.process_file_list if len(item) > 1 and item[1]]
        # 尝试从新文件名中提取序号（例如 "01" 或 "一"），如果没有则使用 "99" 占位
        # 为简化，直接使用 "99" 占位，用户可自行调整
        for fullpath, pages, mtime, size, error_msg in info_list:
            # 尝试从文件名中提取数字序号（如 "03"）
            basename = os.path.basename(fullpath)
            # 使用正则提取数字部分
            import re
            match = re.search(r'(\d{2})', basename)
            if match:
                seq_str = match.group(1)
            else:
                seq_str = "99"  # 默认占位
            # 添加到列表
            self.process_file_list.append((fullpath, seq_str))
            self.process_file_info.append((fullpath, pages, mtime, size, error_msg))
        # 重新加载整个表格（使用 update_process_file_list 会重新获取所有文件信息，但我们已经有 info 了，直接重建表格）
        # 为了简单，我们直接重新构建表格
        self.process_table.setRowCount(0)
        # 重新设置垂直表头并填入数据
        for idx, (fullpath, pages, mtime, size, error_msg) in enumerate(self.process_file_info, start=1):
            row = self.process_table.rowCount()
            self.process_table.insertRow(row)
            header_item = QTableWidgetItem(f"{idx:02d}")
            self.process_table.setVerticalHeaderItem(row, header_item)
            # 获取序号
            seq_str = self.process_file_list[idx - 1][1] if len(self.process_file_list) >= idx else f"{idx:02d}"
            filename = os.path.basename(fullpath)
            file_item = QTableWidgetItem(filename)
            file_item.setCheckState(Qt.Checked)
            file_item.setData(Qt.UserRole, fullpath)
            file_item.setData(Qt.UserRole + 1, seq_str)
            if error_msg:
                file_item.setForeground(Qt.red)
                file_item.setCheckState(Qt.Unchecked)
            file_item.setToolTip(fullpath)
            self.process_table.setItem(row, 0, file_item)
            pages_str = "错误" if error_msg else str(pages)
            pages_item = QTableWidgetItem(pages_str)
            pages_item.setTextAlignment(Qt.AlignCenter)
            self.process_table.setItem(row, 1, pages_item)
            if mtime:
                mtime_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            else:
                mtime_str = "未知"
            mtime_item = QTableWidgetItem(mtime_str)
            mtime_item.setTextAlignment(Qt.AlignCenter)
            self.process_table.setItem(row, 2, mtime_item)
            if size:
                size_kb = size / 1024
                size_str = f"{size_kb / 1024:.2f} MB" if size_kb > 1024 else f"{size_kb:.1f} KB"
            else:
                size_str = "未知"
            size_item = QTableWidgetItem(size_str)
            size_item.setTextAlignment(Qt.AlignCenter)
            self.process_table.setItem(row, 3, size_item)
        self.update_button_state()
        self.update_gen_file_list()
        QMessageBox.information(self, "添加完成", f"成功添加 {len(info_list)} 个文件。")

    def on_table_item_changed(self, item):
        """当表格项发生变化时，如果变化的是第0列（文件名列），则更新生成文件列表"""
        if item.column() == 0:
            self._update_error_state()
            self.update_gen_file_list()

    def update_process_file_list(self):
        self.process_table.setRowCount(0)
        if not self.process_file_list:
            # 【新增】无文件时重置列可见和复选框
            self.process_table.setColumnHidden(1, False)  # 显示页数列
            self.gen_pages_check.setEnabled(True)
            self.gen_pages_check.setChecked(True)  # 默认勾选
            self.update_gen_file_list()
            return
        file_paths = [item[0] for item in self.process_file_list]
        self.file_seq_map = {item[0]: item[1] for item in self.process_file_list}
        row = self.process_table.rowCount()
        self.process_table.insertRow(row)
        self.process_table.setVerticalHeaderItem(row, QTableWidgetItem(""))
        item = QTableWidgetItem("正在获取文件数据，请稍等……")
        item.setTextAlignment(Qt.AlignCenter)
        self.process_table.setItem(row, 0, item)
        self.thread = FileInfoThread(file_paths)
        self.thread.finished.connect(self.on_file_info_ready)
        self.thread.error_signal.connect(lambda msg: QMessageBox.critical(self, "错误", msg))
        self.thread.start()

    def on_file_info_ready(self, info_list):
        self.process_table.setRowCount(0)
        self.process_file_info = info_list
        for idx, (fullpath, pages, mtime, size, error_msg) in enumerate(info_list, start=1):
            row = self.process_table.rowCount()
            self.process_table.insertRow(row)
            header_item = QTableWidgetItem(f"{idx:02d}")
            self.process_table.setVerticalHeaderItem(row, header_item)
            # 列0：文件名
            filename = os.path.basename(fullpath)
            file_item = QTableWidgetItem(filename)
            if error_msg:
                file_item.setForeground(Qt.red)
                flags = file_item.flags()
                flags &= ~Qt.ItemIsUserCheckable
                file_item.setFlags(flags)
            else:
                file_item.setCheckState(Qt.Checked)
            file_item.setData(Qt.UserRole, fullpath)
            seq_str = self.file_seq_map.get(fullpath, f"{idx:02d}")
            file_item.setData(Qt.UserRole + 1, seq_str)
            file_item.setToolTip(fullpath)
            self.process_table.setItem(row, 0, file_item)
            # 页数
            pages_str = "错误" if error_msg else str(pages)
            pages_item = QTableWidgetItem(pages_str)
            pages_item.setTextAlignment(Qt.AlignCenter)
            self.process_table.setItem(row, 1, pages_item)
            # 修改日期
            if mtime:
                mtime_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            else:
                mtime_str = "未知"
            mtime_item = QTableWidgetItem(mtime_str)
            mtime_item.setTextAlignment(Qt.AlignCenter)
            self.process_table.setItem(row, 2, mtime_item)
            # 文件大小
            if size:
                size_kb = size / 1024
                size_str = f"{size_kb / 1024:.2f} MB" if size_kb > 1024 else f"{size_kb:.1f} KB"
            else:
                size_str = "未知"
            size_item = QTableWidgetItem(size_str)
            size_item.setTextAlignment(Qt.AlignCenter)
            self.process_table.setItem(row, 3, size_item)
            # 文件路径（新增）
            path_item = QTableWidgetItem(fullpath)
            path_item.setToolTip(fullpath)
            self.process_table.setItem(row, 4, path_item)
        errors = [info for info in info_list if info[4] is not None]
        if errors:
            msg = "以下文件获取页数失败：\n" + "\n".join([os.path.basename(e[0]) + ": " + e[4] for e in errors])
            QMessageBox.warning(self, "获取页数错误", msg)
        self._update_error_state()
        self.update_button_state()
        self.update_gen_file_list()

    def update_gen_file_list(self):
        self.gen_label.setText(
            "下方显示的是生成文件，请填写文件期望页数并检查文件名是否符合要求：" if self.gen_pages_check.isChecked()
            else "下方显示的是生成文件，请检查文件名是否符合要求："
        )
        # 清空旧内容
        while self.gen_file_layout.count() > 0:
            layout_item = self.gen_file_layout.takeAt(0)
            if layout_item:
                widget = layout_item.widget()
                if widget:
                    widget.deleteLater()
                child_layout = layout_item.layout()
                if child_layout:
                    while child_layout.count() > 0:
                        child_item = child_layout.takeAt(0)
                        if child_item:
                            child_widget = child_item.widget()
                            if child_widget:
                                child_widget.deleteLater()
        # 获取勾选的文件及其序号、页数
        checked_files = []
        row_count = self.process_table.rowCount()
        for row in range(row_count):
            file_item = self.process_table.item(row, 0)
            if file_item and file_item.checkState() == Qt.Checked:
                fullpath = file_item.data(Qt.UserRole)
                seq_str = file_item.data(Qt.UserRole + 1)
                # 从第1列读取页数
                pages_item = self.process_table.item(row, 1)
                # 确保 pages_item 不为空且为数字
                pages = 1
                if pages_item and pages_item.text().isdigit():
                    pages = int(pages_item.text())
                checked_files.append((fullpath, pages, seq_str))
        if not checked_files:
            return
        # 获取生成参数
        prefix = self.gen_prefix_edit.text().strip() or "未命名"
        suffix = self.gen_suffix_edit.text().strip()
        use_date = self.gen_date_check.isChecked()
        date_str = ""
        if use_date:
            fmt = self.gen_date_format.currentText()
            date_str = self.format_date_for_output(QDate.currentDate(), fmt)
        show_pages = self.gen_pages_check.isChecked()
        # 填充布局
        for idx, (fullpath, pages, seq_str) in enumerate(checked_files):
            # 将源文件序号转换为输出格式
            src_type = self.orig_seq_type_combo.currentText() if hasattr(self,
                                                                          'orig_seq_type_combo') else self.seq_type_combo.currentText()
            try:
                seq_num = int(seq_str)
            except ValueError:
                seq_num = 0
                if src_type == "\u6c49\u5b57\u4e00\u3001\u4e8c\u3001\u4e09":
                    seq_list = self.getSeqListFromFile(src_type)
                    if seq_str in seq_list:
                        seq_num = seq_list.index(seq_str) + 1
            output_seq = self.seq_int_to_str(seq_num) if seq_num > 0 else seq_str
            name = prefix
            if use_date and date_str:
                name += date_str
            name += output_seq + suffix + ".docx"
            hbox = QHBoxLayout()
            label = QLabel(f"{idx + 1:02d}  {name}")
            label.setToolTip(name)
            label.setObjectName(f"gen_file_label_{idx}")
            hbox.addWidget(label)
            if show_pages:
                spin = QSpinBox()
                spin.setRange(1, 999)
                spin.setValue(pages if pages > 0 else 1)
                spin.setObjectName(f"expect_pages_{idx}")
                hbox.addWidget(spin)
                hbox.addWidget(QLabel("页"))
            hbox.addStretch()
            self.gen_file_layout.addLayout(hbox)
        # ===== 从加载的预设中恢复期望页数 =====
        if hasattr(self, '_preset_gen_file_pages') and self._preset_gen_file_pages:
            values = list(self._preset_gen_file_pages.values())
            all_same = len(set(values)) == 1
            for i in range(self.gen_file_layout.count()):
                item = self.gen_file_layout.itemAt(i)
                if item and item.layout():
                    lay = item.layout()
                    for j in range(lay.count()):
                        child = lay.itemAt(j)
                        if child and child.widget():
                            cw = child.widget()
                            if isinstance(cw, QSpinBox):
                                if all_same:
                                    cw.setValue(values[0])
                                else:
                                    for k in range(lay.count()):
                                        lbl_item = lay.itemAt(k)
                                        if lbl_item and lbl_item.widget() and isinstance(lbl_item.widget(), QLabel):
                                            lbl_text = lbl_item.widget().toolTip().strip()
                                            if lbl_text in self._preset_gen_file_pages:
                                                cw.setValue(self._preset_gen_file_pages[lbl_text])
                                            break
                                break

    def format_date_for_output(self, date, fmt):
        year = date.year()
        month = date.month()
        day = date.day()
        if fmt == "yyyy年MM月dd日":
            return f"{year:04d}年{month:02d}月{day:02d}日"
        elif fmt == "MM月dd日":
            return f"{month:02d}月{day:02d}日"
        elif fmt == "yyyy.MM.dd":
            return f"{year:04d}.{month:02d}.{day:02d}"
        elif fmt == "MM.dd":
            return f"{month:02d}.{day:02d}"
        elif fmt == "[dd]th,MM(Eng),yyyy":
            months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            mon_eng = months[month - 1]
            if 4 <= day <= 20 or 24 <= day <= 30:
                suffix = "th"
            else:
                suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
            return f"{day:02d}{suffix},{mon_eng},{year:04d}"
        elif fmt == "[dd]th,MM(Eng)":
            months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            mon_eng = months[month - 1]
            if 4 <= day <= 20 or 24 <= day <= 30:
                suffix = "th"
            else:
                suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
            return f"{day:02d}{suffix},{mon_eng}"
        elif fmt == "yyyy-MM-dd":
            return f"{year:04d}-{month:02d}-{day:02d}"
        elif fmt == "MM-dd":
            return f"{month:02d}-{day:02d}"
        else:
            return ""

    def on_gen_pages_check_toggled(self, checked):
        """控制生成文件列表是否显示期望页数输入框"""
        # 直接重新生成列表
        self.update_gen_file_list()

    def apply_first_pages(self):
        """将第一项的期望页数应用到所有项"""
        # 获取第一个生成文件行的期望页数spinbox
        children = self.gen_file_container.children()
        spinboxes = []
        for child in children:
            if isinstance(child, QSpinBox):
                spinboxes.append(child)
        if not spinboxes:
            QMessageBox.information(self, "提示", "没有可应用的页数数据。")
            return
        first_val = spinboxes[0].value()
        for spin in spinboxes[1:]:
            spin.setValue(first_val)
        QMessageBox.information(self, "应用成功", f"已将第一项页数 ({first_val}) 应用到所有项。")

    def update_button_state(self):
        selected = self.process_table.selectedItems()
        has_selection = len(selected) > 0
        self.process_btn.setEnabled(not self._has_checked_errors())
        self.move_down_btn.setEnabled(has_selection)
        self.delete_btn.setEnabled(has_selection)

    def move_up(self):
        selected_rows = self.process_table.selectedItems()
        if not selected_rows:
            return
        current_row = selected_rows[0].row()
        if current_row == 0:
            return
        # 交换表格行
        for col in range(self.process_table.columnCount()):
            item1 = self.process_table.takeItem(current_row, col)
            item2 = self.process_table.takeItem(current_row - 1, col)
            self.process_table.setItem(current_row, col, item2)
            self.process_table.setItem(current_row - 1, col, item1)
        # 交换垂直表头
        header1 = self.process_table.verticalHeaderItem(current_row)
        header2 = self.process_table.verticalHeaderItem(current_row - 1)
        if header1 and header2:
            text1, text2 = header1.text(), header2.text()
            header1.setText(text2)
            header2.setText(text1)
        # 交换复选框状态
        item1 = self.process_table.item(current_row, 0)
        item2 = self.process_table.item(current_row - 1, 0)
        if item1 and item2:
            state1, state2 = item1.checkState(), item2.checkState()
            item1.setCheckState(state2)
            item2.setCheckState(state1)
        # 【关键】同步交换 process_file_list
        self.process_file_list[current_row], self.process_file_list[current_row - 1] = \
            self.process_file_list[current_row - 1], self.process_file_list[current_row]
        self.reorder_table()
        self.process_table.selectRow(current_row - 1)

    def _has_checked_errors(self):
        """检查是否有勾选的错误文件"""
        if not hasattr(self, "process_file_info") or not self.process_file_info:
            return False
        for row in range(self.process_table.rowCount()):
            item = self.process_table.item(row, 0)
            if item is None:
                continue
            if row < len(self.process_file_info) and self.process_file_info[row][4] is not None:
                if item.checkState() == Qt.Checked:
                    return True
        return False

    def _update_error_state(self):
        """根据错误文件状态更新列可见性和页数复选框"""
        has_checked = self._has_checked_errors()
        self.process_table.setColumnHidden(1, has_checked)
        self.gen_pages_check.setEnabled(not has_checked)
        if has_checked:
            self.gen_pages_check.setChecked(False)
        else:
            self.gen_pages_check.setEnabled(True)

    def move_down(self):
        selected_rows = self.process_table.selectedItems()
        if not selected_rows:
            return
        current_row = selected_rows[0].row()
        if current_row == self.process_table.rowCount() - 1:
            return
        # 交换表格行
        for col in range(self.process_table.columnCount()):
            item1 = self.process_table.takeItem(current_row, col)
            item2 = self.process_table.takeItem(current_row + 1, col)
            self.process_table.setItem(current_row, col, item2)
            self.process_table.setItem(current_row + 1, col, item1)
        # 交换垂直表头
        header1 = self.process_table.verticalHeaderItem(current_row)
        header2 = self.process_table.verticalHeaderItem(current_row + 1)
        if header1 and header2:
            text1, text2 = header1.text(), header2.text()
            header1.setText(text2)
            header2.setText(text1)
        # 交换复选框状态
        item1 = self.process_table.item(current_row, 0)
        item2 = self.process_table.item(current_row + 1, 0)
        if item1 and item2:
            state1, state2 = item1.checkState(), item2.checkState()
            item1.setCheckState(state2)
            item2.setCheckState(state1)
        # 【关键】同步交换 process_file_list
        self.process_file_list[current_row], self.process_file_list[current_row + 1] = \
            self.process_file_list[current_row + 1], self.process_file_list[current_row]
        self.reorder_table()
        self.process_table.selectRow(current_row + 1)

    def delete_selected(self):
        rows = sorted(set(item.row() for item in self.process_table.selectedItems()), reverse=True)
        if not rows:
            return
        reply = QMessageBox.question(self, "确认删除", f"确定要删除选中的{len(rows)} 个文件吗？",
                                      QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            # 从后往前删除，避免索引变化
            for row in rows:
                self.process_table.removeRow(row)
                # 同步删除 process_file_list 和 process_file_info 中对应的元素
                del self.process_file_list[row]
                if hasattr(self, "process_file_info") and row < len(self.process_file_info):
                    del self.process_file_info[row]
            self.reorder_table()
            # 检查是否还有错误文件，若无则重新启用页数设置
            self._update_error_state()
            self.update_gen_file_list()

    def clear_list(self):
        reply = QMessageBox.question(self, "确认清空",
                                      "清空后无法处理文件。是不是你的文件夹选错了？请回到“文件名”标签重新选择文件夹。\n是否要继续清空列表？",
                                      QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.process_table.setRowCount(0)
            self.process_file_list.clear()
            self.process_file_info.clear()
            self._update_error_state()
            self.update_gen_file_list()
        self.update_button_state()

    def reorder_table(self):
        """重新排序表格序号（行号）"""
        row_count = self.process_table.rowCount()
        for row in range(row_count):
            header_item = self.process_table.verticalHeaderItem(row)
            if header_item:
                header_item.setText(f"{row + 1:02d}")
            else:
                new_header = QTableWidgetItem(f"{row + 1:02d}")
                self.process_table.setVerticalHeaderItem(row, new_header)
        self.update_gen_file_list()
        self.update_button_state()


    def start_process(self):
        self.log_text.clear()
        self.log_message("开始排版...")
        if self.process_table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "文件列表为空，请先检查输入。")
            return

        if self.date_status_icon.text() != "\u221a":
            QMessageBox.warning(self, "警告", "请在\"标题\"标签页中点击\"检查日期\"按钮来检查！")
            return
        checked_paths = []
        for row in range(self.process_table.rowCount()):
            item = self.process_table.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                fullpath = item.data(Qt.UserRole)
                if fullpath:
                    checked_paths.append(fullpath)
        if not checked_paths:
            QMessageBox.warning(self, "警告", "请至少勾选一个文件。")
            return

        output_dir = self.output_folder_edit.text().strip()
        if not output_dir:
            output_dir = os.path.dirname(checked_paths[0])
        os.makedirs(output_dir, exist_ok=True)

        # 收集期望页数
        expected_pages = {}
        if self.gen_pages_check.isChecked() and hasattr(self, 'gen_file_layout'):
            for i in range(self.gen_file_layout.count()):
                item = self.gen_file_layout.itemAt(i)
                if item and item.layout():
                    lay = item.layout()
                    for j in range(lay.count()):
                        child = lay.itemAt(j)
                        if child and child.widget() and isinstance(child.widget(), QSpinBox):
                            expected_pages[i] = child.widget().value()
                            break

        from docx_formatter import DocumentFormatter, SourceParseError
        formatter = DocumentFormatter(self)
        success = 0
        fail = 0
        fail_msgs = []
        output_paths = []
        for src in checked_paths:
            basename = os.path.basename(src)
            self.log_message(f"正在处理：{basename}")
            QApplication.processEvents()
            try:
                out = formatter.process_file(src, output_dir)
                output_paths.append(out)
                success += 1
                self.log_message(f"已完成 => {os.path.basename(out)}")
            except SourceParseError as e:
                fail += 1
                output_paths.append(None)
                msg = f"文件名格式不匹配 - {str(e)}"
                fail_msgs.append(f"{basename}: {msg}")
                self.log_message(f"失败：{basename} - {msg}")
            except Exception as e:
                fail += 1
                output_paths.append(None)
                msg = str(e)
                fail_msgs.append(f"{basename}: {msg}")
                self.log_message(f"失败：{basename} - {msg}")
            QApplication.processEvents()

        # 页数校验
        mismatches = []
        if self.gen_pages_check.isChecked() and expected_pages and output_paths:
            self.log_message("正在验证页数...")
            QApplication.processEvents()
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            word_app = None
            try:
                try:
                    word_app = win32com.client.Dispatch("Kwps.Application")
                except:
                    try:
                        word_app = win32com.client.Dispatch("wps.Application")
                    except:
                        word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0
                word_app.ScreenUpdating = False
                for idx, out in enumerate(output_paths):
                    if out is None:
                        continue
                    actual_pages = -1
                    try:
                        doc = word_app.Documents.Open(out)
                        doc.Repaginate()
                        actual_pages = doc.ComputeStatistics(2)
                        doc.Close()
                    except Exception:
                        actual_pages = -1
                    expected = expected_pages.get(idx, -1)
                    if expected > 0 and actual_pages > 0 and actual_pages != expected:
                        mismatches.append((os.path.basename(out), expected, actual_pages))
                        self.log_message(f"页数不符：{os.path.basename(out)} 期望{expected}页，实际{actual_pages}页")
            except Exception as e:
                self.log_message(f"无法初始化Word/WPS进行页数校验：{str(e)}")
            finally:
                if word_app:
                    word_app.Quit()
                pythoncom.CoUninitialize()

        # PDF 导出
        output_format = self.output_format.currentText()
        if "PDF" in output_format or "两者" in output_format:
            self.log_message("正在导出PDF...")
            QApplication.processEvents()
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            pdf_word = None
            pdf_success = 0
            pdf_fail = 0
            try:
                try:
                    pdf_word = win32com.client.Dispatch("Kwps.Application")
                except:
                    try:
                        pdf_word = win32com.client.Dispatch("wps.Application")
                    except:
                        pdf_word = win32com.client.Dispatch("Word.Application")
                pdf_word.Visible = False
                pdf_word.DisplayAlerts = 0
                pdf_word.ScreenUpdating = False
                for out in output_paths:
                    if out is None:
                        continue
                    try:
                        pdf_path = out.replace('.docx', '.pdf')
                        doc = pdf_word.Documents.Open(out)
                        doc.SaveAs(pdf_path, FileFormat=17)
                        doc.Close()
                        pdf_success += 1
                        self.log_message(f"PDF已生成：{os.path.basename(pdf_path)}")
                    except Exception as e:
                        pdf_fail += 1
                        self.log_message(f"PDF导出失败：{os.path.basename(out)} - {str(e)}")
            except Exception as e:
                self.log_message(f"无法初始化Word/WPS进行PDF导出：{str(e)}")
            finally:
                if pdf_word:
                    pdf_word.Quit()
                pythoncom.CoUninitialize()
            self.log_message(f"PDF导出完成：成功 {pdf_success} 个，失败 {pdf_fail} 个。")
            # 当只选PDF时删除中间.docx
            if output_format == "Portable Document Format（.pdf）":
                deleted = 0
                for out in output_paths:
                    if out and os.path.exists(out):
                        try:
                            os.remove(out)
                            deleted += 1
                        except Exception:
                            pass
                if deleted > 0:
                    self.log_message(f"已删除 {deleted} 个临时.docx文件")


        # 结果汇总
        result_lines = [f"排版完成：成功 {success} 个，失败 {fail} 个。"]
        if fail_msgs:
            result_lines.append("")
            result_lines.append("失败详情：")
            result_lines.extend(fail_msgs)
        if mismatches:
            result_lines.append("")
            result_lines.append(f"页数不符（共{len(mismatches)}个）：")
            for name, exp, act in mismatches:
                result_lines.append(f"  {name}：期望{exp}页，实际{act}页")

        msg = "\n".join(result_lines)
        self.log_message(msg)

        if mismatches:
            detail_lines = [f"  {name}：期望{exp}页，实际{act}页" for name, exp, act in mismatches]
            detail = "\n".join(detail_lines)
            QMessageBox.warning(self, "排版完成（有页数警告）",
                f"排版完成：成功 {success} 个，失败 {fail} 个。\n\n"
                f"以下 {len(mismatches)} 个文件的页数与期望不符：\n{detail}")
        else:
            QMessageBox.information(self, "排版结果",
                f"排版完成：成功 {success} 个，失败 {fail} 个。")

    def log_message(self, msg):
        """向日志框追加消息"""
        try:
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.append(f"[{timestamp}] {msg}")
            scrollbar = self.log_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
        except Exception:
            pass


    def pickColor(self, btn):
        color = QColorDialog.getColor()
        if color.isValid():
            btn.setStyleSheet(f"background-color: {color.name()};")
            btn.setProperty("color", color)

    def validate_title_inputs(self):
        return True

    def checkDates(self):
        start = self.start_date.date()
        end = self.end_date.date()
        today = QDate.currentDate()
        if start < today or end < today:
            QMessageBox.warning(self, "日期错误", "日期不能早于今天")
            self.setDateStatusIcon(False)
            return
        if end <= start:
            QMessageBox.warning(self, "日期错误", "结束日期必须大于开始日期")
            self.setDateStatusIcon(False)
            return

        # 统计已加载文件数量
        file_count = len(self.process_file_list)

        if file_count == 0:
            QMessageBox.warning(self, "日期检查", "没有加载文件，无法计算日期\n请先在\"文件名\"标签页中加载文件")
            self.setDateStatusIcon(False)
            return

        mode = self.cycle_type_combo.currentText()
        dates = []

        if mode == "按正常工作日安排":
            # 周一至周五工作，周六周日休息
            current = QDate(start)
            i = 0
            while i < file_count:
                day_of_week = current.dayOfWeek()  # 1=周一, 7=周日
                if day_of_week <= 5:  # 周一到周五
                    dates.append(current)
                    i += 1
                current = current.addDays(1)
        else:  # 自定义 - 做N天休M天
            work_days = self.rest_do.value()
            rest_days = self.rest_rest.value()
            if work_days <= 0:
                QMessageBox.warning(self, "日期错误", "休息设置中\"做\"的天数必须大于0")
                self.setDateStatusIcon(False)
                return
            if rest_days == 0:
                current = QDate(start)
                # 不休：每天连续安排
                for _ in range(file_count):
                    dates.append(current)
                    current = current.addDays(1)
            else:
                current = QDate(start)
                i = 0
                cycle_pos = 0
                total_cycle = work_days + rest_days
                if total_cycle == 0:
                    total_cycle = 1
                while i < file_count:
                    if cycle_pos < work_days:
                        dates.append(current)
                        i += 1
                        cycle_pos += 1
                    else:
                        cycle_pos += 1
                        if cycle_pos >= total_cycle:
                            cycle_pos = 0
                    current = current.addDays(1)

        last_date = dates[-1]
        if end < last_date:
            QMessageBox.warning(self, "日期错误",
                f"结束日期太早！\n"
                f"根据当前设置，{file_count}份文件需要安排到 {last_date.toString('yyyy-MM-dd')}。\n"
                f"请将结束日期设置为 {last_date.toString('yyyy-MM-dd')} 或之后。")
            self.setDateStatusIcon(False)
            return

        # 生成日期预览
        date_preview = []
        weekday_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        for idx, d in enumerate(dates):
            seq = self.process_file_list[idx][1] if idx < len(self.process_file_list) else '?'
            wd = weekday_names[d.dayOfWeek() - 1]
            date_preview.append(f"  第{idx+1}份 ({seq}): {d.toString('yyyy-MM-dd')} ({wd})")
        preview_str = "\n".join(date_preview)

        self.setDateStatusIcon(True)
        QMessageBox.information(self, "日期检查",
            f"日期设置有效\n\n"
            f"开始日期：{start.toString('yyyy-MM-dd')}\n"
            f"结束日期：{end.toString('yyyy-MM-dd')}\n"
            f"文件数量：{file_count}\n"
            f"最后文件日期：{last_date.toString('yyyy-MM-dd')}\n"
            f"日期计算方式：{mode}\n\n"
            f"日期安排预览：\n{preview_str}")

    def setDateStatusIcon(self, success):
        if success:
            self.date_status_icon.setText("√")
            self.date_status_icon.setStyleSheet("color: green; font-size: 16px; font-weight: bold;")
        else:
            self.date_status_icon.setText("×")
            self.date_status_icon.setStyleSheet("color: red; font-size: 16px; font-weight: bold;")

    # ==================== 副标题录入对话框 ====================

    def openExaminerDialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("出题人、审题人录入")
        dialog.resize(600, 400)
        layout = QVBoxLayout(dialog)
        label = QLabel("请输入或粘贴出题人和审题人列表，一行一个，分别对应每份试卷的出题人和审题人。")
        layout.addWidget(label)
        hbox = QHBoxLayout()
        vbox1 = QVBoxLayout()
        vbox1.addWidget(QLabel("出题人列表："))
        self.examiner_text = QTextEdit()
        self.examiner_text.setPlaceholderText("每行一个姓名")
        vbox1.addWidget(self.examiner_text)
        hbox.addLayout(vbox1)
        vbox2 = QVBoxLayout()
        vbox2.addWidget(QLabel("审题人列表："))
        self.reviewer_text = QTextEdit()
        self.reviewer_text.setPlaceholderText("每行一个姓名")
        vbox2.addWidget(self.reviewer_text)
        hbox.addLayout(vbox2)
        layout.addLayout(hbox)
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(lambda: self.saveExaminerData(dialog))
        btn_box.rejected.connect(dialog.reject)
        layout.addWidget(btn_box)
        if self.examiners:
            self.examiner_text.setText("\n".join(self.examiners))
        if self.reviewers:
            self.reviewer_text.setText("\n".join(self.reviewers))
        dialog.exec_()

    # ==================== 标题功能开关控制 ====================
    def _toggle_groupbox_by_title(self, enabled, keyword):
        layout = self.tab_title.layout()
        if layout is None:
            return
        for i in range(layout.count()):
            item = layout.itemAt(i)
            if item and item.widget() and isinstance(item.widget(), QGroupBox):
                if keyword in item.widget().title():
                    item.widget().setVisible(enabled)
                    break

    def on_header_enable_toggled(self, enabled):
        self._toggle_groupbox_by_title(enabled, '题头')

    def on_main_enable_toggled(self, enabled):
        self._toggle_groupbox_by_title(enabled, '主标题')

    def on_sub_enable_toggled(self, enabled):
        self._toggle_groupbox_by_title(enabled, '副标题')

    def on_sub_features_toggled(self, enabled):
        """控制副标题内部功能控件的显示/隐藏"""
        if hasattr(self, 'sub_features_container'):
            self.sub_features_container.setVisible(enabled)

    def saveExaminerData(self, dialog):
        examiners = [line.strip() for line in self.examiner_text.toPlainText().splitlines() if line.strip()]
        reviewers = [line.strip() for line in self.reviewer_text.toPlainText().splitlines() if line.strip()]
        self.examiners = examiners
        self.reviewers = reviewers
        dialog.accept()
        QMessageBox.information(self, "保存成功", f"已保存 {len(examiners)} 位出题人，{len(reviewers)} 位审题人。")

    # ==================== 预设配置保存/加载 ====================

    def _get_color_name(self, btn):
        """从颜色按钮的样式表中提取颜色名称"""
        import re
        style = btn.styleSheet()
        m = re.search(r'background-color:\s*(#[0-9a-fA-F]{6}|#[0-9a-fA-F]{3}|[a-zA-Z]+)', style)
        if m:
            return m.group(1)
        return "black"

    def _set_color_btn(self, btn, color_name):
        """设置颜色按钮的背景色"""
        from PyQt5.QtGui import QColor
        btn.setStyleSheet(f"background-color: {color_name};")
        color = QColor(color_name)
        btn.setProperty("color", color)

    def _get_widget_value(self, widget):
        """获取控件当前值，返回适合JSON序列化的类型"""
        cls_name = widget.__class__.__name__
        if cls_name == 'QLineEdit':
            return widget.text()
        elif cls_name == 'QComboBox':
            return widget.currentText()
        elif cls_name in ('QCheckBox',):
            return widget.isChecked()
        elif cls_name == 'QDoubleSpinBox':
            return widget.value()
        elif cls_name == 'QSpinBox':
            return widget.value()
        elif cls_name == 'QDateEdit':
            return widget.date().toString("yyyy-MM-dd")
        elif cls_name == 'QPushButton':
            return self._get_color_name(widget)
        elif cls_name == 'QLabel':
            return {'text': widget.text(), 'styleSheet': widget.styleSheet()}
        return None

    def _set_widget_value(self, widget, value):
        """设置控件值，忽略失败"""
        if value is None:
            return
        try:
            cls_name = widget.__class__.__name__
            if cls_name == 'QLineEdit':
                widget.setText(str(value))
            elif cls_name == 'QComboBox':
                idx = widget.findText(str(value))
                if idx >= 0:
                    widget.setCurrentIndex(idx)
            elif cls_name in ('QCheckBox',):
                widget.blockSignals(True)
                widget.setChecked(bool(value))
                widget.blockSignals(False)
            elif cls_name == 'QDoubleSpinBox':
                widget.setValue(float(value))
            elif cls_name == 'QSpinBox':
                widget.setValue(int(value))
            elif cls_name == 'QDateEdit':
                from PyQt5.QtCore import QDate
                widget.setDate(QDate.fromString(str(value), "yyyy-MM-dd"))
            elif cls_name == 'QPushButton':
                self._set_color_btn(widget, str(value))
            elif cls_name == 'QLabel' and isinstance(value, dict):
                widget.setText(value.get('text', ''))
                widget.setStyleSheet(value.get('styleSheet', ''))
        except Exception:
            pass  # 单个控件加载失败不中断

    def savePreset(self):
        """保存当前所有设置到JSON文件"""
        import json
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        from PyQt5.QtCore import QDate
        from PyQt5.QtWidgets import QLabel, QSpinBox

        default_name = "排版预设.json"
        if self.preset_config_path:
            default_name = self.preset_config_path

        filepath, _ = QFileDialog.getSaveFileName(
            self, "保存预设配置", default_name, "JSON文件 (*.json)"
        )
        if not filepath:
            return

        data = {}

        # ---- 文件名Tab ----
        tab_widgets_file = [
            ('folder_edit', self.folder_edit),
            ('output_folder_edit', self.output_folder_edit),
            ('orig_date_check', self.orig_date_check),
            ('orig_prefix_edit', self.orig_prefix_edit),
            ('orig_date_format', self.orig_date_format),
            ('orig_suffix_edit', self.orig_suffix_edit),
            ('orig_seq_type_combo', self.orig_seq_type_combo),
            ('gen_date_check', self.gen_date_check),
            ('gen_prefix_edit', self.gen_prefix_edit),
            ('gen_date_format', self.gen_date_format),
            ('gen_suffix_edit', self.gen_suffix_edit),
            ('seq_type_combo', self.seq_type_combo),
            ('subject_name_edit', self.subject_name_edit),
        ]
        for name, w in tab_widgets_file:
            data[name] = self._get_widget_value(w)

        # ---- 标题Tab ----
        # Header
        tab_widgets_title_header = [
            ('enable_header_cb', self.enable_header_cb),
            ('header_content_edit', self.header_content_edit),
            ('header_font', self.header_font),
            ('header_font_size', self.header_font_size),
            ('header_bold', self.header_bold),
            ('header_underline', self.header_underline),
            ('header_color_btn', self.header_color_btn),
            ('header_align', self.header_align),
            ('header_line_spacing_type', self.header_line_spacing_type),
            ('header_line_spacing_value', self.header_line_spacing_value),
        ]
        for name, w in tab_widgets_title_header:
            data[name] = self._get_widget_value(w)

        # Main title
        tab_widgets_title_main = [
            ('enable_main_cb', self.enable_main_cb),
            ('main_prefix_edit', self.main_prefix_edit),
            ('main_suffix_edit', self.main_suffix_edit),
            ('main_seq_type', self.main_seq_type),
            ('main_font', self.main_font),
            ('main_font_size', self.main_font_size),
            ('main_bold', self.main_bold),
            ('main_underline', self.main_underline),
            ('main_color_btn', self.main_color_btn),
            ('main_align', self.main_align),
            ('main_line_spacing_type', self.main_line_spacing_type),
            ('main_line_spacing_value', self.main_line_spacing_value),
        ]
        for name, w in tab_widgets_title_main:
            data[name] = self._get_widget_value(w)

        # Sub title
        tab_widgets_title_sub = [
            ('enable_sub_cb', self.enable_sub_cb),
            ('sub_font', self.sub_font),
            ('sub_font_size', self.sub_font_size),
            ('sub_bold', self.sub_bold),
            ('sub_underline', self.sub_underline),
            ('sub_color_btn', self.sub_color_btn),
            ('sub_align', self.sub_align),
            ('sub_line_spacing_type', self.sub_line_spacing_type),
            ('sub_line_spacing_value', self.sub_line_spacing_value),
        ]
        for name, w in tab_widgets_title_sub:
            data[name] = self._get_widget_value(w)

        # Sub features
        tab_widgets_sub_features = [
            ('cycle_type_combo', self.cycle_type_combo),
            ('start_date', self.start_date),
            ('end_date', self.end_date),
            ('enable_examiner_cb', self.enable_examiner_cb),
            ('enable_date_cb', self.enable_date_cb),
            ('rest_do', self.rest_do),
            ('rest_rest', self.rest_rest),

        ]
        for name, w in tab_widgets_sub_features:
            data[name] = self._get_widget_value(w)

        # Examiners and reviewers
        data['examiners'] = self.examiners
        data['reviewers'] = self.reviewers

        # ---- 页面Tab ----
        tab_widgets_page = [
            ('margin_top', self.margin_top),
            ('margin_bottom', self.margin_bottom),
            ('margin_left', self.margin_left),
            ('margin_right', self.margin_right),
            ('header_mode', self.header_mode),
            ('header_text_edit', self.header_text_edit),
            ('insert_page_check', self.insert_page_check),
            ('page_position', self.page_position),
            ('page_style', self.page_style),
            ('page_font', self.page_font),
            ('page_font_size', self.page_font_size),
            ('page_bold', self.page_bold),
            ('page_underline', self.page_underline),
            ('page_italic', self.page_italic),
            ('page_color_btn', self.page_color_btn),
            ('paper_size', self.paper_size),
            ('enforce_body_font', self.enforce_body_font),
            ('question_line_spacing_type', self.question_line_spacing_type),
            ('question_line_spacing_value', self.question_line_spacing_value),
        ]
        for name, w in tab_widgets_page:
            data[name] = self._get_widget_value(w)

        # ---- 处理Tab ----
        data['output_format'] = self._get_widget_value(self.output_format)
        data['gen_pages_check'] = self._get_widget_value(self.gen_pages_check)
        data['log_text'] = {'text': '', 'styleSheet': self.log_text.styleSheet()}

        # 保存每个文件的期望页数
        file_pages = {}
        if hasattr(self, 'gen_file_layout') and self.gen_file_layout is not None:
            for i in range(self.gen_file_layout.count()):
                item = self.gen_file_layout.itemAt(i)
                if item and item.layout():
                    lay = item.layout()
                    filename_label = None
                    pages_spin = None
                    for j in range(lay.count()):
                        child = lay.itemAt(j)
                        if child and child.widget():
                            cw = child.widget()
                            if isinstance(cw, QLabel) and cw.objectName().startswith('gen_file_label_'):
                                filename_label = cw
                            elif isinstance(cw, QSpinBox):
                                pages_spin = cw
                    if filename_label and pages_spin is not None:
                        fname = filename_label.toolTip().strip()
                        file_pages[fname] = pages_spin.value()
        data['gen_file_pages'] = file_pages

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.preset_config_path = filepath
            self.config_path_label.setText(f"配置文件：{filepath}")
            QMessageBox.information(self, "保存成功", f"预设配置已保存到：{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"保存预设时出错：\n{str(e)}")

    def loadPreset(self):
        """从JSON文件加载预设配置"""
        import json
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        from PyQt5.QtCore import QDate

        filepath, _ = QFileDialog.getOpenFileName(
            self, "打开预设配置", "", "JSON文件 (*.json)"
        )
        if not filepath:
            return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"无法读取配置文件：\n{str(e)}")
            return

        failed_items = []

        # 执行加载
        for name, value in data.items():
            if name in ('examiners', 'reviewers', 'date_status_icon', 'footer_left_edit', 'footer_right_edit', 'gen_file_pages', 'preset_source_files', 'header_left_edit', 'header_right_edit', 'gen_pages_check', 'enable_rest_cb'):
                continue
            widget = getattr(self, name, None)
            if widget is None:
                failed_items.append(name)
                continue
            try:
                self._set_widget_value(widget, value)
            except Exception:
                failed_items.append(name)

        # 加载出题人、审题人
        if 'examiners' in data:
            self.examiners = list(data['examiners'])
        if 'reviewers' in data:
            self.reviewers = list(data['reviewers'])

        # 存储期望页数，等生成文件列表后恢复
        if 'gen_file_pages' in data and data['gen_file_pages']:
            self._preset_gen_file_pages = data['gen_file_pages']
        # 确保gen_pages_check按文件加载（显式处理）
        if 'gen_pages_check' in data:
            self.gen_pages_check.blockSignals(True)
            self.gen_pages_check.setChecked(bool(data['gen_pages_check']))
            self.gen_pages_check.blockSignals(False)

        # 更新config label
        self.preset_config_path = filepath
        self.config_path_label.setText(f"配置文件：{filepath}")

        # 报告加载结果
        if not failed_items:
            QMessageBox.information(self, "加载成功",
                                    f"预设配置已从以下文件加载：\n{filepath}\n\n所有设置已成功应用。")
        else:
            QMessageBox.warning(self, "部分加载失败",
                                f"预设配置已从以下文件加载：\n{filepath}\n\n"
                                f"但以下 {len(failed_items)} 项加载失败：\n" + "、".join(failed_items))

    def closeEvent(self, event):
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
