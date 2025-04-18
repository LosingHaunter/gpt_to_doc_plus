import sys
import re
import os
import subprocess
import tempfile
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QPushButton, QPlainTextEdit, QProgressDialog, QFileDialog,
    QDialog, QLineEdit, QDialogButtonBox, QLabel, QGridLayout
)
from PyQt5.QtCore import QSettings  # 修正：从 QtCore 导入 QSettings

DEFAULT_TEMPLATE = 'Temple.docx'
ORGANIZATION = 'YmY'
APPLICATION = 'GPT_to_Doc_PLUS'

def get_app_dir():
    """
    获取应用所在目录：
    如果是打包后的独立 exe（onefile 模式），通过 sys.argv[0] 获取原始 exe 的目录，
    否则返回当前脚本所在目录。
    """
    if getattr(sys, 'frozen', False):
        # 对于打包后，不使用 sys._MEIPASS，而是使用 sys.argv[0] 所在目录
        return os.path.dirname(sys.argv[0])
    else:
        return os.path.abspath(".")

APP_DIR = get_app_dir()
PANDOC_CMD = 'pandoc'

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GPT_to_Doc转换工具")
        self.resize(600, 400)

        # 初始化 QSettings
        self.settings = QSettings(ORGANIZATION, APPLICATION)
        # 加载上次保存的配置
        self.template_file = self.settings.value('template_file', os.path.join(APP_DIR, DEFAULT_TEMPLATE))
        self.output_dir = self.settings.value('output_dir', APP_DIR)
        filename = self.settings.value('filename_base', 'output')
        self.filename_base = filename

        # ---------- 设置按钮，用于打开设置对话框 ----------
        settings_btn = QPushButton("设置")
        settings_btn.clicked.connect(self.open_settings)

        top_layout = QHBoxLayout()
        top_layout.addWidget(settings_btn)
        top_layout.addStretch()
        # --------------------------------------------

        # 创建文本编辑框，支持无限长度的文本输入
        self.editor = QPlainTextEdit()

        # 创建按钮，点击后处理文本并自动调用 pandoc 命令
        self.process_button = QPushButton("处理并转换为 DOCX")
        self.process_button.clicked.connect(self.process_text)

        # 布局
        layout = QVBoxLayout()
        layout.addLayout(top_layout)
        layout.addWidget(self.editor)
        layout.addWidget(self.process_button)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def open_settings(self):
        """
        新增：弹出设置对话框，可设置模板、输出目录及文件名，使用网格布局对齐
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("设置")
        dlg_layout = QVBoxLayout()
        dialog.setLayout(dlg_layout)

        # 使用网格布局，使控件列对齐
        grid = QGridLayout()
        # 模板文件
        grid.addWidget(QLabel("模板文件："), 0, 0)
        label_tpl = QLabel(self.template_file)
        grid.addWidget(label_tpl, 0, 1)
        btn_tpl = QPushButton("选择模板")
        btn_tpl.clicked.connect(lambda: self._choose_file(label_tpl, "Word 文档 (*.docx)"))
        grid.addWidget(btn_tpl, 0, 2)
        # 输出目录
        grid.addWidget(QLabel("输出目录："), 1, 0)
        label_outdir = QLabel(self.output_dir)
        grid.addWidget(label_outdir, 1, 1)
        btn_outdir = QPushButton("选择目录")
        btn_outdir.clicked.connect(lambda: self._choose_dir(label_outdir))
        grid.addWidget(btn_outdir, 1, 2)
        # 文件名
        grid.addWidget(QLabel("文件名："), 2, 0)
        le_name = QLineEdit(self.filename_base)
        grid.addWidget(le_name, 2, 1)
        grid.addWidget(QLabel(".docx"), 2, 2)

        dlg_layout.addLayout(grid)

        # 对话框按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        dlg_layout.addWidget(btn_box)
        btn_box.accepted.connect(lambda: self._save_settings(dialog, label_tpl.text(), label_outdir.text(), le_name.text()))
        btn_box.rejected.connect(dialog.reject)

        dialog.exec_()

    def _choose_file(self, label, filter_str):
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", APP_DIR, filter_str)
        if path:
            label.setText(path)

    def _choose_dir(self, label):
        path = QFileDialog.getExistingDirectory(self, "选择目录", APP_DIR)
        if path:
            label.setText(path)

    def _save_settings(self, dialog, tpl, outdir, name):
        self.template_file = tpl
        self.output_dir = outdir
        self.filename_base = name
        self.settings.setValue('template_file', tpl)
        self.settings.setValue('output_dir', outdir)
        self.settings.setValue('filename_base', name)
        dialog.accept()

    def process_text(self):
        # 初始化进度弹窗：
        # 参数说明：窗口文本、取消按钮文本、最小值、最大值以及父窗口对象
        progress = QProgressDialog("开始处理...", "取消", 0, 4, self)
        progress.setWindowTitle("进度")
        progress.setMinimumDuration(0)  # 立即显示
        progress.setAutoClose(False)  # 防止达到最大值自动关闭，保留错误提示
        progress.setValue(0)
        QApplication.processEvents()  # 刷新界面

        # 获取文本框内的内容
        text = self.editor.toPlainText()

        # 功能一：删除全文中的窄不换行空格 (Unicode U+202F)
        text = text.replace('\u202f', '')

        # 功能二：删除其它特殊标记（例如 cite 相关标记）
        text = re.sub('\ue200.*?\ue201', '', text, flags=re.DOTALL)
        text = text.translate({0xE200: None, 0xE201: None, 0xE202: None})
        # 压缩多余空格与空行
        text = re.sub('[ \t]+\n', '\n', text)  # 清除行尾空格
        text = re.sub('\n{3,}', '\n\n', text)  # 连续 ≥3 行压缩为 2 行

        # 功能三：处理数学公式：转换公式环境
        # 内联公式：\(...\) 转换为 $...$
        text = re.sub(r'\\\(\s*(.*?)\s*\\\)', lambda m: '$' + m.group(1).strip() + '$', text)
        # 显示公式：\[...\] 转换为 $$...$$
        text = re.sub(r'\\\[\s*(.*?)\\\]', lambda m: '$$' + m.group(1).strip() + '$$', text, flags=re.DOTALL)

        # 功能四：删除 Markdown 标题行中的编号
        text = re.sub(
            r'^(\s*#+\s*)((?:(?:\d+(?:\.\d+)*\.?))|(?:[一二三四五六七八九十]+(?:、|[.,，])?))\s*',
            r'\1', text, flags=re.MULTILINE)

        # 功能五：移除仅包含 --- 的分隔行
        lines = text.splitlines()
        text = "\n".join([ln for ln in lines if not re.match(r'^\s*---\s*$', ln)])

        progress.setLabelText("文本处理完成，创建临时文件...")
        progress.setValue(1)
        QApplication.processEvents()

        # ------------------------------
        # 1. 在当前目录生成临时 Markdown 文件
        # ------------------------------
        fd, temp_md = tempfile.mkstemp(suffix=".md")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)

        progress.setLabelText("临时文件创建成功，调用 Pandoc 进行转换...")
        progress.setValue(2)
        QApplication.processEvents()

        # ------------------------------
        # 2. 调用 pandoc 命令转换为 DOCX 文件
        # ------------------------------
        output_filename = f"{self.filename_base}.docx"
        target_path = os.path.join(self.output_dir, output_filename)
        cmd = [PANDOC_CMD, temp_md, '-o', target_path, f'--reference-doc={self.template_file}']
        try:
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
        except subprocess.CalledProcessError as e:
            progress.setLabelText("Pandoc 转换出错：" + e.stderr)
            progress.setValue(4)
            QApplication.processEvents()
            return

        progress.setLabelText("Pandoc 转换成功，删除临时文件...")
        progress.setValue(3)
        QApplication.processEvents()

        # ------------------------------
        # 3. 删除临时 Markdown 文件
        # ------------------------------
        os.remove(temp_md)

        progress.setLabelText("处理完成！")
        progress.setValue(4)
        QApplication.processEvents()
        progress.close()  # 操作完成后手动关闭弹窗


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
