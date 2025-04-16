import sys
import re
import os
import subprocess
import tempfile
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton,
    QPlainTextEdit, QProgressDialog, QFileDialog
)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文本公式转换工具")
        self.resize(600, 400)

        # 创建文本编辑框，支持无限长度的文本输入
        self.editor = QPlainTextEdit()

        # 创建按钮，点击后处理文本并自动调用 pandoc 命令
        self.process_button = QPushButton("处理并转换为 DOCX")
        self.process_button.clicked.connect(self.process_text)

        # 布局
        layout = QVBoxLayout()
        layout.addWidget(self.editor)
        layout.addWidget(self.process_button)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def process_text(self):
        # 初始化进度弹窗：
        # 参数说明：窗口文本、取消按钮文本、最小值、最大值以及父窗口对象
        progress = QProgressDialog("开始处理...", "取消", 0, 4, self)
        progress.setWindowTitle("进度")
        progress.setMinimumDuration(0)  # 立即显示
        progress.setValue(0)
        QApplication.processEvents()  # 刷新界面

        # 获取文本框内的内容
        text = self.editor.toPlainText()

        # ① 删除全文中的窄不换行空格 (Unicode U+202F)
        text = text.replace('\u202f', '')

        # 删除其它特殊标记（例如 cite 相关标记）
        text = re.sub('\ue200.*?\ue201', '', text, flags=re.DOTALL)
        text = text.translate({0xE200: None, 0xE201: None, 0xE202: None})
        # 压缩多余空格与空行
        text = re.sub('[ \t]+\n', '\n', text)  # 清除行尾空格
        text = re.sub('\n{3,}', '\n\n', text)    # 连续 ≥3 行压缩为 2 行

        # 处理数学公式：转换公式环境
        # 内联公式：\(...\) 转换为 $...$
        pattern_inline = re.compile(r'\\\(\s*(.*?)\s*\\\)')
        text = pattern_inline.sub(lambda m: '$' + m.group(1).strip() + '$', text)
        # 显示公式：\[...\] 转换为 $$...$$（支持跨行匹配）
        pattern_display = re.compile(r'\\\[\s*(.*?)\s*\\\]', re.DOTALL)
        text = pattern_display.sub(lambda m: '$$' + m.group(1).strip() + '$$', text)

        # 删除 Markdown 标题行中的编号
        pattern_title = re.compile(
            r'^(\s*#+\s*)((?:(?:\d+(?:\.\d+)*\.?)|(?:[一二三四五六七八九十]+(?:、|[,.，])?))\s*)',
            flags=re.MULTILINE
        )
        text = pattern_title.sub(r'\1', text)

        # 移除仅包含 --- 的分隔行
        lines = text.splitlines()
        filtered_lines = [line for line in lines if not re.match(r'^\s*---\s*$', line)]
        text = "\n".join(filtered_lines)

        # 更新进度：文本处理完成
        progress.setLabelText("文本处理完成，创建临时文件...")
        progress.setValue(1)
        QApplication.processEvents()

        # ------------------------------
        # 1. 在当前目录生成临时 Markdown 文件
        # ------------------------------
        fd, temp_path = tempfile.mkstemp(suffix=".md")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)

        progress.setLabelText("临时文件创建成功，调用 Pandoc 进行转换...")
        progress.setValue(2)
        QApplication.processEvents()

        # ------------------------------
        # 2. 调用 pandoc 命令转换为 DOCX 文件
        # 命令格式：
        # pandoc 临时文件路径 -o D:\Desktop\output.docx --reference-doc=Doc11.docx
        # ------------------------------
        # 假设 Doc11.docx 与 exe 文件在同一目录
        ref_doc_path = os.path.join(os.path.dirname(sys.executable), "Doc11.docx") if getattr(sys, 'frozen',
                                                                                              False) else "Doc11.docx"
        pandoc_command = [
            'pandoc',
            temp_path,
            '-o',
            r'D:\Desktop\output.docx',
            f'--reference-doc={ref_doc_path}'
        ]
        try:
            subprocess.run(
                pandoc_command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True
            )
        except subprocess.CalledProcessError as e:
            progress.setLabelText("Pandoc 转换出错！")
            progress.setValue(4)
            QApplication.processEvents()
            return

        progress.setLabelText("Pandoc 转换成功，删除临时文件...")
        progress.setValue(3)
        QApplication.processEvents()

        # ------------------------------
        # 3. 删除临时 Markdown 文件
        # ------------------------------
        os.remove(temp_path)

        progress.setLabelText("处理完成！")
        progress.setValue(4)
        QApplication.processEvents()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
