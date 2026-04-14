import sys
import os
import json
from pathlib import Path
from typing import Dict, Optional

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QCheckBox, QSlider, QFileDialog,
    QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QDialog, QTextBrowser, QGroupBox
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont

# 假设存在 pdf_keyword_finder 模块，若本地测试可注释掉或提供 mock
try:
    import pdf_keyword_finder
except ImportError:
    class pdf_keyword_finder:
        @staticmethod
        def find_keywords_in_pdf(**kwargs):
            import time
            time.sleep(2)  # 模拟耗时操作
            return {
                'total_matches': 10, 
                'noise_info': [
                    {'text': 'Mock Header Noise', 'reason': 'Header', 'page': 1, 'repeat_rate': 0.85},
                    {'text': 'Mock Footer Noise', 'reason': 'Footer', 'page': 1, 'repeat_rate': 0.75}
                ]
            }


class SearchWorker(QThread):
    """搜索工作线程，避免阻塞UI"""
    search_finished = Signal(dict)
    search_error = Signal(str)

    def __init__(self, params: dict):
        super().__init__()
        self.params = params
        self._cancel_flag = False

    def run(self):
        try:
            results = pdf_keyword_finder.find_keywords_in_pdf(**self.params)
            if not self._cancel_flag:
                self.search_finished.emit(results)
        except Exception as e:
            if not self._cancel_flag:
                self.search_error.emit(str(e))

    def cancel(self):
        self._cancel_flag = True


class NoiseResultDialog(QDialog):
    """噪声检测结果弹窗"""
    def __init__(self, noise_info: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("噪声检测结果")
        self.resize(700, 500)

        layout = QVBoxLayout(self)

        title_label = QLabel("噪声检测完成")
        title_label.setFont(QFont("", 14, QFont.Bold))
        layout.addWidget(title_label)

        # 数据聚合与统计
        noise_aggregation = {}
        for item in noise_info:
            text = item.get('text', '')
            if text not in noise_aggregation:
                noise_aggregation[text] = {
                    'text': text,
                    'reason': item.get('reason', ''),
                    'pages': set(),
                    'repeat_rate': item.get('repeat_rate', 0),
                    'position': item.get('position', '')
                }
            noise_aggregation[text]['pages'].add(item.get('page', ''))

        unique_noise_list = list(noise_aggregation.values())

        # 按类型和重复率排序：先'边缘位置+高频重复'，后'边缘位置+带有数字'
        def sort_key(item):
            reason = item.get('reason', '')
            repeat_rate = item.get('repeat_rate', 0)
            # 排序优先级：高频重复=0，数字=1，其他=2
            if '高频重复' in reason:
                priority = 0
            elif '数字' in reason:
                priority = 1
            else:
                priority = 2
            return (priority, -repeat_rate)  # 负号表示降序

        unique_noise_list.sort(key=sort_key)

        reason_count = {}
        for item in unique_noise_list:
            reason = item.get('reason', '未知')
            reason_count[reason] = reason_count.get(reason, 0) + 1

        total_unique = len(unique_noise_list)
        total_blocks = len(noise_info)

        # 按排序顺序显示统计
        summary = f"共检测到 {total_blocks} 个噪声block，去重后 {total_unique} 种\n\n按类型统计:\n"
        # 按优先级显示
        reason_order = ['边缘位置+高频重复', '边缘位置+带有数字']
        for reason in reason_order:
            if reason in reason_count:
                summary += f"  • {reason}: {reason_count[reason]}种\n"
        # 显示其他类型
        for reason, count in reason_count.items():
            if reason not in reason_order:
                summary += f"  • {reason}: {count}种\n"

        summary_label = QLabel(summary)
        summary_label.setWordWrap(True)
        layout.addWidget(summary_label)

        detail_label = QLabel("去重后的噪声详情（按类型排序，类型内按重复率降序）:")
        detail_label.setFont(QFont("", -1, QFont.Bold))
        layout.addWidget(detail_label)

        # 详细列表显示
        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(False)
        details = []

        # 追踪当前类型，添加分组标题
        current_reason = None
        for item in unique_noise_list[:15]:
            reason = item.get('reason', '未知')
            repeat_rate = item.get('repeat_rate', 0)
            pages = item.get('pages', set())
            text = item.get('text', '')
            display_text = text[:50] + '...' if len(text) > 50 else text

            # 类型变化时添加分组标题
            if current_reason != reason:
                details.append(f"\n【{reason}】")
                current_reason = reason

            details.append(f"  • 重复率:{repeat_rate:.1%} | 页数:{len(pages)}页 | 内容: {display_text}")

        if len(unique_noise_list) > 15:
            details.append(f"\n... 还有 {len(unique_noise_list) - 15} 种噪声未显示")

        text_browser.setText("\n".join(details))
        layout.addWidget(text_browser)

        # 添加查看完整列表的展开功能
        self.expand_btn = QPushButton("查看全部噪声详情")
        self.expand_btn.clicked.connect(lambda: self._show_full_list(unique_noise_list))
        layout.addWidget(self.expand_btn)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("确定")
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)
        layout.addLayout(btn_layout)

    def _show_full_list(self, unique_noise_list: list):
        """显示完整的噪声列表"""
        dialog = QDialog(self)
        dialog.setWindowTitle("全部噪声详情")
        dialog.resize(700, 500)

        layout = QVBoxLayout(dialog)

        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(False)
        details = []

        # 按分组显示
        current_reason = None
        group_idx = 0
        for item in unique_noise_list:
            text = item.get('text', '')
            reason = item.get('reason', '未知')
            repeat_rate = item.get('repeat_rate', 0)
            pages = item.get('pages', set())
            pages_str = ', '.join(sorted(str(p) for p in pages)) if pages else 'N/A'
            position = item.get('position')
            # 类型变化时添加分组标题
            if current_reason != reason:
                if current_reason is not None:
                    details.append("")
                details.append(f"【{reason}】")
                current_reason = reason
                group_idx = 1  # 组内序号

            details.append(f"  {group_idx}. 重复率:{repeat_rate:.1%} | 位置:{position} | 页码: {pages_str}\n     内容: {text}")
            group_idx += 1

        text_browser.setText("\n".join(details))
        layout.addWidget(text_browser)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dialog.close)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        dialog.exec()


class PDFKeywordFinderApp(QMainWindow):
    """PDF关键字搜索工具主窗口 - 基于 PySide6"""

    def __init__(self):
        super().__init__()

        # 窗口基本设置
        self.setWindowTitle("解决方案中心工具")
        self.resize(1000, 750)
        self.setMinimumSize(900, 650)

        # 初始化数据与状态
        self.keywords: Dict[str, int] = {}
        self.search_worker: Optional[SearchWorker] = None
        self.is_searching = False
        self._last_noise_info = []

        self._init_ui()

    def _init_ui(self):
        """初始化界面组件"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 文件选择区域
        main_layout.addWidget(self._create_file_selection())

        # 中间区域：左侧关键字 + 右侧设置
        middle_layout = self._create_middle_section()
        main_layout.addLayout(middle_layout, stretch=1)

        # 底部操作区域
        main_layout.addLayout(self._create_action_section())

    def _create_file_selection(self) -> QGroupBox:
        """创建文件选择区域"""
        group_box = QGroupBox("文件选择")
        layout = QVBoxLayout()

        # PDF文件选择
        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(QLabel("PDF文件:"))
        self.pdf_entry = QLineEdit()
        self.pdf_entry.setPlaceholderText("请选择PDF文件...")
        pdf_layout.addWidget(self.pdf_entry)
        self.pdf_btn = QPushButton("选择文件")
        self.pdf_btn.clicked.connect(self._select_pdf)
        pdf_layout.addWidget(self.pdf_btn)
        layout.addLayout(pdf_layout)

        # 输出目录选择
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_entry = QLineEdit()
        self.output_entry.setPlaceholderText("默认与PDF文件同目录")
        output_layout.addWidget(self.output_entry)
        self.output_btn = QPushButton("选择目录")
        self.output_btn.clicked.connect(self._select_output_dir)
        output_layout.addWidget(self.output_btn)
        self.open_dir_btn = QPushButton("打开目录")
        self.open_dir_btn.clicked.connect(self._open_output_dir)
        output_layout.addWidget(self.open_dir_btn)
        layout.addLayout(output_layout)

        group_box.setLayout(layout)
        return group_box

    def _create_middle_section(self) -> QHBoxLayout:
        """创建中间区域布局"""
        layout = QHBoxLayout()

        # 左侧：关键字管理
        layout.addLayout(self._create_keyword_section(), stretch=3)

        # 右侧：搜索设置面板
        layout.addLayout(self._create_settings_section(), stretch=2)

        return layout

    def _create_keyword_section(self) -> QVBoxLayout:
        """创建关键字管理区域"""
        layout = QVBoxLayout()
        
        header = QLabel("关键字管理")
        header.setFont(QFont("", 12, QFont.Bold))
        layout.addWidget(header)

        # 关键字列表表格
        self.keyword_table = QTableWidget(0, 2)
        self.keyword_table.setHorizontalHeaderLabels(["关键字", "分数"])
        self.keyword_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.keyword_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.keyword_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.keyword_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.keyword_table)

        # 添加关键字控件
        add_layout = QHBoxLayout()
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("输入关键字")
        self.keyword_input.returnPressed.connect(self._add_keyword)
        add_layout.addWidget(self.keyword_input)

        self.score_input = QLineEdit("1")
        self.score_input.setPlaceholderText("分数")
        self.score_input.setFixedWidth(60)
        self.score_input.returnPressed.connect(self._add_keyword)
        add_layout.addWidget(self.score_input)

        add_btn = QPushButton("添加")
        add_btn.clicked.connect(self._add_keyword)
        add_layout.addWidget(add_btn)
        layout.addLayout(add_layout)

        # 批量操作按钮
        action_layout = QHBoxLayout()
        self.save_cfg_btn = QPushButton("保存配置")
        self.save_cfg_btn.clicked.connect(self._save_config)
        action_layout.addWidget(self.save_cfg_btn)

        self.load_cfg_btn = QPushButton("加载配置")
        self.load_cfg_btn.clicked.connect(self._load_config)
        action_layout.addWidget(self.load_cfg_btn)

        delete_btn = QPushButton("删除选中")
        delete_btn.clicked.connect(self._delete_keyword)
        action_layout.addWidget(delete_btn)

        clear_btn = QPushButton("清空")
        clear_btn.clicked.connect(self._clear_keywords)
        action_layout.addWidget(clear_btn)

        action_layout.addStretch()
        layout.addLayout(action_layout)

        return layout

    def _create_settings_section(self) -> QVBoxLayout:
        """创建搜索设置面板"""
        layout = QVBoxLayout()

        header = QLabel("搜索设置")
        header.setFont(QFont("", 12, QFont.Bold))
        layout.addWidget(header)

        # 上下文设置
        context_group = QGroupBox("上下文设置")
        context_layout = QVBoxLayout()
        
        context_layout.addWidget(QLabel("上下文丰富度:"))
        context_slider_layout = QHBoxLayout()
        self.context_slider = QSlider(Qt.Horizontal)
        self.context_slider.setRange(50, 1000)
        self.context_slider.setSingleStep(50)
        self.context_slider.setValue(200)
        self.context_slider.setTickPosition(QSlider.TicksBelow)
        self.context_slider.setTickInterval(50)
        context_slider_layout.addWidget(self.context_slider)
        self.context_value_label = QLabel("200")
        self.context_slider.valueChanged.connect(lambda v: self.context_value_label.setText(str(v)))
        context_slider_layout.addWidget(self.context_value_label)
        context_layout.addLayout(context_slider_layout)

        front_layout = QHBoxLayout()
        front_layout.addWidget(QLabel("前窗口字数:"))
        self.front_entry = QLineEdit("0")
        self.front_entry.setFixedWidth(80)
        front_layout.addWidget(self.front_entry)
        front_layout.addStretch()
        context_layout.addLayout(front_layout)

        context_group.setLayout(context_layout)
        layout.addWidget(context_group)

        # 输出选项
        output_group = QGroupBox("输出选项")
        output_layout = QVBoxLayout()
        self.txt_check = QCheckBox("输出 TXT 文件")
        self.txt_check.setChecked(True)
        output_layout.addWidget(self.txt_check)
        self.excel_check = QCheckBox("输出 Excel 文件")
        self.excel_check.setChecked(True)
        output_layout.addWidget(self.excel_check)
        self.sort_check = QCheckBox("按重要性排序")
        self.sort_check.setChecked(True)
        output_layout.addWidget(self.sort_check)
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # 噪声检测设置
        noise_group = QGroupBox("噪声检测设置")
        noise_layout = QVBoxLayout()
        self.noise_check = QCheckBox("自动检测页眉页脚水印")
        self.noise_check.setChecked(False)
        self.noise_check.toggled.connect(self._toggle_noise_params)
        noise_layout.addWidget(self.noise_check)

        # 页眉区域
        noise_layout.addWidget(QLabel("页眉区域 (%):"))
        h_layout = QHBoxLayout()
        self.header_slider = QSlider(Qt.Horizontal)
        self.header_slider.setRange(0, 40)
        self.header_slider.setValue(10)
        self.header_slider.setEnabled(False)
        h_layout.addWidget(self.header_slider)
        self.header_value_label = QLabel("15%")
        self.header_slider.valueChanged.connect(lambda v: self.header_value_label.setText(f"{v}%"))
        h_layout.addWidget(self.header_value_label)
        noise_layout.addLayout(h_layout)

        # 页脚区域
        noise_layout.addWidget(QLabel("页脚区域 (%):"))
        f_layout = QHBoxLayout()
        self.footer_slider = QSlider(Qt.Horizontal)
        self.footer_slider.setRange(60, 100)
        self.footer_slider.setValue(90)
        self.footer_slider.setEnabled(False)
        f_layout.addWidget(self.footer_slider)
        self.footer_value_label = QLabel("85%")
        self.footer_slider.valueChanged.connect(lambda v: self.footer_value_label.setText(f"{v}%"))
        f_layout.addWidget(self.footer_value_label)
        noise_layout.addLayout(f_layout)

        # 重复率阈值
        noise_layout.addWidget(QLabel("重复率阈值 (%):"))
        r_layout = QHBoxLayout()
        self.threshold_slider = QSlider(Qt.Horizontal)
        self.threshold_slider.setRange(10, 95)
        self.threshold_slider.setValue(30)
        self.threshold_slider.setEnabled(False)
        r_layout.addWidget(self.threshold_slider)
        self.threshold_value_label = QLabel("30%")
        self.threshold_slider.valueChanged.connect(lambda v: self.threshold_value_label.setText(f"{v}%"))
        r_layout.addWidget(self.threshold_value_label)
        noise_layout.addLayout(r_layout)

        noise_group.setLayout(noise_layout)
        layout.addWidget(noise_group)

        layout.addStretch()
        return layout

    def _create_action_section(self) -> QHBoxLayout:
        """创建操作按钮与进度条区域"""
        layout = QHBoxLayout()

        # 左侧：进度信息与进度条
        progress_layout = QVBoxLayout()
        self.progress_label = QLabel("就绪")
        progress_layout.addWidget(self.progress_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        layout.addLayout(progress_layout, stretch=1)

        # 右侧：操作按钮
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_search)
        layout.addWidget(self.cancel_btn)

        self.search_btn = QPushButton("🔍 开始搜索")
        self.search_btn.setFont(QFont("", -1, QFont.Bold))
        self.search_btn.setMinimumSize(120, 35)
        self.search_btn.clicked.connect(self._start_search)
        layout.addWidget(self.search_btn)

        return layout

    # ==================== 界面交互状态控制 ====================

    def _set_ui_state(self, is_searching: bool):
        """在搜索过程中启用/禁用相关控件"""
        state = not is_searching
        self.pdf_btn.setEnabled(state)
        self.output_btn.setEnabled(state)
        self.open_dir_btn.setEnabled(state)
        self.save_cfg_btn.setEnabled(state)
        self.load_cfg_btn.setEnabled(state)
        self.context_slider.setEnabled(state)
        self.front_entry.setEnabled(state)
        self.txt_check.setEnabled(state)
        self.excel_check.setEnabled(state)
        self.sort_check.setEnabled(state)
        
        self.search_btn.setEnabled(state)
        self.cancel_btn.setEnabled(is_searching)
        
        self.noise_check.setEnabled(state)
        self._toggle_noise_params()

    def _toggle_noise_params(self):
        """根据噪声检测开关启用/禁用参数控件"""
        enabled = self.noise_check.isChecked() and self.noise_check.isEnabled()
        self.header_slider.setEnabled(enabled)
        self.footer_slider.setEnabled(enabled)
        self.threshold_slider.setEnabled(enabled)

    # ==================== 文件操作 ====================

    def _select_pdf(self):
        """选择PDF文件"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择PDF文件", "", "PDF文件 (*.pdf);;所有文件 (*.*)")
        if file_path:
            self.pdf_entry.setText(file_path)
            if not self.output_entry.text():
                self.output_entry.setText(str(Path(file_path).parent))

    def _select_output_dir(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_entry.setText(dir_path)

    def _open_output_dir(self):
        """打开输出目录"""
        output_dir = self.output_entry.text()
        if not output_dir:
            pdf_path = self.pdf_entry.text()
            if pdf_path and os.path.exists(pdf_path):
                output_dir = str(Path(pdf_path).parent)
            else:
                output_dir = ""

        if output_dir and os.path.exists(output_dir):
            # 跨平台打开目录
            if sys.platform == "win32":
                os.startfile(output_dir)
            elif sys.platform == "darwin":
                os.system(f'open "{output_dir}"')
            else:
                os.system(f'xdg-open "{output_dir}"')
        else:
            QMessageBox.warning(self, "提示", "输出目录不存在，请先执行搜索或手动指定目录")

    # ==================== 关键字管理 ====================

    def _add_keyword(self):
        """添加关键字"""
        keyword = self.keyword_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "提示", "请输入关键字")
            return

        try:
            score = int(self.score_input.text())
        except ValueError:
            score = 1

        if score < 1:
            score = 1

        if keyword in self.keywords:
            QMessageBox.warning(self, "提示", f"关键字 '{keyword}' 已存在")
            return

        self.keywords[keyword] = score
        self._refresh_keyword_list()

        self.keyword_input.clear()
        self.score_input.setText("1")
        self.keyword_input.setFocus()

    def _delete_keyword(self):
        """删除选中的关键字"""
        row = self.keyword_table.currentRow()
        if row >= 0:
            keyword_item = self.keyword_table.item(row, 0)
            if keyword_item:
                keyword = keyword_item.text()
                if keyword in self.keywords:
                    del self.keywords[keyword]
                    self._refresh_keyword_list()
        else:
            QMessageBox.warning(self, "提示", "请先选择要删除的关键字")

    def _clear_keywords(self):
        """清空所有关键字"""
        if self.keywords and QMessageBox.question(self, "确认", "确定要清空所有关键字吗？") == QMessageBox.Yes:
            self.keywords.clear()
            self._refresh_keyword_list()

    def _refresh_keyword_list(self):
        """刷新关键字列表显示"""
        self.keyword_table.setRowCount(0)
        for keyword, score in self.keywords.items():
            row_position = self.keyword_table.rowCount()
            self.keyword_table.insertRow(row_position)
            self.keyword_table.setItem(row_position, 0, QTableWidgetItem(keyword))
            self.keyword_table.setItem(row_position, 1, QTableWidgetItem(str(score)))

    def _save_config(self):
        """保存配置到JSON文件"""
        if not self.keywords:
            QMessageBox.warning(self, "提示", "没有关键字可保存")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "保存配置", "", "JSON文件 (*.json);;所有文件 (*.*)")
        if file_path:
            config = {
                "keywords": self.keywords,
                "settings": {
                    "context_chars": self.context_slider.value(),
                    "front_window": int(self.front_entry.text() or 0),
                    "output_txt": self.txt_check.isChecked(),
                    "output_excel": self.excel_check.isChecked(),
                    "sort_by_importance": self.sort_check.isChecked(),
                    "auto_clean_noise": self.noise_check.isChecked(),
                    "header_ratio": self.header_slider.value() / 100.0,
                    "footer_ratio": self.footer_slider.value() / 100.0,
                    "repeat_threshold": self.threshold_slider.value() / 100.0
                }
            }
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                QMessageBox.information(self, "成功", f"配置已保存到:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存配置失败:\n{str(e)}")

    def _load_config(self):
        """从JSON文件加载配置"""
        file_path, _ = QFileDialog.getOpenFileName(self, "加载配置", "", "JSON文件 (*.json);;所有文件 (*.*)")
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                if "keywords" in config and isinstance(config["keywords"], dict):
                    self.keywords = {str(k): int(v) for k, v in config["keywords"].items()}
                    self._refresh_keyword_list()

                if "settings" in config and isinstance(config["settings"], dict):
                    settings = config["settings"]
                    if "context_chars" in settings: self.context_slider.setValue(settings["context_chars"])
                    if "front_window" in settings: self.front_entry.setText(str(settings["front_window"]))
                    if "output_txt" in settings: self.txt_check.setChecked(settings["output_txt"])
                    if "output_excel" in settings: self.excel_check.setChecked(settings["output_excel"])
                    if "sort_by_importance" in settings: self.sort_check.setChecked(settings["sort_by_importance"])
                    if "auto_clean_noise" in settings: self.noise_check.setChecked(settings["auto_clean_noise"])
                    if "header_ratio" in settings: self.header_slider.setValue(int(settings["header_ratio"] * 100))
                    if "footer_ratio" in settings: self.footer_slider.setValue(int(settings["footer_ratio"] * 100))
                    if "repeat_threshold" in settings: self.threshold_slider.setValue(int(settings["repeat_threshold"] * 100))

                QMessageBox.information(self, "成功", "配置加载成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载配置失败:\n{str(e)}")

    # ==================== 搜索操作 ====================

    def _start_search(self):
        """开始搜索"""
        pdf_path = self.pdf_entry.text()
        if not pdf_path:
            QMessageBox.warning(self, "提示", "请选择PDF文件")
            return

        if not os.path.exists(pdf_path):
            QMessageBox.warning(self, "提示", "PDF文件不存在")
            return

        if not self.keywords:
            QMessageBox.warning(self, "提示", "请添加至少一个关键字")
            return

        if not self.txt_check.isChecked() and not self.excel_check.isChecked():
            QMessageBox.warning(self, "提示", "请至少选择一种输出格式")
            return

        # 准备输出路径
        pdf_name = Path(pdf_path).stem
        output_dir = self.output_entry.text() or str(Path(pdf_path).parent)
        os.makedirs(output_dir, exist_ok=True)

        output_file = os.path.join(output_dir, f"{pdf_name}_keywords.txt") if self.txt_check.isChecked() else None
        excel_file = os.path.join(output_dir, f"{pdf_name}_keywords.xlsx") if self.excel_check.isChecked() else None

        # 更新UI状态并启动进度条加载动画
        self.is_searching = True
        self._set_ui_state(True)
        self.progress_bar.setRange(0, 0)  # 切换为不确定进度条(滚动动画)
        self.progress_label.setText("正在搜索...")

        # 实例化并启动工作线程
        self.search_worker = SearchWorker({
            "pdf_path": pdf_path,
            "context_rich": self.context_slider.value(),
            "front_window": int(self.front_entry.text() or 0),
            "keywords": self.keywords.copy(),
            "output_file": output_file,
            "excel_file": excel_file,
            "auto_clean_noise": self.noise_check.isChecked(),
            "header_ratio": self.header_slider.value() / 100.0,
            "footer_ratio": self.footer_slider.value() / 100.0,
            "repeat_threshold": self.threshold_slider.value() / 100.0
        })
        self.search_worker.search_finished.connect(self._on_search_finished)
        self.search_worker.search_error.connect(self._on_search_error)
        self.search_worker.start()

    def _cancel_search(self):
        """取消搜索"""
        if self.search_worker and self.search_worker.isRunning():
            self.search_worker.cancel()
            self.progress_label.setText("正在取消... (等待当前处理完成)")
            self.cancel_btn.setEnabled(False)

    def _on_search_finished(self, results: dict):
        """搜索完成回调"""
        self.is_searching = False
        self._set_ui_state(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)

        total_matches = results.get('total_matches', 0)
        message = f"搜索完成，找到 {total_matches} 处匹配，请前往输出目录查看文件。"
        self.progress_label.setText(message)
        
        noise_info = results.get('noise_info', [])

        if noise_info:
            self._show_noise_result(noise_info, message)
        else:
            QMessageBox.information(self, "完成", message)

    def _on_search_error(self, error_msg: str):
        """搜索失败回调"""
        self.is_searching = False
        self._set_ui_state(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_label.setText("搜索失败")
        QMessageBox.critical(self, "错误", f"搜索失败:\n{error_msg}")

    def _show_noise_result(self, noise_info: list, search_message: str):
        """显示噪声检测结果"""
        dialog = NoiseResultDialog(noise_info, self)
        if dialog.exec() == QDialog.Accepted:
            QMessageBox.information(self, "完成", search_message)


def main():
    app = QApplication(sys.argv)

    # 设置默认字体，避免Windows上的字体警告
    from PySide6.QtGui import QFontDatabase
    font_id = QFontDatabase.addApplicationFont("C:/Windows/Fonts/msyh.ttc")
    if font_id != -1:
        font_families = QFontDatabase.applicationFontFamilies(font_id)
        if "微软雅黑" in font_families:
            app.setFont(QFont("微软雅黑"))
        else:
            app.setFont(QFont("Segoe UI"))
    else:
        app.setFont(QFont("Segoe UI"))

    window = PDFKeywordFinderApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()