"""
PDF关键字搜索工具 - 图形界面
基于 CustomTkinter 实现现代化界面
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import os
from pathlib import Path
from typing import Dict, Optional
import pdf_keyword_finder


class PDFKeywordFinderApp(ctk.CTk):
    """PDF关键字搜索工具主窗口"""

    def __init__(self):
        super().__init__()

        # 窗口基本设置
        self.title("解决方案中心工具")
        self.geometry("950x750")
        self.minsize(850, 650)

        # 设置主题
        ctk.set_appearance_mode("system")  # 跟随系统主题
        ctk.set_default_color_theme("blue")  # 蓝色主题

        # 初始化变量
        self.pdf_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self.context_chars = ctk.IntVar(value=200)
        self.front_window = ctk.IntVar(value=0)
        self.output_txt = ctk.BooleanVar(value=True)
        self.output_excel = ctk.BooleanVar(value=True)
        self.sort_by_importance = ctk.BooleanVar(value=True)

        # 关键字数据
        self.keywords: Dict[str, int] = {}

        # 分数输入变量（使用StringVar避免空值异常）
        self.new_keyword = ctk.StringVar()
        self.new_score_str = ctk.StringVar(value="1")

        # 搜索线程和取消标志
        self.search_thread: Optional[threading.Thread] = None
        self.cancel_flag = False
        self.is_searching = False
        self._search_success = False
        self._search_message = ""

        # 创建界面
        self._create_widgets()

        # 居中窗口
        self._center_window()

    def _center_window(self):
        """窗口居中"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _get_row_colors(self, idx):
        """根据当前主题获取交替行颜色"""
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            bg_color = "#3a3a3a" if idx % 2 == 0 else "#454545"
            text_color = "white"
        else:
            bg_color = "#f0f0f0" if idx % 2 == 0 else "#e0e0e0"
            text_color = "black"
        return bg_color, text_color

    def _create_widgets(self):
        """创建所有界面组件"""
        # 主容器，添加内边距
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # 1. 优先固定底部区域，确保操作按钮和进度条始终可见，不被挤出屏幕
        self._create_action_section()

        # 2. 标题
        title_label = ctk.CTkLabel(
            self.main_frame,
            text="📄 解决方案中心PDF标书分析工具",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(fill="x", pady=(0, 15))

        # 3. 文件选择区域
        self._create_file_selection()

        # 4. 中间区域：左侧关键字管理 + 右侧设置面板 (自动占据剩余空间)
        self._create_middle_section()

    def _create_file_selection(self):
        """创建文件选择区域"""
        file_frame = ctk.CTkFrame(self.main_frame)
        file_frame.pack(fill="x", pady=(0, 10))

        # PDF文件选择
        pdf_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        pdf_frame.pack(fill="x", padx=10, pady=10)

        pdf_label = ctk.CTkLabel(pdf_frame, text="PDF文件:", width=80, anchor="w")
        pdf_label.pack(side="left")

        self.pdf_entry = ctk.CTkEntry(pdf_frame, textvariable=self.pdf_path, placeholder_text="请选择PDF文件...")
        self.pdf_entry.pack(side="left", fill="x", expand=True, padx=(10, 10))

        self.pdf_btn = ctk.CTkButton(pdf_frame, text="选择文件", width=80, command=self._select_pdf)
        self.pdf_btn.pack(side="left")

        # 输出目录选择
        output_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        output_frame.pack(fill="x", padx=10, pady=(0, 10))

        output_label = ctk.CTkLabel(output_frame, text="输出目录:", width=80, anchor="w")
        output_label.pack(side="left")

        self.output_entry = ctk.CTkEntry(output_frame, textvariable=self.output_dir, placeholder_text="默认与PDF文件同目录")
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(10, 10))

        self.output_btn = ctk.CTkButton(output_frame, text="选择目录", width=80, command=self._select_output_dir)
        self.output_btn.pack(side="left", padx=(0, 5))

        # 将打开目录按钮移至此处
        self.open_dir_btn = ctk.CTkButton(output_frame, text="打开目录", width=80, command=self._open_output_dir)
        self.open_dir_btn.pack(side="left")

    def _create_middle_section(self):
        """创建中间区域：关键字管理 + 设置面板"""
        middle_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        middle_frame.pack(fill="both", expand=True, pady=10)

        # 左侧：关键字管理
        self._create_keyword_section(middle_frame)

        # 右侧：设置面板
        self._create_settings_section(middle_frame)

    def _create_keyword_section(self, parent):
        """创建关键字管理区域"""
        keyword_frame = ctk.CTkFrame(parent)
        keyword_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        # 标题
        header = ctk.CTkLabel(keyword_frame, text="关键字管理", font=ctk.CTkFont(size=14, weight="bold"))
        header.pack(pady=10)
        # 批量操作按钮 - 先 pack 并置于底部，防止被中间的 expand 挤出视野
        btn_frame = ctk.CTkFrame(keyword_frame, fg_color="transparent")
        btn_frame.pack(side="bottom", fill="x", padx=10, pady=(5, 10))
        self.save_cfg_btn = ctk.CTkButton(btn_frame, text="保存配置", command=self._save_config)
        self.save_cfg_btn.pack(side="left", padx=5)
        self.load_cfg_btn = ctk.CTkButton(btn_frame, text="加载配置", command=self._load_config)
        self.load_cfg_btn.pack(side="left", padx=5)
        self.clear_btn = ctk.CTkButton(btn_frame, text="清空", fg_color="gray", hover_color="gray30", command=self._clear_keywords)
        self.clear_btn.pack(side="left", padx=5)
        # 添加关键字区域 - 紧接在批量操作按钮上方
        add_frame = ctk.CTkFrame(keyword_frame, fg_color="transparent")
        add_frame.pack(side="bottom", fill="x", padx=10, pady=5)
        self.keyword_entry = ctk.CTkEntry(add_frame, textvariable=self.new_keyword, placeholder_text="输入关键字", width=140)
        self.keyword_entry.pack(side="left", padx=(0, 5))
        self.score_entry = ctk.CTkEntry(add_frame, textvariable=self.new_score_str, placeholder_text="分数", width=60)
        self.score_entry.pack(side="left", padx=(0, 5))
        self.add_btn = ctk.CTkButton(add_frame, text="添加", width=60, command=self._add_keyword)
        self.add_btn.pack(side="left")
        # 绑定回车键添加关键字
        self.keyword_entry.bind("<Return>", lambda e: self._add_keyword())
        self.score_entry.bind("<Return>", lambda e: self._add_keyword())
        # 关键字表格区域 - 最后 pack，自动填充上方剩余空间
        table_frame = ctk.CTkFrame(keyword_frame)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        # 表头
        header_frame = ctk.CTkFrame(table_frame, fg_color="gray30", corner_radius=0)
        header_frame.pack(fill="x")
        ctk.CTkLabel(header_frame, text="关键字", width=150, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="分数", width=60, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="操作", width=60, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        # 创建可滚动的关键字列表
        self.keyword_scroll = ctk.CTkScrollableFrame(table_frame, fg_color="transparent")
        self.keyword_scroll.pack(fill="both", expand=True)
        self._refresh_keyword_list()

    def _create_settings_section(self, parent):
        """创建设置面板区域"""
        settings_frame = ctk.CTkFrame(parent, width=280)
        settings_frame.pack(side="right", fill="both", padx=(5, 0))
        settings_frame.pack_propagate(False)  # 固定宽度

        # 标题
        header = ctk.CTkLabel(settings_frame, text="搜索设置", font=ctk.CTkFont(size=14, weight="bold"))
        header.pack(pady=10)

        # 搜索参数
        search_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=15, pady=5)

        # 上下文字符数
        context_label = ctk.CTkLabel(search_frame, text="上下文丰富度:", anchor="w")
        context_label.pack(fill="x")
        context_frame = ctk.CTkFrame(search_frame, fg_color="transparent")
        context_frame.pack(fill="x", pady=(0, 10))

        self.context_slider = ctk.CTkSlider(
            context_frame,
            from_=50, to=1000,
            variable=self.context_chars,
            number_of_steps=19
        )
        self.context_slider.pack(side="left", fill="x", expand=True)

        context_value = ctk.CTkLabel(context_frame, textvariable=self.context_chars, width=40)
        context_value.pack(side="left", padx=(10, 0))

        # 固定窗口字数
        front_frame = ctk.CTkFrame(search_frame, fg_color="transparent")
        front_frame.pack(fill="x", pady=(0, 10))
        front_label = ctk.CTkLabel(front_frame, text="前窗口字数:", width=80, anchor="w")
        front_label.pack(side="left")

        self.front_entry = ctk.CTkEntry(
            front_frame,
            textvariable=self.front_window,
            width=50,
            placeholder_text="默认 0"
            )
        self.front_entry.pack(side="left", fill="x",  padx=(10, 0))

        # 分隔线
        separator = ctk.CTkFrame(settings_frame, height=2, fg_color="gray50")
        separator.pack(fill="x", padx=15, pady=15)

        # 输出选项
        output_label = ctk.CTkLabel(settings_frame, text="输出选项", font=ctk.CTkFont(size=14, weight="bold"))
        output_label.pack(pady=(0, 10))

        output_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        output_frame.pack(fill="x", padx=15)

        self.txt_check = ctk.CTkCheckBox(output_frame, text="输出 TXT 文件", variable=self.output_txt)
        self.txt_check.pack(fill="x", pady=2)

        self.excel_check = ctk.CTkCheckBox(output_frame, text="输出 Excel 文件", variable=self.output_excel)
        self.excel_check.pack(fill="x", pady=2)

        self.sort_check = ctk.CTkCheckBox(output_frame, text="按重要性排序", variable=self.sort_by_importance)
        self.sort_check.pack(fill="x", pady=2)

    def _create_action_section(self):
        """创建操作按钮和进度条区域"""
        action_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        # 明确指定 side="bottom"，确保该区域固定在底部，不会被中间的 expand=True 挤出视野
        action_frame.pack(side="bottom", fill="x", pady=(15, 0))

        # 按钮区域
        btn_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_frame.pack(pady=5)

        self.search_btn = ctk.CTkButton(
            btn_frame,
            text="🔍 开始搜索",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=120, height=35,
            command=self._start_search
        )
        self.search_btn.pack(side="left", padx=10)

        self.cancel_btn = ctk.CTkButton(
            btn_frame,
            text="取消",
            font=ctk.CTkFont(size=14),
            width=80, height=35,
            fg_color="gray",
            hover_color="gray30",
            state="disabled",
            command=self._cancel_search
        )
        self.cancel_btn.pack(side="left", padx=10)

        # 进度条
        progress_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        progress_frame.pack(fill="x", pady=5)

        self.progress_label = ctk.CTkLabel(progress_frame, text="就绪", anchor="w")
        self.progress_label.pack(fill="x")

        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=5)

    # ==================== 界面交互状态控制 ====================

    def _set_ui_state(self, is_searching):
        """在搜索过程中启用/禁用相关控件"""
        state = "disabled" if is_searching else "normal"
        self.pdf_btn.configure(state=state)
        self.output_btn.configure(state=state)
        self.add_btn.configure(state=state)
        self.save_cfg_btn.configure(state=state)
        self.load_cfg_btn.configure(state=state)
        self.clear_btn.configure(state=state)
        self.open_dir_btn.configure(state=state)
        self.context_slider.configure(state=state)
        self.front_entry.configure(state=state)  # 同步禁用输入框
        self.txt_check.configure(state=state)
        self.excel_check.configure(state=state)
        self.sort_check.configure(state=state)

    # ==================== 文件操作 ====================

    def _select_pdf(self):
        """选择PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            # 自动设置输出目录为PDF所在目录
            if not self.output_dir.get():
                self.output_dir.set(str(Path(file_path).parent))

    def _select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir.set(dir_path)

    def _open_output_dir(self):
        """打开输出目录"""
        output_dir = self.output_dir.get()
        # 如果没有手动指定输出目录，尝试打开PDF所在目录
        if not output_dir:
            pdf_path = self.pdf_path.get()
            if pdf_path and os.path.exists(pdf_path):
                output_dir = str(Path(pdf_path).parent)
            else:
                output_dir = ""

        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            messagebox.showwarning("提示", "输出目录不存在，请先执行搜索或手动指定目录")

    # ==================== 关键字管理 ====================

    def _add_keyword(self):
        """添加关键字"""
        keyword = self.new_keyword.get().strip()

        if not keyword:
            messagebox.showwarning("提示", "请输入关键字")
            return

        # 解析分数，处理空值和非数字输入
        try:
            score = int(self.new_score_str.get())
        except ValueError:
            score = 1

        if score < 1:
            score = 1

        if keyword in self.keywords:
            messagebox.showwarning("提示", f"关键字 '{keyword}' 已存在")
            return

        self.keywords[keyword] = score
        self._refresh_keyword_list()

        # 清空输入并聚焦
        self.new_keyword.set("")
        self.new_score_str.set("1")
        self.keyword_entry.focus()

    def _delete_keyword(self, keyword: str):
        """删除关键字"""
        if keyword in self.keywords:
            del self.keywords[keyword]
            self._refresh_keyword_list()

    def _clear_keywords(self):
        """清空所有关键字"""
        if self.keywords and messagebox.askyesno("确认", "确定要清空所有关键字吗？"):
            self.keywords.clear()
            self._refresh_keyword_list()

    def _refresh_keyword_list(self):
        """刷新关键字列表显示"""
        for widget in self.keyword_scroll.winfo_children():
            widget.destroy()

        if not self.keywords:
            placeholder = ctk.CTkLabel(self.keyword_scroll, text="暂无关键字，请在下方添加", text_color="gray")
            placeholder.pack(pady=20)
            return

        for idx, (keyword, score) in enumerate(self.keywords.items()):
            bg_color, text_color = self._get_row_colors(idx)
            row = ctk.CTkFrame(self.keyword_scroll, fg_color=bg_color, corner_radius=5)
            row.pack(fill="x", pady=2, padx=2)

            ctk.CTkLabel(row, text=keyword, width=150, anchor="w", text_color=text_color).pack(side="left", padx=5, pady=3)
            ctk.CTkLabel(row, text=str(score), width=60, text_color=text_color).pack(side="left", padx=5, pady=3)

            delete_btn = ctk.CTkButton(
                row,
                text="删除",
                width=50,
                height=24,
                fg_color="#e74c3c",
                hover_color="#c0392b",
                text_color="white",
                command=lambda k=keyword: self._delete_keyword(k)
            )
            delete_btn.pack(side="left", padx=5, pady=3)

    def _save_config(self):
        """保存配置到JSON文件"""
        if not self.keywords:
            messagebox.showwarning("提示", "没有关键字可保存")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存配置",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if file_path:
            config = {
                "keywords": self.keywords,
                "settings": {
                    "context_chars": self.context_chars.get(),
                    "front_window": self.front_window.get(),
                    "output_txt": self.output_txt.get(),
                    "output_excel": self.output_excel.get(),
                    "sort_by_importance": self.sort_by_importance.get()
                }
            }
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"配置已保存到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败:\n{str(e)}")

    def _load_config(self):
        """从JSON文件加载配置"""
        file_path = filedialog.askopenfilename(
            title="加载配置",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                # 加载关键字
                if "keywords" in config and isinstance(config["keywords"], dict):
                    self.keywords = {str(k): int(v) for k, v in config["keywords"].items()}
                    self._refresh_keyword_list()

                # 加载设置
                if "settings" in config and isinstance(config["settings"], dict):
                    settings = config["settings"]
                    if "context_chars" in settings: self.context_chars.set(settings["context_chars"])
                    if "front_window" in settings: self.front_window.set(settings["front_window"])
                    if "output_txt" in settings: self.output_txt.set(settings["output_txt"])
                    if "output_excel" in settings: self.output_excel.set(settings["output_excel"])
                    if "sort_by_importance" in settings: self.sort_by_importance.set(settings["sort_by_importance"])

                messagebox.showinfo("成功", "配置加载成功")
            except Exception as e:
                messagebox.showerror("错误", f"加载配置失败:\n{str(e)}")

    # ==================== 搜索操作 ====================

    def _start_search(self):
        """开始搜索"""
        # 验证输入
        if not self.pdf_path.get():
            messagebox.showwarning("提示", "请选择PDF文件")
            return

        if not os.path.exists(self.pdf_path.get()):
            messagebox.showwarning("提示", "PDF文件不存在")
            return

        if not self.keywords:
            messagebox.showwarning("提示", "请添加至少一个关键字")
            return

        if not self.output_txt.get() and not self.output_excel.get():
            messagebox.showwarning("提示", "请至少选择一种输出格式")
            return

        # 准备输出路径
        pdf_name = Path(self.pdf_path.get()).stem
        output_dir = self.output_dir.get() or str(Path(self.pdf_path.get()).parent)

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        output_file = None
        excel_file = None

        if self.output_txt.get():
            output_file = os.path.join(output_dir, f"{pdf_name}_keywords.txt")

        if self.output_excel.get():
            excel_file = os.path.join(output_dir, f"{pdf_name}_keywords.xlsx")

        # 更新UI状态
        self._set_ui_state(True)
        self.search_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.is_searching = True
        self.cancel_flag = False
        self.progress_bar.set(0)
        self.progress_label.configure(text="正在搜索...")

        # 在新线程中执行搜索
        self.search_thread = threading.Thread(
            target=self._search_thread,
            args=(
                self.pdf_path.get(),
                self.context_chars.get(),
                self.front_window.get(),
                self.keywords.copy(),
                output_file,
                excel_file,
                output_dir
            ),
            daemon=True
        )
        self.search_thread.start()

        # 定时检查搜索状态
        self.after(100, self._check_search_status)

    def _search_thread(self, pdf_path: str, context_rich:int, front_window : int, keywords: Dict[str, int], output_file: Optional[str], excel_file: Optional[str], output_dir: str):
        """搜索线程"""
        try:
            # 调用搜索函数
            results = pdf_keyword_finder.find_keywords_in_pdf(
                pdf_path=pdf_path,
                keywords=keywords,
                context_rich=context_rich,
                front_window = front_window,
                output_file=output_file,
                excel_file=excel_file,
            )

            self._search_success = True
            self._search_message = f"搜索完成，找到 {results.get('total_matches', 0)} 处匹配，请前往输出目录查看文件。"

        except Exception as e:
            self._search_success = False
            self._search_message = f"搜索失败: {str(e)}"

    def _check_search_status(self):
        """检查搜索状态"""
        if self.search_thread is None:
            return

        if self.search_thread.is_alive():
            # 更新进度动画 (不确定进度条效果)
            current = self.progress_bar.get()
            self.progress_bar.set((current + 0.02) % 1)
            self.after(100, self._check_search_status)
        else:
            # 搜索完成
            self.is_searching = False
            self._set_ui_state(False)
            self.search_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")
            self.progress_bar.set(1)

            self.progress_label.configure(text=self._search_message)
            if self._search_success:
                messagebox.showinfo("完成", self._search_message)

    def _cancel_search(self):
        """取消搜索"""
        self.cancel_flag = True
        self.is_searching = False
        self._set_ui_state(False)
        self.progress_label.configure(text="正在取消... (等待当前处理完成)")
        self.search_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")


def main():
    """主函数"""
    app = PDFKeywordFinderApp()
    app.mainloop()


if __name__ == "__main__":
    main()