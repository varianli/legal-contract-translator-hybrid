import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import hybrid_markdown_run_translator as core


MAC_APP_VERSION = "v1.27-mac"


def friendly_error(exc: Exception) -> str:
    text = str(exc)
    lower = text.lower()
    if "401" in text or "authentication" in lower or "invalid api" in lower:
        return (
            "API Key 认证失败。请确认：\n"
            "1. API 类型选对了，例如 DeepSeek Key 就选 DeepSeek。\n"
            "2. API Key 没有复制多余空格、换行或中文符号。\n"
            "3. Base URL 是否正确：DeepSeek 应为 https://api.deepseek.com。\n\n"
            f"原始错误：{text}"
        )
    return text


class MacHybridApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.main_thread_id = threading.get_ident()
        self.root.title(f"复合方法 Mac版 - {MAC_APP_VERSION}")
        self.root.geometry("980x760")
        self.root.minsize(900, 680)

        config = core.load_config()
        provider = config.get("provider", "DeepSeek")
        if provider not in core.PROVIDER_DEFAULTS:
            provider = "DeepSeek"
        defaults = core.PROVIDER_DEFAULTS[provider]

        self.provider = tk.StringVar(value=provider)
        self.english_font = tk.StringVar(value=core.valid_english_font(config.get("english_font", core.DEFAULT_ENGLISH_FONT)))
        self.chinese_font = tk.StringVar(value=core.valid_chinese_font(config.get("chinese_font", core.DEFAULT_CHINESE_FONT)))
        self.remember_key = tk.BooleanVar(value=bool(config.get("api_key")))
        self.file_paths: list[Path] = []
        self.job_widgets: dict[str, dict] = {}
        self.job_cancel_events: dict[str, threading.Event] = {}
        self.job_results: list[tuple[Path, tuple, bool]] = []
        self.job_errors: list[tuple[Path, str]] = []
        self.active_jobs = 0
        self.is_running = False

        self.build_ui(defaults, config)
        self.log(f"{MAC_APP_VERSION} 已启动。Mac 专用界面不使用拖拽插件，请使用“添加文件”。")

    def build_ui(self, defaults: dict, config: dict) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.page_canvas = tk.Canvas(self.root, highlightthickness=0)
        self.page_scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.page_canvas.yview)
        self.page_canvas.configure(yscrollcommand=self.page_scrollbar.set)
        self.page_canvas.grid(row=0, column=0, sticky="nsew")
        self.page_scrollbar.grid(row=0, column=1, sticky="ns")

        self.content = ttk.Frame(self.page_canvas)
        self.content_window = self.page_canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.content.columnconfigure(0, weight=1)
        self.content.bind("<Configure>", self.on_content_configure)
        self.page_canvas.bind("<Configure>", self.on_page_canvas_configure)
        self.bind_page_scroll()

        api = ttk.LabelFrame(self.content, text="API 设置", padding=12)
        api.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 8))
        api.columnconfigure(1, weight=1)

        ttk.Label(api, text="API 类型").grid(row=0, column=0, sticky="w", pady=4)
        provider_row = ttk.Frame(api)
        provider_row.grid(row=0, column=1, sticky="w", pady=4)
        for provider_name in core.PROVIDER_DEFAULTS:
            ttk.Radiobutton(provider_row, text=provider_name, variable=self.provider, value=provider_name, command=self.on_provider_change).pack(side="left", padx=(0, 18))

        self.key_label = ttk.Label(api, text=defaults["key_label"])
        self.key_label.grid(row=1, column=0, sticky="w", pady=4)
        self.api_key_entry = ttk.Entry(api, show="*", width=80)
        self.api_key_entry.grid(row=1, column=1, sticky="ew", pady=4)
        self.api_key_entry.insert(0, config.get("api_key", ""))

        ttk.Label(api, text="Base URL").grid(row=2, column=0, sticky="w", pady=4)
        self.base_url_entry = ttk.Entry(api, width=80)
        self.base_url_entry.grid(row=2, column=1, sticky="ew", pady=4)
        self.base_url_entry.insert(0, config.get("base_url", defaults["base_url"]))

        ttk.Label(api, text="模型").grid(row=3, column=0, sticky="w", pady=4)
        self.model_entry = ttk.Entry(api, width=80)
        self.model_entry.grid(row=3, column=1, sticky="ew", pady=4)
        self.model_entry.insert(0, config.get("model", defaults["model"]))

        font_row = ttk.Frame(api)
        font_row.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 2))
        ttk.Label(font_row, text="英文字体").pack(side="left")
        for font_name in core.ENGLISH_FONT_OPTIONS:
            ttk.Radiobutton(font_row, text=font_name, variable=self.english_font, value=font_name).pack(side="left", padx=(10, 18))
        ttk.Label(font_row, text=f"中文字体（数字固定 {core.DIGIT_FONT}）").pack(side="left", padx=(18, 0))
        for font_name in core.CHINESE_FONT_OPTIONS:
            ttk.Radiobutton(font_row, text=font_name, variable=self.chinese_font, value=font_name).pack(side="left", padx=(10, 18))

        key_row = ttk.Frame(api)
        key_row.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Checkbutton(key_row, text="记住 API Key（明文保存在本机目录）", variable=self.remember_key).pack(side="left")
        ttk.Button(key_row, text="测试 API Key", command=self.test_api_key).pack(side="right")

        files = ttk.LabelFrame(self.content, text="文件与输出", padding=12)
        files.grid(row=1, column=0, sticky="ew", padx=14, pady=8)
        files.columnconfigure(0, weight=1)
        file_buttons = ttk.Frame(files)
        file_buttons.grid(row=0, column=0, sticky="ew")
        ttk.Button(file_buttons, text="添加文件", command=self.choose_files).pack(side="left")
        ttk.Button(file_buttons, text="删除选中", command=self.remove_selected_files).pack(side="left", padx=8)
        ttk.Button(file_buttons, text="清空列表", command=self.clear_files).pack(side="left")

        self.file_list_box = tk.Listbox(files, height=5, selectmode="extended", exportselection=False)
        self.file_list_box.grid(row=1, column=0, sticky="ew", pady=(8, 10))
        self.file_list_box.bind("<Delete>", lambda _event: self.remove_selected_files())
        self.refresh_file_list()

        output_row = ttk.Frame(files)
        output_row.grid(row=2, column=0, sticky="ew")
        output_row.columnconfigure(1, weight=1)
        ttk.Label(output_row, text="输出目录").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.output_dir_entry = ttk.Entry(output_row)
        self.output_dir_entry.grid(row=0, column=1, sticky="ew")
        self.output_dir_entry.insert(0, str(core.OUTPUT_DIR))
        ttk.Button(output_row, text="选择目录", command=self.choose_output_dir).grid(row=0, column=2, padx=(8, 0))

        actions = ttk.Frame(self.content, padding=(14, 4))
        actions.grid(row=2, column=0, sticky="ew")
        actions.columnconfigure(0, weight=1)
        self.start_button = ttk.Button(actions, text="开始翻译所选文件", command=self.start)
        self.start_button.grid(row=0, column=0, sticky="ew")
        self.stop_button = ttk.Button(actions, text="全部中止并导出当前进度", command=self.request_cancel_all, state="disabled")
        self.stop_button.grid(row=0, column=1, padx=(8, 0))

        progress_outer = ttk.LabelFrame(self.content, text="文件进度", padding=10)
        progress_outer.grid(row=3, column=0, sticky="nsew", padx=14, pady=8)
        progress_outer.columnconfigure(0, weight=1)
        self.jobs_canvas = tk.Canvas(progress_outer, height=230, highlightthickness=0)
        self.jobs_scrollbar = ttk.Scrollbar(progress_outer, orient="vertical", command=self.jobs_canvas.yview)
        self.jobs_frame = ttk.Frame(self.jobs_canvas)
        self.jobs_window = self.jobs_canvas.create_window((0, 0), window=self.jobs_frame, anchor="nw")
        self.jobs_canvas.configure(yscrollcommand=self.jobs_scrollbar.set)
        self.jobs_canvas.grid(row=0, column=0, sticky="nsew")
        self.jobs_scrollbar.grid(row=0, column=1, sticky="ns")
        self.jobs_frame.bind("<Configure>", lambda _e: self.jobs_canvas.configure(scrollregion=self.jobs_canvas.bbox("all")))
        self.jobs_canvas.bind("<Configure>", lambda e: self.jobs_canvas.itemconfigure(self.jobs_window, width=e.width))
        self.empty_jobs_label = ttk.Label(self.jobs_frame, text="开始后会在这里显示每个文件的翻译进度和复核进度。")
        self.empty_jobs_label.pack(anchor="w", pady=8)

        log_outer = ttk.LabelFrame(self.content, text="日志", padding=8)
        log_outer.grid(row=4, column=0, sticky="nsew", padx=14, pady=(8, 14))
        log_outer.columnconfigure(0, weight=1)
        log_outer.rowconfigure(0, weight=1)
        self.log_box = tk.Text(log_outer, height=10, wrap="word")
        self.log_box.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_outer, orient="vertical", command=self.log_box.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_box.configure(yscrollcommand=log_scroll.set)

    def on_content_configure(self, _event=None) -> None:
        self.page_canvas.configure(scrollregion=self.page_canvas.bbox("all"))

    def on_page_canvas_configure(self, event) -> None:
        self.page_canvas.itemconfigure(self.content_window, width=event.width)

    def bind_page_scroll(self) -> None:
        self.root.bind_all("<MouseWheel>", self.on_page_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self.on_page_mousewheel, add="+")
        self.root.bind_all("<Button-5>", self.on_page_mousewheel, add="+")

    def on_page_mousewheel(self, event):
        if event.widget is self.log_box:
            return
        if getattr(event, "num", None) == 4:
            units = -3
        elif getattr(event, "num", None) == 5:
            units = 3
        else:
            delta = getattr(event, "delta", 0)
            if delta == 0:
                return
            units = -3 if delta > 0 else 3
        self.page_canvas.yview_scroll(units, "units")
        return "break"

    def get_api_values(self) -> tuple[str, str, str, str]:
        provider = self.provider.get().strip() or "DeepSeek"
        if provider not in core.PROVIDER_DEFAULTS:
            provider = "DeepSeek"
            self.provider.set(provider)
        api_key = self.api_key_entry.get().strip()
        base_url = self.base_url_entry.get().strip()
        model = self.model_entry.get().strip()
        if provider == "DeepSeek" and not base_url:
            base_url = core.PROVIDER_DEFAULTS["DeepSeek"]["base_url"]
            self.set_entry(self.base_url_entry, base_url)
        return provider, api_key, base_url, model

    def set_entry(self, entry: ttk.Entry, value: str) -> None:
        entry.delete(0, tk.END)
        entry.insert(0, value or "")

    def on_provider_change(self) -> None:
        provider = self.provider.get()
        defaults = core.PROVIDER_DEFAULTS.get(provider, core.PROVIDER_DEFAULTS["DeepSeek"])
        self.key_label.configure(text=defaults["key_label"])
        self.set_entry(self.base_url_entry, defaults["base_url"])
        self.set_entry(self.model_entry, defaults["model"])

    def choose_files(self) -> None:
        if self.is_running:
            messagebox.showerror("正在翻译", "翻译进行中不能添加文件。")
            return
        paths = filedialog.askopenfilenames(title="选择一个或多个 Word 合同", filetypes=[("Word 文档", "*.docx")])
        self.add_files(paths)

    def add_files(self, paths) -> None:
        existing = {str(path.resolve()).lower() for path in self.file_paths if path.exists()}
        added = 0
        for raw in paths or []:
            path = Path(str(raw).strip().strip('"'))
            if path.exists() and path.suffix.lower() == ".docx":
                key = str(path.resolve()).lower()
                if key not in existing:
                    self.file_paths.append(path)
                    existing.add(key)
                    added += 1
        self.refresh_file_list()
        if added:
            self.log(f"已添加 {added} 个文件。")

    def refresh_file_list(self) -> None:
        self.file_list_box.delete(0, tk.END)
        if not self.file_paths:
            self.file_list_box.insert(tk.END, "尚未选择文件。请点击“添加文件”。")
        else:
            for index, path in enumerate(self.file_paths, start=1):
                self.file_list_box.insert(tk.END, f"{index}. {path}")

    def remove_selected_files(self) -> None:
        if self.is_running:
            messagebox.showerror("正在翻译", "翻译进行中不能删除文件。")
            return
        if not self.file_paths:
            return
        selected = list(self.file_list_box.curselection())
        if not selected:
            messagebox.showerror("未选中文件", "请先在文件列表中选中要删除的文件。")
            return
        remove_indices = set(selected)
        self.file_paths = [path for index, path in enumerate(self.file_paths) if index not in remove_indices]
        self.refresh_file_list()
        self.clear_job_rows()

    def clear_files(self) -> None:
        if self.is_running:
            messagebox.showerror("正在翻译", "翻译进行中不能清空文件列表。")
            return
        self.file_paths = []
        self.refresh_file_list()
        self.clear_job_rows()

    def choose_output_dir(self) -> None:
        current = self.output_dir_entry.get().strip() or str(core.OUTPUT_DIR)
        path = filedialog.askdirectory(title="选择输出目录", initialdir=current)
        if path:
            self.set_entry(self.output_dir_entry, path)

    def clear_job_rows(self) -> None:
        for child in self.jobs_frame.winfo_children():
            child.destroy()
        self.job_widgets = {}
        self.empty_jobs_label = ttk.Label(self.jobs_frame, text="开始后会在这里显示每个文件的翻译进度和复核进度。")
        self.empty_jobs_label.pack(anchor="w", pady=8)

    def create_job_row(self, job_id: str, path: Path) -> None:
        if self.empty_jobs_label and self.empty_jobs_label.winfo_exists():
            self.empty_jobs_label.destroy()
        row = ttk.Frame(self.jobs_frame, padding=8, relief="groove")
        row.pack(fill="x", pady=(0, 8))
        row.columnconfigure(0, weight=1)
        ttk.Label(row, text=path.name, font=("TkDefaultFont", 11, "bold")).grid(row=0, column=0, sticky="w")
        cancel = ttk.Button(row, text="中止此文件", command=lambda key=job_id: self.request_cancel_job(key))
        cancel.grid(row=0, column=1, sticky="e", padx=(8, 0))
        ttk.Label(row, text=str(path), wraplength=820).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 6))
        translation_text = tk.StringVar(value="翻译进度：0/100（0.0%）")
        review_text = tk.StringVar(value="复核检查：0/100（0.0%）")
        current_text = tk.StringVar(value="状态：排队中")
        ttk.Label(row, textvariable=translation_text).grid(row=2, column=0, columnspan=2, sticky="w")
        translation_bar = ttk.Progressbar(row, maximum=100, mode="determinate")
        translation_bar.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(2, 6))
        ttk.Label(row, textvariable=review_text).grid(row=4, column=0, columnspan=2, sticky="w")
        review_bar = ttk.Progressbar(row, maximum=100, mode="determinate")
        review_bar.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(2, 6))
        ttk.Label(row, textvariable=current_text, wraplength=850).grid(row=6, column=0, columnspan=2, sticky="w")
        self.job_widgets[job_id] = {
            "cancel": cancel,
            "translation_text": translation_text,
            "translation_bar": translation_bar,
            "review_text": review_text,
            "review_bar": review_bar,
            "current_text": current_text,
        }

    def test_api_key(self) -> None:
        provider, api_key, base_url, model = self.get_api_values()
        if not api_key:
            messagebox.showerror("缺少 API Key", f"请先输入 {provider} API Key。")
            return
        if not model:
            messagebox.showerror("缺少模型", "请填写模型名称。")
            return
        self.log("正在测试 API Key...")
        threading.Thread(target=self._test_api_key_worker, args=(provider, api_key, base_url, model), daemon=True).start()

    def _test_api_key_worker(self, provider: str, api_key: str, base_url: str, model: str) -> None:
        try:
            client = core.build_client(api_key, base_url)
            client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": "Return OK only."}],
                temperature=0,
                max_tokens=8,
            )
        except Exception as exc:
            self.root.after(0, lambda: messagebox.showerror("API Key 测试失败", friendly_error(exc)))
            self.root.after(0, self.log, f"API Key 测试失败：{friendly_error(exc)}")
            return
        self.root.after(0, lambda: messagebox.showinfo("API Key 测试成功", "API Key 可以正常调用。"))
        self.root.after(0, self.log, "API Key 测试成功。")

    def start(self) -> None:
        provider, api_key, base_url, model = self.get_api_values()
        english_font = core.valid_english_font(self.english_font.get().strip())
        chinese_font = core.valid_chinese_font(self.chinese_font.get().strip())
        output_dir = Path((self.output_dir_entry.get().strip().strip('"') or str(core.OUTPUT_DIR)))
        if not api_key:
            messagebox.showerror("缺少 API Key", f"请先输入 {provider} API Key。")
            return
        if not model:
            messagebox.showerror("缺少模型", "请填写模型名称。")
            return
        if not self.file_paths:
            messagebox.showerror("缺少文件", "请先添加至少一个 .docx 文件。")
            return
        invalid = [path for path in self.file_paths if not path.exists() or path.suffix.lower() != ".docx"]
        if invalid:
            messagebox.showerror("文件不支持", "以下文件不存在或不是 .docx：\n" + "\n".join(str(path) for path in invalid[:10]))
            return
        self.set_entry(self.output_dir_entry, str(output_dir))
        core.save_config(provider, api_key if self.remember_key.get() else "", base_url, model, english_font, chinese_font)

        self.is_running = True
        self.job_results = []
        self.job_errors = []
        self.job_cancel_events = {}
        self.active_jobs = len(self.file_paths)
        self.clear_job_rows()
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        for index, path in enumerate(self.file_paths, start=1):
            job_id = str(index)
            cancel_event = threading.Event()
            self.job_cancel_events[job_id] = cancel_event
            self.create_job_row(job_id, path)
            threading.Thread(
                target=self.run_job,
                args=(job_id, path, output_dir, provider, api_key, base_url, model, english_font, chinese_font, cancel_event),
                daemon=True,
            ).start()

    def run_job(
        self,
        job_id: str,
        path: Path,
        output_dir: Path,
        provider: str,
        api_key: str,
        base_url: str,
        model: str,
        english_font: str,
        chinese_font: str,
        cancel_event: threading.Event,
    ) -> None:
        try:
            result = core.translate_docx_hybrid(
                path,
                output_dir,
                provider,
                api_key,
                base_url,
                model,
                english_font,
                chinese_font,
                lambda msg: self.log(f"{path.name}: {msg}"),
                lambda done, total, current: self.update_translation(job_id, done, total, current),
                lambda done, total, current: self.update_review(job_id, done, total, current),
                cancel_event,
            )
        except Exception as exc:
            error = friendly_error(exc)
            self.root.after(0, self.finish_job_error, job_id, path, error)
            return
        self.root.after(0, self.finish_job_success, job_id, path, result, bool(result[-1]))

    def update_translation(self, job_id: str, done: int, total: int, current: str) -> None:
        self.root.after(0, self._update_translation_ui, job_id, done, total, current)

    def _update_translation_ui(self, job_id: str, done: int, total: int, current: str) -> None:
        widgets = self.job_widgets.get(job_id)
        if not widgets:
            return
        percent = 0 if total <= 0 else max(0, min(100, done * 100 / total))
        widgets["translation_bar"]["value"] = percent
        widgets["translation_text"].set(f"翻译进度：{done}/{total}（{percent:.1f}%）")
        widgets["current_text"].set(f"翻译：{current}")

    def update_review(self, job_id: str, done: int, total: int, current: str) -> None:
        self.root.after(0, self._update_review_ui, job_id, done, total, current)

    def _update_review_ui(self, job_id: str, done: int, total: int, current: str) -> None:
        widgets = self.job_widgets.get(job_id)
        if not widgets:
            return
        percent = 0 if total <= 0 else max(0, min(100, done * 100 / total))
        widgets["review_bar"]["value"] = percent
        widgets["review_text"].set(f"复核检查：{done}/{total}（{percent:.1f}%）")
        widgets["current_text"].set(f"复核：{current}")

    def request_cancel_job(self, job_id: str) -> None:
        event = self.job_cancel_events.get(job_id)
        if event:
            event.set()
            widgets = self.job_widgets.get(job_id)
            if widgets:
                widgets["current_text"].set("状态：正在中止并导出当前进度...")
                widgets["cancel"].configure(state="disabled")

    def request_cancel_all(self) -> None:
        for event in self.job_cancel_events.values():
            event.set()
        self.stop_button.configure(state="disabled")
        self.log("已请求全部中止；正在导出当前已完成进度。")

    def finish_job_success(self, job_id: str, path: Path, result: tuple, cancelled: bool) -> None:
        widgets = self.job_widgets.get(job_id)
        if widgets:
            widgets["cancel"].configure(state="disabled")
            widgets["current_text"].set("状态：已中止并导出" if cancelled else "状态：已完成")
        self.job_results.append((path, result, cancelled))
        self.finish_one_job()

    def finish_job_error(self, job_id: str, path: Path, error: str) -> None:
        widgets = self.job_widgets.get(job_id)
        if widgets:
            widgets["cancel"].configure(state="disabled")
            widgets["current_text"].set(f"状态：失败 - {error[:180]}")
        self.job_errors.append((path, error))
        self.finish_one_job()

    def finish_one_job(self) -> None:
        self.active_jobs = max(0, self.active_jobs - 1)
        if self.active_jobs:
            return
        self.is_running = False
        self.job_cancel_events = {}
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        completed = len([item for item in self.job_results if not item[2]])
        cancelled = len([item for item in self.job_results if item[2]])
        failed = len(self.job_errors)
        lines = [f"全部任务结束：完成 {completed} 个，中止导出 {cancelled} 个，失败 {failed} 个。"]
        for input_path, paths, was_cancelled in self.job_results[:10]:
            state = "中止导出" if was_cancelled else "完成"
            lines.append(f"{state}: {input_path.name}")
            lines.append(f"  {paths[0]}")
        for input_path, error in self.job_errors[:10]:
            lines.append(f"失败: {input_path.name} - {error}")
        if failed:
            messagebox.showwarning("部分任务失败", "\n".join(lines))
        else:
            messagebox.showinfo("全部任务结束", "\n".join(lines))

    def log(self, text: str) -> None:
        if threading.get_ident() != self.main_thread_id:
            self.root.after(0, self.log, text)
            return
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)


def main() -> None:
    root = tk.Tk()
    MacHybridApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise
