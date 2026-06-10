from __future__ import annotations

import ctypes
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, WORD, W, BooleanVar, StringVar, Text, Tk, filedialog, messagebox
from tkinter import ttk

from ccxd_us_reports_app.annual_runner import AnnualRunConfig, run_annual_extract
from ccxd_us_reports_app.mapping_template import create_annual_mapping_template, create_quarterly_mapping_template
from ccxd_us_reports_app.quarterly_runner import QuarterlyRunConfig, run_quarterly_extract


APP_TITLE = "美股年报季报提取工具"
FIELD_HELP_COLOR = "#5F6B7A"
CARD_BG = "#F6F8FB"


@dataclass
class FieldGroup:
    label: str
    var: StringVar
    help_text: str
    is_dir: bool = False
    required: bool = True
    test_only: bool = False


def _enable_high_dpi() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1440x1080")
        self.root.minsize(1280, 960)

        self.mode_var = StringVar(value="quarterly")
        self.status_var = StringVar(value="请选择功能并填写路径，然后点击“开始运行”。")
        self.enable_test_var = BooleanVar(value=False)
        self.template_btn_text = StringVar(value="生成季报映射模板")

        self.fields: dict[str, FieldGroup] = {}
        self._configure_styles()
        self._build()
        self._switch_mode()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except Exception:
            pass

        base_font = ("Microsoft YaHei UI", 14)
        style.configure(".", font=base_font)
        style.configure("Header.TLabel", font=("Microsoft YaHei UI", 24, "bold"))
        style.configure("SubHeader.TLabel", font=("Microsoft YaHei UI", 13))
        style.configure("Section.TLabelframe.Label", font=("Microsoft YaHei UI", 14, "bold"))
        style.configure("Hint.TLabel", font=("Microsoft YaHei UI", 11), foreground=FIELD_HELP_COLOR)
        style.configure("Status.TLabel", font=("Microsoft YaHei UI", 13))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("TCheckbutton", font=("Microsoft YaHei UI", 12))
        style.configure("TRadiobutton", font=("Microsoft YaHei UI", 13))
        style.configure("InfoCard.TFrame", background=CARD_BG)
        style.configure("InfoCardTitle.TLabel", font=("Microsoft YaHei UI", 13, "bold"), background=CARD_BG)
        style.configure("InfoCardText.TLabel", font=("Microsoft YaHei UI", 12), background=CARD_BG)

    def _build(self) -> None:
        frame = ttk.Frame(self.root, padding=22)
        frame.pack(fill=BOTH, expand=True)

        ttk.Label(frame, text="美股年报 / 季报提取工具", style="Header.TLabel").pack(anchor=W)
        ttk.Label(
            frame,
            text=(
                "面向业务用户的桌面工具。默认从 SEC note 数据中提取结果，不需要修改代码。"
                "年报和季报是两个独立功能，按界面提示选择数据源、股票池和输出目录即可运行。"
            ),
            style="SubHeader.TLabel",
            wraplength=1360,
        ).pack(anchor=W, pady=(8, 10))

        info_card = ttk.Frame(frame, style="InfoCard.TFrame", padding=16)
        info_card.pack(fill="x", pady=(0, 14))
        ttk.Label(info_card, text="输出说明", style="InfoCardTitle.TLabel").pack(anchor=W)
        ttk.Label(
            info_card,
            text=(
                "1. 一定输出源表。\n"
                "2. 一定输出未匹配公司表。\n"
                "3. 只有在提供映射表时，才输出修改映射表。\n"
                "4. 原件 / HTML 补齐默认关闭，仅在勾选“启用测试功能”后显示。"
            ),
            style="InfoCardText.TLabel",
            justify="left",
            wraplength=1320,
        ).pack(anchor=W, pady=(8, 0))

        mode_bar = ttk.Frame(frame)
        mode_bar.pack(fill="x", pady=(0, 12))
        ttk.Radiobutton(mode_bar, text="季报提取", value="quarterly", variable=self.mode_var, command=self._switch_mode).pack(side=LEFT)
        ttk.Radiobutton(mode_bar, text="年报提取", value="annual", variable=self.mode_var, command=self._switch_mode).pack(side=LEFT, padx=(18, 0))
        ttk.Checkbutton(
            mode_bar,
            text="启用测试功能（显示原件 / HTML 补齐项）",
            variable=self.enable_test_var,
            command=self._switch_mode,
        ).pack(side=LEFT, padx=(28, 0))

        self.form_frame = ttk.LabelFrame(frame, text="参数填写", style="Section.TLabelframe")
        self.form_frame.pack(fill="x")

        action_bar = ttk.Frame(frame)
        action_bar.pack(fill="x", pady=(14, 8))
        self.run_btn = ttk.Button(action_bar, text="开始运行", command=self._on_run, style="Primary.TButton")
        self.run_btn.pack(side=LEFT)
        ttk.Button(action_bar, textvariable=self.template_btn_text, command=self._save_template).pack(side=LEFT, padx=(10, 0))

        self.status_label = ttk.Label(frame, textvariable=self.status_var, style="Status.TLabel", wraplength=1360)
        self.status_label.pack(anchor=W, pady=(0, 8))

        log_wrap = ttk.LabelFrame(frame, text="运行日志", style="Section.TLabelframe")
        log_wrap.pack(fill=BOTH, expand=True)
        self.log = Text(log_wrap, height=16, wrap=WORD, font=("Consolas", 13), padx=14, pady=12)
        self.log.pack(fill=BOTH, expand=True)
        self.log.configure(state="disabled")

    def _clear_form(self) -> None:
        for child in self.form_frame.winfo_children():
            child.destroy()
        self.fields.clear()

    def _add_field(
        self,
        row: int,
        key: str,
        label: str,
        help_text: str,
        *,
        is_dir: bool = False,
        required: bool = True,
        test_only: bool = False,
    ) -> None:
        if test_only and not self.enable_test_var.get():
            return

        var = StringVar()
        self.fields[key] = FieldGroup(label, var, help_text, is_dir=is_dir, required=required, test_only=test_only)
        ttk.Label(self.form_frame, text=label).grid(row=row * 2, column=0, sticky=W, padx=(12, 8), pady=(8, 2))
        ttk.Entry(self.form_frame, textvariable=var, width=98).grid(row=row * 2, column=1, sticky="ew", padx=(0, 8), pady=(8, 2))
        ttk.Button(self.form_frame, text="选择", command=lambda k=key: self._pick_path(k)).grid(
            row=row * 2, column=2, sticky=W, padx=(0, 12), pady=(8, 2)
        )
        required_text = "必填" if required else "可选"
        if test_only:
            required_text = f"{required_text}，仅测试"
        ttk.Label(
            self.form_frame,
            text=f"{required_text}。{help_text}",
            style="Hint.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=row * 2 + 1, column=1, columnspan=2, sticky=W, padx=(0, 12), pady=(0, 6))
        self.form_frame.columnconfigure(1, weight=1)

    def _switch_mode(self) -> None:
        self._clear_form()
        quarterly = self.mode_var.get() == "quarterly"
        self.template_btn_text.set("生成季报映射模板" if quarterly else "生成年报映射模板")

        if quarterly:
            fields = [
                (
                    "notes_data_dir",
                    "季报 note 数据目录",
                    "季报主数据源。程序会优先从这里提取最新季报，再与 note 中该公司的最新年报比较 filed_date，选择更新的一份。",
                    True,
                    True,
                    False,
                ),
                (
                    "main_pool_workbook",
                    "主股票池工作簿",
                    "主股票池文件，决定本次要处理哪些公司。",
                    False,
                    True,
                    False,
                ),
                (
                    "extra_pool_workbook",
                    "补充股票池工作簿",
                    "补充公司清单，不需要时可以留空。",
                    False,
                    False,
                    False,
                ),
                (
                    "mapping_workbook",
                    "季报映射表（可选）",
                    "用于把 segment_axis 映射成 segment_type。只有填写后，才会额外输出“美股季报提取_修改映射.xlsx”。",
                    False,
                    False,
                    False,
                ),
                (
                    "output_dir",
                    "输出目录",
                    "会输出季报源表、季报未匹配公司表；如果提供映射表，还会输出季报修改映射表。",
                    True,
                    True,
                    False,
                ),
                (
                    "company_dir",
                    "本地年报原件目录（测试）",
                    "仅测试时使用。用于 note 无法覆盖时尝试从本地 HTML / 原件补齐。",
                    True,
                    False,
                    True,
                ),
            ]
        else:
            fields = [
                (
                    "notes_data_dir",
                    "年报 note 数据目录",
                    "年报主数据源。程序会直接从这里提取最新年报数据。",
                    True,
                    True,
                    False,
                ),
                (
                    "main_pool_workbook",
                    "主股票池工作簿",
                    "主股票池文件，决定本次要处理哪些公司。",
                    False,
                    True,
                    False,
                ),
                (
                    "extra_pool_workbook",
                    "补充股票池工作簿",
                    "补充公司清单，不需要时可以留空。",
                    False,
                    False,
                    False,
                ),
                (
                    "mapping_workbook",
                    "年报映射表（可选）",
                    "用于把 segment_axis 映射成 segment_type。只有填写后，才会额外输出“美股年报提取_修改映射.xlsx”。",
                    False,
                    False,
                    False,
                ),
                (
                    "output_dir",
                    "输出目录",
                    "会输出年报源表、年报未匹配公司表；如果提供映射表，还会输出年报修改映射表。",
                    True,
                    True,
                    False,
                ),
                (
                    "am_annual_dir",
                    "本地年报原件目录（测试）",
                    "仅测试时使用。用于 note 无法覆盖时尝试从本地原件补齐。",
                    True,
                    False,
                    True,
                ),
            ]

        render_row = 0
        for key, label, help_text, is_dir, required, test_only in fields:
            self._add_field(render_row, key, label, help_text, is_dir=is_dir, required=required, test_only=test_only)
            if not (test_only and not self.enable_test_var.get()):
                render_row += 1

    def _pick_path(self, key: str) -> None:
        field = self.fields[key]
        if field.is_dir:
            path = filedialog.askdirectory(title=f"选择{field.label}")
        else:
            path = filedialog.askopenfilename(title=f"选择{field.label}", filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")])
        if path:
            field.var.set(path)

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(END, f"{message}\n")
        self.log.see(END)
        self.log.configure(state="disabled")

    def _validate(self) -> dict[str, Path]:
        values: dict[str, Path] = {}
        for key, field in self.fields.items():
            raw = field.var.get().strip()
            if not raw:
                if field.required:
                    raise ValueError(f"请先填写：{field.label}")
                continue
            path = Path(raw)
            if not path.exists() and key != "output_dir":
                raise ValueError(f"路径不存在：{field.label}")
            values[key] = path
        return values

    def _save_template(self) -> None:
        quarterly = self.mode_var.get() == "quarterly"
        path = filedialog.asksaveasfilename(
            title="保存映射模板",
            defaultextension=".xlsx",
            initialfile="季报映射模板.xlsx" if quarterly else "年报映射模板.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path:
            return
        if quarterly:
            create_quarterly_mapping_template(Path(path))
        else:
            create_annual_mapping_template(Path(path))
        messagebox.showinfo(APP_TITLE, f"模板已生成：\n{path}")

    def _on_run(self) -> None:
        try:
            values = self._validate()
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        self.run_btn.config(state="disabled")
        self.status_var.set("正在运行，请稍候。")
        self._append_log(f"开始执行：{'季报提取' if self.mode_var.get() == 'quarterly' else '年报提取'}")
        threading.Thread(target=self._run_task, args=(values,), daemon=True).start()

    def _run_task(self, values: dict[str, Path]) -> None:
        try:
            if self.mode_var.get() == "quarterly":
                result = run_quarterly_extract(
                    QuarterlyRunConfig(
                        notes_data_dir=values["notes_data_dir"],
                        main_pool_workbook=values["main_pool_workbook"],
                        output_dir=values["output_dir"],
                        extra_pool_workbook=values.get("extra_pool_workbook"),
                        mapping_workbook=values.get("mapping_workbook"),
                        enable_test_features=self.enable_test_var.get(),
                        company_dir=values.get("company_dir"),
                    )
                )
                self.root.after(0, lambda: self._finish(result, "季报"))
            else:
                result = run_annual_extract(
                    AnnualRunConfig(
                        notes_data_dir=values["notes_data_dir"],
                        main_pool_workbook=values["main_pool_workbook"],
                        output_dir=values["output_dir"],
                        extra_pool_workbook=values.get("extra_pool_workbook"),
                        mapping_workbook=values.get("mapping_workbook"),
                        enable_test_features=self.enable_test_var.get(),
                        am_annual_dir=values.get("am_annual_dir"),
                    )
                )
                self.root.after(0, lambda: self._finish(result, "年报"))
        except Exception:
            error_text = traceback.format_exc()
            self.root.after(0, lambda: self._fail(error_text))

    def _finish(self, result: dict[str, object], mode_name: str) -> None:
        self._append_log(f"{mode_name}源表：{result['source_workbook']}")
        self._append_log(f"{mode_name}未匹配公司表：{result['unmatched_workbook']}")
        if result.get("mapped_written"):
            self._append_log(f"{mode_name}修改映射表：{result['mapped_workbook']}")
        else:
            self._append_log(f"未提供{mode_name}映射表，本次未输出修改映射表。")
        self.status_var.set(f"{mode_name}提取完成。")
        self.run_btn.config(state="normal")
        messagebox.showinfo(APP_TITLE, f"{mode_name}提取完成。")

    def _fail(self, error_text: str) -> None:
        self._append_log(error_text)
        self.status_var.set("运行失败，请查看日志。")
        self.run_btn.config(state="normal")
        messagebox.showerror(APP_TITLE, "运行失败，请查看下方日志。")


def launch() -> None:
    _enable_high_dpi()
    root = Tk()
    root.tk.call("tk", "scaling", 1.45)
    App(root)
    root.mainloop()
