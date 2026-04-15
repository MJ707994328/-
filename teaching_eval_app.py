from __future__ import annotations

import json
import sys
import threading
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox

import customtkinter as ctk
from PIL import Image

THIS_DIR = Path(__file__).resolve().parent
if str(THIS_DIR) not in sys.path:
    sys.path.insert(0, str(THIS_DIR))

from ai_scoring import AIScoringError, ScoreResult, is_ai_scoring_configured, save_score_report, score_experiment
from gds1000e import GDS1000ESerialClient, ScopeIdentity, autodetect_scope


ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

PREVIEW_INTERVAL_MS = 1200
SNAPSHOT_INTERVAL_MS = 5 * 60 * 1000
EXPERIMENT_ROOT = Path(__file__).with_name("experiments")
LOGO_PATH = THIS_DIR / "assets" / "xjtu_logo.png"

APP_BG = ("#F4F7FB", "#0E1621")
CARD_BG = ("#FFFFFF", "#16212B")
CARD_BG_SOFT = ("#F8FAFC", "#1B2733")
CARD_ELEVATED = ("#FFFFFF", "#1D2A36")
TEXT = ("#15202B", "#F5F7FA")
TEXT_MUTED = ("#667085", "#95A2B3")
TEXT_SOFT = ("#8A94A6", "#6F7F92")
ACCENT = ("#1F5EFF", "#6B8CFF")
ACCENT_HOVER = ("#1848C8", "#7B98FF")
ACCENT_SOFT = ("#EAF0FF", "#23324A")
SUCCESS_SOFT = ("#E7F6EF", "#183629")
SUCCESS_TEXT = ("#227A52", "#8EE0B6")
WARNING_SOFT = ("#FFF4DE", "#43331A")
WARNING_TEXT = ("#9A6800", "#F8C867")
DANGER = ("#CB3A31", "#E2564D")
DANGER_HOVER = ("#B72B23", "#EF6C63")
OUTLINE = ("#E5EAF1", "#263645")
PREVIEW_BG = ("#0E1621", "#0B121C")
PREVIEW_FRAME = ("#0F1724", "#0F1724")
PREVIEW_TEXT = ("#E8EEF6", "#E8EEF6")
XJTU_RED = (193, 39, 45)


@dataclass(slots=True)
class ExperimentSession:
    folder: Path
    snapshots_dir: Path
    started_at: datetime
    scope: ScopeIdentity
    expected_duration_seconds: int
    snapshot_count: int = 0

    @property
    def meta_path(self) -> Path:
        return self.folder / "meta.json"

    @property
    def description_path(self) -> Path:
        return self.folder / "description.txt"

    @property
    def final_image_path(self) -> Path:
        return self.folder / "final_screen.png"

    @property
    def score_report_path(self) -> Path:
        return self.folder / "ai_score.json"


class TeachingEvalApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("西安交通大学示波器实验测评系统")
        self.geometry("1580x980")
        self.minsize(1360, 860)
        self.configure(fg_color=APP_BG)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.scope_identity: ScopeIdentity | None = None
        self.session: ExperimentSession | None = None
        self.last_completed_session: ExperimentSession | None = None
        self.preview_job: str | None = None
        self.clock_job: str | None = None
        self.snapshot_job: str | None = None
        self.capture_thread: threading.Thread | None = None
        self.score_thread: threading.Thread | None = None
        self.preview_busy = False
        self.score_busy = False
        self.initial_snapshot_pending = False
        self.next_snapshot_at: datetime | None = None
        self.last_preview_image: Image.Image | None = None
        self.preview_ctk_image: ctk.CTkImage | None = None
        self.logo_image: ctk.CTkImage | None = None

        self.status_var = tk.StringVar(value="正在检测设备...")
        self.scope_var = tk.StringVar(value="未连接示波器")
        self.session_var = tk.StringVar(value="当前没有实验记录")
        self.expected_time_var = tk.StringVar(value="")
        self.expected_time_display_var = tk.StringVar(value="未设置")
        self.elapsed_var = tk.StringVar(value="00:00:00")
        self.snapshot_var = tk.StringVar(value="0")
        self.next_snapshot_var = tk.StringVar(value="尚未开始")
        self.preview_hint_var = tk.StringVar(value="点击“开始实验”后开始实时监控示波器画面。")
        self.ai_status_var = tk.StringVar(value="总分 = 波形得分 * 70% + 时间得分 * 30%")
        self.ai_score_var = tk.StringVar(value="--")
        self.ai_verdict_var = tk.StringVar(value="等待评分")
        self.ai_summary_var = tk.StringVar(value="结束实验后，系统将生成带明确扣分项的标准化评分结果。")
        self.ai_summary_compact_var = tk.StringVar(value="完成实验后，这里会显示简要评分摘要。")
        self.ai_objective_var = tk.StringVar(value="实验目标按教师给定目标处理，只用于判断结果是否达标，不对文字表达单独评分。")
        self.ai_feedback_var = tk.StringVar(value="评分完成后，这里会显示波形分析、优点总结和教师反馈。")
        self.ai_strengths_var = tk.StringVar(value="暂未生成波形扣分项。")
        self.ai_issues_var = tk.StringVar(value="暂未生成时间扣分项。")
        self.appearance_mode_var = tk.StringVar(value="跟随系统")

        self._configure_fonts()
        self._build_ui()
        self.after(80, self.refresh_scope)

    def _configure_fonts(self) -> None:
        self.ui_family = self._pick_font_family(
            ["gothic", "clean", "song ti", "fangsong ti", "DejaVu Sans"]
        )
        self.serif_family = self._pick_font_family(
            ["song ti", "fangsong ti", "mincho", "bitstream charter", "DejaVu Serif"]
        )
        self.mono_family = self._pick_font_family(
            ["courier 10 pitch", "fixed", "DejaVu Sans Mono"]
        )

        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=12, family=self.ui_family)
        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(size=12, family=self.ui_family)
        fixed_font = tkfont.nametofont("TkFixedFont")
        fixed_font.configure(size=11, family=self.mono_family)

        self.brand_font = ctk.CTkFont(family=self.serif_family, size=34, weight="bold")
        self.title_font = ctk.CTkFont(family=self.ui_family, size=21, weight="bold")
        self.subtitle_font = ctk.CTkFont(family=self.ui_family, size=15)
        self.metric_font = ctk.CTkFont(family=self.serif_family, size=38, weight="bold")
        self.metric_label_font = ctk.CTkFont(family=self.ui_family, size=14)
        self.body_font = ctk.CTkFont(family=self.ui_family, size=18)
        self.body_bold_font = ctk.CTkFont(family=self.ui_family, size=18, weight="bold")
        self.small_font = ctk.CTkFont(family=self.ui_family, size=14)
        self.input_font = ctk.CTkFont(family=self.ui_family, size=22)
        self.input_label_font = ctk.CTkFont(family=self.ui_family, size=21, weight="bold")
        self.score_font = ctk.CTkFont(family=self.serif_family, size=60, weight="bold")
        self.preview_hint_font = ctk.CTkFont(family=self.ui_family, size=21)

    def _pick_font_family(self, candidates: list[str]) -> str:
        available = set(tkfont.families(self))
        for family in candidates:
            if family in available:
                return family
        return tkfont.nametofont("TkDefaultFont").cget("family")

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=34, pady=(28, 20))
        header.grid_columnconfigure(1, weight=1)

        self._load_logo()
        if self.logo_image is not None:
            ctk.CTkLabel(header, image=self.logo_image, text="").grid(
                row=0, column=0, rowspan=2, sticky="w", padx=(0, 18)
            )

        brand = ctk.CTkFrame(header, fg_color="transparent")
        brand.grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(
            brand,
            text="西安交通大学示波器实验测评系统",
            font=self.brand_font,
            text_color=TEXT,
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            brand,
            text="更清晰的实时预览，更规范的过程记录，更可信的智能测评。",
            font=self.subtitle_font,
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=2, rowspan=2, sticky="e")

        self.mode_switch = ctk.CTkSegmentedButton(
            actions,
            values=["跟随系统", "浅色", "深色"],
            variable=self.appearance_mode_var,
            command=self._change_appearance_mode,
            font=self.small_font,
            height=40,
        )
        self.mode_switch.grid(row=0, column=0, padx=(0, 10))

        self.refresh_button = self._create_action_button(
            actions, "刷新设备", self.refresh_scope, style="secondary"
        )
        self.refresh_button.grid(row=0, column=1, padx=(0, 10))
        self.start_button = self._create_action_button(
            actions, "开始实验", self.start_experiment, style="primary"
        )
        self.start_button.grid(row=0, column=2, padx=(0, 10))
        self.end_button = self._create_action_button(
            actions, "结束实验", self.end_experiment, style="danger"
        )
        self.end_button.grid(row=0, column=3)
        self.start_button.configure(state="disabled")
        self.end_button.configure(state="disabled")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=34, pady=(0, 30))
        body.grid_columnconfigure(0, weight=7)
        body.grid_columnconfigure(1, weight=5)
        body.grid_rowconfigure(0, weight=8)
        body.grid_rowconfigure(1, weight=3)

        self.preview_card = self._create_card(body, elevated=True)
        self.preview_card.grid(row=0, column=0, sticky="nsew", padx=(0, 20), pady=(0, 18))
        self.preview_card.grid_columnconfigure(0, weight=1)
        self.preview_card.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            self.preview_card,
            text="示波器实时预览",
            font=self.title_font,
            text_color=TEXT,
        ).grid(row=0, column=0, sticky="w", padx=28, pady=(24, 8))
        ctk.CTkLabel(
            self.preview_card,
            text="实验开始后自动抓取首张有效画面，并持续同步示波器当前显示。",
            font=self.small_font,
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, sticky="w", padx=28, pady=(0, 16))

        self.preview_stage = ctk.CTkFrame(
            self.preview_card,
            corner_radius=28,
            fg_color=PREVIEW_FRAME,
            border_width=1,
            border_color=("#1E2A38", "#233342"),
        )
        self.preview_stage.grid(row=2, column=0, sticky="nsew", padx=28, pady=(0, 28))
        self.preview_stage.bind("<Configure>", self._on_preview_stage_resize)

        self.preview_label = ctk.CTkLabel(
            self.preview_stage,
            textvariable=self.preview_hint_var,
            font=self.preview_hint_font,
            text_color=PREVIEW_TEXT,
            justify="center",
        )
        self.preview_label.pack(expand=True, fill="both", padx=16, pady=16)

        sidebar = ctk.CTkFrame(body, fg_color="transparent")
        sidebar.grid(row=0, column=1, sticky="nsew")
        sidebar.grid_columnconfigure(0, weight=1)
        sidebar.grid_rowconfigure(2, weight=1)

        metrics = ctk.CTkFrame(sidebar, fg_color="transparent")
        metrics.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        for col in range(3):
            metrics.grid_columnconfigure(col, weight=1)

        self.status_card = self._create_metric_card(metrics, "当前状态", self.status_var, metric=False)
        self.status_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.duration_card = self._create_metric_card(metrics, "实验用时", self.elapsed_var, metric=True)
        self.duration_card.grid(row=0, column=1, sticky="nsew", padx=5)
        self.snapshot_card = self._create_metric_card(metrics, "过程截图", self.snapshot_var, metric=True)
        self.snapshot_card.grid(row=0, column=2, sticky="nsew", padx=(10, 0))

        self.status_pill = ctk.CTkLabel(
            self.status_card,
            textvariable=self.status_var,
            font=self.body_bold_font,
            text_color=ACCENT,
            fg_color=ACCENT_SOFT,
            corner_radius=999,
            height=40,
            anchor="center",
        )
        self.status_pill.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 18))

        details_card = self._create_card(sidebar)
        details_card.grid(row=1, column=0, sticky="ew", pady=(0, 18))
        details_card.grid_columnconfigure(0, weight=1)
        details_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(details_card, text="实验信息", font=self.title_font, text_color=TEXT).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=22, pady=(20, 14)
        )
        self._add_info_row(details_card, 1, "当前设备", self.scope_var, column=0)
        self._add_info_row(details_card, 1, "实验目录", self.session_var, column=1)
        self._add_info_row(details_card, 2, "预期时长", self.expected_time_display_var, column=0)
        self._add_info_row(details_card, 2, "下次自动截图", self.next_snapshot_var, column=1)

        objective_card = self._create_card(sidebar)
        objective_card.grid(row=2, column=0, sticky="nsew")
        objective_card.grid_columnconfigure(0, weight=1)
        objective_card.grid_rowconfigure(3, weight=1, minsize=380)
        objective_card.bind("<Configure>", self._on_objective_card_resize)
        ctk.CTkLabel(objective_card, text="实验目标", font=self.title_font, text_color=TEXT).grid(
            row=0, column=0, sticky="w", padx=22, pady=(22, 12)
        )

        expected_row = ctk.CTkFrame(objective_card, fg_color="transparent")
        expected_row.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 12))
        expected_row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            expected_row,
            text="预期时长（分钟）",
            font=self.input_label_font,
            text_color=TEXT,
        ).grid(row=0, column=0, sticky="w")
        self.expected_time_entry = ctk.CTkEntry(
            expected_row,
            width=380,
            textvariable=self.expected_time_var,
            font=self.input_font,
            justify="left",
            corner_radius=24,
            height=78,
            border_color=OUTLINE,
            fg_color=CARD_BG_SOFT,
        )
        self.expected_time_entry.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.expected_time_entry._entry.configure(insertwidth=3, relief="flat")

        self.objective_hint_label = ctk.CTkLabel(
            objective_card,
            text="这里填写老师给出的实验目标。系统不会按作文来评分，只会用它判断最终波形是否符合目标。",
            font=self.small_font,
            text_color=TEXT_MUTED,
            justify="left",
            wraplength=460,
        )
        self.objective_hint_label.grid(row=2, column=0, sticky="w", padx=22, pady=(0, 8))

        self.description_text = ctk.CTkTextbox(
            objective_card,
            corner_radius=24,
            border_width=1,
            border_spacing=12,
            border_color=OUTLINE,
            fg_color=CARD_BG_SOFT,
            text_color=TEXT,
            font=self.input_font,
            height=380,
            wrap="word",
        )
        self.description_text.grid(row=3, column=0, sticky="nsew", padx=22, pady=(0, 10))
        self.description_text._textbox.configure(padx=20, pady=18, spacing1=8, spacing3=10)

        self.time_rule_hint_label = ctk.CTkLabel(
            objective_card,
            text="时间评分规则：若实际用时比预期用时超出 X%，则时间部分扣 X 分。",
            font=self.small_font,
            text_color=TEXT_SOFT,
            justify="left",
            wraplength=460,
        )
        self.time_rule_hint_label.grid(row=4, column=0, sticky="w", padx=22, pady=(0, 16))

        score_card = self._create_card(body, elevated=True)
        score_card.grid(row=1, column=0, columnspan=2, sticky="nsew")
        score_card.grid_columnconfigure(0, weight=1)
        score_card.grid_rowconfigure(3, weight=1)

        score_header = ctk.CTkFrame(score_card, fg_color="transparent")
        score_header.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
        score_header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(score_header, text="智能评分结果", font=self.title_font, text_color=TEXT).grid(
            row=0, column=0, sticky="w"
        )

        score_header_summary = ctk.CTkFrame(score_header, fg_color="transparent")
        score_header_summary.grid(row=0, column=1, sticky="ew", padx=(18, 18))
        score_header_summary.grid_columnconfigure(1, weight=1)

        self.score_header_tile = ctk.CTkFrame(
            score_header_summary,
            width=96,
            corner_radius=20,
            fg_color=ACCENT_SOFT,
            border_width=1,
            border_color=("#D7E3FF", "#2B4166"),
        )
        self.score_header_tile.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.score_header_tile.grid_propagate(False)
        ctk.CTkLabel(self.score_header_tile, textvariable=self.ai_score_var, font=self.metric_font, text_color=ACCENT).pack(
            anchor="center", pady=(8, 0)
        )
        ctk.CTkLabel(self.score_header_tile, text="综合得分", font=self.small_font, text_color=TEXT_MUTED).pack(
            anchor="center", pady=(0, 10)
        )

        score_header_text = ctk.CTkFrame(score_header_summary, fg_color="transparent")
        score_header_text.grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(
            score_header_text,
            textvariable=self.ai_verdict_var,
            font=self.body_bold_font,
            text_color=TEXT,
            anchor="w",
            justify="left",
        ).grid(row=0, column=0, sticky="ew")
        self.ai_summary_compact_label = ctk.CTkLabel(
            score_header_text,
            textvariable=self.ai_summary_compact_var,
            font=self.small_font,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=520,
        )
        self.ai_summary_compact_label.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        self.rescore_button = self._create_action_button(score_header, "重新评分", self.run_ai_scoring, style="primary")
        self.rescore_button.grid(row=0, column=2, sticky="e")
        score_card.bind("<Configure>", self._on_score_summary_resize)

        self.ai_status_label = ctk.CTkLabel(
            score_card,
            textvariable=self.ai_status_var,
            font=self.small_font,
            text_color=TEXT_MUTED,
            justify="left",
            anchor="w",
        )
        self.ai_status_label.grid(row=1, column=0, sticky="ew", padx=24)

        self.ai_summary_label = ctk.CTkLabel(
            score_card,
            textvariable=self.ai_summary_var,
            font=self.body_font,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
        )
        self.ai_summary_label.grid(row=2, column=0, sticky="ew", padx=24, pady=(10, 10))

        content = ctk.CTkFrame(score_card, fg_color="transparent")
        content.grid(row=3, column=0, sticky="nsew", padx=24, pady=(0, 18))
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1, uniform="score_rows")
        content.grid_rowconfigure(1, weight=1, uniform="score_rows")

        self.waveform_box = self._create_text_card(content, "波形扣分项")
        self.waveform_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 12))
        self.time_box = self._create_text_card(content, "时间扣分项")
        self.time_box.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=(0, 12))
        self.objective_box = self._create_text_card(content, "目标说明与结构化读屏")
        self.objective_box.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        self.feedback_box = self._create_text_card(content, "教师反馈与优点")
        self.feedback_box.grid(row=1, column=1, sticky="nsew", padx=(10, 0))

        self._reset_ai_panel()

    def _create_action_button(self, parent, text: str, command, *, style: str) -> ctk.CTkButton:
        palette = {
            "primary": (ACCENT, ACCENT_HOVER, "#FFFFFF"),
            "secondary": (CARD_BG_SOFT, ("#EEF2F7", "#223042"), TEXT),
            "danger": (DANGER, DANGER_HOVER, "#FFFFFF"),
        }
        fg_color, hover_color, text_color = palette[style]
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            corner_radius=18,
            height=48,
            font=self.body_bold_font,
            fg_color=fg_color,
            hover_color=hover_color,
            text_color=text_color,
            border_width=0 if style != "secondary" else 1,
            border_color=OUTLINE,
        )

    def _create_card(self, parent, *, elevated: bool = False) -> ctk.CTkFrame:
        return ctk.CTkFrame(
            parent,
            corner_radius=30,
            fg_color=CARD_ELEVATED if elevated else CARD_BG,
            border_width=1,
            border_color=OUTLINE,
        )

    def _create_metric_card(self, parent, title: str, variable: tk.StringVar, *, metric: bool) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(
            parent,
            corner_radius=22,
            fg_color=CARD_BG_SOFT,
            border_width=1,
            border_color=OUTLINE,
        )
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=title, font=self.metric_label_font, text_color=TEXT_MUTED).grid(
            row=0, column=0, sticky="w", padx=20, pady=(18, 8)
        )
        if metric:
            ctk.CTkLabel(frame, textvariable=variable, font=self.metric_font, text_color=TEXT).grid(
                row=1, column=0, sticky="w", padx=20, pady=(0, 20)
            )
        return frame

    def _create_text_card(self, parent, title: str) -> ctk.CTkTextbox:
        wrapper = ctk.CTkFrame(
            parent,
            corner_radius=24,
            fg_color=CARD_BG_SOFT,
            border_width=1,
            border_color=OUTLINE,
        )
        wrapper.grid_columnconfigure(0, weight=1)
        wrapper.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(wrapper, text=title, font=self.body_bold_font, text_color=TEXT).grid(
            row=0, column=0, sticky="w", padx=20, pady=(18, 12)
        )
        textbox = ctk.CTkTextbox(
            wrapper,
            corner_radius=18,
            border_width=0,
            fg_color="transparent",
            text_color=TEXT,
            font=self.body_font,
            wrap="word",
        )
        textbox.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        textbox._textbox.configure(padx=6, pady=6, spacing1=4, spacing3=6)
        textbox.configure(state="disabled")
        wrapper.textbox = textbox
        return wrapper

    def _textbox(self, wrapper: ctk.CTkFrame) -> ctk.CTkTextbox:
        return wrapper.textbox

    def _set_textbox(self, wrapper: ctk.CTkFrame, text: str) -> None:
        textbox = self._textbox(wrapper)
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.insert("1.0", text)
        textbox.configure(state="disabled")

    def _load_logo(self) -> None:
        if not LOGO_PATH.exists():
            return
        image = Image.open(LOGO_PATH).convert("RGBA")
        red_logo = Image.new("RGBA", image.size, (XJTU_RED[0], XJTU_RED[1], XJTU_RED[2], 255))
        red_logo.putalpha(image.getchannel("A"))
        red_logo = red_logo.resize((82, 82), Image.Resampling.LANCZOS)
        self.logo_image = ctk.CTkImage(light_image=red_logo, dark_image=red_logo, size=(82, 82))

    def _add_info_row(self, parent: ctk.CTkFrame, row_index: int, title: str, variable: tk.StringVar, *, column: int = 0) -> None:
        label_row = row_index * 2 - 1
        value_row = row_index * 2
        pad_left = 22 if column == 0 else 14
        pad_right = 14 if column == 0 else 22
        ctk.CTkLabel(parent, text=title, font=self.metric_label_font, text_color=TEXT_MUTED).grid(
            row=label_row, column=column, sticky="w", padx=(pad_left, pad_right), pady=(0, 6)
        )
        ctk.CTkLabel(
            parent,
            textvariable=variable,
            font=self.body_font,
            text_color=TEXT,
            justify="left",
            anchor="w",
            wraplength=250 if column else 260,
        ).grid(row=value_row, column=column, sticky="w", padx=(pad_left, pad_right), pady=(0, 16))

    def _change_appearance_mode(self, value: str) -> None:
        mapping = {"跟随系统": "system", "浅色": "light", "深色": "dark"}
        ctk.set_appearance_mode(mapping.get(value, "system"))

    def refresh_scope(self) -> None:
        self.status_var.set("正在检测设备...")
        self._set_status_style(ACCENT_SOFT, ACCENT)
        self.update_idletasks()
        try:
            identity = autodetect_scope()
        except Exception as exc:
            self.scope_identity = None
            self.scope_var.set("未检测到示波器")
            self.status_var.set(f"设备检测失败：{exc}")
            self._set_status_style((DANGER, DANGER), ("#FFFFFF", "#FFFFFF"))
            self.start_button.configure(state="disabled")
            return

        self.scope_identity = identity
        self.scope_var.set(f"{identity.raw} @ {identity.port}")
        self.status_var.set("设备已就绪")
        self._set_status_style(SUCCESS_SOFT, SUCCESS_TEXT)
        if self.session is None:
            self.start_button.configure(state="normal")

    def start_experiment(self) -> None:
        if self.scope_identity is None:
            messagebox.showwarning("未连接设备", "请先连接示波器，然后点击“刷新设备”。")
            return
        if self.session is not None:
            messagebox.showinfo("实验进行中", "当前已经有一个实验正在进行。")
            return

        objective_text = self.description_text.get("1.0", "end").strip()
        if not objective_text:
            messagebox.showwarning("实验目标缺失", "开始实验前，请先填写老师给出的实验目标。")
            return
        expected_duration_seconds = self._parse_expected_duration_seconds()
        if expected_duration_seconds is None:
            return

        started_at = datetime.now()
        folder = EXPERIMENT_ROOT / started_at.strftime("%Y%m%d_%H%M%S")
        snapshots_dir = folder / "snapshots"
        snapshots_dir.mkdir(parents=True, exist_ok=True)

        self.session = ExperimentSession(
            folder=folder,
            snapshots_dir=snapshots_dir,
            started_at=started_at,
            scope=self.scope_identity,
            expected_duration_seconds=expected_duration_seconds,
        )
        self.last_completed_session = None
        self.initial_snapshot_pending = True
        self.last_preview_image = None
        self.snapshot_var.set("0")
        self.expected_time_display_var.set(self._format_duration(expected_duration_seconds))
        self.next_snapshot_var.set("等待首张有效截图")
        self.preview_hint_var.set("正在抓取第一张有效示波器画面...")
        self._reset_ai_panel()
        self._write_session_metadata(status="running")

        self.start_button.configure(state="disabled")
        self.end_button.configure(state="normal")
        self.refresh_button.configure(state="disabled")
        self.session_var.set(str(folder))
        self.status_var.set("实验进行中")
        self._set_status_style(ACCENT_SOFT, ACCENT)

        self._update_elapsed_clock()
        self._schedule_preview(immediate=True)

    def end_experiment(self) -> None:
        if self.session is None:
            return

        active_session = self.session
        if self.preview_job is not None:
            self.after_cancel(self.preview_job)
            self.preview_job = None
        if self.clock_job is not None:
            self.after_cancel(self.clock_job)
            self.clock_job = None
        if self.snapshot_job is not None:
            self.after_cancel(self.snapshot_job)
            self.snapshot_job = None

        final_image = self.last_preview_image.copy() if self.last_preview_image is not None else self._capture_scope_image()
        if final_image is not None:
            self.last_preview_image = final_image
            self._render_preview()
            final_image.save(active_session.final_image_path)

        ended_at = datetime.now()
        self.next_snapshot_at = None
        self.next_snapshot_var.set("已完成")
        self._write_session_metadata(status="completed", ended_at=ended_at)

        self.session = None
        self.last_completed_session = active_session
        self.status_var.set("实验已完成")
        self._set_status_style(WARNING_SOFT, WARNING_TEXT)
        self.session_var.set(str(active_session.folder))
        self.elapsed_var.set("00:00:00")
        self.initial_snapshot_pending = False
        self.start_button.configure(state="normal" if self.scope_identity else "disabled")
        self.end_button.configure(state="disabled")
        self.refresh_button.configure(state="normal")

        if is_ai_scoring_configured():
            self._start_ai_scoring(active_session)
        else:
            self.ai_status_var.set("智能评分功能已就绪，但当前还没有配置模型 API Key。")
            self.ai_summary_var.set("请先配置 MOONSHOT_API_KEY、KIMI_API_KEY 或 OPENAI_API_KEY 后再进行智能评分。")
            self.ai_objective_var.set("实验目标按教师给定目标处理，不会作为学生自评或作文来评分。")
            self.ai_feedback_var.set("完成 API 配置后，每次实验结束都可以自动生成详细评分。")
            self._sync_score_textboxes()

        messagebox.showinfo(
            "实验已保存",
            f"本次实验记录已保存到：\n{active_session.folder}\n\n评分结果可在下方评分面板查看。",
        )

    def run_ai_scoring(self) -> None:
        if self.session is not None:
            messagebox.showinfo("请先结束实验", "请先结束当前实验，再执行智能评分。")
            return
        if self.last_completed_session is None:
            messagebox.showinfo("暂无实验记录", "请先完成一次实验，再进行智能评分。")
            return
        if not is_ai_scoring_configured():
            messagebox.showwarning("缺少 API Key", "请先配置 MOONSHOT_API_KEY、KIMI_API_KEY 或 OPENAI_API_KEY。")
            return
        self._start_ai_scoring(self.last_completed_session)

    def _start_ai_scoring(self, session: ExperimentSession) -> None:
        if self.score_busy:
            return

        description = session.description_path.read_text(encoding="utf-8") if session.description_path.exists() else ""
        duration_seconds = 0
        expected_duration_seconds = session.expected_duration_seconds
        if session.meta_path.exists():
            try:
                meta = json.loads(session.meta_path.read_text(encoding="utf-8"))
                duration_seconds = int(meta.get("duration_seconds", 0))
                expected_duration_seconds = int(meta.get("expected_duration_seconds", expected_duration_seconds))
            except Exception:
                duration_seconds = 0

        self.score_busy = True
        self.rescore_button.configure(state="disabled")
        self.ai_status_var.set("正在生成智能评分...")
        self.ai_score_var.set("...")
        self.ai_verdict_var.set("正在分析")
        self.ai_summary_var.set("正在结合实验目标、最终截图和实际用时生成评分。")
        self.ai_summary_compact_var.set("正在根据实验目标、波形截图和实验用时生成评分摘要。")
        self.ai_feedback_var.set("正在整理标准化评分条目，请稍候。")
        self.ai_strengths_var.set("正在分析波形扣分项...")
        self.ai_issues_var.set("正在计算时间扣分项...")
        self.ai_objective_var.set("正在提取示波器屏幕中的结构化事实...")
        self._sync_score_textboxes()

        self.score_thread = threading.Thread(
            target=self._score_worker,
            args=(session, description, duration_seconds, expected_duration_seconds),
            daemon=True,
        )
        self.score_thread.start()

    def _score_worker(
        self,
        session: ExperimentSession,
        description: str,
        duration_seconds: int,
        expected_duration_seconds: int,
    ) -> None:
        try:
            result = score_experiment(
                description=description,
                duration_seconds=duration_seconds,
                expected_duration_seconds=expected_duration_seconds,
                final_image_path=session.final_image_path,
            )
            report_path = save_score_report(result, session.score_report_path)
        except Exception as exc:
            self.after(0, self._on_score_error, session, exc)
            return
        self.after(0, self._on_score_ready, session, result, report_path)

    def _on_score_ready(self, session: ExperimentSession, result: ScoreResult, report_path: Path) -> None:
        self.score_busy = False
        self.rescore_button.configure(state="normal")
        self.ai_status_var.set(f"{result.scoring_formula}；时间规则：超时 X%，时间部分扣 X 分。")
        self.ai_score_var.set(str(result.overall_score))
        self.ai_verdict_var.set(f"{result.verdict} / 波形 {result.waveform_score}/100 / 时间 {result.time_score}/100")
        self.ai_summary_var.set(result.summary)
        self.ai_summary_compact_var.set(
            f"综合 {result.overall_score}/100，波形 {result.waveform_score}/100，时间 {result.time_score}/100。"
        )
        self.ai_objective_var.set(
            f"{result.objective_notice}\n\n"
            f"{self._format_screen_facts_summary(result)}\n\n"
            f"预期时长：{self._format_duration(result.expected_duration_seconds)}\n"
            f"实际时长：{self._format_duration(result.actual_duration_seconds)}\n"
            f"时间规则：{result.time_rule}"
        )
        self.ai_feedback_var.set(
            f"屏幕判读：\n{self._format_bullets(result.screen_observations)}\n\n"
            f"幅值判断依据：\n{result.amplitude_evidence}\n\n"
            f"波形分析：\n{result.waveform_summary}\n\n"
            f"优点：\n{self._format_bullets(result.strengths)}\n\n"
            f"教师反馈：\n{result.instructor_feedback}"
        )
        self.ai_strengths_var.set(self._format_deductions(result.waveform_deductions, empty_text="没有波形扣分项，波形部分保持满分。"))
        self.ai_issues_var.set(self._format_deductions(result.time_deductions, empty_text="没有时间扣分项。"))
        self._sync_score_textboxes()
        self._append_score_metadata(session, result, report_path)

    def _on_score_error(self, session: ExperimentSession, exc: Exception) -> None:
        self.score_busy = False
        self.rescore_button.configure(state="normal")
        message = str(exc) if isinstance(exc, AIScoringError) else f"{type(exc).__name__}: {exc}"
        self.ai_status_var.set("智能评分失败")
        self.ai_score_var.set("--")
        self.ai_verdict_var.set("暂无结果")
        self.ai_summary_var.set(message)
        self.ai_summary_compact_var.set("评分暂时未完成，请检查网络或稍后重试。")
        self.ai_objective_var.set("实验目标按教师给定目标处理，不会作为学生自评或作文来评分。")
        self.ai_feedback_var.set("智能评分暂时未完成，请检查网络、模型配置或稍后重试。")
        self.ai_strengths_var.set("由于评分未完成，暂时没有波形扣分明细。")
        self.ai_issues_var.set("请检查 API Key、网络连通性、模型配置以及评分规则。")
        self._sync_score_textboxes()
        self._append_score_metadata(session, None, None, error_message=message)

    def _sync_score_textboxes(self) -> None:
        self._set_textbox(self.waveform_box, self.ai_strengths_var.get())
        self._set_textbox(self.time_box, self.ai_issues_var.get())
        self._set_textbox(self.objective_box, self.ai_objective_var.get())
        self._set_textbox(self.feedback_box, self.ai_feedback_var.get())
        self.ai_summary_compact_label.configure(wraplength=360)

    def on_close(self) -> None:
        if self.session is not None:
            confirmed = messagebox.askyesno(
                "保存并退出",
                "当前仍有实验在进行。\n是否先保存当前实验记录再退出？",
            )
            if not confirmed:
                return
            self.end_experiment()
        self.destroy()

    def _schedule_preview(self, immediate: bool = False) -> None:
        delay = 60 if immediate else PREVIEW_INTERVAL_MS
        self.preview_job = self.after(delay, self._start_preview_capture)

    def _start_preview_capture(self) -> None:
        self.preview_job = None
        if self.session is None or self.preview_busy:
            if self.session is not None:
                self._schedule_preview()
            return

        self.preview_busy = True
        self.capture_thread = threading.Thread(target=self._preview_worker, daemon=True)
        self.capture_thread.start()

    def _preview_worker(self) -> None:
        try:
            image = self._capture_scope_image()
        except Exception as exc:
            self.after(0, self._on_preview_error, exc)
            return
        self.after(0, self._on_preview_ready, image)

    def _on_preview_ready(self, image: Image.Image | None) -> None:
        self.preview_busy = False
        if image is not None:
            self.last_preview_image = image
            self._render_preview()
            if self.session is not None:
                self.preview_hint_var.set("示波器画面已同步到预览区。")
                if self.initial_snapshot_pending:
                    self.initial_snapshot_pending = False
                    self._save_snapshot_from_image(image, reason="start")
                    self._schedule_periodic_snapshot()
        else:
            self.preview_hint_var.set("预览失败，请检查示波器连接状态。")
            self.status_var.set("预览失败")
            self._set_status_style((DANGER, DANGER), ("#FFFFFF", "#FFFFFF"))

        if self.session is not None:
            self._schedule_preview()

    def _on_preview_error(self, exc: Exception) -> None:
        self.preview_busy = False
        self.status_var.set(f"预览失败：{exc}")
        self._set_status_style((DANGER, DANGER), ("#FFFFFF", "#FFFFFF"))
        self.preview_hint_var.set("预览失败，请检查示波器连接状态。")
        if self.session is not None:
            self._schedule_preview()

    def _schedule_periodic_snapshot(self) -> None:
        if self.session is None:
            return
        if self.snapshot_job is not None:
            self.after_cancel(self.snapshot_job)
        self.next_snapshot_at = datetime.now() + timedelta(milliseconds=SNAPSHOT_INTERVAL_MS)
        self.snapshot_job = self.after(SNAPSHOT_INTERVAL_MS, self._take_periodic_snapshot)
        self._refresh_next_snapshot_label()

    def _take_periodic_snapshot(self) -> None:
        self.snapshot_job = None
        if self.session is None:
            return
        if self.last_preview_image is not None:
            self._save_snapshot_from_image(self.last_preview_image, reason="auto")
        self._schedule_periodic_snapshot()

    def _save_snapshot_from_image(self, image: Image.Image, reason: str) -> None:
        if self.session is None:
            return
        self.session.snapshot_count += 1
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{self.session.snapshot_count:03d}_{reason}_{timestamp}.png"
        output_path = self.session.snapshots_dir / filename
        image.save(output_path)
        self.snapshot_var.set(str(self.session.snapshot_count))
        self._write_session_metadata(status="running")

    def _capture_scope_image(self) -> Image.Image | None:
        identity = self.scope_identity
        if identity is None:
            return None
        with GDS1000ESerialClient(identity.port) as scope:
            return scope.capture_display_image()

    def _on_preview_stage_resize(self, _event: tk.Event) -> None:
        self._render_preview()

    def _on_objective_card_resize(self, event: tk.Event) -> None:
        if not hasattr(self, "objective_hint_label") or not hasattr(self, "time_rule_hint_label"):
            return
        wrap = max(event.width - 56, 260)
        self.objective_hint_label.configure(wraplength=wrap)
        self.time_rule_hint_label.configure(wraplength=wrap)

    def _on_score_summary_resize(self, event: tk.Event) -> None:
        if not hasattr(self, "ai_summary_compact_label"):
            return
        wrap = max(event.width - 320, 260)
        self.ai_summary_compact_label.configure(wraplength=wrap)

    def _render_preview(self) -> None:
        if self.last_preview_image is None:
            return

        max_width = max(self.preview_stage.winfo_width() - 48, 520)
        max_height = max(self.preview_stage.winfo_height() - 48, 320)
        width, height = self.last_preview_image.size
        scale = min(max_width / width, max_height / height)
        scale = max(0.55, min(scale, 2.5))
        new_size = (max(1, int(width * scale)), max(1, int(height * scale)))
        preview = self.last_preview_image.resize(new_size, Image.Resampling.LANCZOS)
        self.preview_ctk_image = ctk.CTkImage(light_image=preview, dark_image=preview, size=new_size)
        self.preview_label.configure(image=self.preview_ctk_image, text="")

    def _update_elapsed_clock(self) -> None:
        if self.session is None:
            self.clock_job = None
            return

        elapsed = datetime.now() - self.session.started_at
        total_seconds = int(elapsed.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        self.elapsed_var.set(f"{hours:02d}:{minutes:02d}:{seconds:02d}")
        self._refresh_next_snapshot_label()
        self.clock_job = self.after(1000, self._update_elapsed_clock)

    def _refresh_next_snapshot_label(self) -> None:
        if self.session is None:
            self.next_snapshot_var.set("未开始")
            return
        if self.next_snapshot_at is None:
            self.next_snapshot_var.set("等待首张有效截图")
            return
        remaining = int((self.next_snapshot_at - datetime.now()).total_seconds())
        remaining = max(0, remaining)
        minutes, seconds = divmod(remaining, 60)
        self.next_snapshot_var.set(f"{minutes:02d}:{seconds:02d} 后自动保存")

    def _parse_expected_duration_seconds(self) -> int | None:
        raw = self.expected_time_var.get().strip()
        if not raw:
            messagebox.showwarning("缺少预期时长", "开始实验前，请先填写预期时长（分钟）。")
            return None
        try:
            minutes = float(raw)
        except ValueError:
            messagebox.showwarning("预期时长无效", "预期时长必须是数字，例如 8 或 12.5。")
            return None
        if minutes <= 0:
            messagebox.showwarning("预期时长无效", "预期时长必须大于 0 分钟。")
            return None
        return max(1, round(minutes * 60))

    def _format_duration(self, seconds: int) -> str:
        minutes = seconds / 60.0
        if abs(minutes - round(minutes)) < 1e-9:
            return f"{int(round(minutes))} 分钟"
        return f"{minutes:.1f} 分钟"

    def _write_session_metadata(self, status: str, ended_at: datetime | None = None) -> None:
        if self.session is None:
            return

        description = self.description_text.get("1.0", "end").strip()
        self.session.description_path.write_text(description, encoding="utf-8")

        payload = {
            "status": status,
            "scope": {
                "port": self.session.scope.port,
                "manufacturer": self.session.scope.manufacturer,
                "model": self.session.scope.model,
                "serial_number": self.session.scope.serial_number,
                "firmware": self.session.scope.firmware,
                "raw": self.session.scope.raw,
            },
            "started_at": self.session.started_at.isoformat(timespec="seconds"),
            "description": description,
            "experiment_objective": description,
            "expected_duration_seconds": self.session.expected_duration_seconds,
            "expected_duration_minutes": round(self.session.expected_duration_seconds / 60.0, 2),
            "snapshots_dir": str(self.session.snapshots_dir),
            "snapshot_count": self.session.snapshot_count,
            "snapshot_interval_seconds": SNAPSHOT_INTERVAL_MS // 1000,
        }
        if ended_at is not None:
            duration_seconds = int((ended_at - self.session.started_at).total_seconds())
            payload["ended_at"] = ended_at.isoformat(timespec="seconds")
            payload["duration_seconds"] = duration_seconds
            payload["final_image"] = str(self.session.final_image_path)

        self.session.meta_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _append_score_metadata(
        self,
        session: ExperimentSession,
        result: ScoreResult | None,
        report_path: Path | None,
        *,
        error_message: str | None = None,
    ) -> None:
        if not session.meta_path.exists():
            return
        try:
            payload = json.loads(session.meta_path.read_text(encoding="utf-8"))
        except Exception:
            return

        if report_path is not None:
            payload["ai_score_report"] = str(report_path)
        if result is not None:
            payload["ai_score"] = result.to_dict()
        if error_message is not None:
            payload["ai_score_error"] = error_message

        session.meta_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _reset_ai_panel(self) -> None:
        self.ai_status_var.set("总分 = 波形得分 * 70% + 时间得分 * 30%")
        self.ai_score_var.set("--")
        self.ai_verdict_var.set("等待评分")
        self.ai_summary_var.set("结束实验后，系统将生成带明确扣分项的标准化评分结果。")
        self.ai_summary_compact_var.set("完成实验后，这里会显示简要评分摘要。")
        self.ai_objective_var.set("实验目标按教师给定目标处理，不会作为学生自评或作文来评分。")
        self.ai_feedback_var.set("评分完成后，这里会显示波形分析、优点总结和教师反馈。")
        self.ai_strengths_var.set("暂未生成波形扣分项。")
        self.ai_issues_var.set("暂未生成时间扣分项。")
        self._sync_score_textboxes()
        self.rescore_button.configure(state="normal")

    def _format_bullets(self, items: list[str]) -> str:
        return "\n".join(f"- {item}" for item in items)

    def _format_deductions(self, items, *, empty_text: str) -> str:
        if not items:
            return empty_text
        return "\n".join(f"- 扣 {item.points_deducted} 分：{item.reason}" for item in items)

    def _format_screen_facts_summary(self, result: ScoreResult) -> str:
        facts = result.screen_facts
        lines = [
            f"触发状态：{facts.trigger_status or '未识别'}",
            f"获取模式：{facts.acquisition_mode or '未识别'}",
            f"存储深度 / 采样率：{facts.memory_depth_text or '未识别'} / {facts.sample_rate_text or '未识别'}",
            f"水平时基：{facts.timebase_text or '未识别'}",
            f"垂直刻度：{facts.vertical_scale_text or '未识别'}",
            f"有效通道：{', '.join(facts.active_channels) if facts.active_channels else '未识别'}",
        ]
        if facts.high_low_span_divisions > 0:
            lines.append(f"波形高低差估计：约 {facts.high_low_span_divisions:.2f} 格")
        if facts.estimated_vpp_volts > 0:
            lines.append(f"按格数估算的高低电平差：约 {facts.estimated_vpp_volts:.2f} V")
        if facts.ignored_voltage_readouts:
            lines.append(f"已排除的电压读数：{'；'.join(facts.ignored_voltage_readouts)}")
        if facts.ambiguities:
            lines.append(f"仍存在的不确定项：{'；'.join(facts.ambiguities)}")
        return "\n".join(lines)

    def _set_status_style(self, fg_color, text_color) -> None:
        self.status_pill.configure(fg_color=fg_color, text_color=text_color)


def main() -> None:
    app = TeachingEvalApp()
    app.mainloop()


if __name__ == "__main__":
    main()
