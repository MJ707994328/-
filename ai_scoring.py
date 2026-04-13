from __future__ import annotations

import base64
import http.client
import json
import mimetypes
import os
import time
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path

THIS_DIR = Path(__file__).resolve().parent
LOCAL_CONFIG_PATH = Path(
    os.getenv("AI_SCORING_LOCAL_CONFIG", str(THIS_DIR / ".local_secrets.json"))
)

WAVEFORM_WEIGHT = 0.70
TIME_WEIGHT = 0.30

DEFAULT_OPENAI_BASE_URL = "https://api.openai.com/v1"
DEFAULT_MOONSHOT_BASE_URL = "https://api.moonshot.cn/v1"

DEFAULT_OPENAI_MODEL = os.getenv("OPENAI_SCORING_MODEL", "gpt-5.4-mini")
DEFAULT_MOONSHOT_MODEL = os.getenv("MOONSHOT_SCORING_MODEL", "kimi-k2.5")

TIME_RULE_TEXT = (
    "时间得分从 100 分开始计算。若实际用时超过预期用时，则按超时百分比扣分。"
    "例如超时 20%，时间部分扣 20 分。"
)

OBJECTIVE_NOTICE = (
    "输入框中的文本按教师给出的实验目标处理。"
    "系统只用它来判断最终波形是否符合目标，不会把它当作学生作文或实验报告来评分。"
)

GDS1000E_SCREEN_READING_GUIDE = """
请严格按照 GDS-1000E 中文操作手册中的屏幕说明来判读界面：
1. 左上方的 “10k pts”“2MSa/s”等信息表示存储深度和采样率，不是波形幅值。
2. 上方 Memory Bar 表示波形在内存中的位置和触发位置，不代表波形高度。
3. Trigger Status 的含义必须按手册理解：
   Trig’d = 已触发；PrTrig = 预触发；Trig? = 未触发，屏幕不更新；
   Stop = 触发停止；Roll = 滚动模式；Auto = 自动触发模式。
4. Trigger Configuration 区域显示触发源、斜率、触发电平、耦合。
   其中触发电平不是信号幅值。
5. Horizontal Status 显示水平时基和水平位置，例如 500us/div、1ms/div。
6. Channel Status 显示通道、耦合方式和垂直刻度，例如 CH1、DC、2V/Div。
   注意：2V/Div 是每格电压刻度，不是波形振幅或峰峰值。
7. 通道颜色按手册对应：CH1 黄色，CH2 蓝色，CH3 粉色，CH4 绿色。
8. 只有纯色、清晰激活的通道状态和对应波形轨迹才应视为开启通道；灰色或暗淡的占位标记不要轻易当作有效输入信号。
9. 零电压位置标记只表示参考零电平，不表示波形峰值。
10. 底部若明确显示自动测量项（如 Frequency、Pk-Pk、Max、Min、Amplitude），这些才可作为明确测量值引用。
11. 如果界面没有明确显示某项测量值，就不要臆造该数值；幅值只能通过“明确测量值”或“格数 x V/div 的估算”来判断，并说明依据。
12. 如果实验目标里明确要求某个电压幅值，而界面上没有明确出现 Vpp、Amplitude、Max、Min 等电压测量项，
    则必须优先用“波形上下电平之间占据的垂直格数 x V/div”来估算高低电平差，不允许直接用右下角单个电压读数替代幅值结论。
13. 对于方波截图，若波形上下电平之间明显跨越约 4 个小格，而 CH1 垂直刻度为 500mV/div，
    那么应优先认为高低电平差约为 2V，而不是约 1V，除非屏幕明确显示 Vpp/Amplitude 等相反证据。
"""


class AIScoringError(RuntimeError):
    pass


@dataclass(slots=True)
class DeductionItem:
    reason: str
    points_deducted: int

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


@dataclass(slots=True)
class ScreenFacts:
    trigger_status: str
    acquisition_mode: str
    memory_depth_text: str
    sample_rate_text: str
    timebase_text: str
    vertical_scale_text: str
    channel_status_text: str
    frequency_readout_text: str
    active_channels: list[str]
    vertical_scale_volts_per_div: float
    high_low_span_divisions: float
    estimated_vpp_volts: float
    amplitude_source: str
    amplitude_evidence: str
    ignored_voltage_readouts: list[str]
    screen_observations: list[str]
    ambiguities: list[str]

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


@dataclass(slots=True)
class ScoreResult:
    model: str
    provider: str
    overall_score: int
    verdict: str
    summary: str
    waveform_score: int
    time_score: int
    waveform_weight: float
    time_weight: float
    actual_duration_seconds: int
    expected_duration_seconds: int
    overtime_seconds: int
    overtime_percent: float
    screen_facts: ScreenFacts
    waveform_summary: str
    screen_observations: list[str]
    amplitude_evidence: str
    waveform_deductions: list[DeductionItem]
    time_deductions: list[DeductionItem]
    strengths: list[str]
    instructor_feedback: str
    confidence: str
    objective_notice: str
    time_rule: str
    scoring_formula: str

    def to_dict(self) -> dict[str, object]:
        payload = asdict(self)
        payload["screen_facts"] = self.screen_facts.to_dict()
        payload["waveform_deductions"] = [item.to_dict() for item in self.waveform_deductions]
        payload["time_deductions"] = [item.to_dict() for item in self.time_deductions]
        return payload


def is_ai_scoring_configured() -> bool:
    return bool(_resolve_api_key())


def score_experiment(
    *,
    description: str,
    duration_seconds: int,
    expected_duration_seconds: int,
    final_image_path: str | Path,
    model: str | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
) -> ScoreResult:
    if expected_duration_seconds <= 0:
        raise AIScoringError("预期时长必须大于 0 秒。")

    local_config = _load_local_config()
    resolved_api_key = api_key or _resolve_api_key()
    if not resolved_api_key:
        raise AIScoringError("当前未配置 AI API Key，请设置 MOONSHOT_API_KEY、KIMI_API_KEY 或 OPENAI_API_KEY。")

    image_path = Path(final_image_path)
    if not image_path.exists():
        raise AIScoringError(f"未找到最终截图文件：{image_path}")

    provider = _resolve_provider(base_url=base_url, local_config=local_config)
    resolved_base_url = (base_url or _resolve_base_url(provider)).rstrip("/")
    chosen_model = model or _resolve_default_model(provider)

    screen_facts = _extract_screen_facts(
        provider=provider,
        model=chosen_model,
        base_url=resolved_base_url,
        api_key=resolved_api_key,
        objective_text=description,
        final_image_path=image_path,
    )

    waveform_payload_text = _request_waveform_score(
        provider=provider,
        model=chosen_model,
        base_url=resolved_base_url,
        api_key=resolved_api_key,
        objective_text=description,
        screen_facts=screen_facts,
    )
    parsed = _parse_json_text(waveform_payload_text)

    waveform_deductions = _parse_deduction_items(parsed.get("waveform_deductions", []))
    waveform_score = max(0, 100 - sum(item.points_deducted for item in waveform_deductions))

    time_score, overtime_seconds, overtime_percent, time_deductions = _calculate_time_score(
        actual_duration_seconds=duration_seconds,
        expected_duration_seconds=expected_duration_seconds,
    )

    overall_score = round((waveform_score * WAVEFORM_WEIGHT) + (time_score * TIME_WEIGHT))

    return ScoreResult(
        model=chosen_model,
        provider=provider,
        overall_score=overall_score,
        verdict=_verdict_from_score(overall_score),
        summary=_build_summary(
            overall_score=overall_score,
            waveform_score=waveform_score,
            time_score=time_score,
            overtime_percent=overtime_percent,
        ),
        waveform_score=waveform_score,
        time_score=time_score,
        waveform_weight=WAVEFORM_WEIGHT,
        time_weight=TIME_WEIGHT,
        actual_duration_seconds=duration_seconds,
        expected_duration_seconds=expected_duration_seconds,
        overtime_seconds=overtime_seconds,
        overtime_percent=overtime_percent,
        screen_facts=screen_facts,
        waveform_summary=str(parsed["waveform_summary"]),
        screen_observations=screen_facts.screen_observations,
        amplitude_evidence=screen_facts.amplitude_evidence,
        waveform_deductions=waveform_deductions,
        time_deductions=time_deductions,
        strengths=[str(item) for item in parsed["strengths"]],
        instructor_feedback=str(parsed["instructor_feedback"]),
        confidence=str(parsed["confidence"]),
        objective_notice=OBJECTIVE_NOTICE,
        time_rule=TIME_RULE_TEXT,
        scoring_formula="总分 = 波形得分 * 70% + 时间得分 * 30%",
    )


def save_score_report(result: ScoreResult, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(result.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def _extract_screen_facts(
    *,
    provider: str,
    model: str,
    base_url: str,
    api_key: str,
    objective_text: str,
    final_image_path: Path,
) -> ScreenFacts:
    image_url = _image_to_data_url(final_image_path)
    objective = objective_text.strip() or "（未填写教师实验目标）"

    system_prompt = (
        "你是一位非常严谨的示波器读屏助手。"
        "你的任务不是直接评分，而是先把 GDS-1000E 的屏幕内容结构化提取出来。"
        f"{GDS1000E_SCREEN_READING_GUIDE}\n"
        "所有输出请使用中文，并且只返回 JSON。"
    )
    user_prompt = (
        "请先读取这张 GDS-1000E 示波器截图，并提取结构化屏幕事实。\n\n"
        f"教师给定的实验目标：\n{objective}\n\n"
        "请返回一个 JSON 对象，且必须包含这些字段：\n"
        "{\n"
        '  "trigger_status": string，\n'
        '  "acquisition_mode": string，\n'
        '  "memory_depth_text": string，\n'
        '  "sample_rate_text": string，\n'
        '  "timebase_text": string，\n'
        '  "vertical_scale_text": string，\n'
        '  "channel_status_text": string，\n'
        '  "frequency_readout_text": string，\n'
        '  "active_channels": [string, ...]，\n'
        '  "vertical_scale_volts_per_div": number，\n'
        '  "high_low_span_divisions": number，\n'
        '  "amplitude_source": string，\n'
        '  "screen_observations": [string, ... 3 到 6 项]，\n'
        '  "amplitude_evidence": string，\n'
        '  "ignored_voltage_readouts": [string, ... 0 到 4 项]，\n'
        '  "ambiguities": [string, ... 0 到 4 项]\n'
        "}\n\n"
        "规则：\n"
        "1. 这里只做事实提取，不做扣分、不做总评。\n"
        "2. 如果没有明确的 Vpp/Amplitude/Max/Min 测量值，就必须优先按垂直格数 x V/div 估算幅值。\n"
        "3. 不允许直接把右下角单个电压值当作方波幅值，除非它明确标注为 Vpp、Amplitude、Max 或 Min。\n"
        "4. 如果某些读数不确定，请在 ambiguities 中说明。\n"
        "5. 所有文字都用中文。\n"
        "6. 只输出 JSON。"
    )

    if provider == "moonshot":
        payload = {
            "model": model,
            "temperature": 1,
            "messages": [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": user_prompt},
                        {"type": "image_url", "image_url": {"url": image_url}},
                    ],
                },
            ],
        }
        response_json = _post_chat_completions_request(
            base_url=base_url,
            api_key=api_key,
            payload=payload,
            provider=provider,
        )
        return _parse_screen_facts(_extract_chat_completions_text(response_json))

    payload = {
        "model": model,
        "temperature": 0.2,
        "store": False,
        "input": [
            {
                "role": "developer",
                "content": [{"type": "input_text", "text": system_prompt}],
            },
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": user_prompt},
                    {"type": "input_image", "image_url": image_url, "detail": "high"},
                ],
            },
        ],
    }
    response_json = _post_responses_request(base_url=base_url, api_key=api_key, payload=payload)
    return _parse_screen_facts(_extract_responses_text(response_json))


def _request_waveform_score(
    *,
    provider: str,
    model: str,
    base_url: str,
    api_key: str,
    objective_text: str,
    screen_facts: ScreenFacts,
) -> str:
    objective = objective_text.strip() or "（未填写教师实验目标）"
    facts_json = json.dumps(screen_facts.to_dict(), ensure_ascii=False, indent=2)

    system_prompt = (
        "你是一位严格但公平的大学电子实验教师。"
        "现在不要重新看图，而是只根据已经提取好的结构化读屏结果来做波形评分。"
        "输入文字是教师给定的实验目标，不是学生自评。"
        "不要按文笔打分。"
        "波形部分采用扣分制：从 100 分开始，列出明确扣分项和整数扣分值。"
        "所有输出请使用中文，并且只返回 JSON。"
    )
    user_prompt = (
        "下面是本次实验的教师目标，以及已经从示波器截图中提取出的结构化读屏结果。\n\n"
        f"教师目标：\n{objective}\n\n"
        f"结构化读屏结果：\n{facts_json}\n\n"
        "请只根据这些结构化事实做波形评分，不要重新臆测图片里还有什么。\n"
        "请返回一个 JSON 对象，且必须包含这些字段：\n"
        "{\n"
        '  "waveform_summary": string，\n'
        '  "waveform_deductions": [\n'
        '    {"reason": string, "points_deducted": integer}，最多 5 项\n'
        "  ],\n"
        '  "strengths": [string, ... 2 到 4 项],\n'
        '  "instructor_feedback": string，\n'
        '  "confidence": "low" | "medium" | "high"\n'
        "}\n\n"
        "规则：\n"
        "1. 若结构化结果中的 estimated_vpp_volts 已经由 V/div 和格数估算得出，则把它当作主要幅值依据。\n"
        "2. 若 ignored_voltage_readouts 里提到某些电压读数不应作为幅值，则不要再拿它们反向扣分。\n"
        "3. 若 ambiguities 非空，请在教师反馈中适当提示不确定性。\n"
        "4. 这里只评波形，不评时间。\n"
        "5. 所有文字都用中文。\n"
        "6. 只输出 JSON。"
    )

    if provider == "moonshot":
        payload = {
            "model": model,
            "temperature": 1,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        response_json = _post_chat_completions_request(
            base_url=base_url,
            api_key=api_key,
            payload=payload,
            provider=provider,
        )
        return _extract_chat_completions_text(response_json)

    payload = {
        "model": model,
        "temperature": 0.2,
        "store": False,
        "input": [
            {
                "role": "developer",
                "content": [{"type": "input_text", "text": system_prompt}],
            },
            {
                "role": "user",
                "content": [{"type": "input_text", "text": user_prompt}],
            },
        ],
    }
    response_json = _post_responses_request(base_url=base_url, api_key=api_key, payload=payload)
    return _extract_responses_text(response_json)


def _calculate_time_score(*, actual_duration_seconds: int, expected_duration_seconds: int) -> tuple[int, int, float, list[DeductionItem]]:
    overtime_seconds = max(0, actual_duration_seconds - expected_duration_seconds)
    if overtime_seconds == 0:
        return (
            100,
            0,
            0.0,
            [DeductionItem(reason="实际用时未超过预期时长，因此时间部分不扣分。", points_deducted=0)],
        )

    overtime_percent = (overtime_seconds / expected_duration_seconds) * 100
    deduction_points = min(100, round(overtime_percent))
    time_score = max(0, 100 - deduction_points)
    reason = (
        f"实际用时比预期时长超出 {overtime_percent:.1f}% "
        f"（实际 {actual_duration_seconds} 秒，预期 {expected_duration_seconds} 秒）。"
        f"按照时间规则，时间部分扣 {deduction_points} 分。"
    )
    return time_score, overtime_seconds, overtime_percent, [DeductionItem(reason=reason, points_deducted=deduction_points)]


def _build_summary(*, overall_score: int, waveform_score: int, time_score: int, overtime_percent: float) -> str:
    if overtime_percent > 0:
        return (
            f"综合得分 {overall_score}/100。波形得分 {waveform_score}/100，时间得分 {time_score}/100。"
            f"本次总分主要由波形表现决定，同时超时带来了时间扣分。"
        )
    return (
        f"综合得分 {overall_score}/100。波形得分 {waveform_score}/100，时间得分 {time_score}/100。"
        "本次实验在预期时间内完成，因此时间部分没有扣分。"
    )


def _verdict_from_score(score: int) -> str:
    if score >= 90:
        return "优秀"
    if score >= 80:
        return "良好"
    if score >= 70:
        return "合格"
    if score >= 60:
        return "临界"
    return "需改进"


def _parse_deduction_items(raw_items: object) -> list[DeductionItem]:
    items: list[DeductionItem] = []
    if not isinstance(raw_items, list):
        return items
    for raw in raw_items:
        if not isinstance(raw, dict):
            continue
        reason = str(raw.get("reason", "")).strip()
        if not reason:
            continue
        points = _clamp_score(raw.get("points_deducted"))
        items.append(DeductionItem(reason=reason, points_deducted=points))
    return items


def _parse_screen_facts(raw_text: str) -> ScreenFacts:
    parsed = _parse_json_text(raw_text)

    vertical_scale_volts_per_div = _parse_float(parsed.get("vertical_scale_volts_per_div"))
    high_low_span_divisions = _parse_float(parsed.get("high_low_span_divisions"))
    estimated_vpp_volts = 0.0
    if vertical_scale_volts_per_div > 0 and high_low_span_divisions > 0:
        estimated_vpp_volts = round(vertical_scale_volts_per_div * high_low_span_divisions, 4)

    return ScreenFacts(
        trigger_status=str(parsed.get("trigger_status", "")).strip(),
        acquisition_mode=str(parsed.get("acquisition_mode", "")).strip(),
        memory_depth_text=str(parsed.get("memory_depth_text", "")).strip(),
        sample_rate_text=str(parsed.get("sample_rate_text", "")).strip(),
        timebase_text=str(parsed.get("timebase_text", "")).strip(),
        vertical_scale_text=str(parsed.get("vertical_scale_text", "")).strip(),
        channel_status_text=str(parsed.get("channel_status_text", "")).strip(),
        frequency_readout_text=str(parsed.get("frequency_readout_text", "")).strip(),
        active_channels=[str(item).strip() for item in parsed.get("active_channels", []) if str(item).strip()],
        vertical_scale_volts_per_div=vertical_scale_volts_per_div,
        high_low_span_divisions=high_low_span_divisions,
        estimated_vpp_volts=estimated_vpp_volts,
        amplitude_source=str(parsed.get("amplitude_source", "")).strip(),
        amplitude_evidence=str(parsed.get("amplitude_evidence", "")).strip(),
        ignored_voltage_readouts=[str(item).strip() for item in parsed.get("ignored_voltage_readouts", []) if str(item).strip()],
        screen_observations=[str(item).strip() for item in parsed.get("screen_observations", []) if str(item).strip()],
        ambiguities=[str(item).strip() for item in parsed.get("ambiguities", []) if str(item).strip()],
    )


def _post_responses_request(*, base_url: str, api_key: str, payload: dict[str, object]) -> dict[str, object]:
    url = f"{base_url}/responses"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    if os.getenv("OPENAI_PROJECT"):
        headers["OpenAI-Project"] = os.getenv("OPENAI_PROJECT", "")
    if os.getenv("OPENAI_ORGANIZATION"):
        headers["OpenAI-Organization"] = os.getenv("OPENAI_ORGANIZATION", "")

    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )

    return _perform_json_request(request, label="Responses 接口")


def _post_chat_completions_request(
    *,
    base_url: str,
    api_key: str,
    payload: dict[str, object],
    provider: str,
) -> dict[str, object]:
    url = f"{base_url}/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    return _perform_json_request(request, label=f"{provider} 对话接口")


def _extract_responses_text(response_json: dict[str, object]) -> str:
    output_text = response_json.get("output_text")
    if isinstance(output_text, str) and output_text.strip():
        return output_text.strip()

    for item in response_json.get("output", []):
        if not isinstance(item, dict):
            continue
        for content in item.get("content", []):
            if not isinstance(content, dict):
                continue
            if content.get("type") == "output_text" and isinstance(content.get("text"), str):
                return content["text"].strip()

    raise AIScoringError("无法从 Responses 接口中提取文本输出。")


def _extract_chat_completions_text(response_json: dict[str, object]) -> str:
    choices = response_json.get("choices")
    if not isinstance(choices, list) or not choices:
        raise AIScoringError("对话接口返回结果中缺少 choices。")
    message = choices[0].get("message", {})
    content = message.get("content")
    if isinstance(content, str) and content.strip():
        return content.strip()
    raise AIScoringError("无法从对话接口中提取文本输出。")


def _parse_json_text(raw_text: str) -> dict[str, object]:
    text = raw_text.strip()
    candidates = [text]

    if "```" in text:
        stripped = text.replace("```json", "```").replace("```JSON", "```")
        parts = stripped.split("```")
        for part in parts:
            part = part.strip()
            if part:
                candidates.append(part)

    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        candidates.append(text[start:end + 1].strip())

    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, dict):
            return parsed

    raise AIScoringError("模型返回内容不是合法 JSON。")


def _resolve_api_key() -> str | None:
    local_config = _load_local_config()
    return (
        os.getenv("MOONSHOT_API_KEY")
        or os.getenv("KIMI_API_KEY")
        or os.getenv("OPENAI_API_KEY")
        or _config_value(local_config, "api_key", "moonshot_api_key", "kimi_api_key", "openai_api_key")
    )


def _resolve_provider(*, base_url: str | None, local_config: dict[str, object] | None = None) -> str:
    local_config = local_config or _load_local_config()
    if base_url and "moonshot" in base_url:
        return "moonshot"
    if base_url and "openai" in base_url:
        return "openai"
    if os.getenv("MOONSHOT_API_KEY") or os.getenv("KIMI_API_KEY"):
        return "moonshot"
    configured_provider = str(local_config.get("provider", "")).strip().lower()
    if configured_provider in {"moonshot", "kimi"}:
        return "moonshot"
    configured_model = str(local_config.get("model", "")).strip().lower()
    configured_base_url = str(_config_value(local_config, "base_url", "moonshot_base_url", "openai_base_url") or "").lower()
    if "moonshot" in configured_base_url or configured_model.startswith("kimi"):
        return "moonshot"
    return "openai"


def _resolve_base_url(provider: str) -> str:
    local_config = _load_local_config()
    if provider == "moonshot":
        return str(
            os.getenv("MOONSHOT_BASE_URL")
            or _config_value(local_config, "base_url", "moonshot_base_url")
            or DEFAULT_MOONSHOT_BASE_URL
        ).rstrip("/")
    return str(
        os.getenv("OPENAI_BASE_URL")
        or _config_value(local_config, "openai_base_url")
        or DEFAULT_OPENAI_BASE_URL
    ).rstrip("/")


def _resolve_default_model(provider: str) -> str:
    local_config = _load_local_config()
    if provider == "moonshot":
        return str(
            os.getenv("MOONSHOT_SCORING_MODEL")
            or _config_value(local_config, "model", "moonshot_model", "kimi_model")
            or DEFAULT_MOONSHOT_MODEL
        )
    return str(
        os.getenv("OPENAI_SCORING_MODEL")
        or _config_value(local_config, "openai_model")
        or DEFAULT_OPENAI_MODEL
    )


def _clamp_score(value: object) -> int:
    try:
        score = int(float(value))
    except Exception:
        return 0
    if score < 0:
        return 0
    if score > 100:
        return 100
    return score


def _parse_float(value: object) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        cleaned = []
        dot_used = False
        sign_used = False
        for ch in text:
            if ch.isdigit():
                cleaned.append(ch)
            elif ch == "." and not dot_used:
                cleaned.append(ch)
                dot_used = True
            elif ch in "+-" and not sign_used and not cleaned:
                cleaned.append(ch)
                sign_used = True
        candidate = "".join(cleaned)
        if candidate and candidate not in {".", "+", "-"}:
            try:
                return float(candidate)
            except ValueError:
                return 0.0
    return 0.0


def _image_to_data_url(image_path: Path) -> str:
    mime_type, _ = mimetypes.guess_type(image_path.name)
    if not mime_type:
        mime_type = "image/png"
    encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def _load_local_config() -> dict[str, object]:
    if not LOCAL_CONFIG_PATH.exists():
        return {}
    try:
        raw = json.loads(LOCAL_CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if isinstance(raw, dict):
        return raw
    return {}


def _config_value(config: dict[str, object], *keys: str) -> str | None:
    for key in keys:
        value = config.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return None


def _perform_json_request(request: urllib.request.Request, *, label: str) -> dict[str, object]:
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="replace")
            raise AIScoringError(f"{label} 返回 HTTP {exc.code}: {details}") from exc
        except (urllib.error.URLError, http.client.RemoteDisconnected, TimeoutError) as exc:
            last_error = exc
            if attempt == 2:
                break
            time.sleep(1.2 * (attempt + 1))
    raise AIScoringError(f"无法连接到 {label}：{last_error}")
