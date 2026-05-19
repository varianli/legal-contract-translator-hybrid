import json
import os
import re
import threading
import time
import traceback
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, BooleanVar, Button, Checkbutton, Entry, Frame, Label, OptionMenu, StringVar, Text, Tk, filedialog, messagebox
from tkinter.ttk import Progressbar

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openai import OpenAI

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except Exception:
    TkinterDnD = None
    DND_FILES = None
    DND_AVAILABLE = False


APP_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / "输出"
CONFIG_PATH = APP_DIR / ".hybrid_config.json"

PROVIDER_DEFAULTS = {
    "OpenAI": {"base_url": "", "model": "gpt-4.1", "key_label": "OpenAI API Key"},
    "DeepSeek": {"base_url": "https://api.deepseek.com", "model": "deepseek-v4-flash", "key_label": "DeepSeek API Key"},
}

SYSTEM_PROMPT = """You are a senior bilingual legal translator and document-format alignment specialist.
Translate Chinese legal contracts into precise formal legal English.
Preserve legal meaning, defined terms, numbering, dates, parties, amounts, placeholders, and clause references.
Return only valid JSON when JSON is requested."""

CAPITAL_MARKETS_LEGAL_RAG = """Capital markets legal English retrieval notes:
- This is a capital markets / private equity style legal contract. Prefer formal transactional drafting, not conversational English.
- Translate \u9644\u5f55 as Appendix and keep it consistent. Do not translate \u9644\u5f55 as Schedule.
- Translate \u9644\u4ef6 as Annex. Translate \u9644\u8868 as Schedule.
- Use Roman numerals for Chinese appendix numerals: \u9644\u5f55\u4e00 -> Appendix I; \u9644\u5f55\u4e8c -> Appendix II; \u9644\u5f55\u4e09\uff08A\uff09 -> Appendix III(A); \u9644\u5f55\u4e03 -> Appendix VII.
- Translate \u62ab\u9732\u51fd / \u62ab\u9732\u6e05\u5355 as Disclosure Schedule. If it is titled \u9644\u5f55\u4e94\uff08\u62ab\u9732\u51fd\uff09, use Appendix V (Disclosure Schedule).
- Translate \u4ea4\u5272 as Closing, \u4ea4\u5272\u65e5 as Closing Date, \u6700\u8fdf\u5b8c\u6210\u65e5 as Longstop Date, \u7b7e\u7f72\u65e5 as Signing Date.
- Translate \u672c\u6b21\u4ea4\u6613 as this Transaction, \u6295\u8d44\u6b3e as Investment Amount, \u76ee\u6807\u516c\u53f8 as Target Company, \u96c6\u56e2\u516c\u53f8 as Group Company / Group Companies.
- Preserve defined-term capitalization once established. If the same Chinese term recurs, reuse the same English term.
- Use English punctuation in English output: straight quotes, half-width commas, periods, semicolons, colons, parentheses, and brackets."""

LEGAL_RAG_TERMS = [
    ("\u9644\u5f55", "\u9644\u5f55 = Appendix; use Roman numerals, e.g. Appendix I, Appendix II, Appendix III(A)."),
    ("\u9644\u4ef6", "\u9644\u4ef6 = Annex."),
    ("\u9644\u8868", "\u9644\u8868 = Schedule."),
    ("\u62ab\u9732\u51fd", "\u62ab\u9732\u51fd = Disclosure Schedule."),
    ("\u4ea4\u5272\u65e5", "\u4ea4\u5272\u65e5 = Closing Date."),
    ("\u4ea4\u5272", "\u4ea4\u5272 = Closing."),
    ("\u6700\u8fdf\u5b8c\u6210\u65e5", "\u6700\u8fdf\u5b8c\u6210\u65e5 = Longstop Date."),
    ("\u7b7e\u7f72\u65e5", "\u7b7e\u7f72\u65e5 = Signing Date."),
    ("\u672c\u6b21\u4ea4\u6613", "\u672c\u6b21\u4ea4\u6613 = this Transaction."),
    ("\u6295\u8d44\u6b3e", "\u6295\u8d44\u6b3e = Investment Amount."),
    ("\u76ee\u6807\u516c\u53f8", "\u76ee\u6807\u516c\u53f8 = Target Company."),
    ("\u96c6\u56e2\u516c\u53f8", "\u96c6\u56e2\u516c\u53f8 = Group Company / Group Companies."),
]

CHINESE_NUMERAL_ROMAN = {
    "\u4e00": "I",
    "\u4e8c": "II",
    "\u4e09": "III",
    "\u56db": "IV",
    "\u4e94": "V",
    "\u516d": "VI",
    "\u4e03": "VII",
    "\u516b": "VIII",
    "\u4e5d": "IX",
    "\u5341": "X",
}
ARABIC_ROMAN = {str(index): roman for index, roman in enumerate(["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]) if index}
ROMAN_ARABIC = {roman: arabic for arabic, roman in ARABIC_ROMAN.items()}
APPENDIX_SOURCE_RE = re.compile(r"(\u9644\u5f55|\u9644\u4ef6|\u9644\u8868)([\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\d]+)(?:[\uff08(]\s*([A-Za-z\uff21-\uff3a])\s*[\uff09)])?")
SOURCE_FALLBACK_TARGETS = {
    "\u5b9a\u4e49": ["Definitions"],
    "\u672c\u6b21\u4ea4\u6613\u5b89\u6392": ["Transaction Arrangement", "Transaction Arrangements"],
    "\u4ea4\u5272": ["Closing"],
    "\u4ea4\u5272\u540e\u4e49\u52a1": ["Post-Closing Obligations", "Post-closing Obligations"],
    "\u58f0\u660e\u548c\u4fdd\u8bc1": ["Representations and Warranties"],
}

CJK_PUNCT_TRANSLATION = str.maketrans(
    {
        "\uff0c": ",",
        "\u3002": ".",
        "\uff1b": ";",
        "\uff1a": ":",
        "\uff01": "!",
        "\uff1f": "?",
        "\uff08": "(",
        "\uff09": ")",
        "\u201c": '"',
        "\u201d": '"',
        "\u2018": "'",
        "\u2019": "'",
        "\u3001": ",",
        "\u3010": "[",
        "\u3011": "]",
        "\u300a": '"',
        "\u300b": '"',
    }
)

BLOCK_RE = re.compile(r"<!--\s*BLOCK:(\d+)\s*-->\s*\n(.*?)(?=\n<!--\s*BLOCK:\d+\s*-->\s*\n|\Z)", re.DOTALL)


@dataclass
class FormatSpan:
    span_id: str
    block_id: int
    source_text: str
    style_label: str
    bold: bool
    italic: bool
    underline: object
    highlight_color: object


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_VAL = f"{W_NS}val"
OFF_VALUES = {"0", "false", "off", "none"}


def has_cjk(text: str) -> bool:
    return bool(re.search(r"[\u3400-\u9fff]", text or ""))


def compact_text(text: str, limit: int) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    if len(text) <= limit:
        return text
    half = max(1, limit // 2)
    return text[:half] + "\n...[middle omitted]...\n" + text[-half:]


def legal_rag_for_text(text: str) -> str:
    retrieved = []
    text = text or ""
    for trigger, note in LEGAL_RAG_TERMS:
        if trigger in text and note not in retrieved:
            retrieved.append(note)
    if not retrieved:
        return CAPITAL_MARKETS_LEGAL_RAG
    return CAPITAL_MARKETS_LEGAL_RAG + "\n\nRetrieved terms for this chunk:\n- " + "\n- ".join(retrieved)


def normalize_english_punctuation(text: str) -> str:
    text = (text or "").translate(CJK_PUNCT_TRANSLATION)
    text = re.sub(r"\s+([,.;:!?\]\)])", r"\1", text)
    text = re.sub(r"([\[\(])\s+", r"\1", text)
    text = re.sub(r"([,.;:!?])(?=[A-Za-z0-9\"'])", r"\1 ", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def normalize_roman_token(token: str) -> str:
    token = (token or "").strip().upper()
    if token in CHINESE_NUMERAL_ROMAN:
        return CHINESE_NUMERAL_ROMAN[token]
    if token in ARABIC_ROMAN:
        return ARABIC_ROMAN[token]
    if token in ROMAN_ARABIC:
        return token
    return token


def normalize_legal_english(text: str) -> str:
    text = normalize_english_punctuation(text)

    def appendix_repl(match):
        roman = normalize_roman_token(match.group(1))
        suffix = f"({match.group(2).upper()})" if match.group(2) else ""
        return f"Appendix {roman}{suffix}"

    text = re.sub(r"\bAppendix\s+([1-9]|10|I|II|III|IV|V|VI|VII|VIII|IX|X)(?!\.\d)(?:\s*\(\s*([A-Za-z])\s*\))?", appendix_repl, text, flags=re.IGNORECASE)
    text = re.sub(r"\bSchedule\s+([1-9]|10|I|II|III|IV|V|VI|VII|VIII|IX|X)\s*\(?\s*Disclosure Schedule\s*\)?", lambda m: f"Appendix {normalize_roman_token(m.group(1))} (Disclosure Schedule)", text, flags=re.IGNORECASE)
    text = re.sub(r"\bSchedule\s+([1-9]|10|I|II|III|IV|V|VI|VII|VIII|IX|X)\s+of\s+the\s+Disclosure\s+Schedule\b", lambda m: f"Appendix {normalize_roman_token(m.group(1))} (Disclosure Schedule)", text, flags=re.IGNORECASE)
    text = re.sub(r"\bLong\s+Stop\s+Date\b", "Longstop Date", text, flags=re.IGNORECASE)
    return text


def source_span_fallback_targets(source_text: str) -> list[str]:
    source_text = source_text or ""
    direct_targets = SOURCE_FALLBACK_TARGETS.get(source_text.strip(), [])
    if direct_targets:
        return direct_targets
    match = APPENDIX_SOURCE_RE.search(source_text or "")
    if not match:
        return []
    label_map = {"\u9644\u5f55": "Appendix", "\u9644\u4ef6": "Annex", "\u9644\u8868": "Schedule"}
    label = label_map.get(match.group(1))
    roman = normalize_roman_token(match.group(2))
    if not label or not roman:
        return []
    suffix = ""
    if match.group(3):
        suffix = f"({match.group(3).upper()})"
    targets = [f"{label} {roman}{suffix}"]
    arabic = ROMAN_ARABIC.get(roman)
    if arabic:
        targets.append(f"{label} {arabic}{suffix}")
    return targets


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_config(provider: str, api_key: str, base_url: str, model: str) -> None:
    CONFIG_PATH.write_text(
        json.dumps({"provider": provider, "api_key": api_key, "base_url": base_url, "model": model}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def build_client(api_key: str, base_url: str) -> OpenAI:
    kwargs = {"api_key": api_key, "timeout": 180}
    if base_url.strip():
        kwargs["base_url"] = base_url.strip()
    return OpenAI(**kwargs)


def parse_json_object(raw: str) -> dict:
    raw = (raw or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?", "", raw, flags=re.IGNORECASE).strip()
        raw = re.sub(r"```$", "", raw).strip()
    start = raw.find("{")
    end = raw.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("模型没有返回 JSON 对象。")
    return json.loads(raw[start : end + 1])


def chat_json(client: OpenAI, provider: str, model: str, messages: list[dict], retries: int = 3) -> dict:
    last_error = None
    for attempt in range(retries):
        response_format_options = [True, False] if provider == "DeepSeek" else [True]
        for use_response_format in response_format_options:
            try:
                kwargs = {"model": model, "messages": messages, "temperature": 0.05}
                if use_response_format:
                    kwargs["response_format"] = {"type": "json_object"}
                response = client.chat.completions.create(**kwargs)
                return parse_json_object(response.choices[0].message.content)
            except Exception as exc:
                last_error = exc
        if attempt + 1 < retries:
            time.sleep(2 + attempt * 3)
    raise last_error


def iter_table_paragraphs(table, seen):
    for row in table.rows:
        for cell in row.cells:
            yield from iter_cell_paragraphs(cell, seen)


def iter_cell_paragraphs(cell, seen):
    for paragraph in cell.paragraphs:
        key = id(paragraph._p)
        if key not in seen:
            seen.add(key)
            yield paragraph
    for table in cell.tables:
        yield from iter_table_paragraphs(table, seen)


def iter_part_paragraphs(part, seen):
    for paragraph in part.paragraphs:
        key = id(paragraph._p)
        if key not in seen:
            seen.add(key)
            yield paragraph
    for table in part.tables:
        yield from iter_table_paragraphs(table, seen)


def iter_document_paragraphs(doc: Document):
    seen = set()
    yield from iter_part_paragraphs(doc, seen)
    for section in doc.sections:
        for part in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ):
            yield from iter_part_paragraphs(part, seen)


def paragraph_text(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs)


def paragraph_meta(paragraph) -> dict:
    fmt = paragraph.paragraph_format
    return {
        "style": paragraph.style.name if paragraph.style else "",
        "alignment": str(paragraph.alignment) if paragraph.alignment else "",
        "left_indent": str(fmt.left_indent) if fmt.left_indent else "",
        "first_line_indent": str(fmt.first_line_indent) if fmt.first_line_indent else "",
    }


def markdown_prefix(meta: dict, text: str) -> str:
    style = (meta.get("style") or "").lower()
    stripped = text.strip()
    if "heading 1" in style or style in ("title", "标题 1"):
        return "# "
    if "heading 2" in style or style == "标题 2":
        return "## "
    if "heading 3" in style or style == "标题 3":
        return "### "
    if re.match(r"^第[一二三四五六七八九十百零〇]+条", stripped):
        return "## "
    if re.match(r"^\d+(\.\d+)*\s+", stripped):
        return ""
    return ""


def block_to_markdown(block: dict) -> str:
    meta = block["meta"]
    prefix = markdown_prefix(meta, block["text"])
    meta_json = json.dumps(meta, ensure_ascii=False)
    return f"<!-- BLOCK:{block['id']} -->\n{prefix}{block['text']}\n<!-- META:{meta_json} -->"


def parse_translated_markdown(markdown: str) -> dict[int, str]:
    results = {}
    for match in BLOCK_RE.finditer(markdown.strip()):
        block_id = int(match.group(1))
        body = match.group(2).strip()
        body = re.sub(r"\n<!--\s*META:.*?-->\s*$", "", body, flags=re.DOTALL).strip()
        body = re.sub(r"^#{1,6}\s+", "", body).strip()
        results[block_id] = normalize_legal_english(body)
    return results


def xml_has_off_property(element, tag_name: str) -> bool:
    if element is None:
        return False
    for child in element.iter(f"{W_NS}{tag_name}"):
        value = child.get(W_VAL)
        if value is not None and value.lower() in OFF_VALUES:
            return True
    return False


def xml_has_on_property(element, tag_name: str) -> bool:
    if element is None:
        return False
    for child in element.iter(f"{W_NS}{tag_name}"):
        value = child.get(W_VAL)
        if value is None or value.lower() not in OFF_VALUES:
            return True
    return False


def style_format_defaults(paragraph) -> tuple:
    style = getattr(paragraph, "style", None)
    font = getattr(style, "font", None)
    if font is None:
        return (False, False, None, None)
    return (bool(font.bold), bool(font.italic), font.underline, font.highlight_color)


def run_style_tuple(run, paragraph=None) -> tuple:
    direct = (bool(run.bold), bool(run.italic), run.underline, run.font.highlight_color)
    if paragraph is None:
        return direct
    style_bold, style_italic, style_underline, style_highlight = style_format_defaults(paragraph)
    rpr = run._r.rPr

    bold = True if run.bold is True or xml_has_on_property(rpr, "b") else False
    if not bold and run.bold is not False and not xml_has_off_property(rpr, "b") and style_bold:
        bold = True

    italic = True if run.italic is True or xml_has_on_property(rpr, "i") else False
    if not italic and run.italic is not False and not xml_has_off_property(rpr, "i") and style_italic:
        italic = True

    underline = run.underline
    if not underline and run.underline is not False and not xml_has_off_property(rpr, "u"):
        underline = style_underline if style_underline else (True if xml_has_on_property(rpr, "u") else None)

    highlight = run.font.highlight_color
    if not highlight and not xml_has_off_property(rpr, "highlight"):
        highlight = style_highlight
    return (bold, italic, underline, highlight)


def is_formatted_run(run, paragraph=None) -> bool:
    bold, italic, underline, highlight = run_style_tuple(run, paragraph)
    return bool((bold or italic or underline or highlight) and run.text and run.text.strip())


def style_label(style: tuple) -> str:
    bold, italic, underline, highlight = style
    labels = []
    if bold:
        labels.append("bold")
    if italic:
        labels.append("italic")
    if underline:
        labels.append("underline")
    if highlight:
        labels.append(f"highlight:{highlight}")
    return "+".join(labels)


def extract_format_spans(paragraph, block_id: int) -> list[FormatSpan]:
    spans = []
    current_text = ""
    current_style = None
    span_number = 0

    def flush():
        nonlocal current_text, current_style, span_number
        if current_text.strip() and current_style:
            bold, italic, underline, highlight = current_style
            spans.append(
                FormatSpan(
                    span_id=f"b{block_id}_s{span_number}",
                    block_id=block_id,
                    source_text=current_text,
                    style_label=style_label(current_style),
                    bold=bold,
                    italic=italic,
                    underline=copy(underline),
                    highlight_color=copy(highlight),
                )
            )
            span_number += 1
        current_text = ""
        current_style = None

    for run in paragraph.runs:
        if is_formatted_run(run, paragraph):
            style = run_style_tuple(run, paragraph)
            if style == current_style:
                current_text += run.text
            else:
                flush()
                current_style = style
                current_text = run.text
        else:
            flush()
    flush()
    return spans


def build_chunks(blocks: list[dict], max_chars: int = 35000, max_blocks: int = 45) -> list[list[dict]]:
    chunks = []
    current = []
    current_chars = 0
    for block in blocks:
        block_chars = len(block["markdown"]) + sum(len(span["source_text"]) for span in block["format_spans"])
        if current and (current_chars + block_chars > max_chars or len(current) >= max_blocks):
            chunks.append(current)
            current = []
            current_chars = 0
        current.append(block)
        current_chars += block_chars
    if current:
        chunks.append(current)
    return chunks


def make_translation_memory(client: OpenAI, provider: str, model: str, source_markdown: str, log) -> str:
    log("正在基于 Markdown 结构提取术语表和翻译规则...")
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                "Create a concise bilingual translation memory for this Chinese legal contract represented as Markdown blocks.\n"
                "Return JSON {\"memory\":\"...\"}. Include defined terms, parties, recurring phrases, bracket placeholders, and legal drafting preferences.\n\n"
                f"CAPITAL MARKETS LEGAL RAG:\n{legal_rag_for_text(source_markdown)}\n\n"
                f"MARKDOWN:\n{compact_text(source_markdown, 60000)}"
            ),
        },
    ]
    try:
        return str(chat_json(client, provider, model, messages, retries=2).get("memory", "")).strip()
    except Exception as exc:
        log(f"术语记忆提取失败，将继续翻译。原因：{exc}")
        return ""


def translate_markdown_chunk(client: OpenAI, provider: str, model: str, memory: str, chunk_markdown: str, before: str, after: str) -> str:
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                "Translate this Markdown-structured Chinese legal contract chunk into precise formal legal English.\n"
                "Return JSON {\"markdown\":\"...\"}.\n"
                "Rules:\n"
                "1. Preserve every <!-- BLOCK:n --> marker exactly.\n"
                "2. Preserve Markdown heading/list/paragraph structure where reasonable.\n"
                "3. Translate the visible Chinese text into English using the whole chunk context.\n"
                "4. Do not translate or remove <!-- META:... --> comments.\n"
                "5. Do not leave Chinese characters unless intentionally part of a name or exhibit label.\n\n"
                "Workflow reminder:\n"
                "A. Markdown carries clause hierarchy, indentation/list cues, and some visible formatting signals, but it is not the complete Word run structure.\n"
                "B. Translate under the Markdown structure first so the English keeps context and legal coherence.\n"
                "C. Keep each BLOCK independent after translation because run-level formatting will be checked and applied block by block.\n\n"
                f"CAPITAL MARKETS LEGAL RAG:\n{legal_rag_for_text(chunk_markdown)}\n\n"
                f"TRANSLATION MEMORY:\n{memory or '(none)'}\n\n"
                f"CONTEXT BEFORE:\n{compact_text(before, 3000)}\n\n"
                f"CONTEXT AFTER:\n{compact_text(after, 3000)}\n\n"
                f"CHUNK MARKDOWN:\n{chunk_markdown}"
            ),
        },
    ]
    return str(chat_json(client, provider, model, messages).get("markdown", "")).strip()


def map_format_spans(client: OpenAI, provider: str, model: str, memory: str, source_blocks: list[dict], translated_blocks: dict[int, str]) -> list[dict]:
    payload = []
    for block in source_blocks:
        spans = block["format_spans"]
        if spans:
            payload.append(
                {
                    "block_id": block["id"],
                    "source_text": block["text"],
                    "translation": translated_blocks.get(block["id"], ""),
                    "format_spans": spans,
                }
            )
    if not payload:
        return []
    payload_text = "\n".join(f"{item['source_text']}\n{item['translation']}" for item in payload)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                "Map each formatted Chinese source span to the exact English substring in the translated block that should receive the same formatting.\n"
                "Return JSON {\"mappings\":[{\"span_id\":\"...\",\"block_id\":1,\"target_text\":\"...\",\"confidence\":\"high|medium|low\",\"note\":\"...\"}]}.\n"
                "Rules:\n"
                "1. Every input span_id must appear once.\n"
                "2. target_text should be an exact contiguous substring of the English translation whenever possible.\n"
                "3. If the concept is split in English, choose the closest legally equivalent contiguous phrase and mark confidence low.\n"
                "4. Do not invent styling; only map the supplied spans.\n\n"
                "Mandatory block-by-block checklist:\n"
                "A. Treat Markdown as already carrying only part of the document structure/formatting.\n"
                "B. For each block, inspect the run-derived format_spans as the checklist of formatting not fully represented by Markdown.\n"
                "C. For every span, record the formatted source content, locate the legally corresponding exact phrase in that block's English translation, and return it as target_text.\n"
                "D. Finish all spans in the current block before moving to the next block.\n"
                "E. If a target is uncertain, still return the closest contiguous legally equivalent English phrase and explain the uncertainty in note.\n\n"
                f"CAPITAL MARKETS LEGAL RAG:\n{legal_rag_for_text(payload_text)}\n\n"
                f"TRANSLATION MEMORY:\n{memory or '(none)'}\n\n"
                f"BLOCKS JSON:\n{json.dumps({'blocks': payload}, ensure_ascii=False)}"
            ),
        },
    ]
    return chat_json(client, provider, model, messages).get("mappings", []) or []


def nearest_nonspace(text: str, index: int, direction: int) -> str:
    while 0 <= index < len(text):
        if not text[index].isspace():
            return text[index]
        index += direction
    return ""


def candidate_match_score(text: str, start: int, end: int, exact_case: bool) -> int:
    score = 20 if exact_case else 0
    before = nearest_nonspace(text, start - 1, -1)
    after = nearest_nonspace(text, end, 1)
    if before in {'"', "'", "\u201c", "\u2018"} and after in {'"', "'", "\u201d", "\u2019"}:
        score += 100
    if before in {"(", "["} and after in {")", "]"}:
        score += 20
    if start == 0 or not text[start - 1].isalnum():
        score += 8
    else:
        score -= 50
    if end == len(text) or not text[end : end + 1].isalnum():
        score += 8
    else:
        score -= 50
    prefix = text[max(0, start - 8) : start].lower()
    if '"' in prefix or "\u201c" in prefix or "(" in prefix:
        score += 5
    return score


def find_available_span(text: str, needle: str, occupied: list[tuple[int, int]]) -> tuple[int, int] | None:
    if not needle:
        return None
    candidates = []
    for candidate in [needle, needle.strip(), re.sub(r"\s+", " ", needle).strip(), normalize_legal_english(needle)]:
        if candidate and candidate not in candidates:
            candidates.append(candidate)
    for candidate in candidates:
        if not candidate:
            continue
        matches = []
        start = text.find(candidate)
        while start != -1:
            end = start + len(candidate)
            if all(end <= a or start >= b for a, b in occupied):
                matches.append((candidate_match_score(text, start, end, True), start, end))
            start = text.find(candidate, start + 1)
        lower_text = text.lower()
        lower_candidate = candidate.lower()
        start = lower_text.find(lower_candidate)
        while start != -1:
            end = start + len(candidate)
            if all(end <= a or start >= b for a, b in occupied):
                matches.append((candidate_match_score(text, start, end, False), start, end))
            start = lower_text.find(lower_candidate, start + 1)
        if matches:
            matches.sort(key=lambda item: (-item[0], item[1]))
            _, start, end = matches[0]
            return start, end
    return None


def clear_paragraph_runs(paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            paragraph._p.remove(child)


def paragraph_rpr(paragraph):
    ppr = paragraph._p.pPr
    return None if ppr is None else ppr.find(qn("w:rPr"))


def base_format_overrides(paragraph) -> dict[str, bool]:
    overrides = {"bold": False, "italic": False, "underline": False, "highlight": False}
    rpr = paragraph_rpr(paragraph)
    if xml_has_off_property(rpr, "b"):
        overrides["bold"] = True
    if xml_has_off_property(rpr, "i"):
        overrides["italic"] = True
    if xml_has_off_property(rpr, "u"):
        overrides["underline"] = True
    if xml_has_off_property(rpr, "highlight"):
        overrides["highlight"] = True

    for run in paragraph.runs:
        if not run.text:
            continue
        if run.bold is False:
            overrides["bold"] = True
        if run.italic is False:
            overrides["italic"] = True
        if run.underline is False:
            overrides["underline"] = True
        if xml_has_off_property(run._r.rPr, "highlight"):
            overrides["highlight"] = True
    return overrides


def set_highlight_none(run) -> None:
    rpr = run._r.get_or_add_rPr()
    for existing in list(rpr.findall(qn("w:highlight"))):
        rpr.remove(existing)
    highlight = OxmlElement("w:highlight")
    highlight.set(qn("w:val"), "none")
    rpr.append(highlight)


def apply_base_format(run, base_format: dict[str, bool]) -> None:
    if base_format.get("bold"):
        run.bold = False
    if base_format.get("italic"):
        run.italic = False
    if base_format.get("underline"):
        run.underline = False
    if base_format.get("highlight"):
        set_highlight_none(run)


def add_run_with_base_format(paragraph, text: str, base_format: dict[str, bool]):
    run = paragraph.add_run(text)
    apply_base_format(run, base_format)
    return run


def apply_style(run, span: FormatSpan, base_format: dict[str, bool]) -> None:
    apply_base_format(run, base_format)
    if span.bold:
        run.bold = True
    if span.italic:
        run.italic = True
    if span.underline:
        run.underline = span.underline
    if span.highlight_color:
        run.font.highlight_color = span.highlight_color


def rewrite_paragraph(paragraph, translation: str, intervals: list[dict]) -> None:
    intervals = sorted(intervals, key=lambda item: (item["start"], item["end"]))
    base_format = base_format_overrides(paragraph)
    clear_paragraph_runs(paragraph)
    cursor = 0
    for item in intervals:
        start = max(0, item["start"])
        end = min(len(translation), item["end"])
        if start < cursor or end <= start:
            continue
        if cursor < start:
            add_run_with_base_format(paragraph, translation[cursor:start], base_format)
        run = add_run_with_base_format(paragraph, translation[start:end], base_format)
        apply_style(run, item["span"], base_format)
        cursor = end
    if cursor < len(translation):
        add_run_with_base_format(paragraph, translation[cursor:], base_format)
    if not translation:
        add_run_with_base_format(paragraph, "", base_format)


def safe_save_document(doc: Document, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_name(f"{output_path.stem}.tmp.{int(time.time() * 1000)}{output_path.suffix}")
    doc.save(str(tmp_path))
    try:
        os.replace(str(tmp_path), str(output_path))
        return output_path
    except OSError as exc:
        if getattr(exc, "winerror", None) != 5 and getattr(exc, "errno", None) != 13:
            tmp_path.unlink(missing_ok=True)
            raise
        fallback = output_path.with_name(f"{output_path.stem}_自动另存_{time.strftime('%Y%m%d_%H%M%S')}{output_path.suffix}")
        os.replace(str(tmp_path), str(fallback))
        return fallback


def checklist_markdown(rows: list[dict]) -> str:
    lines = [
        "# 复合方法 Checklist",
        "",
        "| 状态 | Block | 原格式文本 | 格式 | 英文映射 | 置信度 | 备注 |",
        "|---|---:|---|---|---|---|---|",
    ]
    for row in rows:
        lines.append(
            "| {status} | {block_id} | {source_text} | {style} | {target_text} | {confidence} | {note} |".format(
                status=row.get("status", ""),
                block_id=row.get("block_id", ""),
                source_text=str(row.get("source_text", "")).replace("|", "\\|"),
                style=str(row.get("style", "")).replace("|", "\\|"),
                target_text=str(row.get("target_text", "")).replace("|", "\\|"),
                confidence=str(row.get("confidence", "")).replace("|", "\\|"),
                note=str(row.get("note", "")).replace("|", "\\|"),
            )
        )
    return "\n".join(lines) + "\n"


def translate_docx_hybrid(
    input_path: Path,
    output_dir: Path,
    provider: str,
    api_key: str,
    base_url: str,
    model: str,
    log,
    progress,
    cancel_event: threading.Event | None = None,
) -> tuple[Path, Path, Path, Path, Path, bool]:
    output_dir.mkdir(parents=True, exist_ok=True)
    doc = Document(str(input_path))
    paragraphs = list(iter_document_paragraphs(doc))
    blocks = []
    spans_by_id: dict[str, FormatSpan] = {}

    for block_id, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph)
        if not has_cjk(text):
            continue
        spans = extract_format_spans(paragraph, block_id)
        public_spans = []
        for span in spans:
            spans_by_id[span.span_id] = span
            public_spans.append({"span_id": span.span_id, "source_text": span.source_text, "style": span.style_label})
        block = {"id": block_id, "text": text, "meta": paragraph_meta(paragraph), "format_spans": public_spans}
        block["markdown"] = block_to_markdown(block)
        blocks.append(block)

    if not blocks:
        raise ValueError("没有在文档中发现中文内容。")

    source_markdown = "\n\n".join(block["markdown"] for block in blocks)
    client = build_client(api_key, base_url)
    memory = make_translation_memory(client, provider, model, source_markdown, log)
    chunks = build_chunks(blocks)
    log(f"复合方法：{len(blocks)} 个中文 block，分为 {len(chunks)} 个 Markdown 大段批次。")

    translated_blocks: dict[int, str] = {}
    translated_markdown_parts = []
    all_mappings = []
    completed = 0
    cancelled = False
    for chunk_index, chunk in enumerate(chunks, start=1):
        if cancel_event and cancel_event.is_set():
            cancelled = True
            log("收到中止请求，停止进入下一个 Markdown 批次，开始导出当前进度。")
            break
        chunk_markdown = "\n\n".join(block["markdown"] for block in chunk)
        before = "\n\n".join(block["markdown"] for block in blocks[max(0, blocks.index(chunk[0]) - 3) : blocks.index(chunk[0])])
        after_start = blocks.index(chunk[-1]) + 1
        after = "\n\n".join(block["markdown"] for block in blocks[after_start : min(len(blocks), after_start + 3)])
        progress(completed, len(blocks), f"翻译 Markdown 大段 {chunk_index}/{len(chunks)}")
        translated_markdown = translate_markdown_chunk(client, provider, model, memory, chunk_markdown, before, after)
        translated_markdown_parts.append(translated_markdown)
        chunk_translations = parse_translated_markdown(translated_markdown)
        translated_blocks.update(chunk_translations)
        all_mappings.extend(map_format_spans(client, provider, model, memory, chunk, chunk_translations))
        completed += len(chunk)
        if cancel_event and cancel_event.is_set():
            cancelled = True
            log("当前 Markdown 批次已完成；根据中止请求，开始导出当前进度。")
            progress(completed, len(blocks), "已中止，正在导出当前进度")
            break
        progress(completed, len(blocks), f"大段 {chunk_index}/{len(chunks)} 已完成")

    mappings_by_span = {str(mapping.get("span_id")): mapping for mapping in all_mappings if mapping.get("span_id")}
    checklist = []
    if cancelled:
        checklist.append(
            {
                "status": "CANCELLED",
                "block_id": "",
                "source_text": "",
                "style": "",
                "target_text": "",
                "confidence": "",
                "note": "用户中止；已导出当前已完成批次，未完成 block 保留原文。",
            }
        )
    for block in blocks:
        block_id = block["id"]
        translation = normalize_legal_english(translated_blocks.get(block_id, "").strip())
        intervals = []
        occupied = []
        if not translation:
            checklist.append({"status": "MISSING_TRANSLATION", "block_id": block_id, "note": "模型没有返回该 block 的译文。"})
            continue
        for public_span in block["format_spans"]:
            span_id = public_span["span_id"]
            span = spans_by_id[span_id]
            mapping = mappings_by_span.get(span_id)
            row = {
                "status": "",
                "block_id": block_id,
                "source_text": span.source_text,
                "style": span.style_label,
                "target_text": "",
                "confidence": "",
                "note": "",
            }
            if not mapping:
                fallback_found = None
                fallback_target = ""
                for candidate in source_span_fallback_targets(span.source_text):
                    fallback_target = normalize_legal_english(candidate)
                    fallback_found = find_available_span(translation, fallback_target, occupied)
                    if fallback_found:
                        break
                if fallback_found:
                    start, end = fallback_found
                    intervals.append({"start": start, "end": end, "span": span})
                    occupied.append((start, end))
                    row["target_text"] = fallback_target
                    row["confidence"] = "deterministic"
                    row["status"] = "APPLIED_FALLBACK"
                    row["note"] = "Applied deterministic legal-reference mapping because the model omitted this span."
                    checklist.append(row)
                    continue
                row["status"] = "MISSING_MAPPING"
                row["note"] = "模型没有返回该格式 span 的英文映射。"
                checklist.append(row)
                continue
            target_text = normalize_legal_english(str(mapping.get("target_text", "")).strip())
            row["target_text"] = target_text
            row["confidence"] = str(mapping.get("confidence", "")).strip()
            row["note"] = str(mapping.get("note", "")).strip()
            found = find_available_span(translation, target_text, occupied)
            if not found:
                for candidate in source_span_fallback_targets(span.source_text):
                    target_text = normalize_legal_english(candidate)
                    found = find_available_span(translation, target_text, occupied)
                    if found:
                        row["target_text"] = target_text
                        row["confidence"] = "deterministic"
                        row["note"] = (row["note"] + " " if row["note"] else "") + "Used deterministic legal-reference fallback."
                        break
            if not found:
                row["status"] = "TARGET_NOT_FOUND"
                row["note"] = (row["note"] + " " if row["note"] else "") + "英文映射未在译文中精确找到。"
                checklist.append(row)
                continue
            start, end = found
            intervals.append({"start": start, "end": end, "span": span})
            occupied.append((start, end))
            row["status"] = "APPLIED"
            checklist.append(row)
        if has_cjk(translation):
            checklist.append({"status": "HAS_CHINESE_IN_TRANSLATION", "block_id": block_id, "note": "译文仍检测到中文。"})
        rewrite_paragraph(paragraphs[block_id], translation, intervals)

    base = input_path.stem
    suffix = "_已中止" if cancelled else ""
    docx_path = safe_save_document(doc, output_dir / f"{base}_复合方法英文翻译.docx")
    source_md_path = output_dir / f"{base}_source_blocks.md"
    translated_md_path = output_dir / f"{base}_translated_blocks.md"
    checklist_path = output_dir / f"{base}_复合方法checklist.md"
    json_path = output_dir / f"{base}_复合方法明细.json"
    if cancelled:
        cancelled_docx_path = docx_path.with_name(f"{docx_path.stem}{suffix}{docx_path.suffix}")
        os.replace(str(docx_path), str(cancelled_docx_path))
        docx_path = cancelled_docx_path
        translated_md_path = translated_md_path.with_name(f"{translated_md_path.stem}{suffix}{translated_md_path.suffix}")
        checklist_path = checklist_path.with_name(f"{checklist_path.stem}{suffix}{checklist_path.suffix}")
        json_path = json_path.with_name(f"{json_path.stem}{suffix}{json_path.suffix}")

    source_md_path.write_text(source_markdown, encoding="utf-8")
    translated_md_path.write_text("\n\n".join(translated_markdown_parts), encoding="utf-8")
    checklist_path.write_text(checklist_markdown(checklist), encoding="utf-8")
    json_path.write_text(
        json.dumps({"cancelled": cancelled, "blocks": blocks, "mappings": all_mappings, "checklist": checklist}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    progress(completed if cancelled else len(blocks), len(blocks), "已中止并导出当前进度" if cancelled else "全部完成")
    return docx_path, source_md_path, translated_md_path, checklist_path, json_path, cancelled


class HybridApp:
    def __init__(self, root):
        self.root = root
        self.main_thread_id = threading.get_ident()
        self.root.title("复合方法 - Markdown 结构 + run 格式映射")
        self.root.geometry("880x700")
        self.root.minsize(800, 640)

        config = load_config()
        provider = config.get("provider", "DeepSeek")
        if provider not in PROVIDER_DEFAULTS:
            provider = "DeepSeek"
        defaults = PROVIDER_DEFAULTS[provider]
        self.provider = StringVar(value=provider)
        self.api_key = StringVar(value=config.get("api_key", ""))
        self.base_url = StringVar(value=config.get("base_url", defaults["base_url"]))
        self.model = StringVar(value=config.get("model", defaults["model"]))
        self.file_path = StringVar(value="")
        self.output_dir = StringVar(value=str(OUTPUT_DIR))
        self.remember_key = BooleanVar(value=bool(config.get("api_key")))
        self.cancel_event = threading.Event()

        top = Frame(root, padx=16, pady=12)
        top.pack(fill=X)
        Label(top, text="API 类型").pack(anchor="w")
        OptionMenu(top, self.provider, *PROVIDER_DEFAULTS.keys(), command=self.on_provider_change).pack(fill=X, pady=(4, 8))
        self.key_label = Label(top, text=defaults["key_label"])
        self.key_label.pack(anchor="w")
        Entry(top, textvariable=self.api_key, show="*", width=90).pack(fill=X, pady=(4, 8))
        Label(top, text="Base URL").pack(anchor="w")
        Entry(top, textvariable=self.base_url, width=90).pack(fill=X, pady=(4, 8))
        Label(top, text="模型").pack(anchor="w")
        Entry(top, textvariable=self.model, width=90).pack(fill=X, pady=(4, 8))
        Checkbutton(top, text="记住 API Key（明文保存在本机复合方法目录）", variable=self.remember_key).pack(anchor="w")

        file_frame = Frame(root, padx=16, pady=8)
        file_frame.pack(fill=X)
        Label(file_frame, text="合同文件（.docx）").pack(anchor="w")
        row = Frame(file_frame)
        row.pack(fill=X, pady=(4, 8))
        Entry(row, textvariable=self.file_path).pack(side=LEFT, fill=X, expand=True)
        Button(row, text="选择文件", command=self.choose_file, width=12).pack(side=RIGHT, padx=(8, 0))
        Label(file_frame, text="输出目录").pack(anchor="w")
        out_row = Frame(file_frame)
        out_row.pack(fill=X, pady=(4, 0))
        Entry(out_row, textvariable=self.output_dir).pack(side=LEFT, fill=X, expand=True)
        Button(out_row, text="选择目录", command=self.choose_output_dir, width=12).pack(side=RIGHT, padx=(8, 0))

        self.drop_label = Label(root, text="把 Word 合同拖到这里\n输出：DOCX + source Markdown + translated Markdown + checklist + JSON", relief="groove", height=4)
        self.drop_label.pack(fill=X, padx=16, pady=8)
        if DND_AVAILABLE:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self.on_drop)

        progress_frame = Frame(root, padx=16, pady=8)
        progress_frame.pack(fill=X)
        self.progress_text = StringVar(value="进度：0/0")
        self.current_text = StringVar(value="当前：尚未开始")
        Label(progress_frame, textvariable=self.progress_text).pack(anchor="w")
        self.bar = Progressbar(progress_frame, maximum=100)
        self.bar.pack(fill=X, pady=(4, 6))
        Label(progress_frame, textvariable=self.current_text, wraplength=820, justify=LEFT).pack(anchor="w")

        action = Frame(root, padx=16, pady=8)
        action.pack(fill=X)
        self.start_button = Button(action, text="开始复合方法翻译", command=self.start, height=2)
        self.start_button.pack(side=LEFT, fill=X, expand=True)
        self.stop_button = Button(action, text="中止并导出当前进度", command=self.request_cancel, height=2, state="disabled")
        self.stop_button.pack(side=RIGHT, padx=(8, 0))

        self.log_box = Text(root, height=14, wrap="word")
        self.log_box.pack(fill=BOTH, expand=True, padx=16, pady=(4, 16))
        self.log("复合方法：Markdown 负责结构和上下文，run 结构负责格式抽取，模型负责格式片段到英文短语的映射。")

    def on_provider_change(self, provider):
        defaults = PROVIDER_DEFAULTS[provider]
        self.key_label.config(text=defaults["key_label"])
        self.base_url.set(defaults["base_url"])
        self.model.set(defaults["model"])

    def choose_file(self):
        path = filedialog.askopenfilename(title="选择 Word 合同", filetypes=[("Word 文档", "*.docx")])
        if path:
            self.file_path.set(path)

    def choose_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录", initialdir=self.output_dir.get() or str(OUTPUT_DIR))
        if path:
            self.output_dir.set(path)

    def on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        if paths:
            self.file_path.set(paths[0])

    def log(self, text: str):
        if threading.get_ident() != self.main_thread_id:
            self.root.after(0, self.log, text)
            return
        self.log_box.insert(END, text + "\n")
        self.log_box.see(END)
        self.root.update_idletasks()

    def update_progress(self, done: int, total: int, current: str):
        if threading.get_ident() != self.main_thread_id:
            self.root.after(0, self.update_progress, done, total, current)
            return
        percent = 0 if total <= 0 else max(0, min(100, done * 100 / total))
        self.bar["value"] = percent
        self.progress_text.set(f"进度：{done}/{total}（{percent:.1f}%）")
        self.current_text.set(f"当前：{current}")
        self.root.update_idletasks()

    def start(self):
        provider = self.provider.get().strip()
        api_key = self.api_key.get().strip()
        base_url = self.base_url.get().strip()
        model = self.model.get().strip()
        input_path = Path(self.file_path.get().strip().strip('"'))
        output_dir = Path(self.output_dir.get().strip().strip('"'))
        if not api_key:
            messagebox.showerror("缺少 API Key", f"请先输入 {provider} API Key。")
            return
        if not model:
            messagebox.showerror("缺少模型", "请填写模型名称。")
            return
        if not input_path.exists() or input_path.suffix.lower() != ".docx":
            messagebox.showerror("文件不支持", "请选择存在的 .docx 文件。")
            return
        if provider == "DeepSeek" and not base_url:
            base_url = PROVIDER_DEFAULTS["DeepSeek"]["base_url"]
            self.base_url.set(base_url)
        if self.remember_key.get():
            save_config(provider, api_key, base_url, model)
        self.cancel_event.clear()
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        thread = threading.Thread(target=self.worker, args=(input_path, output_dir, provider, api_key, base_url, model), daemon=True)
        thread.start()

    def request_cancel(self):
        self.cancel_event.set()
        self.stop_button.config(state="disabled")
        self.log("已请求中止；当前 API 批次返回后会导出当前进度。")
        self.update_progress(0, 0, "正在等待当前批次返回，然后导出当前进度...")

    def worker(self, input_path: Path, output_dir: Path, provider: str, api_key: str, base_url: str, model: str):
        try:
            paths = translate_docx_hybrid(input_path, output_dir, provider, api_key, base_url, model, self.log, self.update_progress, self.cancel_event)
            self.root.after(0, self.finish_success, paths[:-1], paths[-1])
        except Exception as exc:
            self.log("发生错误：")
            self.log(str(exc))
            self.log(traceback.format_exc())
            self.root.after(0, self.finish_error, str(exc))

    def finish_success(self, paths, cancelled: bool):
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        title = "已中止并导出" if cancelled else "完成"
        messagebox.showinfo(title, "已输出：\n" + "\n".join(str(path) for path in paths))

    def finish_error(self, error: str):
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        messagebox.showerror("翻译失败", error)


def main():
    root_class = TkinterDnD.Tk if DND_AVAILABLE else Tk
    root = root_class()
    HybridApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
