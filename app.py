"""Resume Roast: a small, no-login Streamlit demo with no intentional file persistence."""

from __future__ import annotations

import io
import json
import os
import re
import urllib.error
import urllib.request
from typing import Any

import streamlit as st
from docx import Document
from pypdf import PdfReader

# Load .env if python-dotenv is installed; fail gracefully if it isn't.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ModuleNotFoundError:
    pass


def get_secret(key: str, default: str = "") -> str:
    val = os.getenv(key)
    if val:
        return val
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return default


st.set_page_config(page_title="Resume Roast", page_icon="✳", layout="centered")

WORKSHOP_URL = get_secret("WORKSHOP_URL", "https://nxtwave.tech/")
OLLAMA_API_KEY = get_secret("OLLAMA_API_KEY", "")
OLLAMA_MODEL = get_secret("OLLAMA_MODEL", "beam:cloud")
OLLAMA_API_URL = get_secret("OLLAMA_API_URL", "https://ollama.com/api/chat")
OLLAMA_TIMEOUT_SECONDS = max(30, int(get_secret("OLLAMA_TIMEOUT_SECONDS", "180")))
MAX_FILE_BYTES = 5 * 1024 * 1024
MAX_TEXT_CHARS = 18_000

STYLES = """
<style>
    :root {
        color-scheme: light;
        --text-color: #20221f;
        --background-color: #fbfbfa;
        --secondary-background-color: #ffffff;
        --primary-color: #20221f;
        --border-color: #e3e5df;
    }
    .stApp { background: #fbfbfa; color: #20221f; }
    [data-testid="stHeader"] { background: rgba(251,251,250,.9); }
    .block-container { max-width: 780px; padding-top: 2.2rem; padding-bottom: 4rem; }
    .eyebrow { color:#62665e; font-size:.78rem; font-weight:700; letter-spacing:.13em; text-transform:uppercase; }
    .hero { font-size:3.1rem; line-height:1.03; letter-spacing:-.055em; font-weight:750; margin:.4rem 0 .8rem; }
    .intro { color:#62665e; font-size:1.08rem; max-width:620px; line-height:1.65; }
    .score { font-size:3.7rem; font-weight:760; letter-spacing:-.06em; line-height:1; }
    .muted { color:#74776f; }
    div.stButton > button[kind="primary"] { background:#20221f; border-color:#20221f; border-radius:10px; min-height:2.8rem; }
    div.stButton > button[kind="primary"]:hover { background:#393c36; border-color:#393c36; }
    [data-testid="stWidgetLabel"] p { color:#20221f !important; opacity:1 !important; }
    [data-testid="stSegmentedControl"] label,
    [data-testid="stSegmentedControl"] button,
    [data-testid="stSegmentedControl"] button p,
    [data-testid="stSegmentedControl"] p { color:#20221f !important; opacity:1 !important; }
    a { color:#20221f; }
    .cta-card {
        background: #20221f;
        color: #fbfbfa;
        border-radius: 14px;
        padding: 1.4rem 1.6rem 1.5rem;
        margin-top: 1.6rem;
        margin-bottom: 1.8rem;
    }
    .cta-card .cta-eyebrow {
        font-size: .72rem;
        font-weight: 700;
        letter-spacing: .13em;
        text-transform: uppercase;
        color: #a3a89e;
        margin-bottom: .35rem;
    }
    .cta-card .cta-heading {
        font-size: 1.22rem;
        font-weight: 720;
        letter-spacing: -.03em;
        line-height: 1.25;
        color: #fbfbfa;
        margin-bottom: .5rem;
    }
    .cta-card .cta-body {
        font-size: .95rem;
        color: #c8ccc4;
        line-height: 1.6;
        margin-bottom: 1.1rem;
    }
    .cta-card a.cta-btn {
        display: inline-block;
        background: #fbfbfa;
        color: #20221f !important;
        font-weight: 650;
        font-size: .93rem;
        padding: .6rem 1.25rem;
        border-radius: 8px;
        text-decoration: none;
        letter-spacing: -.01em;
    }
    .cta-card a.cta-btn:hover { background: #e3e5df; }
</style>
"""
st.markdown(STYLES, unsafe_allow_html=True)


def _extract_pdf(data: bytes) -> str:
    """Extract text from PDF, trying layout-aware mode first, then plain."""
    reader = PdfReader(io.BytesIO(data), strict=False)
    pages: list[str] = []
    for page in reader.pages:
        try:
            text = page.extract_text(extraction_mode="layout") or ""
        except TypeError:
            text = ""
        if not text.strip():
            text = page.extract_text() or ""
        pages.append(text)
    return "\n".join(pages)


def _extract_docx(data: bytes) -> str:
    """Extract text from DOCX including tables, headers, and footers."""
    document = Document(io.BytesIO(data))
    parts: list[str] = []
    for para in document.paragraphs:
        stripped = para.text.strip()
        if stripped:
            parts.append(stripped)
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if cells:
                parts.append("\t".join(cells))
    for section in document.sections:
        for hf in (section.header, section.footer):
            if hf is not None:
                for para in hf.paragraphs:
                    stripped = para.text.strip()
                    if stripped:
                        parts.append(stripped)
    return "\n".join(parts)


def extract_text(filename: str, data: bytes) -> str:
    suffix = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if suffix == "pdf":
        return _extract_pdf(data)
    if suffix == "docx":
        return _extract_docx(data)
    if suffix == "txt":
        return data.decode("utf-8-sig", errors="replace")
    raise ValueError("Please upload a PDF, DOCX, or TXT file.")


MODES = ["Standard", "Spicy"]

TONE_GUIDANCE = {
    "Standard": (
        "Balanced and recruiter-style. Be honest but fair. "
        "Name specific weak spots, skip fluff, and give concrete actionable next steps. "
        "No cheerleading, no cruelty."
    ),
    "Spicy 🌶️": (
        "Savage and culturally sharp. Roast this resume hard by comparing it to what "
        "the tech industry actually expects in 2025: LLM/AI literacy, working demos on GitHub, "
        "quantified impact, contributions to open-source, side projects that actually shipped, "
        "and skills that haven't been deprecated. "
        "Call out buzzwords that every bootcamp grad uses (HTML, CSS, 'team player'). "
        "Mock vague bullets with no numbers. Drag any skill that a model can now do better. "
        "Reference real 2024-2025 trends: the vibe-coding era, AI agents replacing junior devs, "
        "RAG pipelines, multimodal models, Cursor/Copilot fluency as baseline, not a differentiator. "
        "Be funny and specific — aim at the document, not the person. "
        "Mild profanity is fine when it lands. Do NOT produce generic corporate coaching."
    ),
}


def local_preview(text: str, severity: str) -> dict[str, Any]:
    """Rules-based preview used when no AI key is configured."""
    low = text.lower()
    signals = {
        "Projects or portfolio evidence": any(x in low for x in ("project", "github", "portfolio", "built ")),
        "Impact or measurable outcomes": bool(re.search(r"\b\d+(?:\.\d+)?\s*(%|users|clients|ms|x|k|m)\b", low)),
        "Technical tools or skills": any(x in low for x in ("python", "java", "sql", "react", "machine learning", "api", "cloud")),
        "Experience or leadership": any(x in low for x in ("intern", "experience", "lead", "volunteer", "hackathon")),
    }
    present = [label for label, found in signals.items() if found]
    missing = [label for label, found in signals.items() if not found]
    score = max(25, min(92, 42 + 12 * len(present)))
    if severity == "Spicy":
        headline = "Written before LLMs could do your job better. Let's fix that."
        roast = [
            f"No {item.lower()} — in 2025 that's table stakes, not optional." for item in missing[:3]
        ] or ["Suspiciously vague throughout. A RAG pipeline would hallucinate better bullets than this."]
    else:
        headline = "Solid foundation — let's add proof where it's missing."
        roast = [f"I couldn't spot {item.lower()} — add one concrete example." for item in missing[:3]] or [
            "The resume is light on specifics. Quantify at least two bullets."
        ]
    return {
        "score": score,
        "headline": headline,
        "strengths": present or ["You have a starting point to build from."],
        "roast": roast,
        "next_steps": [
            "Link one real project (GitHub or live URL) with the stack you used.",
            "Rewrite two bullets: action verb + measurable result.",
            "Add one AI/ML-adjacent tool or project — it's what recruiters scan for first in 2025.",
        ],
        "mode": "Preview (rules-based — add an Ollama Cloud key for AI feedback)",
    }


def _parse_review(content: str) -> dict[str, Any]:
    content = content.strip()
    fenced = re.search(r"```(?:json)?\s*(.*?)\s*```", content, flags=re.I | re.S)
    if fenced:
        content = fenced.group(1).strip()
    start, end = content.find("{"), content.rfind("}")
    if start >= 0 and end > start:
        content = content[start : end + 1]
    result = json.loads(content)
    if not isinstance(result, dict):
        raise ValueError("The AI response was not a JSON object.")
    if not isinstance(result.get("score"), (int, float)):
        raise ValueError("The AI response did not include a valid score.")
    for field in ("headline", "strengths", "roast", "next_steps"):
        expected = list if field in {"strengths", "roast", "next_steps"} else str
        if not isinstance(result.get(field), expected):
            raise ValueError(f"The AI response did not include valid {field} feedback.")
    result["score"] = max(0, min(100, round(result["score"])))
    for field in ("strengths", "roast", "next_steps"):
        result[field] = [str(item).strip() for item in result[field] if str(item).strip()]
    if not result["roast"] or not result["next_steps"]:
        raise ValueError("The AI response was incomplete. Please retry the review.")
    return result


def ollama_review(text: str, severity: str) -> dict[str, Any]:
    if not OLLAMA_API_KEY:
        return local_preview(text, severity)

    style = TONE_GUIDANCE.get(severity, TONE_GUIDANCE["Standard"])
    prompt = f"""Review the resume below. Selected tone: {severity}.
Tone instructions: {style}
Use only facts visible in the resume. Never invent projects, metrics, skills, or outcomes. Evaluate evidence, clarity, project depth, skills, and impact. Do not make a hiring decision or infer personal/protected traits.

Return one compact JSON object only, with exactly these keys:
{{"score": 0, "headline": "one short sentence", "strengths": ["up to 3 short strings"], "roast": ["2 short, specific strings"], "next_steps": ["exactly 3 concise actions"]}}
Keep every array item under 22 words and the headline under 16 words. Do not use markdown or put line breaks inside strings.

Resume text:
---
{text[:MAX_TEXT_CHARS]}
---"""
    messages = [
        {"role": "system", "content": "You are a resume reviewer. Follow the selected tone. Return a complete, valid JSON object and nothing else."},
        {"role": "user", "content": prompt},
    ]

    def request_content(messages_to_send: list[dict[str, str]]) -> tuple[str, str]:
        payload = {
            "model": OLLAMA_MODEL,
            "messages": messages_to_send,
            "think": False,
            "options": {"temperature": 0.45, "num_predict": 1800},
            "stream": False,
        }
        request = urllib.request.Request(
            OLLAMA_API_URL,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {OLLAMA_API_KEY}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            method="POST",
        )
        with urllib.request.urlopen(request, timeout=OLLAMA_TIMEOUT_SECONDS) as response:
            body = json.loads(response.read().decode("utf-8"))
        content = body.get("message", {}).get("content")
        if not isinstance(content, str) or not content.strip():
            finish_reason = body.get("done_reason", "unknown")
            raise ValueError(f"Ollama Cloud returned no final answer (finish reason: {finish_reason}). Please try again.")
        return content, str(body.get("done_reason", "unknown"))

    finish_reason = "unknown"
    try:
        content, finish_reason = request_content(messages)
        result = _parse_review(content)
    except (json.JSONDecodeError, ValueError):
        # Ask for a fresh, shorter response once; don't expose an opaque JSON parser traceback to the user.
        retry_messages = [
            {"role": "system", "content": "Output exactly one complete, compact JSON object. No markdown, analysis, or text outside the JSON."},
            {"role": "user", "content": prompt + f"\n\nYour previous response was incomplete (finish reason: {finish_reason}). Regenerate a shorter complete JSON object. Keep each item very brief."},
        ]
        try:
            content, finish_reason = request_content(retry_messages)
            result = _parse_review(content)
        except (json.JSONDecodeError, ValueError) as exc:
            raise ValueError("The AI returned an incomplete review twice. Please retry, or upload a text-based PDF/DOCX with less content.") from exc

    result["mode"] = f"AI feedback · {severity}"
    return result


def render_result(result: dict[str, Any]) -> None:
    with st.container(border=True):
        left, right = st.columns([1, 3])
        with left:
            st.markdown(f'<div class="score">{int(result["score"])}</div><div class="muted">out of 100</div>', unsafe_allow_html=True)
        with right:
            st.markdown(f"### {result['headline']}")
            st.caption(result.get("mode", "AI feedback"))

    st.markdown("#### What’s working")
    for item in result.get("strengths", []):
        st.markdown(f"- {item}")
    st.markdown("#### The roast")
    for item in result.get("roast", []):
        st.markdown(f"- {item}")
    st.markdown("#### Your next moves")
    for index, item in enumerate(result.get("next_steps", []), start=1):
        st.markdown(f"{index}. {item}")

    score = result.get("score", 50)
    if score >= 70:
        rec_heading = "Start now make it undeniable."
        rec_body = (
            "You've got the foundation. The fastest way to separate yourself "
            "from other candidates is a real, deployed AI project with your name on it. "
            "Our hands-on workshop gets you from zero to a live project in one session."
        )
    elif score >= 45:
        rec_heading = "Good bones. Time to add the proof."
        rec_body = (
            "Recruiters want evidence, not claims. The single highest-leverage move "
            "right now is shipping an AI project you can link from your resume. "
            "Join the workshop and build one today — no prior ML experience needed."
        )
    else:
        rec_heading = "The fastest fix? Ship something real."
        rec_body = (
            "A resume without projects is a list of promises. In 2025, even one "
            "deployed AI project outweighs pages of bullet points. "
            "Our free workshop walks you through building and launching your first one — start today."
        )

    st.markdown(
        f"""
        <div class="cta-card">
            <div class="cta-eyebrow">Recommended next step</div>
            <div class="cta-heading">{rec_heading}</div>
            <div class="cta-body">{rec_body}</div>
            <a class="cta-btn" href="{WORKSHOP_URL}" target="_blank">Join the free workshop →</a>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.markdown('<div class="eyebrow"> </div>', unsafe_allow_html=True)
st.markdown('<div class="hero">Roast my resume.</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="intro">Get a quick, constructive review of your projects, skills, and impact. '
    'Choose the tone. Take the feedback. Then build something worth adding.</div>',
    unsafe_allow_html=True,
)

if "roast_result" in st.session_state:
    render_result(st.session_state["roast_result"])
    if st.button("Review another resume"):
        st.session_state.pop("roast_result", None)
        st.session_state["upload_key"] = st.session_state.get("upload_key", 0) + 1
        st.rerun()
    st.markdown("<div class='muted'>Your resume text is not saved by this app. Close or refresh this page to clear the result.</div>", unsafe_allow_html=True)
else:
    with st.container(border=True):
        severity = st.segmented_control(
            "Pick your roast level",
            MODES,
            default="Standard",
            selection_mode="single",
            help="Standard = honest recruiter feedback. Spicy = dragged against 2025 tech trends.",
        )
        severity = severity or "Standard"
        upload_key = st.session_state.setdefault("upload_key", 0)
        uploaded = st.file_uploader(
            "Drop your resume here",
            type=["pdf", "docx", "txt"],
            accept_multiple_files=False,
            max_upload_size=5,
            key=f"resume_upload_{upload_key}",
            help="PDF, DOCX, or TXT · up to 5 MB",
        )
        consent_copy = "I agree to have the extracted resume text analyzed for this review."
        if OLLAMA_API_KEY:
            consent_copy += " I understand it will be sent to Ollama Cloud and handled under its data terms."
        consent = st.checkbox(consent_copy)
        
        if not OLLAMA_API_KEY:
            st.info("Ollama Cloud is not connected yet. Add OLLAMA_API_KEY to your local .env file to enable Beam AI feedback. This session will use the clearly labeled preview mode until then.")
        analyze = st.button("Roast my resume", type="primary", use_container_width=True)

    if analyze:
        if not uploaded:
            st.error("Upload a resume first.")
        elif not consent:
            st.error("Please agree before sending your resume for analysis.")
        elif uploaded.size > MAX_FILE_BYTES:
            st.error("That file is larger than 5 MB.")
        else:
            try:
                with st.spinner("Reading your resume and finding the useful bits…"):
                    resume_text = extract_text(uploaded.name, uploaded.getvalue()).strip()
                    if len(resume_text) < 80:
                        raise ValueError("I couldn't read enough text. Try an exported PDF, DOCX, or plain-text resume.")
                    result = ollama_review(resume_text, severity)
                st.session_state["roast_result"] = result
                # Replace the uploader widget so the uploaded bytes are released on rerun.
                st.session_state["upload_key"] = upload_key + 1
                del resume_text
                st.rerun()
            except (ValueError, json.JSONDecodeError, KeyError, IndexError) as exc:
                st.error(str(exc) or "The resume could not be read. Try a different file.")
            except urllib.error.HTTPError as exc:
                detail = exc.read(500).decode("utf-8", errors="replace").strip()
                st.error(f"Ollama Cloud returned HTTP {exc.code}. Check OLLAMA_API_KEY, the model name ({OLLAMA_MODEL}), and your account access. {detail}")
            except (urllib.error.URLError, TimeoutError) as exc:
                if isinstance(exc, TimeoutError) or isinstance(getattr(exc, "reason", None), TimeoutError):
                    st.error(f"Ollama Cloud did not return within {OLLAMA_TIMEOUT_SECONDS} seconds. Try again, or check model availability and account access.")
                else:
                    st.error(f"Ollama Cloud could not be reached. Check your connection and OLLAMA_API_URL, then try again. ({exc})")

st.divider()
st.markdown("<div class='muted'>Keep personal contact details out if you prefer.</div>", unsafe_allow_html=True)
