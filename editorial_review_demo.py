"""
AI-Assisted Editorial Review Tool — DEMO VERSION
(Uses mock AI responses to demonstrate output without API key)

For production: replace mock_analyze_chapter() with real OpenAI/Claude API call.

Setup:  pip install python-docx
Usage:  python editorial_review_demo.py input.docx output_annotated.docx
"""

import sys
import json
import re
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── CONFIG ────────────────────────────────────────────────────────────────────
CLOSING_HEADING_TEXT = "A manner of closing"

# ── MOCK AI (replace with real API call in production) ────────────────────────
def mock_analyze_chapter(chapter_text: str, chapter_num: int, is_odd: bool) -> dict:
    """
    DEMO: Detects likely misspellings with simple heuristics.
    In production, replace this with an OpenAI/Claude API call.
    """
    import re

    words = re.findall(r'\b[a-zA-Z]+\b', chapter_text)
    common_misspelling_patterns = [
        r'.+ley$',    # slowley
        r'.+tion$',   # magnifiscent won't match but others might
        r'^joruney$', r'^begining$', r'^futher$', r'^magnifiscent$',
        r'^qestion$', r'^wrold$', r'^remebered$', r'^sucseeded$',
        r'^obstackles$', r'^straange$', r'^recogniced$', r'^trainning$',
        r'^microscop$', r'^specimin$', r'^biologgy$', r'^carefuly$',
        r'^documanted$', r'^discoverey$', r'^coarse$', r'^sience$',
        r'^accross$', r'^aproached$', r'^exausted$', r'^relived$',
        r'^dammage$', r'^expresion$', r'^jurney$', r'^tole$',
        r'^eveything$', r'^compleet$', r'^misssion$', r'^holed$',
        r'^slowley$', r'^stret$', r'^futher$',
    ]

    flagged = []
    for word in set(w.lower() for w in words):
        for pattern in common_misspelling_patterns:
            if re.match(pattern, word):
                flagged.append(word)
                break

    # Find closing position: last non-empty paragraph start
    paras = [p.strip() for p in chapter_text.split('\n') if p.strip()]
    closing_pos = ' '.join(paras[-1].split()[:6]) if paras else ""

    result = {
        "flagged_words": flagged,
        "closing_position": closing_pos,
        "closing_rationale": "Final paragraph — natural closing point"
    }

    if is_odd and paras:
        result["emphasis_sentence"] = paras[0][:120] if paras else ""
        result["blockquote_sentence"] = paras[1][:120] if len(paras) > 1 else ""

    return result


def call_openai_api(chapter_text: str, chapter_num: int, is_odd: bool) -> dict:
    """
    PRODUCTION: Replace mock with real API call.
    Uncomment and set your API key.
    """
    import requests

    # OPENAI_API_KEY = "sk-..."  # or load from env
    # or for Claude:
    # ANTHROPIC_API_KEY = "sk-ant-..."

    odd_instruction = ""
    if is_odd:
        odd_instruction = '\n4. "emphasis_sentence": one sentence best for emphasis.\n5. "blockquote_sentence": one sentence best as block quote.'

    prompt = f"""Analyze this manuscript chapter. Return ONLY JSON with:
1. "flagged_words": list of misspelled/suspicious/typographic words
2. "closing_position": first 5-7 words of paragraph near end to insert closing heading before
3. "closing_rationale": brief reason{odd_instruction}

Chapter {chapter_num}:
---
{chapter_text[:3000]}
---
Return ONLY valid JSON."""

    # OpenAI example:
    # r = requests.post("https://api.openai.com/v1/chat/completions",
    #     headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
    #     json={"model": "gpt-4o-mini", "messages": [{"role": "user", "content": prompt}], "max_tokens": 500})
    # raw = r.json()["choices"][0]["message"]["content"]

    # Claude example:
    # r = requests.post("https://api.anthropic.com/v1/messages",
    #     headers={"x-api-key": ANTHROPIC_API_KEY, "anthropic-version": "2023-06-01", "Content-Type": "application/json"},
    #     json={"model": "claude-haiku-4-5-20251001", "max_tokens": 500, "messages": [{"role": "user", "content": prompt}]})
    # raw = r.json()["content"][0]["text"]

    raw = re.sub(r"```json|```", "", raw).strip()
    return json.loads(raw)


# ── DOCX HELPERS ──────────────────────────────────────────────────────────────
HIGHLIGHT_MAP = {
    "yellow":     (0xFF, 0xFF, 0x00),
    "cyan":       (0x00, 0xFF, 0xFF),
    "green":      (0x00, 0xFF, 0x00),
    "darkYellow": (0xFF, 0xCC, 0x00),
}

def set_highlight(run, color_name: str):
    rPr = run._r.get_or_add_rPr()
    existing = rPr.find(qn('w:highlight'))
    if existing is not None:
        rPr.remove(existing)
    hl = OxmlElement('w:highlight')
    hl.set(qn('w:val'), color_name)
    rPr.append(hl)


def add_annotation(para, text: str):
    run = para.add_run(f"  [{text}]")
    run.italic = True
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    run.font.size = Pt(8)


def highlight_flagged_words(para, flagged: list):
    if not flagged:
        return
    flagged_lower = set(f.lower() for f in flagged)
    for run in para.runs:
        words_in_run = re.findall(r'\b\w+\b', run.text)
        if any(w.lower() in flagged_lower for w in words_in_run):
            set_highlight(run, "yellow")


def insert_heading_before(para, heading_text: str):
    """Insert a Heading 2 paragraph before the given paragraph."""
    new_p = OxmlElement('w:p')

    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'Heading2')
    pPr.append(pStyle)
    new_p.append(pPr)

    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    hl = OxmlElement('w:highlight')
    hl.set(qn('w:val'), 'darkYellow')
    rPr.append(hl)
    r.append(rPr)

    t = OxmlElement('w:t')
    t.text = heading_text
    r.append(t)
    new_p.append(r)

    para._element.addprevious(new_p)


# ── CHAPTER PARSING ───────────────────────────────────────────────────────────
def get_chapters(doc: Document) -> list:
    chapters, current = [], {"heading": None, "heading_para": None, "paragraphs": []}
    for para in doc.paragraphs:
        sname = para.style.name if para.style else ""
        if sname.startswith("Heading 1"):
            if current["heading"] is not None:
                chapters.append(current)
            current = {"heading": para.text, "heading_para": para, "paragraphs": []}
        elif current["heading"] is not None:
            current["paragraphs"].append(para)
    if current["heading"] is not None:
        chapters.append(current)
    return chapters


# ── MAIN ──────────────────────────────────────────────────────────────────────
def process(input_path: str, output_path: str):
    print(f"\n📄 Loading: {input_path}")
    doc = Document(input_path)
    chapters = get_chapters(doc)

    if not chapters:
        print("⚠️  No Heading 1 chapters found.")
        return

    print(f"✅ {len(chapters)} chapters detected\n")

    for i, ch in enumerate(chapters):
        num = i + 1
        is_odd = num % 2 == 1
        body = "\n".join(p.text for p in ch["paragraphs"] if p.text.strip())

        print(f"🔍 Chapter {num}: {ch['heading'][:55]}...")

        # ── AI analysis (swap mock for real API when deploying) ──
        analysis = mock_analyze_chapter(body, num, is_odd)
        # analysis = call_openai_api(body, num, is_odd)  # ← production line

        flagged  = analysis.get("flagged_words", [])
        closing  = analysis.get("closing_position", "")
        emphasis = analysis.get("emphasis_sentence", "")
        bquote   = analysis.get("blockquote_sentence", "")

        print(f"   ⚠️  Flagged: {flagged}")
        if is_odd:
            print(f"   ✏️  Emphasis: {emphasis[:55]}...")
            print(f"   💬  Blockquote: {bquote[:55]}...")

        closing_done = False
        paras = ch["paragraphs"]

        for j, para in enumerate(paras):
            if not para.text.strip():
                continue

            # 1. Spelling/typo highlights
            highlight_flagged_words(para, flagged)

            # 2. Emphasis highlight (odd chapters)
            if is_odd and emphasis and emphasis[:20].lower() in para.text.lower():
                for run in para.runs:
                    set_highlight(run, "cyan")
                add_annotation(para, "✏️ EMPHASIS SUGGESTION")

            # 3. Block quote highlight (odd chapters)
            if is_odd and bquote and bquote[:20].lower() in para.text.lower():
                for run in para.runs:
                    set_highlight(run, "green")
                add_annotation(para, "💬 BLOCK QUOTE SUGGESTION")

            # 4. Closing heading — insert before last 2 paragraphs
            if not closing_done and j >= len(paras) - 2:
                insert_heading_before(para, CLOSING_HEADING_TEXT)
                closing_done = True
                print(f"   📌 Closing heading inserted")

        # Annotate chapter heading with summary
        if flagged:
            add_annotation(ch["heading_para"],
                f"⚠️ Review {len(flagged)} flagged word(s): {', '.join(flagged[:6])}")

    doc.save(output_path)
    print(f"\n💾 Saved: {output_path}")
    print("\nColour legend (open in Microsoft Word):")
    print("  🟡 Yellow     = potential spelling / typo issue")
    print("  🔵 Cyan       = emphasis sentence suggestion  (odd chapters)")
    print("  🟢 Green      = block quote suggestion        (odd chapters)")
    print("  🟠 DarkYellow = inserted closing heading")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python editorial_review_demo.py input.docx output_annotated.docx")
        sys.exit(1)
    process(sys.argv[1], sys.argv[2])
