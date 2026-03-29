from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename

import os
import re
from datetime import datetime
from copy import deepcopy

from dotenv import load_dotenv
from openai import OpenAI
from docx import Document
from docx.shared import Pt as DocxPt
from docx.shared import RGBColor as DocxRGBColor
from PyPDF2 import PdfReader

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR
    from pptx.dml.color import RGBColor as PptRGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False


# ---------------- CONFIG ----------------
app = Flask(__name__, static_folder="static", template_folder="templates")

UPLOAD_FOLDER = "uploads"
GENERATED_FOLDER = "generated"
ALLOWED_EXTENSIONS = {"txt", "docx", "pdf"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["GENERATED_FOLDER"] = GENERATED_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


# ---------------- BASIC HELPERS ----------------
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_text(filepath: str) -> str:
    """Extract readable text from txt, docx, or pdf."""
    ext = filepath.rsplit(".", 1)[-1].lower()

    try:
        if ext == "txt":
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()

        if ext == "docx":
            doc = Document(filepath)
            return "\n".join(p.text for p in doc.paragraphs)

        if ext == "pdf":
            reader = PdfReader(filepath)
            parts = []
            for page in reader.pages:
                parts.append(page.extract_text() or "")
            return "\n".join(parts)

        return ""

    except Exception as e:
        print(f"Error extracting text from {filepath}: {e}")
        return ""


def normalize_whitespace(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def parse_drug_brand_blocks(text: str):
    """
    Handles:

    A) TOP 200 format:
       Generic
       Brand®
       Clinical Use

    B) Class-based format:
       Class
       Generic
       Brand®
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    drugs = []
    i = 0

    while i < len(lines) - 1:
        if i + 2 < len(lines):
            l1, l2, l3 = lines[i], lines[i + 1], lines[i + 2]

            if ("®" in l2) and ("®" not in l1) and ("®" not in l3):
                drugs.append({"generic": l1, "brand": l2, "extra": l3})
                i += 3
                continue

            if ("®" in l3) and ("®" not in l1) and ("®" not in l2):
                drugs.append({"generic": l2, "brand": l3, "extra": l1})
                i += 3
                continue

        l1, l2 = lines[i], lines[i + 1]
        if "®" in l2:
            drugs.append({"generic": l1, "brand": l2, "extra": ""})
            i += 2
        else:
            i += 1

    return drugs


# ---------------- MCQ NORMALIZATION / PARSING ----------------
def build_mcq_normalization_prompt(raw_text: str) -> str:
    return f"""
You are an expert pharmacy exam editor.

Read the text below and extract ALL multiple-choice questions EXACTLY as written.

CRITICAL RULES:
- If a question includes a case scenario, vignette, or clinical stem BEFORE the question sentence,
  include the FULL vignette as part of the question text.
- Do NOT summarize.
- Do NOT shorten.
- Do NOT rewrite unless needed only to repair broken formatting.
- Preserve original wording as much as possible.

OUTPUT RULES:
- Plain text only
- No markdown
- No bullets
- No commentary
- Use EXACTLY this structure for every question:

1. Full question text ending with a question mark?

A) Option text
B) Option text
C) Option text
D) Option text
Ans: C
Tips: Short explanation sentence.

RULES:
- Number sequentially
- Each question must include A) B) C) D)
- Use exactly "Ans:" and "Tips:"
- Leave one blank line between questions
- If the source has no explanation, create a short helpful explanation
- Fix broken option labels if needed
- If source uses A. / B. / C. / D., convert to A) / B) / C) / D)

TEXT:
{raw_text}
""".strip()


def split_text_for_llm(text: str, max_chars: int = 14000, max_chunks: int = 12):
    """
    Split large exam files into chunks while trying to preserve question boundaries.
    """
    text = normalize_whitespace(text)

    # Prefer splitting on numbered questions
    parts = re.split(r"(?=\n?\s*\d+\.\s)", "\n" + text)
    parts = [p.strip() for p in parts if p.strip()]

    if len(parts) <= 1:
        parts = re.split(r"\n\s*\n", text)
        parts = [p.strip() for p in parts if p.strip()]

    chunks = []
    current = ""

    for part in parts:
        candidate = f"{current}\n\n{part}".strip() if current else part
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                chunks.append(current)
            current = part

    if current:
        chunks.append(current)

    if not chunks:
        chunks = [text[:max_chars]]

    return chunks[:max_chunks]


def normalize_mcq_output_text(text: str) -> str:
    """
    Clean model output into one strict format.
    """
    text = text.replace("```", "")
    text = normalize_whitespace(text)

    lines = []
    for raw_line in text.split("\n"):
        line = raw_line.strip()

        line = re.sub(r"^([A-Da-d])\.\s+", r"\1) ", line)
        line = re.sub(r"^([A-Da-d])\)\s*", lambda m: f"{m.group(1).upper()}) ", line)
        line = re.sub(r"^(Ans|Answer)\s*[:\-]?\s*([A-Da-d])\b", r"Ans: \2", line, flags=re.I)
        line = re.sub(r"^(Tips|Explanation)\s*[:\-]?\s*", "Tips: ", line, flags=re.I)

        lines.append(line)

    cleaned = "\n".join(lines)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def parse_normalized_mcq_text(normalized_text: str):
    """
    Parse strict normalized MCQ text into structured objects.
    """
    normalized_text = normalize_mcq_output_text(normalized_text)

    pattern = re.compile(
        r"""
        ^\s*(\d+)\.\s*(.+?)\n
        A\)\s*(.+?)\n
        B\)\s*(.+?)\n
        C\)\s*(.+?)\n
        D\)\s*(.+?)
        (?:\nAns:\s*([A-D]))?
        (?:\nTips:\s*(.+?))?
        (?=\n\s*\d+\.|\Z)
        """,
        re.S | re.M | re.X,
    )

    mcqs = []
    for match in pattern.finditer(normalized_text):
        question = match.group(2).strip()
        a_opt = match.group(3).strip()
        b_opt = match.group(4).strip()
        c_opt = match.group(5).strip()
        d_opt = match.group(6).strip()
        answer = (match.group(7) or "").strip().upper()
        explanation = (match.group(8) or "No explanation provided.").strip()

        mcqs.append(
            {
                "question": question,
                "options": [
                    f"A) {a_opt}",
                    f"B) {b_opt}",
                    f"C) {c_opt}",
                    f"D) {d_opt}",
                ],
                "answer": answer or "N/A",
                "explanation": explanation,
            }
        )

    return mcqs


def render_normalized_mcq_text(mcqs) -> str:
    """
    Render structured MCQs back into strict normalized text.
    """
    blocks = []

    for idx, mcq in enumerate(mcqs, start=1):
        options = mcq.get("options", [])
        while len(options) < 4:
            label = chr(65 + len(options))
            options.append(f"{label}) ")

        answer = (mcq.get("answer") or "N/A").strip().upper()
        explanation = (mcq.get("explanation") or "No explanation provided.").strip()

        block = "\n".join(
            [
                f"{idx}. {mcq.get('question', '').strip()}",
                options[0].strip(),
                options[1].strip(),
                options[2].strip(),
                options[3].strip(),
                f"Ans: {answer}",
                f"Tips: {explanation}",
            ]
        )
        blocks.append(block)

    return "\n\n".join(blocks).strip()


def regex_extract_mcqs_fallback(text: str):
    """
    Fallback extractor if model normalization fails.
    More robust than the old simple splitter.
    """
    text = normalize_whitespace(text)

    pattern = re.compile(
        r"""
        (?:
            ^|\n
        )
        \s*(\d+)\.\s*(.+?)
        (?=
            \n[A-Da-d][\)\.]
        )
        \n[Aa][\)\.]\s*(.+?)
        \n[Bb][\)\.]\s*(.+?)
        \n[Cc][\)\.]\s*(.+?)
        \n[Dd][\)\.]\s*(.+?)
        (?:
            \n(?:Ans|Answer)\s*[:\-]?\s*([A-Da-d])
        )?
        (?:
            \n(?:Tips|Explanation)\s*[:\-]?\s*(.+?)
        )?
        (?=\n\s*\d+\.|\Z)
        """,
        re.S | re.M | re.X,
    )

    mcqs = []
    for m in pattern.finditer(text):
        question = m.group(2).strip()
        options = [
            f"A) {m.group(3).strip()}",
            f"B) {m.group(4).strip()}",
            f"C) {m.group(5).strip()}",
            f"D) {m.group(6).strip()}",
        ]
        answer = (m.group(7) or "").upper()
        explanation = (m.group(8) or "No explanation provided.").strip()

        mcqs.append(
            {
                "question": question,
                "options": options,
                "answer": answer or "N/A",
                "explanation": explanation,
            }
        )

    return mcqs


def normalize_mcqs_with_gpt(text: str) -> str:
    """
    Use the same CBT-parser standard to normalize MCQs for all generators.
    """
    chunks = split_text_for_llm(text)
    all_mcqs = []

    for chunk in chunks:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You normalize messy exam content into strict MCQ format.",
                },
                {
                    "role": "user",
                    "content": build_mcq_normalization_prompt(chunk),
                },
            ],
            temperature=0,
        )

        chunk_text = normalize_mcq_output_text(response.choices[0].message.content or "")
        chunk_mcqs = parse_normalized_mcq_text(chunk_text)
        all_mcqs.extend(chunk_mcqs)

    return render_normalized_mcq_text(all_mcqs)


def extract_mcqs_with_cbt_standard(text: str):
    """
    One universal MCQ extraction path for:
    - CBT Parser export
    - MCQ Generator 1
    - MCQ Generator 2
    - MCQ Mobile
    - MCQ Grader
    - future MCQ generators
    """
    try:
        normalized_text = normalize_mcqs_with_gpt(text)
        mcqs = parse_normalized_mcq_text(normalized_text)

        if mcqs:
            return mcqs, normalized_text
    except Exception as e:
        print(f"GPT normalization failed, falling back to regex extractor: {e}")

    mcqs = regex_extract_mcqs_fallback(text)
    normalized_text = render_normalized_mcq_text(mcqs) if mcqs else ""
    return mcqs, normalized_text


# ---------------- DOCX OUTPUT HELPERS ----------------
def write_normalized_mcqs_to_docx(normalized_text: str, output_path: str):
    doc = Document()

    def style(run, bold=False, italic=False, color=None):
        run.font.name = "Times New Roman"
        run.font.size = DocxPt(14)
        run.bold = bold
        run.italic = italic
        if color:
            run.font.color.rgb = color

    for line in normalized_text.split("\n"):
        line = line.rstrip()
        p = doc.add_paragraph()
        p.paragraph_format.space_before = DocxPt(0)
        p.paragraph_format.space_after = DocxPt(0)

        m_q = re.match(r"^(\d+\.)(\s*)(.+)", line)
        m_opt = re.match(r"^([A-D]\))(\s*)(.+)", line)

        if m_q:
            style(p.add_run(m_q.group(1)), bold=True)
            p.add_run(m_q.group(2))
            style(p.add_run(m_q.group(3)))
            continue

        if m_opt:
            p.paragraph_format.left_indent = DocxPt(18)
            style(p.add_run(m_opt.group(1)), bold=True)
            p.add_run(m_opt.group(2))
            style(p.add_run(m_opt.group(3)))
            continue

        if line.startswith("Ans:"):
            style(p.add_run("Ans:"), bold=True, color=DocxRGBColor(200, 0, 0))
            p.add_run(" ")
            style(
                p.add_run(line.replace("Ans:", "").strip()),
                color=DocxRGBColor(200, 0, 0),
            )
            continue

        if line.startswith("Tips:"):
            style(p.add_run("Tips:"), bold=True)
            p.add_run(" ")
            style(p.add_run(line.replace("Tips:", "").strip()))
            continue

        style(p.add_run(line))

    doc.save(output_path)


def convert_text_to_html_via_gpt(text: str) -> str:
    prompt = f"""
You are a document formatter.

Convert the content below into clean, valid HTML.

STRICT RULES:
- Return HTML ONLY
- Use semantic tags: h1, h2, h3, p, ul, ol, li, table, tr, th, td where appropriate
- Preserve logical structure
- Do NOT add explanations
- Do NOT use markdown
- Do NOT wrap in <html> or <body> tags

CONTENT:
{text}
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "system",
                "content": "You convert documents into clean HTML for web platforms.",
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        temperature=0,
    )

    return (response.choices[0].message.content or "").strip()


# ---------------- PPT HELPERS ----------------
def set_text_frame_defaults(text_frame, margins=(20, 20, 18, 18), align=PP_ALIGN.LEFT):
    text_frame.word_wrap = True
    text_frame.margin_left = Pt(margins[0])
    text_frame.margin_right = Pt(margins[1])
    text_frame.margin_top = Pt(margins[2])
    text_frame.margin_bottom = Pt(margins[3])
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    if text_frame.paragraphs:
        text_frame.paragraphs[0].alignment = align


def question_font_size(question: str, large=34, medium=30, small=26, tiny=22):
    n = len(re.sub(r"\s+", " ", question).strip())
    if n > 520:
        return Pt(tiny)
    if n > 360:
        return Pt(small)
    if n > 220:
        return Pt(medium)
    return Pt(large)


def option_font_size(option_text: str, large=24, medium=21, small=18):
    n = len(option_text.strip())
    if n > 80:
        return Pt(small)
    if n > 50:
        return Pt(medium)
    return Pt(large)


def ensure_four_options(mcq: dict):
    options = list(mcq.get("options", []))
    while len(options) < 4:
        label = chr(65 + len(options))
        options.append(f"{label}) ")
    return options[:4]


def shrink_text_to_fit(text_frame, paragraph, min_size=22):
    while paragraph.font.size and paragraph.font.size.pt > min_size:
        if len(text_frame.text.split("\n")) <= 4:
            break
        paragraph.font.size = Pt(paragraph.font.size.pt - 2)


# ---------------- PPT TEMPLATES ----------------
def create_vba_template_presentation(mcqs, output_path):
    if not PPTX_AVAILABLE:
        return False

    try:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        BLUE = PptRGBColor(100, 149, 237)
        CUSTOMBG = PptRGBColor(249, 248, 242)
        DARK = PptRGBColor(22, 22, 26)
        WHITE = PptRGBColor(255, 255, 255)

        for i, mcq in enumerate(mcqs, 1):
            slide1 = prs.slides.add_slide(prs.slide_layouts[6])

            bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg1.fill.solid()
            bg1.fill.fore_color.rgb = CUSTOMBG
            bg1.line.fill.background()

            qBox1 = slide1.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.7), Inches(12.3), Inches(2.0)
            )
            qBox1.fill.solid()
            qBox1.fill.fore_color.rgb = CUSTOMBG
            qBox1.line.color.rgb = BLUE
            qBox1.line.width = Pt(8)
            qBox1.adjustments[0] = 0.08

            qText1 = qBox1.text_frame
            qText1.clear()
            qText1.word_wrap = True
            qText1.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            qText1.margin_left = Pt(25)
            qText1.margin_right = Pt(25)
            qText1.margin_top = Pt(20)
            qText1.margin_bottom = Pt(20)

            p = qText1.paragraphs[0]
            p.text = f"Q{i}: {mcq['question']}"
            p.font.name = "Arial"
            p.font.size = question_font_size(mcq["question"], large=36, medium=32, small=28, tiny=24)
            p.font.bold = True
            p.font.color.rgb = DARK

            badge1 = slide1.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(0.15), Inches(0.45), Inches(1.1), Inches(1.1)
            )
            badge1.fill.solid()
            badge1.fill.fore_color.rgb = BLUE
            badge1.line.fill.background()

            badgeText1 = badge1.text_frame
            badgeText1.text = str(i)
            badgeText1.paragraphs[0].font.name = "Arial"
            badgeText1.paragraphs[0].font.size = Pt(36)
            badgeText1.paragraphs[0].font.bold = True
            badgeText1.paragraphs[0].font.color.rgb = WHITE
            badgeText1.paragraphs[0].alignment = PP_ALIGN.CENTER

            option_top = Inches(3.0)
            options = ensure_four_options(mcq)

            for j, option in enumerate(options):
                option_box = slide1.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Inches(0.5),
                    option_top + (j * Inches(1.0)),
                    Inches(12.3),
                    Inches(0.85),
                )
                option_box.fill.solid()
                option_box.fill.fore_color.rgb = CUSTOMBG
                option_box.line.color.rgb = BLUE
                option_box.line.width = Pt(8)
                option_box.adjustments[0] = 0.08

                option_text = option_box.text_frame
                option_text.word_wrap = True
                option_text.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                option_text.margin_left = Pt(25)
                option_text.margin_right = Pt(25)
                option_text.margin_top = Pt(12)
                option_text.margin_bottom = Pt(12)
                option_text.text = option

                option_text.paragraphs[0].font.name = "Arial"
                option_text.paragraphs[0].font.size = option_font_size(option)
                option_text.paragraphs[0].font.italic = True
                option_text.paragraphs[0].font.color.rgb = DARK

            slide2 = prs.slides.add_slide(prs.slide_layouts[6])

            bg2 = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg2.fill.solid()
            bg2.fill.fore_color.rgb = CUSTOMBG
            bg2.line.fill.background()

            qBox2 = slide2.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.7), Inches(12.3), Inches(2.0)
            )
            qBox2.fill.solid()
            qBox2.fill.fore_color.rgb = CUSTOMBG
            qBox2.line.color.rgb = BLUE
            qBox2.line.width = Pt(8)
            qBox2.adjustments[0] = 0.08

            qText2 = qBox2.text_frame
            qText2.clear()
            qText2.word_wrap = True
            qText2.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            qText2.margin_left = Pt(25)
            qText2.margin_right = Pt(25)
            qText2.margin_top = Pt(20)
            qText2.margin_bottom = Pt(20)

            p = qText2.paragraphs[0]
            p.text = f"Q{i}: {mcq['question']}"
            p.font.name = "Arial"
            p.font.size = question_font_size(mcq["question"], large=36, medium=32, small=28, tiny=24)
            p.font.bold = True
            p.font.color.rgb = DARK

            badge2 = slide2.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(0.15), Inches(0.45), Inches(1.1), Inches(1.1)
            )
            badge2.fill.solid()
            badge2.fill.fore_color.rgb = BLUE
            badge2.line.fill.background()

            badgeText2 = badge2.text_frame
            badgeText2.text = str(i)
            badgeText2.paragraphs[0].font.name = "Arial"
            badgeText2.paragraphs[0].font.size = Pt(36)
            badgeText2.paragraphs[0].font.bold = True
            badgeText2.paragraphs[0].font.color.rgb = WHITE
            badgeText2.paragraphs[0].alignment = PP_ALIGN.CENTER

            answer_box = slide2.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(0.5),
                Inches(3.0),
                Inches(12.3),
                Inches(3.2),
            )
            answer_box.fill.solid()
            answer_box.fill.fore_color.rgb = CUSTOMBG
            answer_box.line.color.rgb = BLUE
            answer_box.line.width = Pt(8)
            answer_box.adjustments[0] = 0.08

            answer_content = []
            if mcq["answer"]:
                answer_content.append(f"Correct Answer: {mcq['answer']}")
            if mcq["explanation"]:
                answer_content.append(f"Explanation: {mcq['explanation']}")

            answer_text = answer_box.text_frame
            answer_text.text = "\n\n".join(answer_content)
            answer_text.word_wrap = True
            answer_text.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            answer_text.margin_left = Pt(30)
            answer_text.margin_right = Pt(30)
            answer_text.margin_top = Pt(25)
            answer_text.margin_bottom = Pt(25)

            for paragraph in answer_text.paragraphs:
                paragraph.font.name = "Arial"
                paragraph.font.size = Pt(28)
                paragraph.font.color.rgb = DARK
                paragraph.alignment = PP_ALIGN.LEFT

            if answer_text.paragraphs:
                answer_text.paragraphs[0].font.bold = True
                answer_text.paragraphs[0].font.italic = True

        prs.save(output_path)
        return True

    except Exception as e:
        print(f"Error creating VBA template presentation: {e}")
        return False


def create_mcq_generator2_exact(mcqs, output_path):
    """
    MCQ Generator 2
    - balanced blue question slide
    - clean red answer slide
    - adaptive text sizing
    - better spacing and proportions
    """
    if not PPTX_AVAILABLE:
        return False

    try:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        BG = PptRGBColor(239, 243, 248)
        RIGHT_BG = PptRGBColor(240, 243, 246)
        BLUE = PptRGBColor(45, 105, 220)
        RED = PptRGBColor(210, 35, 35)
        WHITE = PptRGBColor(255, 255, 255)
        DARK = PptRGBColor(24, 28, 35)
        CARD_FILL = PptRGBColor(255, 255, 255)
        CARD_LINE = PptRGBColor(200, 207, 216)

        left_panel_width = Inches(6.35)
        right_start = left_panel_width
        right_card_x = Inches(7.35)
        right_card_w = Inches(5.15)

        def fit_question_size(text):
            n = len((text or "").strip())
            if n > 500:
                return Pt(18)
            if n > 350:
                return Pt(22)
            if n > 220:
                return Pt(25)
            return Pt(29)

        def fit_option_size(text):
            n = len((text or "").strip())
            if n > 90:
                return Pt(15)
            if n > 65:
                return Pt(17)
            return Pt(20)

        def fit_explanation_size(text):
            n = len((text or "").strip())
            if n > 700:
                return Pt(15)
            if n > 500:
                return Pt(17)
            if n > 320:
                return Pt(19)
            return Pt(21)

        for i, mcq in enumerate(mcqs, 1):
            options = ensure_four_options(mcq)

            # =========================
            # SLIDE 1 — QUESTION
            # =========================
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = BG
            bg.line.fill.background()

            left_panel = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                0,
                0,
                left_panel_width,
                prs.slide_height,
            )
            left_panel.fill.solid()
            left_panel.fill.fore_color.rgb = BLUE
            left_panel.line.fill.background()

            right_panel = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                right_start,
                0,
                prs.slide_width - left_panel_width,
                prs.slide_height,
            )
            right_panel.fill.solid()
            right_panel.fill.fore_color.rgb = RIGHT_BG
            right_panel.line.fill.background()

            num_circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(0.55),
                Inches(0.50),
                Inches(0.78),
                Inches(0.78),
            )
            num_circle.fill.solid()
            num_circle.fill.fore_color.rgb = WHITE
            num_circle.line.fill.background()

            num_tf = num_circle.text_frame
            num_tf.clear()
            num_tf.text = str(i)
            num_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            num_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            num_tf.paragraphs[0].font.name = "Arial"
            num_tf.paragraphs[0].font.size = Pt(24)
            num_tf.paragraphs[0].font.bold = True
            num_tf.paragraphs[0].font.color.rgb = BLUE

            q_box = slide.shapes.add_textbox(
                Inches(0.78),
                Inches(1.00),
                Inches(5.00),
                Inches(5.95),
            )
            q_tf = q_box.text_frame
            q_tf.clear()
            q_tf.word_wrap = True
            q_tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            q_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            q_tf.margin_left = Pt(20)
            q_tf.margin_right = Pt(20)
            q_tf.margin_top = Pt(20)
            q_tf.margin_bottom = Pt(20)

            q_p = q_tf.paragraphs[0]
            q_p.text = f"{i}. {mcq['question']}"
            q_p.alignment = PP_ALIGN.LEFT
            q_p.font.name = "Arial"
            q_p.font.size = fit_question_size(mcq["question"])
            q_p.font.bold = True
            q_p.font.color.rgb = WHITE

            option_y_positions = [Inches(0.90), Inches(2.08), Inches(3.26), Inches(4.44)]

            for idx, option in enumerate(options):
                opt_box = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    right_card_x,
                    option_y_positions[idx],
                    right_card_w,
                    Inches(0.92),
                )
                opt_box.fill.solid()
                opt_box.fill.fore_color.rgb = CARD_FILL
                opt_box.line.color.rgb = CARD_LINE
                opt_box.line.width = Pt(1.3)
                opt_box.adjustments[0] = 0.06

                opt_tf = opt_box.text_frame
                opt_tf.clear()
                opt_tf.word_wrap = True
                opt_tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                opt_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                opt_tf.margin_left = Pt(18)
                opt_tf.margin_right = Pt(18)
                opt_tf.margin_top = Pt(8)
                opt_tf.margin_bottom = Pt(8)

                opt_p = opt_tf.paragraphs[0]
                opt_p.text = option
                opt_p.alignment = PP_ALIGN.CENTER
                opt_p.font.name = "Arial"
                opt_p.font.size = fit_option_size(option)
                opt_p.font.bold = True
                opt_p.font.color.rgb = DARK

            # =========================
            # SLIDE 2 — ANSWER
            # =========================
            slide2 = prs.slides.add_slide(prs.slide_layouts[6])

            bg2 = slide2.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
            )
            bg2.fill.solid()
            bg2.fill.fore_color.rgb = BG
            bg2.line.fill.background()

            correct_letter = (mcq.get("answer") or "N/A").strip().upper()
            correct_text = next(
                (o[3:].strip() for o in options if o.startswith(correct_letter)),
                "",
            )

            answer_box = slide2.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.75),
                Inches(0.72),
                Inches(7.20),
                Inches(1.35),
            )
            answer_box.fill.solid()
            answer_box.fill.fore_color.rgb = RED
            answer_box.line.fill.background()

            ans_tf = answer_box.text_frame
            ans_tf.clear()
            ans_tf.word_wrap = True
            ans_tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            ans_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            ans_tf.margin_left = Pt(18)
            ans_tf.margin_right = Pt(18)
            ans_tf.margin_top = Pt(10)
            ans_tf.margin_bottom = Pt(10)

            ans_p = ans_tf.paragraphs[0]
            ans_p.text = f"Answer: {correct_letter}) {correct_text}"
            ans_p.alignment = PP_ALIGN.LEFT
            ans_p.font.name = "Arial"
            ans_p.font.size = Pt(24)
            ans_p.font.bold = True
            ans_p.font.color.rgb = WHITE

            exp_box = slide2.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(0.75),
                Inches(2.40),
                Inches(11.85),
                Inches(2.45),
            )
            exp_box.fill.solid()
            exp_box.fill.fore_color.rgb = WHITE
            exp_box.line.color.rgb = CARD_LINE
            exp_box.line.width = Pt(1.3)
            exp_box.adjustments[0] = 0.06

            exp_tf = exp_box.text_frame
            exp_tf.clear()
            exp_tf.word_wrap = True
            exp_tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            exp_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
            exp_tf.margin_left = Pt(24)
            exp_tf.margin_right = Pt(24)
            exp_tf.margin_top = Pt(18)
            exp_tf.margin_bottom = Pt(18)

            exp_title = exp_tf.paragraphs[0]
            exp_title.text = "Explanation:"
            exp_title.alignment = PP_ALIGN.LEFT
            exp_title.font.name = "Arial"
            exp_title.font.size = Pt(20)
            exp_title.font.bold = True
            exp_title.font.color.rgb = DARK

            exp_body = exp_tf.add_paragraph()
            exp_body.text = mcq.get("explanation") or "No explanation provided."
            exp_body.alignment = PP_ALIGN.LEFT
            exp_body.font.name = "Arial"
            exp_body.font.size = fit_explanation_size(mcq.get("explanation") or "")
            exp_body.font.color.rgb = DARK

        prs.save(output_path)
        return True

    except Exception as e:
        print(f"MCQ2 ERROR: {e}")
        return False


def create_vba_template_presentation_mobile(mcqs, output_path):
    if not PPTX_AVAILABLE:
        return False

    try:
        prs = Presentation()
        prs.slide_width = Inches(7.5)
        prs.slide_height = Inches(13.33)

        BLUE = PptRGBColor(100, 149, 237)
        BG = PptRGBColor(249, 248, 242)
        DARK = PptRGBColor(22, 22, 26)
        WHITE = PptRGBColor(255, 255, 255)

        LEFT_MARGIN = Inches(0.6)
        BOX_WIDTH = prs.slide_width - Inches(1.2)

        for i, mcq in enumerate(mcqs, start=1):
            options = ensure_four_options(mcq)

            slide = prs.slides.add_slide(prs.slide_layouts[6])

            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = BG
            bg.line.fill.background()

            qbox_top = Inches(1.8)
            qbox_height = Inches(3.2)

            qbox = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                LEFT_MARGIN,
                qbox_top,
                BOX_WIDTH,
                qbox_height,
            )
            qbox.fill.solid()
            qbox.fill.fore_color.rgb = BG
            qbox.line.color.rgb = BLUE
            qbox.line.width = Pt(6)
            qbox.adjustments[0] = 0.08

            qtf = qbox.text_frame
            qtf.word_wrap = True
            qtf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            qtf.margin_left = Inches(0.35)
            qtf.margin_right = Inches(0.35)
            qtf.margin_top = Inches(0.35)
            qtf.margin_bottom = Inches(0.35)

            qp = qtf.paragraphs[0]
            qp.text = f"Q{i}: {mcq['question']}"
            qp.font.name = "Arial"
            qp.font.size = question_font_size(mcq["question"], large=32, medium=28, small=24, tiny=20)
            qp.font.bold = True
            qp.font.color.rgb = DARK
            qp.alignment = PP_ALIGN.CENTER

            badge = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                LEFT_MARGIN - Inches(0.5),
                qbox_top - Inches(0.4),
                Inches(1.2),
                Inches(1.2),
            )
            badge.fill.solid()
            badge.fill.fore_color.rgb = BLUE
            badge.line.fill.background()

            bt = badge.text_frame
            bt.clear()
            bp = bt.paragraphs[0]
            bp.text = str(i)
            bp.font.name = "Arial"
            bp.font.size = Pt(34)
            bp.font.bold = True
            bp.font.color.rgb = WHITE
            bp.alignment = PP_ALIGN.CENTER

            start_y = qbox_top + qbox_height + Inches(0.8)
            option_height = Inches(1.05)
            option_spacing = Inches(1.45)

            for opt in options:
                obox = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    LEFT_MARGIN,
                    start_y,
                    BOX_WIDTH,
                    option_height,
                )
                obox.fill.solid()
                obox.fill.fore_color.rgb = BG
                obox.line.color.rgb = BLUE
                obox.line.width = Pt(6)
                obox.adjustments[0] = 0.08

                ot = obox.text_frame
                ot.word_wrap = True
                ot.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                ot.margin_left = Inches(0.3)
                ot.margin_right = Inches(0.3)

                op = ot.paragraphs[0]
                op.text = opt
                op.font.name = "Arial"
                op.font.size = option_font_size(opt, large=26, medium=22, small=18)
                op.font.italic = True
                op.font.color.rgb = DARK
                op.alignment = PP_ALIGN.CENTER

                start_y += option_spacing

            slide2 = prs.slides.add_slide(prs.slide_layouts[6])

            bg2 = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg2.fill.solid()
            bg2.fill.fore_color.rgb = BG
            bg2.line.fill.background()

            qbox2_top = Inches(1.8)

            qbox2 = slide2.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                LEFT_MARGIN,
                qbox2_top,
                BOX_WIDTH,
                Inches(2.6),
            )
            qbox2.fill.solid()
            qbox2.fill.fore_color.rgb = BG
            qbox2.line.color.rgb = BLUE
            qbox2.line.width = Pt(6)
            qbox2.adjustments[0] = 0.08

            qtf2 = qbox2.text_frame
            qtf2.word_wrap = True
            qtf2.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            qtf2.margin_left = Inches(0.35)
            qtf2.margin_right = Inches(0.35)

            qtf2.paragraphs[0].text = mcq["question"]
            qtf2.paragraphs[0].font.size = question_font_size(mcq["question"], large=28, medium=24, small=20, tiny=18)
            qtf2.paragraphs[0].font.bold = True
            qtf2.paragraphs[0].font.color.rgb = DARK
            qtf2.paragraphs[0].alignment = PP_ALIGN.CENTER

            badge2 = slide2.shapes.add_shape(
                MSO_SHAPE.OVAL,
                LEFT_MARGIN - Inches(0.5),
                qbox2_top - Inches(0.4),
                Inches(1.2),
                Inches(1.2),
            )
            badge2.fill.solid()
            badge2.fill.fore_color.rgb = BLUE
            badge2.line.fill.background()

            bt2 = badge2.text_frame
            bt2.clear()
            bp2 = bt2.paragraphs[0]
            bp2.text = str(i)
            bp2.font.size = Pt(34)
            bp2.font.bold = True
            bp2.font.color.rgb = WHITE
            bp2.alignment = PP_ALIGN.CENTER

            abox = slide2.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                LEFT_MARGIN,
                qbox2_top + Inches(3.6),
                BOX_WIDTH,
                Inches(4.2),
            )
            abox.fill.solid()
            abox.fill.fore_color.rgb = BG
            abox.line.color.rgb = BLUE
            abox.line.width = Pt(6)
            abox.adjustments[0] = 0.08

            atf = abox.text_frame
            atf.word_wrap = True
            atf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            atf.margin_left = Inches(0.35)
            atf.margin_right = Inches(0.35)
            atf.margin_top = Inches(0.35)

            p1 = atf.paragraphs[0]
            p1.text = f"Correct Answer: {mcq['answer']}"
            p1.font.size = Pt(30)
            p1.font.bold = True
            p1.font.color.rgb = DARK
            p1.alignment = PP_ALIGN.CENTER

            p2 = atf.add_paragraph()
            p2.text = mcq["explanation"]
            p2.font.size = question_font_size(mcq["explanation"], large=24, medium=22, small=20, tiny=18)
            p2.font.color.rgb = DARK
            p2.alignment = PP_ALIGN.CENTER

        prs.save(output_path)
        return True

    except Exception as e:
        print(f"Error creating mobile MCQ PPT: {e}")
        return False


def create_brand_template_presentation_from_ppt(drug_blocks, output_path):
    if not PPTX_AVAILABLE:
        return False

    try:
        template_path = "drug output template.pptx"
        prs = Presentation(template_path)
        template_slide = prs.slides[0]

        sldIdLst = list(prs.slides._sldIdLst)
        for sld in sldIdLst:
            rId = sld.rId
            prs.part.drop_rel(rId)
            prs.slides._sldIdLst.remove(sld)

        def clone_template_slide():
            slide_layout = template_slide.slide_layout
            new_slide = prs.slides.add_slide(slide_layout)

            for shp in list(new_slide.shapes):
                el = shp._element
                el.getparent().remove(el)

            for shp in template_slide.shapes:
                new_el = deepcopy(shp._element)
                new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

            return new_slide

        for idx, drug in enumerate(drug_blocks, start=1):
            slide = clone_template_slide()
            shapes = slide.shapes

            generic = drug.get("generic", "")
            brand = drug.get("brand", "")
            header = drug.get("header", "")

            if len(shapes) >= 1:
                shapes[0].text = generic
            if len(shapes) >= 2:
                shapes[1].text = brand
            if len(shapes) >= 3:
                shapes[2].text = header
            if len(shapes) >= 4:
                shapes[3].text = str(idx)

        prs.save(output_path)
        return True

    except Exception as e:
        print(f"Error creating brand template PPT from PPT: {e}")
        return False


def create_brand_template_presentation(drug_blocks, output_path):
    if not PPTX_AVAILABLE:
        return False

    try:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        BG = PptRGBColor(249, 244, 230)
        BLUE = PptRGBColor(51, 102, 204)
        YELLOW = PptRGBColor(255, 204, 0)
        BLACK = PptRGBColor(0, 0, 0)
        WHITE = PptRGBColor(255, 255, 255)
        GREY = PptRGBColor(120, 120, 120)

        for idx, drug in enumerate(drug_blocks, start=1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            generic = drug.get("generic", "")
            brand = drug.get("brand", "")
            extra = drug.get("extra", "")

            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = BG
            bg.line.fill.background()

            blue_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.5),
                Inches(1.3),
                Inches(12.3),
                Inches(1.2),
            )
            blue_box.fill.solid()
            blue_box.fill.fore_color.rgb = BLUE
            blue_box.line.fill.background()

            gtf = blue_box.text_frame
            gtf.word_wrap = True
            gtf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            p = gtf.paragraphs[0]
            p.text = generic
            p.font.name = "Arial"
            p.font.size = Pt(44)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

            brand_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.0), Inches(12.3), Inches(1.0))
            btf = brand_box.text_frame
            btf.word_wrap = True
            btf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            pb = btf.paragraphs[0]
            pb.text = brand
            pb.font.name = "Arial"
            pb.font.size = Pt(40)
            pb.font.bold = True
            pb.font.color.rgb = BLACK
            pb.alignment = PP_ALIGN.CENTER

            yellow_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.5),
                Inches(4.4),
                Inches(12.3),
                Inches(1.6),
            )
            yellow_box.fill.solid()
            yellow_box.fill.fore_color.rgb = YELLOW
            yellow_box.line.fill.background()

            etf = yellow_box.text_frame
            etf.word_wrap = True
            etf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            etf.margin_left = Pt(20)
            etf.margin_right = Pt(20)
            pe = etf.paragraphs[0]
            pe.text = extra
            pe.font.name = "Arial"
            pe.font.size = Pt(32)
            pe.font.color.rgb = BLACK
            pe.alignment = PP_ALIGN.CENTER

            num = slide.shapes.add_textbox(
                prs.slide_width - Inches(1.0),
                prs.slide_height - Inches(0.9),
                Inches(0.8),
                Inches(0.8),
            )
            ntf = num.text_frame
            pn = ntf.paragraphs[0]
            pn.text = str(idx)
            pn.font.name = "Arial"
            pn.font.size = Pt(28)
            pn.font.color.rgb = GREY
            pn.alignment = PP_ALIGN.RIGHT

        prs.save(output_path)
        return True

    except Exception as e:
        print(f"Error in create_brand_template_presentation: {e}")
        return False


def create_brand_template_presentation_mobile(drug_blocks, output_path):
    if not PPTX_AVAILABLE:
        return False

    try:
        prs = Presentation()
        prs.slide_width = Inches(7.5)
        prs.slide_height = Inches(13.33)

        BG = PptRGBColor(249, 244, 230)
        BLUE = PptRGBColor(51, 102, 204)
        YELLOW = PptRGBColor(255, 204, 0)
        BLACK = PptRGBColor(0, 0, 0)
        WHITE = PptRGBColor(255, 255, 255)

        for idx, drug in enumerate(drug_blocks, start=1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            generic = drug.get("generic", "")
            brand = drug.get("brand", "")
            extra = drug.get("extra", "")

            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = BG
            bg.line.fill.background()

            generic_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.5),
                Inches(1.0),
                prs.slide_width - Inches(1.0),
                Inches(1.8),
            )
            generic_box.fill.solid()
            generic_box.fill.fore_color.rgb = BLUE
            generic_box.line.fill.background()

            tf = generic_box.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            p = tf.paragraphs[0]
            p.text = generic
            p.font.name = "Arial"
            p.font.size = Pt(42)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

            brand_box = slide.shapes.add_textbox(
                Inches(0.5),
                Inches(3.4),
                prs.slide_width - Inches(1.0),
                Inches(2.0),
            )
            btf = brand_box.text_frame
            btf.word_wrap = True
            btf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            bp = btf.paragraphs[0]
            bp.text = brand
            bp.font.name = "Arial"
            bp.font.size = Pt(40)
            bp.font.bold = True
            bp.font.color.rgb = BLACK
            bp.alignment = PP_ALIGN.CENTER

            extra_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.5),
                prs.slide_height - Inches(3.0),
                prs.slide_width - Inches(1.0),
                Inches(1.8),
            )
            extra_box.fill.solid()
            extra_box.fill.fore_color.rgb = YELLOW
            extra_box.line.fill.background()

            etf = extra_box.text_frame
            etf.word_wrap = True
            etf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            ep = etf.paragraphs[0]
            ep.text = extra
            ep.font.name = "Arial"
            ep.font.size = Pt(30)
            ep.font.color.rgb = BLACK
            ep.alignment = PP_ALIGN.CENTER

        prs.save(output_path)
        return True

    except Exception as e:
        print(f"Error creating mobile brand template: {e}")
        return False


def create_ppt_template_presentation(mcqs, output_path):
    if not PPTX_AVAILABLE:
        return False

    try:
        prs = Presentation("templates/ppt_template.pptx")

        layout = None
        for slide_layout in prs.slide_layouts:
            if slide_layout.placeholders:
                layout = slide_layout
                break

        if not layout:
            layout = prs.slide_layouts[0]

        for _ in range(len(prs.slides)):
            rId = prs.slides._sldIdLst[0].rId
            prs.part.drop_rel(rId)

        for i, mcq in enumerate(mcqs, 1):
            slide = prs.slides.add_slide(layout)

            if slide.shapes.title:
                slide.shapes.title.text = f"Q{i}: {mcq['question']}"

            content_shape = None
            for shape in slide.placeholders:
                if shape.placeholder_format.type == 7:
                    content_shape = shape
                    break

            if content_shape:
                tf = content_shape.text_frame
                tf.clear()
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

                for option in ensure_four_options(mcq):
                    p = tf.add_paragraph()
                    p.text = option
                    p.font.size = Pt(18)

                if mcq["answer"] or mcq["explanation"]:
                    p = tf.add_paragraph()
                    answer_text = f"Answer: {mcq['answer']}" if mcq["answer"] else ""
                    explanation_text = (
                        f"Explanation: {mcq['explanation']}" if mcq["explanation"] else ""
                    )
                    p.text = f"{answer_text}\n{explanation_text}".strip()
                    p.font.bold = True
                    p.font.size = Pt(16)

        prs.save(output_path)
        return True

    except Exception as e:
        print(f"Error creating PPT template presentation: {e}")
        return False


# ---------------- ROUTES ----------------
@app.route("/")
def home():
    return send_from_directory(".", "index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(filepath)
        return jsonify({"message": "✅ File uploaded successfully", "filename": filename})

    return jsonify({"error": "Invalid file type"}), 400


@app.route("/generate", methods=["POST"])
def generate_output():
    try:
        data = request.get_json(silent=True) or request.form

        filename = data.get("filename")
        template_choice = data.get("template")
        cbt_topic = (data.get("cbt_topic") or "").strip()

        # ---------- CBT GENERATOR ----------
        if template_choice == "cbt":
            if not cbt_topic:
                return jsonify({"error": "Please enter a topic for CBT Generator"}), 400

            prompt = f"""
Generate 25 MCQs on the topic: "{cbt_topic}"

Format MUST match EXACTLY this structure:

**1.** *Question text here…*

a. Option A
b. Option B
c. Option C
d. Option D

**Answer: c**
**Explanation:** *Short explanation here.*

Rules:
- Same line breaks
- Numbered 1–25
- Keep formatting consistent
"""

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": "You generate perfectly formatted medical MCQs.",
                    },
                    {
                        "role": "user",
                        "content": prompt,
                    },
                ],
                temperature=0,
            )

            content = response.choices[0].message.content or ""

            output_filename = f"CBT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            output_path = os.path.join(app.config["GENERATED_FOLDER"], output_filename)

            doc = Document()

            for line in content.split("\n"):
                line = line.rstrip()

                if line.startswith("**") and line.endswith("**"):
                    p = doc.add_paragraph()
                    r = p.add_run(line.replace("**", ""))
                    r.bold = True
                    continue

                if line.startswith("**Answer"):
                    p = doc.add_paragraph()
                    r = p.add_run(line.replace("**", ""))
                    r.bold = True
                    continue

                if line.startswith("**Explanation"):
                    p = doc.add_paragraph()
                    r = p.add_run(line.replace("**", ""))
                    r.bold = True
                    r.italic = True
                    continue

                if line.startswith("*") and line.endswith("*"):
                    p = doc.add_paragraph()
                    r = p.add_run(line.replace("*", ""))
                    r.italic = True
                    continue

                p = doc.add_paragraph()
                p.add_run(line)

            doc.save(output_path)

            return jsonify(
                {
                    "message": "✅ CBT Word file generated!",
                    "download_url": f"/download/{output_filename}",
                }
            )

        # ---------- ALL FILE-BASED MODES ----------
        if not filename or not template_choice:
            return jsonify({"error": "Missing parameters"}), 400

        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        if not os.path.exists(filepath):
            return jsonify({"error": f"Uploaded file not found: {filepath}"}), 404

        file_content = extract_text(filepath)
        if not file_content.strip():
            return jsonify({"error": "File is empty"}), 400

        # ---------- DOCX TO HTML ----------
        if template_choice == "docx_to_html":
            html_output = convert_text_to_html_via_gpt(file_content)

            output_filename = f"converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
            output_path = os.path.join(app.config["GENERATED_FOLDER"], output_filename)

            with open(output_path, "w", encoding="utf-8") as f:
                f.write(html_output)

            return jsonify(
                {
                    "message": "✅ HTML generated successfully",
                    "download_url": f"/download/{output_filename}",
                }
            )

        # ---------- BRAND / DRUG TEMPLATES ----------
        if template_choice in {"brand_template", "brand_template_mobile"}:
            drug_blocks = parse_drug_brand_blocks(file_content)

            if not drug_blocks:
                return jsonify({"error": "No brand/generic blocks were found in this file."}), 400

            output_filename = f"generated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            output_path = os.path.join(app.config["GENERATED_FOLDER"], output_filename)

            if template_choice == "brand_template":
                success = create_brand_template_presentation(drug_blocks, output_path)
            else:
                success = create_brand_template_presentation_mobile(drug_blocks, output_path)

            if success:
                return jsonify(
                    {
                        "message": "File generated successfully",
                        "download_url": f"/download/{output_filename}",
                    }
                )

            return jsonify({"error": "Failed to generate file"}), 500

        # ---------- UNIFIED MCQ EXTRACTION FOR ALL MCQ TOOLS ----------
        mcqs, normalized_mcq_text = extract_mcqs_with_cbt_standard(file_content)

        if not mcqs:
            return jsonify({"error": "No MCQs could be extracted from this file."}), 400

        # ---------- CBT PARSER EXPORT ----------
        if template_choice == "cbt_parser":
            output_filename = f"CBT_PARSED_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            output_path = os.path.join(app.config["GENERATED_FOLDER"], output_filename)
            write_normalized_mcqs_to_docx(normalized_mcq_text, output_path)

            return jsonify(
                {
                    "message": "✅ File parsed and formatted successfully",
                    "download_url": f"/download/{output_filename}",
                }
            )

        # ---------- MCQ GRADER ----------
        if template_choice == "mcq_grader":
            grader_mcqs = []
            for mcq in mcqs:
                ans = (mcq.get("answer") or "").upper()
                if ans in {"A", "B", "C", "D"}:
                    grader_mcqs.append(
                        {
                            "question": mcq["question"],
                            "answer": ans,
                        }
                    )

            if not grader_mcqs:
                return jsonify(
                    {"error": "No MCQs with clear A–D answers were found in this file."}
                ), 400

            return jsonify(
                {
                    "message": "MCQs loaded for grading.",
                    "mcqs": grader_mcqs,
                }
            )

        # ---------- PPT / VBA GENERATION ----------
        output_filename = f"generated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        output_path = os.path.join(app.config["GENERATED_FOLDER"], output_filename)

        if template_choice == "vba":
            success = create_vba_template_presentation(mcqs, output_path)
        elif template_choice == "mcq2":
            success = create_mcq_generator2_exact(mcqs, output_path)
        elif template_choice == "vba_mobile":
            success = create_vba_template_presentation_mobile(mcqs, output_path)
        elif template_choice == "ppt":
            success = create_ppt_template_presentation(mcqs, output_path)
        else:
            return jsonify({"error": "Invalid template type"}), 400

        if success:
            return jsonify(
                {
                    "message": "File generated successfully",
                    "download_url": f"/download/{output_filename}",
                }
            )

        return jsonify({"error": "Failed to generate file"}), 500

    except Exception as e:
        print(f"/generate error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(app.config["GENERATED_FOLDER"], filename, as_attachment=True)


# ---------------- RUN APP ----------------
if __name__ == "__main__":
    if not PPTX_AVAILABLE:
        print("WARNING: python-pptx is not installed. Please run: pip install python-pptx")
    app.run(debug=True)