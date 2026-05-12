"""Build embeddable index.html from outline.json + docx media."""
from __future__ import annotations

import html
import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]  # gce-level-1-embed
SITE_ROOT = ROOT.parent
DOCX = SITE_ROOT / "gce level 1.docx"
OUTLINE = Path(__file__).resolve().parent / "outline.json"
OUT_HTML = ROOT / "index.html"
MEDIA_DIR = ROOT / "media"

# Anchor id -> block index (first paragraph / figure at or after this step)
NAV: list[tuple[str, str, int | None, list[tuple[str, str, int]] | None]] = [
    ("cara-masuk", "Cara masuk Lab Drive", 0, None),
    ("bahagian-kuiz", "Bahagian Kuiz", 41, None),
    (
        "lab-1",
        "Lab 1 · Google Classroom",
        53,
        [
            ("lab-1-login", "Log in GC", 63),
            ("lab-1-class", "Tambah / create class", 72),
            ("lab-1-nama", "Nama kelas & assignment", 79),
            ("lab-1-assignment", "Classwork → Assignment", 83),
            ("lab-1-topic", "Topic", 124),
            ("lab-1-due", "Tarikh due", 143),
            ("lab-1-stream", "Stream → announcement", 163),
            ("lab-1-selesai", "Tamat Lab 1", 185),
        ],
    ),
    (
        "lab-2",
        "Lab 2 · Google Doc",
        199,
        [
            ("lab-2-t1", "Task 1 · Docs", 201),
            ("lab-2-t2", "Task 2 · Share", 230),
            ("lab-2-t3", "Task 3 · Tetapan", 246),
            ("lab-2-t4", "Task 4 · Comment", 268),
            ("lab-2-t5", "Task 5 · Folder", 299),
        ],
    ),
    (
        "lab-3",
        "Lab 3 · Google Form",
        340,
        [
            ("lab-3-t1", "Task 1 · Form", 342),
            ("lab-3-t2", "Task 2 · Imej", 344),
            ("lab-3-t3", "Task 3 · Kuiz", 382),
            ("lab-3-t5", "Task 5 · People / Sheets", 418),
        ],
    ),
    ("selesai", "Selesai", 450, None),
]

URL_RE = re.compile(r"https?://[^\s<>\")\]]+")

# Paragraph presentation (flat Word → visual hierarchy)
STEP_NUM_RE = re.compile(r"^(\d+)\.\s*(.*)$", re.DOTALL)
# Task 1–5 headings: colon, equals, or space after number (e.g. "Task 5 klik …", "Task 4 = bonus").
TASK_HEAD_RE = re.compile(
    r"^(Task\s+[1-5]|TASK\s+[1-5])(?:\s*:\s*|\s*=\s*|\s+)(.*)$",
    re.IGNORECASE | re.DOTALL,
)


def classify_paragraph(text: str) -> str:
    s = text.strip()
    if not s:
        return "body"
    if re.match(r"^LAB\s+\d", s, re.IGNORECASE):
        return "lab_heading"
    if re.match(r"^Bahagian\s+Kuiz", s, re.IGNORECASE):
        return "section_heading"
    if re.match(r"^Cara\s+untuk\s+masuk", s, re.IGNORECASE):
        return "section_heading"
    if re.match(r"^Tahniah\b", s, re.IGNORECASE):
        return "done_banner"
    if TASK_HEAD_RE.match(s):
        return "task_heading"
    if STEP_NUM_RE.match(s):
        return "numbered_step"
    if s.startswith("Reminder:"):
        return "english_callout"
    if "[" in s and "]" in s and ">" in s and len(s) < 140:
        return "ui_breadcrumb"
    if re.match(r"^Selepas masuk Lab Drive", s, re.IGNORECASE):
        return "lead"
    if re.match(r"^Guna cara yang sama", s, re.IGNORECASE):
        return "lead"
    if re.match(r"^Ambil perhatian\s+semasa\s+log\s+in\s+ke\s+GC", s, re.IGNORECASE) and "tengok baik-baik gambar" in s.lower():
        return "gc_login_notice"
    return "body"

# block_index -> list of (button_label, exact_string_to_copy)
MANUAL_COPIES_STATIC: dict[int, list[tuple[str, str]]] = {
    42: [("Arahan Gemini", "bagi jawapan tanpa penerangan")],
    79: [("Nama kelas", "Geography")],
    93: [("Nama assignment", "Capital Cities")],
    125: [("Nama topic", "Cities")],
    221: [("Tajuk dokumen", "Supply List")],
}


def extract_media() -> None:
    if not DOCX.is_file():
        raise FileNotFoundError(DOCX)
    MEDIA_DIR.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(DOCX, "r") as zf:
        for name in zf.namelist():
            if not name.startswith("word/media/") or name.endswith("/"):
                continue
            dest = MEDIA_DIR / Path(name).name
            dest.write_bytes(zf.read(name))


def sanitize_exam_copy_text(text: str) -> str:
    """Strip common Malay instructional prefixes before English exam strings (clipboard accuracy)."""
    s = (text or "").strip()
    if not s:
        return s
    for pat in (
        r"^dan\s+kandungan\s+ayat\s*[-–\u2013:\u00A0]\s*",
        r"^kandungan\s+ayat\s*[-–\u2013:\u00A0]\s*",
        r"^dan\s+kandungan\s+ayat\s+",
        r"^kandungan\s+ayat\s+",
    ):
        s = re.sub(pat, "", s, flags=re.IGNORECASE).strip()
    return s


def trim_paste_snippet_tail(s: str) -> str:
    """Remove Malay tails (parentheticals, 'lalu klik …') from extracted paste text."""
    s = (s or "").strip()
    if not s:
        return s
    s = re.sub(r"\s+lalu\s+klik\b.*$", "", s, flags=re.IGNORECASE).strip()
    s = re.sub(r"\s+Kemudian\b.*$", "", s, flags=re.IGNORECASE).strip()
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"\s*\([^()]{0,400}\)\s*$", "", s).strip()
    return s


def paste_ayat_snippets(text: str) -> list[str]:
    """English strings after 'Paste ayat' / 'paste ayat' (dash/colon); excludes Malay after the snippet."""
    out: list[str] = []
    seen: set[str] = set()
    # Word-boundary so mid-sentence 'paste ayat:' (e.g. Task 4) matches; stop before Pilih / lalu klik / EOS.
    pap = re.compile(
        r"\bPaste\s+ayat\s*[–\u2013\-:]\s*(.+?)(?=\s+Pilih\b|\s+pilih\b|\s+Type\b|\s+Pada\b|\s+lalu\s+klik\b|\s*$)",
        re.IGNORECASE | re.DOTALL,
    )
    for m in pap.finditer(text):
        s = trim_paste_snippet_tail(m.group(1))
        s = sanitize_exam_copy_text(s)
        if s and len(s) < 200 and s not in seen:
            seen.add(s)
            out.append(s)
    return out


def solar_title_from_task1(text: str) -> str | None:
    m = re.search(
        r"tajuk\s+google\s+form\s+kepada\s*[–\u2013\-]\s*Our\s+Solar\s+System",
        text,
        flags=re.IGNORECASE,
    )
    return "Our Solar System" if m else None


def student_name_snippet(text: str) -> str | None:
    m = re.search(
        r"soalan\s+1\s+masukkan\s*[–\u2013\-]\s*Student\s+Name",
        text,
        flags=re.IGNORECASE,
    )
    return "Student Name" if m else None


def solar_section_name(text: str) -> str | None:
    m = re.search(
        r"Namakan\s+ruangan\s+ini\s*[–\u2013\-]\s*Solar\s+System",
        text,
        flags=re.IGNORECASE,
    )
    return "Solar System" if m else None


def mcq_options_line(text: str) -> str | None:
    """Legacy single-line options (only if not handled by split_lab3_planet_options)."""
    if "Mercury" in text and "Venus" in text and "Earth" in text:
        m = re.search(
            r"(Mercury\s+Venus\s+Earth\s+Mars|Jupiter\s+Saturn\s+Uranus\s+Neptune)",
            text,
        )
        return m.group(1) if m else None
    return None


# Lab 3 Google Form — planet names for MC / Dropdown (order matches doc)
LAB3_PLANETS_INNER = ("Mercury", "Venus", "Earth", "Mars")
LAB3_PLANETS_OUTER = ("Jupiter", "Saturn", "Uranus", "Neptune")


def split_lab3_planet_options(text: str) -> tuple[str, ...] | None:
    """Planet option tuples for Lab 3 MC / Dropdown paragraphs (order matches doc)."""
    t = text
    if re.search(r"Mercury\s+Venus\s+Earth\s+Mars", t, re.IGNORECASE):
        return LAB3_PLANETS_INNER
    if re.search(r"Jupiter\s+Saturn\s+Uranus\s+Neptune", t, re.IGNORECASE):
        return LAB3_PLANETS_OUTER
    return None


# Lab 1 GC login: Word order is role picker then Continue (2,1); teach Continue then role.
LAB1_GC_LOGIN_FIGURE_ORDER = ("image7.png", "image6.png")

# Lab 1 Stream → New announcement: Word order is compose dialog then class Stream (2,1); teach Stream first.
LAB1_STREAM_ANNOUNCEMENT_FIGURE_ORDER = ("image22.png", "image21.png")

# Lab 2: "Jika tidak dapat drag" — Word order is Move dialog then Drive (2,1); teach as 1 then 2.
LAB2_DRAG_FIGURE_ORDER = ("image39.png", "image38.png")

# Lab 2 Task 4: Word order is comment pop-up then highlight (2,1); teach highlight first.
LAB2_TASK4_COMMENT_FIGURE_ORDER = ("image33.png", "image32.png")

# Lab 2 TASK 5: Word order is dialog then menu then Drive (3,2,1); teach 1 → 2 → 3.
LAB2_TASK5_FOLDER_FIGURE_ORDER = ("image37.png", "image36.png", "image35.png")

# Lab 3 Task 2: Word media order is result → dialog → add-image (3,2,1); teach as 1,2,3.
LAB3_TASK2_FIGURE_STEP_ORDER = ("image44.png", "image43.png", "image42.png")


def lab1_gc_login_figure_order(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Reorder GC login screenshots; numbered badges 1–2."""
    if len(imgs) != 2:
        return imgs, False
    names = {Path(rel).name for rel in imgs}
    if names != {"image6.png", "image7.png"}:
        return imgs, False
    t = (paragraph_text or "").lower()
    if "ambil perhatian" not in t or "log in ke gc" not in t:
        return imgs, False
    by_name = {Path(rel).name: rel for rel in imgs}
    ordered = [by_name[n] for n in LAB1_GC_LOGIN_FIGURE_ORDER]
    return ordered, True


def lab1_stream_announcement_figure_order(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Reorder Stream UI (1) before announcement dialog (2); screenshots already have step marks."""
    if len(imgs) != 2:
        return imgs, False
    names = {Path(rel).name for rel in imgs}
    if names != {"image21.png", "image22.png"}:
        return imgs, False
    t = (paragraph_text or "").lower()
    if "task terakhir" not in t or "stream" not in t or "new announcement" not in t:
        return imgs, False
    by_name = {Path(rel).name: rel for rel in imgs}
    ordered = [by_name[n] for n in LAB1_STREAM_ANNOUNCEMENT_FIGURE_ORDER]
    return ordered, False


def lab2_drag_figure_order(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Reorder the two drag-fallback screenshots; numbered badges 1–2."""
    if len(imgs) != 2:
        return imgs, False
    names = {Path(rel).name for rel in imgs}
    if names != {"image38.png", "image39.png"}:
        return imgs, False
    t = (paragraph_text or "").lower()
    if "tidak dapat drag" not in t:
        return imgs, False
    by_name = {Path(rel).name: rel for rel in imgs}
    ordered = [by_name[n] for n in LAB2_DRAG_FIGURE_ORDER]
    return ordered, True


def lab2_task4_figure_order(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Reorder highlight step (1) before comment box (2); screenshot labels match."""
    if len(imgs) != 2:
        return imgs, False
    names = {Path(rel).name for rel in imgs}
    if names != {"image32.png", "image33.png"}:
        return imgs, False
    t = (paragraph_text or "").strip()
    if not re.match(r"^Task\s*4\s*:", t, re.IGNORECASE):
        return imgs, False
    if "get red pens" not in t.lower():
        return imgs, False
    by_name = {Path(rel).name: rel for rel in imgs}
    ordered = [by_name[n] for n in LAB2_TASK4_COMMENT_FIGURE_ORDER]
    return ordered, False


def lab2_task5_figure_order(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Reorder new-folder flow: Drive home (1), +New menu (2), name dialog (3)."""
    if len(imgs) != 3:
        return imgs, False
    names = {Path(rel).name for rel in imgs}
    if names != {"image35.png", "image36.png", "image37.png"}:
        return imgs, False
    t = (paragraph_text or "").strip()
    if not re.match(r"^TASK\s*5\s*:", t, re.IGNORECASE):
        return imgs, False
    if "supply list committee" not in t.lower():
        return imgs, False
    by_name = {Path(rel).name: rel for rel in imgs}
    ordered = [by_name[n] for n in LAB2_TASK5_FOLDER_FIGURE_ORDER]
    return ordered, False


def lab2_doc_task1_figure_badges(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Lab 2 Task 1 (+new → Google Docs): two steps, show 1–2 on merged screenshots."""
    if len(imgs) != 2:
        return imgs, False
    if [Path(rel).name for rel in imgs] != ["image25.png", "image26.png"]:
        return imgs, False
    t = (paragraph_text or "").strip()
    if not re.match(r"^Task\s*1\s*:", t, re.IGNORECASE) or "google docs" not in t.lower():
        return imgs, False
    return imgs, True


def lab1_balik_endlab_figure_badges(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Lab 1: question page (1) then lab progress / End Lab (2)."""
    if len(imgs) != 2:
        return imgs, False
    if [Path(rel).name for rel in imgs] != ["image24.png", "image23.png"]:
        return imgs, False
    t = (paragraph_text or "").strip()
    if not t.startswith("Balik ke laman web soalan") or "Check semua task" not in t:
        return imgs, False
    return imgs, True


def merge_lab2_task1_images_under_heading(blocks: list[dict]) -> None:
    """Word puts first screenshot on the intro paragraph; move it under Task 1 (chronological order)."""
    for i in range(1, len(blocks)):
        prev, cur = blocks[i - 1], blocks[i]
        ptxt = (prev.get("text") or "").strip().lower()
        ctxt = (cur.get("text") or "").strip()
        if "masuk ke drive lab 2" not in ptxt:
            continue
        if not re.match(r"^Task\s*1\s*:\s*Klik\s+butang", ctxt, re.IGNORECASE):
            continue
        if "google docs" not in ctxt.lower():
            continue
        pim = [Path(x).name for x in (prev.get("images") or [])]
        cim = [Path(x).name for x in (cur.get("images") or [])]
        if pim == ["image25.png"] and cim == ["image26.png"]:
            cur["images"] = ["media/image25.png", "media/image26.png"]
            prev["images"] = []
            return


def merge_lab2_task2_share_image(blocks: list[dict]) -> None:
    """Word attaches Share screenshot (image27) to supplies paragraph; it belongs under Task 2."""
    for i, cur in enumerate(blocks):
        t = (cur.get("text") or "").strip()
        if not t.lower().startswith("dan kandungan ayat") or "supplies needed" not in t.lower():
            continue
        imgs = [Path(x).name for x in (cur.get("images") or [])]
        if imgs != ["image27.png", "image28.png"]:
            continue
        for j in range(i + 1, min(i + 50, len(blocks))):
            tj = (blocks[j].get("text") or "").strip()
            if re.match(r"^Task\s*2\s*:\s*klik\s+share", tj, re.IGNORECASE):
                share_imgs = list(blocks[j].get("images") or [])
                blocks[j]["images"] = ["media/image27.png"] + share_imgs
                cur["images"] = ["media/image28.png"]
                return


def merge_lab1_balik_endlab_figures(blocks: list[dict]) -> None:
    """Place question-site (image24) and lab UI (image23) in one row: Task 1 then Task 2."""
    for i, cur in enumerate(blocks):
        t = (cur.get("text") or "").strip()
        if not t.startswith("Balik ke laman web soalan") or "Check semua task" not in t:
            continue
        cur_im = [Path(x).name for x in (cur.get("images") or [])]
        if cur_im != ["image23.png"]:
            continue
        for j in range(i + 1, min(i + 8, len(blocks))):
            blk = blocks[j]
            tt = (blk.get("text") or "").strip()
            jim = [Path(x).name for x in (blk.get("images") or [])]
            if not tt and jim == ["image24.png"]:
                cur["images"] = ["media/image24.png", "media/image23.png"]
                blk["images"] = []
                return


def merge_lab2_task5_folder_figures(blocks: list[dict]) -> None:
    """Word puts TASK 5 text and figures on separate blocks; merge and order Drive → menu → dialog."""
    for i in range(1, len(blocks)):
        prev, cur = blocks[i - 1], blocks[i]
        pt = (prev.get("text") or "").strip()
        if not re.match(r"^TASK\s*5\s*:", pt, re.IGNORECASE):
            continue
        if "supply list committee" not in pt.lower():
            continue
        if prev.get("images"):
            continue
        if (cur.get("text") or "").strip():
            continue
        cur_im = [Path(x).name for x in (cur.get("images") or [])]
        if cur_im != ["image35.png", "image36.png", "image37.png"]:
            continue
        prev["images"] = ["media/image37.png", "media/image36.png", "media/image35.png"]
        cur["images"] = []
        return


def lab3_task2_figure_order(imgs: list[str], paragraph_text: str) -> tuple[list[str], bool]:
    """Return (images, show_step_badge). Reorder only the known three-step add-image strip."""
    if len(imgs) != 3:
        return imgs, False
    names = {Path(rel).name for rel in imgs}
    want = set(LAB3_TASK2_FIGURE_STEP_ORDER)
    if names != want:
        return imgs, False
    t = (paragraph_text or "").strip()
    if not re.search(r"Task\s*2\s*:", t, re.IGNORECASE) or "add image" not in t.lower():
        return imgs, False
    by_name = {Path(rel).name: rel for rel in imgs}
    ordered = [by_name[n] for n in LAB3_TASK2_FIGURE_STEP_ORDER]
    return ordered, True


def build_extra_copies_for_block(
    idx: int, text: str, manual: dict[int, list[tuple[str, str]]]
) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = list(manual.get(idx, []))
    if idx == 42:
        pass
    else:
        for snip in paste_ayat_snippets(text):
            if snip not in [x[1] for x in found]:
                found.append(("Ayat tampal", snip))
    st = solar_title_from_task1(text)
    if st:
        found.append(("Tajuk form", st))
    sn = student_name_snippet(text)
    if sn:
        found.append(("Soalan 1 (pendek)", sn))
    ss = solar_section_name(text)
    if ss:
        found.append(("Nama ruangan imej", ss))
    names = split_lab3_planet_options(text)
    if names:
        found.append(
            (
                "4 pilihan (satu baris satu nama)",
                "\n".join(names),
            )
        )
    else:
        mc = mcq_options_line(text)
        if mc:
            found.append(("Pilihan jawapan", mc))
    return found


def split_urls(s: str) -> list[tuple[str, str | None]]:
    """Return list of (fragment, url_or_none)."""
    parts: list[tuple[str, str | None]] = []
    pos = 0
    for m in URL_RE.finditer(s):
        if m.start() > pos:
            parts.append((s[pos : m.start()], None))
        parts.append((m.group(0), m.group(0)))
        pos = m.end()
    if pos < len(s):
        parts.append((s[pos:], None))
    if not parts:
        parts.append((s, None))
    return parts


def render_text_with_urls(s: str) -> str:
    chunks: list[str] = []
    for frag, url in split_urls(s):
        esc = html.escape(frag)
        if url:
            chunks.append(
                '<span class="inline-flex flex-wrap items-center gap-2 align-middle">'
                f'<code class="rounded bg-slate-100 px-2 py-1 text-sm break-all">{esc}</code>'
                '<button type="button" class="copy-btn shrink-0 rounded border border-slate-300 bg-white '
                'px-2 py-1 text-xs font-medium text-slate-700 shadow-sm hover:bg-slate-50" '
                f'data-copy="{html.escape(url, quote=True)}">Salin</button></span>'
            )
        else:
            chunks.append(esc)
    return "".join(chunks)


def merge_adjacent_runs_list(runs: list[dict]) -> list[dict]:
    if not runs:
        return []
    out = [dict(runs[0])]
    for r in runs[1:]:
        if r["b"] == out[-1]["b"] and r.get("c") == out[-1].get("c"):
            out[-1]["t"] += r["t"]
        else:
            out.append(dict(r))
    return out


def slice_runs(runs: list[dict], start: int, end: int) -> list[dict]:
    if start >= end or not runs:
        return []
    out: list[dict] = []
    pos = 0
    for r in runs:
        t = r["t"]
        ln = len(t)
        seg_a = max(0, start - pos)
        seg_b = min(ln, end - pos)
        if seg_b > seg_a:
            out.append({"t": t[seg_a:seg_b], "b": r["b"], "c": r.get("c")})
        pos += ln
        if pos >= end:
            break
    return merge_adjacent_runs_list(out)


def render_run_chunk_to_html(t: str, bold: bool, color: str | None) -> str:
    inner = render_text_with_urls(t)
    classes: list[str] = []
    style_parts: list[str] = []
    if color and len(color) == 6 and all(c in "0123456789ABCDEFabcdef" for c in color):
        style_parts.append(f"color:#{color.upper()}")
    if bold:
        classes.append("font-bold")
    if not classes and not style_parts:
        return inner
    cls = f' class="{" ".join(classes)}"' if classes else ""
    sty = f' style="{";".join(style_parts)}"' if style_parts else ""
    return f"<span{cls}{sty}>{inner}</span>"


def render_runs_to_html(runs: list[dict]) -> str:
    return "".join(
        render_run_chunk_to_html(r["t"], bool(r.get("b")), r.get("c"))
        for r in runs
        if r.get("t") or r.get("b") or r.get("c")
    )


def render_copy_row(label: str, value: str) -> str:
    safe = html.escape(value, quote=True)
    shown = html.escape(value)
    return (
        f'<div class="mt-3 flex flex-wrap items-start gap-2 rounded-lg border border-emerald-200/90 bg-emerald-50/90 p-3 shadow-sm">'
        f'<div class="min-w-0 flex-1"><span class="text-[11px] font-bold uppercase tracking-wider text-emerald-800">'
        f"{html.escape(label)}</span>"
        f'<pre class="mt-1.5 whitespace-pre-wrap break-words text-sm leading-relaxed text-slate-800">{shown}</pre></div>'
        f'<button type="button" class="copy-btn shrink-0 rounded-md bg-emerald-600 px-3 py-1.5 text-sm font-medium text-white '
        f'shadow-sm hover:bg-emerald-700" data-copy="{safe}">Salin</button></div>'
    )


def wrap_paragraph_block(
    kind: str,
    raw_text: str,
    inner_rich: str,
    copy_rows: str,
    *,
    body_inner: str | None = None,
    task_inner: str | None = None,
) -> str:
    """Wrap paragraph HTML by semantic kind (raw_text used for step/task parsing)."""
    st = raw_text.strip()
    if kind == "lab_heading":
        return (
            f'<h2 class="not-prose mb-5 mt-10 scroll-mt-32 border-b border-indigo-200 pb-3 text-xl font-bold '
            f'tracking-tight text-indigo-950 first:mt-0 sm:text-2xl">{inner_rich}</h2>{copy_rows}'
        )
    if kind == "section_heading":
        return (
            f'<h2 class="not-prose mb-4 mt-8 scroll-mt-32 border-b border-slate-200 pb-2 text-xl font-bold '
            f'tracking-tight text-slate-900 first:mt-0">{inner_rich}</h2>{copy_rows}'
        )
    if kind == "done_banner":
        return (
            f'<p class="not-prose my-8 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-center text-base '
            f'font-semibold text-amber-900 shadow-sm">{inner_rich}</p>{copy_rows}'
        )
    if kind == "task_heading":
        m = TASK_HEAD_RE.match(st)
        if m:
            tag, rest = m.group(1), m.group(2).strip()
            if task_inner is not None:
                inner_rest = task_inner
            else:
                inner_rest = render_text_with_urls(rest) if rest else ""
            head = html.escape(tag)
            return (
                f'<h3 class="not-prose mb-3 mt-8 flex flex-wrap items-baseline gap-2 border-l-4 border-slate-700 pl-3 '
                f'text-base font-bold text-slate-950 sm:text-[17px]">'
                f'<span class="shrink-0 rounded-md bg-slate-800 px-2 py-0.5 text-xs font-bold uppercase tracking-wide '
                f'text-white">{head}</span>'
                f'<span class="min-w-0 flex-1 font-bold leading-snug text-slate-950">{inner_rest}</span></h3>{copy_rows}'
            )
    if kind == "numbered_step":
        m = STEP_NUM_RE.match(st)
        if m:
            num, rest = m.group(1), m.group(2)
            if body_inner is not None:
                body = body_inner
            else:
                body = render_text_with_urls(rest)
            return (
                f'<div class="not-prose mb-4 flex gap-3 sm:gap-4">'
                f'<span class="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-slate-200 text-sm '
                f'font-bold text-slate-800 shadow-sm ring-1 ring-slate-300/80">{html.escape(num)}</span>'
                f'<div class="min-w-0 flex-1 pt-0.5"><p class="text-[15px] leading-7 text-slate-700">{body}</p>{copy_rows}</div></div>'
            )
    if kind == "english_callout":
        return (
            f'<blockquote class="not-prose my-4 border-l-4 border-blue-400 bg-blue-50/80 py-2 pl-4 pr-3 text-[15px] '
            f'leading-relaxed text-slate-800 shadow-sm"><p>{inner_rich}</p></blockquote>{copy_rows}'
        )
    if kind == "ui_breadcrumb":
        return (
            f'<p class="not-prose my-3 rounded-lg border border-slate-200 bg-slate-100 px-3 py-2.5 font-mono text-sm '
            f'leading-relaxed text-slate-800 shadow-inner">{inner_rich}</p>{copy_rows}'
        )
    if kind == "lead":
        return (
            f'<p class="not-prose mb-4 text-[15px] font-medium leading-7 text-slate-700">{inner_rich}</p>{copy_rows}'
        )
    if kind == "gc_login_notice":
        return (
            f'<p class="not-prose mb-3.5 text-[15px] font-bold leading-7 text-slate-950 last:mb-0">{inner_rich}</p>{copy_rows}'
        )
    # body
    return (
        f'<p class="not-prose mb-3.5 text-[15px] leading-7 text-slate-700 last:mb-0">{inner_rich}</p>{copy_rows}'
    )


def render_block(idx: int, block: dict, anchor_ids: set[str], manual: dict[int, list[tuple[str, str]]]) -> str:
    t = block.get("text") or ""
    imgs = block.get("images") or []
    rows = block.get("rows")

    parts: list[str] = []
    bid = f"b-{idx}"

    if block.get("type") == "tbl" and rows:
        id_attr = f' id="{bid}"' if bid in anchor_ids else ""
        scroll = " scroll-mt-28" if id_attr else ""
        parts.append(f'<div class="my-4 overflow-x-auto{scroll}"{id_attr}>')
        parts.append('<table class="min-w-full border border-slate-200 text-sm">')
        for r in rows:
            parts.append("<tr>")
            for c in r:
                parts.append(f'<td class="border border-slate-200 px-3 py-2">{html.escape(c)}</td>')
            parts.append("</tr>")
        parts.append("</table></div>")
        return "".join(parts)

    if not t.strip() and not imgs:
        return ""

    is_anchor = bid in anchor_ids
    id_attr = f' id="{bid}"' if is_anchor else ""
    inner_parts: list[str] = []

    if t.strip():
        runs = block.get("runs") or []
        flat = "".join(r["t"] for r in runs)
        st = flat.strip()
        text_for_kind = st if st else t.strip()
        kind = classify_paragraph(text_for_kind)
        extras_text = (block.get("text") or "").strip()
        extras = build_extra_copies_for_block(idx, extras_text, manual)
        copy_rows = "".join(render_copy_row(lbl, val) for lbl, val in extras if val)

        loff = len(flat) - len(flat.lstrip())
        hi = loff + len(st) if st else loff

        body_inner: str | None = None
        task_inner: str | None = None
        if runs and st:
            inner_rich = render_runs_to_html(slice_runs(runs, loff, hi))
            if kind == "numbered_step":
                m = STEP_NUM_RE.match(st)
                if m:
                    pre_len = len(m.group(0)) - len(m.group(2))
                    body_inner = render_runs_to_html(slice_runs(runs, loff + pre_len, hi))
            elif kind == "task_heading":
                m = TASK_HEAD_RE.match(st)
                if m:
                    pre_len = len(m.group(0)) - len(m.group(2))
                    task_inner = render_runs_to_html(slice_runs(runs, loff + pre_len, hi))
        else:
            inner_rich = render_text_with_urls(t.strip())

        inner_parts.append(
            wrap_paragraph_block(
                kind,
                text_for_kind,
                inner_rich,
                copy_rows,
                body_inner=body_inner,
                task_inner=task_inner,
            )
        )

    if imgs:
        imgs, step_badge = lab1_gc_login_figure_order(imgs, t)
        if not step_badge:
            imgs, step_badge = lab1_stream_announcement_figure_order(imgs, t)
        if not step_badge:
            imgs, step_badge = lab2_drag_figure_order(imgs, t)
        if not step_badge:
            imgs, step_badge = lab3_task2_figure_order(imgs, t)
        if not step_badge:
            imgs, step_badge = lab2_task5_figure_order(imgs, t)
        if not step_badge:
            imgs, step_badge = lab2_doc_task1_figure_badges(imgs, t)
        if not step_badge:
            imgs, step_badge = lab2_task4_figure_order(imgs, t)
        if not step_badge:
            imgs, step_badge = lab1_balik_endlab_figure_badges(imgs, t)
        fig_inner: list[str] = []
        for step_i, rel in enumerate(imgs):
            name = Path(rel).name
            src = "media/" + name
            badge = ""
            if step_badge:
                n = str(step_i + 1)
                badge = (
                    f'<span class="pointer-events-none absolute left-2 top-2 z-10 flex h-8 w-8 items-center '
                    f'justify-center rounded-full bg-rose-600 text-sm font-bold text-white shadow-md ring-2 ring-white" '
                    f'aria-hidden="true">{html.escape(n)}</span>'
                )
            fig_inner.append(
                f'<div class="relative">'
                f'{badge}'
                f'<figure class="content-figure not-prose flex min-h-[80px] flex-col overflow-hidden rounded-xl border '
                f'border-slate-200/90 bg-slate-50 shadow-md ring-1 ring-slate-200/50">'
                f'<div class="flex flex-1 items-center justify-center p-2 sm:p-4 lg:p-3">'
                f'<img src="{html.escape(src)}" alt="" '
                f'class="content-figure-img w-full cursor-zoom-in rounded-md object-contain transition hover:opacity-95 '
                f'max-h-[min(68vh,520px)] lg:max-h-[min(48vh,540px)]" '
                f'loading="lazy" decoding="async"></div></figure></div>'
            )
        if len(fig_inner) == 1:
            inner_parts.append(
                f'<div class="not-prose my-6 lg:mx-auto lg:max-w-[42rem]">{fig_inner[0]}</div>'
            )
        else:
            inner_parts.append(
                '<div class="not-prose my-6 grid grid-cols-1 gap-4 sm:grid-cols-2 sm:gap-5 lg:mx-auto lg:max-w-5xl lg:gap-6">'
                + "".join(fig_inner)
                + "</div>"
            )

    body = "".join(inner_parts)
    if is_anchor:
        return f'<article{id_attr} class="anchor-block mb-2 scroll-mt-28">{body}</article>'
    return body


def collect_anchor_ids() -> set[str]:
    ids: set[str] = set()
    for _sid, _label, main_idx, children in NAV:
        if main_idx is not None:
            ids.add(f"b-{main_idx}")
        if children:
            for cid, _cl, cidx in children:
                ids.add(f"b-{cidx}")
    return ids


def render_nav_html() -> str:
    lines: list[str] = []
    for sid, label, main_idx, children in NAV:
        lines.append('<div class="mb-4 border-b border-slate-100 pb-3 last:mb-0 last:border-0 last:pb-0">')
        lines.append(
            f'<a href="#b-{main_idx}" class="nav-link nav-parent block rounded-md px-2 py-1.5 text-sm font-semibold '
            f'text-slate-900 hover:bg-indigo-50" data-nav-target="b-{main_idx}">{html.escape(label)}</a>'
        )
        if children:
            lines.append(
                '<div class="nav-children mt-2 space-y-0.5 border-l-2 border-indigo-200 pl-3 ml-1.5">'
            )
            for cid, cl, cidx in children:
                lines.append(
                    f'<a href="#b-{cidx}" class="nav-link nav-child block rounded px-2 py-1 text-xs leading-snug '
                    f'text-slate-600 hover:bg-slate-100 hover:text-slate-900" data-nav-target="b-{cidx}">'
                    f'<span class="mr-1.5 text-indigo-400">·</span>{html.escape(cl)}</a>'
                )
            lines.append("</div>")
        lines.append("</div>")
    return "\n".join(lines)


def build_html(blocks: list[dict], manual: dict[int, list[tuple[str, str]]]) -> str:
    merge_lab2_task1_images_under_heading(blocks)
    merge_lab2_task2_share_image(blocks)
    merge_lab1_balik_endlab_figures(blocks)
    merge_lab2_task5_folder_figures(blocks)
    anchor_ids = collect_anchor_ids()
    no_anchor: set[str] = set()
    body_chunks: list[str] = []

    body_chunks.append(
        '<div id="b-0" class="doc-shell mx-auto mb-10 max-w-[52rem] scroll-mt-28 rounded-2xl border border-slate-200/90 '
        'bg-white p-6 shadow-sm sm:p-8">'
        '<p class="not-prose mb-6 rounded-lg bg-slate-100 px-3 py-2 text-xs font-medium uppercase tracking-wide '
        'text-slate-500">Bahagian A · Masuk ke Lab Drive</p>'
    )
    for i in range(0, 41):
        chunk = render_block(i, blocks[i], no_anchor, manual)
        if chunk:
            body_chunks.append(chunk)
    body_chunks.append("</div>")

    body_chunks.append(
        '<div id="b-41" class="doc-shell mx-auto mb-10 max-w-[52rem] scroll-mt-28 rounded-2xl border border-slate-200/90 '
        'bg-white p-6 shadow-sm sm:p-8">'
        '<p class="not-prose mb-6 rounded-lg bg-violet-50 px-3 py-2 text-xs font-medium uppercase tracking-wide '
        'text-violet-700">Bahagian B · Kuiz</p>'
    )
    for i in range(41, 53):
        chunk = render_block(i, blocks[i], no_anchor, manual)
        if chunk:
            body_chunks.append(chunk)
    body_chunks.append("</div>")

    body_chunks.append(
        '<section class="lab-section mx-auto mb-12 max-w-[52rem] scroll-mt-24 rounded-2xl border-2 border-indigo-100 '
        'bg-gradient-to-b from-indigo-50/40 to-white p-6 shadow-sm sm:p-8" aria-label="Lab 1">'
        '<header class="not-prose mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-indigo-700">'
        '<span class="h-1.5 w-1.5 rounded-full bg-indigo-500"></span>Praktikal · Lab 1</header>'
    )
    for i in range(53, 199):
        chunk = render_block(i, blocks[i], anchor_ids, manual)
        if chunk:
            body_chunks.append(chunk)
    body_chunks.append("</section>")

    body_chunks.append(
        '<section class="lab-section mx-auto mb-12 max-w-[52rem] scroll-mt-24 rounded-2xl border-2 border-emerald-100 '
        'bg-gradient-to-b from-emerald-50/40 to-white p-6 shadow-sm sm:p-8" aria-label="Lab 2">'
        '<header class="not-prose mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-emerald-800">'
        '<span class="h-1.5 w-1.5 rounded-full bg-emerald-500"></span>Praktikal · Lab 2</header>'
    )
    for i in range(199, 340):
        chunk = render_block(i, blocks[i], anchor_ids, manual)
        if chunk:
            body_chunks.append(chunk)
    body_chunks.append("</section>")

    body_chunks.append(
        '<section class="lab-section mx-auto mb-8 max-w-[52rem] scroll-mt-24 rounded-2xl border-2 border-amber-100 '
        'bg-gradient-to-b from-amber-50/35 to-white p-6 shadow-sm sm:p-8" aria-label="Lab 3">'
        '<header class="not-prose mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-amber-900">'
        '<span class="h-1.5 w-1.5 rounded-full bg-amber-500"></span>Praktikal · Lab 3</header>'
    )
    for i in range(340, len(blocks)):
        chunk = render_block(i, blocks[i], anchor_ids, manual)
        if chunk:
            body_chunks.append(chunk)
    body_chunks.append("</section>")

    nav_html = render_nav_html()

    return f"""<!DOCTYPE html>
<html lang="ms">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>GCE Level 1 — Panduan Lab</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="min-h-screen bg-slate-100 text-slate-900 antialiased">
  <header class="sticky top-0 z-40 border-b border-slate-200 bg-white/95 backdrop-blur">
    <div class="mx-auto flex max-w-6xl items-center justify-between gap-3 px-4 py-3">
      <h1 class="text-lg font-semibold tracking-tight text-slate-900">GCE Level 1</h1>
      <p class="hidden text-sm text-slate-500 sm:block">Panduan langkah demi langkah · salin teks dengan selamat</p>
    </div>
  </header>

  <!-- Mobile: overlay + drawer (same links as desktop sidebar) -->
  <div id="mobile-nav-layer" class="fixed inset-0 z-[100] hidden lg:hidden" aria-hidden="true">
    <div id="mobile-nav-overlay" class="absolute inset-0 bg-slate-900/50 backdrop-blur-[2px]"></div>
    <div class="absolute inset-x-0 bottom-0 flex max-h-[min(85vh,36rem)] flex-col rounded-t-2xl border border-slate-200 bg-white shadow-2xl">
      <div class="flex shrink-0 items-center justify-between border-b border-slate-100 px-4 py-3">
        <span class="text-sm font-bold text-slate-900">Navigasi</span>
        <button type="button" id="mobile-nav-close" class="rounded-lg px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-100">Tutup</button>
      </div>
      <nav class="min-h-0 flex-1 overflow-y-auto overscroll-contain p-4 text-sm" aria-label="Navigasi mudah alih">
        {nav_html}
      </nav>
    </div>
  </div>

  <button type="button" id="mobile-nav-open" class="fixed left-1/2 z-[90] flex -translate-x-1/2 items-center gap-2 rounded-full border border-indigo-200 bg-indigo-600 px-5 py-3 text-sm font-semibold text-white shadow-lg ring-2 ring-white/90 hover:bg-indigo-700 lg:hidden" style="bottom: max(1rem, env(safe-area-inset-bottom, 0px));">
    <svg class="h-5 w-5 shrink-0 opacity-90" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor" aria-hidden="true"><path stroke-linecap="round" stroke-linejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" /></svg>
    Navigasi
  </button>

  <div class="mx-auto flex max-w-6xl flex-col gap-6 px-4 py-8 lg:flex-row lg:items-stretch">
    <!-- items-stretch so aside column is as tall as main; inner nav stays sticky while scrolling -->
    <aside class="hidden w-60 shrink-0 lg:block lg:min-h-0" aria-label="Navigasi sisi">
      <nav class="sticky top-20 z-30 max-h-[calc(100vh-5.5rem)] overflow-y-auto rounded-xl border border-slate-200 bg-white p-4 text-sm shadow-sm" id="sidebar-nav">
        <p class="mb-2 text-xs font-bold uppercase tracking-wide text-slate-500">Navigasi</p>
        {nav_html}
      </nav>
    </aside>

    <main class="min-w-0 flex-1 rounded-2xl border border-slate-200/80 bg-white p-6 pb-24 shadow-md sm:p-10 sm:pb-24 lg:pb-10 lg:shadow-lg" id="main-content">
      <div class="not-prose mb-8 rounded-xl border border-slate-200 bg-slate-50/80 px-4 py-3 text-sm leading-relaxed text-slate-600">
        <strong class="text-slate-800">Cara guna:</strong> gunakan butang <span class="rounded bg-white px-1.5 py-0.5 font-medium text-slate-800 ring-1 ring-slate-200">Salin</span> untuk URL dan teks penting.
        Hoskan pada <strong class="text-slate-800">HTTPS</strong> supaya salin ke papan kereta stabil pada semua pelayar.
      </div>
      <div class="prose-flow max-w-none">
      {"".join(body_chunks)}
      </div>
    </main>
  </div>

  <!-- Image lightbox: click any figure in article to enlarge (desktop + mobile); pinch / buttons to zoom -->
  <dialog id="img-lightbox" class="z-[200] m-0 h-[100dvh] w-full max-w-none border-0 bg-slate-950 p-0 text-white outline-none [&::backdrop]:bg-black/70 [&::backdrop]:backdrop-blur-sm" aria-label="Gambar diperbesar">
    <div class="flex h-full min-h-0 flex-col">
      <div class="flex shrink-0 flex-col gap-2 border-b border-white/10 px-3 py-2 sm:flex-row sm:items-center sm:justify-between">
        <span class="max-w-[min(100%,20rem)] text-xs font-medium leading-snug opacity-90 sm:max-w-none">Cubit jari (pinch) atau butang +/− untuk zum. Bila dizum, seret satu jari pada gambar untuk gerak. Pada komputer: klik kiri pada gambar lalu seret (drag). Klik kawasan kosong, Tutup, atau Esc untuk keluar.</span>
        <div class="flex flex-wrap items-center justify-end gap-2">
          <div class="flex items-center gap-1 rounded-lg bg-white/10 p-1" role="group" aria-label="Zoom gambar">
            <button type="button" id="img-lb-zoom-out" class="min-h-[44px] min-w-[44px] rounded-md text-xl font-bold leading-none text-white hover:bg-white/15 active:bg-white/25">−</button>
            <button type="button" id="img-lb-zoom-reset" class="min-h-[44px] min-w-[52px] rounded-md text-xs font-semibold tabular-nums text-white hover:bg-white/15 active:bg-white/25">100%</button>
            <button type="button" id="img-lb-zoom-in" class="min-h-[44px] min-w-[44px] rounded-md text-xl font-bold leading-none text-white hover:bg-white/15 active:bg-white/25">+</button>
          </div>
          <form method="dialog">
            <button type="submit" class="rounded-lg bg-white/10 px-4 py-2 text-sm font-semibold hover:bg-white/20">Tutup</button>
          </form>
        </div>
      </div>
      <div id="img-lightbox-body" class="relative flex min-h-0 flex-1 overflow-hidden overscroll-none p-3" style="touch-action: none;">
        <div id="img-lb-zoom-wrap" class="m-auto flex min-h-min min-w-min touch-none items-center justify-center p-2 will-change-transform" style="transform: translate(0px, 0px) scale(1); transform-origin: center center;">
          <img id="img-lightbox-img" src="" alt="" class="max-h-[min(86dvh,1100px)] w-auto max-w-[min(100vw,1100px)] object-contain select-none" draggable="false">
        </div>
      </div>
    </div>
  </dialog>

  <script>
  (function () {{
    document.querySelectorAll(".copy-btn").forEach(function (btn) {{
      btn.addEventListener("click", function () {{
        var t = btn.getAttribute("data-copy");
        if (!t) return;
        function done(ok) {{
          var prev = btn.textContent;
          btn.textContent = ok ? "Disalin" : "Gagal";
          btn.disabled = true;
          setTimeout(function () {{
            btn.textContent = prev;
            btn.disabled = false;
          }}, 1600);
        }}
        if (navigator.clipboard && navigator.clipboard.writeText) {{
          navigator.clipboard.writeText(t).then(function () {{ done(true); }}).catch(function () {{ fallback(); }});
        }} else {{
          fallback();
        }}
        function fallback() {{
          var ta = document.createElement("textarea");
          ta.value = t;
          ta.style.position = "fixed";
          ta.style.left = "-9999px";
          document.body.appendChild(ta);
          ta.select();
          try {{
            document.execCommand("copy");
            done(true);
          }} catch (e) {{
            done(false);
          }}
          document.body.removeChild(ta);
        }}
      }});
    }});

    var mobileLayer = document.getElementById("mobile-nav-layer");
    var mobileOpen = document.getElementById("mobile-nav-open");
    var mobileClose = document.getElementById("mobile-nav-close");
    var mobileOverlay = document.getElementById("mobile-nav-overlay");
    function openMobileNav() {{
      if (!mobileLayer) return;
      mobileLayer.classList.remove("hidden");
      mobileLayer.setAttribute("aria-hidden", "false");
      document.body.style.overflow = "hidden";
    }}
    function closeMobileNav() {{
      if (!mobileLayer) return;
      mobileLayer.classList.add("hidden");
      mobileLayer.setAttribute("aria-hidden", "true");
      document.body.style.overflow = "";
    }}
    if (mobileOpen) mobileOpen.addEventListener("click", openMobileNav);
    if (mobileClose) mobileClose.addEventListener("click", closeMobileNav);
    if (mobileOverlay) mobileOverlay.addEventListener("click", closeMobileNav);
    document.querySelectorAll("#mobile-nav-layer a.nav-link").forEach(function (a) {{
      a.addEventListener("click", closeMobileNav);
    }});
    document.addEventListener("keydown", function (e) {{
      if (e.key === "Escape" && mobileLayer && !mobileLayer.classList.contains("hidden")) closeMobileNav();
    }});

    var imgLb = document.getElementById("img-lightbox");
    var imgLbImg = document.getElementById("img-lightbox-img");
    var imgLbBody = document.getElementById("img-lightbox-body");
    var imgLbWrap = document.getElementById("img-lb-zoom-wrap");
    var imgLbZoomIn = document.getElementById("img-lb-zoom-in");
    var imgLbZoomOut = document.getElementById("img-lb-zoom-out");
    var imgLbZoomReset = document.getElementById("img-lb-zoom-reset");
    var lbScale = 1;
    var lbTx = 0;
    var lbTy = 0;
    var LB_MIN = 0.35;
    var LB_MAX = 5;
    var pinchStartDist = 0;
    var pinchStartScale = 1;
    var panArmed = false;
    var panStartX = 0;
    var panStartY = 0;
    var panStartTx = 0;
    var panStartTy = 0;
    var mousePan = false;
    var mouseStartX = 0;
    var mouseStartY = 0;
    var mouseStartTx = 0;
    var mouseStartTy = 0;
    function lbClamp(s) {{
      return Math.max(LB_MIN, Math.min(LB_MAX, s));
    }}
    function lbClampPan() {{
      if (!imgLbBody || !imgLbImg) return;
      var vw = imgLbBody.clientWidth;
      var vh = imgLbBody.clientHeight;
      var iw = imgLbImg.offsetWidth * lbScale;
      var ih = imgLbImg.offsetHeight * lbScale;
      var pad = 32;
      var maxX = Math.max(0, (iw - vw) / 2 + pad);
      var maxY = Math.max(0, (ih - vh) / 2 + pad);
      lbTx = Math.max(-maxX, Math.min(maxX, lbTx));
      lbTy = Math.max(-maxY, Math.min(maxY, lbTy));
    }}
    function lbApplyScale() {{
      if (!imgLbWrap) return;
      lbScale = lbClamp(lbScale);
      if (lbScale <= 1.02) {{
        lbTx = 0;
        lbTy = 0;
      }}
      lbClampPan();
      imgLbWrap.style.transform = "translate(" + lbTx + "px, " + lbTy + "px) scale(" + lbScale + ")";
      imgLbWrap.style.transformOrigin = "center center";
      if (imgLbZoomReset) imgLbZoomReset.textContent = Math.round(lbScale * 100) + "%";
      if (imgLbWrap)
        imgLbWrap.style.cursor = mousePan ? "grabbing" : lbScale > 1.02 ? "grab" : "";
    }}
    function lbEndMousePan() {{
      if (!mousePan) return;
      mousePan = false;
      document.removeEventListener("mousemove", lbMousePanMove);
      document.removeEventListener("mouseup", lbMousePanUp);
      if (imgLbWrap) imgLbWrap.style.cursor = lbScale > 1.02 ? "grab" : "";
    }}
    function lbMousePanMove(e) {{
      if (!mousePan || !imgLb || !imgLb.open) return;
      e.preventDefault();
      lbTx = mouseStartTx + (e.clientX - mouseStartX);
      lbTy = mouseStartTy + (e.clientY - mouseStartY);
      lbClampPan();
      lbApplyScale();
    }}
    function lbMousePanUp() {{
      lbEndMousePan();
    }}
    function lbResetZoom() {{
      lbScale = 1;
      lbTx = 0;
      lbTy = 0;
      pinchStartDist = 0;
      panArmed = false;
      lbEndMousePan();
      lbApplyScale();
    }}
    function openImgLightbox(src) {{
      if (!imgLb || !imgLbImg || !src) return;
      lbResetZoom();
      imgLbImg.src = src;
      imgLbImg.onload = function () {{
        lbClampPan();
        lbApplyScale();
      }};
      imgLb.showModal();
    }}
    function closeImgLightbox() {{
      if (imgLb && imgLb.open) imgLb.close();
      if (imgLbImg) imgLbImg.removeAttribute("src");
      lbResetZoom();
    }}
    document.querySelectorAll("#main-content .content-figure img.content-figure-img").forEach(function (im) {{
      im.addEventListener("click", function () {{
        openImgLightbox(im.currentSrc || im.getAttribute("src") || "");
      }});
    }});
    if (imgLbZoomIn) imgLbZoomIn.addEventListener("click", function (e) {{
      e.stopPropagation();
      lbScale = lbClamp(lbScale + 0.25);
      lbApplyScale();
    }});
    if (imgLbZoomOut) imgLbZoomOut.addEventListener("click", function (e) {{
      e.stopPropagation();
      lbScale = lbClamp(lbScale - 0.25);
      lbApplyScale();
    }});
    if (imgLbZoomReset) imgLbZoomReset.addEventListener("click", function (e) {{
      e.stopPropagation();
      lbResetZoom();
    }});
    function lbTouchDist(a, b) {{
      return Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
    }}
    if (imgLbBody) {{
      imgLbBody.addEventListener("click", function (e) {{
        if (e.target === imgLbBody) closeImgLightbox();
      }});
      imgLbBody.addEventListener("touchstart", function (e) {{
        if (!imgLb || !imgLb.open) return;
        if (e.touches.length === 2) {{
          panArmed = false;
          pinchStartDist = lbTouchDist(e.touches[0], e.touches[1]);
          pinchStartScale = lbScale;
        }} else if (e.touches.length === 1 && lbScale > 1.02 && (e.target === imgLbImg || e.target === imgLbWrap)) {{
          panArmed = true;
          panStartX = e.touches[0].clientX;
          panStartY = e.touches[0].clientY;
          panStartTx = lbTx;
          panStartTy = lbTy;
        }}
      }}, {{ passive: true }});
      imgLbBody.addEventListener("touchmove", function (e) {{
        if (!imgLb || !imgLb.open) return;
        if (e.touches.length === 2 && pinchStartDist > 10) {{
          e.preventDefault();
          var d = lbTouchDist(e.touches[0], e.touches[1]);
          lbScale = lbClamp(pinchStartScale * (d / pinchStartDist));
          lbClampPan();
          lbApplyScale();
        }} else if (e.touches.length === 1 && panArmed && pinchStartDist === 0 && lbScale > 1.02) {{
          e.preventDefault();
          var t = e.touches[0];
          lbTx = panStartTx + (t.clientX - panStartX);
          lbTy = panStartTy + (t.clientY - panStartY);
          lbClampPan();
          lbApplyScale();
        }}
      }}, {{ passive: false }});
      imgLbBody.addEventListener("touchend", function (e) {{
        if (!e.touches || e.touches.length < 2) pinchStartDist = 0;
        if (!e.touches || e.touches.length === 0) panArmed = false;
      }}, {{ passive: true }});
      imgLbBody.addEventListener("wheel", function (e) {{
        if (!imgLb || !imgLb.open) return;
        if (e.ctrlKey) {{
          e.preventDefault();
          lbScale = lbClamp(lbScale - e.deltaY * 0.012);
          lbApplyScale();
        }}
      }}, {{ passive: false }});
    }}
    if (imgLbWrap) {{
      imgLbWrap.addEventListener("mousedown", function (e) {{
        if (!imgLb || !imgLb.open || e.button !== 0) return;
        if (lbScale <= 1.02) return;
        if (e.target !== imgLbImg && e.target !== imgLbWrap) return;
        e.preventDefault();
        mousePan = true;
        mouseStartX = e.clientX;
        mouseStartY = e.clientY;
        mouseStartTx = lbTx;
        mouseStartTy = lbTy;
        imgLbWrap.style.cursor = "grabbing";
        document.addEventListener("mousemove", lbMousePanMove);
        document.addEventListener("mouseup", lbMousePanUp);
      }});
    }}
    if (imgLb) {{
      imgLb.addEventListener("close", function () {{
        if (imgLbImg) imgLbImg.removeAttribute("src");
        lbResetZoom();
      }});
    }}

    var navLinks = [].slice.call(document.querySelectorAll("a.nav-link[data-nav-target]"));
    var byId = {{}};
    navLinks.forEach(function (a) {{
      var id = a.getAttribute("data-nav-target");
      var el = document.getElementById(id);
      if (el) byId[id] = el;
    }});
    if (!window.IntersectionObserver) return;
    var obs = new IntersectionObserver(
      function (entries) {{
        entries.forEach(function (en) {{
          if (!en.isIntersecting) return;
          var id = en.target.id;
          navLinks.forEach(function (a) {{
            var on = a.getAttribute("data-nav-target") === id;
            a.classList.toggle("bg-indigo-100", on);
            a.classList.toggle("ring-1", on);
            a.classList.toggle("ring-inset", on);
            a.classList.toggle("ring-indigo-200", on);
            a.classList.toggle("font-semibold", on);
          }});
        }});
      }},
      {{ rootMargin: "-20% 0px -55% 0px", threshold: 0 }}
    );
    Object.keys(byId).forEach(function (k) {{ obs.observe(byId[k]); }});
  }})();
  </script>
</body>
</html>
"""


def main() -> int:
    if not OUTLINE.is_file():
        print("Run extract_docx.py first:", OUTLINE, file=sys.stderr)
        return 1
    extract_media()
    blocks = json.loads(OUTLINE.read_text(encoding="utf-8"))
    manual: dict[int, list[tuple[str, str]]] = {k: list(v) for k, v in MANUAL_COPIES_STATIC.items()}
    if 95 < len(blocks):
        t95 = sanitize_exam_copy_text(blocks[95].get("text") or "")
        manual[95] = [("Instructions (English)", t95)]
    if 172 < len(blocks):
        t172 = sanitize_exam_copy_text(blocks[172].get("text") or "")
        manual[172] = [("Announcement (English)", t172)]
    if 222 < len(blocks):
        t222 = sanitize_exam_copy_text(blocks[222].get("text") or "")
        manual[222] = [("Isi dokumen (English)", t222)]

    out = build_html(blocks, manual)
    OUT_HTML.write_text(out, encoding="utf-8")
    print("Wrote", OUT_HTML)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
