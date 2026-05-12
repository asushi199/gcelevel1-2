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
TASK_HEAD_RE = re.compile(r"^(Task\s+\d+|TASK\s+\d+)\s*:\s*(.*)$", re.IGNORECASE | re.DOTALL)


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
    return "body"

# block_index -> list of (button_label, exact_string_to_copy)
MANUAL_COPIES_STATIC: dict[int, list[tuple[str, str]]] = {
    42: [("Arahan Gemini", "bagi jawapan tanpa penerangan")],
    79: [("Nama kelas", "Geography.")],
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


def paste_ayat_snippets(text: str) -> list[str]:
    out: list[str] = []
    for m in re.finditer(
        r"Paste\s+ayat\s*[–\u2013\-:]\s*([^\n.]+?)(?:\s+Pilih|\s+pilih|\s+Type|\s+Pada|$)",
        text,
        flags=re.IGNORECASE,
    ):
        s = m.group(1).strip()
        if s and len(s) < 200:
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
                f'text-base font-semibold text-slate-900 sm:text-[17px]">'
                f'<span class="shrink-0 rounded-md bg-slate-800 px-2 py-0.5 text-xs font-bold uppercase tracking-wide '
                f'text-white">{head}</span>'
                f'<span class="min-w-0 flex-1 font-medium leading-snug text-slate-800">{inner_rest}</span></h3>{copy_rows}'
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
        fig_inner: list[str] = []
        for rel in imgs:
            name = Path(rel).name
            src = "media/" + name
            fig_inner.append(
                f'<figure class="not-prose flex min-h-[100px] flex-col overflow-hidden rounded-xl border '
                f'border-slate-200/90 bg-slate-50 shadow-md ring-1 ring-slate-200/50">'
                f'<div class="flex flex-1 items-center justify-center p-3 sm:p-5">'
                f'<img src="{html.escape(src)}" alt="" '
                f'class="max-h-[min(72vh,800px)] w-full object-contain" loading="lazy"></div></figure>'
            )
        if len(fig_inner) == 1:
            inner_parts.append(f'<div class="not-prose my-6">{fig_inner[0]}</div>')
        else:
            inner_parts.append(
                '<div class="not-prose my-6 grid grid-cols-1 gap-5 sm:grid-cols-2 lg:gap-6">'
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

  <div class="mx-auto flex max-w-6xl flex-col gap-6 px-4 py-8 lg:flex-row lg:items-stretch">
    <!-- Mobile TOC -->
    <details class="group rounded-xl border border-slate-200 bg-white p-4 shadow-sm lg:hidden">
      <summary class="cursor-pointer text-sm font-semibold text-slate-800">Isi kandungan</summary>
      <nav class="mt-3 max-h-[50vh] overflow-y-auto text-sm" aria-label="Navigasi">
        {nav_html}
      </nav>
    </details>

    <!-- items-stretch so aside column is as tall as main; inner nav stays sticky while scrolling -->
    <aside class="hidden w-60 shrink-0 lg:block lg:min-h-0" aria-label="Navigasi sisi">
      <nav class="sticky top-20 z-30 max-h-[calc(100vh-5.5rem)] overflow-y-auto rounded-xl border border-slate-200 bg-white p-4 text-sm shadow-sm" id="sidebar-nav">
        <p class="mb-2 text-xs font-bold uppercase tracking-wide text-slate-500">Navigasi</p>
        {nav_html}
      </nav>
    </aside>

    <main class="min-w-0 flex-1 rounded-2xl border border-slate-200/80 bg-white p-6 shadow-md sm:p-10 lg:shadow-lg" id="main-content">
      <div class="not-prose mb-8 rounded-xl border border-slate-200 bg-slate-50/80 px-4 py-3 text-sm leading-relaxed text-slate-600">
        <strong class="text-slate-800">Cara guna:</strong> gunakan butang <span class="rounded bg-white px-1.5 py-0.5 font-medium text-slate-800 ring-1 ring-slate-200">Salin</span> untuk URL dan teks penting.
        Hoskan pada <strong class="text-slate-800">HTTPS</strong> supaya salin ke papan kereta stabil pada semua pelayar.
      </div>
      <div class="prose-flow max-w-none">
      {"".join(body_chunks)}
      </div>
    </main>
  </div>

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
        t95 = blocks[95].get("text") or ""
        manual[95] = [("Instructions (English)", t95)]
    if 172 < len(blocks):
        t172 = blocks[172].get("text") or ""
        manual[172] = [("Announcement (English)", t172)]
    if 222 < len(blocks):
        t222 = blocks[222].get("text") or ""
        manual[222] = [("Isi dokumen (English)", t222)]

    out = build_html(blocks, manual)
    OUT_HTML.write_text(out, encoding="utf-8")
    print("Wrote", OUT_HTML)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
