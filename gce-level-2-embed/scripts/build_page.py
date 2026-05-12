"""Build embeddable index.html for GCE Level 2 from outline.json + docx media."""
from __future__ import annotations

import copy
import html
import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]  # gce-level-2-embed


def _find_site_root() -> Path:
    p = Path(__file__).resolve()
    for k in range(min(10, len(p.parents))):
        root = p.parents[k]
        if (root / "GCE LEVEL 2" / "GCE LEVEL 2.docx").is_file():
            return root
    return ROOT.parent


SITE_ROOT = _find_site_root()
DOCX = SITE_ROOT / "GCE LEVEL 2" / "GCE LEVEL 2.docx"
OUTLINE = Path(__file__).resolve().parent / "outline.json"
OUT_HTML = ROOT / "index.html"
MEDIA_DIR = ROOT / "media"

URL_RE = re.compile(r"https?://[^\s<>\")\]]+")
STEP_NUM_RE = re.compile(r"^(\d+)\.\s*(.*)$", re.DOTALL)
TASK_HEAD_RE = re.compile(
    r"^(TASK|Task)\s+(\d+)(?:\s*:\s*|\s*=\s*|\s+)(.*)$",
    re.IGNORECASE | re.DOTALL,
)

# (salin) — ASCII or fullwidth parentheses
SALIN_MARKER_RE = re.compile(r"（\s*salin\s*）|\(\s*salin\s*\)", re.IGNORECASE)
# (salin untuk semua kandungan) — copy the following consecutive short paragraphs (slide labels, etc.)
SALIN_BULK_MARKER_RE = re.compile(
    r"（\s*salin\s+untuk\s+semua\s+kandungan\s*）|\(\s*salin\s+untuk\s+semua\s+kandungan\s*\)",
    re.IGNORECASE,
)
# Reference-only: (task1a), (gambar task3a), (task2b hingga task2g), etc.
TASK_REF_PAREN_RE = re.compile(
    r"\([^)]*\b(?:gambar\s+)?task\s*\d+\s*[a-z]?[^)]*\)",
    re.IGNORECASE,
)
# Editor-only picture markers (not shown on site)
PICTURE_REF_PAREN_RE = re.compile(r"\(\s*picture\s*\d+\s*\)", re.IGNORECASE)

NavEntry = tuple[str, str, int | None, list[tuple[str, str, int]] | None]


def classify_paragraph(text: str) -> str:
    s = text.strip()
    if not s:
        return "body"
    if re.match(r"^LAB\s+\d", s, re.IGNORECASE):
        return "lab_heading"
    if re.match(r"^GCE\s+LEVEL\s+2", s, re.IGNORECASE) and re.search(r"\bLAB\b", s, re.IGNORECASE):
        return "lab_heading"
    if re.match(r"^Cara\s+end\s+exam\s*$", s, re.IGNORECASE):
        return "section_heading"
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
    if re.match(
        r"^Ambil perhatian\s+semasa\s+log\s+in\s+ke\s+GC",
        s,
        re.IGNORECASE,
    ) and "tengok baik-baik gambar" in s.lower():
        return "gc_login_notice"
    return "body"


def _merge_intervals(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    if not spans:
        return []
    spans = sorted(spans)
    out: list[tuple[int, int]] = []
    for s, e in spans:
        if out and s <= out[-1][1]:
            out[-1] = (out[-1][0], max(out[-1][1], e))
        else:
            out.append((s, e))
    return out


def _remove_regex_spans_from_runs(runs: list[dict], pattern: re.Pattern[str]) -> list[dict]:
    """Remove every match of pattern from flattened runs (char-accurate)."""
    if not runs:
        return []
    flat, meta = _run_flat_meta(runs)
    spans = [m.span() for m in pattern.finditer(flat)]
    if not spans:
        return runs
    merged = _merge_intervals(spans)
    new_chars: list[tuple[str, bool, str | None]] = []
    for i, ch in enumerate(flat):
        if any(s <= i < e for s, e in merged):
            continue
        new_chars.append((ch, meta[i][0], meta[i][1]))
    return _chars_to_runs(new_chars)


def _following_bulk_copy_bundle(blocks: list[dict], idx: int) -> tuple[list[str], list[int]]:
    """Lines after (salin untuk semua kandungan) and their block indices (for skipping duplicate render)."""
    out_lines: list[str] = []
    out_ix: list[int] = []
    j = idx + 1
    while j < len(blocks):
        b = blocks[j]
        if b.get("type") != "p":
            break
        t = (b.get("text") or "").strip()
        if not t:
            break
        tl = t.lower()
        if "(task" in tl:
            break
        if "salin untuk semua kandungan" in tl:
            break
        if re.match(
            r"^(Klik|Pilih|Taip|Masukkan|Buka|Tetapkan|Pastikan|Selepas|Kemudian|Seterusnya|Untuk|Dalam|Guna|Ambil)\b",
            t,
            re.IGNORECASE,
        ):
            break
        if "[" in t and "]" in t and re.search(r"\[(?:Share|Editor|Send|Add)\b", t, re.IGNORECASE):
            break
        if len(t) > 160:
            break
        out_lines.append(t)
        out_ix.append(j)
        j += 1
        if len(out_lines) > 40:
            break
    return out_lines, out_ix


def collect_bulk_salin_skip_indices(blocks: list[dict]) -> set[int]:
    """Paragraph indices already folded into a preceding (salin untuk semua kandungan) clipboard."""
    skip: set[int] = set()
    for i, blk in enumerate(blocks):
        if blk.get("type") != "p":
            continue
        flat0 = "".join(r["t"] for r in (blk.get("runs") or []))
        if SALIN_BULK_MARKER_RE.search(flat0):
            _, ix = _following_bulk_copy_bundle(blocks, i)
            skip.update(ix)
    return skip


def _run_flat_meta(runs: list[dict]) -> tuple[str, list[tuple[bool, str | None]]]:
    chars: list[str] = []
    meta: list[tuple[bool, str | None]] = []
    for r in runs:
        bold = bool(r.get("b"))
        col = r.get("c")
        for ch in r.get("t") or "":
            chars.append(ch)
            meta.append((bold, col))
    return "".join(chars), meta


def _chars_to_runs(chars: list[tuple[str, bool, str | None]]) -> list[dict]:
    if not chars:
        return []
    runs = [{"t": "", "b": chars[0][1], "c": chars[0][2]}]
    for ch, b, c in chars:
        if b == runs[-1]["b"] and c == runs[-1].get("c"):
            runs[-1]["t"] += ch
        else:
            runs.append({"t": ch, "b": b, "c": c})
    return runs


def _find_salin_copy_text(flat: str, meta: list[tuple[bool, str | None]], m_start: int) -> str:
    head = flat[:m_start]

    # "… [Add location] > 123 Main Street (salin)" — copy only the part after the last ">"
    if ">" in head:
        after_gt = head.rsplit(">", 1)[-1].strip()
        after_gt = re.sub(r"\s+", " ", after_gt)
        after_gt = TASK_REF_PAREN_RE.sub("", after_gt)
        after_gt = PICTURE_REF_PAREN_RE.sub("", after_gt)
        after_gt = SALIN_MARKER_RE.sub("", after_gt).strip()
        if after_gt and len(after_gt) < 200:
            return after_gt

    pos = m_start - 1
    while pos >= 0 and flat[pos] in " \t\n\u00a0":
        pos -= 1
    end_pos = pos
    while pos >= 0 and meta[pos][0]:
        pos -= 1
    bold_text = flat[pos + 1 : end_pos + 1].strip()
    if bold_text:
        return bold_text
    # Do not split on ":" or ";" — English lines like "Reminder: …" or "A; B; C."
    # would wrongly keep only the last clause. ">" is handled above.
    segs = re.split(r"[–\u2013\-=]", head)
    if len(segs) >= 2:
        tail = segs[-1].strip()
        tail = re.sub(r"\s+", " ", tail)
        tail = TASK_REF_PAREN_RE.sub("", tail)
        tail = PICTURE_REF_PAREN_RE.sub("", tail)
        tail = SALIN_MARKER_RE.sub("", tail).strip()
        if tail and len(tail) < 200:
            return tail
    tail = head.strip()
    tail = TASK_REF_PAREN_RE.sub("", tail)
    tail = PICTURE_REF_PAREN_RE.sub("", tail)
    tail = SALIN_MARKER_RE.sub("", tail)
    tail = re.sub(r"\s+", " ", tail).strip()
    if tail and len(tail) < 200:
        return tail
    print(f"Warning: (salin) without bold/heuristic: {flat[:100]!r}", file=sys.stderr)
    return ""


def extract_salin_and_strip_reference_parens(runs: list[dict]) -> tuple[list[dict], list[tuple[str, str]]]:
    """Remove (salin) / task refs from runs; return Salin clipboard rows (label, value)."""
    if not runs:
        return [], []
    flat, meta = _run_flat_meta(runs)
    copies: list[tuple[str, str]] = []
    remove_spans: list[tuple[int, int]] = []
    for m in SALIN_MARKER_RE.finditer(flat):
        ct = _find_salin_copy_text(flat, meta, m.start())
        if ct:
            copies.append(("Ayat tampal", ct))
        remove_spans.append(m.span())
    for m in TASK_REF_PAREN_RE.finditer(flat):
        remove_spans.append(m.span())
    for m in PICTURE_REF_PAREN_RE.finditer(flat):
        remove_spans.append(m.span())
    if not remove_spans:
        return merge_adjacent_runs_list(runs), copies
    merged = _merge_intervals(remove_spans)
    new_chars: list[tuple[str, bool, str | None]] = []
    for i, ch in enumerate(flat):
        if any(s <= i < e for s, e in merged):
            continue
        new_chars.append((ch, meta[i][0], meta[i][1]))
    new_runs = _chars_to_runs(new_chars)
    return merge_adjacent_runs_list(new_runs), copies


def find_end_exam_range(blocks: list[dict]) -> tuple[int, int] | None:
    for i, blk in enumerate(blocks):
        t = (blk.get("text") or "").strip().lower()
        if re.match(r"^cara\s+end\s+exam\s*$", t):
            return (i, len(blocks))
    return None


def move_cara_end_exam_to_end(blocks: list[dict]) -> list[dict]:
    """Keep Cara end exam as the closing section (after all labs)."""
    rng = find_end_exam_range(blocks)
    if not rng:
        return blocks
    s, e = rng
    chunk = blocks[s:e]
    rest = blocks[:s] + blocks[e:]
    return rest + chunk


def cara_end_exam_figure_order(imgs: list[str], paragraph_text: str) -> list[str]:
    """Word order was dashboard then breadcrumb; teach breadcrumb (image62) first, then End Exam (image61)."""
    if len(imgs) != 2:
        return imgs
    names = [Path(x).name for x in imgs]
    if set(names) != {"image61.png", "image62.png"}:
        return imgs
    t = (paragraph_text or "").lower()
    if "laman web soalan" not in t:
        return imgs
    by_name = {Path(rel).name: rel for rel in imgs}
    return [by_name["image62.png"], by_name["image61.png"]]


def merge_lab2_task2_invite_images(blocks: list[dict]) -> None:
    """Merge image4–7 + image8–9 into the TASK2 invite paragraph (Word splits figures across empty runs)."""
    for i, blk in enumerate(blocks):
        t = (blk.get("text") or "").strip().lower()
        if "invite mereka ke dalam google classroom" not in t:
            continue
        if "task2b" not in t or "task2g" not in t:
            continue
        merged = list(blk.get("images") or [])
        k = i + 1
        while k < len(blocks):
            nxt = blocks[k]
            if (nxt.get("text") or "").strip():
                break
            extra = list(nxt.get("images") or [])
            if extra:
                merged.extend(extra)
                nxt["images"] = []
            k += 1
        blk["images"] = merged
        return


def lab2_task2_invite_figure_order(imgs: list[str]) -> tuple[list[str], bool]:
    """Badge order 1→6 vs Word export order (layout row1 2|1, row3 6|5). Returns (reordered, show 1–6 badges)."""
    names_set = {Path(x).name for x in imgs}
    want = {f"image{n}.png" for n in range(4, 10)}
    if len(imgs) != 6 or names_set != want:
        return imgs, False
    by_name = {Path(rel).name: rel for rel in imgs}
    ordered = [by_name[f"image{n}.png"] for n in (5, 4, 6, 7, 9, 8)]
    return ordered, True


def lab2_calendar_event_intro_figure_order(imgs: list[str]) -> list[str]:
    """Lab 2 Google Calendar: Word embeds panel 2 (left) then panel 1 (right); teach red badge 1 then 2."""
    if len(imgs) != 2:
        return imgs
    names = {Path(x).name for x in imgs}
    if names != {"image27.png", "image28.png"}:
        return imgs
    by_name = {Path(rel).name: rel for rel in imgs}
    return [by_name["image28.png"], by_name["image27.png"]]


def lab2_calendar_event_intro_step_badges(imgs: list[str]) -> bool:
    return len(imgs) == 2 and {Path(x).name for x in imgs} == {"image27.png", "image28.png"}


def lab2_task2ef_slide_placeholder_figure_order(imgs: list[str]) -> list[str]:
    """Slides Image placeholder steps: Word order was Rename (2) then menu (1); teach 1 then 2 left-to-right."""
    if len(imgs) != 2:
        return imgs
    names = {Path(x).name for x in imgs}
    if names != {"image46.png", "image47.png"}:
        return imgs
    by_name = {Path(rel).name: rel for rel in imgs}
    return [by_name["image47.png"], by_name["image46.png"]]


def build_nav_l2(blocks: list[dict]) -> list[NavEntry]:
    out: list[NavEntry] = []
    cara_idx: int | None = None
    for i, blk in enumerate(blocks):
        if (blk.get("text") or "").strip().lower() == "cara end exam":
            cara_idx = i
            break
    lab_count = 0
    children: list[tuple[str, str, int]] | None = None
    for i, blk in enumerate(blocks):
        t = (blk.get("text") or "").strip()
        if re.match(r"^GCE\s+LEVEL\s+2", t, re.IGNORECASE) and re.search(r"\bLAB\b", t, re.IGNORECASE):
            lab_count += 1
            sid = f"lab-{lab_count}"
            children = []
            out.append((sid, t[:88], i, children))
        elif children is not None and re.match(r"^(TASK|Task)\s+\d+", t):
            m = re.match(r"^(TASK|Task)\s+(\d+)", t, re.IGNORECASE)
            assert m is not None
            tid = f"lab-{lab_count}-task-{m.group(2)}"
            children.append((tid, t[:72], i))
    if cara_idx is not None:
        out.append(("cara-end-exam", "Cara end exam", cara_idx, None))
    return out


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


def build_extra_copies_for_block(
    _idx: int, _text: str, manual: dict[int, list[tuple[str, str]]]
) -> list[tuple[str, str]]:
    return list(manual.get(_idx, []))


def split_urls(s: str) -> list[tuple[str, str | None]]:
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


def _norm_ws_one_line(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


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
            tag = f"{m.group(1)} {m.group(2)}".strip()
            rest = (m.group(3) or "").strip()
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
            body = body_inner if body_inner is not None else render_text_with_urls(rest)
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
    return (
        f'<p class="not-prose mb-3.5 text-[15px] leading-7 text-slate-700 last:mb-0">{inner_rich}</p>{copy_rows}'
    )


def render_block(
    idx: int,
    block: dict,
    blocks: list[dict],
    anchor_ids: set[str],
    manual: dict[int, list[tuple[str, str]]],
    bulk_skip: set[int],
) -> str:
    if idx in bulk_skip:
        return ""
    blk = copy.deepcopy(block)
    runs_in = list(blk.get("runs") or [])
    flat0 = "".join(r["t"] for r in runs_in)
    bulk_rows: list[tuple[str, str]] = []
    if SALIN_BULK_MARKER_RE.search(flat0):
        bulk_lines, _ = _following_bulk_copy_bundle(blocks, idx)
        if bulk_lines:
            bullet = "\u2022 "  # plain-text bullet for paste into slides
            bulk_rows.append(
                ("Ayat tampal", "\n".join(f"{bullet}{ln.strip()}" for ln in bulk_lines)),
            )
        runs_in = _remove_regex_spans_from_runs(runs_in, SALIN_BULK_MARKER_RE)
        runs_in = merge_adjacent_runs_list(runs_in)
    runs_proc, salin_rows = extract_salin_and_strip_reference_parens(runs_in)
    salin_rows = bulk_rows + salin_rows
    blk["runs"] = runs_proc
    blk["text"] = "".join(r["t"] for r in runs_proc).strip()

    t = blk.get("text") or ""
    imgs = list(blk.get("images") or [])
    imgs = cara_end_exam_figure_order(imgs, t)
    imgs, invite_step_badges = lab2_task2_invite_figure_order(imgs)
    imgs = lab2_calendar_event_intro_figure_order(imgs)
    imgs = lab2_task2ef_slide_placeholder_figure_order(imgs)
    calendar_intro_badges = lab2_calendar_event_intro_step_badges(imgs)
    rows = blk.get("rows")
    bid = f"b-{idx}"

    if blk.get("type") == "tbl" and rows:
        id_attr = f' id="{bid}"' if bid in anchor_ids else ""
        scroll = " scroll-mt-28" if id_attr else ""
        parts = [
            f'<div class="my-4 overflow-x-auto{scroll}"{id_attr}>',
            '<table class="min-w-full border border-slate-200 text-sm">',
        ]
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
        runs = blk.get("runs") or []
        flat = "".join(r["t"] for r in runs)
        st = flat.strip()
        text_for_kind = st if st else t.strip()
        kind = classify_paragraph(text_for_kind)
        extras = list(salin_rows)
        extras.extend(build_extra_copies_for_block(idx, t, manual))
        copy_rows = "".join(render_copy_row(lbl, val) for lbl, val in extras if val)

        copy_vals = [v for _lb, v in extras if v]
        skip_dup_body = (
            kind == "body"
            and len(copy_vals) == 1
            and _norm_ws_one_line(t) == _norm_ws_one_line(copy_vals[0])
        )

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
                    pre_len = len(m.group(0)) - len(m.group(3))
                    task_inner = render_runs_to_html(slice_runs(runs, loff + pre_len, hi))
        else:
            inner_rich = render_text_with_urls(t.strip())

        if not skip_dup_body:
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
        elif copy_rows:
            inner_parts.append(f'<div class="not-prose mb-3.5 last:mb-0">{copy_rows}</div>')

    if imgs:
        fig_inner: list[str] = []
        for step_i, rel in enumerate(imgs):
            name = Path(rel).name
            src = "media/" + name
            badge = ""
            if invite_step_badges or calendar_intro_badges:
                n = str(step_i + 1)
                badge = (
                    f'<span class="pointer-events-none absolute left-2 top-2 z-10 flex h-8 w-8 items-center '
                    f'justify-center rounded-full bg-rose-600 text-sm font-bold text-white shadow-md ring-2 ring-white" '
                    f'aria-hidden="true">{html.escape(n)}</span>'
                )
            fig_inner.append(
                f'<div class="relative">'
                f"{badge}"
                f'<figure class="content-figure not-prose flex min-h-[80px] flex-col overflow-hidden rounded-xl border '
                f'border-slate-200/90 bg-slate-50 shadow-md ring-1 ring-slate-200/50">'
                f'<div class="flex flex-1 items-center justify-center p-2 sm:p-4 lg:p-3">'
                f'<img src="{html.escape(src)}" alt="" '
                f'class="content-figure-img w-full cursor-zoom-in rounded-md object-contain transition hover:opacity-95 '
                f'max-h-[min(68vh,520px)] lg:max-h-[min(48vh,540px)]" '
                f'loading="lazy" decoding="async"></div></figure></div>'
            )
        if len(fig_inner) == 1:
            inner_parts.append(f'<div class="not-prose my-6 lg:mx-auto lg:max-w-[42rem]">{fig_inner[0]}</div>')
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


def collect_anchor_ids(nav: list[NavEntry]) -> set[str]:
    ids: set[str] = set()
    for _sid, _label, main_idx, children in nav:
        if main_idx is not None:
            ids.add(f"b-{main_idx}")
        if children:
            for _cid, _cl, cidx in children:
                ids.add(f"b-{cidx}")
    return ids


def render_nav_html(nav: list[NavEntry]) -> str:
    lines: list[str] = []
    for sid, label, main_idx, children in nav:
        lines.append('<div class="mb-4 border-b border-slate-100 pb-3 last:mb-0 last:border-0 last:pb-0">')
        if main_idx is None:
            lines.append(
                f'<span class="block px-2 py-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">'
                f"{html.escape(label)}</span>"
            )
        else:
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


def _lab_section_shell_open(li: int, aria: str) -> str:
    esc = html.escape(aria)
    shells = [
        (
            '<section class="lab-section mx-auto mb-12 max-w-[52rem] scroll-mt-24 rounded-2xl border-2 border-indigo-100 '
            'bg-gradient-to-b from-indigo-50/40 to-white p-6 shadow-sm sm:p-8" '
            f'aria-label="{esc}">'
            '<header class="not-prose mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-indigo-700">'
            '<span class="h-1.5 w-1.5 rounded-full bg-indigo-500"></span>'
            f'Praktikal · {esc}</header>'
        ),
        (
            '<section class="lab-section mx-auto mb-12 max-w-[52rem] scroll-mt-24 rounded-2xl border-2 border-emerald-100 '
            'bg-gradient-to-b from-emerald-50/40 to-white p-6 shadow-sm sm:p-8" '
            f'aria-label="{esc}">'
            '<header class="not-prose mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-emerald-800">'
            '<span class="h-1.5 w-1.5 rounded-full bg-emerald-500"></span>'
            f'Praktikal · {esc}</header>'
        ),
        (
            '<section class="lab-section mx-auto mb-8 max-w-[52rem] scroll-mt-24 rounded-2xl border-2 border-amber-100 '
            'bg-gradient-to-b from-amber-50/35 to-white p-6 shadow-sm sm:p-8" '
            f'aria-label="{esc}">'
            '<header class="not-prose mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-amber-900">'
            '<span class="h-1.5 w-1.5 rounded-full bg-amber-500"></span>'
            f'Praktikal · {esc}</header>'
        ),
    ]
    return shells[li % len(shells)]


def build_html(blocks: list[dict], manual: dict[int, list[tuple[str, str]]], nav: list[NavEntry]) -> str:
    anchor_ids = collect_anchor_ids(nav)
    bulk_skip = collect_bulk_salin_skip_indices(blocks)
    nav_html = render_nav_html(nav)

    lab_starts: list[int] = []
    for i, blk in enumerate(blocks):
        t = (blk.get("text") or "").strip()
        if re.match(r"^GCE\s+LEVEL\s+2", t, re.IGNORECASE) and re.search(r"\bLAB\b", t, re.IGNORECASE):
            lab_starts.append(i)

    cara_start: int | None = None
    for i, blk in enumerate(blocks):
        if (blk.get("text") or "").strip().lower() == "cara end exam":
            cara_start = i
            break

    body_chunks: list[str] = []
    if not lab_starts:
        body_chunks.append(
            '<div class="doc-shell mx-auto mb-10 max-w-[52rem] scroll-mt-28 rounded-2xl border border-slate-200/90 '
            'bg-white p-6 shadow-sm sm:p-8">'
        )
        for i in range(len(blocks)):
            ch = render_block(i, blocks[i], blocks, anchor_ids, manual, bulk_skip)
            if ch:
                body_chunks.append(ch)
        body_chunks.append("</div>")
    else:
        for li, start in enumerate(lab_starts):
            next_start = lab_starts[li + 1] if li + 1 < len(lab_starts) else len(blocks)
            end = next_start
            if cara_start is not None and start <= cara_start < end:
                end = cara_start
            title = (blocks[start].get("text") or f"Lab {li + 1}").strip()
            short = re.sub(r"^GCE\s+LEVEL\s+2\s*[–-]\s*", "", title, flags=re.IGNORECASE).strip() or title
            body_chunks.append(_lab_section_shell_open(li, short))
            for i in range(start, end):
                ch = render_block(i, blocks[i], blocks, anchor_ids, manual, bulk_skip)
                if ch:
                    body_chunks.append(ch)
            body_chunks.append("</section>")
        if cara_start is not None:
            body_chunks.append(
                '<section class="lab-section mx-auto mb-8 max-w-[52rem] scroll-mt-24 rounded-2xl border-2 border-slate-200 '
                'bg-gradient-to-b from-slate-50/90 to-white p-6 shadow-sm sm:p-8" aria-label="Cara end exam">'
                '<header class="not-prose mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-slate-600">'
                '<span class="h-1.5 w-1.5 rounded-full bg-slate-400"></span>Tamat peperiksaan</header>'
            )
            for i in range(cara_start, len(blocks)):
                ch = render_block(i, blocks[i], blocks, anchor_ids, manual, bulk_skip)
                if ch:
                    body_chunks.append(ch)
            body_chunks.append("</section>")

    return f"""<!DOCTYPE html>
<html lang="ms">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>GCE Level 2 — Panduan Lab</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="min-h-screen bg-slate-100 text-slate-900 antialiased">
  <header class="sticky top-0 z-40 border-b border-slate-200 bg-white/95 backdrop-blur">
    <div class="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-3 px-4 py-3">
      <div class="flex min-w-0 flex-wrap items-center gap-2 sm:gap-3">
        <a href="../levels.html" class="shrink-0 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm font-medium text-slate-700 shadow-sm hover:bg-slate-50">← Senarai</a>
        <h1 class="text-lg font-semibold tracking-tight text-slate-900">GCE Level 2</h1>
      </div>
      <p class="hidden text-sm text-slate-500 sm:block">Panduan langkah demi langkah · salin teks dengan selamat</p>
    </div>
  </header>

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
    blocks: list[dict] = json.loads(OUTLINE.read_text(encoding="utf-8"))
    blocks = move_cara_end_exam_to_end(blocks)
    merge_lab2_task2_invite_images(blocks)
    nav = build_nav_l2(blocks)
    manual: dict[int, list[tuple[str, str]]] = {}
    out = build_html(blocks, manual, nav)
    OUT_HTML.write_text(out, encoding="utf-8")
    print("Wrote", OUT_HTML)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
