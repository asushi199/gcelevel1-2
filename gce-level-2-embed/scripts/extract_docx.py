"""Extract ordered blocks from GCE LEVEL 2.docx (same schema as Level 1)."""
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


BLIP = f"{{{A}}}blip"
EMBED = f"{{{R_REL}}}embed"


def local_tag(elem: ET.Element) -> str:
    if not elem.tag or "}" not in elem.tag:
        return elem.tag or ""
    return elem.tag.split("}", 1)[1]


def attr_val(elem: ET.Element, local: str) -> str | None:
    if elem is None:
        return None
    full = f"{{{W}}}{local}"
    return elem.get(full) or elem.get(local)


def is_on(elem: ET.Element | None) -> bool:
    if elem is None:
        return False
    val = attr_val(elem, "val")
    if val is None:
        return True
    return val.lower() not in ("0", "false", "off", "none")


def run_text_and_props(r: ET.Element) -> tuple[str, bool, str | None]:
    parts: list[str] = []
    for child in r:
        lt = local_tag(child)
        if lt == "t":
            parts.append(child.text or "")
        elif lt == "tab":
            parts.append("\t")
        elif lt == "br":
            parts.append("\n")
        elif lt == "noBreakHyphen":
            parts.append("\u2011")
        elif lt == "softHyphen":
            parts.append("\u00ad")
    text = "".join(parts)

    rpr = r.find(w("rPr"))
    bold = False
    color_hex: str | None = None
    if rpr is not None:
        if is_on(rpr.find(w("b"))) or is_on(rpr.find(w("bCs"))):
            bold = True
        c_el = rpr.find(w("color"))
        if c_el is not None:
            raw = attr_val(c_el, "val")
            if raw and raw.lower() != "auto":
                raw = raw.upper().replace("#", "")
                if len(raw) == 8 and all(c in "0123456789ABCDEF" for c in raw):
                    color_hex = raw[2:8]
                elif len(raw) == 6 and all(c in "0123456789ABCDEF" for c in raw):
                    color_hex = raw
    return text, bold, color_hex


def iter_runs_in_paragraph(p: ET.Element):

    def walk(el: ET.Element):
        for child in el:
            lt = local_tag(child)
            if lt == "r":
                yield child
            elif lt in (
                "hyperlink",
                "ins",
                "del",
                "moveFrom",
                "moveTo",
                "smartTag",
                "sdt",
                "dir",
                "fldSimple",
            ):
                yield from walk(child)

    yield from walk(p)


def run_has_drawing(r: ET.Element) -> bool:
    for d in r.iter():
        if local_tag(d) in ("drawing", "pict"):
            return True
    return False


def run_has_text(r: ET.Element) -> bool:
    return any(local_tag(c) == "t" for c in r.iter())


def merge_adjacent_runs(runs: list[dict]) -> list[dict]:
    if not runs:
        return []
    out = [dict(runs[0])]
    for r in runs[1:]:
        if r["b"] == out[-1]["b"] and r.get("c") == out[-1].get("c"):
            out[-1]["t"] += r["t"]
        else:
            out.append(dict(r))
    return out


def paragraph_to_block(p: ET.Element, rid_to_target: dict[str, str]) -> dict:
    runs_out: list[dict] = []

    for r in iter_runs_in_paragraph(p):
        if run_has_drawing(r) and not run_has_text(r):
            continue
        t, bold, col = run_text_and_props(r)
        if t == "" and not bold and col is None:
            continue
        runs_out.append({"t": t, "b": bold, "c": col})

    runs_out = merge_adjacent_runs(runs_out)
    joined = "".join(r["t"] for r in runs_out)
    para_text = joined.strip()
    imgs: list[str] = []
    for blip in p.iter(BLIP):
        rid = blip.get(EMBED)
        if rid and rid in rid_to_target:
            imgs.append(rid_to_target[rid])

    if not runs_out and para_text:
        runs_out = [{"t": para_text, "b": False, "c": None}]

    return {"type": "p", "text": para_text, "runs": runs_out, "images": imgs}


def _find_site_root() -> Path:
    p = Path(__file__).resolve()
    for k in range(min(10, len(p.parents))):
        root = p.parents[k]
        if (root / "GCE LEVEL 2" / "GCE LEVEL 2.docx").is_file():
            return root
    return p.parents[2]


def main() -> int:
    site_root = _find_site_root()
    doc_path = site_root / "GCE LEVEL 2" / "GCE LEVEL 2.docx"
    out_path = Path(__file__).resolve().parent / "outline.json"

    if not doc_path.is_file():
        print("Missing docx:", doc_path, file=sys.stderr)
        return 1

    with zipfile.ZipFile(doc_path, "r") as zf:
        rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        rid_to_target: dict[str, str] = {}
        for rel in rels_root:
            rid = rel.get("Id")
            tgt = rel.get("Target")
            if rid and tgt:
                rid_to_target[rid] = tgt

        doc_root = ET.fromstring(zf.read("word/document.xml"))
        body = doc_root.find(w("body"))
        if body is None:
            print("No body", file=sys.stderr)
            return 1

        blocks: list[dict] = []
        for child in list(body):
            tag = local_tag(child)
            if tag == "p":
                blocks.append(paragraph_to_block(child, rid_to_target))
            elif tag == "tbl":
                rows_text: list[list[str]] = []
                for tr in child.iter(w("tr")):
                    row: list[str] = []
                    for tc in tr.iter(w("tc")):
                        cell_parts: list[str] = []
                        for node in tc.iter(w("t")):
                            if node.text:
                                cell_parts.append(node.text)
                            if node.tail:
                                cell_parts.append(node.tail)
                        row.append("".join(cell_parts).strip())
                    if any(row):
                        rows_text.append(row)
                blocks.append({"type": "tbl", "rows": rows_text, "images": [], "runs": [], "text": ""})

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(blocks, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Wrote", out_path, "blocks", len(blocks))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
