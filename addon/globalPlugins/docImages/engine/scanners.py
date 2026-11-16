"""Format scanners: list the images of a document (with alt text and location) without extracting them.
Pure stdlib, no NVDA imports."""
import posixpath
import xml.etree.ElementTree as ET

from .model import DocImage, DocImageError

MAX_IMAGE_BYTES = 100 * 1024 * 1024
MAX_TOTAL_BYTES = 500 * 1024 * 1024

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
XLINK = "http://www.w3.org/1999/xlink"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
RD = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
RVR = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"

T_P, T_T, T_TR = f"{{{W}}}p", f"{{{W}}}t", f"{{{W}}}tr"
T_BR, T_LRPB = f"{{{W}}}br", f"{{{W}}}lastRenderedPageBreak"
T_PPR, T_SECT, T_PBB = f"{{{W}}}pPr", f"{{{W}}}sectPr", f"{{{W}}}pageBreakBefore"
T_DRAWING = f"{{{W}}}drawing"
T_TXBX = f"{{{W}}}txbxContent"
T_BLIP = f"{{{A}}}blip"
T_CHART = f"{{{C}}}chart"
T_DOCPR = f"{{{WP}}}docPr"
T_IMAGEDATA = f"{{{V}}}imagedata"
T_FALLBACK = f"{{{MC}}}Fallback"
A_EMBED, A_RID = f"{{{R}}}embed", f"{{{R}}}id"
A_OTITLE = f"{{{O}}}title"
A_TYPE, A_VAL = f"{{{W}}}type", f"{{{W}}}val"
A_RELID = f"{{{O}}}relid"
T_VML_SHAPE = f"{{{V}}}shape"
HF_CODES = ("LH", "CH", "RH", "LF", "CF", "RF")

# Word counts the end-of-row mark of every table row as a paragraph in Document.Paragraphs.
COUNT_ROW_END_AS_PARAGRAPH = True
PART_ORDER = ("header", "footer", "footnotes", "endnotes")


class Budget:
    """Guards against zip bombs: per-image and total size limits (from the zip directory, no decompression)."""

    def __init__(self, z):
        self.z = z
        self.total = 0
        self.seen = set()
        self.charts = 0  # native charts found while scanning (they are not pictures)
        self.chart_items = []  # xlsx/pptx charts as list entries (they can be exported through the live app)

    def ok(self, member):
        if member in self.seen:  # the same media file used again costs nothing
            return True
        try:
            size = self.z.getinfo(member).file_size
        except KeyError:
            return False
        if size > MAX_IMAGE_BYTES or self.total + size > MAX_TOTAL_BYTES:
            return False
        self.total += size
        self.seen.add(member)
        return True


# ---------------------------------------------------------------- relationships
def _rels_path(member):
    d, b = posixpath.split(member)
    return posixpath.join(d, "_rels", b + ".rels")


def _all_rels(z, names, member):
    """[(rId, type suffix such as 'image', resolved zip member)] for internal relationships of a part."""
    relpath = _rels_path(member)
    if relpath not in names:
        return []
    base = posixpath.dirname(member)
    out = []
    for r in ET.fromstring(z.read(relpath)).findall(f"{{{PKG_REL}}}Relationship"):
        if r.get("TargetMode") == "External":
            continue
        t = r.get("Target") or ""
        target = t.lstrip("/") if t.startswith("/") else posixpath.join(base, t)
        out.append((r.get("Id"), (r.get("Type") or "").rsplit("/", 1)[-1], posixpath.normpath(target)))
    return out


def _image_rels(z, names, member):
    return {i: m for i, k, m in _all_rels(z, names, member) if k == "image"}


def _walk(elem, skip=(), stop=()):
    for child in elem:
        if child.tag in skip:
            continue
        yield child
        if child.tag not in stop:
            yield from _walk(child, skip, stop)


def _clean(text, limit=120):
    return " ".join((text or "").split())[:limit]


# ---------------------------------------------------------------- DOCX
def _is_on(el):
    return el is not None and el.get(A_VAL, "true").lower() not in ("0", "false", "off")


class _Scanner:
    def __init__(self, has_lrpb):
        self.has_lrpb = has_lrpb
        self.main = 0
        self.page = 1
        self.hits = []        # pictures: (paragraph, context, page_est, rid, alt, in_textbox, slot, name)
        self.chart_hits = []  # charts, same tuple
        self.slots = {}       # paragraph -> how many picture-like objects it holds so far

    def scan(self, node, anchor):
        for child in node:
            tag = child.tag
            if tag == T_FALLBACK:
                continue
            if tag == T_P:
                self.para(child, anchor)
            elif tag == T_TR:
                self.scan(child, anchor)
                if anchor is None and COUNT_ROW_END_AS_PARAGRAPH:
                    self.main += 1
            else:
                self.scan(child, anchor)

    def para(self, p, anchor):
        in_tx = anchor is not None
        if not in_tx:
            self.main += 1
        idx = anchor if in_tx else self.main
        ppr = p.find(T_PPR)
        if not in_tx and not self.has_lrpb and ppr is not None and _is_on(ppr.find(T_PBB)):
            self.page += 1
        text, found, boxes = [], [], []
        for e in _walk(p, {T_FALLBACK}, {T_TXBX}):
            tag = e.tag
            if tag == T_T:
                text.append(e.text or "")
            elif tag == T_LRPB:
                if self.has_lrpb and not in_tx:
                    self.page += 1
            elif tag == T_BR:
                if e.get(A_TYPE) == "page" and not self.has_lrpb and not in_tx:
                    self.page += 1
            elif tag == T_DRAWING:
                docpr = next((d for d in _walk(e, {T_FALLBACK}, {T_TXBX}) if d.tag == T_DOCPR), None)
                alt = (docpr.get("descr") or docpr.get("title") or "") if docpr is not None else ""
                name = (docpr.get("name") or "") if docpr is not None else ""
                for b in _walk(e, {T_FALLBACK}, {T_TXBX}):
                    if b.tag == T_BLIP and b.get(A_EMBED):
                        found.append(("pic", b.get(A_EMBED), alt, self.page, name))
                    elif b.tag == T_CHART and b.get(A_RID):
                        found.append(("chart", b.get(A_RID), alt, self.page, name))
            elif tag == T_IMAGEDATA and e.get(A_RID):
                found.append(("pic", e.get(A_RID), e.get(A_OTITLE) or "", self.page, ""))
            elif tag == T_TXBX:
                boxes.append(e)
        ctx = _clean("".join(text))
        for kind, rid, alt, page, name in found:
            slot = self.slots.get(idx, 0)
            self.slots[idx] = slot + 1
            (self.chart_hits if kind == "chart" else self.hits).append((idx, ctx, page, rid, alt, in_tx, slot, name))
        for box in boxes:
            self.scan(box, idx)
        if not in_tx and not self.has_lrpb and ppr is not None:
            sect = ppr.find(T_SECT)
            if sect is not None:
                t = sect.find(f"{{{W}}}type")
                if t is None or t.get(A_VAL) != "continuous":
                    self.page += 1


def scan_docx(z, names, budget):
    if "word/document.xml" not in names:
        raise DocImageError("nocontent")
    images = []
    root = ET.fromstring(z.read("word/document.xml"))
    rels = _image_rels(z, names, "word/document.xml")
    sc = _Scanner(any(True for _ in root.iter(T_LRPB)))
    sc.scan(root, None)
    budget.charts += len(sc.chart_hits)
    for idx, ctx, page, rid, alt, in_tx, slot, _name in sc.hits:
        member = rels.get(rid)
        if member and budget.ok(member):
            images.append(DocImage(len(images) + 1, member, "docx", "body", idx, alt.strip(), ctx,
                                   in_tx, page_est=page, slot=slot))
    crels = {i: m for i, k, m in _all_rels(z, names, "word/document.xml") if k == "chart"}
    for idx, ctx, page, rid, alt, in_tx, slot, name in sc.chart_hits:
        member = crels.get(rid, "")
        budget.chart_items.append(DocImage(
            len(budget.chart_items) + 1, member, "docx", "body", idx, alt.strip(), ctx, in_tx, kind="chart",
            title=_chart_title(z, names, member), shape_name=name, page_est=page, slot=slot))
    parts, seen = [], set()
    for _id, kind, member in _all_rels(z, names, "word/document.xml"):
        if kind in PART_ORDER and member in names and member not in seen:
            seen.add(member)
            parts.append((PART_ORDER.index(kind), member, kind))
    for _order, member, kind in sorted(parts):
        proot = ET.fromstring(z.read(member))
        prels = _image_rels(z, names, member)
        ps = _Scanner(False)
        ps.scan(proot, None)
        for _idx, ctx, _page, rid, alt, in_tx, _slot, _name in ps.hits:
            target = prels.get(rid)
            if target and budget.ok(target):
                images.append(DocImage(len(images) + 1, target, "docx", kind, 0, alt.strip(), ctx, in_tx))
    return images


def _chart_title(z, names, member):
    """Title text of a chart part, or an empty string."""
    if not member or member not in names:
        return ""
    try:
        root = ET.fromstring(z.read(member))
    except ET.ParseError:
        return ""
    chart = root.find(f"{{{C}}}chart")
    title = chart.find(f"{{{C}}}title") if chart is not None else None
    if title is None:
        return ""
    return _clean(" ".join(t.text or "" for t in title.iter(f"{{{A}}}t")))


# ---------------------------------------------------------------- PPTX
def _shape_text(sp):
    return _clean(" ".join(t.text or "" for t in sp.iter(f"{{{A}}}t")))


def _slide_title(root):
    first = ""
    for sp in root.iter(f"{{{P}}}sp"):
        txt = _shape_text(sp)
        if not txt:
            continue
        ph = sp.find(f".//{{{P}}}nvPr/{{{P}}}ph")
        if ph is not None and ph.get("type") in ("title", "ctrTitle"):
            return txt
        first = first or txt
    return first


def scan_pptx(z, names, budget):
    prs = "ppt/presentation.xml"
    if prs not in names:
        raise DocImageError("nocontent")
    root = ET.fromstring(z.read(prs))
    rels = {i: m for i, _k, m in _all_rels(z, names, prs)}
    slides = [rels.get(s.get(A_RID)) for s in root.iter(f"{{{P}}}sldId")]
    images, number = [], 0
    for member in slides:
        if not member or member not in names:
            continue
        number += 1
        sroot = ET.fromstring(z.read(member))
        srels = _all_rels(z, names, member)
        irels = {i: m for i, k, m in srels if k == "image"}
        crels = {i: m for i, k, m in srels if k == "chart"}
        title = _slide_title(sroot)
        budget.charts += sum(1 for _ in sroot.iter(f"{{{C}}}chart"))
        for pic in sroot.iter(f"{{{P}}}pic"):
            cnv = next(pic.iter(f"{{{P}}}cNvPr"), None)
            alt = (cnv.get("descr") or cnv.get("title") or "") if cnv is not None else ""
            blip = next(pic.iter(T_BLIP), None)
            target = irels.get(blip.get(A_EMBED)) if blip is not None else None
            if target and budget.ok(target):
                images.append(DocImage(len(images) + 1, target, "pptx", alt=alt.strip(), context=title,
                                       slide=number, shape_name=(cnv.get("name") or "") if cnv is not None else ""))
        for frame in sroot.iter(f"{{{P}}}graphicFrame"):
            chart = next(frame.iter(f"{{{C}}}chart"), None)
            if chart is None:
                continue
            cnv = next(frame.iter(f"{{{P}}}cNvPr"), None)
            member_chart = crels.get(chart.get(A_RID), "")
            ctitle = _chart_title(z, names, member_chart)
            budget.chart_items.append(DocImage(
                len(budget.chart_items) + 1, member_chart, "pptx", kind="chart", slide=number,
                alt=((cnv.get("descr") or cnv.get("title") or "") if cnv is not None else "").strip(),
                context=ctitle, title=ctitle,
                shape_name=(cnv.get("name") or "") if cnv is not None else ""))
    return images


# ---------------------------------------------------------------- XLSX
def _col_letters(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _rich_cell_images(z, names):
    """{vm index: (media member, alt text)} for pictures placed in cells ("Place in Cell").
    A cell points to value metadata (vm) -> future metadata -> rich value -> structure keys -> relationship."""
    needed = ("xl/metadata.xml", "xl/richData/rdrichvalue.xml",
              "xl/richData/rdrichvaluestructure.xml", "xl/richData/richValueRel.xml")
    if any(n not in names for n in needed):
        return {}
    try:
        meta = ET.fromstring(z.read("xl/metadata.xml"))
        types = [t.get("name") for t in meta.iter(f"{{{S}}}metadataType")]
        rvb = []
        for fm in meta.iter(f"{{{S}}}futureMetadata"):
            if fm.get("name") == "XLRICHVALUE":
                for bk in fm.findall(f"{{{S}}}bk"):
                    node = next(bk.iter(f"{{{RD}}}rvb"), None)
                    rvb.append(int(node.get("i")) if node is not None else -1)
        vm_to_rv = {}
        vmeta = meta.find(f"{{{S}}}valueMetadata")
        for n, bk in enumerate(vmeta.findall(f"{{{S}}}bk") if vmeta is not None else [], 1):
            for rc in bk.findall(f"{{{S}}}rc"):
                t, v = int(rc.get("t")) - 1, int(rc.get("v"))
                if 0 <= t < len(types) and types[t] == "XLRICHVALUE" and 0 <= v < len(rvb) and rvb[v] >= 0:
                    vm_to_rv[n] = rvb[v]
        structs = [[k.get("n") for k in s.findall(f"{{{RD}}}k")]
                   for s in ET.fromstring(z.read("xl/richData/rdrichvaluestructure.xml")).findall(f"{{{RD}}}s")]
        rel_ids = [r.get(A_RID) for r in
                   ET.fromstring(z.read("xl/richData/richValueRel.xml")).findall(f"{{{RVR}}}rel")]
        rel_members = {i: m for i, k, m in _all_rels(z, names, "xl/richData/richValueRel.xml") if k == "image"}
        by_rv = {}
        for idx, rv in enumerate(ET.fromstring(z.read("xl/richData/rdrichvalue.xml")).findall(f"{{{RD}}}rv")):
            keys = structs[int(rv.get("s"))]
            vals = [v.text or "" for v in rv.findall(f"{{{RD}}}v")]
            if "_rvRel:LocalImageIdentifier" not in keys:
                continue
            member = rel_members.get(rel_ids[int(vals[keys.index("_rvRel:LocalImageIdentifier")])])
            alt = ""
            if "Text" in keys and keys.index("Text") < len(vals):
                alt = vals[keys.index("Text")]
            if member:
                by_rv[idx] = (member, alt)
        return {vm: by_rv[rv] for vm, rv in vm_to_rv.items() if rv in by_rv}
    except (ET.ParseError, ValueError, IndexError, KeyError, TypeError):
        return {}


def _cells_with_vm(z, member):
    """[(cell address, vm)] for cells that carry value metadata. Streams the sheet, which can be huge."""
    out = []
    cell_tag = f"{{{S}}}c"
    with z.open(member) as f:
        for _event, el in ET.iterparse(f, events=("end",)):
            if el.tag == cell_tag:
                vm = el.get("vm")
                if vm:
                    out.append((el.get("r") or "", int(vm)))
                el.clear()
    return out


def scan_xlsx(z, names, budget):
    wb = "xl/workbook.xml"
    if wb not in names:
        raise DocImageError("nocontent")
    root = ET.fromstring(z.read(wb))
    wrels = {i: m for i, _k, m in _all_rels(z, names, wb)}
    rich = _rich_cell_images(z, names)
    images = []

    def add(member, sname, alt="", **extra):
        if member and budget.ok(member):
            images.append(DocImage(len(images) + 1, member, "xlsx", alt=(alt or "").strip(), sheet=sname, **extra))

    for sheet in root.iter(f"{{{S}}}sheet"):
        sname = sheet.get("name") or ""
        smember = wrels.get(sheet.get(A_RID))
        if not smember or smember not in names:
            continue
        srels = _all_rels(z, names, smember)
        # 1. pictures floating over the cells
        for _i, kind, dmember in srels:
            if kind != "drawing" or dmember not in names:
                continue
            droot = ET.fromstring(z.read(dmember))
            drels = _all_rels(z, names, dmember)
            irels = {i: m for i, k, m in drels if k == "image"}
            crels = {i: m for i, k, m in drels if k == "chart"}
            budget.charts += sum(1 for _ in droot.iter(f"{{{C}}}chart"))
            for anchor in droot:
                frm = anchor.find(f"{{{XDR}}}from")
                cell = ""
                if frm is not None:
                    try:
                        col = int(frm.find(f"{{{XDR}}}col").text)
                        row = int(frm.find(f"{{{XDR}}}row").text)
                        cell = f"{_col_letters(col)}{row + 1}"
                    except (AttributeError, TypeError, ValueError):
                        pass
                for pic in anchor.iter(f"{{{XDR}}}pic"):
                    cnv = next(pic.iter(f"{{{XDR}}}cNvPr"), None)
                    alt = (cnv.get("descr") or cnv.get("title") or "") if cnv is not None else ""
                    blip = next(pic.iter(T_BLIP), None)
                    add(irels.get(blip.get(A_EMBED)) if blip is not None else None, sname, alt, cell=cell,
                        shape_name=(cnv.get("name") or "") if cnv is not None else "")
                for frame in anchor.iter(f"{{{XDR}}}graphicFrame"):
                    chart = next(frame.iter(f"{{{C}}}chart"), None)
                    if chart is None:
                        continue
                    cnv = next(frame.iter(f"{{{XDR}}}cNvPr"), None)
                    member_chart = crels.get(chart.get(A_RID), "")
                    ctitle = _chart_title(z, names, member_chart)
                    budget.chart_items.append(DocImage(
                        len(budget.chart_items) + 1, member_chart, "xlsx", kind="chart", sheet=sname, cell=cell,
                        alt=((cnv.get("descr") or cnv.get("title") or "") if cnv is not None else "").strip(),
                        context=ctitle, title=ctitle,
                        shape_name=(cnv.get("name") or "") if cnv is not None else "",
                        chartsheet=smember.startswith("xl/chartsheets/")))
        # 2. pictures placed inside cells
        if rich:
            for address, vm in _cells_with_vm(z, smember):
                if vm in rich:
                    member, alt = rich[vm]
                    add(member, sname, alt, cell=address, kind="incell")
        # 3. pictures in the page header and footer (legacy VML drawing)
        for _i, kind, vmember in srels:
            if kind != "vmlDrawing" or vmember not in names:
                continue
            try:
                vroot = ET.fromstring(z.read(vmember))
            except ET.ParseError:
                continue
            vrels = _image_rels(z, names, vmember)
            for shape in vroot.iter(T_VML_SHAPE):
                code = shape.get("id") or ""
                if code not in HF_CODES:
                    continue
                for im in shape.iter(T_IMAGEDATA):
                    add(vrels.get(im.get(A_RELID) or im.get(A_RID)), sname, "", kind="headerfooter", hf=code)
        # 4. the sheet background picture
        for _i, kind, member in srels:
            if kind == "image":
                add(member, sname, "", kind="background")
    return images


# ---------------------------------------------------------------- ODF (odt, ods, odp)
def _odf_text(elem):
    parts = [elem.text or ""]
    for child in elem:
        if child.tag != f"{{{DRAW}}}frame":
            parts.append(_odf_text(child))
        parts.append(child.tail or "")
    return "".join(parts)


def scan_odf(z, names, budget):
    if "content.xml" not in names:
        raise DocImageError("nocontent")
    root = ET.fromstring(z.read("content.xml"))
    parent = {c: p for p in root.iter() for c in p}
    pages = {el: n for n, el in enumerate(root.iter(f"{{{DRAW}}}page"), 1)}
    images = []
    for frame in root.iter(f"{{{DRAW}}}frame"):
        img = frame.find(f"{{{DRAW}}}image")
        href = img.get(f"{{{XLINK}}}href") if img is not None else None
        if not href or "://" in href:
            continue
        member = posixpath.normpath(href)
        if member not in names or not budget.ok(member):
            continue
        desc = frame.find(f"{{{SVG}}}desc")
        title = frame.find(f"{{{SVG}}}title")
        alt = "".join(desc.itertext()) if desc is not None else ""
        alt = alt or ("".join(title.itertext()) if title is not None else "")
        ctx, slide, sheet = "", None, ""
        node = parent.get(frame)
        while node is not None:
            if not ctx and node.tag in (f"{{{TEXT}}}p", f"{{{TEXT}}}h"):
                ctx = _clean(_odf_text(node))
            if slide is None and node in pages:
                slide = pages[node]
            if not sheet and node.tag == f"{{{TABLE}}}table":
                sheet = node.get(f"{{{TABLE}}}name") or ""
            node = parent.get(node)
        images.append(DocImage(len(images) + 1, member, "odf", alt=alt.strip(), context=ctx,
                               slide=slide, sheet=sheet))
    return images
