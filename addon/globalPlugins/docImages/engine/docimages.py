"""Public API: scan a document, extract its images, prepare/export them. No NVDA imports."""
import os
import re
import shutil
import zipfile

from . import scanners
from .model import DocImage, DocImageError  # noqa: F401  (re-exported)

DOCX_EXTS = {".docx", ".docm", ".dotx", ".dotm"}
PPTX_EXTS = {".pptx", ".pptm", ".potx", ".ppsx"}
XLSX_EXTS = {".xlsx", ".xlsm", ".xltx"}
ODF_EXTS = {".odt", ".ods", ".odp", ".ott"}
LEGACY_WORD_EXTS = {".doc", ".dot"}  # converted to .docx first (needs Word or LibreOffice)
SUPPORTED_EXTS = DOCX_EXTS | PPTX_EXTS | XLSX_EXTS | ODF_EXTS | LEGACY_WORD_EXTS
SHARE_OK = {".png", ".jpg", ".jpeg", ".gif"}

_SCANNERS = {"docx": scanners.scan_docx, "pptx": scanners.scan_pptx,
             "xlsx": scanners.scan_xlsx, "odf": scanners.scan_odf}


class ScanResult(list):
    """The images found, plus the native charts of the document. Charts are not pictures: `charts` counts
    them all, `chart_items` lists those that can be exported through the live application."""
    charts = 0
    chart_items = ()


def family(path):
    ext = os.path.splitext(path)[1].lower()
    for name, exts in (("docx", DOCX_EXTS), ("pptx", PPTX_EXTS), ("xlsx", XLSX_EXTS), ("odf", ODF_EXTS)):
        if ext in exts:
            return name
    return None


def scan_images(path):
    """List the images (alt text, location, context) without extracting them. Fast."""
    fam = family(path)
    if fam is None:
        raise DocImageError("unsupported")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise DocImageError("corrupt")
    with z:
        budget = scanners.Budget(z)
        try:
            images = _SCANNERS[fam](z, set(z.namelist()), budget)
        except (zipfile.BadZipFile, KeyError):
            raise DocImageError("corrupt")
        except DocImageError:
            raise
        except Exception as e:
            import xml.etree.ElementTree as ET
            if isinstance(e, ET.ParseError):
                raise DocImageError("corrupt")
            raise
    result = ScanResult(images)
    result.charts = budget.charts
    result.chart_items = list(budget.chart_items)
    return result


def extract_files(path, images, out_dir):
    """Write each image to out_dir and set image.path."""
    os.makedirs(out_dir, exist_ok=True)
    written = {}  # a media file used several times is written once
    with zipfile.ZipFile(path) as z:
        for img in images:
            if img.member not in written:
                ext = os.path.splitext(img.member)[1].lower() or ".bin"
                dest = os.path.join(out_dir, f"image_{img.index:03d}{ext}")
                with open(dest, "wb") as f:
                    f.write(z.read(img.member))
                written[img.member] = dest
            img.path = written[img.member]


MEDIA_PREFIXES = ("word/media/", "xl/media/", "ppt/media/", "Pictures/")
PICTURE_EXTS = {".png", ".jpg", ".jpeg", ".jfif", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".webp"}


def unlisted_count(path, images):
    """How many picture files the document stores that no listed image points to (fills, chart
    parts, kinds of placement this add-on does not know). Lets the UI avoid staying silent about them."""
    try:
        with zipfile.ZipFile(path) as z:
            media = {n for n in z.namelist()
                     if n.startswith(MEDIA_PREFIXES) and os.path.splitext(n)[1].lower() in PICTURE_EXTS}
    except zipfile.BadZipFile:
        return 0
    return len(media - {i.member for i in images})


def prepare_for_sharing(path):
    """Return a path to a PNG/JPEG/GIF version of the image (converts EMF/WMF/TIFF/etc. with Pillow)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in SHARE_OK:
        return path
    try:
        from PIL import Image
    except ImportError:  # NVDA builds differ in which libraries they ship
        raise DocImageError("nopillow")
    im = Image.open(path)
    try:
        im.load(dpi=200)
    except TypeError:
        im.load()
    out = os.path.splitext(path)[0] + ".png"
    im.convert("RGBA" if "A" in im.getbands() else "RGB").save(out, "PNG")
    return out


def _slug(text):
    return re.sub(r"[^\w\-]+", "_", text).strip("_")


def export_name(img):
    """File name for exports, carrying the location: p03_image_007.png, slide02_image_004.jpg ..."""
    if img.kind == "chart":
        page = img.page or img.page_est
        if img.slide:
            base = f"slide{img.slide:02d}"
        elif img.sheet:
            base = _slug(img.sheet) or "sheet"
        elif page:
            base = f"p{page:03d}"
        else:
            base = "document"
        return f"{base}_chart_{img.index:03d}.png"
    ext = os.path.splitext(img.path or img.member)[1].lower() or ".bin"
    tail = f"image_{img.index:03d}{ext}"
    if img.fmt == "pptx" or (img.fmt == "odf" and img.slide):
        return f"slide{img.slide:02d}_{tail}"
    if img.sheet:
        tag = img.cell or img.hf or ("background" if img.kind == "background" else "")
        tag = f"_{tag}" if tag else ""
        return f"{_slug(img.sheet) or 'sheet'}{tag}_{tail}"
    if img.fmt == "docx":
        if img.part != "body":
            return f"{img.part}_{tail}"
        page = img.page or img.page_est
        if page:
            return f"p{page:03d}_{tail}"
    return tail


def export_all(images, folder):
    """Copy every extracted image into folder with location-based names. Returns the count."""
    n = 0
    for img in images:
        if img.path and os.path.exists(img.path):
            shutil.copyfile(img.path, os.path.join(folder, export_name(img)))
            n += 1
    return n
