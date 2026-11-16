"""Data model shared by all scanners. No NVDA imports."""
import os
from dataclasses import dataclass
from typing import Optional


class DocImageError(Exception):
    """code is one of: unsupported, corrupt, nocontent. The UI turns it into a translated message."""

    def __init__(self, code):
        super().__init__(code)
        self.code = code


@dataclass
class DocImage:
    index: int
    member: str                       # path of the media file inside the zip
    fmt: str = "docx"                 # docx, pptx, xlsx, odf
    part: str = "body"                # docx only: body, header, footer, footnotes, endnotes
    paragraph: int = 0                # docx body: Word-style paragraph index of the anchor paragraph
    alt: str = ""
    context: str = ""
    in_textbox: bool = False
    slide: Optional[int] = None       # pptx, odp
    sheet: str = ""                   # xlsx, ods
    cell: str = ""                    # xlsx
    kind: str = ""                    # xlsx: "" (floating), incell, background, headerfooter; xlsx/pptx: chart
    hf: str = ""                      # xlsx header/footer slot: LH CH RH LF CF RF
    title: str = ""                   # charts: the title text of the chart
    shape_name: str = ""              # xlsx/pptx: name of the shape in the document (used to find it live)
    chartsheet: bool = False          # xlsx: the chart fills a whole chart sheet
    page_est: Optional[int] = None    # docx: estimated from page-break markers in the file
    page: Optional[int] = None        # docx: exact, from Word layout
    line: Optional[int] = None        # docx: exact, from Word layout
    char_start: Optional[int] = None  # docx: character offset in Word (for Go to / alt text editing)
    slot: int = 0                     # docx: n-th image inside its paragraph
    path: str = ""                    # extracted file, set by docimages.extract_files

    @property
    def name(self):
        return os.path.basename(self.path)
