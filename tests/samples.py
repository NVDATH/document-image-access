"""Builds small real documents with the libraries that write them (XlsxWriter, python-docx, python-pptx)."""
import os
import zipfile

from PIL import Image


def png(path, color):
    Image.new("RGB", (40, 40), color).save(path)
    return str(path)


def build_all(folder):
    folder = str(folder)
    pics = {name: png(os.path.join(folder, f"{name}.png"), name)
            for name in ("red", "blue", "green", "yellow", "purple", "orange")}
    out = {"pics": pics}
    out["xlsx_floating"] = xlsx_floating(folder, pics)
    out["xlsx_header"] = xlsx_header(folder, pics)
    out["xlsx_background"] = xlsx_background(folder, pics)
    out["xlsx_incell"] = xlsx_incell(folder, pics)
    out["xlsx_chartsheet"] = xlsx_chartsheet(folder)
    out["xlsx_shapefill"] = xlsx_shape_fill(folder, pics)
    out["docx"] = docx_body(folder, pics)
    out["docx_chart"] = docx_chart(folder, pics)
    out["pptx"] = pptx_deck(folder, pics)
    out["odt"] = odf(folder, pics, "odt")
    out["odp"] = odf(folder, pics, "odp")
    return out


def xlsx_floating(folder, pics):
    import xlsxwriter
    path = os.path.join(folder, "floating.xlsx")
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Pictures")
    ws.insert_image("B3", pics["red"], {"description": "Red square"})
    ws.insert_image("F3", pics["blue"], {"description": "Blue circle"})
    ws.write_column("L1", [3, 5, 2, 8])
    chart = wb.add_chart({"type": "column"})
    chart.add_series({"values": "=Pictures!$L$1:$L$4"})
    chart.set_title({"name": "Quarterly sales"})
    ws.insert_chart("H12", chart, {"description": "Column chart of sales"})
    wb.close()
    return path


def xlsx_header(folder, pics):
    import xlsxwriter
    path = os.path.join(folder, "header.xlsx")
    wb = xlsxwriter.Workbook(path)
    wb.add_worksheet("Header").set_header("&L&G", {"image_left": pics["yellow"]})
    wb.close()
    return path


def xlsx_background(folder, pics):
    import xlsxwriter
    path = os.path.join(folder, "background.xlsx")
    wb = xlsxwriter.Workbook(path)
    wb.add_worksheet("Background").set_background(pics["orange"])
    wb.close()
    return path


def xlsx_incell(folder, pics):
    import xlsxwriter
    path = os.path.join(folder, "incell.xlsx")
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("InCell")
    ws.embed_image("A1", pics["green"], {"description": "Green triangle in a cell"})
    ws.embed_image("A3", pics["purple"])
    for row in range(10, 60):  # the same picture again and again
        ws.embed_image(row, 0, pics["red"])
    wb.close()
    return path


def xlsx_chartsheet(folder):
    import xlsxwriter
    path = os.path.join(folder, "chartsheet.xlsx")
    wb = xlsxwriter.Workbook(path)
    data = wb.add_worksheet("chart")  # a normal sheet that only has this name
    data.write_column("A1", [1, 4, 9, 16])
    sheet = wb.add_chartsheet("Chart1")
    chart = wb.add_chart({"type": "line"})
    chart.add_series({"values": "=chart!$A$1:$A$4"})
    chart.set_title({"name": "Growth over time"})
    sheet.set_chart(chart)
    wb.close()
    return path


def xlsx_shape_fill(folder, pics):
    """A picture used as the fill of a shape: stored in media, but not placed where the scanner looks."""
    import xlsxwriter
    base = os.path.join(folder, "base_shapefill.xlsx")
    wb = xlsxwriter.Workbook(base)
    wb.add_worksheet("Data").insert_image("B2", pics["red"], {"description": "Red square"})
    wb.close()
    path = os.path.join(folder, "shapefill.xlsx")
    shape = ('<xdr:twoCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>'
             '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff>'
             '<xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink="">'
             '<xdr:nvSpPr><xdr:cNvPr id="9" name="Rectangle 8"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>'
             '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip xmlns:r="http://schemas.'
             'openxmlformats.org/officeDocument/2006/relationships" r:embed="rId9"/><a:stretch><a:fillRect/>'
             '</a:stretch></a:blipFill></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>')
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/drawings/drawing1.xml":
                data = data.replace(b"</xdr:wsDr>", shape.encode() + b"</xdr:wsDr>")
            elif item.filename == "xl/drawings/_rels/drawing1.xml.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rId9" Type="http://schemas.'
                                    b'openxmlformats.org/officeDocument/2006/relationships/image" '
                                    b'Target="../media/image9.png"/></Relationships>')
            zout.writestr(item, data)
        zout.write(pics["yellow"], "xl/media/image9.png")
    return path


def docx_body(folder, pics):
    import docx
    from docx.shared import Inches
    doc = docx.Document()
    doc.add_paragraph("First paragraph.")
    run = doc.add_paragraph("Picture with alt text: ").add_run()
    run.add_picture(pics["red"], width=Inches(1))
    run._element.xpath(".//wp:docPr")[0].set("descr", "Red square")
    table = doc.add_table(rows=2, cols=2)
    table.cell(1, 1).paragraphs[0].add_run("In a table: ").add_picture(pics["blue"], width=Inches(1))
    doc.sections[0].header.paragraphs[0].add_run("Header ").add_picture(pics["green"], width=Inches(0.5))
    doc.sections[0].footer.paragraphs[0].add_run("Footer ").add_picture(pics["yellow"], width=Inches(0.5))
    doc.add_page_break()
    doc.add_paragraph("Second page: ").add_run().add_picture(pics["purple"], width=Inches(1))
    path = os.path.join(folder, "body.docx")
    doc.save(path)
    return path


def docx_chart(folder, pics):
    """A Word document with a native chart. The chart part comes from python-pptx's chart writer."""
    import docx
    from docx.shared import Inches
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches as PptInches

    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    data = CategoryChartData()
    data.categories = ["North", "South"]
    data.add_series("Sales", (4, 7))
    frame = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, PptInches(1), PptInches(1), PptInches(4),
                                   PptInches(3), data)
    frame.chart.has_title = True
    frame.chart.chart_title.text_frame.text = "Sales by region"
    source = os.path.join(folder, "chart_source.pptx")
    deck.save(source)
    with zipfile.ZipFile(source) as z:
        chart_xml, chart_rels = z.read("ppt/charts/chart1.xml"), z.read("ppt/charts/_rels/chart1.xml.rels")
        embedded = [n for n in z.namelist() if n.startswith("ppt/embeddings/")][0]
        embedded_blob = z.read(embedded)

    doc = docx.Document()
    doc.add_paragraph("Intro.")
    doc.add_paragraph("Chart follows this line:")
    doc.add_paragraph("Picture after the chart: ").add_run().add_picture(pics["red"], width=Inches(1))
    plain = os.path.join(folder, "_plain.docx")
    doc.save(plain)
    namespaces = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    paragraph = (f'<w:p {namespaces}><w:r><w:drawing><wp:inline><wp:extent cx="4572000" cy="2743200"/>'
                 '<wp:docPr id="50" name="Chart 1" descr="Column chart of sales"/><wp:cNvGraphicFramePr/>'
                 '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                 '<c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
    path = os.path.join(folder, "chart.docx")
    with zipfile.ZipFile(plain) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml = blob.decode()
                at = xml.index("</w:p>", xml.index("Chart follows this line:")) + len("</w:p>")
                blob = (xml[:at] + paragraph + xml[at:]).encode()
            elif item.filename == "word/_rels/document.xml.rels":
                blob = blob.decode().replace("</Relationships>", '<Relationship Id="rIdChart1" Type="http://schemas.'
                                             'openxmlformats.org/officeDocument/2006/relationships/chart" '
                                             'Target="charts/chart1.xml"/></Relationships>').encode()
            elif item.filename == "[Content_Types].xml":
                text = blob.decode().replace("</Types>", '<Override PartName="/word/charts/chart1.xml" ContentType='
                                             '"application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                                             '</Types>')
                if 'Extension="xlsx"' not in text:
                    text = text.replace("<Default ", '<Default Extension="xlsx" ContentType="application/vnd.'
                                        'openxmlformats-officedocument.spreadsheetml.sheet"/><Default ', 1)
                blob = text.encode()
            zout.writestr(item, blob)
        zout.writestr("word/charts/chart1.xml", chart_xml)
        zout.writestr("word/charts/_rels/chart1.xml.rels", chart_rels)
        zout.writestr("word/embeddings/" + embedded.split("/")[-1], embedded_blob)
    return path


def pptx_deck(folder, pics):
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches
    deck = Presentation()
    first = deck.slides.add_slide(deck.slide_layouts[5])
    first.shapes.title.text = "Company overview"
    picture = first.shapes.add_picture(pics["blue"], Inches(1), Inches(2), width=Inches(3))
    picture._element.nvPicPr.cNvPr.set("descr", "Blue circle logo")
    second = deck.slides.add_slide(deck.slide_layouts[5])
    second.shapes.title.text = "Results"
    for i, name in enumerate(("red", "green")):
        shape = second.shapes.add_picture(pics[name], Inches(0.5 + 4.5 * i), Inches(2), width=Inches(3))
        shape._element.nvPicPr.cNvPr.set("descr", "")
    data = CategoryChartData()
    data.categories = ["A", "B"]
    data.add_series("S", (1, 3))
    frame = second.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(5), Inches(5), Inches(2), data)
    frame.chart.has_title = True
    frame.chart.chart_title.text_frame.text = "Results by group"
    frame._element.nvGraphicFramePr.cNvPr.set("descr", "Column chart of results")
    path = os.path.join(folder, "deck.pptx")
    deck.save(path)
    return path


def odf(folder, pics, kind):
    ns = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
          'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
          'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
          'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
          'xmlns:xlink="http://www.w3.org/1999/xlink"')
    if kind == "odt":
        content = (f'<office:document-content {ns}><office:body><office:text><text:p>Hello <draw:frame>'
                   '<draw:image xlink:href="Pictures/a.png"/><svg:title>short</svg:title>'
                   '<svg:desc>A long description</svg:desc></draw:frame> world</text:p></office:text>'
                   '</office:body></office:document-content>')
    else:
        content = (f'<office:document-content {ns}><office:body><office:presentation><draw:page><draw:frame>'
                   '<draw:image xlink:href="Pictures/a.png"/></draw:frame></draw:page><draw:page><draw:frame>'
                   '<draw:image xlink:href="Pictures/a.png"/><svg:title>Second slide</svg:title></draw:frame>'
                   '</draw:page></office:presentation></office:body></office:document-content>')
    path = os.path.join(folder, f"doc.{kind}")
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("content.xml", content)
        z.write(pics["red"], "Pictures/a.png")
    return path
