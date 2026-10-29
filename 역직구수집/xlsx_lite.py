# -*- coding: utf-8 -*-
"""openpyxl 없이 표준 라이브러리만으로 xlsx를 만드는 최소 작성기.
셀 값: 숫자, 문자열, "=수식". 스타일: head / input / link / money / pct / plain."""
import zipfile
from xml.sax.saxutils import escape


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


STYLES = {"plain": 0, "head": 1, "input": 2, "link": 3, "money": 4, "pct": 5, "input_money": 6, "title": 7, "note": 8}

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0;(#,##0);-"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts>
<fonts count="6">
<font><sz val="10"/><name val="맑은 고딕"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="맑은 고딕"/></font>
<font><sz val="10"/><color rgb="FF0000FF"/><name val="맑은 고딕"/></font>
<font><u/><sz val="10"/><color rgb="FF0563C1"/><name val="맑은 고딕"/></font>
<font><b/><sz val="13"/><name val="맑은 고딕"/></font>
<font><i/><sz val="10"/><color rgb="FF7F7F7F"/><name val="맑은 고딕"/></font>
</fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="9">
<xf fontId="0" xfId="0"/>
<xf fontId="1" fillId="2" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf fontId="2" fillId="3" applyFont="1" applyFill="1"/>
<xf fontId="3" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf fontId="0" numFmtId="164" applyNumberFormat="1"/>
<xf fontId="0" numFmtId="165" applyNumberFormat="1"/>
<xf fontId="2" fillId="3" numFmtId="164" applyFont="1" applyFill="1" applyNumberFormat="1"/>
<xf fontId="4" applyFont="1"/>
<xf fontId="5" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def _cell(ref, value, style):
    s = STYLES.get(style, 0)
    if value is None or value == "":
        return f'<c r="{ref}" s="{s}"/>' if s else ""
    if isinstance(value, bool):
        value = str(value)
    if isinstance(value, (int, float)):
        return f'<c r="{ref}" s="{s}"><v>{value}</v></c>'
    value = str(value)
    if value.startswith("="):
        return f'<c r="{ref}" s="{s}"><f>{escape(value[1:])}</f></c>'
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'


def _sheet_xml(sheet):
    cols = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in sorted(sheet.get("widths", {}).items()))
    rows = []
    for r_idx, row in enumerate(sheet["rows"], 1):
        cells = "".join(_cell(f"{col_letter(c_idx)}{r_idx}", v, st) for c_idx, (v, st) in enumerate(row, 1))
        rows.append(f'<row r="{r_idx}">{cells}</row>')
    pane = ""
    if sheet.get("freeze"):
        r, c = sheet["freeze"]
        pane = (f'<sheetViews><sheetView workbookViewId="0"><pane xSplit="{c}" ySplit="{r}" '
                f'topLeftCell="{col_letter(c + 1)}{r + 1}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>')
    af = f'<autoFilter ref="{sheet["filter"]}"/>' if sheet.get("filter") else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'{pane}{"<cols>" + cols + "</cols>" if cols else ""}<sheetData>{"".join(rows)}</sheetData>{af}</worksheet>')


def save(path, sheets):
    """sheets: [{"name":..., "rows":[[(value, style), ...], ...], "widths":{col:width}, "freeze":(row,col), "filter":"A4:Z10"}]"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        n = len(sheets)
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, n + 1))
                   + '</Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        defined = "".join(
            f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">\'{s["name"]}\'!${s["filter"].split(":")[0][0]}${s["filter"].split(":")[0][1:]}:${"".join(ch for ch in s["filter"].split(":")[1] if ch.isalpha())}${"".join(ch for ch in s["filter"].split(":")[1] if ch.isdigit())}</definedName>'
            for i, s in enumerate(sheets) if s.get("filter"))
        z.writestr("xl/workbook.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                   + "".join(f'<sheet name="{escape(s["name"])}" sheetId="{i}" r:id="rId{i}"/>' for i, s in enumerate(sheets, 1))
                   + '</sheets>' + (f'<definedNames>{defined}</definedNames>' if defined else '')
                   + '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, n + 1))
                   + f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '</Relationships>')
        z.writestr("xl/styles.xml", STYLES_XML)
        for i, s in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(s))
