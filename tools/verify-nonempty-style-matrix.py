#!/usr/bin/env python3
"""Generated-data-only compatibility integration test. Requires openpyxl and Pillow.

--tool 'dotnet tools/RelaxCat.ExcelTool/bin/Release/net10.0/RelaxCat.ExcelTool.dll'
Runs the built optimizer; without --tool only creates and validates input fixtures.
This is package/model verification, NOT Microsoft Excel rendering or COM verification.
"""
import argparse
import copy
import hashlib
import json
from pathlib import Path
import posixpath
import shlex
import subprocess
import zipfile
from xml.etree import ElementTree as ET

from PIL import Image as PillowImage
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
def tag(name): return '{' + NS + '}' + name

def read(path):
    with zipfile.ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}

def write(path, parts):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in sorted(parts.items()):
            info = zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)

def xml(node): return ET.tostring(node, encoding='utf-8', xml_declaration=True)

def generate(folder, count):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Synthetic data'
    ws.append(['Item', 'Quantity', 'Price', 'Total'])
    for row in range(2, 102):
        ws.append(['Generated %03d' % row, row % 7 + 1, row * 1.25, '=B%d*C%d' % (row, row)])
        ws.cell(row, 3).number_format = '0.0000 "units"'
        ws.cell(row, 4).number_format = '#,##0.00'
        ws.cell(row, 1).alignment = Alignment(wrap_text=True)
    ws.freeze_panes = 'B2'
    ws.oddHeader.center.text = 'Generated &Bheader&B &G'
    ws.oddHeader.center.font = 'Microsoft YaHei,Regular'
    ws.oddHeader.center.size = 14
    ws.oddFooter.right.text = 'Page &P of &N'
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['A'].font = Font(name='Calibri', bold=True)
    ws.row_dimensions[2].font = Font(name='Calibri', italic=True)
    table = Table(displayName='SyntheticTable', ref='A1:D101')
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws.add_table(table)
    ws['F1'] = 'Synthetic link'
    ws['F1'].hyperlink = 'https://example.invalid/synthetic'
    ws['F2'] = 'Comment anchor'
    ws['F2'].comment = Comment('Generated comment only', 'Synthetic generator')
    ws.merge_cells('F4:H4')
    ws['F4'] = 'Generated merged heading'
    ws.conditional_formatting.add('B2:B101', CellIsRule(operator='greaterThan', formula=['3'], fill=PatternFill('solid', fgColor='FFFF0000')))
    validation = DataValidation(type='whole', operator='between', formula1=1, formula2=1000)
    validation.add('B2:B101')
    ws.add_data_validation(validation)
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=8), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=8))
    ws.add_chart(chart, 'F6')
    image_path = folder / 'synthetic.png'
    PillowImage.new('RGB', (16, 16), (30, 140, 90)).save(image_path)
    ws.add_image(Image(image_path), 'F22')
    summary = wb.create_sheet('Summary')
    summary['A1'] = "=SUM('Synthetic data'!D2:D101)"
    summary['A2'] = '=SUM(SyntheticTable[Quantity])'
    hidden = wb.create_sheet('Hidden data')
    hidden['A1'] = 123
    hidden.sheet_state = 'hidden'
    path = folder / 'rich-nonempty.xlsx'
    wb.save(path)
    parts = read(path)
    # A worksheet textbox/shape is DrawingML, NOT a print header.
    drawing = ET.fromstring(parts['xl/drawings/drawing1.xml'])
    drawing.append(ET.fromstring('''<xdr:oneCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="2000000" cy="400000"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="99" name="Synthetic title textbox"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="1800" b="1"><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>Synthetic worksheet heading</a:t></a:r><a:endParaRPr lang="zh-CN"/></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>'''))
    parts['xl/drawings/drawing1.xml'] = xml(drawing)
    # A header image uses VML, separate from the worksheet DrawingML picture.
    sheet = ET.fromstring(parts['xl/worksheets/sheet1.xml'])
    header_drawing = ET.Element(tag('legacyDrawingHF'), {'{' + REL + '}id': 'headerImage'})
    tables = sheet.find(tag('tableParts'))
    sheet.insert(list(sheet).index(tables) if tables is not None else len(sheet), header_drawing)
    parts['xl/worksheets/sheet1.xml'] = xml(sheet)
    rels = ET.fromstring(parts['xl/worksheets/_rels/sheet1.xml.rels'])
    ET.SubElement(rels, '{' + PKG + '}Relationship', Id='headerImage', Type=REL + '/vmlDrawing', Target='../drawings/vmlDrawingHeader.vml')
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = xml(rels)
    parts['xl/drawings/vmlDrawingHeader.vml'] = ('''<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="CH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:12pt;height:12pt;z-index:1"><v:imagedata o:relid="headerPng" o:title="Synthetic header image"/><o:lock v:ext="edit" rotation="t"/></v:shape></xml>''').encode()
    parts['xl/drawings/_rels/vmlDrawingHeader.vml.rels'] = (f'<Relationships xmlns="{PKG}"><Relationship Id="headerPng" Type="{REL}/image" Target="../media/image1.png"/></Relationships>').encode()
    root = ET.fromstring(parts['xl/styles.xml'])
    bases, names, cells = (root.find(tag(s)) for s in ['cellStyleXfs', 'cellStyles', 'cellXfs'])
    for i in range(1, count):
        bases.append(copy.deepcopy(bases[0]))
        ET.SubElement(names, tag('cellStyle'), name='SyntheticCopy_%05d' % i, xfId=str(i), hidden='1')
    bases.set('count', str(len(bases)))
    names.set('count', str(len(names)))
    # Make live cells depend on copied base identities, not just unused debris.
    for i, cell in enumerate(cells):
        cell.set('xfId', str((i * 7919) % count))
    parts['xl/styles.xml'] = xml(root)
    write(path, parts)
    # Unknown namespace/extension must remain a blocker, even in destructive mode.
    unknown = dict(parts)
    root = ET.fromstring(unknown['xl/styles.xml'])
    ext = ET.SubElement(ET.SubElement(root, tag('extLst')), tag('ext'), uri='urn:synthetic:unknown-reference')
    ET.SubElement(ext, '{urn:synthetic:unknown}styleBinding', xfId=str(count - 1))
    unknown['xl/styles.xml'] = xml(root)
    write(folder / 'unknown-extension.xlsx', unknown)
    external = dict(parts)
    workbook = ET.fromstring(external['xl/workbook.xml'])
    # Schema order: externalReferences belongs after definedNames and before calcPr.
    refs = ET.Element(tag('externalReferences'))
    ET.SubElement(refs, tag('externalReference'), {'{' + REL + '}id': 'syntheticExternal'})
    calc = workbook.find(tag('calcPr'))
    workbook.insert(list(workbook).index(calc) if calc is not None else len(workbook), refs)
    external['xl/workbook.xml'] = xml(workbook)
    relations = ET.fromstring(external['xl/_rels/workbook.xml.rels'])
    ET.SubElement(relations, '{' + PKG + '}Relationship', Id='syntheticExternal', Type=REL + '/externalLink', Target='externalLinks/externalLink1.xml')
    external['xl/_rels/workbook.xml.rels'] = xml(relations)
    external['xl/externalLinks/externalLink1.xml'] = (f'<externalLink xmlns="{NS}" xmlns:r="{REL}"><externalBook r:id="source"><sheetNames><sheetName val="Synthetic source"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1" t="n"><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>').encode()
    external['xl/externalLinks/_rels/externalLink1.xml.rels'] = (f'<Relationships xmlns="{PKG}"><Relationship Id="source" Type="{REL}/externalLinkPath" Target="file:///synthetic/source.xlsx" TargetMode="External"/></Relationships>').encode()
    types = ET.fromstring(external['[Content_Types].xml'])
    ET.SubElement(types, '{' + CT + '}Override', PartName='/xl/externalLinks/externalLink1.xml', ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml')
    external['[Content_Types].xml'] = xml(types)
    sheet = ET.fromstring(external['xl/worksheets/sheet2.xml'])
    sheet.find(tag('dimension')).set('ref', 'A1:A3')
    ET.SubElement(sheet.find(tag("sheetData")), tag("row"), r="3").append(ET.fromstring(f"<c xmlns='{NS}' r='A3'><f>'[1]Synthetic source'!A1</f><v>42</v></c>"))
    external['xl/worksheets/sheet2.xml'] = xml(sheet)
    write(folder / 'external-links.xlsx', external)
    return [path, folder / 'external-links.xlsx', folder / 'unknown-extension.xlsx']

def stats(path):
    with zipfile.ZipFile(path) as z:
        info = z.getinfo('xl/styles.xml')
        root = ET.fromstring(z.read(info))
        return dict(package_bytes=path.stat().st_size, styles_raw_bytes=info.file_size,
                    styles_zip_bytes=info.compress_size,
                    **{name: len(root.find(tag(name))) for name in ['cellXfs', 'cellStyleXfs', 'cellStyles']})

def validate_indices(path):
    parts = read(path)
    style = ET.fromstring(parts['xl/styles.xml'])
    sizes = {x.tag.rsplit('}', 1)[-1]: len(x) for x in style}
    for element in style:
        if 'count' in element.attrib:
            assert int(element.get('count')) == len(element), element.tag
    for table in ['cellXfs', 'cellStyleXfs']:
        for xf in style.find(tag(table)):
            for attr, target in [('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders'), ('xfId', 'cellStyleXfs')]:
                if attr in xf.attrib:
                    assert 0 <= int(xf.get(attr)) < sizes[target], (table, attr)
    for name in style.find(tag('cellStyles')):
        assert int(name.get('xfId')) < sizes['cellStyleXfs']
    for name, data in parts.items():
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            for elem in ET.fromstring(data).iter():
                attr = 'style' if elem.tag == tag('col') else 's' if elem.tag in [tag('c'), tag('row')] else None
                if attr and attr in elem.attrib:
                    assert int(elem.get(attr)) < sizes['cellXfs'], (name, attr)
                if 'dxfId' in elem.attrib:
                    assert int(elem.get('dxfId')) < sizes['dxfs']
        if name.endswith('.rels'):
            # Every internal OPC relationship resolves, not just worksheet styles.
            base = '' if name == '_rels/.rels' else name.rsplit('/_rels/', 1)[0]
            for rel in ET.fromstring(data):
                if rel.get('TargetMode') != 'External':
                    target = rel.get('Target')
                    dest = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))
                    assert dest in parts, (name, dest)

def snapshot(path):
    wb = load_workbook(path, data_only=False)
    result = []
    for ws in wb:
        cells = []
        for row in ws:
            for cell in row:
                if cell.value is not None:
                    cells.append((cell.coordinate, cell.value, cell.data_type, cell.number_format,
                                  str(cell.font), str(cell.fill), str(cell.border), str(cell.alignment), str(cell.protection),
                                  cell.hyperlink.target if cell.hyperlink else None))
        result.append((ws.title, ws.sheet_state, cells, sorted(str(x) for x in ws.merged_cells.ranges),
                       list(ws.tables), len(ws._charts), len(ws._images),
                       [(key, str(dim.font), str(dim.fill), str(dim.border), dim.number_format,
                         str(dim.alignment), str(dim.protection)) for key, dim in ws.row_dimensions.items()],
                       [(key, str(dim.font), str(dim.fill), str(dim.border), dim.number_format,
                         str(dim.alignment), str(dim.protection)) for key, dim in ws.column_dimensions.items()]))
    wb.close()
    return result

def assert_unchanged_content(before, after):
    a, b = read(before), read(after)
    assert set(a) == set(b)
    for name in a:
        if name == 'xl/styles.xml':
            continue
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            # Strip only style references; all formulas, cells, row/column dimensions,
            # merges, CF, validations, drawings, tables and links must remain.
            left, right = ET.fromstring(a[name]), ET.fromstring(b[name])
            for root in [left, right]:
                for element in root.iter():
                    attr = 'style' if element.tag == tag('col') else 's' if element.tag in [tag('c'), tag('row')] else None
                    if attr:
                        element.attrib.pop(attr, None)
            assert ET.tostring(left) == ET.tostring(right), name
        else:
            assert a[name] == b[name], name
    assert snapshot(before) == snapshot(after), 'cell formatting/model changed'

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--styles', type=int, default=48846)
    parser.add_argument('--tool', help='built CLI command, quoted as one argument')
    args = parser.parse_args()
    if args.styles < 2:
        parser.error('--styles must be at least 2')
    args.output_dir.mkdir(parents=True, exist_ok=True)
    report = {'scope': 'Synthetic package and openpyxl model validation; not Excel rendering or native save/reopen', 'cases': []}
    for path in generate(args.output_dir, args.styles):
        validate_indices(path)
        original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
        case = {'case': path.name, 'before': stats(path)}
        if args.tool:
            out = path.with_name(path.stem + '.optimized.xlsx')
            command = shlex.split(args.tool)
            result = subprocess.run(command + ['optimize', str(path), '--output', str(out), '--destructive-styles', '--json'], check=True, text=True, capture_output=True)
            status = json.loads(result.stdout)
            validate_indices(out)
            assert_unchanged_content(path, out)
            assert hashlib.sha256(path.read_bytes()).hexdigest() == original_hash
            if path.name.startswith('unknown'):
                assert status['AdvancedOptimizationSkipReason'], status
                assert read(path) == read(out)
            else:
                assert status['AdvancedOptimizationSkipReason'] is None, status['AdvancedOptimizationSkipReason']
                assert stats(out)['cellStyleXfs'] == 1
                assert stats(out)['cellStyles'] == 1
            twice = path.with_name(path.stem + '.twice.xlsx')
            subprocess.run(command + ['optimize', str(out), '--output', str(twice), '--destructive-styles', '--json'], check=True, capture_output=True)
            assert out.read_bytes() == twice.read_bytes(), 'not byte-idempotent'
            case.update(after=stats(out), status='passed', byte_idempotent=True)
            if path.name == 'rich-nonempty.xlsx':
                safe = path.with_name(path.stem + '.safe.xlsx')
                safe_result = subprocess.run(command + ['optimize', str(path), '--output', str(safe), '--gc', '--json'], check=True, text=True, capture_output=True)
                safe_status = json.loads(safe_result.stdout)
                assert safe_status['AdvancedOptimizationSkipReason'] is None, safe_status['AdvancedOptimizationSkipReason']
                validate_indices(safe)
                assert_unchanged_content(path, safe)
                assert stats(safe)['cellStyles'] == stats(path)['cellStyles']
                case['safe_gc'] = 'passed; named identities retained'
        else:
            snapshot(path)
            case['status'] = 'fixture validated only; optimizer not run'
        report['cases'].append(case)
    (args.output_dir / 'matrix-report.json').write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps(report, indent=2))

if __name__ == '__main__':
    main()
