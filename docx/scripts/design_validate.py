from __future__ import annotations
import argparse, json, re, shutil, subprocess, sys, tempfile
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'w': W, 'wp': WP, 'a': A}

def twips(node, attr='w:val', default=0):
    if node is None: return default
    return int(node.get(f'{{{W}}}{attr.split(":")[-1]}', default))

def inspect_docx(path: Path):
    with ZipFile(path) as z:
        xml = ET.fromstring(z.read('word/document.xml'))
        styles_xml = ET.fromstring(z.read('word/styles.xml')) if 'word/styles.xml' in z.namelist() else None
        fonts = sorted({v for n in styles_xml.findall('.//w:rFonts', NS) if (v := n.get(f'{{{W}}}eastAsia') or n.get(f'{{{W}}}ascii') or n.get(f'{{{W}}}hAnsi'))}) if styles_xml is not None else []
        sect = xml.find('.//w:sectPr', NS)
        pg = sect.find('w:pgSz', NS); mar = sect.find('w:pgMar', NS)
        width = twips(pg, 'w:w', 11906); height = twips(pg, 'w:h', 16838)
        left = twips(mar, 'w:left', 1440); right = twips(mar, 'w:right', 1440)
        usable = width - left - right
        tables=[]
        for table in xml.findall('.//w:tbl', NS):
            grid = [twips(n, 'w:w') for n in table.findall('w:tblGrid/w:gridCol', NS)]
            rows=[]
            for row in table.findall('w:tr', NS):
                cursor=0; cells=[]
                for c in row.findall('w:tc', NS):
                    span_node=c.find('w:tcPr/w:gridSpan', NS); span=twips(span_node,'w:val',1)
                    tcw=c.find('w:tcPr/w:tcW', NS); cw=twips(tcw,'w:w',0)
                    expected=sum(grid[cursor:cursor+span])
                    cells.append({'span':span,'width':cw,'expected':expected,'ok':cw==expected})
                    cursor += span
                rows.append({'cells':cells,'grid_complete':cursor==len(grid)})
            tables.append({'grid':grid,'width':sum(grid),'fits_usable':sum(grid)<=usable,'rows':rows})
        images=[]
        for ext in xml.findall('.//wp:inline/wp:extent', NS):
            images.append({'width_emu':int(ext.get('cx',0)), 'height_emu':int(ext.get('cy',0)), 'fits_usable':int(ext.get('cx',0)) <= usable*635})
        return {'page': {'width_dxa':width,'height_dxa':height,'margins_dxa':{'left':left,'right':right},'usable_width_dxa':usable}, 'tables':tables, 'images':images, 'table_grid_ok': all(c['ok'] and r['grid_complete'] for t in tables for r in t['rows'] for c in r['cells']) if tables else True, 'all_tables_fit': all(t['fits_usable'] for t in tables), 'all_images_fit': all(i['fits_usable'] for i in images), 'declared_fonts': fonts, 'fonts_explicit': bool(fonts)}

def inspect_pdf(path: Path, expected_fonts=()):
    try:
        import fitz
    except ImportError:
        return {'available':False,'reason':'PyMuPDF is not installed'}
    doc=fitz.open(path); pages=[]; blank=[]; out_of_bounds=[]; orphan_headings=[]; font_names=set()
    for i,page in enumerate(doc):
        rect=page.rect; text=page.get_text('text').strip(); blocks=page.get_text('blocks')
        pages.append({'page':i+1,'text_chars':len(text),'width':rect.width,'height':rect.height})
        for f in page.get_fonts(full=True):
            if len(f) > 3 and f[3]: font_names.add(f[3])
        if len(text)<10: blank.append(i+1)
        for b in blocks:
            if b[0] < -1 or b[2] > rect.width+1 or b[1] < -1 or b[3] > rect.height+1: out_of_bounds.append({'page':i+1,'rect':list(b[:4])})
        for block in page.get_text('dict').get('blocks', []):
            for line in block.get('lines', []):
                label=''.join(span.get('text','') for span in line.get('spans',[])).strip()
                if re.match(r'^(第[一二三四五六七八九十百]+条|[一二三四五六七八九十]+、)', label) and line.get('bbox',[0,0,0,0])[3] > rect.height-90:
                    orphan_headings.append({'page':i+1,'text':label,'bbox':line['bbox']})
    normalize = lambda value: re.sub(r'[^a-z0-9]', '', value.lower())
    expected_ok = not expected_fonts or any(any(normalize(e) in normalize(f) for f in font_names) for e in expected_fonts)
    return {'available':True,'page_count':len(doc),'pages':pages,'blank_pages':blank,'out_of_bounds':out_of_bounds,'orphan_headings':orphan_headings,'fonts':sorted(font_names),'expected_fonts':list(expected_fonts),'fonts_match':expected_ok,'pass':not blank and not out_of_bounds and not orphan_headings and expected_ok}

def main():
    ap=argparse.ArgumentParser(description='Validate DOCX geometry and optional rendered PDF.')
    ap.add_argument('docx', type=Path); ap.add_argument('--pdf', type=Path); ap.add_argument('--json', type=Path); ap.add_argument('--expected-font', action='append', default=[])
    a=ap.parse_args(); result={'docx':str(a.docx.resolve()),'structure':inspect_docx(a.docx)}
    if a.pdf: result['pdf']=inspect_pdf(a.pdf, a.expected_font)
    result['pass']=result['structure']['table_grid_ok'] and result['structure']['all_tables_fit'] and result['structure']['all_images_fit'] and result['structure']['fonts_explicit'] and (not a.pdf or result['pdf'].get('pass',False))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if a.json: a.json.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return 0 if result['pass'] else 1
if __name__=='__main__': raise SystemExit(main())
