"""Create deterministic, entirely synthetic OOXML; standard library only."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.sax.saxutils import escape
import hashlib
import json

HEADERS = ['Order', 'Channel', 'Amount']
ROWS = [['S001', 'Web', 100], ['S002', 'Web', 300], ['S003', 'Store', 50], ['S004', 'Store', None], ['S005', 'Store', 150]]
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'

def workbook(path, rows=ROWS):
    cells = []
    for i, row in enumerate([HEADERS] + rows, 1):
        cs = []
        for j, value in enumerate(row):
            ref = chr(65+j)+str(i)
            if value is None:
                continue
            if isinstance(value, str):
                cs.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>')
            else:
                cs.append(f'<c r="{ref}"><v>{value}</v></c>')
        cells.append(f'<row r="{i}">'+''.join(cs)+'</row>')
    files = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="Orders" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{S}" xmlns:r="{R}"><sheetData>'+''.join(cells)+'</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>',
        'xl/worksheets/_rels/sheet1.xml.rels': f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/table" Target="../tables/table1.xml"/></Relationships>',
        'xl/tables/table1.xml': f'<table xmlns="{S}" id="1" name="OrdersTable" displayName="OrdersTable" ref="A1:C{len(rows)+1}" totalsRowShown="0"><autoFilter ref="A1:C{len(rows)+1}"/><tableColumns count="3">'+''.join(f'<tableColumn id="{i}" name="{h}"/>' for i,h in enumerate(HEADERS,1))+'</tableColumns></table>'
    }
    with ZipFile(path, 'w', compression=ZIP_DEFLATED) as archive:
        for name, value in files.items():
            info = ZipInfo(name, (2026,1,1,0,0,0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, value.encode())


def plan(source, sha):
    return {'schema_version':1,'source':{'path':source,'sha256':sha},'question':'各渠道已知订单金额合计、均值与订单数是多少？','interpretation':{'basis':'全部5笔合成订单；金额单位CNY；空金额跳过且披露有效分母，订单数包含空金额订单。','clarifications':[]},'table':'OrdersTable','formula_policy':'reject','field_units':{'Amount':'CNY'},'missing_tokens':[],'group_by':['Channel'],'metrics':[{'id':'orders','label':'订单数','op':'count_rows','missing':'error','unit':'count','scope':'group'},{'id':'total','label':'已知金额合计','op':'sum','expr':{'field':'Amount'},'missing':'skip','unit':'CNY','scope':'group'},{'id':'average','label':'已知金额均值','op':'mean','expr':{'field':'Amount'},'missing':'skip','unit':'CNY','scope':'group'}]}

if __name__ == '__main__':
    root=Path(__file__).parent
    target=root/'orders.xlsx'
    if target.exists() or (root/'plan.json').exists():
        raise SystemExit('Sample already exists; refusing to overwrite.')
    workbook(target)
    (root/'plan.json').write_text(json.dumps(plan('examples/orders.xlsx',hashlib.sha256(target.read_bytes()).hexdigest()),ensure_ascii=False,indent=2)+'\n')
