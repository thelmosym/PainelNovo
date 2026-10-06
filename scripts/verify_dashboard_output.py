import zipfile, xml.etree.ElementTree as ET

with zipfile.ZipFile('Painel de controle MemoriaPetrobras V6.xlsm', 'r') as z:
    wb_xml = z.read('xl/workbook.xml')
    root = ET.fromstring(wb_xml)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rels_xml = z.read('xl/_rels/workbook.xml.rels')
    rels_root = ET.fromstring(rels_xml)
    r_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    rel_map = {r.attrib['Id']: r.attrib['Target'] for r in rels_root.findall('.//rel:Relationship', r_ns)}

    sheets = [s.attrib['name'] for s in root.findall('.//main:sheet', ns)]
    print('First 6 sheets in workbook:')
    for i, s in enumerate(sheets[:6]):
        print(f'  {i+1}: {s}')

    # Check shared strings
    sst = []
    if 'xl/sharedStrings.xml' in z.namelist():
        sst_xml = z.read('xl/sharedStrings.xml')
        sst_root = ET.fromstring(sst_xml)
        for si in sst_root.findall('.//main:si', ns):
            t = si.find('.//main:t', ns)
            if t is not None and t.text:
                sst.append(t.text)
            else:
                text_parts = [r.find('.//main:t', ns).text for r in si.findall('.//main:r', ns) if r.find('.//main:t', ns) is not None and r.find('.//main:t', ns).text]
                sst.append(''.join(text_parts))

    def get_val(c):
        v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        if v is None or v.text is None:
            return ''
        t = c.attrib.get('t')
        if t == 's':
            idx = int(v.text)
            return sst[idx] if idx < len(sst) else v.text
        return v.text

    # Read 01_DASHBOARD
    for s in root.findall('.//main:sheet', ns):
        if s.attrib['name'] == '01_DASHBOARD':
            target = 'xl/' + rel_map[s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
            s_xml = z.read(target)
            s_root = ET.fromstring(s_xml)
            print('\n=== 01_DASHBOARD ROWS ===')
            for r in s_root.findall('.//main:row', ns):
                r_num = int(r.attrib.get('r'))
                if r_num in [1, 4, 6, 9, 10, 14, 15, 16, 17, 18, 19, 20, 39, 40]:
                    c_vals = [f"{c.attrib.get('r')}:{get_val(c)}" for c in r.findall('./main:c', ns) if get_val(c) != '']
                    if c_vals:
                        print(f"Row {r_num}: {' | '.join(c_vals[:7])}")
