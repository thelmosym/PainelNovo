import zipfile, xml.etree.ElementTree as ET

with zipfile.ZipFile('Painel de controle MemoriaPetrobras V6.xlsm', 'r') as z:
    sst = []
    if 'xl/sharedStrings.xml' in z.namelist():
        sst_xml = z.read('xl/sharedStrings.xml')
        sst_root = ET.fromstring(sst_xml)
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
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

    s_xml = z.read('xl/worksheets/sheet1.xml')
    s_root = ET.fromstring(s_xml)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    for r in s_root.findall('.//main:row', ns):
        c_vals = [f"{c.attrib.get('r')}:{get_val(c)}" for c in r.findall('./main:c', ns) if get_val(c) != '']
        if c_vals:
            print(f"Row {r.attrib.get('r')}: {' | '.join(c_vals)}")
