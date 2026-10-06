import zipfile, xml.etree.ElementTree as ET
from collections import Counter

with zipfile.ZipFile('Painel de controle MemoriaPetrobras V6.xlsm', 'r') as z:
    # 1. Read shared strings
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

    # Read DADOS
    dados_xml = z.read('xl/worksheets/sheet28.xml')
    dados_root = ET.fromstring(dados_xml)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rows = dados_root.findall('.//main:row', ns)
    
    # Headers are in row 2
    row2_cells = rows[1].findall('./main:c', ns)
    headers = {c.attrib.get('r')[0]: get_val(c) for c in row2_cells}
    # map column letter to header
    # actually multiple letters like AA, AB
    import re
    col_map = {}
    for c in row2_cells:
        m = re.match(r'([A-Z]+)', c.attrib.get('r'))
        if m:
            col_map[m.group(1)] = get_val(c)
    
    print('Column Mapping:')
    for k, v in col_map.items():
        v_clean = v.replace('\n', ' ').strip()
        print(f'  {k}: {v_clean}')

    # Process data rows (rows 3 onwards)
    data = []
    for r in rows[2:]:
        row_dict = {}
        for c in r.findall('./main:c', ns):
            m = re.match(r'([A-Z]+)', c.attrib.get('r'))
            if m:
                col = m.group(1)
                row_dict[col] = get_val(c)
        data.append(row_dict)

    print(f'\nTotal data rows: {len(data)}')
    
    # Analyze contracts
    contracts = Counter(r.get('A', '') for r in data)
    print('\nContracts distribution (Col A):')
    for c, cnt in contracts.most_common():
        print(f'  {c}: {cnt}')

    # Analyze activities
    activities = Counter(r.get('B', '') for r in data)
    print('\nActivities distribution (Col B):')
    for a, cnt in activities.most_common(10):
        print(f'  {a}: {cnt}')

    # Sum metrics
    def to_float(val):
        try:
            return float(val.replace(',', '.'))
        except:
            return 0.0

    total_solic = sum(to_float(r.get('H', '0')) for r in data)
    total_atend = sum(to_float(r.get('M', '0')) for r in data)
    total_qexec = sum(to_float(r.get('AE', '0')) for r in data)
    total_km = sum(to_float(r.get('Y', '0')) for r in data)

    print(f'\nMetrics:')
    print(f'  Total Solicitada (Col H): {total_solic:,.2f}')
    print(f'  Total Atendida (Col M): {total_atend:,.2f}')
    print(f'  Total QExec (Col AE): {total_qexec:,.2f}')
    print(f'  Total KM Adicional (Col Y): {total_km:,.2f}')

    # FDM / SLA
    fdm_types = Counter(r.get('O', '') for r in data)
    print('\nClassificação FDM (Col O):')
    for f, cnt in fdm_types.most_common():
        print(f'  {f}: {cnt}')

    # Missing OS
    missing_os = sum(1 for r in data if not r.get('I', '').strip())
    print(f'\nMissing OS (Col I): {missing_os}')

    # Missing Centro
    missing_centro = sum(1 for r in data if not r.get('AB', '').strip())
    print(f'Missing Centro (Col AB): {missing_centro}')
