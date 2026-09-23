#!/usr/bin/env python3
"""
inspect_workbook.py
Utilitário para inspeção estática de pastas de trabalho Excel (.xlsm e .xlsx).
Utiliza exclusivamente a biblioteca padrão do Python (zipfile, xml.etree).
Não requer instalação de pacotes externos.

Uso:
    python inspect_workbook.py "caminho/para/arquivo.xlsm"
"""

import sys
import os
import zipfile
import xml.etree.ElementTree as ET

NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'rel': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

def inspect_excel(filepath):
    if not os.path.exists(filepath):
        print(f"Erro: Arquivo '{filepath}' não encontrado.")
        sys.exit(1)

    print(f"\n==================================================")
    print(f" INSPETOR DE WORKBOOK EXCEL")
    print(f" Arquivo: {os.path.basename(filepath)}")
    print(f" Tamanho: {os.path.getsize(filepath) / (1024*1024):.2f} MB")
    print(f"==================================================\n")

    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            namelist = set(z.namelist())

            # 1. Checar Presença de Macros VBA
            has_vba = "xl/vbaProject.bin" in namelist
            print(f"[*] Projeto VBA (Macros): {'PRESENTE (xl/vbaProject.bin)' if has_vba else 'NÃO DETECTADO'}")

            # 2. Listar Abas (Worksheets)
            if "xl/workbook.xml" in namelist:
                wb_xml = z.read("xl/workbook.xml")
                root = ET.fromstring(wb_xml)
                sheets = root.findall('.//main:sheet', NS)
                print(f"\n[*] Abas Encontradas ({len(sheets)}):")
                for s in sheets:
                    name = s.attrib.get('name', 'Sem Nome')
                    sheet_id = s.attrib.get('sheetId', '?')
                    state = s.attrib.get('state', 'visible')
                    print(f"    - [{state}] {name} (ID: {sheet_id})")

            # 3. Listar Tabelas (ListObjects)
            table_files = [f for f in namelist if f.startswith("xl/tables/table") and f.endswith(".xml")]
            print(f"\n[*] Tabelas Estruturadas ({len(table_files)}):")
            for tf in sorted(table_files):
                tbl_xml = z.read(tf)
                t_root = ET.fromstring(tbl_xml)
                tbl_name = t_root.attrib.get('name', 'Sem Nome')
                tbl_disp = t_root.attrib.get('displayName', tbl_name)
                ref = t_root.attrib.get('ref', '?')
                print(f"    - {tbl_disp} (Ref: {ref})")

            # 4. Listar Conexões (Power Query / OLEDB)
            if "xl/connections.xml" in namelist:
                conn_xml = z.read("xl/connections.xml")
                c_root = ET.fromstring(conn_xml)
                conns = c_root.findall('.//main:connection', NS)
                print(f"\n[*] Conexões de Dados / Power Query ({len(conns)}):")
                for c in conns:
                    c_name = c.attrib.get('name', 'Sem Nome')
                    c_type = c.attrib.get('type', '?')
                    print(f"    - {c_name} (Tipo: {c_type})")
            else:
                print(f"\n[*] Conexões Externas: Nenhuma detectada.")

            # 5. Modelos de Dados / Mashup (Power Query Package)
            has_pq_mashup = any("customXml/item" in f for f in namelist) or any("powerquery" in f.lower() for f in namelist)
            print(f"\n[*] Pacote Mashup/Power Query: {'Detectado no pacote' if has_pq_mashup else 'Estrutura padrão'}")

    except zipfile.BadZipFile:
        print(f"Erro: O arquivo '{filepath}' não é um arquivo Excel OpenXML válido.")
    except Exception as e:
        print(f"Erro durante a inspeção: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python inspect_workbook.py <caminho_arquivo_excel>")
    else:
        inspect_excel(sys.argv[1])
