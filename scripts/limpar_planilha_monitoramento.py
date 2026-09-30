#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
limpar_planilha_monitoramento.py
Utilitário para execução automatizada da rotina de limpeza e ajuste
da aba 'Monitoramento' para exatamente 100 linhas, preservando validações de dados.
"""

import os
import sys
import glob

def limpar_planilha(caminho_arquivo, silencioso=False):
    import win32com.client as win32

    caminho_absoluto = os.path.abspath(caminho_arquivo)
    if not os.path.exists(caminho_absoluto):
        print(f"[ERRO] Arquivo não encontrado: {caminho_absoluto}")
        return False

    nome_arquivo = os.path.basename(caminho_absoluto)
    print(f"\n[PROCESSANDO] {nome_arquivo}...")

    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        wb = excel.Workbooks.Open(caminho_absoluto)

        # 1. Obter aba Monitoramento
        ws = None
        for sheet in wb.Sheets:
            if sheet.Name.lower() == "monitoramento":
                ws = sheet
                break

        if ws is None:
            print(f"  [AVISO] Aba 'Monitoramento' não encontrada em {nome_arquivo}.")
            wb.Close(False)
            return False

        # 2. Obter ListObject
        tbl = None
        if ws.ListObjects.Count > 0:
            try:
                tbl = ws.ListObjects("Monitoramento")
            except Exception:
                tbl = ws.ListObjects(1)

        if tbl is None:
            print(f"  [AVISO] Nenhuma tabela estruturada encontrada na aba 'Monitoramento'.")
            wb.Close(False)
            return False

        linhas_iniciais = tbl.ListRows.Count
        print(f"  - Tabela: '{tbl.Name}' | Linhas atuais: {linhas_iniciais} | Intervalo: {tbl.Range.Address}")

        # 3. Mapear validações de dados existentes
        validacoes = {}
        for c in range(1, tbl.ListColumns.Count + 1):
            col_rng = tbl.ListColumns(c).DataBodyRange
            if col_rng is not None:
                first_cell = col_rng.Cells(1, 1)
                try:
                    v = first_cell.Validation
                    if v.Type != 0:
                        validacoes[c] = {
                            "Type": v.Type,
                            "AlertStyle": v.AlertStyle,
                            "Operator": v.Operator,
                            "Formula1": v.Formula1,
                            "Formula2": getattr(v, "Formula2", ""),
                            "IgnoreBlank": v.IgnoreBlank,
                            "InCellDropdown": v.InCellDropdown,
                            "ShowInput": v.ShowInput,
                            "ShowError": v.ShowError,
                            "InputTitle": v.InputTitle,
                            "InputMessage": v.InputMessage,
                            "ErrorTitle": v.ErrorTitle,
                            "ErrorMessage": v.ErrorMessage,
                        }
                except Exception:
                    pass

        # Fórmulas padrão de fallback (TabelaA)
        fallback_formulas = {
            3: "=TabelaA!$B$2:$B$5",
            4: "=TabelaA!$E$2:$E$6",
            5: "=TabelaA!$AD$2:$AD$57",
            6: "=TabelaA!$AO$2:$AO$35",
            7: "=TabelaA!$AQ$2:$AQ$53",
            8: "=TabelaA!$AH$2:$AH$3",
            25: "=TabelaA!$AW$2:$AW$3"
        }

        # 4. Ajustar linhas da tabela para 100 linhas
        QTD_DESEJADA = 100
        if linhas_iniciais > QTD_DESEJADA:
            rng_del = ws.Range(tbl.ListRows(QTD_DESEJADA + 1).Range, tbl.ListRows(linhas_iniciais).Range)
            rng_del.Delete() # xlShiftUp
            print(f"  - Excesso de {linhas_iniciais - QTD_DESEJADA} linhas removido.")
        elif linhas_iniciais < QTD_DESEJADA:
            while tbl.ListRows.Count < QTD_DESEJADA:
                tbl.ListRows.Add()
            print(f"  - Linhas adicionadas até completar {QTD_DESEJADA} linhas.")

        # 5. Limpeza de dados e formatações
        if tbl.DataBodyRange is not None:
            tbl.DataBodyRange.ClearContents()
            tbl.DataBodyRange.FormatConditions.Delete()
            tbl.DataBodyRange.ClearFormats()
            tbl.DataBodyRange.RowHeight = 20
            tbl.DataBodyRange.VerticalAlignment = -4108 # xlCenter

        # 6. Limpar resíduos abaixo da tabela na planilha
        fim_tabela = tbl.Range.Row + tbl.Range.Rows.Count - 1
        last_row = ws.Cells.SpecialCells(11).Row # xlCellTypeLastCell
        if last_row > fim_tabela:
            ws.Rows(f"{fim_tabela + 1}:{last_row}").Delete()
            print(f"  - Linhas residuais abaixo da tabela ({fim_tabela + 1} a {last_row}) eliminadas.")

        # 7. Reaplicar e garantir validações de dados
        for c in range(1, tbl.ListColumns.Count + 1):
            col_rng = tbl.ListColumns(c).DataBodyRange
            if col_rng is not None:
                info = validacoes.get(c)
                if not info and c in fallback_formulas:
                    info = {
                        "Type": 3,
                        "AlertStyle": 1,
                        "Operator": 1,
                        "Formula1": fallback_formulas[c],
                        "Formula2": "",
                        "IgnoreBlank": True,
                        "InCellDropdown": True,
                        "ShowInput": True,
                        "ShowError": True,
                        "InputTitle": "",
                        "InputMessage": "",
                        "ErrorTitle": "",
                        "ErrorMessage": ""
                    }
                if info:
                    try:
                        col_rng.Validation.Delete()
                        if info.get("Formula2"):
                            col_rng.Validation.Add(info["Type"], info["AlertStyle"], info["Operator"], info["Formula1"], info["Formula2"])
                        else:
                            col_rng.Validation.Add(info["Type"], info["AlertStyle"], info["Operator"], info["Formula1"])
                        col_rng.Validation.IgnoreBlank = info["IgnoreBlank"]
                        col_rng.Validation.InCellDropdown = info["InCellDropdown"]
                        col_rng.Validation.ShowInput = info["ShowInput"]
                        col_rng.Validation.ShowError = info["ShowError"]
                    except Exception as e_val:
                        pass

        # 8. Formatações complementares de colunas
        try:
            if tbl.ListColumns.Count >= 2:
                tbl.ListColumns(2).DataBodyRange.NumberFormat = "dd/mm/aaaa (ddd)"
                tbl.ListColumns(2).DataBodyRange.HorizontalAlignment = -4108
            for c_align in [3, 4, 5, 8, 15, 20, 25]:
                if tbl.ListColumns.Count >= c_align:
                    tbl.ListColumns(c_align).DataBodyRange.HorizontalAlignment = -4108
            for c_date in [12, 17, 18]:
                if tbl.ListColumns.Count >= c_date:
                    tbl.ListColumns(c_date).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm"
        except Exception:
            pass

        # Ativar estilos e posicionar em A3
        try:
            tbl.ShowTableStyleRowStripes = True
            tbl.ShowAutoFilter = True
            ws.Activate()
            ws.Range("A3").Select()
        except Exception:
            pass

        wb.Save()
        wb.Close(False)
        print(f"  [OK] Concluído com sucesso! Tabela ajustada para 100 linhas (A3:AK102).")
        return True

    except Exception as e:
        print(f"  [ERRO] Falha durante o processamento de {nome_arquivo}: {e}")
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        return False
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python limpar_planilha_monitoramento.py <caminho_arquivo_ou_pasta>")
        sys.exit(1)

    alvo = sys.argv[1]
    if os.path.isfile(alvo):
        limpar_planilha(alvo)
    elif os.path.isdir(alvo):
        padrao = os.path.join(alvo, "Monitoramento individual_*.xlsm")
        arquivos = glob.glob(padrao)
        if not arquivos:
            print(f"Nenhum arquivo encontrado no padrão: {padrao}")
        for arq in sorted(arquivos):
            if "~$" not in arq:
                limpar_planilha(arq)
