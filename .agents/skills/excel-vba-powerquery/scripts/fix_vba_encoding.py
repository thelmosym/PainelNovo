#!/usr/bin/env python3
"""
fix_vba_encoding.py
Utilitário para auditoria e conversão de arquivos de código VBA (.bas, .cls, .frm)
para a codificação Windows-1252 (CP1252 / ANSI) com quebras de linha CRLF.

Garante que textos e acentuações em português (á, é, í, ó, ú, ç, ã, etc.) apareçam
perfeitamente legíveis dentro do Visual Basic Editor (VBE) do Microsoft Excel,
eliminando mojibake e substituindo emojis de 4 bytes por tags textuais equivalentes.

Uso:
    python fix_vba_encoding.py "caminho/para/arquivo.bas"
    python fix_vba_encoding.py "d:/PainelNovo/Codes/VBA"
"""

import sys
import os

EMOJI_REPLACEMENTS = {
    "🟡": "[AMARELO]",
    "🟢": "[VERDE]",
    "🔴": "[VERMELHO]",
    "⚪": "[BRANCO]",
    "✅": "[OK]",
    "⚠️": "[AVISO]",
    "❌": "[ERRO]",
    "ℹ️": "[INFO]",
    "📊": "[DASHBOARD]",
    "📁": "[PASTA]",
    "💾": "[SALVO]",
    "ðŸŸ¡": "[AMARELO]",
    "ðŸŸ¢": "[VERDE]",
    "ðŸ”´": "[VERMELHO]",
    "âœ…": "[OK]",
    "âš ï¸": "[AVISO]",
    "âŒ": "[ERRO]"
}

VBA_EXTENSIONS = {".bas", ".cls", ".frm", ".vba"}

def sanitize_and_convert_file(filepath, dry_run=False):
    with open(filepath, "rb") as f:
        raw_bytes = f.read()

    # Detectar codificação atual
    detected_encoding = None
    text = None

    if raw_bytes.startswith(b"\xef\xbb\xbf"):
        try:
            text = raw_bytes[3:].decode("utf-8")
            detected_encoding = "UTF-8-BOM"
        except UnicodeDecodeError:
            pass

    if text is None:
        try:
            text = raw_bytes.decode("utf-8")
            detected_encoding = "UTF-8"
        except UnicodeDecodeError:
            pass

    if text is None:
        try:
            text = raw_bytes.decode("cp1252")
            detected_encoding = "CP1252"
        except UnicodeDecodeError:
            try:
                text = raw_bytes.decode("latin-1")
                detected_encoding = "Latin-1"
            except Exception as e:
                print(f"[!] Falha ao decodificar {filepath}: {e}")
                return False

    # Substituir emojis e mojibake por tags limpas
    replacements_made = 0
    for emoji, replacement in EMOJI_REPLACEMENTS.items():
        if emoji in text:
            count = text.count(emoji)
            text = text.replace(emoji, replacement)
            replacements_made += count

    # Padronizar terminações de linha para CRLF
    normalized_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    crlf_text = "\r\n".join(normalized_lines)

    # Codificar em CP1252
    try:
        final_bytes = crlf_text.encode("cp1252")
    except UnicodeEncodeError as e:
        # Se contiver algum caractere exótico fora do CP1252, substitui por equivalente aproximado
        final_bytes = crlf_text.encode("cp1252", errors="replace")
        print(f"[Aviso] Caracteres fora do CP1252 foram substituídos em {os.path.basename(filepath)}.")

    has_changed = (final_bytes != raw_bytes) or (detected_encoding != "CP1252")

    if has_changed:
        if not dry_run:
            with open(filepath, "wb") as f:
                f.write(final_bytes)
        print(f"[CONVERTIDO] {os.path.basename(filepath)}: De {detected_encoding} para CP1252/CRLF ({replacements_made} emojis ajustados).")
    else:
        print(f"[OK] {os.path.basename(filepath)}: Já está em CP1252/CRLF.")

    return True

def process_path(target_path, dry_run=False):
    if not os.path.exists(target_path):
        print(f"Erro: Caminho '{target_path}' não encontrado.")
        return

    if os.path.isfile(target_path):
        sanitize_and_convert_file(target_path, dry_run)
    else:
        for root, dirs, files in os.walk(target_path):
            for file in sorted(files):
                ext = os.path.splitext(file)[1].lower()
                if ext in VBA_EXTENSIONS:
                    full_path = os.path.join(root, file)
                    sanitize_and_convert_file(full_path, dry_run)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        default_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../Codes/VBA"))
        print(f"Nenhum caminho especificado. Processando pasta padrão: {default_dir}")
        process_path(default_dir)
    else:
        process_path(sys.argv[1])
