#!/usr/bin/env bash
#===================================================================================================================
# SCRIPT BASH: criar_macro_limpar_monitoramento.sh
# OBJETIVO: Gerar, validar e aplicar a macro VBA (modLimparMonitoramento.bas) para limpeza completa da aba
#           "Monitoramento" na planilha de Monitoramento Individual, mantendo 100 linhas na tabela,
#           removendo o excesso e preservando 100% da validação de dados das células.
#
# USO:
#   ./scripts/criar_macro_limpar_monitoramento.sh                 # Gera/valida o arquivo .bas em CP1252/CRLF
#   ./scripts/criar_macro_limpar_monitoramento.sh --aplicar       # Executa a limpeza em todas as planilhas
#   ./scripts/criar_macro_limpar_monitoramento.sh --aplicar "Monitoramento/Monitoramento individual_A4UU_v3.0.xlsm"
#   ./scripts/criar_macro_limpar_monitoramento.sh --help          # Exibe manual de uso e parâmetros
#===================================================================================================================

set -e

DIR_RAIZ="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIR_VBA="${DIR_RAIZ}/Codes/VBA"
DIR_VBA_NOVAS="${DIR_VBA}/Funções Novas"
ARQ_BAS="${DIR_VBA}/modLimparMonitoramento.bas"
ARQ_BAS_NOVAS="${DIR_VBA_NOVAS}/modLimparMonitoramento.bas"
SCRIPT_ENCODING="${DIR_RAIZ}/.agents/skills/excel-vba-powerquery/scripts/fix_vba_encoding.py"
SCRIPT_PS_LIMPEZA="${DIR_RAIZ}/scripts/limpar_planilha_monitoramento.ps1"
DIR_MONITORAMENTO="${DIR_RAIZ}/Monitoramento"

# Cores para terminal
VERDE='\033[0;32m'
AMARELO='\033[1;33m'
AZUL='\033[0;34m'
VERMELHO='\033[0;31m'
NC='\033[0m' # No Color

exibir_ajuda() {
    cat << EOF
---------------------------------------------------------------------------------------------------
 UTILITÁRIO BASH: MACRO DE LIMPEZA E REDIMENSIONAMENTO (100 LINHAS) - ABA MONITORAMENTO
---------------------------------------------------------------------------------------------------
Este script gerencia o módulo VBA 'modLimparMonitoramento.bas' e a automação de limpeza para as
planilhas de Monitoramento Individual (A4UU, DPBR, GPZ1, GQ6S, S2IJ).

FUNCIONALIDADES DA MACRO:
  1. Limpa todos os dados digitados na tabela da aba 'Monitoramento'.
  2. Limpa formatações manuais e regras duplicadas/fragmentadas de formatação condicional.
  3. Redimensiona a tabela para conter exatamente 100 linhas de dados (A3:AK102), eliminando o excesso.
  4. Exclui linhas residuais e sujeiras abaixo da tabela (linha 103 em diante).
  5. Mapeia e reaplica 100% das regras de Validação de Dados (listas suspensas de Situação,
     Contrato, UF, Descrição da Atividade, Item, Aplicação, OS disponibilizada).

PARÂMETROS:
  -g, --gerar              Gera e audita os arquivos .bas em Windows-1252 (CP1252) com quebras CRLF.
  -a, --aplicar [ARQUIVO]  Executa a limpeza diretamente no arquivo informado ou em todos da pasta.
  -h, --help               Exibe esta mensagem de ajuda.

EXEMPLOS DE USO:
  ./scripts/criar_macro_limpar_monitoramento.sh --gerar
  ./scripts/criar_macro_limpar_monitoramento.sh --aplicar "Monitoramento/Monitoramento individual_A4UU_v3.0.xlsm"
  ./scripts/criar_macro_limpar_monitoramento.sh --aplicar
---------------------------------------------------------------------------------------------------
EOF
}

gerar_macro() {
    echo -e "${AZUL}[*] Gerando e auditando arquivo VBA: modLimparMonitoramento.bas...${NC}"

    mkdir -p "${DIR_VBA}"
    mkdir -p "${DIR_VBA_NOVAS}"

    if [ ! -f "${ARQ_BAS}" ]; then
        echo -e "${VERMELHO}[ERRO] Arquivo base '${ARQ_BAS}' não encontrado.${NC}"
        exit 1
    fi

    # Sincroniza cópia na pasta 'Funções Novas'
    cp -f "${ARQ_BAS}" "${ARQ_BAS_NOVAS}"

    # Valida e converte encoding para CP1252 (ANSI) e CRLF
    if [ -f "${SCRIPT_ENCODING}" ]; then
        echo -e "${AMARELO}[*] Auditando codificação Windows-1252 (ANSI) e terminações CRLF...${NC}"
        python "${SCRIPT_ENCODING}" "${ARQ_BAS}"
        python "${SCRIPT_ENCODING}" "${ARQ_BAS_NOVAS}"
    fi

    echo -e "${VERDE}[OK] Módulo VBA pronto para uso:${NC}"
    echo -e "     -> ${ARQ_BAS}"
    echo -e "     -> ${ARQ_BAS_NOVAS}"
}

aplicar_limpeza() {
    local alvo="$1"
    if [ -z "${alvo}" ]; then
        alvo="${DIR_MONITORAMENTO}"
    fi

    echo -e "${AZUL}[*] Iniciando rotina de limpeza na planilha: ${alvo}${NC}"
    if [ ! -f "${SCRIPT_PS_LIMPEZA}" ]; then
        echo -e "${VERMELHO}[ERRO] Script PowerShell '${SCRIPT_PS_LIMPEZA}' não encontrado.${NC}"
        exit 1
    fi

    # Executa PowerShell nativo do Windows para automação Excel COM
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(cygpath -w "${SCRIPT_PS_LIMPEZA}")" -Caminho "$(cygpath -w "${alvo}")"
}

# Processamento de argumentos
case "$1" in
    -h|--help)
        exibir_ajuda
        exit 0
        ;;
    -g|--gerar)
        gerar_macro
        exit 0
        ;;
    -a|--aplicar)
        gerar_macro
        aplicar_limpeza "$2"
        exit 0
        ;;
    *)
        if [ -n "$1" ]; then
            echo -e "${AMARELO}[AVISO] Opção '$1' desconhecida. Executando modo padrão (geração/auditoria).${NC}\n"
        fi
        gerar_macro
        echo ""
        echo -e "${VERDE}===================================================================================${NC}"
        echo -e "${VERDE} COMO IMPORTAR E USAR A MACRO NO EXCEL:${NC}"
        echo -e " 1. Abra a sua planilha (ex: Monitoramento individual_A4UU_v3.0.xlsm)."
        echo -e " 2. Pressione ${AMARELO}ALT + F11${NC} para abrir o Editor do VBA (VBE)."
        echo -e " 3. No menu superior, clique em ${AMARELO}Arquivo > Importar Arquivo... (Ctrl + M)${NC}."
        echo -e " 4. Selecione o arquivo: ${AMARELO}${ARQ_BAS}${NC}."
        echo -e " 5. Feche o VBE e volte ao Excel."
        echo -e " 6. Pressione ${AMARELO}ALT + F8${NC}, selecione ${AMARELO}LimparMonitoramento${NC} e clique em Executar."
        echo -e "    (Ou vincule a macro a um botão de sua escolha na aba Menu ou Monitoramento)."
        echo -e "${VERDE}===================================================================================${NC}"
        echo -e " DICA: Para aplicar a limpeza diretamente via terminal bash em um arquivo:"
        echo -e "   ${AMARELO}./scripts/criar_macro_limpar_monitoramento.sh --aplicar \"caminho/da/planilha.xlsm\"${NC}"
        echo -e "${VERDE}===================================================================================${NC}"
        ;;
esac
