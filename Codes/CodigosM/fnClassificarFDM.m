let
    // -------------------------------------------------------------------
    // fnClassificarFDM
    // -----------------------------------------------------------------
    // Equivalente M dos procedimentos VBA analiseDeDados()/
    // preenchimentoPrazoFDM() (módulo Planilha7) — classifica cada
    // linha do painel em uma categoria de FDM (Normal/Expresso/Isento/
    // FDM definido) e um Status de Prazo (NP/FP), com base no SLA
    // calculado e na referência de prazo da atividade.
    //
    // O QUE FAZ:
    //   - calculoFDM = "S" ? linha já é duplicata de FDM (mesma
    //     atividade/solicitação/OS já vista antes) ? "FDM definido".
    //   - refFDM = null/0 ? atividade fora do escopo de medição de
    //     FDM ? "Isento".
    //   - refPrazo = "Tabela" ? valida campos obrigatórios (contrato,
    //     datas, município, UF, grupo); calcula o SLA via
    //     fnCalcularPrazoSLA e classifica em Normal/Expresso e NP/FP,
    //     com regras específicas para IPF/ABF/RFT.
    //   - refPrazo = "Fiscalização" ? compara prazo combinado x data
    //     de fechamento; se prazo combinado estiver vazio, assume
    //     Normal/NP (ajuste aplicado nesta conversa, sem log de
    //     pendência).
    //   - Qualquer outro valor de refPrazo ? log de inconsistência.
    //
    // ? AJUSTE APLICADO NESTA CONVERSA:
    //   refPrazo = "Fiscalização" e prazoCombinado = null agora resulta
    //   em Classificacao = "Normal", StatusPrazo = "NP", Log = "" (em
    //   vez do comportamento original de null/null/mensagem de log).
    //
    // PARÂMETROS:
    //   calculoFDM, refFDM, refPrazo, descricaoAtividade, contrato,
    //   dataAbertura, horaAbertura, dataFechamento, prazoCombinado,
    //   municipio, uf, grupo, dPrazo, aplicacao
    //
    // RETORNO:
    //   record [ Classificacao, StatusPrazo, Log ]
    //
    // DEPENDÊNCIA:
    //   fnCalcularPrazoSLA (usada dentro do ramo refPrazo="Tabela")
    // -------------------------------------------------------------------
    fnClassificarFDM = (
        calculoFDM as text,
        refFDM as any,
        refPrazo as nullable text,
        descricaoAtividade as text,
        contrato as nullable text,
        dataAbertura as nullable date,
        horaAbertura as nullable time,
        dataFechamento as nullable datetime,
        prazoCombinado as nullable datetime,
        municipio as nullable text,
        uf as nullable text,
        grupo as nullable text,
        dPrazo as nullable number,
        aplicacao as nullable text
    ) as record =>
    let
        Resultado =
            if calculoFDM = "S" then
                [ Classificacao = "FDM definido", StatusPrazo = null, Log = "" ]

            else if refFDM = null or refFDM = 0 then
                [ Classificacao = "Isento", StatusPrazo = null, Log = "" ]

            else if refPrazo = "Tabela" then
                let
                    ValidacaoCampoVazio =
                        if contrato = null or contrato = "" then "Contrato vazio; "
                        else if dataAbertura = null then "D. abertura vazio; "
                        else if dataFechamento = null then "D. fechamento vazio; "
                        else if municipio = null or municipio = "" then "Municipio vazio; "
                        else if uf = null or uf = "" then "UF vazio; "
                        else if grupo = null or grupo = "Nao encontrado" then "Municipio sem grupo valido para a UF; "
                        else "",

                    ResultadoTabela =
                        if ValidacaoCampoVazio <> "" then
                            [ Classificacao = null, StatusPrazo = null, Log = ValidacaoCampoVazio ]
                        else if dPrazo = 0 or dPrazo = null then
                            [ Classificacao = null, StatusPrazo = null, Log = "FDM nao permitido para este grupo/atividade; " ]
                        else
                            let
                                dtReSLA = fnCalcularPrazoSLA(dPrazo, dataAbertura, horaAbertura),
                                DentroDoPrazo = dtReSLA >= dataFechamento,
                                EhExpresso = Text.Contains(Text.Upper(descricaoAtividade), "EXPRESSO"),
                                ResultadoFinal =
                                    if DentroDoPrazo then
                                        if EhExpresso then
                                            if aplicacao = "IPF" then
                                                [ Classificacao = "Normal", StatusPrazo = "FP", Log = "" ]
                                            else
                                                [ Classificacao = "Expresso", StatusPrazo = "NP", Log = "" ]
                                        else
                                            [ Classificacao = "Normal", StatusPrazo = "NP", Log = "" ]
                                    else
                                        if aplicacao = "ABF" or aplicacao = "RFT" then
                                            [ Classificacao = if EhExpresso then "Expresso" else "Normal", StatusPrazo = "NP", Log = "" ]
                                        else
                                            [ Classificacao = "Normal", StatusPrazo = "FP", Log = "" ]
                            in
                                ResultadoFinal
                in
                    ResultadoTabela

            else if refPrazo = "Fiscalização" then
                if dataFechamento = null then
                    [ Classificacao = null, StatusPrazo = null, Log = "D. fechamento vazio; " ]
                else if prazoCombinado = null then
                    // Ajuste aplicado nesta conversa: assume Normal/NP,
                    // sem registrar pendência de log
                    [ Classificacao = "Normal", StatusPrazo = "NP", Log = "" ]
                else if prazoCombinado >= dataFechamento then
                    [ Classificacao = "Normal", StatusPrazo = "NP", Log = "" ]
                else
                    [ Classificacao = "Normal", StatusPrazo = "FP", Log = "" ]

            else
                [ Classificacao = null, StatusPrazo = null, Log = "Ref. prazo nao reconhecida; " ]
    in
        Resultado
in
    fnClassificarFDM
