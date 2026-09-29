let
    // -------------------------------------------------------------------
    // fnClassificarFDM
    // -------------------------------------------------------------------
    // Equivalente M dos procedimentos VBA analiseDeDados()/
    // preenchimentoPrazoFDM() (modulo Planilha7) - classifica cada
    // linha do painel em uma categoria de FDM (Normal/Expresso/Isento/
    // FDM definido) e um Status de Prazo (NP/FP), com base no SLA
    // calculado e na referencia de prazo da atividade.
    //
    // O QUE FAZ:
    //   - calculoFDM = "S" -> linha ja e duplicata de FDM (mesma
    //     atividade/solicitacao/OS ja vista antes) -> "FDM definido".
    //   - refFDM = null/0 -> atividade fora do escopo de medicao de
    //     FDM -> "Isento".
    //   - refPrazo = "Tabela" -> valida campos obrigatorios (contrato,
    //     datas, municipio, UF, grupo); calcula o SLA via
    //     fnCalcularPrazoSLA e classifica em Normal/Expresso e NP/FP,
    //     com regras especificas para IPF/ABF/RFT.
    //   - refPrazo = "Fiscalizacao" -> compara prazo combinado x data
    //     de fechamento; se prazo combinado estiver vazio, assume
    //     Normal/NP (ajuste sem log de pendencia).
    //   - Qualquer outro valor de refPrazo -> log de inconsistencia.
    //
    // AJUSTES APLICADOS:
    //   - refPrazo = "Fiscalizacao" e prazoCombinado = null resulta em
    //     Classificacao = "Normal", StatusPrazo = "NP", Log = "".
    //   - Mensagens de observacao/log padronizadas sem acentos para
    //     evitar problemas de incompatibilidade de codificacao.
    //   - Suporte opcional a tabela de feriados para o calculo de SLA.
    //
    // PARAMETROS:
    //   calculoFDM, refFDM, refPrazo, descricaoAtividade, contrato,
    //   dataAbertura, horaAbertura, dataFechamento, prazoCombinado,
    //   municipio, uf, grupo, dPrazo, aplicacao, optional tbFeriados
    //
    // RETORNO:
    //   record [ Classificacao, StatusPrazo, Log ]
    //
    // DEPENDENCIA:
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
        aplicacao as nullable text,
        optional tbFeriados as table
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
                                dtReSLA = fnCalcularPrazoSLA(dPrazo, dataAbertura, horaAbertura, municipio, uf, tbFeriados),
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
            else if refPrazo = "Fiscalizacao" or refPrazo = "Fiscalização" or (refPrazo <> null and Text.StartsWith(Text.Upper(refPrazo), "FISCALIZA")) then
                if dataFechamento = null then
                    [ Classificacao = null, StatusPrazo = null, Log = "D. fechamento vazio; " ]
                else if prazoCombinado = null then
                    // Ajuste aplicado nesta conversa: assume Normal/NP,
                    // sem registrar pendencia de log
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
