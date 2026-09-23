let
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
                        [ Classificacao = null, StatusPrazo = null, Log = "Prazo combinado vazio; " ]
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