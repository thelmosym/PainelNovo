let
    // -------------------------------------------------------------------
    // fnValidarLocalizacao
    // Verifica se a atividade esta entre as que permitem validacao de
    // dados entre localidades (usado para gerar o campo Localidade do
    // galpao de guarda / calculo de KM adicional).
    // -------------------------------------------------------------------
    fnValidarLocalizacao = (sAtividade as nullable text) as logical =>
        List.Contains({
            "DEVOLUCAO DE EMPRESTIMO",
            "COLETA DE EMBALAGEM",
            "COLETA DE ITEM AVULSO",
            "ENTREGA DE EMBALAGEM EXPRESSO",
            "ENTREGA DE EMBALAGEM NORMAL",
            "ENTREGA DE ITEM EXPRESSO",
            "ENTREGA DE ITEM NORMAL"
        }, sAtividade)
in
    fnValidarLocalizacao
