let
    // -------------------------------------------------------------------
    // fnBuscarFDMAnterior
    // -------------------------------------------------------------------
    // Equivalente M da funcao VBA validar_FDMAnterior (modulo
    // fncTabelaB).
    //
    // O QUE FAZ:
    //   Busca o valor de FDM ja calculado e registrado historicamente
    //   para um periodo (mes/ano), contrato e item especificos, na
    //   tabela FDM_IAPFARQ (aba da tabelaB). Esse valor historico e
    //   necessario para o recalculo do FDM vigente conforme a regra
    //   de degradacao em cascata do item 8.4.3 do contrato (ver funcao
    //   fnRecalcularFDM, que consome este resultado).
    //
    // ? CORRECAO DE BUG APLICADA (em relacao a versao VBA original):
    //   O VBA original (validar_FDMAnterior) retornava 0 (valor padrao
    //   de uma variavel Double nao inicializada) quando o registro nao
    //   era encontrado - isso mascarava silenciosamente a ausencia do
    //   dado, podendo levar a um recalculo de FDM incorreto sem nenhum
    //   aviso. Esta versao M retorna NULL explicitamente nesse caso,
    //   permitindo que a funcao chamadora (fnRecalcularFDM) detecte a
    //   ausencia do dado e lance um erro claro em vez de tratar a
    //   ausencia como um FDM=0 valido.
    //
    // PARAMETROS:
    //   periodo  - texto no formato "AAAA#MM" (ex.: "2026#07")
    //   contrato - texto com o nome do contrato (ex.: "IRON-LT1")
    //   item     - numero identificando o servico (1 ou 2, conforme
    //              "Ref. FDM"/"Item contrato")
    //
    // RETORNO:
    //   nullable number - o FDM historico encontrado, ou null se nao
    //   houver nenhum registro correspondente a periodo+contrato+item
    //
    // DEPENDENCIA:
    //   RefFDMHistorico - consulta bufferizada da tabela FDM_IAPFARQ
    //                     (ver definicao da consulta RefFDMHistorico)
    // -------------------------------------------------------------------
    fnBuscarFDMAnterior = (periodo as text, contrato as text, item as number) as nullable number =>
        let
            // Filtra a tabela historica pela combinacao exata de
            // periodo + contrato + item (equivalente as 3 condicoes
            // combinadas com "And" no VBA original)
            Filtrado = Table.SelectRows(RefFDMHistorico, each
                [Medição] = periodo and [Contrato] = contrato and [#"Item contrato"] = item
            ),
            // Se encontrou exatamente o registro, retorna o valor de
            // FDM daquela linha; caso contrario, retorna null
            // (em vez do "0" silencioso do VBA original)
            Resultado = if Table.RowCount(Filtrado) > 0 then Filtrado{0}[FDM] else null
        in
            Resultado
in
    fnBuscarFDMAnterior
