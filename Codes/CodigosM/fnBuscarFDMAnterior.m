let
    // -------------------------------------------------------------------
    // fnBuscarFDMAnterior
    // -----------------------------------------------------------------
    // Equivalente M da função VBA validar_FDMAnterior (módulo
    // fncTabelaB).
    //
    // O QUE FAZ:
    //   Busca o valor de FDM já calculado e registrado historicamente
    //   para um período (mês/ano), contrato e item específicos, na
    //   tabela FDM_IAPFARQ (aba da tabelaB). Esse valor histórico é
    //   necessário para o recálculo do FDM vigente conforme a regra
    //   de degradação em cascata do item 8.4.3 do contrato (ver função
    //   fnRecalcularFDM, que consome este resultado).
    //
    // ? CORREÇÃO DE BUG APLICADA (em relação à versão VBA original):
    //   O VBA original (validar_FDMAnterior) retornava 0 (valor padrão
    //   de uma variável Double não inicializada) quando o registro não
    //   era encontrado — isso mascarava silenciosamente a ausência do
    //   dado, podendo levar a um recálculo de FDM incorreto sem nenhum
    //   aviso. Esta versão M retorna NULL explicitamente nesse caso,
    //   permitindo que a função chamadora (fnRecalcularFDM) detecte a
    //   ausência do dado e lance um erro claro em vez de tratar a
    //   ausência como um FDM=0 válido.
    //
    // PARÂMETROS:
    //   periodo  - texto no formato "AAAA#MM" (ex.: "2026#07")
    //   contrato - texto com o nome do contrato (ex.: "IRON-LT1")
    //   item     - número identificando o serviço (1 ou 2, conforme
    //              "Ref. FDM"/"Item contrato")
    //
    // RETORNO:
    //   nullable number — o FDM histórico encontrado, ou null se não
    //   houver nenhum registro correspondente a período+contrato+item
    //
    // DEPENDÊNCIA:
    //   RefFDMHistorico - consulta bufferizada da tabela FDM_IAPFARQ
    //                     (ver definição da consulta RefFDMHistorico)
    // -------------------------------------------------------------------
    fnBuscarFDMAnterior = (periodo as text, contrato as text, item as number) as nullable number =>
        let
            // Filtra a tabela histórica pela combinação exata de
            // período + contrato + item (equivalente às 3 condições
            // combinadas com "And" no VBA original)
            Filtrado = Table.SelectRows(RefFDMHistorico, each
                [Medição] = periodo and [Contrato] = contrato and [#"Item contrato"] = item
            ),

            // Se encontrou exatamente o registro, retorna o valor de
            // FDM daquela linha; caso contrário, retorna null
            // (em vez do "0" silencioso do VBA original)
            Resultado = if Table.RowCount(Filtrado) > 0 then Filtrado{0}[FDM] else null
        in
            Resultado
in
    fnBuscarFDMAnterior
