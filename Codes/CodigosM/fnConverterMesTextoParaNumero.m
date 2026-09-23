let
    // -------------------------------------------------------------------
    // fnConverterMesTextoParaNumero
    // -----------------------------------------------------------------
    // Converte o nome do mês em português (ex.: "Julho") para seu
    // número correspondente (1 a 12). Necessária porque o parâmetro
    // PeriodoMes do pipeline está definido como TEXTO (nome do mês),
    // enquanto fnCalcularPeriodosAnteriores e fnRecalcularFDM operam
    // com o mês em formato NUMÉRICO (equivalente ao "iPerMes" do VBA,
    // que já vinha convertido via Left(Planilha6.Range("C9").Value,2)
    // na origem — aqui replicamos essa conversão de forma explícita).
    //
    // PARÂMETROS:
    //   mesTexto - nome do mês em português, com a mesma grafia usada
    //              no parâmetro PeriodoMes (ex.: "Julho")
    //
    // RETORNO:
    //   number — o número do mês (1 a 12)
    //
    // Lança erro explícito se o texto não corresponder a nenhum dos
    // 12 meses cadastrados (evita falha silenciosa/confusa mais
    // adiante no cálculo de períodos anteriores).
    // -------------------------------------------------------------------
    fnConverterMesTextoParaNumero = (mesTexto as text) as number =>
        let
            Meses = {
                "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
            },
            Indice = List.PositionOf(Meses, mesTexto)
        in
            if Indice = -1 then
                error Error.Record(
                    "MesInvalido",
                    "Mês não reconhecido: '" & mesTexto & "'. Esperado um dos 12 nomes de mês em português (ex.: 'Julho')."
                )
            else
                Indice + 1
in
    fnConverterMesTextoParaNumero
