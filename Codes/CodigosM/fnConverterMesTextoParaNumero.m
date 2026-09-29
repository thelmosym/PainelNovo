let
    // -------------------------------------------------------------------
    // fnConverterMesTextoParaNumero
    // -------------------------------------------------------------------
    // Converte o nome do mes em portugues (ex.: "Julho") para seu
    // numero correspondente (1 a 12). Necessaria porque o parametro
    // PeriodoMes do pipeline esta definido como TEXTO (nome do mes),
    // enquanto fnCalcularPeriodosAnteriores e fnRecalcularFDM operam
    // com o mes em formato NUMERICO (equivalente ao "iPerMes" do VBA,
    // que ja vinha convertido via Left(Planilha6.Range("C9").Value,2)
    // na origem - aqui replicamos essa conversao de forma explicita).
    //
    // PARAMETROS:
    //   mesTexto - nome do mes em portugues, com a mesma grafia usada
    //              no parametro PeriodoMes (ex.: "Julho")
    //
    // RETORNO:
    //   number - o numero do mes (1 a 12)
    //
    // Lanca erro explicito se o texto nao corresponder a nenhum dos
    // 12 meses cadastrados (evita falha silenciosa/confusa mais
    // adiante no calculo de periodos anteriores).
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
