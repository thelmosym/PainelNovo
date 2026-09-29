let
    // -------------------------------------------------------------------
    // fnCalcularPeriodosAnteriores
    // -------------------------------------------------------------------
    // Equivalente M do trecho de calculo de periodo anterior presente
    // dentro de btMCGuardaExterna/btMC_GE_individual (modulo
    // subMCGuardaExterna) - usado para localizar o FDM historico dos
    // 2 meses anteriores ao periodo vigente (item 8.4.3 do contrato).
    //
    // O QUE FAZ:
    //   Calcula as chaves de periodo no formato "AAAA#MM" para o mes
    //   anterior (PerAnt1) e para o mes anterior a esse (PerAnt2),
    //   tratando corretamente a VIRADA DE ANO:
    //
    //   - PerAnt1 (1 mes atras):
    //       Se o mes vigente for Janeiro (mes=1), o mes anterior e
    //       Dezembro do ANO ANTERIOR.
    //       Caso contrario, e o mes vigente - 1, no mesmo ano.
    //
    //   - PerAnt2 (2 meses atras):
    //       Se o mes vigente for Janeiro (mes=1), 2 meses atras e
    //       Novembro do ANO ANTERIOR.
    //       Se o mes vigente for Fevereiro (mes=2), 2 meses atras e
    //       Dezembro do ANO ANTERIOR.
    //       Caso contrario, e o mes vigente - 2, no mesmo ano.
    //
    // FORMATO DE SAIDA:
    //   Cada periodo e formatado como "AAAA#MM" (ex.: "2026#07"), com
    //   o mes sempre em 2 digitos (zero a esquerda quando necessario),
    //   compativel com a chave usada na tabela historica FDM_IAPFARQ
    //   (coluna "Medicao").
    //
    // PARAMETROS:
    //   mes - numero do mes vigente (1 a 12)
    //   ano - numero do ano vigente (ex.: 2026)
    //
    // RETORNO:
    //   Um record com 2 campos: [PerAnt1, PerAnt2]
    // -------------------------------------------------------------------
    fnCalcularPeriodosAnteriores = (mes as number, ano as number) as record =>
        let
            // Calcula o periodo de 1 mes atras, tratando virada de ano
            PerAnt1 = if mes - 1 = 0
                then Text.From(ano - 1) & "#12"
                else Text.From(ano) & "#" & Text.PadStart(Text.From(mes - 1), 2, "0"),
            // Calcula o periodo de 2 meses atras, tratando virada de ano
            // (2 casos possiveis: mes vigente = Janeiro ou Fevereiro)
            PerAnt2 = if mes - 2 = -1
                then Text.From(ano - 1) & "#11"
                else if mes - 2 = 0
                then Text.From(ano - 1) & "#12"
                else Text.From(ano) & "#" & Text.PadStart(Text.From(mes - 2), 2, "0")
        in
            [ PerAnt1 = PerAnt1, PerAnt2 = PerAnt2 ]
in
    fnCalcularPeriodosAnteriores
