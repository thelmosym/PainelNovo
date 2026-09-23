let
    // -------------------------------------------------------------------
    // fnCalcularPeriodosAnteriores
    // -----------------------------------------------------------------
    // Equivalente M do trecho de cálculo de período anterior presente
    // dentro de btMCGuardaExterna/btMC_GE_individual (módulo
    // subMCGuardaExterna) — usado para localizar o FDM histórico dos
    // 2 meses anteriores ao período vigente (item 8.4.3 do contrato).
    //
    // O QUE FAZ:
    //   Calcula as chaves de período no formato "AAAA#MM" para o mês
    //   anterior (PerAnt1) e para o mês anterior a esse (PerAnt2),
    //   tratando corretamente a VIRADA DE ANO:
    //
    //   - PerAnt1 (1 mês atrás):
    //       Se o mês vigente for Janeiro (mes=1), o mês anterior é
    //       Dezembro do ANO ANTERIOR.
    //       Caso contrário, é o mês vigente - 1, no mesmo ano.
    //
    //   - PerAnt2 (2 meses atrás):
    //       Se o mês vigente for Janeiro (mes=1), 2 meses atrás é
    //       Novembro do ANO ANTERIOR.
    //       Se o mês vigente for Fevereiro (mes=2), 2 meses atrás é
    //       Dezembro do ANO ANTERIOR.
    //       Caso contrário, é o mês vigente - 2, no mesmo ano.
    //
    // FORMATO DE SAÍDA:
    //   Cada período é formatado como "AAAA#MM" (ex.: "2026#07"), com
    //   o mês sempre em 2 dígitos (zero à esquerda quando necessário),
    //   compatível com a chave usada na tabela histórica FDM_IAPFARQ
    //   (coluna "Medição").
    //
    // PARÂMETROS:
    //   mes - número do mês vigente (1 a 12)
    //   ano - número do ano vigente (ex.: 2026)
    //
    // RETORNO:
    //   Um record com 2 campos: [PerAnt1, PerAnt2]
    // -------------------------------------------------------------------
    fnCalcularPeriodosAnteriores = (mes as number, ano as number) as record =>
        let
            // Calcula o período de 1 mês atrás, tratando virada de ano
            PerAnt1 = if mes - 1 = 0
                then Text.From(ano - 1) & "#12"
                else Text.From(ano) & "#" & Text.PadStart(Text.From(mes - 1), 2, "0"),

            // Calcula o período de 2 meses atrás, tratando virada de ano
            // (2 casos possíveis: mês vigente = Janeiro ou Fevereiro)
            PerAnt2 = if mes - 2 = -1
                then Text.From(ano - 1) & "#11"
                else if mes - 2 = 0
                then Text.From(ano - 1) & "#12"
                else Text.From(ano) & "#" & Text.PadStart(Text.From(mes - 2), 2, "0")
        in
            [ PerAnt1 = PerAnt1, PerAnt2 = PerAnt2 ]
in
    fnCalcularPeriodosAnteriores
