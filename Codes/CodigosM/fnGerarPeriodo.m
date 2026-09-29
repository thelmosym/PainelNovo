let
    fn = (iPerMes as text, iPerAno as number) as text =>
    let
        // Tabela de conversao: nome do mes ? numero
        meses = [
            Janeiro   = 1,  Fevereiro = 2,  Março    = 3,
            Abril     = 4,  Maio      = 5,  Junho    = 6,
            Julho     = 7,  Agosto    = 8,  Setembro = 9,
            Outubro   = 10, Novembro  = 11, Dezembro = 12
        ],
        // Converte o nome para numero
        iPerMesNum = Record.Field(meses, iPerMes),
        // Logica original
        iMesAnt      = iPerMesNum - 1,
        iMesAntFinal = if iMesAnt = 0 then 12 else iMesAnt,
        iAnoAnt      = if iMesAnt = 0 then iPerAno - 1 else iPerAno,
        sMesAntFmt   = Text.PadStart(Text.From(iMesAntFinal), 2, "0"),
        sMesFmt      = Text.PadStart(Text.From(iPerMesNum), 2, "0"),
        sPerAnt      = "26/" & sMesAntFmt & "/" & Text.From(iAnoAnt),
        sPerAtual    = "25/" & sMesFmt    & "/" & Text.From(iPerAno),
        resultado    = sPerAnt & " à " & sPerAtual
    in
        resultado
in
    fn
