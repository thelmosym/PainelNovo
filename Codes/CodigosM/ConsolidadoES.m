let
    Fonte = Excel.Workbook(File.Contents(localLote1&localES), null, true),
    ConsolidadoES_Table = Fonte{[Item="ConsolidadoES",Kind="Table"]}[Data],
    #"Linhas Filtradas" = Table.SelectRows(ConsolidadoES_Table, each ([MEDIDAS CONTRATO] <> null)),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "Contrato-Reg.", each "IRON-LT1-ES"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2", "FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "UA", "$ Unitário", "$ Total", "Tabela", "Contrato"})
in
    #"Colunas Reordenadas"
