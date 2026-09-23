let
    Fonte = Excel.Workbook(File.Contents(localLote2&localBA), null, true),
    ConsolidadoPA_Table = Fonte{[Item="ConsolidadoPA",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(ConsolidadoPA_Table,{{"MEDIDAS CONTRATO", type text}, {"Padrão", type text}, {"Ø", type any}, {"A", type any}, {"L", type any}, {"P", type any}, {"cm³ = V1", type number}, {"cm³ = V2", Int64.Type}, {"FC = V1/V2", type number}, {"QUAm-1", Int64.Type}, {"QBm", type any}, {"QDm", type any}, {"QNm", Int64.Type}, {"QUAm", Int64.Type}, {"UA", type number}, {"$ Unitário", type number}, {"$ Total", type any}, {"Tabela", type text}, {"Contrato", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([MEDIDAS CONTRATO] <> null)),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "Contrato-Reg.", each "PA-LT2-BA"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2", "FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "UA", "$ Unitário", "$ Total", "Tabela", "Contrato"})
in
    #"Colunas Reordenadas"
