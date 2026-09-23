let
    Fonte = Excel.Workbook(File.Contents(localLote1&localRJ), null, true),
    ConsolidadoRJ_Table = Fonte{[Item="ConsolidadoRJ",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(ConsolidadoRJ_Table,{{"MEDIDAS CONTRATO", type text}, {"Padrão", type text}, {"Ø", Int64.Type}, {"A", type number}, {"L", type number}, {"P", type number}, {"cm³ = V1", type number}, {"cm³ = V2", type number}, {"FC = V1/V2", type number}, {"QUAm-1", Int64.Type}, {"QBm", Int64.Type}, {"QDm", Int64.Type}, {"QNm", Int64.Type}, {"QUAm", Int64.Type}, {"UA", type number}, {"$ Unitário", type number}, {"$ Total", type number}, {"Tabela", type text}, {"Contrato", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([MEDIDAS CONTRATO] <> null)),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "Contrato-Reg.", each "IRON-LT1-RJ"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2", "FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "UA", "$ Unitário", "$ Total", "Tabela", "Contrato"})
in
    #"Colunas Reordenadas"
