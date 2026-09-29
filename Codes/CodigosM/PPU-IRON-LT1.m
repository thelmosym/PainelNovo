let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    PPU_Table = Fonte{[Item="PPU",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(PPU_Table,{{"Contrato", type text}, {"Item PPU", type text}, {"Linha de Serviço PPU", type text}, {"Descrição da Linha de Serviço", type text}, {"Unidade da Linha de Serviço", type text}, {"P. Unitário (R$)", type number}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([Contrato] = "IRON-LT1")),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Linhas Filtradas",{{"P. Unitário (R$)", Currency.Type}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado1",{{"Linha de Serviço PPU", "Linha de Serviço"}, {"Unidade da Linha de Serviço", "Unidade"}})
in
    #"Colunas Renomeadas"
