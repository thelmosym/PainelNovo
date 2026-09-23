let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    PPU_Table = Fonte{[Item="PPU",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(PPU_Table,{{"Contrato", type text}, {"Item PPU", type text}, {"Linha de Serviço PPU", type text}, {"Descrição da Linha de Serviço", type text}, {"Unidade da Linha de Serviço", type text}, {"P. Unitário (R$)", type number}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([Contrato] = "PA-LT2")),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Linhas Filtradas",{{"P. Unitário (R$)", Currency.Type}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado1",{{"Unidade da Linha de Serviço", "Unidade"}, {"Linha de Serviço PPU", "Linha de Serviço"}})
in
    #"Colunas Renomeadas"
