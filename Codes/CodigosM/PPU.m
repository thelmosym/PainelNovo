let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    PPU_Table = Fonte{[Item="PPU",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(PPU_Table,{{"Contrato", type text}, {"Item PPU", type text}, {"Linha de Serviço PPU", type text}, {"Descrição da Linha de Serviço", type text}, {"Unidade da Linha de Serviço", type text}, {"P. Unitário (R$)", type number}})
in
    #"Tipo Alterado"
