let
    Fonte = Excel.CurrentWorkbook(){[Name="tabelaAtividadeItemPPU"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Item PPU", type text}, {"Atividade do painel", type text}, {"Item", type text}, {"Unidade", type text}, {"Linha de serviço PPU", type text}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "CodLinhaPPU", each [Atividade do painel]&[Item])
in
    #"Personalização Adicionada"
