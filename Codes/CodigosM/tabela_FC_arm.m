let
    Fonte = Excel.CurrentWorkbook(){[Name="tabela_FC_arm"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Descrição da tabela", type text}, {"Item", type text}, {"Diâmetro", Int64.Type}, {"A", type number}, {"L", type number}, {"P", type number}, {"FC", type number}})
in
    #"Tipo Alterado"
