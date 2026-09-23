let
    Fonte = Excel.CurrentWorkbook(){[Name="tabela_FC_arm"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Descrição da tabela", type text}, {"Item", type text}, {"Diâmetro", Int64.Type}, {"A", type number}, {"L", type number}, {"P", type number}, {"FC", type number}}),
    #"Consulta Acrescentada" = Table.Combine({#"Tipo Alterado", tabela_FC_migracao, tabela_FC_digitalizacao}),
    #"Personalização Adicionada" = Table.AddColumn(#"Consulta Acrescentada", "cod_Tabela_FC", each [Atividade do painel]
&
[Item])
in
    #"Personalização Adicionada"
