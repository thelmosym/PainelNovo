let
    Fonte = Excel.CurrentWorkbook(){[Name="tabelaPrazoContrato"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Contrato", type text}, {"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Grupo de município", type text}, {"Classificação", type text}, {"Prazo", type number}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "codPrazoContrato", each [Contrato]
&"\"&
[Atividade do painel]
&"\"&
[Grupo de município]),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"codPrazoContrato", "Contrato", "Atividade do painel", "Tab. Contrato", "Grupo de município", "Classificação", "Prazo"})
in
    #"Colunas Reordenadas"
