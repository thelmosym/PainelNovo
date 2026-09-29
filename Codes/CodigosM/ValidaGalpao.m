let
    Fonte = Excel.CurrentWorkbook(){[Name="ValidaGalpao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Validação Descrição da atividade para GALPAO", type text}})
in
    #"Tipo Alterado"
