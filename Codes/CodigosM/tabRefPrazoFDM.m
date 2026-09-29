let
    Fonte = Excel.CurrentWorkbook(){[Name="tabRefPrazoFDM"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Serviço", type text}, {"Nº contrato", Int64.Type}, {"Descrição da atividade", type text}, {"Atividade do painel", type text}, {"Ref. FDM", Int64.Type}, {"Ref. QExec", type text}, {"Ref. prazo", type text}, {"Obs", type text}})
in
    #"Tipo Alterado"
