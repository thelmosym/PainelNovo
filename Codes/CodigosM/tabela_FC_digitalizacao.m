let
    Fonte = Excel.CurrentWorkbook(){[Name="tabela_FC_digitalizacao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Descrição da tabela", type text}, {"Item", type text}, {"Unidade", type text}, {"FC", type number}})
in
    #"Tipo Alterado"
