let
    Fonte = Excel.CurrentWorkbook(){[Name="valEnderecoGalpao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Galpãp", type text}, {"Galpãp_Logradouro", type text}, {"Galpãp_CEP", type text}, {"Galpãp_Bairro", type text}, {"Galpãp_Cidade", type text}, {"Galpãp_UF", type text}})
in
    #"Tipo Alterado"
