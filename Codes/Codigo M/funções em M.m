let
    Fonte = Excel.Workbook(File.Contents("C:\Users\DOK4\OneDrive - PETROBRAS\Área de Trabalho\Thelmo\Tabelas Especificas\TabelaB-V2.xlsx"), null, true),
    Centro_Table = Fonte{[Item="Centro",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Centro_Table,{{"Centro", type text}, {"Código Centro", Int64.Type}, {"CNPJ", type text}, {"Endereço", type text}, {"CEP", type text}, {"Bairro", type text}, {"Município", type text}, {"UF", type text}}),
    #"Linhas Classificadas" = Table.Sort(#"Tipo Alterado",{{"Centro", Order.Ascending}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Linhas Classificadas",{{"Centro", Text.Upper, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto em Maiúscula",{{"Centro", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"Centro", Text.Clean, type text}})
in
    #"Texto Limpo"


    