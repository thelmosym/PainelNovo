let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Origem_Destino_Table = Fonte{[Item="Origem_Destino",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Origem_Destino_Table,{{"DE_LocalCentro", type text}, {"DE_RegMetrop", type text}, {"DE_Tipo", type text}, {"DE_Logradouro", type text}, {"DE_CEP", type text}, {"DE_Bairro", type text}, {"DE_Cidade", type text}, {"DE_UF", type text}, {"PARA_LocalCentro", type text}, {"PARA_RegMetrop", type text}, {"PARA_Tipo", type text}, {"PARA_Logradouro", type text}, {"PARA_CEP", type text}, {"PARA_Bairro", type text}, {"PARA_Cidade", type text}, {"PARA_UF", type text}, {"KM", type number}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{"DE_LocalCentro", "DE_RegMetrop", "DE_Tipo", "DE_Logradouro", "DE_CEP", "DE_Bairro", "DE_Cidade", "DE_UF", "PARA_RegMetrop", "KM"}),
    #"Duplicatas Removidas" = Table.Distinct(#"Colunas Removidas", {"PARA_LocalCentro"}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Duplicatas Removidas",{"PARA_Tipo"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Colunas Removidas1",{{"PARA_LocalCentro", "Galpãp"}, {"PARA_Logradouro", "Galpãp_Logradouro"}, {"PARA_CEP", "Galpãp_CEP"}, {"PARA_Bairro", "Galpãp_Bairro"}, {"PARA_Cidade", "Galpãp_Cidade"}, {"PARA_UF", "Galpãp_UF"}})
in
    #"Colunas Renomeadas"
