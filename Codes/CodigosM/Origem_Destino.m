let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Origem_Destino_Table = Fonte{[Item="Origem_Destino",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Origem_Destino_Table,{{"DE_LocalCentro", type text}, {"DE_RegMetrop", type text}, {"DE_Tipo", type text}, {"DE_Logradouro", type text}, {"DE_CEP", type text}, {"DE_Bairro", type text}, {"DE_Cidade", type text}, {"DE_UF", type text}, {"PARA_LocalCentro", type text}, {"PARA_RegMetrop", type text}, {"PARA_Tipo", type text}, {"PARA_Logradouro", type text}, {"PARA_CEP", type text}, {"PARA_Bairro", type text}, {"PARA_Cidade", type text}, {"PARA_UF", type text}, {"KM", type number}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "codOrigenDestino", each [DE_LocalCentro]&
[DE_Cidade]&
[PARA_LocalCentro]&
[PARA_Cidade])
in
    #"Personalização Adicionada"
