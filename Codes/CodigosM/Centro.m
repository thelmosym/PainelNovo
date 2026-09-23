let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    
    Centro_Table = Fonte{[Item="Centro", Kind="Table"]}[Data],

    // Tipos e limpeza da coluna Centro combinados em um único passo
    #"Tipos e Transformações" = Table.TransformColumnTypes(
        Centro_Table,
        {
            {"Código Centro", Int64.Type},
            {"CNPJ",          type text},
            {"Endereço",      type text},
            {"CEP",           type text},
            {"Bairro",        type text},
            {"Município",     type text},
            {"UF",            type text}
        }
    ),

    // Upper + Trim + Clean aplicados juntos na coluna Centro
    #"Centro Padronizado" = Table.TransformColumns(
        #"Tipos e Transformações",
        {{"Centro", each Text.Clean(Text.Trim(Text.Upper(_))), type text}}
    ),

    #"Linhas Classificadas" = Table.Sort(
        #"Centro Padronizado",
        {{"Centro", Order.Ascending}}
    )

in
    #"Linhas Classificadas"
