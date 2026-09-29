let
    ParametroTabela = Excel.CurrentWorkbook(){[Name="localTabelaB"]}[Content],
    CaminhoTabelaB = Text.From(ParametroTabela{0}[Column1]),
    ArquivoTabelaB = Binary.Buffer(File.Contents(CaminhoTabelaB)),
    Fonte = Excel.Workbook(ArquivoTabelaB, null, true),
    Contrato_Table = Fonte{[Item="Contrato", Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(
        Contrato_Table,
        {
            {"Contrato", type text},
            {"Nome Fornecedor", type text},
            {"CNPJ", type text},
            {"Contrato SAP", Int64.Type},
            {"Instrumento Contratual Jurídico", Int64.Type},
            {"Data Início", type date},
            {"Data Término", type date},
            {"Região", type text},
            {"Endereço", type text},
            {"Município", type text},
            {"UF", type text},
            {"Filial/Matriz", type text}
        }
    )
in
    #"Tipo Alterado"
