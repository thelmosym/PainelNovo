let
    Fonte = Folder.Files(localMonitIndividualT2M),
    ArquivosValidos = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$") and ([Extension] = ".xlsm" or [Extension] = ".xlsx")),
    Navegação1 = ArquivosValidos{0}[Content]
in
    Navegação1
