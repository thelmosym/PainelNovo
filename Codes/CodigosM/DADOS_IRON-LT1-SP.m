let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-SP"))
in
    #"Linhas Filtradas"
