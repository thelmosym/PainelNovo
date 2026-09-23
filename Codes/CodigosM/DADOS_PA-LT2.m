let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "PA-LT2-BA"))
in
    #"Linhas Filtradas"
