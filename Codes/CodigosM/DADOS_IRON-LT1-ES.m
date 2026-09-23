let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-ES"))
in
    #"Linhas Filtradas"
