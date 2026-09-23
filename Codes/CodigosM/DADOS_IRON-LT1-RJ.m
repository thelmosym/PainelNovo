let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-RJ"))
in
    #"Linhas Filtradas"
