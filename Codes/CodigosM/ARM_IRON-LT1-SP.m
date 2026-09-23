let
    Fonte = ARM,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([#"Contrato-Reg."] = "IRON-LT1-SP"))
in
    #"Linhas Filtradas"
