let
    Fonte = ARM,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([#"Contrato-Reg."] = "IRON-LT1-ES"))
in
    #"Linhas Filtradas"
