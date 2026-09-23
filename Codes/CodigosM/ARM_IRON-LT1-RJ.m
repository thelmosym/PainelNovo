let
    Fonte = ARM,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([#"Contrato-Reg."] = "IRON-LT1-RJ")),
    #"Colunas Removidas" = Table.RemoveColumns(#"Linhas Filtradas",{"Contrato-Reg."})
in
    #"Colunas Removidas"
