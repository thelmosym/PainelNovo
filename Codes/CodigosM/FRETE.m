let
    Fonte = DADOS,
    #"Colunas Removidas" = Table.RemoveColumns(Fonte,{"Aplicação", "Código da#(lf)solicitação", "Código OS", "D. abertura", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "FC", "QExec", "Código Centro"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Removidas", each ([Linha de serviço PPU] = "FRE-EXP" or [Linha de serviço PPU] = "FRE-NRM"))
in
    #"Linhas Filtradas"
