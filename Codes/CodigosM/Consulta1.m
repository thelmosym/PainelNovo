let
    Fonte = tabelaPrazoContrato,
    Amostra = Table.FirstN(Fonte, 10)
in
    Amostra
