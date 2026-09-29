let
    Combinado = Table.Combine({ConsolidadoES, ConsolidadoRJ, ConsolidadoSP, ConsolidadoPA}),
    // Bufferiza o resultado final - evita reabrir os 4 arquivos fisicos
    // a cada referencia subsequente desta consulta em outras etapas.
    ResultadoFinal = Table.Buffer(Combinado),
    #"Colunas Renomeadas" = Table.RenameColumns(ResultadoFinal,{{"Contrato", "Linha de serviço PPU"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Removidas",{{"FC = V1/V2", type number}}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Tipo Alterado",{"$ Unitário", "$ Total", "Tabela", "UA"}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Colunas Removidas1",{{"FC = V1/V2", type number}, {"QUAm-1", type number}, {"QBm", type number}, {"QDm", type number}, {"QNm", type number}, {"QUAm", type number}}),
    #"Valor Substituído" = Table.ReplaceValue(#"Tipo Alterado1",null,0,Replacer.ReplaceValue,{"FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Valor Substituído",{{"FC = V1/V2", "FC"}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Colunas Renomeadas1",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "Linha de serviço PPU"}),
    // QUAm-1 - (QBm + QDm) + QNm
    //
    QUA = Table.AddColumn(#"Colunas Reordenadas", "QUA", each [#"QUAm-1"]-([QBm]+[QDm]) + [QNm]),
    #"Colunas Reordenadas1" = Table.ReorderColumns(QUA,{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "QUA", "Linha de serviço PPU"}),
    #"Colunas Renomeadas2" = Table.RenameColumns(#"Colunas Reordenadas1",{{"Padrão", "Item"}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Renomeadas2", each ([#"Contrato-Reg."] = "PA-LT2-BA")),
    #"Colunas Removidas2" = Table.RemoveColumns(#"Linhas Filtradas",{"Contrato-Reg."})
in
    #"Colunas Removidas2"
