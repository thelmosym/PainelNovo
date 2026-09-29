let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    FDM_IAPFARQ_Table = Fonte{[Item="FDM_IAPFARQ",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(FDM_IAPFARQ_Table,{{"Medição", type text}, {"Contrato", type text}, {"Item contrato", Int64.Type}, {"IAPFARQ (%)", type number}, {"FDM", type number}})
in
    #"Tipo Alterado"
