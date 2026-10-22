let
    // -------------------------------------------------------------------
    // fnValidarPeriodoAnterior
    // -------------------------------------------------------------------
    // Equivalente M da função VBA validar_periodo_anterior (módulo fncTabelaB).
    // -------------------------------------------------------------------
    fnValidarPeriodoAnterior = (periodo as nullable text) as logical =>
        if periodo = null or periodo = "" then
            false
        else
            List.Contains(RefFDMHistorico[Medição], periodo)
in
    fnValidarPeriodoAnterior
