let
    // -------------------------------------------------------------------
    // fnCalcularIAPFARQ
    // -----------------------------------------------------------------
    // Equivalente M da função VBA calcular_IAPFARQ (módulo
    // fncEditalGuarda) — item nº 8.4.2 do contrato de guarda externa.
    //
    // O QUE FAZ:
    //   Calcula o índice IAPFARQ (percentual de solicitações atendidas
    //   DENTRO DO PRAZO em relação ao total de solicitações atendidas
    //   NO PERÍODO), pela fórmula:
    //
    //       IAPFARQ = (TSPR / TSPE) * 100
    //
    // CASOS ESPECIAIS:
    //   - Se TSPR = 0 e TSPE = 0 (nenhuma solicitação atendida no
    //     período, nem dentro nem fora do prazo): retorna 100 —
    //     considera-se desempenho pleno, já que não houve demanda a
    //     penalizar.
    //   - Se TSPR = 0 e TSPE > 0 (houve atendimentos, mas nenhum
    //     dentro do prazo — pior cenário possível): retorna 1 em vez
    //     de 0. Isso é fiel ao comportamento do VBA original; não
    //     altera o resultado final do FDM, pois ambos os valores (0
    //     ou 1) cairiam na mesma faixa "< 90" em fnCalcularFDMMensal.
    //
    // OBSERVAÇÃO SOBRE A VERSÃO VBA:
    //   O VBA original tinha duas formas equivalentes de calcular esta
    //   fórmula, sendo a segunda (usada aqui) preferível por evitar
    //   overflow de Integer no VBA (irrelevante em M, que não tem esse
    //   limite, mas mantido por fidelidade ao cálculo original):
    //     'dIAPFARQ = (iTSPR * 100) / iTSPE   ' comentada no VBA
    //     dIAPFARQ = (iTSPR / iTSPE) * 100    ' versão usada
    //
    // PARÂMETROS:
    //   TSPE - Total de Solicitações atendidas no Período
    //   TSPR - Total de Solicitações atendidas dentro do Prazo
    //
    // RETORNO:
    //   number — o índice IAPFARQ em percentual (0 a 100)
    // -------------------------------------------------------------------
    fnCalcularIAPFARQ = (TSPE as number, TSPR as number) as number =>
        if TSPR = 0 then
            if TSPE = 0 then 100 else 1
        else
            (TSPR / TSPE) * 100
in
    fnCalcularIAPFARQ
