let
    // -------------------------------------------------------------------
    // fnCalcularIAPFARQ
    // -------------------------------------------------------------------
    // Equivalente M da funcao VBA calcular_IAPFARQ (modulo
    // fncEditalGuarda) - item no 8.4.2 do contrato de guarda externa.
    //
    // O QUE FAZ:
    //   Calcula o indice IAPFARQ (percentual de solicitacoes atendidas
    //   DENTRO DO PRAZO em relacao ao total de solicitacoes atendidas
    //   NO PERIODO), pela formula:
    //
    //       IAPFARQ = (TSPR / TSPE) * 100
    //
    // CASOS ESPECIAIS:
    //   - Se TSPR = 0 e TSPE = 0 (nenhuma solicitacao atendida no
    //     periodo, nem dentro nem fora do prazo): retorna 100 -
    //     considera-se desempenho pleno, ja que nao houve demanda a
    //     penalizar.
    //   - Se TSPR = 0 e TSPE > 0 (houve atendimentos, mas nenhum
    //     dentro do prazo - pior cenario possivel): retorna 1 em vez
    //     de 0. Isso e fiel ao comportamento do VBA original; nao
    //     altera o resultado final do FDM, pois ambos os valores (0
    //     ou 1) cairiam na mesma faixa "< 90" em fnCalcularFDMMensal.
    //
    // OBSERVACAO SOBRE A VERSAO VBA:
    //   O VBA original tinha duas formas equivalentes de calcular esta
    //   formula, sendo a segunda (usada aqui) preferivel por evitar
    //   overflow de Integer no VBA (irrelevante em M, que nao tem esse
    //   limite, mas mantido por fidelidade ao calculo original):
    //     'dIAPFARQ = (iTSPR * 100) / iTSPE   ' comentada no VBA
    //     dIAPFARQ = (iTSPR / iTSPE) * 100    ' versao usada
    //
    // PARAMETROS:
    //   TSPE - Total de Solicitacoes atendidas no Periodo
    //   TSPR - Total de Solicitacoes atendidas dentro do Prazo
    //
    // RETORNO:
    //   number - o indice IAPFARQ em percentual (0 a 100)
    // -------------------------------------------------------------------
    fnCalcularIAPFARQ = (TSPE as number, TSPR as number) as number =>
        if TSPR = 0 then
            if TSPE = 0 then 100 else 1
        else
            (TSPR / TSPE) * 100
in
    fnCalcularIAPFARQ
