let
    // -------------------------------------------------------------------
    // fnCalcularQExecAgrupado
    // -----------------------------------------------------------------
    // Cálculo ADICIONAL/PARALELO ao QExec por linha (fnCalcularQExec).
    // Cria um índice de agrupamento a partir de (Data Abertura + Local
    // PETROBRAS + Local da empresa de guarda), soma o QExec de todas
    // as linhas com o mesmo índice, e atribui o resultado da fórmula
    // de frete/agrupamento (soma/10, piso 1) APENAS na linha de
    // primeira ocorrência do índice — mesmo padrão de deduplicação
    // usado em concatenar_dados_FDM (VBA) e CalcularQuilometragemAdicional
    // (Sistema 2, via Collection).
    //
    // O QUE FAZ:
    //   resultado = 1,          se (soma do grupo / 10) <= 1
    //   resultado = soma / 10,  caso contrário
    //   As demais linhas do mesmo grupo (não a 1ª ocorrência) recebem
    //   null nesta coluna.
    //
    // ? CORREÇÃO APLICADA (durante os testes em produção):
    //   Se TODAS as linhas de um grupo tiverem QExec = null (ex.:
    //   grupo formado inteiramente por "MATERIAL PARA ARQUIVAMENTO"/
    //   "BAIXA PERMANENTE"), List.Sum de uma lista vazia retornaria
    //   null, causando o erro "Não conseguimos converter o valor null
    //   em tipo Logical" ao comparar null <= 1. Corrigido verificando
    //   explicitamente List.IsEmpty antes de somar, e propagando null
    //   como resultado final nesse caso, em vez de tentar a divisão/
    //   comparação.
    //
    // PARÂMETROS:
    //   tabela            - tabela de entrada (já com QExec calculado
    //                       por fnCalcularQExec)
    //   colDataFechamento   - nome da coluna de Data de Fechamento
    //   colLocalPetrobras - nome da coluna de Localidade PETROBRAS
    //   colLocalGuarda    - nome da coluna de Localidade da empresa de
    //                       guarda (galpão)
    //   colQExec          - nome da coluna já contendo o QExec por
    //                       linha
    //
    // RETORNO:
    //   table — a mesma tabela de entrada, com a coluna adicional
    //   "QExecAgrupado" preenchida apenas na 1ª ocorrência de cada
    //   grupo
    // -------------------------------------------------------------------
    fnCalcularQExecAgrupado = (
        tabela as table,
        colDataFechamento as text,
        colLocalPetrobras as text,
        colLocalGuarda as text,
        colQExec as text
    ) as table =>
        let
            // 1) Monta a chave de agrupamento (Data + Local PETROBRAS +
            //    Local Guarda), equivalente ao "sDado1" concatenado
            //    visto nos códigos VBA analisados
            ComChave = Table.AddColumn(tabela, "_ChaveAgrupamento", each
                Text.From(Record.Field(_, colDataFechamento), "pt-BR")
                & "|" & Text.From(Record.Field(_, colLocalPetrobras))
                & "|" & Text.From(Record.Field(_, colLocalGuarda))
            ),

            // 2) Índice sequencial estável — necessário para localizar
            //    a primeira ocorrência de cada chave
            ComIndice = Table.AddIndexColumn(ComChave, "_Indice", 0, 1, Int64.Type),

            // 3) Agrupa por chave: acha o índice mínimo (1ª ocorrência)
            //    e soma o QExec de todas as linhas do grupo
            Agrupado = Table.Group(ComIndice, {"_ChaveAgrupamento"}, {
                {"_IndiceMinimo", each List.Min([_Indice]), Int64.Type},
                {"_SomaQExec", each
                    let
                        valores = List.RemoveNulls(Table.Column(_, colQExec)),
                        // ? Se não houver nenhum valor válido no
                        // grupo, retorna null em vez de deixar
                        // List.Sum({}) produzir null "por acidente"
                        soma = if List.IsEmpty(valores) then null else List.Sum(valores)
                    in
                        soma,
                    type nullable number
                }
            }),

            // 4) Traz de volta o índice mínimo e a soma do grupo para
            //    cada linha original
            Mesclado = Table.NestedJoin(ComIndice, {"_ChaveAgrupamento"},
                Agrupado, {"_ChaveAgrupamento"}, "_Grupo", JoinKind.LeftOuter),
            Expandido = Table.ExpandTableColumn(Mesclado, "_Grupo",
                {"_IndiceMinimo", "_SomaQExec"}, {"_IndiceMinimo", "_SomaQExec"}),

            // 5) Aplica o cálculo APENAS na linha de 1ª ocorrência
            //    (Indice = IndiceMinimo); demais linhas = null.
            //    Também protege contra soma nula (grupo sem QExec
            //    válido nenhum).
            ComResultado = Table.AddColumn(Expandido, "QExecAgrupado", each
                if [_Indice] = [_IndiceMinimo] then
                    if [_SomaQExec] = 0 or [_SomaQExec] = null then
                        0
                    else
                        let divisao = [_SomaQExec] / 10 in
                            if divisao <= 1 then 1 else divisao
                else
                    null
            ),

            // 6) Remove colunas auxiliares, mantendo só o resultado final
            Limpeza = Table.RemoveColumns(ComResultado,
                {"_ChaveAgrupamento", "_Indice", "_IndiceMinimo", "_SomaQExec"}
            )
        in
            Limpeza
in
    fnCalcularQExecAgrupado
