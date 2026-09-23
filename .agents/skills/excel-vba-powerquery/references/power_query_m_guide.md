# Guia de Linguagem M (Power Query)

Este documento complementa a skill `excel-vba-powerquery` com padrões essenciais de modelagem, funções utilitárias e dicas de performance em Power Query M.

---

## 1. Estrutura Canônica de uma Consulta M

Toda consulta em M utiliza o bloco `let ... in`:

```powerquery
let
    // 1. Fonte de dados
    Fonte = Excel.Workbook(File.Contents("C:\Caminho\Dados.xlsx"), null, true),
    
    // 2. Navegação na tabela/aba
    Tabela_Sheet = Fonte{[Item="Tabela",Kind="Sheet"]}[Data],
    
    // 3. Promoção de cabeçalhos
    CabecalhosPromovidos = Table.PromoteHeaders(Tabela_Sheet, [PromoteAllScalars=true]),
    
    // 4. Tipagem estrita
    TipoAlterado = Table.TransformColumnTypes(CabecalhosPromovidos, {
        {"ID", Int64.Type},
        {"Data", type date},
        {"Descricao", type text},
        {"Valor", type number}
    }),
    
    // 5. Filtros e transformações
    LinhasFiltradas = Table.SelectRows(TipoAlterado, each ([Valor] <> null and [Valor] > 0))
in
    LinhasFiltradas
```

---

## 2. Leitura Dinâmica de Parâmetros da Planilha

Para evitar caminhos de arquivo fixos (hardcoded) no M, crie uma tabela de parâmetros no Excel chamada `tblParametros` com colunas `Parametro` e `Valor`, e acesse via função M:

```powerquery
let
    fnObterParametro = (nomeParametro as text) =>
        let
            Fonte = Excel.CurrentWorkbook(){[Name="tblParametros"]}[Content],
            Linha = Table.SelectRows(Fonte, each ([Parametro] = nomeParametro)),
            Valor = if Table.IsEmpty(Linha) then null else Linha{0}[Valor]
        in
            Valor
in
    fnObterParametro
```

---

## 3. Criação de Funções Customizadas (`fn...`)

Funções customizadas facilitam cálculos complexos reaproveitados em várias consultas:

```powerquery
(dataRegistro as nullable date, diasUteis as nullable number) as nullable date =>
let
    DataResultado = 
        if dataRegistro = null or diasUteis = null then
            null
        else
            Date.AddDays(dataRegistro, diasUteis)
in
    DataResultado
```

---

## 4. Mesclagem e Combinação de Tabelas (Joins)

Prefira selecionar apenas as colunas necessárias ao expandir tabelas combinadas:

```powerquery
let
    // Realiza o merge (Left Outer Join)
    Merge = Table.NestedJoin(TabelaPrincipal, {"ID_Contrato"}, TabelaContratos, {"ID_Contrato"}, "ContratoExpandido", JoinKind.LeftOuter),
    
    // Expande apenas os campos necessários para não sobrecarregar memória
    Expandido = Table.ExpandTableColumn(Merge, "ContratoExpandido", {"NomeContrato", "UF"}, {"Contrato.Nome", "Contrato.UF"})
in
    Expandido
```

---

## 5. Cuidados com `Table.Buffer` e Performance

* **Quando usar `Table.Buffer`:**
  Use ao carregar uma tabela pequena/média que será consultada dezenas ou centenas de vezes dentro de uma função customizada linha a linha. Sem o buffer, o Power Query pode recarregar a fonte a cada linha avaliada.
* **Quando NÃO usar:**
  Não use em tabelas gigantes, pois consome memória RAM desnecessariamente e impede o *Query Folding* contra fontes de banco de dados.
