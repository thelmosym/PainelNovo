# Manual Técnico de Regras: Quantidade Executada (QExec) e Fatores de Conversão

Este documento detalha as fórmulas matemáticas, tabelas de conversão e regras de negócio para apuração da **Quantidade Executada Padronizada (QExec)** nos contratos da Petrobras.

---

## 1. Fundamento Contratual do QExec

Na sistemática de faturamento da Petrobras, os contratos de guarda documental e apoio administrativo possuem uma **Planilha de Preços Unitários (PPU)** baseada em itens de serviço padrão (geralmente referenciados na manipulação de **caixas-padrão tipo box** ou **unidades consolidadas de custódia**).

Como a operação física movimenta materiais de naturezas heterogêneas (pastas finas, documentos avulsos, caixas pesadas, mídias digitais, livros e plantas de engenharia), a quantidade física atendida deve ser convertida para a unidade de cobrança contratual por meio de fatores de equivalência técnica.

---

## 2. Matriz de Fatores de Conversão Padronizados

A função [`fnCalcularQExec.m`](Codes/CodigosM/fnCalcularQExec.m) implementa as seguintes regras de conversão:

| Atividade / Categoria | Item Físico | Fator ($FC$) | Fórmula de Cálculo | Justificativa Operacional |
| :--- | :--- | :---: | :---: | :--- |
| **Embalagem** | Caixas Box de arquivo | **`0,10`** | $QExec = Qtd \times 0,10$ | O manuseio de 10 caixas físicas de arquivo equivale a 1 unidade faturável no PPU. |
| **Item Avulso** | Pastas, Dossiês, Prontuários | **`0,017`** | $QExec = Qtd \times 0,017$ | Equivalência volumétrica: aproximadamente 58 a 60 itens avulsos compõem o volume de uma caixa box padrão. |
| **Organização Analítica** | Caixa / Lote indexado | **`1,00`** | $QExec = Qtd \times 1,00$ | Serviço de alta complexidade intelectual (higienização, classificação temporal, inventário folha a folha). Cobrado na razão integral 1:1. |
| **Organização Simples** | Caixa / Volume | **`0,50`** | $QExec = Qtd \times 0,50$ | Reordenação básica ou conferência física de lotes. Cada 2 volumes equivalem a 1 unidade faturável. |
| **Passagem Direta / Consulta** | Atendimento em balcão | **`1,00`** | $QExec = Qtd \times 1,00$ | Quantidade atendida sem redutor, correspondendo diretamente às unidades demandadas. |
| **Digitalização / Migração** | Páginas / Imagens | **Tabela `FC`** | $QExec = Qtd \times FC$ | O fator varia de acordo com o tamanho do documento (A4, A3, grandes formatos de engenharia) e resolução. |
| **Material p/ Arquivamento** | Entrada inicial de acervo | **`null`** | $QExec = \text{null}$ | Movimentação de recepção e custódia inicial sem remuneração avulsa de execução (remunerada na armazenagem). |
| **Baixa Permanente** | Descarte / Expugo | **`null`** | $QExec = \text{null}$ | Eliminação formal de acervo conforme tabela de temporalidade documental. |

---

## 3. Busca Dinâmica de Fator na Tabela `FC`

Para serviços especializados de digitalização, microfilmagem, higienização profunda e conversão eletrônica:

1. **Geração da Chave Relacional:**
   $$\text{cod\_Tabela\_FC} = \text{Descrição da atividade} \ \& \ \text{Item}$$
2. **Junção Relacional (Left Outer Join):**
   A base é cruzada com a tabela de referência [`tabela_FC`](Codes/CodigosM/tabela_FC.m) pela chave `cod_Tabela_FC`.
3. **Tratamento de Nulos:**
   Se a consulta não localizar correspondência para a combinação na tabela `FC`, o fator padrão atribuído é `0`, gerando $QExec = 0$ para alertar a necessidade de cadastro na matriz contratual.

---

## 4. Regras Especiais e Casos de Borda

### 4.1 Devolução de Empréstimo
- Quando um documento ou caixa previamente consultada retorna ao galpão para recolocação em prateleira (*refile*):
  - Se for devolução de caixa completa: segue a regra da atividade de recolocação.
  - Se for devolução parcial ou item avulso: pode ser isento de cobrança adicional quando o contrato cobrir a guarda em ciclo fechado.

### 4.2 Arredondamento e Precisão Decimal
- Os valores de $QExec$ devem manter precisão matemática de até **4 casas decimais** nas etapas intermediárias do Power Query, sendo formatados com **2 casas decimais** apenas na apresentação das tabelas finais da Memória de Cálculo (`MC`).
- Não se aplica truncamento manual que prejudique centavos de faturamento.

### 4.3 Quantidades Nulas ou Negativas
- Se a `Qtd. Atendida` for `0` ou `null`, o $QExec$ resultante é obrigatoriamente `0` ou `null`.
- Qualquer valor negativo na base operacional é classificado como erro impeditivo no módulo de auditoria (`modAuditoriaLog.bas`).
