# Manual Técnico de Regras: FDM (Fator de Desempenho Mensal) e Qualidade Contratual

Este documento detalha os fundamentos contratuais, a classificação analítica e os algoritmos de apuração do **Fator de Desempenho Mensal (FDM)** e indicadores de conformidade de SLA nos contratos da Petrobras.

---

## 1. Conceito e Objetivo do FDM

O **Fator de Desempenho Mensal (FDM)** é um indicador de nível de serviço contratual (SLA) que atua diretamente como **multiplicador de ajuste financeiro** sobre a fatura de serviços.

- **Meta de Excelência:** A prestadora deve cumprir rigorosamente os prazos de atendimento das solicitações operacionais.
- **Efeito Financeiro:**
  - Se a contratada cumprir as metas contratuais (ex.: conformidade $\ge 98\%$), recebe $100\%$ do valor faturado ($FDM = 1,00$).
  - Se a performance cair abaixo dos níveis acordados, aplica-se um fator redutor sobre a medição mensal ($FDM < 1,00$), retendo valores proporcionais aos atrasos.

---

## 2. Classificação Analítica do Atendimento (`fnClassificarFDM.m`)

Cada solicitação na base consolidada é categorizada para efeitos de apuração de qualidade:

| Classificação FDM | Critério Operacional | Tratamento no Cálculo de SLA |
| :--- | :--- | :--- |
| **`Normal`** | Atendimento regular com prazo padrão de tabela (1 a 5 dias úteis). | Computado no índice $IAPFARQ_1$ e no indicador $TSPE_1 / TSPR_1$. |
| **`Expresso`** | Atendimento urgente / emergencial com prazo reduzido (0,5 dia / 4 horas úteis). | Computado no índice $IAPFARQ_2$ e no indicador $TSPE_2 / TSPR_2$. |
| **`Isento`** | Chamados com atrasos decorrentes de causas externas formalmente justificadas (ex.: greve de transporte, paralisação predial, falta de energia na unidade Petrobras ou autorização expressa do fiscal). | **Desconsiderado** do denominador de penalidades; não impacta negativamente o FDM da contratada. |
| **`FDM definido`** | Atendimentos com regras pré-pactuadas ou prazos combinados extraordinários. | Avaliado estritamente contra a meta acordada entre as partes. |

---

## 3. Indicadores Contratuais Principais

A consulta [`TabelaFDMPorContrato.m`](Codes/CodigosM/TabelaFDMPorContrato.m) orquestra a consolidação dos seguintes indicadores:

### 3.1 Índice de Atendimento aos Prazos ($IAPFARQ$)
Mede o percentual de solicitações entregues dentro da Data Limite de SLA:
$$IAPFARQ = \left( \frac{\text{Total de Atendimentos Dentro do Prazo}}{\text{Total de Atendimentos Elegíveis (Normal + Expresso)}} \right) \times 100$$

### 3.2 Indicadores $TSPE$ e $TSPR$
- **$TSPE$ (Tempo de Solicitação Padrão Esperado):** A soma das horas úteis contratualmente previstas para todos os chamados executados.
- **$TSPR$ (Tempo de Solicitação Padrão Realizado):** A soma das horas úteis efetivamente consumidas pela contratada até o encerramento de cada chamado.
- A relação $\frac{TSPR}{TSPE}$ avalia se a contratada está operando com folga ou no limite do esgotamento de prazo.

---

## 4. Escala Contratual de Aplicação do FDM

Embora cada instrumento jurídico possua faixas específicas, a sistemática padrão adota os seguintes patamares de fechamento:

| Faixa de Cumprimento ($IAPFARQ$) | Fator Multiplicador ($FDM$) | Impacto Financeiro na Medição |
| :---: | :---: | :--- |
| **$\ge 98,00\%$** | **`1,0000`** | Pagamento integral (100% da Memória de Cálculo). |
| **$95,00\%$ a $97,99\%$** | **`0,9800`** | Glosa / Redutor de 2% sobre o montante de serviços. |
| **$90,00\%$ a $94,99\%$** | **`0,9500`** | Glosa / Redutor de 5% sobre o montante de serviços. |
| **$< 90,00\%$** | **`0,9000`** (ou inferior) | Glosa máxima de 10% e abertura de notificação de descumprimento contratual. |

---

## 5. Recálculo e Histórico de Períodos Anteriores

Para evitar distorções sazonais, alguns editais determinam que o FDM definitivo seja calculado pela média móvel ponderada dos últimos ciclos:
- `fnCalcularPeriodosAnteriores.m`: mapeia os ciclos anteriores de faturamento.
- `fnBuscarFDMAnterior.m`: recupera as notas de qualidade homologadas nos meses passados.
- `fnRecalcularFDM.m`: aplica a ponderação trimestral ou semestral para estabilização do índice financeiro.

---

## 6. Unicidade de Apontamento e Prevenção de Falsos Positivos

Para que o FDM reflita a real qualidade operacional sem distorções artificiais:
1. **Atendimentos Cancelados (`Situação <> "CO"`):** São excluídos da base de cálculo.
2. **Ordens de Serviço com Múltiplos Itens:** Um chamado com 3 itens entregues no mesmo momento conta como **1 único evento de atendimento** para conformidade de SLA, impedindo que um atraso pontual gere 3 penalidades duplicadas no cálculo do $IAPFARQ$.
