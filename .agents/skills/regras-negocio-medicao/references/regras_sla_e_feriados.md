# Manual Técnico de Regras: SLA, Prazos Úteis e Feriados

Este documento detalha o algoritmo, as regras contratuais e as premissas matemáticas para apuração do prazo limite de atendimento (**SLA - Service Level Agreement**) e conformidade operacional nos contratos da Petrobras.

---

## 1. Janela de Expediente Útil

O cálculo de prazo útil é fundamentado em uma jornada comercial estrita:

- **Expediente da Manhã:** Das **08:00 às 12:00** (4 horas úteis = 240 minutos).
- **Intervalo Intrajornada (Almoço):** Das **12:00 às 13:00** (1 hora de intervalo suspensa; o relógio de SLA **não corre**).
- **Expediente da Tarde:** Das **13:00 às 17:00** (4 horas úteis = 240 minutos).
- **Carga Horária Útil por Dia Útil:** **8 horas** (480 minutos).
- **Período Noturno / Fora de Expediente:** Das **17:00 às 08:00** do próximo dia útil (relógio suspenso).
- **Fins de Semana:** Sábado e Domingo possuem 0 horas úteis.

---

## 2. Ajuste Inicial da Data/Hora de Abertura

Quando uma solicitação ou Ordem de Serviço (OS) é registrada no sistema fora do horário comercial útil, o marco zero de contagem do prazo é transposto automaticamente:

| Horário de Registro Original | Dia do Registro | Marco Inicial da Contagem (Início Efetivo) |
| :--- | :--- | :--- |
| **Antes das 08:00** (ex: 06:30) | Dia útil | **08:00** do mesmo dia útil. |
| **Durante o Almoço** (12:00 às 13:00) | Dia útil | **13:00** do mesmo dia útil. |
| **Após as 17:00** (ex: 18:15) | Dia útil | **08:00** do primeiro dia útil subsequente. |
| **Qualquer horário** | Sábado ou Domingo | **08:00** da próxima segunda-feira útil (ou dia útil após feriado). |
| **Qualquer horário** | Feriado Local/Nacional | **08:00** do primeiro dia útil imediatamente seguinte. |

---

## 3. Prazos Contratuais Típicos (Conversão em Horas Úteis)

Os prazos são definidos em dias úteis a partir da tabela de referência [`tabelaPrazoContrato`](Tabelas%20Especificas/TabelaB-V2.xlsx):

| Prazo em Dias Úteis | Carga Útil Equivalente | Exemplo de Aplicação |
| :---: | :---: | :--- |
| **`0,0` dia** | Imediato / 0 horas | Consulta simples ou liberação em balcão. |
| **`0,5` dia** | **4 horas úteis** | Atendimentos emergenciais / Frete Expresso local. Se aberto às 08:00, vence às 12:00 do mesmo dia. Se aberto às 10:00, vence às 15:00 (suspende das 12h às 13h). |
| **`1,0` dia** | **8 horas úteis** | Atendimentos normais da mesma localidade. Se aberto às 08:00 de segunda, vence às 17:00 de segunda. |
| **`2,0` dias** | **16 horas úteis** | Atendimentos de rotina intermunicipais. |
| **`3,0` dias** | **24 horas úteis** | Organização ou migração simples de acervo. |
| **`5,0` dias** | **40 horas úteis** | Demandas analíticas complexas de arquivos ou transferências massivas. |

---

## 4. Integração do Calendário de Feriados (`Feriados_2026_2030.csv`)

O cálculo consome a base tabular [`Tabelas Especificas/Feriados_2026_2030.csv`](Tabelas%20Especificas/Feriados_2026_2030.csv). Para cada data percorrida no cálculo, verifica-se se ela coincide com um feriado aplicável ao local da prestação do serviço:

1. **Feriados Nacionais:**
   - Aplicáveis a todos os contratos e localidades (ex.: 01/01, 21/04, 01/05, 07/09, 12/10, 02/11, 15/11, 25/12).
2. **Feriados Estaduais:**
   - Filtrados pela sigla do estado (`UF`):
     - `BA` (Dois de Julho - Independência da Bahia);
     - `RJ` (São Jorge - 23/04);
     - `SP` (Revolução Constitucionalista - 09/07);
     - `ES` (Nossa Senhora da Penha).
3. **Feriados Municipais:**
   - Filtrados pela combinação estrita de `Município` e `UF` (ex.: Padroeiros municipais, aniversários de cidades dos polos como Rio de Janeiro, Santos, Macaé, Salvador, São Mateus, etc.).

Se a data for classificada como feriado para aquele município/UF, a data inteira é desconsiderada como dia útil, avançando o relógio para o próximo dia útil às 08:00.

---

## 5. Algoritmo de Avanço do SLA em Power Query M (`fnCalcularDataSLA.m`)

O motor de cálculo simula o fluxo do relógio operacional:

```
[Início: Data/Hora Abertura + Prazo em Horas]
                     │
                     ▼
             É Dia Útil? ──── NÃO ───► Avança para próximo dia às 08:00
                     │
                    SIM
                     │
                     ▼
             É Feriado Local? ── SIM ──► Avança para próximo dia às 08:00
                     │
                    NÃO
                     │
                     ▼
             Horário < 08:00? ── SIM ──► Ajusta para 08:00 do mesmo dia
                     │
                    NÃO
                     │
                     ▼
             12:00 às 13:00? ── SIM ──► Pula para 13:00 (sem abater horas)
                     │
                    NÃO
                     │
                     ▼
             Horário >= 17:00? ── SIM ──► Avança para próximo dia às 08:00
                     │
                    NÃO
                     │
                     ▼
       Consome tempo restante até o fim da janela útil atual
       (ou até esgotar o prazo).
                     │
                     ▼
        [Prazo Restante = 0?] ── NÃO ──► Repete ciclo no próximo dia
                     │
                    SIM
                     ▼
           [Data Limite SLA Apurada]
```

---

## 6. Determinação do Status de Conformidade

Após a apuração da **Data Limite SLA**, o atendimento é comparado com o momento do encerramento oficial:

1. **Atendimento Padrão:**
   - Se $\text{Data Fechamento} \le \text{Data Limite SLA} \implies \mathbf{"Dentro\ do\ prazo"}$
   - Se $\text{Data Fechamento} > \text{Data Limite SLA} \implies \mathbf{"Fora\ do\ prazo"}$
2. **Prazo Combinado (Extraordinário):**
   - Se houver registro explícito e justificado na coluna `Prazo combinado`, a apuração confronta a data de fechamento contra o prazo combinado pactuado com a fiscalização Petrobras.
3. **Atendimentos Cancelados ou em Andamento:**
   - Não recebem status de conformidade de SLA, pois não compõem a medição faturável do ciclo.
