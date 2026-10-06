---
name: regras-negocio-medicao
description: >-
  Especialista nas regras de negócio de medição contratual, faturamento de serviços e indicadores operacionais da Petrobras (Contratos Lote 1 IRON e Lote 2 PA).
  Use sempre que for necessário implementar, validar, auditar ou esclarecer regras de negócio do painel, incluindo: ciclo de faturamento (dia 26 ao 25), cálculo de SLA e prazos úteis com feriados, classificação e fórmulas de FDM (TSPE/TSPR/IAPFARQ), conversão de Quantidade Executada (QExec), regras de agrupamento de frete (normal vs. expresso), cálculo de KM adicional, custódia de armazenagem (ARM), auditoria de dados e estrutura dos cadernos oficiais de saída (.xlsx).
---

# Skill: Regras de Negócio de Medição e Faturamento Contratual Petrobras

Esta skill consolida e governa todas as regras de negócio, premissas operacionais, fórmulas matemáticas e critérios contratuais aplicados no **Painel de Controle de Medição da Petrobras** (Contratos **PA-LT2** e **IRON-LT1**).

---

## 📌 Guia de Referência Rápida

| Domínio de Regra | Arquivo M / VBA Responsável | Resumo da Regra | Documento Detalhado |
| :--- | :--- | :--- | :--- |
| **Janela de Faturamento** | `fnGerarPeriodo.m` | Ciclo fechado do dia **26** (mês anterior) ao dia **25** (mês de referência). Apenas `Situação = "CO"`. | [Ver Seção 1](#1-ciclo-e-janela-de-faturamento) |
| **Contratos & Lotes** | `Painel T2M.m`, `DADOS.m` | Segregação estrita entre `PA-LT2` (BA - SAP 4600687006) e `IRON-LT1` (RJ, SP, ES - SAP 4600686987). | [Ver Seção 2](#2-contratos-lotes-e-polos-operacionais) |
| **SLA & Prazos Úteis** | `fnCalcularDataSLA.m` | Expediente útil (08:00–17:00), intervalo de almoço (12:00–13:00) e cruzamento de feriados locais (`Feriados_2026_2030.csv`). | [regras_sla_e_feriados.md](./references/regras_sla_e_feriados.md) |
| **QExec Padronizado** | `fnCalcularQExec.m` | Conversão PPU: Embalagem (0,1), Item Avulso (0,017), Org. Analítica (1,0), Org. Simples (0,5), Fatores FC. | [regras_qexec_e_fatores.md](./references/regras_qexec_e_fatores.md) |
| **Agrupamento de Frete** | `fnCalcularQExecAgrupado.m` | Chave composta `Data | Localidade | Galpão | Linha de serviço PPU`. Segregação estrita `FRE-EXP` de `FRE-NRM`. Fórmula $\text{soma}/10$ (piso 1). | [regras_frete_e_km.md](./references/regras_frete_e_km.md) |
| **KM Adicional** | `Painel T2M.m` | Deslocamentos excedentes à franquia contratual (`valKmAdicional`). Deduplicado por viagem. | [regras_frete_e_km.md](./references/regras_frete_e_km.md) |
| **FDM & Indicadores** | `fnClassificarFDM.m`, `TabelaFDMPorContrato.m` | Conformidade de prazos, $TSPE/TSPR$, índice $IAPFARQ$ e redutor/bonificador de fatura. | [regras_fdm_e_qualidade.md](./references/regras_fdm_e_qualidade.md) |
| **Armazenamento (ARM)** | `B_Calc_ARM.bas`, `SaidaARM.m` | Custódia física de caixas box (20kg), mídias e tubos em galpão no último dia do ciclo. | [regras_contratos_e_exportacao.md](./references/regras_contratos_e_exportacao.md) |
| **Auditoria da Base** | `modAuditoriaLog.bas` | Fiscalização em memória RAM. Permite múltiplos itens por OS legítima. Apontamentos na `LOG_CRITICAS`. | [Ver Seção 8](#8-auditoria-de-qualidade-e-integridade-de-dados) |
| **Cadernos de Saída** | `modGerarArquivos.bas` | Geração de cadernos limpos `.xlsx` em `MEDIÇÃO/` com as 4 abas oficiais: `MC`, `ARM`, `DADOS`, `FRETE`. | [regras_contratos_e_exportacao.md](./references/regras_contratos_e_exportacao.md) |

---

## 1. Ciclo e Janela de Faturamento

1. **Definição de Período:**
   - A medição **NÃO segue o mês civil** (1º ao último dia do mês).
   - Início: **Dia 26** do mês $M-1$.
   - Término: **Dia 25** do mês de referência $M$.
   - *Exemplo (Mês 08/2026):* `26/07/2026 à 25/08/2026`.
2. **Critério de Elegibilidade Estrita:**
   - **Situação:** Apenas atendimentos com `Situação = "CO"` (Concluído).
   - Registros com status `CA` (Cancelado), `EA` (Em Atendimento), `AB` (Aberto) ou vazios são descartados da medição corrente.
   - A data do atendimento (`Período faturamento` / `Data Fechamento`) deve estar contida dentro do intervalo de corte.

---

## 2. Contratos, Lotes e Polos Operacionais

A solução atende a dois contratos de guarda documental e apoio operacional da Petrobras:

```
                          ┌────────────────────────┐
                          │ MEDIÇÃO PETROBRAS      │
                          └───────────┬────────────┘
                                      │
            ┌─────────────────────────┴─────────────────────────┐
            ▼                                                   ▼
┌───────────────────────┐                           ┌───────────────────────┐
│ LOTE 2 - PA           │                           │ LOTE 1 - IRON MOUNTAIN│
│ Contrato: 4600687006  │                           │ Contrato: 4600686987  │
│ Jurídico: 5900.0133214│                           │ Jurídico: 5900.0133210│
│ Polo: BA (Bahia)      │                           │ Polos: RJ, SP, ES     │
└───────────────────────┘                           └───────────┬───────────┘
                                                                │
                                        ┌───────────────────────┼───────────────────────┐
                                        ▼                       ▼                       ▼
                            ┌───────────────────────┐┌───────────────────────┐┌───────────────────────┐
                            │ IRON-LT1-RJ           ││ IRON-LT1-SP           ││ IRON-LT1-ES           │
                            │ Polo Rio de Janeiro   ││ Polo São Paulo        ││ Polo Espírito Santo   │
                            └───────────────────────┘└───────────────────────┘└───────────────────────┘
```

- Cada polo regional gera um caderno individual em `MEDIÇÃO/`.
- No contrato `PA-LT2`, o sufixo estadual é padronizado como `PA-LT2-BA`.

---

## 3. SLA e Prazos Úteis (`fnCalcularDataSLA.m`)

Para detalhes exaustivos, consulte o guia [regras_sla_e_feriados.md](./references/regras_sla_e_feriados.md).

### Parâmetros de Jornada Útil:
- **Expediente Comercial:** **08:00 às 17:00** (8 horas úteis por dia útil).
- **Intervalo Intrajornada (Almoço):** **12:00 às 13:00** (não conta como tempo útil).
- **Fins de Semana:** Sábados e Domingos são ignorados na contagem.
- **Feriados Multilocais:** Integrados via `Feriados_2026_2030.csv`, considerando calendário Nacional, Estadual e Municipal com base no `Município` e `UF` da prestação do serviço.

### Regras de Entrada de Chamado:
- Abertura antes das 08:00 $\to$ tempo começa a contar às 08:00 do mesmo dia.
- Abertura após as 17:00 $\to$ tempo começa a contar às 08:00 do próximo dia útil.
- Abertura entre 12:00 e 13:00 $\to$ tempo começa a contar às 13:00.
- Abertura em sábado, domingo ou feriado $\to$ tempo começa a contar às 08:00 do primeiro dia útil seguinte.

### Avaliação de Conformidade:
- Se $\text{Data/Hora Fechamento} \le \text{Data/Hora Limite SLA} \implies \text{"Dentro do prazo"}$
- Se $\text{Data/Hora Fechamento} > \text{Data/Hora Limite SLA} \implies \text{"Fora do prazo"}$
- **Prazo Combinado:** Prevalece sobre o SLA quando preenchido e justificado.

---

## 4. Quantidade Executada Padronizada - QExec (`fnCalcularQExec.m`)

Para detalhes exaustivos, consulte o guia [regras_qexec_e_fatores.md](./references/regras_qexec_e_fatores.md).

O faturamento Petrobras precifica atividades a partir de um valor unitário padrão de caixa (PPU). Volumes fracionários ou especiais são convertidos matematicamente:

| Atividade / Item | Fator de Conversão | Regra / Justificativa Contratual |
| :--- | :---: | :--- |
| **Embalagem (Caixas)** | `0,10` | 10 caixas físicas equivalem a 1 unidade PPU faturável. |
| **Item Avulso (Dossiês/Pastas)** | `0,017` | ~60 documentos avulsos equivalem a 1 unidade PPU. |
| **Organização Analítica** | `1,00` | Conversão integral por caixa/unidade organizada com índice. |
| **Organização Simples** | `0,50` | 2 unidades equivalem a 1 unidade PPU faturável. |
| **Passagem Direta / Consulta** | `1,00` | Quantidade atendida faturada 1:1 sem fator redutor. |
| **Digitalização / Migração** | Tabela `FC` | Fator obtido dinamicamente via `cod_Tabela_FC` (`tabela_FC.m`). |
| **Material p/ Arquivamento** | `null` | Apenas movimentação física inicial de acervo; sem cobrança de QExec. |
| **Baixa Permanente** | `null` | Expugo de acervo; faturamento via destinação/descarte específico. |

---

## 5. Regras de Frete e Deslocamentos

Para detalhes exaustivos, consulte o guia [regras_frete_e_km.md](./references/regras_frete_e_km.md).

### 5.1 Agrupamento de Frete (`fnCalcularQExecAgrupado.m`)
- **Problema de Negócio:** Se 5 solicitações forem entregues no mesmo prédio Petrobras na mesma viagem, a Petrobras paga pelo deslocamento do veículo (viagem agrupada), não por cada ordem individual.
- **Chave de Agrupamento Obrigatória:**
  $$\text{Chave} = \text{Data Fechamento} \mid \text{Localidade Petrobras} \mid \text{Galpão\_Cidade} \mid \text{Linha de serviço PPU}$$
- **Separação Rigorosa de Serviços:**
  A coluna `Linha de serviço PPU` **DEVE** fazer parte da chave. Frete Normal (`FRE-NRM`) e Frete Expresso (`FRE-EXP`) possuem SLAs, veículos dedicados e tabelas de preço distintas. **Nunca devem ser agrupados na mesma viagem**.
- **Fórmula de Frete:**
  $$\text{Soma} = \sum QExec \text{ do grupo}$$
  $$\text{Se } (\text{Soma} / 10 \le 1) \implies QExecAgrupado = 1$$
  $$\text{Se } (\text{Soma} / 10 > 1) \implies QExecAgrupado = \text{Soma} / 10$$
- **Deduplicação:** O resultado é gravado **apenas na primeira ocorrência** (menor índice). As demais ordens do grupo recebem `null`.

### 5.2 Deduplicação de Quilometragem Adicional (`Painel T2M.m`)
- Calculado a partir da distância entre o Galpão da Contratada e o Centro/Localidade da Petrobras (`Origem_Destino.KM`).
- **Franquia Contratual (`valKmAdicional`):** Raio de tolerância incluído no frete básico.
- Se $\text{KM} \le \text{Franquia} \implies \text{KM Adicional} = 0$.
- Se $\text{KM} > \text{Franquia} \implies \text{KM Adicional} = \text{KM} - \text{Franquia}$.
- Cobrado somente na primeira ocorrência da viagem, evitando cobrança repetida do trajeto para ordens conjuntas.

---

## 6. FDM (Fator de Desempenho Mensal) e Indicadores

Para detalhes exaustivos, consulte o guia [regras_fdm_e_qualidade.md](./references/regras_fdm_e_qualidade.md).

- **Classificação Analítica (`fnClassificarFDM.m`):**
  - `Normal`: atendimento regular com prazo padrão.
  - `Expresso`: chamado emergencial sujeito a índice específico.
  - `Isento`: solicitações sem ônus de tempo (justificativas validadas pelo fiscal).
- **Indicadores Mensais:**
  - $TSPE_1 / TSPR_1$: Padrão vs. Realizado para Atendimentos Normais.
  - $TSPE_2 / TSPR_2$: Padrão vs. Realizado para Atendimentos Expressos.
  - $IAPFARQ$: Índice de Cumprimento de Prazos Global.
- **Aplicação Financeira:** O FDM mensal atua como multiplicador da fatura (redutor quando abaixo da meta mínima contratual de 95% ou 98%).

---

## 7. Armazenamento e Custódia Física (`ARM`)

- Apuração do saldo físico sob custódia nos galpões no dia **25** do ciclo.
- Itens controlados:
  - Caixa Box padrão (20kg);
  - Fitas magnéticas / Mídias ópticas em cofre climatizado;
  - Tubos especiais de perfil de poço e desenhos de engenharia.
- O cálculo mensal é gerado na aba `ARM` e consolidado como item de linha fixa na Memória de Cálculo (`MC`).

---

## 8. Auditoria de Qualidade e Integridade de Dados

- Fiscalização automatizada em memória RAM via [`modAuditoriaLog.bas`](Codes/VBA/Fun%C3%A7%C3%B5es%20Novas/modAuditoriaLog.bas).
- **Críticas Impeditivas (CRÍTICO):**
  - Data de Fechamento anterior à Data de Abertura;
  - Quantidade Atendida nula ou negativa;
  - Código OS não preenchido ou sem correspondência;
  - Inconsistência de código de centro ou cidade de galpão.
- **Múltiplos Itens por OS (Regra Operacional):**
  - **Permitido:** É perfeitamente legítimo existirem múltiplas linhas com a mesma chave `Atividade + Solicitação + OS` quando há entregas fracionadas, múltiplos volumes ou etapas complementares. A auditoria **NÃO** deve bloquear chamados com múltiplos itens legítimos.
- Diagnóstico exibido na aba `LOG_CRITICAS` com contadores executivos e links diretos para correção.

---

## 9. Cadernos Contratuais Oficiais de Saída

- Gerados na pasta `MEDIÇÃO/` em formato `.xlsx` limpo e autônomo:
  1. **`MC` (Memória de Cálculo):** Fatura consolidada por item PPU, quantitativos apurados, preços unitários e valores totais. **Mantém fórmulas ativas** (`SOMASE`, subtotais e totais) vinculadas às abas `ARM` e `DADOS` locais, com tabelas estruturadas recriadas e tabelas auxiliares externas convertidas via `BreakLink` (sem avisos de vínculos quebrados).
  2. **`ARM` (Armazenagem):** Demonstrativo de caixas e mídias custodiadas no mês (dados analíticos estáticos e tabela estruturada `ListObject` recriada).
  3. **`DADOS` (Base Analítica):** Relatório detalhado chamado a chamado com SLA, datas e QExec (dados analíticos estáticos e tabela estruturada `ListObject` recriada).
  4. **`FRETE` (Transporte):** Rastreabilidade de viagens, rotas, frete agrupado e KM adicional (dados analíticos estáticos).
- Nomenclatura oficial: `<NumeroSAP>-PLA-MemoriaPetrobras-<Contrato>_<DataHora>.xlsx`.

---

## 📚 Documentos de Referência Aprofundada

Os detalhes matemáticos, códigos M e exceções estão documentados em:
- ⏱️ [Regras de SLA, Prazos Úteis e Feriados](./references/regras_sla_e_feriados.md)
- 📦 [Regras de Quantidade Executada (QExec) e Fatores de Conversão](./references/regras_qexec_e_fatores.md)
- 🚚 [Regras de Frete, Agrupamento e KM Adicional](./references/regras_frete_e_km.md)
- 📈 [Regras de FDM, Indicadores e Qualidade Contratual](./references/regras_fdm_e_qualidade.md)
- 📑 [Regras de Contratos, Armazenamento e Exportação](./references/regras_contratos_e_exportacao.md)
