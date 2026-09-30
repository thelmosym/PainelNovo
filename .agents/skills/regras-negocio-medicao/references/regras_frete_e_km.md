# Manual Técnico de Regras: Frete, Agrupamento e KM Adicional

Este documento estabelece as regras de negócio, premissas de transporte, fórmulas matemáticas e salvaguardas de integridade para a apuração de fretes normais, fretes expressos e quilometragem adicional nos contratos da Petrobras.

---

## 1. Princípio Contratual da Remuneração de Transporte

A Petrobras contrata serviços de transporte de documentos e mídias entre seus prédios administrativos/operacionais (EDISEN, CENPES, REDUC, UO-BA, etc.) e os galpões de guarda das empresas contratadas (Itaboraí, Cajamar, Simões Filho, etc.).

A remuneração de frete segue a premissa de **viagem/deslocamento efetivo**:
- Quando uma van, caminhão ou veículo dedicado realiza um trajeto em determinado dia, transportando caixas de múltiplas Ordens de Serviço (OS), a contratada é remunerada pela **capacidade de transporte da viagem agrupada**, e não por uma tarifa cheia multiplicada por cada chamado individual.
- Isso exige duas deduplicações essenciais:
  1. **Agrupamento de Frete (Viagem Básica):** Consolidado via `fnCalcularQExecAgrupado.m`.
  2. **Deduplicação de KM Adicional:** Apurado em `Painel T2M.m`.

---

## 2. Tipos de Serviço de Frete

O catálogo PPU define duas categorias operacionais distintas:

| Código Linha PPU | Serviço de Frete | SLA de Entrega | Características Operacionais |
| :--- | :--- | :---: | :--- |
| **`FRE-NRM`** | Frete Normal | 1 a 2 dias úteis | Veículo compartilhado com roteirização programada de entregas regulares. |
| **`FRE-EXP`** | Frete Expresso | Imediato / 0,5 dia (4 horas úteis) | Veículo ou motofrete dedicado com saída emergencial e custo unitário superior. |

> [!CRITICAL]
> **Segregação Obrigatória de Tipos de Frete:**  
> Frete Normal e Frete Expresso **NUNCA** devem ser agrupados na mesma chave de viagem, mesmo que ocorram na mesma data e entre os mesmos polos. Tratam-se de veículos, contratos, SLAs e itens de faturamento distintos.

---

## 3. Função de Agrupamento de Frete (`fnCalcularQExecAgrupado.m`)

### 3.1 Composição da Chave de Agrupamento (`_ChaveAgrupamento`)
A chave composta identifica univocamente uma viagem de determinado serviço:
```powerquery
_ChaveAgrupamento = 
    (if dt = null then "" else Text.From(dt, "pt-BR"))
    & "|" & (if locP = null then "" else Text.From(locP))
    & "|" & (if locG = null then "" else Text.From(locG))
    & "|" & (if serv = null then "" else Text.From(serv))
```
- **Campos componentes:**
  1. `Data Fechamento`: data em que o frete foi efetivamente executado.
  2. `Localidade Petrobras`: polo/prédio de origem ou destino da Petrobras (ex.: `EDISEN`, `REDUC`).
  3. `Galpão_Cidade`: cidade do galpão da contratada (ex.: `ITABORAI`).
  4. `Linha de serviço PPU`: identificador do serviço (`FRE-NRM` vs `FRE-EXP`).

### 3.2 Regra de Cálculo do Grupo ($\text{soma}/10$, Piso 1)
1. Para cada chave única, cria-se um índice numérico sequencial estável (`Table.AddIndexColumn`).
2. Agrupa-se por chave para obter:
   - `_IndiceMinimo`: o menor índice da chave (a **primeira ocorrência** física no arquivo).
   - `_SomaQExec`: a soma do $QExec$ individual de todas as linhas do grupo.
3. **Fórmula de Atribuição:**
   - **Na linha da primeira ocorrência (`_Indice = _IndiceMinimo`):**
     $$\text{Se } \_SomaQExec = 0 \text{ ou nulo} \implies QExecAgrupado = 0$$
     $$\text{Se } (\_SomaQExec / 10) \le 1 \implies QExecAgrupado = 1$$
     $$\text{Se } (\_SomaQExec / 10) > 1 \implies QExecAgrupado = \frac{\_SomaQExec}{10}$$
   - **Nas demais linhas do mesmo grupo:**
     $$QExecAgrupado = \mathbf{null}$$

### 3.3 Caso Prático Real (Contrato `IRON-LT1-RJ`)
- **Cenário:** Em 11/08/2026, ocorreram atendimentos para `EDISEN` $\to$ `ITABORAI`:
  - Linha 28: Frete Normal (`FRE-NRM`) de rotina.
  - Linha 88: Frete Expresso (`FRE-EXP`) urgente (OS `CSC0892005`, Sol. `307152736`).
- **Problema anterior:** A chave não continha a linha de serviço. Ambas as ordens caíram na chave `11/08/2026|EDISEN|ITABORAI`. Como a linha 28 apareceu primeiro, ela recebeu a contagem e a linha 88 foi marcada como duplicata (`null`), fazendo com que o Frete Expresso registrasse apenas 1 viagem em vez de 2 no fechamento mensal.
- **Solução implementada:** Com a inclusão de `Linha de serviço PPU`, a chave gerou `11/08/2026|EDISEN|ITABORAI|FRE-NRM` e `11/08/2026|EDISEN|ITABORAI|FRE-EXP`, garantindo que ambas as ordens recebessem `QExecAgrupado = 1` em suas respectivas categorias.

---

## 4. Apuração de Quilometragem Adicional (KM Adicional)

### 4.1 Cruzamento com a Matriz de Distâncias
A quilometragem entre os polos é obtida pelo cruzamento relacional:
$$\text{codOrigemDestino} = \text{Centro} \ \& \ \text{Município} \ \& \ \text{Galpão} \ \& \ \text{Galpão\_Cidade}$$
A chave é associada à tabela [`Origem_Destino`](Tabelas%20Especificas/TabelaB-V2.xlsx), retornando a distância rodoviária oficial (`Origem_Destino.KM`).

### 4.2 Dedução da Franquia Contratual (`valKmAdicional`)
Os contratos preveem uma franquia de deslocamento incluída no valor básico do frete (ex.: raio de 30 km ou 50 km conforme o lote):
- Se $\text{Origem\_Destino.KM} \le valKmAdicional \implies \mathbf{KM\ Adicional = 0}$.
- Se $\text{Origem\_Destino.KM} > valKmAdicional \implies \mathbf{KM\ Adicional = KM - valKmAdicional}$.
- Se a rota não for cadastrada ou a distância for nula $\implies \mathbf{KM\ Adicional = null}$ (gera alerta na auditoria).

### 4.3 Deduplicação por Viagem
A quilometragem adicional refere-se ao deslocamento do veículo. Portanto, ela só é faturada uma única vez por viagem agrupada (primeira ocorrência da rota na data). Linhas complementares da mesma viagem recebem `KM Adicional = null` ou `0`.
