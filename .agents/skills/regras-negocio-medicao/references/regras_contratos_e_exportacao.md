# Manual Técnico de Regras: Contratos, Armazenamento (ARM) e Exportação de Cadernos

Este documento consolida os dados jurídicos e operacionais dos contratos Petrobras, as regras de faturamento de armazenagem e custódia física (**ARM**) e as especificações técnicas de geração dos cadernos oficiais de faturamento.

---

## 1. Cadastro Geral de Contratos e Lotes

A prestação de serviços de apoio operacional e guarda documental da Petrobras é dividida em dois lotes geográficos principais:

### 1.1 Lote 2 - Bahia (`PA-LT2`)
- **Número do Contrato SAP:** `4600687006`
- **Instrumento Jurídico Petrobras:** `5900.0133214.25.2`
- **Prestadora Contratada:** PA
- **Abrangência Geográfica:** Estado da Bahia (`BA`)
- **Polos Atendidos:** Salvador, Camaçari, Simões Filho, São Sebastião do Passé e bases operacionais da UO-BA.
- **Identificador de Saída:** `PA-LT2-BA`

### 1.2 Lote 1 - Sudeste (`IRON-LT1`)
- **Número do Contrato SAP:** `4600686987`
- **Instrumento Jurídico Petrobras:** `5900.0133210.25.2`
- **Prestadora Contratada:** IRON MOUNTAIN
- **Abrangência Geográfica:** Estados do Rio de Janeiro (`RJ`), São Paulo (`SP`) e Espírito Santo (`ES`).
- **Segmentação Operacional dos Cadernos:**
  - **`IRON-LT1-RJ`:** Atendimento aos complexos EDISEN, CENPES, REDUC, Macaé, Imbetiba e Cabiúnas. Galpão principal em Itaboraí.
  - **`IRON-LT1-SP`:** Atendimento aos complexos da Av. Paulista, Santos, RPBC (Cubatão), RECAP (Mauá) e REPLAN (Paulínia). Galpão principal em Cajamar.
  - **`IRON-LT1-ES`:** Atendimento à Sede Vitória e Base Operacional de São Mateus (UO-ES).

---

## 2. Regras de Negócio de Armazenagem e Custódia (`ARM`)

O serviço de armazenagem remunera a contratada pela guarda segura, climatização, seguro predial e custódia física do acervo documental e mídias técnicas da Petrobras.

### 2.1 Posição e Saldo de Faturamento
- A apuração é **mensal**, com corte estrito na data de fechamento do ciclo (**dia 25 do mês**).
- **Fórmula de Fechamento de Saldo:**
  $$\text{Saldo Final} = \text{Saldo Inicial} + \text{Entradas (Novos Arquivamentos)} - \text{Saídas (Baixas Definitivas)}$$
- Materiais em empréstimo temporário (consultas externas) continuam sob responsabilidade contratual e compõem o saldo faturável da armazenagem.

### 2.2 Itens e Unidades de Custódia Controlados

| Item de Armazenagem | Unidade PPU | Especificação Física |
| :--- | :---: | :--- |
| **Caixa Box Padrão (20kg)** | Caixa/Mês | Dimensões regulamentares (~0,04 m³), suportando até 20 kg de documentos papel. |
| **Mídia Magnética / Óptica** | Fita/Mês | Fitas LTO, fitas sísmicas e mídias graváveis mantidas em cofre climatizado de alta segurança. |
| **Tubos Especiais** | Tubo/Mês | Embalagens tubulares para perfis de poço de petróleo e pranchas cartográficas/arquitetura. |
| **Prateleiras / Metros Cúbicos** | $m^3$/Mês | Cobrança por espaço dedicado em estantes industriais quando estipulado em termo aditivo. |

---

## 3. Especificações de Geração dos Cadernos Oficiais (`MEDIÇÃO/*.xlsx`)

A rotina automatizada [`modGerarArquivos.bas`](Codes/VBA/Fun%C3%A7%C3%B5es%20Novas/modGerarArquivos.bas) (`GerarArquivosPorContrato`) é responsável por produzir os arquivos oficiais entregues para ateste da fiscalização técnica da Petrobras.

### 3.1 Padrão de Nomenclatura dos Arquivos
Os cadernos exportados seguem o padrão formal:
```text
[NúmeroSAP]-PLA-MemoriaPetrobras-[IdentificadorContrato]_[DDMMAAAA]_[HHMMSS].xlsx
```
*Exemplos Reais:*
- `4600687006-PLA-MemoriaPetrobras-PA-LT2-BA_29092026_180000.xlsx`
- `4600686987-PLA-MemoriaPetrobras-IRON-LT1-RJ_29092026_180000.xlsx`
- `4600686987-PLA-MemoriaPetrobras-IRON-LT1-SP_29092026_180000.xlsx`
- `4600686987-PLA-MemoriaPetrobras-IRON-LT1-ES_29092026_180000.xlsx`

### 3.2 Estrutura Padronizada das 4 Abas Contratuais
Cada caderno gerado em formato puro `.xlsx` (desprovido de macros e sem conexões externas) deve conter exatamente as seguintes abas:

```
┌────────────────────────────────────────────────────────────────────────┐
│                        ESTRUTURA DO CADERNO DE MEDIÇÃO                 │
├───────────────┬────────────────────────────────────────────────────────┤
│ 1. MC         │ Memória de Cálculo Consolidada                         │
│               │ Tabela resumo com Item PPU, Descrição do Serviço,      │
│               │ Unidade, Quantidade Total Apurada, Preço Unitário      │
│               │ Contratual e Valor Total a Faturar (R$).               │
├───────────────┼────────────────────────────────────────────────────────┤
│ 2. ARM        │ Demonstrativo de Armazenagem                           │
│               │ Saldo físico de caixas, fitas e tubos sob custódia,    │
│               │ movimentações de entrada/saída e valor mensal apurado. │
├───────────────┼────────────────────────────────────────────────────────┤
│ 3. DADOS      │ Base Analítica de Atendimentos                         │
│               │ Rastreabilidade completa chamada a chamada:            │
│               │ OS, Solicitação, Datas (Abertura/Fechamento), Prazo    │
│               │ Limite SLA, Status de Prazo, Qtd Atendida e QExec.     │
├───────────────┼────────────────────────────────────────────────────────┤
│ 4. FRETE      │ Relatório de Transporte e Deslocamentos                │
│               │ Viagens agrupadas, rotas, KM excedente apurado, fretes │
│               │ normais e expressos faturados sem duplicidade.         │
└───────────────┴────────────────────────────────────────────────────────┘
```

### 3.3 Regras de Limpeza e Integridade na Exportação
1. **Valores Estáticos (Paste Values):** Fórmulas voláteis e referências externas a outros cadernos são convertidas em valores puros, garantindo que o fiscal da Petrobras abra o arquivo sem avisos de "Vínculos Quebrados".
2. **Registro de Log e Hiperlinks no Painel:** Após a geração, os caminhos absolutos e hiperlinks clicáveis para os arquivos gerados são registrados na aba `Painel` (intervalo `K7:K10`), permitindo abertura imediata pelo operador.
