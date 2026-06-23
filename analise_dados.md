# Análise de Dados — E-commerce Brasil

> Todos os números deste documento foram recalculados diretamente a partir da execução real de `generate_ecommerce_data.py` + `process_ecommerce_sales.py` (não estimados). Cada valor corresponde exatamente às tabelas da aba **Análise** em `sales_report_FINAL.xlsx`. O gerador usa uma data-fim fixa (`DATA_FIM = datetime(2026, 6, 1)`), então rodar o pipeline em qualquer máquina, hoje ou daqui a um ano, reproduz exatamente os mesmos números.

---

## Resumo Executivo

| Métrica | Valor |
|---|---|
| Receita total | R$ 3.436.060,91 |
| Pedidos | 1.000 |
| Ticket médio | R$ 3.436,06 |
| Clientes novos | 406 (40,6%) |
| Clientes recorrentes | 594 (59,4%) |

---

## 1. Vendas por Mês

| Mês | Vendas Totais | Pedidos | Ticket Médio |
|---|---|---|---|
| 2025-12 | R$ 603.513,97 | 164 | R$ 3.679,96 |
| 2026-01 | R$ 656.229,75 | 176 | R$ 3.728,58 |
| 2026-02 | R$ 521.162,57 | 153 | R$ 3.406,29 |
| 2026-03 | R$ 510.489,86 | 169 | R$ 3.020,65 |
| 2026-04 | R$ 580.756,03 | 154 | R$ 3.771,14 |
| 2026-05 | R$ 554.023,64 | 175 | R$ 3.165,85 |
| 2026-06 | R$ 9.885,09 | 9 | R$ 1.098,34 |

**Leitura:** Junho de 2026 aparece com apenas 9 pedidos porque a janela de geração termina em `2026-06-01` — esse mês está incompleto por desenho, não por queda real de vendas. Os demais 6 meses têm volume comparável (153 a 176 pedidos), sem tendência clara de crescimento ou queda no período.

---

## 2. Concentração por Cidade

| Cidade | Receita | % do Total | Pedidos | Ticket Médio |
|---|---|---|---|---|
| São Paulo | R$ 1.267.032,81 | 36,9% | 367 | R$ 3.452,41 |
| Rio de Janeiro | R$ 729.200,54 | 21,2% | 214 | R$ 3.407,48 |
| Brasília | R$ 647.805,29 | 18,9% | 165 | R$ 3.926,09 |
| Curitiba | R$ 495.778,18 | 14,4% | 156 | R$ 3.178,07 |
| Salvador | R$ 296.244,09 | 8,6% | 98 | R$ 3.022,90 |

**Leitura:** São Paulo concentra 36,9% da receita — coerente com o peso real do mercado paulista em e-commerce no Brasil. Brasília tem o **maior ticket médio** (R$ 3.926,09) com menor volume.

---

## 3. Mix de Produtos

| Produto | Receita | % do Total | Pedidos | Ticket Médio |
|---|---|---|---|---|
| Notebook | R$ 2.086.544,47 | 60,7% | 181 | R$ 11.527,87 |
| Monitor | R$ 976.221,79 | 28,4% | 226 | R$ 4.319,57 |
| Webcam | R$ 190.022,10 | 5,5% | 202 | R$ 940,70 |
| Teclado | R$ 127.384,54 | 3,7% | 199 | R$ 640,12 |
| Mouse | R$ 55.888,01 | 1,6% | 192 | R$ 291,08 |

**Leitura:** Notebook responde por **60,7% da receita com apenas 18,1% dos pedidos** — produto de maior valor unitário e maior concentração de risco. Eletrônicos (Notebook + Monitor) somam 89,1% da receita total.

---

## 4. Clientes: Novos vs. Recorrentes

| Tipo | Pedidos | % |
|---|---|---|
| Novos | 406 | 40,6% |
| Recorrentes | 594 | 59,4% |

**Leitura:** Quase 6 em cada 10 pedidos vêm de clientes recorrentes. O dataset não tem identificador único de cliente — apenas o rótulo Novo/Recorrente por pedido — então LTV, churn ou frequência de recompra real não podem ser calculados a partir dele.

---

## 5. Frete Grátis vs. Pago

| Frete | Receita | Pedidos | Ticket Médio |
|---|---|---|---|
| Grátis | R$ 1.957.826,44 | 479 | **R$ 4.087,32** |
| Pago | R$ 1.478.234,47 | 521 | R$ 2.837,30 |

**Diferença no ticket médio: +44,1%**

**Leitura importante:** Essa diferença é causada pela própria regra de geração dos dados — pedidos acima de R$ 500 têm probabilidade muito maior de receber frete grátis (60% vs. 25%). **Frete grátis não está causando ticket maior aqui — é o ticket maior que aumenta a chance de o pedido ter frete grátis.** É uma correlação esperada pelo desenho do dataset sintético, não uma descoberta de comportamento real de consumidor.

---

## 6. Desconto Aplicado

| % Desconto | Receita | Pedidos | Ticket Médio |
|---|---|---|---|
| 0% | R$ 1.620.986,86 | 430 | R$ 3.769,74 |
| 5% | R$ 659.215,94 | 194 | R$ 3.398,02 |
| 10% | R$ 654.411,86 | 196 | R$ 3.338,84 |
| 15% | R$ 289.749,89 | 104 | R$ 2.786,06 |
| 20% | R$ 211.696,36 | 76 | R$ 2.785,48 |

**Leitura:** O ticket médio cai conforme o desconto aumenta. Como os descontos foram atribuídos de forma aleatória no gerador, essa tendência reflete o desenho do dataset, não um padrão real de comportamento de compra.

---

## Limitações deste Dataset (declaradas, não escondidas)

Este é um **dataset sintético gerado para fins de portfólio**, não dados de um negócio real:

- Correlações como "frete grátis aumenta ticket médio" refletem **regras do gerador**, não comportamento real de consumidor.
- Não há identificador único de cliente, então **LTV, churn e recompra real não podem ser calculados**.
- O mês de 2026-06 está incompleto por desenho — não deve ser comparado com os demais meses em gráficos de tendência.
- Os valores de receita, embora realistas em ordem de grandeza, são arbitrários.

O objetivo deste projeto é demonstrar o **pipeline técnico completo** e a capacidade de interpretar dados corretamente — incluindo reconhecer os limites do que os dados sustentam.

---

## Reprodutibilidade

`generate_ecommerce_data.py` usa `random.seed(42)` e uma data-fim fixa (`datetime(2026, 6, 1)`). Verificado executando o pipeline duas vezes em momentos diferentes e comparando o hash MD5 de `sales_data.csv` — o arquivo é byte a byte idêntico nas duas execuções.

## Fonte

- `sales_data.csv` — gerado por `generate_ecommerce_data.py`
- `sales_processed.csv` / `sales_report_FINAL.xlsx` — gerados por `process_ecommerce_sales.py`
