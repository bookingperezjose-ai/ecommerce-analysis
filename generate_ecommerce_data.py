# -*- coding: utf-8 -*-
"""
generate_ecommerce_data.py

Gera uma base sintética de 1.000 transações de e-commerce no Brasil,
com regras de negócio simples para tornar os dados plausíveis
(ex.: frete grátis é mais provável em pedidos de maior valor).

Saída: sales_data.csv
"""
import random
from datetime import datetime, timedelta
import pandas as pd

# Catálogo de produtos: cada um com sua categoria e faixa de preço unitário (R$)
PRODUTOS = {
    'Notebook': ('Eletrônicos', 1800, 6000),
    'Mouse':    ('Acessórios',    30,  180),
    'Teclado':  ('Acessórios',    60,  350),
    'Monitor':  ('Eletrônicos',  700, 2500),
    'Webcam':   ('Periféricos',  120,  600),
}

CIDADES = ['São Paulo', 'Rio de Janeiro', 'Curitiba', 'Brasília', 'Salvador']

# Pesos de distribuição de pedidos por cidade (soma = 1.0)
# Refletem, de forma simplificada, o maior volume de e-commerce em São Paulo
PESOS_CIDADES = [0.35, 0.22, 0.15, 0.16, 0.12]

TOTAL_PEDIDOS = 1000

# Data fim FIXA (não datetime.now()): random.seed(42) só garante que os
# valores aleatórios saem na mesma ordem — a janela de datas usada para
# calculá-los também precisa ser fixa, ou o dataset muda todo mês conforme
# a data de execução, e os números documentados em analise_dados.md
# deixam de corresponder ao que o script gera. Mantém o pipeline 100%
# reproduzível, hoje e daqui a um ano.
DATA_FIM = datetime(2026, 6, 1)
DATA_INICIO = DATA_FIM - timedelta(days=180)  # janela de 6 meses


def gerar_dataset(total_pedidos=TOTAL_PEDIDOS, seed=42):
    """Gera um DataFrame de `total_pedidos` transações sintéticas de e-commerce."""
    random.seed(seed)
    linhas = []
    for i in range(1, total_pedidos + 1):
        produto = random.choice(list(PRODUTOS.keys()))
        categoria, preco_min, preco_max = PRODUTOS[produto]
        preco_unitario = round(random.uniform(preco_min, preco_max), 2)
        quantidade = random.randint(1, 5)

        # Data aleatória dentro da janela de 6 meses
        dias_atras = random.randint(0, 180)
        data_pedido = DATA_INICIO + timedelta(days=dias_atras)

        cidade = random.choices(CIDADES, weights=PESOS_CIDADES, k=1)[0]
        tipo_cliente = random.choices(['Novo', 'Recorrente'], weights=[0.4, 0.6], k=1)[0]

        # Desconto: a maioria dos pedidos não tem desconto ou tem um valor baixo
        desconto_pct = random.choices(
            [0, 5, 10, 15, 20],
            weights=[0.40, 0.20, 0.20, 0.12, 0.08],
            k=1
        )[0]

        valor_bruto = round(preco_unitario * quantidade, 2)

        # Regra de negócio: pedidos acima de R$ 500 têm mais chance de frete grátis
        # (simula uma política comum de "frete grátis a partir de X reais")
        prob_frete_gratis = 0.6 if valor_bruto > 500 else 0.25
        frete = 'Grátis' if random.random() < prob_frete_gratis else 'Pago'

        valor_final = round(valor_bruto * (1 - desconto_pct / 100), 2)

        linhas.append({
            'ID_Pedido': f'PED-{i:05d}',
            'Data': data_pedido.strftime('%Y-%m-%d'),
            'Produto': produto,
            'Categoria': categoria,
            'Preco_Unitario': preco_unitario,
            'Quantidade': quantidade,
            'Cidade': cidade,
            'Tipo_Cliente': tipo_cliente,
            'Desconto_Pct': desconto_pct,
            'Frete': frete,
            'Valor_Bruto': valor_bruto,
            'Valor_Final': valor_final,
        })

    return pd.DataFrame(linhas)


def main():
    df = gerar_dataset()
    df.to_csv('sales_data.csv', index=False, encoding='utf-8-sig')
    print(f'OK: {len(df)} registros gerados -> sales_data.csv')
    print(df.head(3).to_string(index=False))


if __name__ == '__main__':
    main()

