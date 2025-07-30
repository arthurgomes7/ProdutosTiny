import pandas as pd
import os
from datetime import datetime

# === CAMINHO DO ARQUIVO LOCAL EXISTENTE ===
caminho_original = 'produtos.xls'  # <-- ALTEREs AQUI

# Carrega a planilha existente
df_existente = pd.read_excel(caminho_original)

# Lista para novos produtos
novos_produtos = []
os.system('cls')
print("\nAdicione novos produtos. Digite 'sair' em qualquer campo obrigatório para encerrar.\n")

while True:
    descricao = input("Nome (OBRIGATÓRIO): ")
    if descricao.lower() == 'sair':
        break

    classificacao_fiscal = input("Classificação fiscal (OBRIGATÓRIO): ")
    if classificacao_fiscal.lower() == 'sair':
        break

    preco = input("Preço (OBRIGATÓRIO): ")
    if preco.lower() == 'sair':
        break

    gtin = input("GTIN/EAN (OBRIGATÓRIO): ")
    if gtin.lower() == 'sair':
        break

    # Campos opcionais
    sku = input("Código (SKU) (opcional): ")
    if sku.lower() == 'sair':
        break

    unidade = input("Unidade (opcional): ")
    if unidade.lower() == 'sair':
        break

    # Campos fixos
    origem = 0
    situacao = 'ATIVO'
    estoque = 1000
    categoria = 'LOJA'
    marca = 'Loja Fisica'
    sob_encomenda = 'NÃO'

    produto = {
        'Código (SKU)': sku,
        'Descrição': descricao,
        'Unidade': unidade,
        'NCM (Classificação fiscal)': classificacao_fiscal,
        'Origem': origem,
        'Preço': preco,
        'Situação': situacao,
        'Estoque': estoque,
        'GTIN/EAN': gtin,
        'Categoria': categoria,
        'Marca': marca,
        'Sob encomenda': sob_encomenda
    }

    novos_produtos.append(produto)
    print("✅ Produto adicionado!\n")

# DataFrame com novos produtos
df_novos = pd.DataFrame(novos_produtos)

# Junta com os produtos antigos
df_final = pd.concat([df_existente, df_novos], ignore_index=True)

# Cria nome da nova planilha
data = datetime.now().strftime('%Y%m%d_%H%M%S')
novo_arquivo = f'produtos_atualizados_{data}.xlsx'  # <-- ALTERE AQUI

# Salva nova planilha
df_final.to_excel(novo_arquivo, index=False)
print(f"\n📁 Nova planilha salva com sucesso em: {novo_arquivo}")
