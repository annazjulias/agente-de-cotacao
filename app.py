import pandas as pd
import os


# ==============================
# PADRONIZAÇÃO DE COLUNAS
# ==============================
def tratar_colunas(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("ç", "c")
        .str.replace("ã", "a")
        .str.replace("á", "a")
        .str.replace("é", "e")
        .str.replace("í", "i")
        .str.replace("ó", "o")
        .str.replace("ú", "u")
    )
    return df


# ==============================
# TRATAR VALORES MONETÁRIOS
# ==============================
def tratar_valor(coluna):
    return (
        coluna.astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )


# ==============================
# CARREGAR FORNECEDOR (CORRIGIDO)
# ==============================
def carregar_fornecedor(caminho):
    df = pd.read_excel(caminho, engine="openpyxl")
    df = tratar_colunas(df)

    # Garantir nome padrão
    df = df.rename(columns={
        "codigo_de_barra": "codigo"
    })

    # 🔥 Pegar última coluna como preço
    ultima_coluna = df.columns[-1]

    df["preco"] = df[ultima_coluna]

    # Limpar valores
    df["preco"] = tratar_valor(df["preco"])
    df["preco"] = pd.to_numeric(df["preco"], errors="coerce")

    # Regra: 0 vira vazio
    df["preco"] = df["preco"].replace(0, None)

    # Remover linhas inválidas
    df = df.dropna(subset=["codigo"])

    return df[["codigo", "preco"]]


# ==============================
# NOME DO FORNECEDOR
# ==============================
def extrair_nome_fornecedor(nome_arquivo):
    nome = os.path.basename(nome_arquivo)
    nome = nome.replace(".xlsx", "")
    nome = nome.replace("fornecedor_", "")
    return nome.upper()


# ==============================
# ADICIONAR VÁRIOS FORNECEDORES
# ==============================
def adicionar_multiplos_fornecedores(base_df, pasta):
    if not os.path.exists(pasta):
        print("⚠️ Pasta 'fornecedores' não encontrada. Criando automaticamente...")
        os.makedirs(pasta)
        return base_df

    arquivos = [f for f in os.listdir(pasta) if f.endswith(".xlsx")]

    if not arquivos:
        print("⚠️ Nenhum fornecedor encontrado.")
        return base_df

    df_final = base_df.copy()

    for arquivo in arquivos:
        caminho = os.path.join(pasta, arquivo)

        print(f"📥 Lendo: {arquivo}")

        fornecedor_df = carregar_fornecedor(caminho)
        nome = extrair_nome_fornecedor(arquivo)

        df_final = df_final.merge(
            fornecedor_df,
            on="codigo",
            how="left"
        )

        df_final = df_final.rename(columns={"preco": nome})

    return df_final


# ==============================
# FORMATAÇÃO EXCEL (VERMELHO)
# ==============================
def aplicar_formatacao(caminho_saida, nomes_fornecedores):
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

    wb = load_workbook(caminho_saida)
    ws = wb.active

    fill_vermelho = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    # Identificar colunas
    colunas = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}

    col_ultimo = colunas.get("ultimo_valor")

    for fornecedor in nomes_fornecedores:
        col_forn = colunas.get(fornecedor)

        if not col_forn:
            continue

        for row in range(2, ws.max_row + 1):
            valor_atual = ws.cell(row=row, column=col_forn).value
            ultimo_valor = ws.cell(row=row, column=col_ultimo).value

            if valor_atual and ultimo_valor:
                if valor_atual > ultimo_valor:
                    ws.cell(row=row, column=col_forn).fill = fill_vermelho

    wb.save(caminho_saida)


# ==============================
# MAIN
# ==============================
def main():
    # Base tratada
    base = pd.read_excel("base_tratada.xlsx")

    pasta_fornecedores = "./fornecedores"

    df_final = adicionar_multiplos_fornecedores(base, pasta_fornecedores)

    print("\n📊 Preview:")
    print(df_final.head())

    # Salvar Excel
    caminho_saida = "cotacao_final.xlsx"
    df_final.to_excel(caminho_saida, index=False)

    # Pegar nomes dos fornecedores
    nomes_fornecedores = [
        col for col in df_final.columns
        if col not in ["codigo", "descricao", "apresentacao", "laboratorio", "quantidade", "ultimo_valor"]
    ]

    # Aplicar formatação
    aplicar_formatacao(caminho_saida, nomes_fornecedores)

    print("\n✅ Cotação final gerada com sucesso!")


if __name__ == "__main__":
    main()