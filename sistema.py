import pandas as pd
from datetime import datetime

def carregar_dados():
    estoque = pd.read_excel("estoque.xlsx")
    revendedores = pd.read_excel("revendedores.xlsx")
    vendas = pd.read_excel("vendas.xlsx")
    return estoque, revendedores, vendas

def salvar_dados(estoque, revendedores, vendas):
    estoque.to_excel("estoque.xlsx", index=False)
    revendedores.to_excel("revendedores.xlsx", index=False)
    vendas.to_excel("vendas.xlsx", index=False)

def cadastrar_revendedor(revendedores):
    nome = input("Nome do revendedor: ").strip()
    if nome in revendedores["Nome"].values:
        print("Revendedor já cadastrado.")
    else:
        revendedores.loc[len(revendedores)] = [nome]
        print("Revendedor cadastrado com sucesso.")

def registrar_saida(estoque):
    print("\nProdutos disponíveis:")
    print(estoque[["Produto", "Estoque"]])
    saida = {}
    for i, row in estoque.iterrows():
        qtd = input(f"Quantidade de '{row['Produto']}' para o revendedor: ")
        try:
            qtd = int(qtd)
        except:
            qtd = 0
        if qtd > 0:
            saida[row['Produto']] = qtd
            estoque.at[i, "Estoque"] -= qtd
    return saida, estoque

def registrar_retorno(saida, estoque, vendas, revendedor):
    print("\nRegistrando retorno do revendedor...")
    for produto, qtd_saida in saida.items():
        qtd_retorno = input(f"Quantidade devolvida de '{produto}': ")
        try:
            qtd_retorno = int(qtd_retorno)
        except:
            qtd_retorno = 0
        qtd_vendida = qtd_saida - qtd_retorno
        estoque.loc[estoque["Produto"] == produto, "Estoque"] += qtd_retorno
        preco = estoque.loc[estoque["Produto"] == produto, "Preço (R$)"].values[0]
        total = qtd_vendida * preco
        vendas.loc[len(vendas)] = [
            datetime.today().strftime("%d/%m/%Y"),
            revendedor,
            produto,
            qtd_vendida,
            round(total, 2)
        ]
    print("Vendas registradas com sucesso.\n")

def menu():
    estoque, revendedores, vendas = carregar_dados()

    while True:
        print("==== MENU ====")
        print("1. Cadastrar Revendedor")
        print("2. Registrar Saída para Revendedor")
        print("3. Registrar Retorno do Revendedor (Vendas)")
        print("4. Mostrar Estoque Atual")
        print("0. Sair")

        op = input("Escolha uma opção: ").strip()

        if op == "1":
            cadastrar_revendedor(revendedores)
        elif op == "2":
            rev = input("Nome do revendedor: ").strip()
            if rev not in revendedores["Nome"].values:
                print("Revendedor não encontrado. Cadastre primeiro.")
                continue
            global saida  # Para ser usada depois na função de retorno
            saida, estoque = registrar_saida(estoque)
            print("Saída registrada. Agora registre o retorno após a venda.")
        elif op == "3":
            rev = input("Nome do revendedor: ").strip()
            if rev not in revendedores["Nome"].values:
                print("Revendedor não encontrado.")
                continue
            registrar_retorno(saida, estoque, vendas, rev)
        elif op == "4":
            print("\nEstoque Atual:")
            print(estoque)
        elif op == "0":
            break
        else:
            print("Opção inválida.")

    salvar_dados(estoque, revendedores, vendas)
    print("Dados salvos. Encerrando.")

if __name__ == "__main__":
    menu()
