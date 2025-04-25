import pandas as pd
from datetime import datetime
import os
from tkinter import Tk, filedialog, messagebox, Label, Button, Toplevel

# Função para exibir a janela de apresentação
def exibir_janela_apresentacao():
    janela = Toplevel()
    janela.title("Bem vindo")
    janela.geometry("400x150")  # Tamanho da janela

    # Mensagem de boas-vindas
    mensagem = Label(
        janela,
        text="Programa para gerar planilha de proventos para controle manual",
        font=("Arial", 12),
        wraplength=350  # Define a largura máxima do texto antes de quebrar a linha
    )
    mensagem.pack(pady=20)

    # Variável para controlar se o usuário clicou em "Iniciar"
    iniciar_programa = False

    # Função chamada quando o botão "Iniciar" é clicado
    def iniciar():
        nonlocal iniciar_programa
        iniciar_programa = True
        janela.destroy()  # Fecha a janela de apresentação

    # Botão para iniciar o programa
    botao_iniciar = Button(
        janela,
        text="Iniciar",
        font=("Arial", 12),
        command=iniciar  # Chama a função iniciar ao clicar
    )
    botao_iniciar.pack(pady=10)

    # Esperar o usuário clicar em "Iniciar"
    janela.wait_window()  # Bloqueia até que a janela seja fechada

    return iniciar_programa

# Função para salvar o arquivo com verificação de nome existente
def salvar_arquivo_com_verificacao(caminho_padrao):
    while True:
        # Abrir caixa de diálogo para salvar o arquivo
        caminho_salvar = filedialog.asksaveasfilename(
            title="Salvar como",
            defaultextension=".xlsx",  # Forçar extensão .xlsx
            filetypes=[("Excel files", "*.xlsx")],  # Permitir apenas .xlsx
            initialfile=os.path.basename(caminho_padrao),  # Sugerir o nome padrão
            initialdir=os.path.dirname(caminho_padrao)  # Sugerir o diretório padrão
        )

        # Se o usuário cancelar a operação
        if not caminho_salvar:
            return None

        # Verificar se o arquivo já existe
        if os.path.exists(caminho_salvar):
            resposta = messagebox.askyesno(
                "Arquivo já existe",
                f"O arquivo '{os.path.basename(caminho_salvar)}' já existe. Deseja sobrescrever?"
            )
            if not resposta:
                continue  # Pedir para escolher outro nome
            else:
                return caminho_salvar  # Sobrescrever o arquivo
        else:
            return caminho_salvar  # Salvar o arquivo

# Configurar a interface gráfica principal
root = Tk()
root.withdraw()  # Esconder a janela principal

# Exibir a janela de apresentação e esperar o usuário clicar em "Iniciar"
iniciar_programa = exibir_janela_apresentacao()

# Se o usuário não clicou em "Iniciar", encerrar o programa
if not iniciar_programa:
    print("Programa encerrado pelo usuário.")
    exit()

# Solicitar ao usuário o caminho do arquivo de entrada
arquivo_path = filedialog.askopenfilename(
    title="Selecione o arquivo da B3 (xls, xlsx ou csv):",
    filetypes=[("Excel files", "*.xls *.xlsx"), ("CSV files", "*.csv")]  # Permitir .xls, .xlsx e .csv
)

# Verificar se o usuário selecionou um arquivo
if not arquivo_path:
    print("Nenhum arquivo selecionado. O programa será encerrado.")
    exit()

# Verificar se o arquivo existe
if not os.path.exists(arquivo_path):
    print(f"Arquivo {arquivo_path} não encontrado.")
    exit()

# Verificar a extensão do arquivo de entrada
_, extensao = os.path.splitext(arquivo_path)
extensao = extensao.lower()  # Garantir que a extensão esteja em minúsculas

# Verificar se o arquivo é para o exterior
is_exterior = "exterior" in arquivo_path.lower()

# Criar um DataFrame em branco com as colunas padrão
colunas_padrao = [
    "Ativo", "Associado", "Data", "Qtde", "Preço Médio", "Valor", "Strike",
    "Fluxo", "Estratégia1", "Estratégia2", "Estratégia3", "Tipo", "Geografia", "Operação", "Corretora"
]
proventos_df = pd.DataFrame(columns=colunas_padrao)

# Ler o arquivo de entrada com base na extensão
try:
    if extensao == ".xlsx":
        william_df = pd.read_excel(arquivo_path)
    elif extensao == ".xls":
        william_df = pd.read_excel(arquivo_path, engine="xlrd")  # Usar engine xlrd para .xls
    elif extensao == ".csv":
        william_df = pd.read_csv(arquivo_path)  # Ler arquivo CSV
    else:
        print(f"Formato de arquivo não suportado: {extensao}")
        exit()
except Exception as e:
    print(f"Erro ao ler arquivo: {e}")
    exit()

# Imprimir as colunas do arquivo carregado para ver o nome exato
print(f"Colunas do arquivo {arquivo_path}: {william_df.columns.tolist()}")

# Preencher o DataFrame com base no tipo de arquivo (exterior ou nacional)
if is_exterior:
    # Caso exterior
    proventos_df["Ativo"] = william_df["Ativo"]
    proventos_df["Associado"] = proventos_df["Ativo"]
    proventos_df["Data"] = pd.to_datetime(william_df["Data pgto."], format="%d/%m/%Y", errors="coerce").dt.strftime("%d/%m/%Y")  # Formatar data
    proventos_df["Qtde"] = pd.to_numeric(william_df["Cotas"], errors="coerce").fillna(0).astype(int)
    proventos_df["Preço Médio"] = pd.to_numeric(william_df["Preço médio"], errors="coerce").fillna(0)  # Corrigir preço médio
    proventos_df["Valor"] = william_df["Recebido"].replace('[\$,]', '', regex=True).str.replace('.', ',', regex=False)  # Formatar valor
    proventos_df["Strike"] = None  # Não há correspondência direta
    proventos_df["Fluxo"] = "Entrada"
    proventos_df["Estratégia1"] = "Proventos"
    proventos_df["Estratégia2"] = "Ações"
    proventos_df["Estratégia3"] = "Proventos"
    proventos_df["Tipo"] = "Proventos"
    proventos_df["Geografia"] = "Exterior"
    proventos_df["Operação"] = "Ações Exterior"
    proventos_df["Corretora"] = "Avenue"
else:
    # Caso padrão (nacional)
    proventos_df["Ativo"] = william_df["Produto"].str.split(" - ").str[0]
    proventos_df["Associado"] = proventos_df["Ativo"]
    if 'Previsão de pagamento' in william_df.columns:
        proventos_df["Data"] = pd.to_datetime(william_df["Previsão de pagamento"], format="%d/%m/%Y", errors="coerce").dt.strftime("%d/%m/%Y")  # Formatar data
    elif 'Pagamento' in william_df.columns:
        proventos_df["Data"] = pd.to_datetime(william_df["Pagamento"], format="%d/%m/%Y", errors="coerce").dt.strftime("%d/%m/%Y")  # Formatar data
    else:
        proventos_df["Data"] = pd.NaT  # Se não houver, preencher com NaT (valor nulo de data)

    # Se não houver data, definir como 01/12 do ano atual
    data_padrao = datetime(datetime.now().year, 12, 1).strftime("%d/%m/%Y")
    proventos_df["Data"] = proventos_df["Data"].apply(lambda x: x if pd.notna(x) else data_padrao)

    proventos_df["Qtde"] = pd.to_numeric(william_df["Quantidade"], errors="coerce").fillna(0).astype(int)
    proventos_df["Preço Médio"] = pd.to_numeric(william_df["Preço unitário"], errors="coerce").fillna(0)
    proventos_df["Valor"] = william_df["Valor líquido"].fillna(0)
    proventos_df["Strike"] = None  # Não há correspondência direta
    proventos_df["Fluxo"] = "Entrada"
    proventos_df["Estratégia1"] = "Proventos"
    proventos_df["Estratégia2"] = "Ações"
    proventos_df["Estratégia3"] = "Proventos"  # Valor padrão para Estratégia3
    proventos_df["Tipo"] = "Proventos"
    proventos_df["Geografia"] = "Nacional"
    proventos_df["Operação"] = "Ações"
    proventos_df["Corretora"] = "BTG"

    # Ajuste para o caso de Isadora, William Brasil e William Exterior
    if "isadora" in arquivo_path.lower():
        proventos_df["Estratégia3"] = "Isadora"
    elif "brasil" in arquivo_path.lower():
        proventos_df["Estratégia3"] = "Proventos"  # Valor para William Brasil
    elif "exterior" in arquivo_path.lower():
        proventos_df["Estratégia3"] = "Proventos"  # Valor para William Exterior
        proventos_df["Geografia"] = "Exterior"  # Alteração para William Exterior
    else:
        print("Tipo de arquivo não reconhecido. Verifique o nome do arquivo.")
        exit()

# Remover linhas onde "Ativo" é vazio
proventos_df = proventos_df[proventos_df["Ativo"].notna() & (proventos_df["Ativo"] != "")]

# Solicitar ao usuário onde salvar o arquivo
caminho_padrao_salvar = os.path.join(os.getcwd(), "proventos.xlsx")  # Caminho padrão
caminho_salvar = salvar_arquivo_com_verificacao(caminho_padrao_salvar)

# Se o usuário cancelar a operação de salvar
if not caminho_salvar:
    print("Operação de salvar cancelada pelo usuário.")
    exit()

# Salvar o novo arquivo (sempre como .xlsx)
proventos_df.to_excel(caminho_salvar, index=False)

print(f"Arquivo salvo como {caminho_salvar}")