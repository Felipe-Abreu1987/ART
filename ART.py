import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from sqlite3 import Error
import re
import unicodedata
import os

# Função para normalizar texto (remover acentos e caracteres especiais)
def normalizar_texto(texto):
    if isinstance(texto, str):
        texto_normalizado = unicodedata.normalize("NFKD", texto)
        texto_sem_acentos = texto_normalizado.encode("ASCII", "ignore").decode("ASCII")
        return texto_sem_acentos.upper().strip()
    return texto

# Função para corrigir palavras separadas
def corrigir_palavras_separadas(texto):
    # Lista de palavras que devem ser unidas
    palavras_para_unir = {
        "mecânica": "mecânica",
        "produção": "produção",
        "elétrica": "elétrica",
        "civil": "civil",
        "química": "química",
        "ambiental": "ambiental",
        "sanitária": "sanitária",
        "industrial": "industrial",
        "arquitetura": "arquitetura",
        "engenharia": "engenharia",
        "tecnologia": "tecnologia",
        "construção": "construção",
        "manutenção": "manutenção",
        "instalação": "instalação",
        "projeto": "projeto",
        "sistema": "sistema",
        "gestão": "gestão",
        "qualidade": "qualidade",
        "segurança": "segurança",
        "meio ambiente": "meio ambiente",
        "infraestrutura": "infraestrutura",
        "topografia": "topografia",
        "geotecnia": "geotecnia",
        "hidráulica": "hidráulica",
        "pneumática": "pneumática",
        "telecomunicações": "telecomunicações",
        "automação": "automação",
        "energia": "energia",
        "sustentabilidade": "sustentabilidade",
        "avaliação": "avaliação"
    }

    # Verifica cada palavra na lista
    for palavra_errada, palavra_correta in palavras_para_unir.items():
        # Procura por padrões como "mec ânica" ou "produç ão"
        padrao_errado = re.compile(rf"\b{re.escape(palavra_errada)}\b", re.IGNORECASE)
        texto = padrao_errado.sub(palavra_correta, texto)

    return texto

# Função para conectar ao banco de dados SQLite
def conectar_banco():
    try:
        conn = sqlite3.connect("planilha.db")
        return conn
    except Error as e:
        messagebox.showerror("Erro", f"Erro ao conectar ao banco de dados: {e}")
        return None

# Função para criar a tabela no banco de dados
def criar_tabela(conn, colunas):
    try:
        cursor = conn.cursor()
        # Verifica se a tabela já existe
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='dados_planilha';")
        tabela_existe = cursor.fetchone()
        
        if tabela_existe:
            # Remove a tabela existente para recriá-la
            cursor.execute("DROP TABLE dados_planilha;")
            print("Tabela existente removida.")  # Depuração
        
        # Cria a tabela com as colunas esperadas
        colunas_sql = ", ".join([f"{coluna} TEXT" for coluna in colunas])
        cursor.execute(f"""
            CREATE TABLE dados_planilha (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {colunas_sql}
            );
        """)
        print("Tabela criada com as colunas:", colunas)  # Depuração
        conn.commit()
    except Error as e:
        messagebox.showerror("Erro", f"Erro ao criar tabela: {e}")

# Função para carregar a planilha no banco de dados
def carregar_planilha():
    caminho_planilha = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if caminho_planilha:
        try:
            # Remove o banco de dados existente (se houver)
            if os.path.exists("planilha.db"):
                os.remove("planilha.db")
                print("Banco de dados existente removido.")  # Depuração
            
            df = pd.read_excel(caminho_planilha)
            # Normaliza os nomes das colunas (remove acentos e espaços extras)
            df.columns = [normalizar_texto(col) for col in df.columns]
            print("Colunas normalizadas:", df.columns.tolist())  # Depuração
            
            # Define as colunas esperadas (também normalizadas)
            colunas_esperadas = ["AREA_DE_ATUACAO", "SUB_AREA_DE_ATUACAO", "OBRAS_E_SERVICOS", "COMPLEMENTO"]
            colunas_esperadas = [normalizar_texto(col) for col in colunas_esperadas]
            
            # Verifica se as colunas esperadas estão presentes na planilha
            colunas_faltando = [coluna for coluna in colunas_esperadas if coluna not in df.columns]
            if colunas_faltando:
                messagebox.showerror("Erro", f"As seguintes colunas estão faltando na planilha: {', '.join(colunas_faltando)}")
                return
            
            # Corrige as palavras separadas em todas as colunas de texto
            for coluna in df.columns:
                if df[coluna].dtype == "object":  # Verifica se a coluna é de texto
                    df[coluna] = df[coluna].apply(lambda x: corrigir_palavras_separadas(str(x)) if pd.notna(x) else "")
            
            conn = conectar_banco()
            if conn:
                # Cria a tabela no banco de dados
                criar_tabela(conn, colunas_esperadas)
                cursor = conn.cursor()
                # Insere os dados da planilha no banco de dados
                df.to_sql("dados_planilha", conn, if_exists="append", index=False)
                conn.commit()
                conn.close()
                messagebox.showinfo("Sucesso", "Planilha carregada e corrigida com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")

def buscar_palavra():
    palavra = entry_palavra.get().strip().lower()  # Obtém a palavra digitada e remove espaços extras
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()
            # Busca a palavra em todas as colunas da tabela
            colunas = ["AREA_DE_ATUACAO", "SUB_AREA_DE_ATUACAO", "OBRAS_E_SERVICOS", "COMPLEMENTO"]
            query = " OR ".join([f"{coluna} LIKE ?" for coluna in colunas])
            cursor.execute(f"SELECT * FROM dados_planilha WHERE {query};", [f"%{palavra}%"] * len(colunas))
            resultados = cursor.fetchall()
            # Limpa a Treeview antes de exibir novos resultados
            for row in treeview.get_children():
                treeview.delete(row)
            # Exibe os resultados na Treeview
            if resultados:
                for linha in resultados:
                    treeview.insert("", tk.END, values=linha[1:])  # Ignora a coluna 'id'
            else:
                messagebox.showinfo("Informação", "Nenhum resultado encontrado.")
            conn.close()
        except Error as e:
            messagebox.showerror("Erro", f"Erro ao buscar dados: {e}")

def limpar_busca():
    entry_palavra.delete(0, tk.END)  # Limpa o campo de entrada
    for row in treeview.get_children():
        treeview.delete(row)  # Limpa a Treeview

# Interface gráfica
root = tk.Tk()
root.title("Sistema de Busca em Planilha com SQLite")
root.geometry("1000x600")  # Define o tamanho da janela

# Estilo para melhorar a aparência
style = ttk.Style()
style.configure("TButton", padding=5, font=("Arial", 10))
style.configure("TLabel", font=("Arial", 12))
style.configure("TEntry", font=("Arial", 12))

# Frame para organizar os componentes
frame_superior = ttk.Frame(root)
frame_superior.pack(pady=10, padx=10, fill=tk.X)

# Botão para carregar a planilha
btn_carregar = ttk.Button(frame_superior, text="Carregar Planilha", command=carregar_planilha)
btn_carregar.grid(row=0, column=0, padx=5, pady=5)

# Caixa de entrada para a palavra
label_palavra = ttk.Label(frame_superior, text="Insira a palavra ou expressão:")
label_palavra.grid(row=0, column=1, padx=5, pady=5)

entry_palavra = ttk.Entry(frame_superior, width=50)
entry_palavra.grid(row=0, column=2, padx=5, pady=5)

# Botão para buscar
btn_buscar = ttk.Button(frame_superior, text="Buscar", command=buscar_palavra)
btn_buscar.grid(row=0, column=3, padx=5, pady=5)

# Botão para limpar a busca
btn_limpar = ttk.Button(frame_superior, text="Limpar", command=limpar_busca)
btn_limpar.grid(row=0, column=4, padx=5, pady=5)

# Treeview para exibir os resultados
frame_resultado = ttk.Frame(root)
frame_resultado.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# Cria a Treeview com as colunas específicas
colunas = ["AREA DE ATUACAO", "SUB AREA DE ATUACAO", "OBRAS E SERVICOS", "COMPLEMENTO"]
treeview = ttk.Treeview(frame_resultado, columns=colunas, show="headings")
treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Configura as colunas da Treeview
for coluna in colunas:
    treeview.heading(coluna, text=coluna)
    treeview.column(coluna, width=150, anchor=tk.W)

# Adiciona uma barra de rolagem
scrollbar = ttk.Scrollbar(frame_resultado, orient=tk.VERTICAL, command=treeview.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
treeview.configure(yscrollcommand=scrollbar.set)

# Inicializa a interface
root.mainloop()




