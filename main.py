import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import pandas as pd

class AppBIExcel:
    def __init__(self, root):
        self.root = root
        self.root.title("Dashboard Interativo - Excel BI")
        self.root.geometry("1200x750")
        
        # Configuração de estilo para os seletores modernos do Tkinter
        style = ttk.Style()
        style.theme_use('clam')

        # Variáveis de controle interno de dados
        self.df = None  
        self.df_filtrado = None  
        self.caminho_arquivo = ""

        # Layout Principal (Esquerda: Filtros | Direita: Gráfico BI)
        self.frame_esquerdo = tk.Frame(self.root, width=280, bg="#f5f5f7", bd=1, relief=tk.SOLID)
        self.frame_esquerdo.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        self.frame_esquerdo.pack_propagate(False)

        self.frame_direito = tk.Frame(self.root, bg="white")
        self.frame_direito.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Inicialização dos componentes gráficos
        self.criar_widgets_filtros()
        self.criar_area_grafico()

    def criar_widgets_filtros(self):
        """Cria os botões de carregar arquivo e os seletores de colunas/filtros."""
        self.btn_carregar = tk.Button(
            self.frame_esquerdo,
            text="📁 Carregar Excel",
            command=self.carregar_excel,
            bg="#007AFF",
            fg="white",
            font=("Arial", 10, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            pady=6
        )
        self.btn_carregar.pack(fill=tk.X, padx=10, pady=15)

        self.lbl_arquivo = tk.Label(
            self.frame_esquerdo, text="Nenhum arquivo selecionado", wraplength=240, fg="#666666", bg="#f5f5f7", font=("Arial", 9)
        )
        self.lbl_arquivo.pack(fill=tk.X, padx=10, pady=5)

        tk.Frame(self.frame_esquerdo, height=1, bg="#d1d1d6").pack(fill=tk.X, pady=10, padx=10)

        tk.Label(self.frame_esquerdo, text="Eixo X (Categorias):", bg="#f5f5f7", font=("Arial", 9, "bold"), fg="#333333").pack(
            anchor=tk.W, padx=10, pady=(5, 2)
        )
        self.cb_eixo_x = ttk.Combobox(self.frame_esquerdo, state="readonly")
        self.cb_eixo_x.pack(fill=tk.X, padx=10, pady=5)
        self.cb_eixo_x.bind("<<ComboboxSelected>>", self.atualizar_filtros_disponiveis)

        tk.Label(self.frame_esquerdo, text="Eixo Y (Métricas):", bg="#f5f5f7", font=("Arial", 9, "bold"), fg="#333333").pack(
            anchor=tk.W, padx=10, pady=(5, 2)
        )
        self.cb_eixo_y = ttk.Combobox(self.frame_esquerdo, state="readonly")
        self.cb_eixo_y.pack(fill=tk.X, padx=10, pady=5)
        self.cb_eixo_y.bind("<<ComboboxSelected>>", lambda e: self.gerar_grafico())

        tk.Label(self.frame_esquerdo, text="Tipo de Gráfico:", bg="#f5f5f7", font=("Arial", 9, "bold"), fg="#333333").pack(
            anchor=tk.W, padx=10, pady=(5, 2)
        )
        self.cb_tipo_grafico = ttk.Combobox(self.frame_esquerdo, state="readonly")
        self.cb_tipo_grafico["values"] = ["Barras", "Linhas", "Pizza", "Área"]
        self.cb_tipo_grafico.set("Barras")
        self.cb_tipo_grafico.pack(fill=tk.X, padx=10, pady=5)
        self.cb_tipo_grafico.bind("<<ComboboxSelected>>", lambda e: self.gerar_grafico())

        tk.Frame(self.frame_esquerdo, height=1, bg="#d1d1d6").pack(fill=tk.X, pady=10, padx=10)

        tk.Label(self.frame_esquerdo, text="Filtrar Elementos do Eixo X:", bg="#f5f5f7", font=("Arial", 9, "bold"), fg="#333333").pack(
            anchor=tk.W, padx=10, pady=(5, 2)
        )

        frame_listbox = tk.Frame(self.frame_esquerdo, bg="#f5f5f7")
        frame_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.scrollbar_filtro = tk.Scrollbar(frame_listbox, orient=tk.VERTICAL)
        self.listbox_filtros = tk.Listbox(
            frame_listbox, selectmode=tk.MULTIPLE, yscrollcommand=self.scrollbar_filtro.set, bd=1, relief=tk.SOLID
        )
        self.scrollbar_filtro.config(command=self.listbox_filtros.yview)

        self.scrollbar_filtro.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_filtros.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.btn_filtrar = tk.Button(
            self.frame_esquerdo,
            text="⚡ Aplicar Filtros",
            command=self.aplicar_filtros_dados,
            bg="#34C759",
            fg="white",
            font=("Arial", 10, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            pady=6
        )
        self.btn_filtrar.pack(fill=tk.X, padx=10, pady=15)

    def criar_area_grafico(self):
        """Prepara o container onde o gráfico do Matplotlib será renderizado."""
        self.fig, self.ax = plt.subplots(figsize=(8, 5), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.frame_direito)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill=tk.BOTH, expand=True)

        self.toolbar = NavigationToolbar2Tk(self.canvas, self.frame_direito)
        self.toolbar.update()

    def carregar_excel(self):
        """Garante a leitura automática sempre da primeira aba ('casa') do Excel."""
        self.caminho_arquivo = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )

        if self.caminho_arquivo:
            try:
                # Pega SEMPRE a primeira aba da planilha
                self.df = pd.read_excel(self.caminho_arquivo, sheet_name=0)
                self.df_filtrado = self.df.copy()

                self.lbl_arquivo.config(text=self.caminho_arquivo.split("/")[-1])

                colunas = list(self.df.columns)
                self.cb_eixo_x["values"] = colunas
                self.cb_eixo_y["values"] = colunas

                if len(colunas) >= 2:
                    self.cb_eixo_x.set(colunas[0])
                    self.cb_eixo_y.set(colunas[1])

                self.atualizar_filtros_disponiveis()

            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{str(e)}")

    def atualizar_filtros_disponiveis(self, event=None):
        """Atualiza as opções de filtro baseadas no Eixo X."""
        if self.df is None:
            return

        coluna_x = self.cb_eixo_x.get()
        self.listbox_filtros.delete(0, tk.END)

        valores_unicos = self.df[coluna_x].dropna().unique()

        for valor in valores_unicos:
            self.listbox_filtros.insert(tk.END, str(valor))

        self.listbox_filtros.select_set(0, tk.END)
        self.df_filtrado = self.df.copy()
        self.gerar_grafico()

    def aplicar_filtros_dados(self):
        """Aplica os critérios de filtragem selecionados no Listbox."""
        if self.df is None:
            return

        coluna_x = self.cb_eixo_x.get()
        indices_selecionados = self.listbox_filtros.curselection()
        valores_selecionados = [self.listbox_filtros.get(i) for i in indices_selecionados]

        if not valores_selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um item para filtrar.")
            return

        self.df_filtrado = self.df[self.df[coluna_x].astype(str).isin(valores_selecionados)]
        self.gerar_grafico()

    def gerar_grafico(self):
        """Gera o gráfico limpando duplicatas, aplicando Top N e ajustando margens manuais."""
        if self.df_filtrado is None:
            return

        col_x = self.cb_eixo_x.get()
        col_y = self.cb_eixo_y.get()
        tipo = self.cb_tipo_grafico.get()

        if not col_x or not col_y:
            return

        try:
            self.ax.clear()

            # 1. Agrupamento (Soma para números, contagem para texto)
            if pd.api.types.is_numeric_dtype(self.df_filtrado[col_y]):
                dados_agrupados = self.df_filtrado.groupby(col_x)[col_y].sum().reset_index()
                label_y = f"Soma de {col_y}"
            else:
                dados_agrupados = self.df_filtrado.groupby(col_x)[col_y].count().reset_index()
                label_y = f"Contagem de {col_y}"

            # 2. Ordenação Decrescente (Padrão BI)
            dados_agrupados = dados_agrupados.sort_values(by=col_y, ascending=False)
            
            # Limite máximo de categorias (Top 10 + Outros) para evitar poluição visual
            LIMITE_ITENS = 10 

            if len(dados_agrupados) > LIMITE_ITENS:
                top_itens = dados_agrupados.head(LIMITE_ITENS).copy()
                resto = dados_agrupados.tail(len(dados_agrupados) - LIMITE_ITENS)
                
                soma_outros = resto[col_y].sum()
                linha_outros = pd.DataFrame([{col_x: 'Outros', col_y: soma_outros}])
                
                dados_agrupados = pd.concat([top_itens, linha_outros], ignore_index=True)

            # 3. Plotagem do gráfico escolhido
            cor_padrao = "#2196F3"
            
            if tipo == "Barras":
                self.ax.bar(dados_agrupados[col_x].astype(str), dados_agrupados[col_y], color=cor_padrao, edgecolor="#1976D2")
            
            elif tipo == "Linhas":
                self.ax.plot(
                    dados_agrupados[col_x].astype(str), dados_agrupados[col_y], marker="o", color="#E91E63", linewidth=2.5
                )
            
            elif tipo == "Área":
                self.ax.fill_between(dados_agrupados[col_x].astype(str), dados_agrupados[col_y], color="#00BCD4", alpha=0.3)
                self.ax.plot(dados_agrupados[col_x].astype(str), dados_agrupados[col_y], color="#00BCD4", linewidth=2)
            
            elif tipo == "Pizza":
                wedges, texts, autotexts = self.ax.pie(
                    dados_agrupados[col_y], 
                    labels=dados_agrupados[col_x].astype(str), 
                    autopct="%1.1f%%", 
                    startangle=90
                )
                plt.setp(autotexts, size=8, weight="bold")
                plt.setp(texts, size=9)

            # 4. Ajustes de títulos e rotações do eixo X
            if tipo != "Pizza":
                self.ax.set_title(f"{label_y} por {col_x}\n(Top {LIMITE_ITENS} + Outros)", fontsize=12, fontweight="bold", pad=15)
                self.ax.set_xlabel(col_x, fontsize=10, labelpad=10)
                self.ax.set_ylabel(label_y, fontsize=10, labelpad=10)
                
                self.ax.set_xticklabels(dados_agrupados[col_x].astype(str), rotation=35, ha='right', fontsize=9)
                self.ax.grid(axis='y', linestyle='--', alpha=0.5)
            else:
                self.ax.set_title(f"Distribuição de {label_y} por {col_x}", fontsize=12, fontweight="bold", pad=15)

            # 5. AJUSTE DE MARGEM MANUAL REFORÇADO (Substitui completamente o tight_layout)
            if tipo != "Pizza":
                self.fig.subplots_adjust(bottom=0.28, left=0.15, right=0.95, top=0.86)
            else:
                self.fig.subplots_adjust(bottom=0.12, left=0.12, right=0.95, top=0.86)

            self.canvas.draw()

        except Exception as e:
            messagebox.showerror("Erro ao gerar gráfico", f"Ocorreu um erro ao processar os dados:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppBIExcel(root)
    root.mainloop()