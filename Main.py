import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk


class TopProductsApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Análises de Produtos")
        self.master.geometry("800x600")
        self.master.configure(bg="#283042")

        button_params = {
            "bg": "#6A7B8D",
            "fg": "white",
            "bd": 0,
            "padx": 10,
            "pady": 5,
            "relief": tk.FLAT,
            "font": ('Arial', 10)
        }

        self.btn_open_file = tk.Button(
            self.master, text="Abrir Arquivo", command=self.open_file, **button_params)
        self.btn_open_file.grid(row=0, column=0, padx=(
            20, 10), pady=(20, 5), sticky="nsew")

        padx_between_buttons = 10

        self.btn_most_profitable = tk.Button(
            self.master, text="Produtos Mais Lucrativos", command=self.show_most_profitable, **button_params)
        self.btn_most_profitable.grid(
            row=0, column=1, pady=(20, 5), sticky="nsew", padx=(padx_between_buttons, 0))

        self.btn_most_sold = tk.Button(
            self.master, text="Produtos Mais Vendidos", command=self.show_most_sold, **button_params)
        self.btn_most_sold.grid(
            row=0, column=2, pady=(20, 5), sticky="nsew", padx=(padx_between_buttons, 0))

        self.btn_least_sold = tk.Button(
            self.master, text="Produtos Menos Vendidos", command=self.show_least_sold, **button_params)
        self.btn_least_sold.grid(
            row=0, column=3, pady=(20, 5), sticky="nsew", padx=(padx_between_buttons, 0))

        self.btn_bar_chart = tk.Button(
            self.master, text="Gráfico de Barras", command=self.show_bar_chart, **button_params)
        self.btn_bar_chart.grid(
            row=1, column=0, columnspan=4, pady=(20, 5), sticky="nsew")

        self.tree_all = ttk.Treeview(self.master)
        self.tree_all["columns"] = (
            "DESCRICAO DO PRODUTO", "QTDE VENDIDA PDV", "VALOR VENDIDO PDV", "VALOR_MEDIO_POR_UNIDADE")
        self.configure_tree(self.tree_all)
        self.tree_all.grid(row=2, column=0, columnspan=4,
                           padx=20, pady=10, sticky="nsew")

        button_frame = tk.Frame(self.master, bg="#283042")
        button_frame.grid(row=3, column=0, columnspan=4,
                          pady=(0, 10), sticky="nsew")
        button_frame.grid_rowconfigure(0, weight=1)
        button_frame.grid_columnconfigure(0, weight=1)

        self.master.grid_rowconfigure(2, weight=1)
        self.master.grid_rowconfigure(3, weight=1)
        self.master.grid_rowconfigure(4, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)

    def configure_tree(self, tree):
        tree.column("#0", width=0, stretch=tk.NO)
        tree.column("DESCRICAO DO PRODUTO", anchor=tk.W, width=200)
        tree.column("QTDE VENDIDA PDV", anchor=tk.W, width=100)
        tree.column("VALOR VENDIDO PDV", anchor=tk.W, width=150)
        tree.column("VALOR_MEDIO_POR_UNIDADE", anchor=tk.W, width=150)

        tree.heading("#0", text="", anchor=tk.W)
        tree.heading("DESCRICAO DO PRODUTO",
                     text="Descrição do Produto", anchor=tk.W)
        tree.heading("QTDE VENDIDA PDV",
                     text="Quantidade Vendida PDV", anchor=tk.W)
        tree.heading("VALOR VENDIDO PDV",
                     text="Valor Vendido PDV", anchor=tk.W)
        tree.heading("VALOR_MEDIO_POR_UNIDADE",
                     text="Valor Médio por Unidade", anchor=tk.W)

    def open_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.df.columns = self.df.columns.str.strip()
                self.df["VALOR_MEDIO_POR_UNIDADE"] = self.df["VALOR VENDIDO PDV"] / \
                    self.df["QTDE VENDIDA PDV"]

                self.df = self.df.dropna()

                self.display_result(self.df)

            except Exception as e:
                error_message = "Erro ao processar o arquivo: {}".format(
                    str(e))
                print(error_message)
                self.display_result(error_message)

    def display_result(self, result):
        for item in self.tree_all.get_children():
            self.tree_all.delete(item)

        if isinstance(result, pd.DataFrame):
            for index, row in result.iterrows():
                self.tree_all.insert("", index, values=(
                    row["DESCRICAO DO PRODUTO"], row["QTDE VENDIDA PDV"], row["VALOR VENDIDO PDV"], row["VALOR_MEDIO_POR_UNIDADE"]))
        else:
            self.tree_all.insert("", 0, values=("Erro", result))

    def show_most_profitable(self):
        if hasattr(self, 'df') and not self.df.empty:
            most_profitable = self.get_most_profitable(self.df)
            self.display_result(most_profitable)

    def get_most_profitable(self, df):
        most_profitable = df.sort_values(
            by="VALOR_MEDIO_POR_UNIDADE", ascending=False).head(10)
        return most_profitable[["DESCRICAO DO PRODUTO", "QTDE VENDIDA PDV", "VALOR VENDIDO PDV", "VALOR_MEDIO_POR_UNIDADE"]]

    def show_most_sold(self):
        if hasattr(self, 'df') and not self.df.empty:
            most_sold = self.get_most_sold(self.df)
            self.display_result(most_sold)

    def show_least_sold(self):
        if hasattr(self, 'df') and not self.df.empty:
            least_sold = self.get_least_sold(self.df)
            self.display_result(least_sold)

    def get_most_sold(self, df):
        most_sold = df.sort_values(
            by="QTDE VENDIDA PDV", ascending=False).head(10)
        return most_sold[["DESCRICAO DO PRODUTO", "QTDE VENDIDA PDV", "VALOR VENDIDO PDV", "VALOR_MEDIO_POR_UNIDADE"]]

    def get_least_sold(self, df):
        least_sold = df.sort_values(
            by="QTDE VENDIDA PDV", ascending=True).head(10)
        return least_sold[["DESCRICAO DO PRODUTO", "QTDE VENDIDA PDV", "VALOR VENDIDO PDV", "VALOR_MEDIO_POR_UNIDADE"]]

    def show_bar_chart(self):
        if hasattr(self, 'df') and not self.df.empty:
            most_sold = self.get_most_sold(self.df)
            most_sold = most_sold.dropna(subset=['VALOR VENDIDO PDV'])
            most_sold["DESCRICAO DO PRODUTO"] = most_sold["DESCRICAO DO PRODUTO"].astype(
                str)
            self.display_bar_chart(most_sold)

    def display_bar_chart(self, data):
        fig, ax = plt.subplots()

        bars = ax.bar(data["DESCRICAO DO PRODUTO"],
                      data["VALOR VENDIDO PDV"], color="#6A7B8D")

        ax.set_ylabel("Valor Vendido PDV")
        ax.set_xlabel("Produtos")
        ax.set_title(
            "Relação entre Valor Vendido e Quantidade dos 10 Mais Vendidos")

        ax.set_xticklabels(data["DESCRICAO DO PRODUTO"],
                           rotation=10, ha="right", fontsize=5)

        def autolabel(rects):
            for rect in rects:
                height = rect.get_height()
                ax.annotate('{}'.format(height),
                            xy=(rect.get_x() + rect.get_width() / 2, height),
                            xytext=(0, 3),
                            textcoords="offset points",
                            ha='center', va='bottom')

        autolabel(bars)

        canvas = FigureCanvasTkAgg(fig, master=self.master)
        canvas.get_tk_widget().grid(row=4, column=0, columnspan=4,
                                    padx=20, pady=10, sticky="nsew")
        canvas.draw()

        toolbar = NavigationToolbar2Tk(canvas, self.master)
        toolbar.update()
        toolbar.grid(row=5, column=0, columnspan=4,
                     pady=(0, 10), sticky="nsew")

        canvas.get_tk_widget().configure(scrollregion=canvas.get_tk_widget().bbox("all"))


if __name__ == "__main__":
    root = tk.Tk()
    root.configure(bg="#283042")
    app = TopProductsApp(root)
    root.mainloop()
