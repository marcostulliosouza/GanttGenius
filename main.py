import tkinter as tk
from datetime import datetime, timedelta
from tkinter import ttk, filedialog, messagebox  # Import messagebox module
from tkcalendar import DateEntry
import pandas as pd
import configparser
"""
V2.1
Adicionado o cliente LOCALIZA
"""
"""
V3.5

Adicionado o config.ini para deixar a opcao de clientes editavel
"""


config = configparser.ConfigParser()
config.read('config.ini')

# # Imprime todas as seções e suas chaves
# for section in config.sections():
#     print(f"Seção: {section}")
#     for key in config[section]:
#         print(f"  {key} = {config[section][key.upper()]}")
#
# # Tenta acessar a seção 'Substituicoes'
# if 'Substituicoes' in config:
#     substituicoes = dict(config['Substituicoes'])
#     print(substituicoes)
# else:
#     print("A seção 'Substituicoes' não foi encontrada.")

__version__ = "3.5"
class AplicacaoDesktop:
    def __init__(self, root):
        self.root = root
        self.root.title(f"GanttGenius - Conversor de Tabela para BetterCall - V{__version__} - 13/08/24")

        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(self.frame, text="Tabela 1:").grid(row=0, column=0, sticky=tk.W)
        self.tabela1_entry = ttk.Entry(self.frame, width=30)
        self.tabela1_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar a Programação de Recursos", command=self.carregar_tabela1).grid(row=0, column=2, sticky=tk.W)

        ttk.Label(self.frame, text="Tabela 2:").grid(row=1, column=0, sticky=tk.W)
        self.tabela2_entry = ttk.Entry(self.frame, width=30)
        self.tabela2_entry.grid(row=1, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar a Programação por Hora_V2", command=self.carregar_tabela2).grid(row=1, column=2, sticky=tk.W)

        ttk.Label(self.frame, text="Data de produção:").grid(row=2, column=0, sticky=tk.W)
        self.data_entry = DateEntry(self.frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.data_entry.grid(row=2, column=1, sticky=(tk.W, tk.E))

        ttk.Button(self.frame, text="Gerar Gantt BetterCall", command=self.analisar_gerar_tabela).grid(row=3, column=0, columnspan=3, pady=10)

    def carregar_tabela(self, entry_widget, title):
        filepath = filedialog.askopenfilename(title=f"Selecionar {title}",
                                              filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

    def carregar_tabela1(self):
        self.carregar_tabela(self.tabela1_entry, "Tabela 1")

    def carregar_tabela2(self):
        self.carregar_tabela(self.tabela2_entry, "Tabela 2")

    def analisar_gerar_tabela(self):
        tabela1_path = self.tabela1_entry.get()
        tabela2_path = self.tabela2_entry.get()
        data_comparacao = self.data_entry.get() + ' 00:00:00'
        tabela1 = pd.read_excel(tabela1_path)
        tabela1_filtrada = tabela1[(tabela1.iloc[:, 27].astype(str).str.contains("TESTE"))]

        # Converter 'Unnamed: 15' para datetime
        tabela1_filtrada['Unnamed: 15'] = pd.to_datetime(tabela1_filtrada['Unnamed: 15'], errors='coerce', dayfirst=True)

        tabela1_filtrada_fil = tabela1_filtrada[tabela1_filtrada['Unnamed: 15'].dt.date == pd.to_datetime(data_comparacao).date()]
        tabela2 = pd.read_excel(tabela2_path)

        for index, cell_value in tabela2.iloc[:, 1].items():
            if 0 <= index < len(tabela2) and tabela2.iloc[index, 1] in tabela2.iloc[:, 1][tabela2.iloc[:, 1].duplicated(keep=False) & tabela2.iloc[:, 1].notnull()]:
                mask = tabela2.iloc[:, 1] == tabela2.iloc[index, 1]
                tabela2.loc[mask, 'Unnamed: 1'] = cell_value
        tabela2['Unnamed: 1'] = tabela2['Unnamed: 1'].ffill()
        mescladas = tabela2['Unnamed: 1'].duplicated(keep=False) & tabela2['Unnamed: 1'].notnull()
        tabela2['Unnamed: 1'] = tabela2['Unnamed: 1'].mask(mescladas, tabela2['Unnamed: 1'].ffill())

        for index, cell_value in tabela2.iloc[:, 2].items():
            if 0 <= index < len(tabela2) and tabela2.iloc[index, 2] in tabela2.iloc[:, 2][tabela2.iloc[:, 2].duplicated(keep=False) & tabela2.iloc[:, 2].notnull()]:
                mask = tabela2.iloc[:, 2] == tabela2.iloc[index, 2]
                tabela2.loc[mask, 'Unnamed: 2'] = cell_value
        tabela2['Unnamed: 2'] = tabela2['Unnamed: 2'].ffill()
        mescladas = tabela2['Unnamed: 2'].duplicated(keep=False) & tabela2['Unnamed: 2'].notnull()
        tabela2['Unnamed: 2'] = tabela2['Unnamed: 2'].mask(mescladas, tabela2['Unnamed: 2'].ffill())

        for index, cell_value in tabela2.iloc[:, 3].items():
            if 0 <= index < len(tabela2) and tabela2.iloc[index, 3] in tabela2.iloc[:, 3][tabela2.iloc[:, 3].duplicated(keep=False) & tabela2.iloc[:, 3].notnull()]:
                mask = tabela2.iloc[:, 3] == tabela2.iloc[index, 3]
                tabela2.loc[mask, 'Unnamed: 3'] = cell_value
        tabela2['Unnamed: 3'] = tabela2['Unnamed: 3'].ffill()
        mescladas = tabela2['Unnamed: 2'].duplicated(keep=False) & tabela2['Unnamed: 3'].notnull()
        tabela2['Unnamed: 3'] = tabela2['Unnamed: 3'].mask(mescladas, tabela2['Unnamed: 3'].ffill())

        for index, cell_value in tabela2.iloc[:, 6].items():
            if 0 <= index < len(tabela2) and tabela2.iloc[index, 6] in tabela2.iloc[:, 6][tabela2.iloc[:, 6].duplicated(keep=False) & tabela2.iloc[:, 6].notnull()]:
                mask = tabela2.iloc[:, 6] == tabela2.iloc[index, 6]
                tabela2.loc[mask, 'Unnamed: 6'] = cell_value
        tabela2['Unnamed: 6'] = tabela2['Unnamed: 6'].ffill()
        mescladas = tabela2['Unnamed: 6'].duplicated(keep=False) & tabela2['Unnamed: 6'].notnull()
        tabela2['Unnamed: 6'] = tabela2['Unnamed: 6'].mask(mescladas, tabela2['Unnamed: 6'].ffill())

        valores_filtro = ["HP", "TESTE FUNCIONAL",
                          "PÓS COMPOSIÇÃO",
                          "TRI",
                          "GRAVAÇÃO DO CI PTH",
                          "GRAVAÇÃO DO CI SMT",
                          "MONTAGEM MECÂNICA"]
        tabela2_filtrada = tabela2[tabela2.iloc[:, 6].isin(valores_filtro)]
        tabela3 = pd.DataFrame()
        for index, row_tabela2 in tabela2_filtrada.iterrows():
            for index, row_tabela1 in tabela1_filtrada.iterrows():
                condicao = (row_tabela2['Unnamed: 2'] == row_tabela1['Unnamed: 12']) and (row_tabela2['Unnamed: 6'] == row_tabela1['Unnamed: 14'])
                if condicao:
                    tabela3 = pd.concat([tabela3, pd.DataFrame(row_tabela2).T], ignore_index=True)

        # Verificar se há algum valor string na coluna 32 e excluir as linhas sem valor
        tabela3 = tabela3[tabela3['Unnamed: 34'].notnull()]

        for index_tabela3, row_tabela3 in tabela3.iterrows():
            for index_tabela1, row_tabela1 in tabela1_filtrada_fil.iterrows():
                condicao_verificacao = (row_tabela3['Unnamed: 2'] == row_tabela1['Unnamed: 12']) and (
                        row_tabela3['Unnamed: 6'] == row_tabela1['Unnamed: 14']) and (row_tabela3['Unnamed: 7'] == row_tabela1['Unnamed: 2'])
                if condicao_verificacao:
                    if ((row_tabela1['Unnamed: 15'].hour == 5) &
                            (row_tabela1['Unnamed: 15'].minute >= 30) &
                            (row_tabela1['Unnamed: 15'].minute <= 59)):
                        tabela3.at[index_tabela3, 'Unnamed: 13'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 6) &
                          (row_tabela1['Unnamed: 15'].hour < 7)):
                        tabela3.at[index_tabela3, 'Unnamed: 14'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 7) &
                          (row_tabela1['Unnamed: 15'].hour < 8)):
                        tabela3.at[index_tabela3, 'Unnamed: 15'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 8) &
                          (row_tabela1['Unnamed: 15'].hour < 9)):
                        tabela3.at[index_tabela3, 'Unnamed: 16'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 9) &
                          (row_tabela1['Unnamed: 15'].hour < 10)):
                        tabela3.at[index_tabela3, 'Unnamed: 17'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 10) &
                          (row_tabela1['Unnamed: 15'].hour < 11)):
                        tabela3.at[index_tabela3, 'Unnamed: 18'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 11) &
                          (row_tabela1['Unnamed: 15'].hour < 12)):
                        tabela3.at[index_tabela3, 'Unnamed: 19'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 12) &
                          (row_tabela1['Unnamed: 15'].hour < 13)):
                        tabela3.at[index_tabela3, 'Unnamed: 20'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 13) &
                          (row_tabela1['Unnamed: 15'].hour < 14)):
                        tabela3.at[index_tabela3, 'Unnamed: 21'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 14) &
                          (row_tabela1['Unnamed: 15'].hour < 15)):
                        tabela3.at[index_tabela3, 'Unnamed: 22'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour == 15) &
                          (row_tabela1['Unnamed: 15'].minute >= 0) &
                          (row_tabela1['Unnamed: 15'].minute < 18)):
                        tabela3.at[index_tabela3, 'Unnamed: 23'] = 1

                    elif ((row_tabela1['Unnamed: 15'].hour == 15) &
                          (row_tabela1['Unnamed: 15'].minute >= 18) &
                          (row_tabela1['Unnamed: 15'].minute <= 59)):
                        tabela3.at[index_tabela3, 'Unnamed: 24'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 16) &
                          (row_tabela1['Unnamed: 15'].hour < 17)):
                        tabela3.at[index_tabela3, 'Unnamed: 25'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 17) &
                          (row_tabela1['Unnamed: 15'].hour < 18)):
                        tabela3.at[index_tabela3, 'Unnamed: 26'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 18) &
                          (row_tabela1['Unnamed: 15'].hour < 19)):
                        tabela3.at[index_tabela3, 'Unnamed: 27'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 19) &
                          (row_tabela1['Unnamed: 15'].hour < 20)):
                        tabela3.at[index_tabela3, 'Unnamed: 28'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 20) &
                          (row_tabela1['Unnamed: 15'].hour < 21)):
                        tabela3.at[index_tabela3, 'Unnamed: 29'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 21) &
                          (row_tabela1['Unnamed: 15'].hour < 22)):
                        tabela3.at[index_tabela3, 'Unnamed: 30'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 22) &
                          (row_tabela1['Unnamed: 15'].hour < 23)):
                        tabela3.at[index_tabela3, 'Unnamed: 31'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour >= 23) &
                          (row_tabela1['Unnamed: 15'].hour < 24)):
                        tabela3.at[index_tabela3, 'Unnamed: 32'] = 1
                    elif ((row_tabela1['Unnamed: 15'].hour == 0) &
                          (row_tabela1['Unnamed: 15'].minute >= 0) &
                          (row_tabela1['Unnamed: 15'].minute <= 38)):
                        tabela3.at[index_tabela3, 'Unnamed: 33'] = 1
        if (str(tabela2.iloc[5,13]))==data_comparacao:
            # tabela2 = tabela2[tabela2['Unnamed: 10'].notnull()]
            nova_tabela = pd.DataFrame(
                columns=['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5',
                         'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11',
                         'Unnamed: 12', 'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16', 'Unnamed: 17',
                         'Unnamed: 18', 'Unnamed: 19', 'Unnamed: 20', 'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23',
                         'Unnamed: 24', 'Unnamed: 25', 'Unnamed: 26', 'Unnamed: 27', 'Unnamed: 28', 'Unnamed: 29',
                         'Unnamed: 30', 'Unnamed: 31', 'Unnamed: 32', 'Unnamed: 33', 'Unnamed: 34', 'Unnamed: 35',
                         'Unnamed: 36', 'Unnamed: 37', 'Unnamed: 38', 'Unnamed: 39', 'Unnamed: 40', 'Unnamed: 41',
                         'Unnamed: 42'])
            tabela3 = tabela3.iloc[:].drop_duplicates(keep='first')
            for i in range(31):
                nova_tabela.loc[i] = 0
            nova_tabela = nova_tabela.fillna(0)
            for i in range(31, len(tabela3) + 31):
                nova_tabela.at[i, 'Unnamed: 0'] = str(tabela3.iloc[i - 31, 1])  # cliente
                nova_tabela.at[i, 'Unnamed: 1'] = str(tabela3.iloc[i - 31, 2])  # produto
                nova_tabela.at[i, 'Unnamed: 2'] = str(tabela3.iloc[i - 31, 3])  # descrição
                nova_tabela.at[i, 'Unnamed: 3'] = str(tabela3.iloc[i - 31, 6])  # etapa
                nova_tabela.at[i, 'Unnamed: 4'] = str(tabela3.iloc[i - 31, 10])  # rate
                nova_tabela.at[i, 'Unnamed: 5'] = 0
                nova_tabela.at[i, 'Unnamed: 6'] = 0
                nova_tabela.at[i, 'Unnamed: 7'] = 0
                nova_tabela.at[i, 'Unnamed: 8'] = str(tabela3.iloc[i - 31, 9])  # celula / pós
                nova_tabela.at[i, 'Unnamed: 9'] = 0
                nova_tabela.at[i, 'Unnamed: 10'] = str('COM TESTE')  # haste
                nova_tabela.at[i, 'Unnamed: 11'] = 0
                nova_tabela.at[i, 'Unnamed: 12'] = 0
                nova_tabela.at[i, 'Unnamed: 13'] = 0  # T3BL
                nova_tabela.at[i, 'Unnamed: 14'] = 0
                nova_tabela.at[i, 'Unnamed: 15'] = 0  # T3BL
                nova_tabela.at[i, 'Unnamed: 16'] = 0  # T3AL
                nova_tabela.at[i, 'Unnamed: 17'] = 0
                nova_tabela.at[i, 'Unnamed: 18'] = 0
                nova_tabela.at[i, 'Unnamed: 19'] = 0  # T3AL
                nova_tabela.at[i, 'Unnamed: 20'] = tabela3.iloc[i - 31, 13]  # T1BL
                nova_tabela.at[i, 'Unnamed: 21'] = tabela3.iloc[i - 31, 14]
                nova_tabela.at[i, 'Unnamed: 22'] = tabela3.iloc[i - 31, 15]
                nova_tabela.at[i, 'Unnamed: 23'] = tabela3.iloc[i - 31, 16]
                nova_tabela.at[i, 'Unnamed: 24'] = tabela3.iloc[i - 31, 17]
                nova_tabela.at[i, 'Unnamed: 25'] = tabela3.iloc[i - 31, 18]  # T1BL
                nova_tabela.at[i, 'Unnamed: 26'] = tabela3.iloc[i - 31, 19]  # T1AL
                nova_tabela.at[i, 'Unnamed: 27'] = tabela3.iloc[i - 31, 20]
                nova_tabela.at[i, 'Unnamed: 28'] = tabela3.iloc[i - 31, 21]
                nova_tabela.at[i, 'Unnamed: 29'] = tabela3.iloc[i - 31, 22] + tabela3.iloc[i - 31, 23]
                nova_tabela.at[i, 'Unnamed: 30'] = 0
                nova_tabela.at[i, 'Unnamed: 31'] = tabela3.iloc[i - 31, 24]  # T2BL
                nova_tabela.at[i, 'Unnamed: 32'] = tabela3.iloc[i - 31, 25]
                nova_tabela.at[i, 'Unnamed: 33'] = tabela3.iloc[i - 31, 26]
                nova_tabela.at[i, 'Unnamed: 34'] = tabela3.iloc[i - 31, 27]
                nova_tabela.at[i, 'Unnamed: 35'] = tabela3.iloc[i - 31, 28]  # T2BL
                nova_tabela.at[i, 'Unnamed: 36'] = tabela3.iloc[i - 31, 29]  # T2AL
                nova_tabela.at[i, 'Unnamed: 37'] = tabela3.iloc[i - 31, 30]
                nova_tabela.at[i, 'Unnamed: 38'] = tabela3.iloc[i - 31, 31]
                nova_tabela.at[i, 'Unnamed: 39'] = tabela3.iloc[i - 31, 32]
                nova_tabela.at[i, 'Unnamed: 40'] = tabela3.iloc[i - 31, 33]
                nova_tabela.at[i, 'Unnamed: 41'] = 0
                nova_tabela.at[i, 'Unnamed: 42'] = tabela3.iloc[i - 31, 34]  # totalprodução
        else:
            messagebox.showerror("Erro", "Arquivo não possui plano de produção para a data solicitada.")
            #nova_tabela = None
            return
        substituicoes = {k.upper(): v for k, v in config['Substituicoes'].items()}
        nova_tabela['Unnamed: 0'] = nova_tabela['Unnamed: 0'].map(substituicoes)
        for i in range(len(nova_tabela)):
               if "HP" in str(nova_tabela.at[i, 'Unnamed: 3']) or "TRI" in str(nova_tabela.at[i, 'Unnamed: 3']):
                    nova_tabela.at[i, 'Unnamed: 1'] = str(nova_tabela.at[i, 'Unnamed: 1']) + ' - ' + str(
                    nova_tabela.at[i, 'Unnamed: 3'])
        nova_tabela.iloc[7, 0] = datetime.strptime(data_comparacao, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")
        if nova_tabela is not None:
            nova_tabela_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                             filetypes=[("Arquivos Excel", "*.xlsx")])
            if nova_tabela_path:
                try:
                    nova_tabela.to_excel(nova_tabela_path, index=False)
                    messagebox.showinfo("Sucesso", f"Plano para o BetterCall gerado com sucesso: {nova_tabela_path}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar a nova tabela: {str(e)}")
        else:
            messagebox.showinfo("Aviso", "Não foi possível gerar a tabela. A condição não foi atendida.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacaoDesktop(root)
    root.mainloop()