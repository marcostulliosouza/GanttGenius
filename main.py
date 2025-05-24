import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import configparser

__version__ = "3.5.2"  # Versão com correção para coluna 7

config = configparser.ConfigParser()
config.read('config.ini')


class AplicacaoDesktop:
    def __init__(self, root):
        self.root = root
        self.root.title(f"GanttGenius - Conversor de Tabela para BetterCall - V{__version__}")

        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Interface (mantida igual)
        ttk.Label(self.frame, text="Tabela 1:").grid(row=0, column=0, sticky=tk.W)
        self.tabela1_entry = ttk.Entry(self.frame, width=30)
        self.tabela1_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar a Programação de Recursos", command=self.carregar_tabela1).grid(row=0,
                                                                                                              column=2,
                                                                                                              sticky=tk.W)

        ttk.Label(self.frame, text="Tabela 2:").grid(row=1, column=0, sticky=tk.W)
        self.tabela2_entry = ttk.Entry(self.frame, width=30)
        self.tabela2_entry.grid(row=1, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar a Programação por Hora_V2", command=self.carregar_tabela2).grid(row=1,
                                                                                                              column=2,
                                                                                                              sticky=tk.W)

        ttk.Label(self.frame, text="Data de produção:").grid(row=2, column=0, sticky=tk.W)
        self.data_entry = DateEntry(self.frame, width=12, background='darkblue', foreground='white', borderwidth=2,
                                    date_pattern='yyyy-mm-dd')
        self.data_entry.grid(row=2, column=1, sticky=(tk.W, tk.E))

        ttk.Button(self.frame, text="Gerar Gantt BetterCall", command=self.analisar_gerar_tabela).grid(row=3, column=0,
                                                                                                       columnspan=3,
                                                                                                       pady=10)

    def carregar_tabela(self, entry_widget, title):
        filepath = filedialog.askopenfilename(title=f"Selecionar {title}",
                                              filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

    def carregar_tabela1(self):
        self.carregar_tabela(self.tabela1_entry, "Tabela 1")

    def carregar_tabela2(self):
        self.carregar_tabela(self.tabela2_entry, "Tabela 2")

    def get_column_name(self, df, position):
        """Obtém o nome da coluna pela posição, tratando o caso de colunas não nomeadas ou com nomes numéricos"""
        if position < len(df.columns):
            col_name = df.columns[position]
            # Se for um número, assume que é uma coluna "Unnamed"
            if isinstance(col_name, (int, float)):
                return f'Unnamed: {position}'
            return col_name
        return None

    def analisar_gerar_tabela(self):
        try:
            tabela1_path = self.tabela1_entry.get()
            tabela2_path = self.tabela2_entry.get()
            data_comparacao = self.data_entry.get() + ' 00:00:00'

            if not tabela1_path or not tabela2_path:
                messagebox.showerror("Erro", "Selecione ambos os arquivos!")
                return

            # Carrega as tabelas
            tabela1 = pd.read_excel(tabela1_path)
            tabela2 = pd.read_excel(tabela2_path)

            # Verifica coluna 27 (TESTE) na tabela1
            col_test = self.get_column_name(tabela1, 27)
            if col_test is None or col_test not in tabela1.columns:
                messagebox.showerror("Erro", "A Tabela 1 não possui a coluna de teste necessária (posição 27)!")
                return

            tabela1_filtrada = tabela1[(tabela1[col_test].astype(str).str.contains("TESTE"))].copy()

            # Verifica coluna 15 (data) na tabela1
            col_data = self.get_column_name(tabela1, 15)
            if col_data is None or col_data not in tabela1_filtrada.columns:
                messagebox.showerror("Erro", "A Tabela 1 não possui a coluna de data necessária (posição 15)!")
                return

            tabela1_filtrada[col_data] = pd.to_datetime(tabela1_filtrada[col_data], errors='coerce', dayfirst=True)
            tabela1_filtrada_fil = tabela1_filtrada[
                tabela1_filtrada[col_data].dt.date == pd.to_datetime(data_comparacao).date()]

            # Renomeia colunas numéricas para o padrão "Unnamed: X" na tabela2
            tabela2.columns = [self.get_column_name(tabela2, i) or f'Unnamed: {i}' for i in range(len(tabela2.columns))]

            # Preenche valores mesclados na tabela2
            for col in [1, 2, 3, 6]:
                col_name = f'Unnamed: {col}'
                if col_name in tabela2.columns:
                    tabela2[col_name] = tabela2[col_name].ffill()

            # Filtra tabela2
            valores_filtro = ["HP", "TESTE FUNCIONAL", "PÓS COMPOSIÇÃO", "TRI",
                              "GRAVAÇÃO DO CI PTH", "GRAVAÇÃO DO CI SMT", "MONTAGEM MECÂNICA"]

            if 'Unnamed: 6' not in tabela2.columns:
                messagebox.showerror("Erro", "A Tabela 2 não possui a coluna 'Unnamed: 6' necessária!")
                return

            tabela2_filtrada = tabela2[tabela2['Unnamed: 6'].isin(valores_filtro)]
            tabela3 = pd.DataFrame()

            # Verifica se a coluna 7 existe na tabela2
            col7_tabela2 = 'Unnamed: 7'
            col7_exists = col7_tabela2 in tabela2_filtrada.columns

            # Verifica se a coluna 2 existe na tabela1
            col2_tabela1 = self.get_column_name(tabela1, 2)
            if col2_tabela1 is None or col2_tabela1 not in tabela1_filtrada.columns:
                messagebox.showerror("Erro", "A Tabela 1 não possui a coluna 2 necessária!")
                return

            for _, row_tabela2 in tabela2_filtrada.iterrows():
                for _, row_tabela1 in tabela1_filtrada.iterrows():
                    condicao = (row_tabela2['Unnamed: 2'] == row_tabela1['Unnamed: 12']) and \
                               (row_tabela2['Unnamed: 6'] == row_tabela1['Unnamed: 14'])

                    # Só adiciona a condição da coluna 7 se ela existir em ambos os dataframes
                    if col7_exists and 'Unnamed: 2' in row_tabela1:
                        condicao = condicao and (row_tabela2[col7_tabela2] == row_tabela1['Unnamed: 2'])

                    if condicao:
                        tabela3 = pd.concat([tabela3, pd.DataFrame(row_tabela2).T], ignore_index=True)

            # Restante do código mantido igual...
            if 'Unnamed: 34' not in tabela3.columns:
                messagebox.showerror("Erro", "A tabela gerada não possui a coluna 'Unnamed: 34' necessária!")
                return

            tabela3 = tabela3[tabela3['Unnamed: 34'].notnull()]

            # Preenche os horários (código original)
            for index_tabela3, row_tabela3 in tabela3.iterrows():
                for index_tabela1, row_tabela1 in tabela1_filtrada_fil.iterrows():
                    condicao = (row_tabela3['Unnamed: 2'] == row_tabela1['Unnamed: 12']) and \
                               (row_tabela3['Unnamed: 6'] == row_tabela1['Unnamed: 14'])

                    if col7_exists and 'Unnamed: 2' in row_tabela1:
                        condicao = condicao and (row_tabela3[col7_tabela2] == row_tabela1['Unnamed: 2'])

                    if condicao and 'Unnamed: 15' in row_tabela1:
                        hora = row_tabela1['Unnamed: 15']
                        if pd.notna(hora):
                            # Lógica original de preenchimento dos horários
                            if ((hora.hour == 5) & (hora.minute >= 30) & (hora.minute <= 59)):
                                tabela3.at[index_tabela3, 'Unnamed: 13'] = 1
                            elif ((hora.hour >= 6) & (hora.hour < 7)):
                                tabela3.at[index_tabela3, 'Unnamed: 14'] = 1
                            elif ((hora.hour >= 7) & (hora.hour < 8)):
                                tabela3.at[index_tabela3, 'Unnamed: 15'] = 1
                            elif ((hora.hour >= 8) & (hora.hour < 9)):
                                tabela3.at[index_tabela3, 'Unnamed: 16'] = 1
                            elif ((hora.hour >= 9) & (hora.hour < 10)):
                                tabela3.at[index_tabela3, 'Unnamed: 17'] = 1
                            elif ((hora.hour >= 10) & (hora.hour < 11)):
                                tabela3.at[index_tabela3, 'Unnamed: 18'] = 1
                            elif ((hora.hour >= 11) & (hora.hour < 12)):
                                tabela3.at[index_tabela3, 'Unnamed: 19'] = 1
                            elif ((hora.hour >= 12) & (hora.hour < 13)):
                                tabela3.at[index_tabela3, 'Unnamed: 20'] = 1
                            elif ((hora.hour >= 13) & (hora.hour < 14)):
                                tabela3.at[index_tabela3, 'Unnamed: 21'] = 1
                            elif ((hora.hour >= 14) & (hora.hour < 15)):
                                tabela3.at[index_tabela3, 'Unnamed: 22'] = 1
                            elif ((hora.hour == 15) & (hora.minute >= 0) & (hora.minute < 18)):
                                tabela3.at[index_tabela3, 'Unnamed: 23'] = 1
                            elif ((hora.hour == 15) & (hora.minute >= 18)):
                                tabela3.at[index_tabela3, 'Unnamed: 24'] = 1
                            elif ((hora.hour >= 16) & (hora.hour < 17)):
                                tabela3.at[index_tabela3, 'Unnamed: 25'] = 1
                            elif ((hora.hour >= 17) & (hora.hour < 18)):
                                tabela3.at[index_tabela3, 'Unnamed: 26'] = 1
                            elif ((hora.hour >= 18) & (hora.hour < 19)):
                                tabela3.at[index_tabela3, 'Unnamed: 27'] = 1
                            elif ((hora.hour >= 19) & (hora.hour < 20)):
                                tabela3.at[index_tabela3, 'Unnamed: 28'] = 1
                            elif ((hora.hour >= 20) & (hora.hour < 21)):
                                tabela3.at[index_tabela3, 'Unnamed: 29'] = 1
                            elif ((hora.hour >= 21) & (hora.hour < 22)):
                                tabela3.at[index_tabela3, 'Unnamed: 30'] = 1
                            elif ((hora.hour >= 22) & (hora.hour < 23)):
                                tabela3.at[index_tabela3, 'Unnamed: 31'] = 1
                            elif ((hora.hour >= 23) & (hora.hour < 24)):
                                tabela3.at[index_tabela3, 'Unnamed: 32'] = 1
                            elif ((hora.hour == 0) & (hora.minute <= 38)):
                                tabela3.at[index_tabela3, 'Unnamed: 33'] = 1

            # Verifica a data de produção
            if len(tabela2.columns) > 13 and (str(tabela2.iloc[5, 13]) == data_comparacao):
                nova_tabela = pd.DataFrame(
                    columns=[f'Unnamed: {i}' for i in range(43)])

                tabela3 = tabela3.iloc[:].drop_duplicates(keep='first')

                # Preenche as primeiras 31 linhas com 0
                for i in range(31):
                    nova_tabela.loc[i] = 0

                nova_tabela = nova_tabela.fillna(0)

                # Preenche os dados
                for i in range(31, len(tabela3) + 31):
                    if i - 31 >= len(tabela3):
                        continue

                    # Mapeamento seguro das colunas
                    col_mapping = {
                        0: (1, str),  # cliente
                        1: (2, str),  # produto
                        2: (3, str),  # descrição
                        3: (6, str),  # etapa
                        4: (10, str),  # rate
                        8: (9, str),  # celula / pós
                        10: ('COM TESTE', str),  # haste
                        20: (13, int),  # T1BL
                        21: (14, int),
                        22: (15, int),
                        23: (16, int),
                        24: (17, int),
                        25: (18, int),
                        26: (19, int),
                        27: (20, int),
                        28: (21, int),
                        29: (22, int),
                        31: (24, int),
                        32: (25, int),
                        33: (26, int),
                        34: (27, int),
                        35: (28, int),
                        36: (29, int),
                        37: (30, int),
                        38: (31, int),
                        39: (32, int),
                        40: (33, int),
                        42: (34, int)
                    }

                    for col, (src_col, conv_func) in col_mapping.items():
                        col_name = f'Unnamed: {col}'
                        if isinstance(src_col, int):
                            if src_col < len(tabela3.columns):
                                try:
                                    value = tabela3.iloc[i - 31, src_col]
                                    nova_tabela.at[i, col_name] = conv_func(value) if pd.notna(
                                        value) else 0 if conv_func == int else ''
                                except:
                                    nova_tabela.at[i, col_name] = 0 if conv_func == int else ''
                        else:
                            nova_tabela.at[i, col_name] = src_col

                    # Soma das colunas 22 e 23 para a coluna 29
                    if 22 < len(tabela3.columns) and 23 < len(tabela3.columns):
                        try:
                            soma = tabela3.iloc[i - 31, 22] + tabela3.iloc[i - 31, 23]
                            nova_tabela.at[i, 'Unnamed: 29'] = int(soma) if pd.notna(soma) else 0
                        except:
                            nova_tabela.at[i, 'Unnamed: 29'] = 0

                # Aplica substituições de clientes
                if 'Substituicoes' in config:
                    substituicoes = {k.upper(): v for k, v in config['Substituicoes'].items()}
                    nova_tabela['Unnamed: 0'] = nova_tabela['Unnamed: 0'].str.upper().map(substituicoes).fillna(
                        nova_tabela['Unnamed: 0'])

                # Adiciona descrição para HP e TRI
                for i in range(len(nova_tabela)):
                    if pd.notna(nova_tabela.at[i, 'Unnamed: 3']) and (
                            "HP" in str(nova_tabela.at[i, 'Unnamed: 3']) or "TRI" in str(
                            nova_tabela.at[i, 'Unnamed: 3'])):
                        nova_tabela.at[
                            i, 'Unnamed: 1'] = f"{nova_tabela.at[i, 'Unnamed: 1']} - {nova_tabela.at[i, 'Unnamed: 3']}"

                # Adiciona a data formatada
                nova_tabela.iloc[7, 0] = datetime.strptime(data_comparacao, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")

                # Salva o arquivo
                nova_tabela_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Arquivos Excel", "*.xlsx")])

                if nova_tabela_path:
                    try:
                        nova_tabela.to_excel(nova_tabela_path, index=False)
                        messagebox.showinfo("Sucesso",
                                            f"Plano para o BetterCall gerado com sucesso: {nova_tabela_path}")
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao salvar a nova tabela: {str(e)}")
            else:
                messagebox.showerror("Erro", "Arquivo não possui plano de produção para a data solicitada.")
                return

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacaoDesktop(root)
    root.mainloop()