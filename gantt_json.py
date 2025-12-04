import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import configparser
import json

__version__ = "4.0.0"  # Nova versão com saída JSON

config = configparser.ConfigParser()
config.read('config.ini')


class AplicacaoDesktop:

    def __init__(self, root):
        self.root = root
        self.root.title(f"GanttGenius - Conversor de Tabela para JSON - V{__version__}")

        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Interface
        ttk.Label(self.frame, text="Tabela 1:").grid(row=0, column=0, sticky=tk.W)
        self.tabela1_entry = ttk.Entry(self.frame, width=30)
        self.tabela1_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar a Programação de Recursos",
                   command=self.carregar_tabela1).grid(row=0, column=2, sticky=tk.W)

        ttk.Label(self.frame, text="Tabela 2:").grid(row=1, column=0, sticky=tk.W)
        self.tabela2_entry = ttk.Entry(self.frame, width=30)
        self.tabela2_entry.grid(row=1, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar a Programação por Hora_V2",
                   command=self.carregar_tabela2).grid(row=1, column=2, sticky=tk.W)

        ttk.Label(self.frame, text="Data de produção:").grid(row=2, column=0, sticky=tk.W)
        self.data_entry = DateEntry(self.frame, width=12, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.data_entry.grid(row=2, column=1, sticky=(tk.W, tk.E))

        ttk.Button(self.frame, text="Gerar JSON BetterCall",
                   command=self.analisar_gerar_json).grid(row=3, column=0, columnspan=3, pady=10)

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
        """Obtém o nome da coluna pela posição"""
        if position < len(df.columns):
            col_name = df.columns[position]
            if isinstance(col_name, (int, float)):
                return f'Unnamed: {position}'
            return col_name
        return None

    def processar_horario(self, hora):
        """Mapeia o horário para o slot correspondente"""
        horarios_map = {
            range(530, 600): "05:30-06:00",
            range(600, 700): "06:00-07:00",
            range(700, 800): "07:00-08:00",
            range(800, 900): "08:00-09:00",
            range(900, 1000): "09:00-10:00",
            range(1000, 1100): "10:00-11:00",
            range(1100, 1200): "11:00-12:00",
            range(1200, 1300): "12:00-13:00",
            range(1300, 1400): "13:00-14:00",
            range(1400, 1500): "14:00-15:00",
            range(1500, 1518): "15:00-15:18",
            range(1518, 1600): "15:18-16:00",
            range(1600, 1700): "16:00-17:00",
            range(1700, 1800): "17:00-18:00",
            range(1800, 1900): "18:00-19:00",
            range(1900, 2000): "19:00-20:00",
            range(2000, 2100): "20:00-21:00",
            range(2100, 2200): "21:00-22:00",
            range(2200, 2300): "22:00-23:00",
            range(2300, 2400): "23:00-00:00",
        }

        hora_minuto = hora.hour * 100 + hora.minute

        # Caso especial para meia-noite
        if hora.hour == 0 and hora.minute <= 38:
            return "00:00-00:38"

        for intervalo, slot in horarios_map.items():
            if hora_minuto in intervalo:
                return slot

        return None

    def analisar_gerar_json(self):
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

            # Processamento da tabela1 (mantido igual)
            col_test = self.get_column_name(tabela1, 27)
            if col_test is None or col_test not in tabela1.columns:
                messagebox.showerror("Erro", "A Tabela 1 não possui a coluna de teste necessária (posição 27)!")
                return

            tabela1_filtrada = tabela1[(tabela1[col_test].astype(str).str.contains("TESTE"))].copy()

            col_data = self.get_column_name(tabela1, 15)
            if col_data is None or col_data not in tabela1_filtrada.columns:
                messagebox.showerror("Erro", "A Tabela 1 não possui a coluna de data necessária (posição 15)!")
                return

            tabela1_filtrada[col_data] = pd.to_datetime(tabela1_filtrada[col_data], errors='coerce', dayfirst=True)
            tabela1_filtrada_fil = tabela1_filtrada[
                tabela1_filtrada[col_data].dt.date == pd.to_datetime(data_comparacao).date()]

            # Processamento da tabela2 (mantido igual)
            tabela2.columns = [self.get_column_name(tabela2, i) or f'Unnamed: {i}' for i in range(len(tabela2.columns))]

            for col in [1, 2, 3, 6]:
                col_name = f'Unnamed: {col}'
                if col_name in tabela2.columns:
                    tabela2[col_name] = tabela2[col_name].ffill()

            valores_filtro = ["HP", "TESTE FUNCIONAL", "PÓS COMPOSIÇÃO", "TRI",
                              "GRAVAÇÃO DO CI PTH", "GRAVAÇÃO DO CI SMT", "MONTAGEM MECÂNICA",
                              "EMBALAGEM", "TESTE ELÉTRICO"]

            if 'Unnamed: 6' not in tabela2.columns:
                messagebox.showerror("Erro", "A Tabela 2 não possui a coluna 'Unnamed: 6' necessária!")
                return

            tabela2_filtrada = tabela2[tabela2['Unnamed: 6'].isin(valores_filtro)]
            tabela3 = pd.DataFrame()

            col7_tabela2 = 'Unnamed: 7'
            col7_exists = col7_tabela2 in tabela2_filtrada.columns

            col2_tabela1 = self.get_column_name(tabela1, 2)
            if col2_tabela1 is None or col2_tabela1 not in tabela1_filtrada.columns:
                messagebox.showerror("Erro", "A Tabela 1 não possui a coluna 2 necessária!")
                return

            # Combina tabelas
            for _, row_tabela2 in tabela2_filtrada.iterrows():
                for _, row_tabela1 in tabela1_filtrada.iterrows():
                    condicao = (row_tabela2['Unnamed: 2'] == row_tabela1['Unnamed: 12']) and \
                               (row_tabela2['Unnamed: 6'] == row_tabela1['Unnamed: 14'])

                    if col7_exists and 'Unnamed: 2' in row_tabela1:
                        condicao = condicao and (row_tabela2[col7_tabela2] == row_tabela1['Unnamed: 2'])

                    if condicao:
                        tabela3 = pd.concat([tabela3, pd.DataFrame(row_tabela2).T], ignore_index=True)

            if 'Unnamed: 34' not in tabela3.columns:
                messagebox.showerror("Erro", "A tabela gerada não possui a coluna 'Unnamed: 34' necessária!")
                return

            tabela3 = tabela3[tabela3['Unnamed: 34'].notnull()]

            # Cria estrutura JSON
            plano_producao = {
                "metadata": {
                    "data_producao": datetime.strptime(data_comparacao, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d"),
                    "data_geracao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "versao": __version__
                },
                "atividades": []
            }

            # Processa cada linha e cria objetos JSON
            atividades_dict = {}  # Para agrupar horários por atividade

            for index_tabela3, row_tabela3 in tabela3.iterrows():
                # Cria chave única para a atividade
                chave_atividade = f"{row_tabela3.get('Unnamed: 2', '')}_{row_tabela3.get('Unnamed: 6', '')}_{row_tabela3.get('Unnamed: 7', '')}"

                if chave_atividade not in atividades_dict:
                    # Aplica substituições de clientes
                    cliente = str(row_tabela3.get('Unnamed: 1', ''))
                    if 'Substituicoes' in config:
                        substituicoes = {k.upper(): v for k, v in config['Substituicoes'].items()}
                        cliente = substituicoes.get(cliente.upper(), cliente)

                    produto = str(row_tabela3.get('Unnamed: 2', ''))
                    descricao = str(row_tabela3.get('Unnamed: 3', ''))
                    etapa = str(row_tabela3.get('Unnamed: 6', ''))

                    # Adiciona descrição para HP e TRI
                    if "HP" in etapa or "TRI" in etapa:
                        produto = f"{produto} - {descricao}"

                    atividades_dict[chave_atividade] = {
                        "cliente": cliente,
                        "produto": produto,
                        "descricao": descricao,
                        "etapa": etapa,
                        "rate": str(row_tabela3.get('Unnamed: 10', '')),
                        "celula": str(row_tabela3.get('Unnamed: 9', '')),
                        "haste": "COM TESTE",
                        "horarios": []
                    }

                # Processa horários para esta atividade
                for _, row_tabela1 in tabela1_filtrada_fil.iterrows():
                    condicao = (row_tabela3['Unnamed: 2'] == row_tabela1['Unnamed: 12']) and \
                               (row_tabela3['Unnamed: 6'] == row_tabela1['Unnamed: 14'])

                    if col7_exists and 'Unnamed: 2' in row_tabela1:
                        condicao = condicao and (row_tabela3[col7_tabela2] == row_tabela1['Unnamed: 2'])

                    if condicao and 'Unnamed: 15' in row_tabela1:
                        hora = row_tabela1['Unnamed: 15']
                        if pd.notna(hora):
                            slot_horario = self.processar_horario(hora)
                            if slot_horario and slot_horario not in atividades_dict[chave_atividade]["horarios"]:
                                atividades_dict[chave_atividade]["horarios"].append(slot_horario)

            # Remove atividades sem horários
            for chave, atividade in atividades_dict.items():
                if len(atividade["horarios"]) > 0:
                    atividade["horarios"].sort()  # Ordena os horários
                    plano_producao["atividades"].append(atividade)

            # Adiciona estatísticas
            plano_producao["estatisticas"] = {
                "total_atividades": len(plano_producao["atividades"]),
                "total_horarios": sum(len(a["horarios"]) for a in plano_producao["atividades"])
            }

            # Salva o arquivo JSON
            json_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("Arquivos JSON", "*.json")])

            if json_path:
                try:
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(plano_producao, f, ensure_ascii=False, indent=2)

                    mensagem = f"Plano de produção JSON gerado com sucesso!\n\n"
                    mensagem += f"📊 Estatísticas:\n"
                    mensagem += f"• Atividades: {plano_producao['estatisticas']['total_atividades']}\n"
                    mensagem += f"• Total de horários: {plano_producao['estatisticas']['total_horarios']}\n"
                    mensagem += f"• Arquivo: {json_path}"

                    messagebox.showinfo("Sucesso", mensagem)
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar o JSON: {str(e)}")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacaoDesktop(root)
    root.mainloop()