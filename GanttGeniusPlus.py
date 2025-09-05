import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import configparser
import json
from pathlib import Path
import unicodedata

__version__ = "4.1.0"  # Versão com correções de captura

config = configparser.ConfigParser()
config.read('config.ini')


class GanttJSONProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title(f"GanttGenius JSON Processor - V{__version__}")

        # Mapeamento de turnos (expandido)
        self.turnos_mapping = {
            'T1': {
                'inicio': 5.5,  # 5:30
                'fim': 15.3,  # 15:18
                'colunas': list(range(13, 24))  # Unnamed: 13-23
            },
            'T2': {
                'inicio': 15.3,  # 15:18
                'fim': 5.5,  # Até 05:30 do próximo dia
                'colunas': list(range(24, 34))  # Unnamed: 24-33
            }
        }

        # Etapas válidas expandidas
        self.etapas_validas = [
            "HP", "TESTE FUNCIONAL", "PÓS COMPOSIÇÃO", "TRI",
            "GRAVAÇÃO DO CI PTH", "GRAVAÇÃO DO CI SMT",
            "MONTAGEM MECÂNICA", "EMBALAGEM", "TESTE ELÉTRICO",
            "POS COMPOSICAO", "TESTE EL?TRICO", "MONTAGEM MEC?NICA",
            "GRAVACAO DO CI PTH", "GRAVACAO DO CI SMT"
        ]

        self.debug_mode = True  # Habilitar debug
        self.setup_ui()

    def setup_ui(self):
        """Interface com funcionalidade de extração dinâmica de datas"""
        self.frame = ttk.Frame(self.root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Tabela 1 - Programação de Recursos
        ttk.Label(self.frame, text="Programação de Recursos:").grid(row=0, column=0, sticky=tk.W)
        self.tabela1_entry = ttk.Entry(self.frame, width=40)
        self.tabela1_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar", command=self.carregar_tabela1).grid(row=0, column=2, sticky=tk.W)

        # Tabela 2 - Programação por Hora
        ttk.Label(self.frame, text="Programação por Hora:").grid(row=1, column=0, sticky=tk.W)
        self.tabela2_entry = ttk.Entry(self.frame, width=40)
        self.tabela2_entry.grid(row=1, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.frame, text="Carregar", command=self.carregar_tabela2).grid(row=1, column=2, sticky=tk.W)

        # NOVO: Botão para extrair datas da tabela
        ttk.Button(self.frame, text="Extrair Datas", command=self.atualizar_datas_disponiveis).grid(
            row=1, column=3, sticky=tk.W, padx=(5, 0))

        # Data de produção (inicialmente como DateEntry normal)
        ttk.Label(self.frame, text="Data de produção:").grid(row=2, column=0, sticky=tk.W)
        self.data_entry = DateEntry(self.frame, width=12, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.data_entry.grid(row=2, column=1, sticky=(tk.W, tk.E))

        # Label de status das datas
        self.status_datas_label = ttk.Label(self.frame,
                                            text="Clique em 'Extrair Datas' após carregar a tabela de horas",
                                            foreground='gray')
        self.status_datas_label.grid(row=2, column=2, columnspan=2, sticky=tk.W, padx=(5, 0))

        # Produto para debug
        ttk.Label(self.frame, text="Debug produto específico:").grid(row=3, column=0, sticky=tk.W)
        self.debug_entry = ttk.Entry(self.frame, width=20)
        self.debug_entry.grid(row=3, column=1, sticky=tk.W)
        ttk.Button(self.frame, text="Debug", command=self.debug_produto).grid(row=3, column=2, sticky=tk.W)

        # Botões de ação
        ttk.Button(self.frame, text="Gerar JSON", command=self.processar_dados).grid(
            row=4, column=0, columnspan=3, pady=10)

        # Área de resultado
        ttk.Label(self.frame, text="Resultado:").grid(row=5, column=0, sticky=tk.W)
        self.resultado_text = tk.Text(self.frame, height=25, width=100)
        self.resultado_text.grid(row=6, column=0, columnspan=4, pady=5)

        # Scrollbar para o resultado
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.resultado_text.yview)
        scrollbar.grid(row=6, column=4, sticky="ns")
        self.resultado_text.configure(yscrollcommand=scrollbar.set)

        # Inicializa variáveis para datas extraídas
        self.datas_extraidas = []
        self.datas_disponiveis = {}
        self.mapeamento_colunas_datas = {}

    def carregar_tabela1(self):
        filepath = filedialog.askopenfilename(
            title="Selecionar Programação de Recursos",
            filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
        )
        self.tabela1_entry.delete(0, tk.END)
        self.tabela1_entry.insert(0, filepath)

    def carregar_tabela2(self):
        filepath = filedialog.askopenfilename(
            title="Selecionar Programação por Hora",
            filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
        )
        self.tabela2_entry.delete(0, tk.END)
        self.tabela2_entry.insert(0, filepath)

    def normalizar_texto(self, texto):
        """Normaliza texto removendo acentos e caracteres especiais"""
        if pd.isna(texto) or texto == "":
            return ""
        texto_str = str(texto)
        # Remove acentos
        texto_normalizado = unicodedata.normalize('NFKD', texto_str)
        texto_normalizado = texto_normalizado.encode('ASCII', 'ignore').decode('ASCII')
        return texto_normalizado.upper().strip()

    def get_column_name(self, df, position):
        """Obtém o nome da coluna pela posição"""
        if position < len(df.columns):
            col_name = df.columns[position]
            if isinstance(col_name, (int, float)):
                return f'Unnamed: {position}'
            return col_name
        return None

    def hora_para_decimal(self, hora_obj):
        """Converte hora para decimal (ex: 15:18 = 15.3)"""
        if pd.isna(hora_obj):
            return None
        if hasattr(hora_obj, 'hour'):
            return hora_obj.hour + (hora_obj.minute / 60)
        return None

    def identificar_turno(self, hora_decimal):
        """Identifica o turno baseado na hora (corrigido)"""
        if hora_decimal is None:
            return None

        # T1: 5:30 às 15:18
        if 5.5 <= hora_decimal < 15.3:
            return 'T1'
        # T2: todo o resto (15:18 às 05:30 do próximo dia)
        else:
            return 'T2'

    def extrair_datas_cabecalho(self, tabela_hora):
        """Extrai as datas disponíveis na linha 5 das colunas Unnamed: 13 até Unnamed: 40"""

        # Linha 5  contém as datas
        linha_datas = 5

        # Colunas de data (13 até 40)
        colunas_data_inicio = 13
        colunas_data_fim = 68

        datas_disponiveis = {}

        if len(tabela_hora) <= linha_datas:
            print(f"AVISO: Tabela tem apenas {len(tabela_hora)} linhas. Linha {linha_datas + 1} não existe.")
            return datas_disponiveis

        if self.debug_mode:
            print(f"DEBUG: Extraindo datas da linha {linha_datas + 1} (índice {linha_datas})")

        # Percorre as colunas de 13 até 40
        for col_num in range(colunas_data_inicio, colunas_data_fim + 1):
            col_name = f'Unnamed: {col_num}'

            if col_name in tabela_hora.columns:
                # Pega o valor da linha 5 (índice 4)
                valor_data = tabela_hora.iloc[linha_datas][col_name]

                if pd.notna(valor_data):
                    # Tenta converter para data se for string
                    if isinstance(valor_data, str):
                        try:
                            # Tenta diferentes formatos de data
                            for formato in ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d']:
                                try:
                                    data_convertida = pd.to_datetime(valor_data, format=formato)
                                    datas_disponiveis[col_name] = {
                                        'coluna_numero': col_num,
                                        'valor_original': valor_data,
                                        'data_convertida': data_convertida,
                                        'data_formatada': data_convertida.strftime('%Y-%m-%d')
                                    }
                                    break
                                except:
                                    continue
                            else:
                                # Se não conseguiu converter, guarda o valor original
                                datas_disponiveis[col_name] = {
                                    'coluna_numero': col_num,
                                    'valor_original': valor_data,
                                    'data_convertida': None,
                                    'data_formatada': str(valor_data)
                                }
                        except:
                            datas_disponiveis[col_name] = {
                                'coluna_numero': col_num,
                                'valor_original': valor_data,
                                'data_convertida': None,
                                'data_formatada': str(valor_data)
                            }

                    # Se já é um objeto datetime
                    elif hasattr(valor_data, 'strftime'):
                        datas_disponiveis[col_name] = {
                            'coluna_numero': col_num,
                            'valor_original': valor_data,
                            'data_convertida': valor_data,
                            'data_formatada': valor_data.strftime('%Y-%m-%d')
                        }

                    # Para outros tipos, converte para string
                    else:
                        datas_disponiveis[col_name] = {
                            'coluna_numero': col_num,
                            'valor_original': valor_data,
                            'data_convertida': None,
                            'data_formatada': str(valor_data)
                        }

        if self.debug_mode:
            print(f"DEBUG: Datas encontradas: {len(datas_disponiveis)}")
            for col, info in list(datas_disponiveis.items())[:5]:  # Mostra apenas as primeiras 5
                print(f"  {col}: {info['data_formatada']} (original: {info['valor_original']})")

            if len(datas_disponiveis) > 5:
                print(f"  ... e mais {len(datas_disponiveis) - 5} datas")

        return datas_disponiveis

    def mapear_colunas_para_datas(self, datas_cabecalho):
        """Cria um mapeamento de colunas para datas para uso posterior"""

        mapeamento = {}

        for col_name, info_data in datas_cabecalho.items():
            col_numero = info_data['coluna_numero']
            data_formatada = info_data['data_formatada']

            mapeamento[col_numero] = {
                'coluna_nome': col_name,
                'data': data_formatada,
                'data_obj': info_data['data_convertida']
            }

        if self.debug_mode:
            print(f"DEBUG: Mapeamento coluna -> data criado para {len(mapeamento)} colunas")

        return mapeamento
    def debug_produto(self):
        """Função de debug para produto específico"""
        produto = self.debug_entry.get().strip()
        if not produto:
            messagebox.showwarning("Aviso", "Digite um código de produto para debug")
            return

        try:
            tabela1_path = self.tabela1_entry.get()
            tabela2_path = self.tabela2_entry.get()

            if not tabela1_path or not tabela2_path:
                messagebox.showerror("Erro", "Carregue ambos os arquivos primeiro!")
                return

            # Carrega as tabelas
            tabela1 = pd.read_excel(tabela1_path)
            tabela2 = pd.read_excel(tabela2_path)

            self.debug_produto_especifico(produto, tabela1, tabela2)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro no debug: {str(e)}")

    def debug_produto_especifico(self, produto_codigo, tabela1, tabela2):
        """Debug detalhado para um produto específico"""
        debug_info = f"=== DEBUG PRODUTO {produto_codigo} ===\n\n"

        # 1. Buscar na tabela de recursos
        col_produto_recurso = self.get_column_name(tabela1, 12)  # Unnamed: 12
        col_teste = self.get_column_name(tabela1, 27)  # Unnamed: 27
        col_data_recurso = self.get_column_name(tabela1, 15)  # Unnamed: 15

        recursos_match = tabela1[tabela1[col_produto_recurso] == produto_codigo]
        debug_info += f"1. TABELA RECURSOS:\n"
        debug_info += f"   - Linhas encontradas: {len(recursos_match)}\n"

        if len(recursos_match) > 0:
            for idx, row in recursos_match.iterrows():
                tem_teste = str(row.get(col_teste, "")).upper().strip()
                data_prod = row.get(col_data_recurso)
                etapa = row.get(self.get_column_name(tabela1, 14), "")

                debug_info += f"   - Linha {idx}: TESTE={tem_teste}, Data={data_prod}, Etapa={etapa}\n"

        # 2. Buscar na tabela de horas
        col_produto_hora = self.get_column_name(tabela2, 2)  # Unnamed: 2
        col_etapa_hora = self.get_column_name(tabela2, 6)  # Unnamed: 6

        # Renomear colunas da tabela2
        tabela2_renamed = tabela2.copy()
        tabela2_renamed.columns = [self.get_column_name(tabela2, i) or f'Unnamed: {i}'
                                   for i in range(len(tabela2.columns))]

        # Preencher valores mesclados
        for col in [1, 2, 3, 6]:
            col_name = f'Unnamed: {col}'
            if col_name in tabela2_renamed.columns:
                tabela2_renamed[col_name] = tabela2_renamed[col_name].ffill()

        hora_match = tabela2_renamed[tabela2_renamed[col_produto_hora] == produto_codigo]
        debug_info += f"\n2. TABELA HORAS:\n"
        debug_info += f"   - Linhas encontradas: {len(hora_match)}\n"

        if len(hora_match) > 0:
            for idx, row in hora_match.iterrows():
                etapa = row.get(col_etapa_hora, "")
                cliente = row.get('Unnamed: 1', "")
                debug_info += f"   - Linha {idx}: Cliente={cliente}, Etapa={etapa}\n"

        # 3. Verificar correlação
        debug_info += f"\n3. CORRELAÇÃO:\n"
        if len(recursos_match) > 0 and len(hora_match) > 0:
            correlacoes = 0
            for _, row_hora in hora_match.iterrows():
                for _, row_recurso in recursos_match.iterrows():
                    if row_hora.get(col_etapa_hora) == row_recurso.get(self.get_column_name(tabela1, 14)):
                        correlacoes += 1
                        debug_info += f"   - Correlação encontrada: {row_hora.get(col_etapa_hora)}\n"
            debug_info += f"   - Total de correlações: {correlacoes}\n"
        else:
            debug_info += "   - Não é possível correlacionar (produto não encontrado em uma das tabelas)\n"

        # Exibir resultado
        self.resultado_text.delete(1.0, tk.END)
        self.resultado_text.insert(tk.END, debug_info)

    def processar_programacao_recursos(self, tabela1):
        """Processa a tabela de programação de recursos (melhorado)"""
        # Verifica coluna de teste (posição 27)
        col_test = self.get_column_name(tabela1, 27)
        if not col_test or col_test not in tabela1.columns:
            raise ValueError("Tabela de recursos não possui coluna de teste (posição 27)")

        # Filtra apenas linhas com TESTE (mais flexível)
        tabela_teste = tabela1[
            tabela1[col_test].astype(str).str.upper().str.contains("TESTE", na=False)
        ].copy()

        if tabela_teste.empty:
            # Debug: mostrar valores únicos na coluna de teste
            valores_teste = tabela1[col_test].dropna().unique()
            raise ValueError(f"Nenhuma linha com TESTE encontrada. Valores únicos na coluna: {valores_teste[:10]}")

        # Verifica coluna de data (posição 15)
        col_data = self.get_column_name(tabela1, 15)
        if not col_data or col_data not in tabela_teste.columns:
            raise ValueError("Tabela de recursos não possui coluna de data (posição 15)")

        if 'Unnamed: 2' in tabela_teste.columns:
            tabela_teste = tabela_teste[
                ~tabela_teste['Unnamed: 2'].astype(str).str.upper().str.contains('ORDEM', na=False)
            ]

        tabela_filtrada = tabela_teste

        if self.debug_mode:
            print(f"DEBUG: Recursos com TESTE: {len(tabela_teste)}")

        return tabela_filtrada

    def processar_programacao_hora(self, tabela2):
        """Processa a tabela de programação por hora - versão simplificada"""

        # Renomeia colunas para padronizar
        tabela2.columns = [self.get_column_name(tabela2, i) or f'Unnamed: {i}'
                           for i in range(len(tabela2.columns))]

        if self.debug_mode:
            print(f"DEBUG: Colunas da tabela hora: {tabela2.columns.tolist()}")
            print(f"DEBUG: Total de linhas originais: {len(tabela2)}")

        # Preenche valores mesclados nas colunas importantes
        colunas_para_preencher = ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 6']
        for col in colunas_para_preencher:
            if col in tabela2.columns:
                tabela2[col] = tabela2[col].ffill()

        # Remove apenas as linhas que contêm "Qtd H.C." na coluna Unnamed: 11
        if 'Unnamed: 11' in tabela2.columns:
            # Filtra removendo linhas que contêm "Qtd H.C."
            tabela_filtrada = tabela2[
                ~tabela2['Unnamed: 11'].astype(str).str.contains('Qtd H.C.', na=False, case=False)
            ].copy()

            if self.debug_mode:
                linhas_removidas = len(tabela2) - len(tabela_filtrada)
                print(f"DEBUG: Linhas removidas com 'Qtd H.C.': {linhas_removidas}")
        else:
            print("AVISO: Coluna 'Unnamed: 11' não encontrada. Mantendo todas as linhas.")
            tabela_filtrada = tabela2.copy()

        if self.debug_mode:
            print(f"DEBUG: Linhas após remoção de 'Qtd H.C.': {len(tabela_filtrada)}")

            # Mostra exemplos de dados nas colunas principais
            if not tabela_filtrada.empty:
                print(f"DEBUG: Exemplo de dados:")
                colunas_exemplo = ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 6', 'Unnamed: 11']
                for col in colunas_exemplo:
                    if col in tabela_filtrada.columns:
                        valores_unicos = tabela_filtrada[col].dropna().unique()[:5]
                        print(f"  {col}: {valores_unicos}")

        return tabela_filtrada
    def correlacionar_dados(self, tabela_recursos, tabela_hora):
        """Correlaciona dados entre as duas tabelas - turno baseado na tabela_hora"""
        dados_correlacionados = []

        if self.debug_mode:
            print(f"DEBUG: Iniciando correlação - Recursos: {len(tabela_recursos)}, Horas: {len(tabela_hora)}")

        for idx_hora, row_hora in tabela_hora.iterrows():
            ordem = row_hora.get('Unnamed: 7')
            produto_hora = row_hora.get('Unnamed: 2')
            etapa_hora = row_hora.get('Unnamed: 6')
            total_prod_t1 = row_hora.get('')

            # Verifica se os valores não são nulos
            if pd.notna(ordem) and pd.notna(produto_hora) and pd.notna(etapa_hora):

                # Busca correspondências na tabela de recursos
                produtos_match = tabela_recursos[
                    (tabela_recursos['Unnamed: 12'] == produto_hora) &
                    (tabela_recursos['Unnamed: 2'] == ordem) &
                    (tabela_recursos['Unnamed: 14'] == etapa_hora)
                    ]

                # Se encontrou correspondência, adiciona a linha ORIGINAL da tabela_hora
                if not produtos_match.empty:
                    # Adiciona a linha completa da tabela_hora (não da tabela_recursos)
                    dados_correlacionados.append(row_hora)

        # Converte a lista de Series para DataFrame
        if dados_correlacionados:
            dados_correlacionados = pd.DataFrame(dados_correlacionados)
        else:
            dados_correlacionados = pd.DataFrame()  # Retorna DataFrame vazio se não houver correspondências

        if self.debug_mode:
            print(f"DEBUG: Total de correlações encontradas: {len(dados_correlacionados)}")
            if not dados_correlacionados.empty:
                print(f"DEBUG: Primeiras linhas correlacionadas:\n{dados_correlacionados.head()}")

        return dados_correlacionados

    # def correlacionar_dados(self, tabela_recursos, tabela_hora):
    #     """Correlaciona dados entre as duas tabelas (corrigido)"""
    #     dados_correlacionados = []
    #
    #     if self.debug_mode:
    #         print(f"DEBUG: Iniciando correlação - Recursos: {len(tabela_recursos)}, Horas: {len(tabela_hora)}")
    #
    #     for idx_hora, row_hora in tabela_hora.iterrows():
    #         ordem = row_hora.get('Unnamed: 7')
    #         produto_hora = row_hora.get('Unnamed: 2')
    #         etapa_hora = row_hora.get('Unnamed: 6')
    #
    #         if pd.isna(produto_hora) or pd.isna(etapa_hora):
    #             continue
    #
    #         # Busca correspondências na tabela de recursos
    #         produtos_match = tabela_recursos[
    #             (tabela_recursos['Unnamed: 12'] == produto_hora) &
    #             (tabela_recursos['Unnamed: 2'] == ordem) &
    #             (tabela_recursos['Unnamed: 14'] == etapa_hora)
    #             ]
    #
    #         for idx_recurso, row_recurso in produtos_match.iterrows():
    #             etapa_recurso = row_recurso.get('Unnamed: 14')
    #
    #
    #             # Extrai informações do horário
    #             hora_obj = row_recurso.get('Unnamed: 15')
    #             hora_decimal = None
    #             turno = None
    #
    #             if pd.notna(hora_obj) and hasattr(hora_obj, 'hour'):
    #                 hora_decimal = self.hora_para_decimal(hora_obj)
    #                 turno = self.identificar_turno(hora_decimal)
    #
    #             # Aceita todos os turnos válidos (T1 e T2)
    #             if turno in ['T1', 'T2']:
    #                 item = {
    #                     'cliente': str(row_hora.get('Unnamed: 1', '')),
    #                     'produto': str(row_hora.get('Unnamed: 2', '')),
    #                     'descricao': str(row_hora.get('Unnamed: 3', '')),
    #                     'etapa': str(row_hora.get('Unnamed: 6', '')),
    #                     'celula': str(row_hora.get('Unnamed: 9', '')),
    #                     'rate': str(row_hora.get('Unnamed: 10', '')),
    #                     'turno': turno,
    #                     'hora_programada': hora_obj.strftime('%H:%M') if pd.notna(hora_obj) else None,
    #                     'hora_decimal': hora_decimal,
    #                     'possui_teste': True,  # Já filtrado por TESTE
    #                     'dados_hora': {k: (v if pd.notna(v) else None) for k, v in row_hora.to_dict().items()},
    #                     'dados_recurso': {k: (v if pd.notna(v) else None) for k, v in
    #                                       row_recurso.to_dict().items()},
    #                     'motivo_correlacao': f'produto_etapa_match'
    #                 }
    #                 dados_correlacionados.append(item)
    #
    #                 if self.debug_mode:
    #                     print(f"DEBUG: Correlação encontrada - {produto_hora} - {etapa_hora} - {turno}")
    #
    #     if self.debug_mode:
    #         print(f"DEBUG: Total de correlações encontradas: {len(dados_correlacionados)}")
    #
    #     return dados_correlacionados

    def processar_dados(self):
        """Método principal de processamento (melhorado)"""
        try:
            # Validações iniciais
            tabela1_path = self.tabela1_entry.get()
            tabela2_path = self.tabela2_entry.get()
            # data_comparacao = self.data_entry.get() + ' 00:00:00'

            if not tabela1_path or not tabela2_path:
                messagebox.showerror("Erro", "Selecione ambos os arquivos!")
                return

            # Carrega as tabelas
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "Carregando arquivos...\n")
            self.root.update()

            try:
                tabela1 = pd.read_excel(tabela1_path)
                tabela2 = pd.read_excel(tabela2_path)
            except Exception as e:
                raise ValueError(f"Erro ao carregar arquivos Excel: {str(e)}")

            self.resultado_text.insert(tk.END, f"✓ Tabela recursos: {len(tabela1)} linhas\n")
            self.resultado_text.insert(tk.END, f"✓ Tabela hora: {len(tabela2)} linhas\n")
            self.root.update()

            # Processa programação de recursos
            self.resultado_text.insert(tk.END, "Processando programação de recursos...\n")
            self.root.update()

            try:
                tabela_recursos = self.processar_programacao_recursos(tabela1)
            except Exception as e:
                self.resultado_text.insert(tk.END, f"❌ Erro nos recursos: {str(e)}\n")
                return

            self.resultado_text.insert(tk.END, f"✓ Linhas com TESTE na data: {len(tabela_recursos)}\n")

            # Debug: mostrar alguns produtos encontrados
            if len(tabela_recursos) > 0:
                produtos_recursos = tabela_recursos['Unnamed: 12'].unique()[:5]
                self.resultado_text.insert(tk.END, f"✓ Produtos exemplo: {list(produtos_recursos)}\n")

            self.root.update()

            # Processa programação por hora
            self.resultado_text.insert(tk.END, "Processando programação por hora...\n")
            self.root.update()

            try:
                tabela_hora = self.processar_programacao_hora(tabela2)
            except Exception as e:
                self.resultado_text.insert(tk.END, f"❌ Erro na tabela hora: {str(e)}\n")
                return

            self.resultado_text.insert(tk.END, f"✓ Etapas filtradas: {len(tabela_hora)}\n")

            # Debug: mostrar alguns produtos da tabela hora
            if len(tabela_hora) > 0:
                produtos_hora = tabela_hora['Unnamed: 2'].unique()[:5]
                self.resultado_text.insert(tk.END, f"✓ Produtos hora exemplo: {list(produtos_hora)}\n")

            self.root.update()

            # Correlaciona dados
            self.resultado_text.insert(tk.END, "Correlacionando dados...\n")
            self.root.update()

            dados_correlacionados = self.correlacionar_dados(tabela_recursos, tabela_hora)

            # Aplica substituições de clientes se configurado
            if 'Substituicoes' in config:
                substituicoes = {k.upper(): v for k, v in config['Substituicoes'].items()}
                for item in dados_correlacionados:
                    cliente_original = item['cliente'].upper()
                    if cliente_original in substituicoes:
                        item['cliente'] = substituicoes[cliente_original]

            # Estatísticas por turno
            stats_turno = {}
            for item in dados_correlacionados:
                turno = item.get('turno', 'desconhecido')
                stats_turno[turno] = stats_turno.get(turno, 0) + 1

            # Produtos capturados
            produtos_capturados = [item['produto'] for item in dados_correlacionados]
            produtos_unicos = list(set(produtos_capturados))

            # Monta resultado final
            resultado = {
                'metadata': {
                    'versao': __version__,
                    'data_processamento': datetime.now().isoformat(),
                    'data_producao': self.data_entry.get(),
                    'arquivo_recursos': Path(tabela1_path).name,
                    'arquivo_hora': Path(tabela2_path).name,
                    'total_itens': len(dados_correlacionados),
                    'produtos_unicos': len(produtos_unicos),
                    'estatisticas_turno': stats_turno,
                    'produtos_capturados': produtos_unicos[:20],  # Primeiros 20 para debug
                    'debug_info': {
                        'linhas_recursos_original': len(tabela1),
                        'linhas_hora_original': len(tabela2),
                        'linhas_recursos_com_teste': len(tabela_recursos),
                        'linhas_hora_filtradas': len(tabela_hora),
                        'correlacoes_encontradas': len(dados_correlacionados)
                    }
                },
                'configuracao_turnos': self.turnos_mapping,
                'dados_producao': dados_correlacionados
            }

            # Exibe resultado
            json_resultado = json.dumps(resultado, indent=2, ensure_ascii=False, default=str)

            self.resultado_text.insert(tk.END, f"\n✓ Processamento concluído!\n")
            self.resultado_text.insert(tk.END, f"✓ Total de itens: {len(dados_correlacionados)}\n")
            self.resultado_text.insert(tk.END, f"✓ Produtos únicos: {len(produtos_unicos)}\n")

            for turno, count in stats_turno.items():
                self.resultado_text.insert(tk.END, f"✓ {turno}: {count} itens\n")

            self.resultado_text.insert(tk.END, "=" * 50 + "\n")

            # Mostrar alguns produtos capturados para verificação
            self.resultado_text.insert(tk.END, f"Produtos capturados (primeiros 10):\n")
            for produto in produtos_unicos[:10]:
                self.resultado_text.insert(tk.END, f"  - {produto}\n")

            self.resultado_text.insert(tk.END, "=" * 50 + "\n")
            self.resultado_text.insert(tk.END, json_resultado)

            # Verificar se produto específico foi capturado
            produto_teste = "31000002046"
            if produto_teste in produtos_unicos:
                self.resultado_text.insert(tk.END, f"\n✅ PRODUTO {produto_teste} FOI CAPTURADO!\n")
            else:
                self.resultado_text.insert(tk.END, f"\n❌ PRODUTO {produto_teste} NÃO FOI CAPTURADO\n")
                # Executar debug automático
                self.debug_entry.delete(0, tk.END)
                self.debug_entry.insert(0, produto_teste)

            # Opção de salvar JSON
            salvar = messagebox.askyesno("Salvar JSON",
                                         f"Processamento concluído!\n\n"
                                         f"Total de itens: {len(dados_correlacionados)}\n"
                                         f"Produtos únicos: {len(produtos_unicos)}\n"
                                         f"Deseja salvar o JSON?")

            if salvar:
                json_path = filedialog.asksaveasfilename(
                    defaultextension=".json",
                    filetypes=[("Arquivos JSON", "*.json")]
                )

                if json_path:
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(resultado, f, indent=2, ensure_ascii=False, default=str)
                    messagebox.showinfo("Sucesso", f"JSON salvo em: {json_path}")

        except Exception as e:
            error_msg = f"Erro durante o processamento: {str(e)}"
            self.resultado_text.insert(tk.END, f"\n❌ {error_msg}\n")
            messagebox.showerror("Erro", error_msg)

    def criar_metodo_banco(self):
        """Método para futura injeção no banco de dados"""
        # Placeholder para implementação futura
        pass


if __name__ == "__main__":
    root = tk.Tk()
    app = GanttJSONProcessor(root)
    root.mainloop()
