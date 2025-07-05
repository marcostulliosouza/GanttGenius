# **GanttGenius - Conversor de Tabela para BetterCall**  
**Versão:** 3.5.7  

---

## **1. Visão Geral**  
O **GanttGenius** é uma ferramenta desktop desenvolvida em **Python** para conversão automatizada de tabelas de programação de produção em um formato compatível com o sistema **BetterCall**.  

A aplicação realiza:  
- **Processamento de duas tabelas de entrada** (Programação de Recursos e Programação por Hora_V2).  
- **Filtragem inteligente** com base na data de produção selecionada.  
- **Validação de estrutura** das colunas obrigatórias.  
- **Substituição automática** de códigos de clientes por nomes completos (via `config.ini`).  
- **Geração de um arquivo Excel** formatado para importação no BetterCall.  

---

## **2. Arquitetura e Tecnologias**  
### **2.1. Linguagem e Bibliotecas**  
- **Python 3.7+**  
- **Bibliotecas principais:**  
  - `tkinter` → Interface gráfica (GUI)  
  - `pandas` → Manipulação de dados tabulares  
  - `configparser` → Leitura do arquivo `config.ini`  
  - `tkcalendar` → Seleção de datas na interface  

### **2.2. Fluxo de Processamento**  
1. **Entrada de Dados:**  
   - Carrega **Tabela 1** (Programação de Recursos).  
   - Carrega **Tabela 2** (Programação por Hora_V2).  
2. **Filtragem e Validação:**  
   - Filtra registros com a tag **"TESTE"**.  
   - Valida se a **data de produção** existe nas tabelas.  
3. **Junção de Dados:**  
   - Combina as tabelas com base em **chaves comuns** (`Unnamed: 12`, `Unnamed: 14`).  
4. **Mapeamento de Horários:**  
   - Converte horários em colunas específicas (ex: `05:30-06:00` → `Unnamed: 13`).  
5. **Pós-processamento:**  
   - Remove linhas duplicadas.  
   - Aplica **substituições de clientes** (ex: `WAY → WAYNE`).  
6. **Saída:**  
   - Gera um **arquivo Excel** no formato esperado pelo BetterCall.  

---

## **3. Configuração (`config.ini`)**  
### **3.1. Seção `[Substituicoes]`**  
Define mapeamentos de **códigos de clientes** para **nomes completos**:  
```ini
WAY = WAYNE  
LNA = LANDIS+GYR NORTE AMERICANA  
FAN = FANEM  
... (outros clientes)
```  

### **3.2. Seção `[Settings]`**  
Configurações gerais:  
```ini
default_test_mode = true  
log_level = INFO  
```  

---

## **4. Validações Implementadas**  
### **4.1. Verificação de Estrutura**  
- **Tabela 1:**  
  - Coluna 27 (`TESTE`) → Deve conter a tag "TESTE".  
  - Coluna 15 (`Data`) → Deve ser uma data válida.  
- **Tabela 2:**  
  - Coluna 6 (`Unnamed: 6`) → Deve conter valores como `HP`, `TESTE FUNCIONAL`, etc.  

### **4.2. Filtros Automáticos**  
- Remove linhas onde:  
  - **Não há horários reais** (apenas "1" automático).  
  - **Produto está vazio ou inválido**.  

### **4.3. Tratamento de Dados**  
- **Preenchimento de células mescladas** (`.ffill()` nas colunas 1, 2, 3, 6).  
- **Formatação de datas** (`YYYY-MM-DD` → `DD/MM/YYYY`).  

---

## **5. Interface do Usuário (GUI)**  
### **5.1. Componentes**  
- **Tabela 1:** Campo para carregar a **Programação de Recursos**.  
- **Tabela 2:** Campo para carregar a **Programação por Hora_V2**.  
- **Seletor de Data:** Usa `DateEntry` (tkcalendar).  
- **Botão "Gerar Gantt BetterCall":** Inicia o processamento.  

### **5.2. Mensagens de Feedback**  
- **Erros:**  
  - "A Tabela 1 não possui a coluna de teste necessária!"  
  - "Arquivo não possui plano para a data solicitada."  
- **Sucesso:**  
  - Exibe estatísticas (linhas processadas, removidas, finais).  

---

## **6. Instalação e Execução**  
### **6.1. Pré-requisitos**  
```bash
pip install pandas tkcalendar openpyxl
```  

### **6.2. Executando a Aplicação**  
```bash
python main.py
```  

---

## **7. Exemplo de Saída (Excel)**  
| Unnamed: 0 (Cliente) | Unnamed: 1 (Produto) | Unnamed: 3 (Etapa) | ... | Unnamed: 34 (Horário) |  
|----------------------|----------------------|--------------------|-----|-----------------------|  
| WAYNE                | Produto X - HP       | TESTE FUNCIONAL    | ... | 1 (06:00-07:00)       |  

--- 

**Nota:** Esta documentação refere-se à versão **3.5.7**. Consulte o código-fonte para atualizações.