import ttkbootstrap as ttkb 
from tkinter import messagebox, Listbox, Canvas, Text, Toplevel, simpledialog, filedialog
from ttkbootstrap import ttk
from datetime import datetime
import json
import pandas as pd
import numpy as np
import concurrent.futures as cf 
import os
import threading 
import warnings 

try:
    from escpos.printer import Usb, Network
    
    IMPRESSORA_CONFIG = {"tipo": "USB", "vid": 0x1FC9, "pid": 0x2016} 
    PRINTER_LOADED = True
    print(" Biblioteca ESC/POS carregada com sucesso!")
    
except ImportError as e:
    print(f" AVISO: Biblioteca ESC/POS não encontrada: {e}")
    print(" Criando classes dummy para evitar erros...")
    PRINTER_LOADED = False
    
    class Usb:
        def __init__(self, *args, **kwargs): 
            raise RuntimeError("Biblioteca ESC/POS não instalada. Instale com: pip install python-escpos")
        def text(self, *args, **kwargs): pass
        def set(self, *args, **kwargs): pass
        def image(self, *args, **kwargs): pass
        def cut(self): pass
        def hw(self, *args, **kwargs): pass
    
    class Network(Usb): pass
    
    IMPRESSORA_CONFIG = {"tipo": "USB", "vid": 0x1FC9, "pid": 0x2016}

# -----------------------------------------------

# SILENCIAR FUTURE WARNING
warnings.simplefilter(action='ignore', category=FutureWarning)

# ---------------- CONFIGURAÇÃO DO EXCEL LOCAL ----------------
EXCEL_FILE = 'pedidos.xlsx' 
WORKSHEET_NAME = 'Pedidos' 
EXCEL_LOCK = threading.Lock() 

# VARIÁVEIS GLOBAIS
pedidos_cache = [] 
data_atual = datetime.now().strftime("%d/%m/%Y")
historico_clientes = {} 
tabela_ativos = None 
tabela_arquivados = None 
tabela_clientes = None 
relatorio_label = None 
executor = cf.ThreadPoolExecutor(max_workers=2) 
janela = None 
HEADERS = ['cliente', 'telefone', 'endereco', 'forma_pagamento', 'precisa_troco', 
           'troco_valor', 'taxa_entrega', 'status', 'valor_total', 'hora', 'itens']

# ---------------- Preços ----------------
precos_pizzas = {
    "Muçarela": 45, "Calabresa": 45, "Presunto": 45, "Mista": 50, "Bacon": 50,
    "Marguerita": 50, "Portuguesa": 50, "Frango": 50, "Calabacon": 50,
    "Calabresa Apimentada": 50, "Quatro Queijos": 50, "Frango com Requeijao": 55,
    "Nordestina": 60, "Frango Cremoso": 60, "Siciliana": 65, "Peito de Peru": 65,
    "À Moda da Casa": 65, "Paraense": 65, "Filé": 65, "Camarão Rosa": 70,
    "Filé com Bacon": 70, "Savino Especial": 70, "Brigadeiro": 35, "Disqueti": 35
}
precos_bordas = {"Catupiry": 16, "Chocolate": 11, "Cheddar": 14, "Requeijão": 12}
precos_adicionais = {"Ovo": 2, "Cebola": 2, "Tomate": 3, "Pimentão Verde": 3, "Jambu": 7, "Bacon": 8,
                     "Calabresa": 8, "Muçarela": 10, "Frango": 10, "Requeijão": 10, "Cheddar": 10,
                     "Peito de Peru": 11, "Catupiry": 12}

def calcular_preco_pizza(sabores_selecionados):
    precos = [precos_pizzas.get(s, 0) for s in sabores_selecionados]
    if not precos:
        return 0
    
    # Se tiver mais de uma pizza, soma e divide pela quantidade (ou 2, se quiser fixo)
    if len(precos) > 1:
        return sum(precos) / len(precos)  # média de todas
        # ou, se quiser fixar exatamente 2 sabores:
        # return (precos[0] + precos[1]) / 2
    else:
        return precos[0]

# ---------------- Classe de Autocomplete ----------------
class AutocompleteEntry(ttkb.Entry):
    def __init__(self, parent, autocomplete_list, callback_on_select=None, *args, **kwargs):
        if "textvariable" in kwargs: self.var = kwargs["textvariable"]
        else: self.var = ttkb.StringVar(); kwargs["textvariable"] = self.var
        
        super().__init__(parent, *args, **kwargs)

        self.autocomplete_list = autocomplete_list
        self.callback_on_select = callback_on_select 
        self.var.trace_add("write", self.changed)
        
        self.bind("<Right>", self.selection); self.bind("<Up>", self.move_up); self.bind("<Down>", self.move_down)
        self.bind("<FocusOut>", lambda e: self.after(200, self.hide_listbox))
        self.bind("<Return>", self.selection) 
        self.bind("<KeyRelease-Return>", self.selection) 

        self.listbox_up = False

    def changed(self, name, index, mode):
        if str(self) != str(self.focus_get()): return
        
        if self.var.get() == "":
            if self.listbox_up: self.listbox.destroy(); self.listbox_up = False
        else:
            words = self.comparison()
            if words:
                if not self.listbox_up:
                    x = self.winfo_rootx() - self.winfo_toplevel().winfo_rootx()
                    y = self.winfo_rooty() - self.winfo_toplevel().winfo_rooty() + self.winfo_height()
                    self.listbox = Listbox(self.winfo_toplevel(), width=self["width"], height=min(len(words), 8))
                    self.listbox.bind("<Button-1>", self.selection)
                    self.listbox.bind("<KeyRelease-Return>", self.selection)
                    self.listbox.place(x=x, y=y); self.listbox_up = True
                
                current_items = set(self.listbox.get(0, ttkb.END))
                if current_items != set(words):
                    self.listbox.delete(0, ttkb.END)
                    for w in words: self.listbox.insert(ttkb.END, w)
            else:
                if self.listbox_up: self.listbox.destroy(); self.listbox_up = False
    
    def selection(self, event):
        if self.listbox_up and self.listbox.curselection():
            selected_text = self.listbox.get(ttkb.ACTIVE)
            self.var.set(selected_text)
            
            if self.callback_on_select:
                self.callback_on_select(selected_text)
            
            # CORRIGIDO: Fecha a Listbox após a seleção (seja por clique ou Enter)
            self.listbox.destroy(); self.listbox_up = False; self.icursor(ttkb.END)
            
        elif self.listbox_up:
            # Se for um evento Return (Enter) e não havia seleção, apenas fecha
            self.listbox.destroy(); self.listbox_up = False
            
    def move_up(self, event):
        if self.listbox_up:
            index = self.listbox.curselection()[0] if self.listbox.curselection() else 0
            if index > 0:
                self.listbox.selection_clear(first=index)
                new_index = str(index - 1)
                self.listbox.selection_set(first=new_index)
                self.listbox.activate(new_index)
                self.listbox.see(new_index)
    
    def move_down(self, event):
        if self.listbox_up:
            index = self.listbox.curselection()[0] if self.listbox.curselection() else -1
            if index < self.listbox.size() - 1:
                self.listbox.selection_clear(first=index)
                new_index = str(index + 1)
                self.listbox.selection_set(first=new_index)
                self.listbox.activate(new_index)
                self.listbox.see(new_index)
    
    def comparison(self):
        return [w for w in self.autocomplete_list if self.var.get().lower() in w.lower()][:10]
    
    def hide_listbox(self, event=None):
        if self.listbox_up: 
            if str(self.winfo_toplevel().focus_get()) == str(self.listbox):
                return
            self.listbox.destroy()
            self.listbox_up = False

# ---------------- Funções de Conexão e Persistência (Excel) ----------------

def _garantir_arquivo_excel():
    """Cria o arquivo Excel e o cabeçalho se ele não existir."""
    if not os.path.exists(EXCEL_FILE):
        try:
            df_vazio = pd.DataFrame(columns=HEADERS)
            df_vazio.to_excel(EXCEL_FILE, sheet_name=WORKSHEET_NAME, index=False)
        except Exception as e:
            messagebox.showerror("Erro de Inicialização", f"Falha ao criar o arquivo Excel ({EXCEL_FILE}). Verifique permissões. Erro: {e}")
            return False
    return True

def carregar_pedidos():
    """Carrega todos os pedidos do Excel para o cache e monta o histórico de clientes."""
    global pedidos_cache
    global historico_clientes
    
    if not _garantir_arquivo_excel():
        pedidos_cache = []
        historico_clientes = {}
        return
        
    try:
        with EXCEL_LOCK:
            df = pd.read_excel(EXCEL_FILE, sheet_name=WORKSHEET_NAME)
        
        df = df.fillna('')
        data = df.to_dict('records')
        
        historico_clientes = {} 
        
        for row in data:
            if 'itens' in row and row['itens']:
                try:
                    row['itens'] = json.loads(str(row['itens']).strip())
                except json.JSONDecodeError:
                    row['itens'] = [] 
            
            row['valor_total'] = float(row.get('valor_total', 0.0))
            row['taxa_entrega'] = float(row.get('taxa_entrega', 0.0))
            
            # Popula o histórico de clientes (incluindo telefone)
            cliente = str(row.get('cliente', '')).strip()
            telefone_bruto = str(row.get('telefone', '')).strip()
            # Garante que o telefone é uma string e remove .0 se for float
            telefone = telefone_bruto.split('.')[0] if '.' in telefone_bruto else telefone_bruto
            
            # Chave principal de busca é sempre o Telefone (ou Nome se Telefone for vazio)
            chave_principal = telefone if telefone else cliente
            
            if chave_principal:
                if chave_principal not in historico_clientes:
                    historico_clientes[chave_principal] = {
                        'cliente': cliente,
                        'telefone': telefone,
                        'endereco': str(row.get('endereco', '')).strip(),
                        'pedidos_feitos': 0,
                        'ultima_hora': ''
                    }
                
                historico_clientes[chave_principal]['pedidos_feitos'] += 1
                historico_clientes[chave_principal]['ultima_hora'] = row.get('hora', '') 
            
        pedidos_cache = data
    except Exception as e:
        messagebox.showerror("Erro ao Carregar", f"Erro ao ler dados do Excel. Feche o arquivo se estiver aberto. Erro: {e}")
        pedidos_cache = []
        historico_clientes = {}

def salvar_pedido(dados_do_pedido, modo_edicao=False, hora_original=None):
    """Salva um novo pedido ou atualiza um existente no Excel. RODA EM THREAD DE FUNDO."""
    
    if not _garantir_arquivo_excel():
        return False
        
    try:
        with EXCEL_LOCK:
            df = pd.read_excel(EXCEL_FILE, sheet_name=WORKSHEET_NAME)
            
            linha_para_salvar = dados_do_pedido.copy()
            
            # 1. TRATAMENTO DE JSON
            if linha_para_salvar.get('itens') is None or (isinstance(linha_para_salvar['itens'], float) and np.isnan(linha_para_salvar['itens'])):
                linha_para_salvar['itens'] = '[]'
            else:
                 linha_para_salvar['itens'] = json.dumps(linha_para_salvar['itens'], ensure_ascii=False)
            
            df_novo_pedido = pd.DataFrame([linha_para_salvar], columns=HEADERS)

            if modo_edicao and hora_original:
                # Buscamos pela combinação hora + cliente (a chave única para edição)
                indices_para_atualizar = df[(df['hora'] == hora_original) & 
                                            (df['cliente'] == dados_do_pedido['cliente'])].index
                
                if not indices_para_atualizar.empty:
                    df.iloc[indices_para_atualizar[0]] = df_novo_pedido.iloc[0]
                else:
                    print("Pedido original não encontrado para edição.")
                    return False
            else:
                df = pd.concat([df, df_novo_pedido], ignore_index=True)

            # 2. TRATAMENTO DE SALVAMENTO: Garante que os valores NaN/None nas células sejam strings vazias no Excel
            df = df.fillna('')

            df.to_excel(EXCEL_FILE, sheet_name=WORKSHEET_NAME, index=False)
            
        return True
    
    except Exception as e:
        # Se for um Permission denied (arquivo aberto), avisa de forma mais útil
        if "Permission denied" in str(e) or "access is denied" in str(e):
             janela.after(0, lambda: messagebox.showerror("Erro Crítico de Salvamento", "Falha ao salvar no Excel. O arquivo 'pedidos.xlsx' está ABERTO ou BLOQUEADO. Por favor, feche-o e tente novamente."))
        else:
             print(f"Erro no Thread de Salvamento do Excel: {e}")
             janela.after(0, lambda: messagebox.showerror("Erro de Salvamento", f"Ocorreu um erro inesperado ao salvar: {e}"))
        return False

# ---------------- Funções de Relatório ----------------

def gerar_relatorio_geral():
    """Calcula e exibe o resumo geral dos pedidos."""
    carregar_pedidos() 
    global pedidos_cache
    
    total_pedidos = len(pedidos_cache)
    soma_total = sum(p.get('valor_total', 0.0) for p in pedidos_cache)
    
    if total_pedidos == 0:
        if relatorio_label:
            relatorio_label.config(text="Nenhum pedido encontrado.")
        return

    resumo_pagamento = {}
    for pedido in pedidos_cache:
        forma = pedido.get('forma_pagamento', 'Desconhecida')
        valor = float(pedido.get('valor_total', 0.0))
        resumo_pagamento[forma] = resumo_pagamento.get(forma, 0) + valor

    relatorio_texto = [
        "--- RESUMO GERAL (TODOS OS PEDIDOS) ---",
        f"📋 Quantidade Total de Pedidos: {total_pedidos}",
        f"💵 Soma Total de Vendas: R$ {soma_total:.2f}".replace('.', ','), 
        "\n--- Vendas por Forma de Pagamento ---"
    ]
    for forma, total in resumo_pagamento.items():
        relatorio_texto.append(f"  > {forma}: R$ {total:.2f}".replace('.', ','))

    if relatorio_label:
        relatorio_label.config(text="\n".join(relatorio_texto)) 

def _gerar_relatorio_dados(data_filtro=None):
    """
    Função central que filtra os dados e prepara o relatório estatístico.
    data_filtro deve ser uma string DD/MM ou None.
    """
    carregar_pedidos()
    
    pedidos_filtrados = pedidos_cache
    if data_filtro:
        # CORRIGIDO: Filtrar pedidos que começam com a data (DD/MM)
        pedidos_filtrados = [p for p in pedidos_cache if str(p.get('hora', '')).startswith(data_filtro)]

    if not pedidos_filtrados:
        return None, "Nenhum pedido encontrado para o período selecionado.", []

    total_pedidos = len(pedidos_filtrados)
    soma_total = sum(p.get('valor_total', 0.0) for p in pedidos_filtrados)
    
    resumo_pagamento = {}
    for pedido in pedidos_filtrados:
        forma = pedido.get('forma_pagamento', 'Desconhecida')
        valor = float(pedido.get('valor_total', 0.0))
        resumo_pagamento[forma] = resumo_pagamento.get(forma, 0) + valor

    relatorio_texto = [
        f"--- RESUMO DO DIA: {data_filtro or 'TODOS OS PEDIDOS'} ---",
        f"📋 Quantidade Total de Pedidos: {total_pedidos}",
        f"💵 Soma Total de Vendas: R$ {soma_total:.2f}".replace('.', ','), 
        "\n--- Vendas por Forma de Pagamento ---"
    ]
    for forma, total in resumo_pagamento.items():
        relatorio_texto.append(f"  > {forma}: R$ {total:.2f}".replace('.', ','))
        
    return "\n".join(relatorio_texto), total_pedidos > 0, pedidos_filtrados

def _executar_limpeza_planilha():
    """
    Função de limpeza que zera o arquivo Excel (mantendo apenas o cabeçalho).
    """
    if not _garantir_arquivo_excel(): 
        return False

    try:
        with EXCEL_LOCK:
            df_vazio = pd.DataFrame(columns=HEADERS)
            df_vazio.to_excel(EXCEL_FILE, sheet_name=WORKSHEET_NAME, index=False)
        return True
    except Exception as e:
        print(f"Erro de limpeza no thread: {e}")
        janela.after(0, lambda: messagebox.showerror("Erro de Limpeza", "Falha ao limpar o Excel. Certifique-se de que o arquivo 'pedidos.xlsx' está FECHADO."))
        return False

def verificar_limpeza(future_task):
    """Verifica se a limpeza terminou (Roda na Thread Principal)."""
    if future_task.done():
        try:
            sucesso = future_task.result()
            if sucesso:
                messagebox.showinfo("Sucesso", "Dados de pedidos foram limpos com sucesso! O histórico de clientes (Nome, Tel, Endereço) também foi zerado.")
                atualizar_lista()
                atualizar_lista_arquivados() 
                atualizar_lista_clientes() 
                gerar_relatorio_geral()
            else:
                messagebox.showerror("Erro", "Falha ao limpar os dados da planilha. Verifique o log.")
        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado: {e}")
    else:
        janela.after(100, lambda: verificar_limpeza(future_task))

def limpar_dados_planilha():
    """Inicia a limpeza da planilha em thread de fundo."""
    global janela
    # CORRIGIDO: Mensagem de aviso mais clara sobre a perda total de histórico de pedidos.
    if not messagebox.askyesno("Confirmar Limpeza Total", 
                               "ATENÇÃO! Esta ação apagará TODOS os pedidos, VENDAS e HISTÓRICO DE PEDIDOS dos clientes.\n\n"
                               "Você tem certeza que deseja limpar a planilha?"):
        return
    
    future = executor.submit(_executar_limpeza_planilha)
    janela.after(100, lambda: verificar_limpeza(future))

# ---------------- Funções de Atualização das Tabelas ----------------

def atualizar_lista():
    """Atualiza a lista de pedidos ATIVOS (não 'Entregue') na interface usando o cache."""
    carregar_pedidos() 
    
    global tabela_ativos
    if tabela_ativos:
        for row in tabela_ativos.get_children():
            tabela_ativos.delete(row)
            
        global pedidos_cache
        for pedido in pedidos_cache:
            if pedido.get("status", "N/A") == "Entregue": 
                continue 
            
            cliente = pedido.get("cliente", "Cliente Desconhecido")
            valor_total = float(pedido.get("valor_total", 0.00)) 
            status = pedido.get("status", "N/A")
            # Usa a hora completa (DD/MM HH:MM:SS) como a chave única de busca na tabela
            hora_completa = str(pedido.get("hora", "N/A")) 
            
            num_itens = len(pedido.get('itens', []))
            first_item_desc = pedido['itens'][0]['descricao'] if num_itens > 0 and 'itens' in pedido else "Vazio"
            descricao_pedido = f"{num_itens} Itens ({first_item_desc[:40]}{'...' if len(first_item_desc) > 40 else ''})"
            
            # Tags de cor para status (CORES MAIS DISCRETAS)
            if status == "Em Preparação": tag = 'preparacao' 
            elif status == "Saiu pra entrega": tag = 'entrega' 
            elif status == "Pendente": tag = 'pendente' 
            else: tag = ''

            tabela_ativos.insert("", ttkb.END, values=(
                cliente, 
                descricao_pedido, 
                f"R$ {valor_total:.2f}".replace('.', ','), 
                status, 
                hora_completa
            ), tags=(tag,))


def atualizar_lista_arquivados():
    """Atualiza a lista de pedidos ARQUIVADOS ('Entregue')."""
    carregar_pedidos() 
    
    global tabela_arquivados
    if tabela_arquivados:
        for row in tabela_arquivados.get_children():
            tabela_arquivados.delete(row)
            
        global pedidos_cache
        for pedido in pedidos_cache:
            if pedido.get("status", "N/A") != "Entregue": 
                continue 
            
            cliente = pedido.get("cliente", "Cliente Desconhecido")
            valor_total = float(pedido.get("valor_total", 0.00)) 
            status = pedido.get("status", "N/A")
            hora_completa = str(pedido.get("hora", "N/A")) # CHAVE ÚNICA DE BUSCA
            
            num_itens = len(pedido.get('itens', []))
            first_item_desc = pedido['itens'][0]['descricao'] if num_itens > 0 and 'itens' in pedido else "Vazio"
            descricao_pedido = f"{num_itens} Itens ({first_item_desc[:40]}{'...' if len(first_item_desc) > 40 else ''})"
            
            tabela_arquivados.insert("", ttkb.END, values=(
                cliente, 
                descricao_pedido, 
                f"R$ {valor_total:.2f}".replace('.', ','), 
                status, 
                hora_completa
            ), tags=('entregue',))

def atualizar_lista_clientes():
    """Atualiza a lista de Clientes para a aba Fidelidade."""
    carregar_pedidos()
    
    global tabela_clientes
    if tabela_clientes:
        for row in tabela_clientes.get_children():
            tabela_clientes.delete(row)
            
        global historico_clientes
        
        lista_clientes = []
        clientes_unicos = {}
        
        for chave, dados in historico_clientes.items():
            
            chave_exibicao = dados['telefone'] if dados['telefone'] else dados['cliente']
            
            if chave_exibicao and chave_exibicao not in clientes_unicos:
                
                telefone = str(dados['telefone'])
                if telefone.endswith('.0'):
                     telefone = telefone[:-2]
                    
                lista_clientes.append((dados['cliente'] or "N/A",      # 1. Nome
                    telefone or "N/A",              # 2. Telefone
                    dados['endereco'],              # 3. Endereço
                    dados['pedidos_feitos']         # 4. Total Pedidos
                ))
                clientes_unicos[chave_exibicao] = True
            
        # Ordena a lista pelo número de pedidos (maior para o menor)
        lista_clientes.sort(key=lambda x: x[3], reverse=True)
        
        for cliente in lista_clientes:
            tabela_clientes.insert("", ttkb.END, values=cliente, tags=('fidelidade',))

# ---------------- Funções de Ação (e Impressão) ----------------

def gerar_cupom(pedido):
    """Monta o conteúdo formatado do cupom para impressão."""
    
    valor_total_cupom = pedido.get('valor_total', 0.00)
    
    # Formatação do texto
    texto = "\n"
    texto += "{:^32}\n".format("------ PEDIDO ------")
    texto += "Cliente: {:<15}\n".format(str(pedido.get('cliente', 'N/A'))[:25]) 
    texto += "Telefone: {:<15}\n".format(str(pedido.get('telefone', 'N/A'))[:15]) 
    texto += "Endereco: {:<30}\n".format(str(pedido.get('endereco', 'N/A'))[:30]) 
    texto += "Data: {:<10}\n".format(data_atual)
    texto += "Hora: {:<20}\n".format(str(pedido.get('hora', 'N/A')).split()[-1])
    texto += "--------------------------------\n"
    texto += "ITENS:\n"
    
    for i, item in enumerate(pedido.get('itens', [])):
        preco = item.get('valor', 0.0)
        
        # Nome do Item em Negrito
        texto += "{}. {:<20} R$ {:>6.2f}\n".format(i+1, item.get('descricao', 'Item').upper()[:20], preco)
        
        if item.get('borda') and item['borda'] != 'Não':
             texto += "  -> BORDA: {}\n".format(item['borda'])
        if item.get('adicionais'):
             texto += "  -> ADICIONAIS: {}\n".format(', '.join(item['adicionais']))
        if item.get('observacoes', '').strip():
             texto += "  -> OBS: {}\n".format(item['observacoes'])
             
        texto += "--------------------------------\n"

    texto += "\n"
    taxa_entrega = pedido.get('taxa_entrega', 0.0)
    subtotal = valor_total_cupom - taxa_entrega
    
    texto += "Subtotal: R${:>21.2f}\n".format(subtotal)
    texto += "Taxa Entrega: R${:>18.2f}\n".format(taxa_entrega)
    texto += "VALOR TOTAL: R${:>17.2f}\n".format(valor_total_cupom)
    texto += "Pagamento: {}\n".format(pedido.get('forma_pagamento', 'N/A'))
    
    if pedido.get("precisa_troco") == True:
        # CORRIGIDO: Garante que o troco é formatado como string
        troco_valor = str(pedido.get('troco_valor', '0,00')).replace('.', ',')
        texto += "TROCO P/ R$ {:>16}\n".format(troco_valor)
        
    texto += "--------------------------------\n"
    texto += "Endereco: {}\n".format(pedido.get('endereco', 'N/A'))
    texto += "================================\n\n"
    
    return texto

def imprimir_pedido():
    if not PRINTER_LOADED:
        messagebox.showerror("Erro de Impressão", "A biblioteca ESC/POS não está funcional. Por favor, instale 'python-escpos' e as dependências USB.")
        return
        
    global tabela_ativos
    item_sel = tabela_ativos.focus()
    if not item_sel:
        messagebox.showwarning("Aviso", "Selecione um pedido para imprimir.")
        return
        
    valores = tabela_ativos.item(item_sel)["values"]
    cliente_nome = valores[0]
    hora_completa = valores[4]
    
    pedido = next((p for p in pedidos_cache if p.get("cliente") == cliente_nome and p.get("hora") == hora_completa), None)
    
    if not pedido:
        messagebox.showerror("Erro", "Pedido não encontrado no cache. Recarregue a lista.")
        return

    # 1. Monta o cupom
    cupom_texto = gerar_cupom(pedido)

    # 2. Envia para a impressora (em um thread para não travar a UI)
    def _enviar_impressao():
        try:
            print("Iniciando conexão com a impressora...")
            config = IMPRESSORA_CONFIG
            
            if config["tipo"] == "USB":
                print(f" Tentando conectar via USB: VID={hex(config['vid'])}, PID={hex(config['pid'])}")
                
                # 1. CORREÇÃO: Inicializa com codificação CP850 para suportar caracteres especiais
                p = Usb(config["vid"], config["pid"], profile=config.get("profile", "default"), encoding='cp850')
                
                print("Conexão USB estabelecida com sucesso!")
            elif config["tipo"] == "Network":
                print(f" Tentando conectar via Network: {config['host']}")
                p = Network(config["host"], profile=config.get("profile", "default"), encoding='cp850')
                print("Conexão Network estabelecida!")
            else:
                return False, "Tipo de impressora não configurado (USB ou Network)."

            print("Enviando cupom...")
            
            # --- FLUXO DE IMPRESSÃO CORRIGIDO E FINALIZADO ---
            
            # 2. Comando ESC/POS Bruto para CodePage 858 (ID 19)
            # Esta CodePage é a mais provável de resolver o problema do 'ã' e 'õ'.
            p._raw(b'\x1b\x74\x13') 
            
            # Imprime cabeçalho
            p.text("\n========== PIZZARIA SAVINO ==========\n")
            
            # 3. Imprime o texto REAL do cupom
            # Não use p.set(codepage) aqui, pois o p._raw(b'\x1b\x74\x13') já fez o trabalho.
            p.text(cupom_texto) 
            
            # 4. Garante linhas vazias e o corte
            p.text("\n\n\n") # Adicionei mais um \n para mais espaço antes do corte.
            p.cut() 
            
            print("Impressão concluída com sucesso!")
            return True, "Impressão enviada com sucesso!"

        except OSError as e:
            # O bloco 'except' permanece o mesmo para tratamento de erros
            print(f"ERRO DE SISTEMA: {e}")
            return False, (
                f"ERRO: Impressora não encontrada ou problema de comunicação!\n\n"
                f"IDs configurados:\n"
                f"  VID: {hex(IMPRESSORA_CONFIG.get('vid', 0))}\n"
                f"  PID: {hex(IMPRESSORA_CONFIG.get('pid', 0))}\n\n"
                f"Possíveis causas:\n"
                f"  • Impressora desligada ou desconectada\n"
                f"  • Driver USB não instalado (Zadig)\n"
                f"  • Permissões insuficientes (tente executar como administrador)\n\n"
                f"Erro técnico: {e}"
            )
        except IOError as e:
            # ... (Mensagem de erro mantida)
            return False, f"Erro de comunicação com a impressora: {e}\n\nVerifique se a impressora está ligada e conectada."
        except Exception as e:
            # ... (Mensagem de erro mantida)
            return False, f"Erro desconhecido na impressão: {e}"
    # Executa a impressão em thread
    future = executor.submit(_enviar_impressao)

    # Função de verificação (mantida)
    def verificar_impressao(future_task):
        if future_task.done():
            try:
                sucesso, msg = future_task.result()
                if sucesso:
                    messagebox.showinfo("Impressão", msg)
                else:
                    messagebox.showerror("Erro de Impressão", msg)
            except Exception as e:
                 messagebox.showerror("Erro de Thread", f"Ocorreu um erro ao finalizar a impressão: {e}")
        else:
            # Passa 'future_task' na chamada recursiva
            janela.after(100, lambda: verificar_impressao(future_task))

    janela.after(100, lambda: verificar_impressao(future))

def _finalizar_edicao_em_thread(dados_do_pedido, hora_original):
    """Roda a edição no thread de fundo e atualiza a interface no thread principal."""
    sucesso = salvar_pedido(dados_do_pedido, modo_edicao=True, hora_original=hora_original)
    if sucesso:
        janela.after(0, atualizar_lista)
        janela.after(0, atualizar_lista_arquivados) 
        janela.after(0, atualizar_lista_clientes)
        janela.after(0, gerar_relatorio_geral)
    else:
        janela.after(0, lambda: messagebox.showerror("Erro de Salvamento", "Falha ao salvar a edição do pedido no Excel."))

def _modificar_status_pedido(tabela_widget, novo_status):
    """Função genérica para alterar o status do pedido."""
    try:
        selected_item = tabela_widget.focus()
        if not selected_item:
            messagebox.showwarning("Atenção", "Selecione um pedido na tabela primeiro.")
            return

        # O valor da coluna 4 (hora) na tabela AGORA é a chave completa, garantindo a unicidade
        hora_original = tabela_widget.item(selected_item, 'values')[4] 
        # O cliente é a primeira coluna na tabela (índice 0)
        cliente_nome = tabela_widget.item(selected_item, 'values')[0]
        
        # CORRIGIDO: Busca com Cliente + Hora para garantir a unicidade
        pedido = next((p for p in pedidos_cache if str(p.get('hora')) == str(hora_original) and p.get('cliente') == cliente_nome), None)

        if not pedido:
            messagebox.showerror("Erro", "Pedido não encontrado no cache. Recarregue a lista.")
            return

        if novo_status == "Entregue" and not messagebox.askyesno("Confirmar Entrega", f"Marcar o pedido de {pedido['cliente']} como ENTREGUE e arquivar?"):
            return

        pedido['status'] = novo_status

        executor.submit(_finalizar_edicao_em_thread, pedido, hora_original)
        
        if novo_status == "Entregue":
             messagebox.showinfo("Sucesso", f"Pedido de {pedido['cliente']} arquivado com sucesso!")
        elif novo_status == "Em Preparação":
             messagebox.showinfo("Sucesso", f"Pedido de {pedido['cliente']} movido para 'Em Preparação'.")
        elif novo_status == "Saiu pra entrega":
             messagebox.showinfo("Sucesso", f"Pedido de {pedido['cliente']} saiu para entrega.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao modificar o status: {e}")

def marcar_entregue():
    """Chama a função de modificação para marcar como 'Entregue'."""
    global tabela_ativos
    _modificar_status_pedido(tabela_ativos, "Entregue")

def marcar_preparacao():
    """Chama a função de modificação para marcar como 'Em Preparação'."""
    global tabela_ativos
    _modificar_status_pedido(tabela_ativos, "Em Preparação")

def marcar_saiu_entrega():
    """Chama a função de modificação para marcar como 'Saiu pra entrega'."""
    global tabela_ativos
    _modificar_status_pedido(tabela_ativos, "Saiu pra entrega")

def reabrir_pedido_arquivado():
    """Reabre um pedido com status 'Entregue' para 'Pendente'."""
    global tabela_arquivados
    
    try:
        selected_item = tabela_arquivados.focus()
        if not selected_item:
            messagebox.showwarning("Atenção", "Selecione um pedido na tabela de arquivados primeiro.")
            return

        hora_original = tabela_arquivados.item(selected_item, 'values')[4] 
        cliente_nome = tabela_arquivados.item(selected_item, 'values')[0]
        
        pedido = next((p for p in pedidos_cache if str(p.get('hora')) == str(hora_original) and p.get('cliente') == cliente_nome), None)

        if not pedido:
            messagebox.showerror("Erro", "Pedido não encontrado. Recarregue a lista.")
            return
            
        if not messagebox.askyesno("Confirmar Reabertura", f"Reabrir o pedido de {pedido['cliente']} para 'Pendente'?"):
            return

        pedido['status'] = "Pendente"

        executor.submit(_finalizar_edicao_em_thread, pedido, hora_original)
        messagebox.showinfo("Sucesso", f"Pedido de {pedido['cliente']} reaberto com sucesso!")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao reabrir o pedido: {e}")

def mostrar_detalhes(event):
    """Mostra os detalhes de um pedido do cache (chamada por clique duplo)."""
    widget_clicado = event.widget 
    
    if widget_clicado == tabela_ativos:
        tabela_atual = tabela_ativos
    elif widget_clicado == tabela_arquivados:
        tabela_atual = tabela_arquivados
    else:
        return 

    item_sel = tabela_atual.focus()
    if not item_sel: return
    
    valores = tabela_atual.item(item_sel)["values"]
    cliente = valores[0]
    hora = valores[4]
    
    global pedidos_cache
    pedido = next((p for p in pedidos_cache if p.get("cliente") == cliente and p.get("hora") == hora), None)
    if not pedido: return

    janela_det = ttkb.Toplevel(janela)
    janela_det.title(f"Detalhes do Pedido - {pedido.get('cliente')}")
    janela_det.geometry("450x650") 
    janela_det.resizable(False, False)

    canvas = Canvas(janela_det)
    scroll_y = ttk.Scrollbar(janela_det, orient="vertical", command=canvas.yview)
    scroll_y.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.configure(yscrollcommand=scroll_y.set)

    frame_conteudo = ttkb.Frame(canvas)
    canvas.create_window((0, 0), window=frame_conteudo, anchor="nw")
    frame_conteudo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    ttkb.Label(frame_conteudo, text=f"🧍 Cliente: {pedido.get('cliente', 'N/A')}", font=("Arial", 12, "bold")).pack(pady=5, anchor="w", padx=10)
    ttkb.Label(frame_conteudo, text="🍕 Itens do Pedido:", font=("Arial", 11, "bold")).pack(pady=5, anchor="w", padx=10)
    
    detalhes_frame = ttk.Frame(frame_conteudo, padding=10) 
    detalhes_frame.pack(fill="x", padx=10, pady=5)
    
    for i, item in enumerate(pedido.get('itens', [])):
        preco = item.get('valor', 0.0) 
        ttkb.Label(detalhes_frame, text=f"{i+1}. {item.get('descricao', 'Item Sem Descrição')} (R$ {preco:.2f})", font=("Arial", 10, "bold")).pack(anchor="w")
        ttkb.Label(detalhes_frame, text=f"  Borda: {item.get('borda', 'Não')} ({item.get('sabor_borda','-')})", font=("Arial", 10)).pack(anchor="w")
        adicionais_str = ', '.join(item.get('adicionais', [])) or '-'
        ttkb.Label(detalhes_frame, text=f"  Adicionais: {adicionais_str}", font=("Arial", 10), wraplength=400, justify="left").pack(anchor="w")
        
        observacoes = item.get('observacoes', '-').strip()
        if observacoes and observacoes != '':
             ttk.Label(detalhes_frame, text=f"  ⚠️ Obs: {observacoes}", style="Danger.TLabel", font=("Arial", 10, "italic"), wraplength=400, justify="left").pack(anchor="w")
        else:
             ttkb.Label(detalhes_frame, text="  Obs: -", font=("Arial", 10)).pack(anchor="w")

    ttkb.Label(frame_conteudo, text="---").pack(fill="x", padx=10)
    
    ttkb.Label(frame_conteudo, text=f"💳 Pagamento: {pedido.get('forma_pagamento', 'N/A')}", font=("Arial", 11)).pack(pady=5, anchor="w", padx=10)

    if pedido.get("forma_pagamento") == "Dinheiro":
        troco_info = "Não precisa de troco"
        if pedido.get("precisa_troco") == True:
            troco_info = f"Troco pra: R$ {pedido.get('troco_valor', '0,00')}"
        ttkb.Label(frame_conteudo, text=f"💰 {troco_info}", font=("Arial", 11)).pack(pady=5, anchor="w", padx=10)

    ttkb.Label(frame_conteudo, text=f"📍 Endereço: {pedido.get('endereco', 'N/A')}", font=("Arial", 11), wraplength=420, justify="left").pack(pady=5, anchor="w", padx=10)
    taxa_entrega = pedido.get('taxa_entrega', 0.0)
    if isinstance(taxa_entrega, str): 
          try: taxa_entrega = float(taxa_entrega)
          except ValueError: taxa_entrega = 0.0
          
    ttkb.Label(frame_conteudo, text=f"Taxa de entrega: R$ {taxa_entrega:.2f}", font=("Arial", 11)).pack(pady=5, anchor="w", padx=10)
    
    ttkb.Label(frame_conteudo, text=f"🕒 Hora: {pedido.get('hora', 'N/A')}", font=("Arial", 11)).pack(pady=5, anchor="w", padx=10)
    
    valor_total_exibicao = pedido.get('valor_total', 0.00) 
    ttk.Label(frame_conteudo, text=f"💵 Valor Total: R$ {valor_total_exibicao:.2f}".replace('.', ','), font=("Arial", 12, "bold"), style="Primary.TLabel").pack(pady=10, anchor="w", padx=10)
    
    ttkb.Label(frame_conteudo, text=f"🚦 Status: {pedido.get('status', 'N/A')}", font=("Arial", 11, "bold")).pack(pady=5, anchor="w", padx=10)


def salvar_pedido_principal(janela_pedido, var_cliente, entry_endereco, combo_pagamento, var_troco_check, entry_troco, var_taxa, itens_pedido, modo_edicao, pedido_existente, var_telefone):
    """Prepara os dados e inicia o salvamento assíncrono."""
    cliente = var_cliente.get().strip()
    endereco = entry_endereco.get().strip()
    forma_pagamento = combo_pagamento.get().strip()
    telefone = var_telefone.get().strip()
    
    if not cliente and not telefone: messagebox.showerror("Erro", "O nome do Cliente ou Telefone é obrigatório."); return
    if not itens_pedido: messagebox.showerror("Erro", "O pedido deve ter pelo menos um item."); return

    # Validação do Troco
    precisa_troco = var_troco_check.get()
    if precisa_troco and forma_pagamento == "Dinheiro":
        try:
            troco_valor = float(entry_troco.get().replace(',', '.'))
            if troco_valor <= 0: 
                messagebox.showwarning("Erro", "Valor do troco deve ser positivo."); return
        except ValueError:
            messagebox.showwarning("Erro", "Valor do troco inválido."); return
    else:
         troco_valor = 0.0

    # Validação da Taxa
    try:
        taxa_entrega = float(var_taxa.get().replace(',', '.'))
    except ValueError:
        messagebox.showwarning("Erro", "Valor da taxa de entrega inválido."); return
    
    # Recalcula o valor total para garantir
    valor_total = sum(item['valor'] for item in itens_pedido) + taxa_entrega
    
    # CORRIGIDO: Data salva no formato DD/MM HH:MM:SS
    hora_original_edicao = pedido_existente.get('hora', datetime.now().strftime("%d/%m %H:%M:%S"))
    
    dados_do_pedido = {
        'cliente': cliente,
        'telefone': telefone, 
        'endereco': endereco,
        'forma_pagamento': forma_pagamento,
        'precisa_troco': precisa_troco,
        'troco_valor': troco_valor,
        'taxa_entrega': taxa_entrega,
        'status': pedido_existente.get('status', 'Pendente') if modo_edicao else 'Pendente', 
        'valor_total': valor_total,
        'hora': hora_original_edicao,
        'itens': itens_pedido
    }
    
    # Roda o salvamento em thread de fundo
    future = executor.submit(salvar_pedido, dados_do_pedido, modo_edicao, hora_original_edicao)
    
    def verificar_salvamento_e_fechar(future_task):
        if future_task.done():
            if future_task.result():
                janela_pedido.destroy()
                messagebox.showinfo("Sucesso", f"Pedido de {cliente} {'editado' if modo_edicao else 'criado'} com sucesso!")
                atualizar_lista()
                atualizar_lista_arquivados()
                atualizar_lista_clientes() 
                gerar_relatorio_geral()
            else:
                messagebox.showerror("Erro", "Falha ao salvar o pedido no Excel.")
        else:
            janela_pedido.after(100, lambda: verificar_salvamento_e_fechar(future_task))

    janela_pedido.after(100, lambda: verificar_salvamento_e_fechar(future))


def abrir_janela_pedido(modo_edicao=False, pedido_existente=None):
    """
    Função principal que constrói a interface de criação/edição de pedido.
    """
    
    pedido_existente = pedido_existente if pedido_existente is not None else {}
    itens_do_pedido = [item.copy() for item in pedido_existente.get('itens', [])]
    
    # Variáveis de controle
    var_cliente = ttkb.StringVar(value=pedido_existente.get('cliente', ''))
    var_telefone = ttkb.StringVar(value=pedido_existente.get('telefone', ''))
    var_endereco = ttkb.StringVar(value=pedido_existente.get('endereco', ''))
    var_pagamento = ttkb.StringVar(value=pedido_existente.get('forma_pagamento', 'Dinheiro'))
    var_troco_check = ttkb.BooleanVar(value=pedido_existente.get('precisa_troco', False))
    var_troco = ttkb.StringVar(value=str(pedido_existente.get('troco_valor', 0.0)).replace('.', ','))
    
    # Valor padrão da taxa de entrega para 0,00
    var_taxa = ttkb.StringVar(value=str(pedido_existente.get('taxa_entrega', 0.0)).replace('.', ',') if modo_edicao else '0,00')
    
    # Variáveis de controle para a pizza temporária
    var_sabor_1 = ttkb.StringVar()
    var_sabor_2 = ttkb.StringVar()
    var_borda = ttkb.StringVar(value="Não")
    var_adicional = ttkb.StringVar()
    
    # Variáveis para Labels que precisam de acesso nonlocal
    total_label = ttkb.Label() 
    total_itens_label = ttkb.Label()
    pizza_atual_label = ttkb.Label()
    carrinho_listbox = Listbox()
    sabores_listbox = Listbox()
    adicionais_listbox = Listbox()
    frame_sabor_borda = ttkb.Frame()
    text_observacoes = Text()
    frame_troco = ttkb.Frame()
    frame_troco_valor = ttkb.Frame()


    # ------------------ FUNÇÕES AUXILIARES (COM NONLOCAL) ------------------
    
    def calcular_preco_pizza_atual():
        sabores_selecionados = sabores_listbox.get(0, ttkb.END)
        preco_base = calcular_preco_pizza(sabores_selecionados)
        preco_borda = precos_bordas.get(var_borda.get(), 0)
        
        adicionais_selecionados = adicionais_listbox.get(0, ttkb.END)
        preco_adicionais = sum(precos_adicionais.get(ad, 0) for ad in adicionais_selecionados)
        
        return preco_base + preco_borda + preco_adicionais

    def atualizar_total_pedido(*args):
        nonlocal total_itens_label, total_label
        total_itens = sum(item['valor'] for item in itens_do_pedido)
        try: 
            total_taxa = float(var_taxa.get().replace(",", "."))
        except ValueError: 
            total_taxa = 0
            
        total_label.config(text=f"Valor Total do Pedido: R$ {total_itens + total_taxa:.2f}".replace('.',','))
        total_itens_label.config(text=f"Subtotal dos Itens: R$ {total_itens:.2f}".replace('.',','))
        return total_itens + total_taxa

    def atualizar_total_pizza_atual(*args):
        nonlocal pizza_atual_label
        total = calcular_preco_pizza_atual()
        pizza_atual_label.config(text=f"Preço da Pizza Atual: R$ {total:.2f}".replace('.',','))

    def preencher_por_cliente(chave_selecionada):
        """Preenche Endereço, Nome ou Telefone ao selecionar no Autocomplete."""
        nonlocal entry_endereco, entry_cliente, entry_telefone, var_cliente, var_telefone, var_endereco
        
        info = None
        
        # 1. Tenta busca direta (pela chave que foi usada no Autocomplete)
        info = historico_clientes.get(chave_selecionada)

        # 2. Se a busca direta falhar, itera sobre os valores para ver se o Nome ou Telefone corresponde
        if not info:
            for dados in historico_clientes.values():
                if dados['cliente'] == chave_selecionada or dados['telefone'] == chave_selecionada:
                    info = dados
                    break
        
        if info:
            entry_cliente.delete(0, ttkb.END)
            entry_telefone.delete(0, ttkb.END)
            entry_endereco.delete(0, ttkb.END)

            var_cliente.set(info['cliente'])
            var_telefone.set(info['telefone'])
            var_endereco.set(info['endereco'])

    def toggle_troco(*args):
        nonlocal frame_troco, frame_troco_valor
        if var_pagamento.get() == "Dinheiro": 
            frame_troco.grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=5) # Ajustado a linha
            toggle_valor_troco()
        else:
            frame_troco.grid_forget()
            frame_troco_valor.grid_forget()
            var_troco_check.set(False)

    def toggle_valor_troco():
        nonlocal frame_troco_valor
        if var_troco_check.get(): 
            frame_troco_valor.grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=2) # Ajustado a linha
        else: 
            frame_troco_valor.grid_forget()

    # FUNÇÕES QUE APENAS CONFIGURAM A LISTBOX TEMPORÁRIA
    def adicionar_sabor_temp():
        sabor = var_sabor_1.get()
        if not sabor or sabor not in precos_pizzas.keys(): return messagebox.showwarning("Aviso", "Sabor inválido.")
        
        sabores_atuais = sabores_listbox.get(0, ttkb.END)
        if len(sabores_atuais) >= 2: return messagebox.showwarning("Aviso", "Máximo de 2 sabores por pizza atingido.")
        
        if sabor not in sabores_atuais:
            sabores_listbox.insert(ttkb.END, sabor) 
        
        atualizar_total_pizza_atual()
    
    def remover_sabor_temp():
        selecionado = sabores_listbox.curselection()
        if not selecionado: return
        sabores_listbox.delete(selecionado[0])
        atualizar_total_pizza_atual()
    
    def adicionar_adicional_temp():
        ad = var_adicional.get()
        if not ad or ad not in precos_adicionais.keys(): return messagebox.showwarning("Aviso", "Adicional inválido.")
        
        adicionais_atuais = adicionais_listbox.get(0, ttkb.END)
        if ad not in adicionais_atuais:
            adicionais_listbox.insert(ttkb.END, ad) 
        
        atualizar_total_pizza_atual()
    
    def remover_adicional_temp():
        selecionado = adicionais_listbox.curselection()
        if not selecionado: return
        adicionais_listbox.delete(selecionado[0])
        atualizar_total_pizza_atual()

    # FUNÇÃO QUE MANDA PARA O CARRINHO PRINCIPAL (itens_do_pedido)
    def adicionar_item_ao_carrinho():
        sabores_selecionados = sabores_listbox.get(0, ttkb.END)
        if not sabores_selecionados:
             messagebox.showerror("Erro", "Configure a pizza (adicione o sabor) primeiro.")
             return
             
        adicionais_selecionados = adicionais_listbox.get(0, ttkb.END)
        
        preco_total = calcular_preco_pizza_atual()
        
        descricao = f"{' / '.join(sabores_selecionados)}"
        if var_borda.get() != "Não": descricao += f" | Borda: {var_borda.get()}"
        if adicionais_selecionados: descricao += f" | Add: {', '.join(adicionais_selecionados)}"
        
        itens_do_pedido.append({
            'tipo': 'Pizza',
            'descricao': descricao,
            'valor': preco_total,
            'adicionais': list(adicionais_selecionados), 
            'borda': var_borda.get(),
            'observacoes': text_observacoes.get("1.0", ttkb.END).strip()
        })
        
        # Limpa e reseta a área de configuração temporária
        limpar_configuracao_pizza() 
        
        atualizar_lista_carrinho()
        atualizar_total_pedido()
        messagebox.showinfo("Sucesso", "Item adicionado ao carrinho!")

    def remover_pizza_do_carrinho():
        selected_item = carrinho_listbox.curselection()
        if not selected_item:
            messagebox.showwarning("Atenção", "Selecione um item na lista do carrinho para remover.")
            return

        item_index = selected_item[0]
        
        if 0 <= item_index < len(itens_do_pedido):
            itens_do_pedido.pop(item_index)
            atualizar_lista_carrinho()
            atualizar_total_pedido()

    def limpar_configuracao_pizza():
        nonlocal sabores_listbox, adicionais_listbox, text_observacoes
        sabores_listbox.delete(0, ttkb.END) 
        var_sabor_1.set(''); var_sabor_2.set(''); var_borda.set('Não')
        adicionais_listbox.delete(0, ttkb.END) 
        var_adicional.set("")
        text_observacoes.delete("1.0", ttkb.END) 
        atualizar_total_pizza_atual()

    def atualizar_lista_carrinho():
        nonlocal carrinho_listbox
        carrinho_listbox.delete(0, ttkb.END) 
        for item in itens_do_pedido:
            preco = item.get('valor', 0.0)
            carrinho_listbox.insert(ttkb.END, f"R$ {preco:.2f}".replace('.', ',') + f" - {item.get('descricao', 'Item Desconhecido')}")

    
    # ------------------ CONSTRUÇÃO DA JANELA ------------------

    janela_add = ttkb.Toplevel(janela)
    titulo = "Editar Pedido" if modo_edicao else "Adicionar Pedido"
    janela_add.title(titulo)
    janela_add.geometry("1100x700") 
    janela_add.grid_columnconfigure(0, weight=1)
    janela_add.grid_rowconfigure(0, weight=1)

    main_frame = ttkb.Frame(janela_add)
    main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    main_frame.grid_columnconfigure(0, weight=1, minsize=500) 
    main_frame.grid_columnconfigure(1, weight=1, minsize=400) 

    # ---------- COLUNA ESQUERDA: Cliente e Configuração da Pizza ----------

    coluna_esquerda = ttkb.Frame(main_frame)
    coluna_esquerda.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    coluna_esquerda.grid_columnconfigure(0, weight=1) 

    # Frame de Dados do Cliente
    frame_cliente = ttk.LabelFrame(coluna_esquerda, text="Dados do Cliente e Entrega", padding=10)
    frame_cliente.grid(row=0, column=0, sticky="ew", pady=5)
    frame_cliente.grid_columnconfigure(1, weight=1) 
    
    # LISTAS FILTRADAS PARA AUTOCOMPLETE
    lista_telefones_busca = [d['telefone'] for d in historico_clientes.values() if d['telefone']]
    lista_nomes_busca = [d['cliente'] for d in historico_clientes.values() if d['cliente']]
    
    # Chave 1: Telefone
    ttkb.Label(frame_cliente, text="Telefone:").grid(row=0, column=0, sticky="w", padx=5)
    entry_telefone = AutocompleteEntry(frame_cliente, lista_telefones_busca, callback_on_select=preencher_por_cliente, textvariable=var_telefone, width=40)
    entry_telefone.grid(row=0, column=1, sticky="ew", pady=2, padx=5)

    # Chave 2: Nome
    ttkb.Label(frame_cliente, text="Cliente:").grid(row=1, column=0, sticky="w", padx=5)
    entry_cliente = AutocompleteEntry(frame_cliente, lista_nomes_busca, callback_on_select=preencher_por_cliente, textvariable=var_cliente, width=40)
    entry_cliente.grid(row=1, column=1, sticky="ew", pady=2, padx=5)

    ttkb.Label(frame_cliente, text="Endereço:").grid(row=2, column=0, sticky="w", padx=5)
    entry_endereco = ttk.Entry(frame_cliente, textvariable=var_endereco, width=40)
    entry_endereco.grid(row=2, column=1, sticky="ew", pady=2, padx=5)
    
    ttkb.Label(frame_cliente, text="Taxa de entrega (R$):").grid(row=3, column=0, sticky="w", padx=5, pady=2)
    entry_taxa = ttk.Entry(frame_cliente, textvariable=var_taxa, width=10)
    entry_taxa.grid(row=3, column=1, sticky="w", padx=5, pady=2)
    var_taxa.trace_add("write", atualizar_total_pedido)

    ttkb.Label(frame_cliente, text="Forma de Pagamento:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
    combo_pagamento = ttk.Combobox(frame_cliente, textvariable=var_pagamento, values=["Dinheiro", "Pix", "Cartao"], state="readonly", width=15)
    combo_pagamento.grid(row=4, column=1, sticky="w", padx=5, pady=2)
    combo_pagamento.bind("<<ComboboxSelected>>", toggle_troco)
    
    frame_troco = ttkb.Frame(frame_cliente)
    ttk.Checkbutton(frame_troco, text="Precisa de troco?", variable=var_troco_check, bootstyle="secondary", command=toggle_valor_troco).pack(anchor="w")
    
    frame_troco_valor = ttkb.Frame(frame_cliente)
    ttkb.Label(frame_troco_valor, text="Troco para R$:", ).pack(side="left")
    entry_troco = ttk.Entry(frame_troco_valor, textvariable=var_troco, width=10)
    entry_troco.pack(side="left")
    
    toggle_troco()


    # Frame de Configuração da Pizza
    frame_pizza = ttk.LabelFrame(coluna_esquerda, text="Configurar Pizza Atual", padding=10)
    frame_pizza.grid(row=1, column=0, sticky="nsew", pady=5)
    frame_pizza.grid_columnconfigure(0, weight=1)
    frame_pizza.grid_columnconfigure(1, weight=1)
    frame_pizza.grid_columnconfigure(2, weight=1)

    # --- Sabores ---

    ttkb.Label(frame_pizza, text="🍕 Sabores (Max: 2):", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", columnspan=3, pady=(5,0))
    
    frame_sabores_input = ttkb.Frame(frame_pizza)
    frame_sabores_input.grid(row=1, column=0, columnspan=3, sticky="ew")
    frame_sabores_input.grid_columnconfigure(0, weight=1)
    
    sabor_cb1 = ttk.Combobox(frame_sabores_input, textvariable=var_sabor_1, values=list(precos_pizzas.keys()), state="readonly", width=20)
    sabor_cb1.grid(row=0, column=0, sticky="ew")
    sabor_cb2 = ttk.Combobox(frame_sabores_input, textvariable=var_sabor_2, values=[""] + list(precos_pizzas.keys()), state="readonly", width=20)
    sabor_cb2.grid(row=0, column=1, sticky="ew")
    
    ttk.Button(frame_sabores_input, text="Add Sabor", command=adicionar_sabor_temp, bootstyle="info", width=8).grid(row=0, column=2, padx=5)
    
    sabores_listbox = Listbox(frame_pizza, height=2)
    sabores_listbox.grid(row=2, column=0, sticky="ew", columnspan=2, pady=5)
    ttk.Button(frame_pizza, text="Remover Sabor Selecionado", command=remover_sabor_temp, bootstyle="danger").grid(row=2, column=2, sticky="e")

    # --- Borda Recheada ---

    frame_borda_principal = ttkb.Frame(frame_pizza)
    frame_borda_principal.grid(row=3, column=0, columnspan=3, sticky="w", pady=5)
    
    ttkb.Label(frame_borda_principal, text="Borda Recheada?").grid(row=0, column=0, sticky="w")
    borda_cb = ttk.Combobox(frame_borda_principal, textvariable=var_borda, values=["Não"] + list(precos_bordas.keys()), state="readonly", width=10)
    borda_cb.grid(row=0, column=1, padx=5)
    borda_cb.bind("<<ComboboxSelected>>", atualizar_total_pizza_atual) 

    # --- Adicionais ---

    ttkb.Label(frame_pizza, text="📝 Adicionais:", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky="w", columnspan=3, pady=(5,0))
    frame_adicionais = ttkb.Frame(frame_pizza)
    frame_adicionais.grid(row=5, column=0, columnspan=3, sticky="ew")
    frame_adicionais.grid_columnconfigure(0, weight=1)
    
    adicionais_cb = ttk.Combobox(frame_adicionais, textvariable=var_adicional, state="readonly", values=list(precos_adicionais.keys()))
    adicionais_cb.grid(row=0, column=0, sticky="ew")
    ttk.Button(frame_adicionais, text="+", command=adicionar_adicional_temp, bootstyle="info", width=2).grid(row=0, column=1, padx=5)
    ttk.Button(frame_adicionais, text="-", command=remover_adicional_temp, bootstyle="danger", width=2).grid(row=0, column=2)
    
    adicionais_listbox = Listbox(frame_pizza, height=3)
    adicionais_listbox.grid(row=6, column=0, columnspan=3, sticky="ew", pady=5)
    
    # --- Observações ---

    ttkb.Label(frame_pizza, text="💬 Observações:", font=("Arial", 10, "bold")).grid(row=7, column=0, sticky="w", columnspan=3, pady=(5,0))
    text_observacoes = Text(frame_pizza, height=3, width=50)
    text_observacoes.grid(row=8, column=0, columnspan=3, sticky="ew", pady=5)
    
    pizza_atual_label = ttkb.Label(frame_pizza, text="Preço da Pizza Atual: R$ 0.00", font=("Arial", 12, "bold"))
    pizza_atual_label.grid(row=9, column=0, sticky="w", pady=10)
    
    # Botão final para mover da CONFIGURAÇÃO para o CARRINHO PRINCIPAL
    ttk.Button(frame_pizza, text="➕ Adicionar Item ao Carrinho", bootstyle="success", command=adicionar_item_ao_carrinho).grid(row=9, column=2, sticky="e", pady=10)

    # ---------- COLUNA DIREITA: Carrinho e Pagamento ----------

    coluna_direita = ttkb.Frame(main_frame)
    coluna_direita.grid(row=0, column=1, sticky="nsew")
    coluna_direita.grid_columnconfigure(0, weight=1)
    
    # Frame do Carrinho
    frame_carrinho = ttk.LabelFrame(coluna_direita, text="Itens do Pedido (Carrinho)", padding=10)
    frame_carrinho.grid(row=1, column=0, sticky="ew", pady=5)
    frame_carrinho.grid_columnconfigure(0, weight=1)
    
    carrinho_listbox = Listbox(frame_carrinho, height=12) 
    carrinho_listbox.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)
    
    frame_carrinho_botoes = ttkb.Frame(frame_carrinho)
    frame_carrinho_botoes.grid(row=1, column=0, columnspan=2, sticky="ew")
    frame_carrinho_botoes.grid_columnconfigure(1, weight=1)
    
    ttk.Button(frame_carrinho_botoes, text="Remover Item", command=remover_pizza_do_carrinho, bootstyle="danger").grid(row=0, column=0, sticky="w")
    
    total_itens_label = ttkb.Label(frame_carrinho_botoes, text="Subtotal dos Itens: R$ 0.00", font=("Arial", 11, "bold"))
    total_itens_label.grid(row=0, column=1, sticky="e")
    
    # Frame de Pagamento (Valores)
    frame_pagamento = ttk.LabelFrame(coluna_direita, text="Valores", padding=10)
    frame_pagamento.grid(row=2, column=0, sticky="ew", pady=5)
    frame_pagamento.grid_columnconfigure(0, weight=1)
    
    total_label = ttk.Label(frame_pagamento, text="Valor Total do Pedido: R$ 0.00", style="Primary.TLabel", font=("Arial", 14, "bold"))
    total_label.grid(row=0, column=0, sticky="w", pady=10)

    # NOVO STATUS DE SALVAMENTO
    status_label = ttkb.Label(coluna_direita, text="", font=("Arial", 10, "italic"))
    status_label.grid(row=3, column=0, sticky="ew", pady=(5, 0)) 
    texto_botao_salvar = "✅ Salvar Alterações no Pedido" if modo_edicao else "✅ Finalizar Pedido"
    btn_finalizar = ttk.Button(coluna_direita, text=texto_botao_salvar, bootstyle="success")
    btn_finalizar.grid(row=4, column=0, sticky="ew", pady=15)
    
    btn_finalizar.config(command=lambda: salvar_pedido_principal(
        janela_add, var_cliente, entry_endereco, combo_pagamento, 
        var_troco_check, entry_troco, var_taxa, itens_do_pedido, 
        modo_edicao, pedido_existente, var_telefone
    ))
    
    # Chamadas iniciais
    atualizar_total_pedido() 
    atualizar_lista_carrinho() 
    janela_add.grab_set() 
    janela_add.mainloop()


def abrir_janela_edicao(tabela_widget):
    """Função para abrir a janela de edição a partir de qualquer tabela."""
    selected_item = tabela_widget.focus()
    if not selected_item:
        messagebox.showwarning("Atenção", "Selecione um pedido na tabela primeiro.")
        return

    valores = tabela_widget.item(selected_item, 'values')
    hora_original = valores[4] 
    cliente_nome = valores[0]

    pedido = next((p for p in pedidos_cache if str(p.get('hora')) == str(hora_original) and p.get('cliente') == cliente_nome), None)

    if pedido:
        # Usa cópia limpa e profunda do pedido
        pedido_copia = json.loads(json.dumps(pedido)) 
        abrir_janela_pedido(modo_edicao=True, pedido_existente=pedido_copia)
    else:
        messagebox.showerror("Erro", "Pedido não encontrado. Recarregue a lista.")

def abrir_janela_edicao_cliente(tabela_widget):
    """Abre a janela para editar o nome, telefone e endereço de um cliente."""
    selected_item = tabela_widget.focus()
    if not selected_item:
        messagebox.showwarning("Atenção", "Selecione um cliente na tabela primeiro.")
        return

    valores = tabela_widget.item(selected_item)['values']
    
    # Valores atuais da tabela: (Nome, Telefone, Endereço, Pedidos)
    nome_atual = valores[0]
    telefone_atual = valores[1]
    endereco_atual = valores[2]
            
    janela_edit_cli = ttkb.Toplevel(janela)
    janela_edit_cli.title(f"Editar Dados: {nome_atual}")
    janela_edit_cli.geometry("400x300")
    
    frame = ttkb.Frame(janela_edit_cli, padding=20)
    frame.pack(fill="both", expand=True)
    
    var_nome = ttkb.StringVar(value=nome_atual)
    var_telefone = ttkb.StringVar(value=telefone_atual)
    var_endereco = ttkb.StringVar(value=endereco_atual)

    ttkb.Label(frame, text="Nome:").grid(row=0, column=0, sticky="w", pady=5)
    entry_nome = ttk.Entry(frame, textvariable=var_nome, width=30)
    entry_nome.grid(row=0, column=1, sticky="ew", padx=5)

    ttkb.Label(frame, text="Telefone:").grid(row=1, column=0, sticky="w", pady=5)
    entry_tel = ttk.Entry(frame, textvariable=var_telefone, width=30)
    entry_tel.grid(row=1, column=1, sticky="ew", padx=5)

    ttkb.Label(frame, text="Endereço:").grid(row=2, column=0, sticky="w", pady=5)
    entry_end = ttk.Entry(frame, textvariable=var_endereco, width=30)
    entry_end.grid(row=2, column=1, sticky="ew", padx=5)

    def salvar_edicao_cliente():
        novo_nome = var_nome.get().strip()
        novo_tel = var_telefone.get().strip()
        novo_end = var_endereco.get().strip()
        
        if not novo_nome and not novo_tel:
            messagebox.showerror("Erro", "Nome ou Telefone deve ser preenchido.")
            return

        # Busca todos os pedidos desse cliente
        pedidos_para_atualizar = [p for p in pedidos_cache if p['cliente'] == nome_atual or p['telefone'] == telefone_atual]

        if not pedidos_para_atualizar:
            messagebox.showerror("Erro", "Nenhum pedido encontrado para atualizar.")
            return

        # Para cada pedido, atualizamos as 3 colunas e salvamos.
        for pedido in pedidos_para_atualizar:
            pedido_editado = pedido.copy()
            pedido_editado['cliente'] = novo_nome
            pedido_editado['telefone'] = novo_tel
            pedido_editado['endereco'] = novo_end
            
            # Submete a atualização ao Excel (em thread)
            executor.submit(salvar_pedido, pedido_editado, modo_edicao=True, hora_original=pedido['hora'])

        # Fecha a janela e atualiza as listas
        janela_edit_cli.destroy()
        messagebox.showinfo("Sucesso", f"Dados do cliente '{novo_nome}' serão atualizados no Excel e na lista.")
        # Chamada de atualização para refletir as mudanças no UI
        janela.after(200, atualizar_lista_clientes) 
        janela.after(400, atualizar_lista) # Atualiza pedidos ativos também


    ttkb.Button(frame, text="Salvar Alterações", command=salvar_edicao_cliente, bootstyle="success").grid(row=3, column=0, columnspan=2, pady=20)
    janela_edit_cli.grab_set() 

# ---------------- FUNÇÕES DE CRIAÇÃO DE ABAS ----------------

def _criar_aba_pedidos_ativos(notebook_principal):
    """Cria a aba de Pedidos Ativos com botões de ação e Treeview."""
    global tabela_ativos
    
    frame_pedidos = ttkb.Frame(notebook_principal, padding=10)
    notebook_principal.add(frame_pedidos, text="Pedidos Ativos")

    # Frame para botões de ação e lista
    frame_acoes = ttkb.Frame(frame_pedidos)
    frame_acoes.pack(fill="x", pady=5)
    
    # Botões de Ação
    ttkb.Button(frame_acoes, text="Novo Pedido", command=lambda: abrir_janela_pedido(modo_edicao=False), bootstyle="success").pack(side="left", padx=5)
    ttkb.Button(frame_acoes, text="Editar Pedido", command=lambda: abrir_janela_edicao(tabela_ativos), bootstyle="warning").pack(side="left", padx=5)
    ttkb.Button(frame_acoes, text="Atualizar Lista (Recarregar)", command=atualizar_lista, bootstyle="light").pack(side="right", padx=5)
    
    # BOTÃO DE IMPRIMIR
    ttkb.Button(frame_acoes, text="🖨️ Imprimir Pedido", command=imprimir_pedido, bootstyle="secondary").pack(side="left", padx=5)

    # Botões de Status
    frame_status = ttkb.LabelFrame(frame_pedidos, text=" Atualizar Status ", padding=5)
    frame_status.pack(fill="x", pady=10)
    ttkb.Button(frame_status, text="Em Preparação", command=marcar_preparacao, bootstyle="info").pack(side="left", padx=5, expand=True)
    ttkb.Button(frame_status, text="Saiu p/ Entrega", command=marcar_saiu_entrega, bootstyle="primary").pack(side="left", padx=5, expand=True)
    ttkb.Button(frame_status, text="Marcar como Entregue (Arquivar)", command=marcar_entregue, bootstyle="success-outline").pack(side="left", padx=5, expand=True)

    # Tabela de Pedidos Ativos (Treeview)
    tabela_ativos = ttkb.Treeview(frame_pedidos, columns=('cliente', 'descricao', 'valor', 'status', 'hora'), show='headings', height=15)
    tabela_ativos.heading('cliente', text='Cliente'); tabela_ativos.column('cliente', width=120)
    tabela_ativos.heading('descricao', text='Descrição'); tabela_ativos.column('descricao', width=200)
    tabela_ativos.heading('valor', text='Total'); tabela_ativos.column('valor', width=80, anchor=ttkb.E)
    tabela_ativos.heading('status', text='Status'); tabela_ativos.column('status', width=120)
    tabela_ativos.heading('hora', text='Hora'); tabela_ativos.column('hora', width=120, anchor=ttkb.E)
    
    # Tags de cor para status (CORES MAIS DISCRETAS)
    tabela_ativos.tag_configure('preparacao', background='', foreground='orange') 
    tabela_ativos.tag_configure('entrega', background='', foreground='skyblue')   
    tabela_ativos.tag_configure('pendente', background='', foreground='#A9A9A9')  # Cinza discreto

    tabela_ativos.pack(fill="both", expand=True)
    
    # Configura o evento de clique duplo para mostrar detalhes
    tabela_ativos.bind("<Double-1>", mostrar_detalhes) 

    # Carrega a lista ao iniciar
    atualizar_lista()

def _criar_aba_pedidos_arquivados(notebook_principal):
    """Cria a aba de Pedidos Arquivados."""
    global tabela_arquivados
    
    frame_arquivados = ttkb.Frame(notebook_principal, padding=10)
    notebook_principal.add(frame_arquivados, text="Pedidos Arquivados")
    
    # Frame para botões de ação e lista
    frame_acoes = ttkb.Frame(frame_arquivados)
    frame_acoes.pack(fill="x", pady=5)
    
    # Botão de Reabrir/Desarquivar
    ttkb.Button(frame_acoes, text="Reabrir Pedido (Mover para Ativos)", command=reabrir_pedido_arquivado, bootstyle="danger-outline").pack(side="left", padx=5)
    ttkb.Button(frame_acoes, text="Atualizar Lista (Recarregar)", command=atualizar_lista_arquivados, bootstyle="light").pack(side="right", padx=5)

    # Tabela de Pedidos Arquivados (Treeview)
    tabela_arquivados = ttkb.Treeview(frame_arquivados, columns=('cliente', 'descricao', 'valor', 'status', 'hora'), show='headings', height=15)
    tabela_arquivados.heading('cliente', text='Cliente'); tabela_arquivados.column('cliente', width=120)
    tabela_arquivados.heading('descricao', text='Descrição'); tabela_arquivados.column('descricao', width=200)
    tabela_arquivados.heading('valor', text='Total'); tabela_arquivados.column('valor', width=80, anchor=ttkb.E)
    tabela_arquivados.heading('status', text='Status'); tabela_arquivados.column('status', width=120)
    tabela_arquivados.heading('hora', text='Hora'); tabela_arquivados.column('hora', width=120, anchor=ttkb.E)
    
    tabela_arquivados.tag_configure('entregue', foreground='gray', background='#E0E0E0')

    tabela_arquivados.pack(fill="both", expand=True)
    
    # Configura o evento de clique duplo para mostrar detalhes
    tabela_arquivados.bind("<Double-1>", mostrar_detalhes) 

    # Carrega a lista ao iniciar
    atualizar_lista_arquivados()

def _criar_aba_clientes(notebook_principal):
    """Cria a aba de Fidelidade/Clientes."""
    global tabela_clientes
    
    frame_clientes = ttkb.Frame(notebook_principal, padding=10)
    notebook_principal.add(frame_clientes, text="Clientes")
    
    ttkb.Label(frame_clientes, text="Histórico e Fidelidade", font=("Arial", 16, "bold")).pack(pady=10)
    
    # Tabela de Clientes
    tabela_clientes = ttkb.Treeview(frame_clientes, columns=('nome', 'telefone', 'endereco', 'pedidos'), show='headings', height=15)
    # ORDEM CORRIGIDA: Nome, Telefone, Endereço, Pedidos
    tabela_clientes.heading('nome', text='Nome'); tabela_clientes.column('nome', width=150)
    tabela_clientes.heading('telefone', text='Telefone'); tabela_clientes.column('telefone', width=100)
    tabela_clientes.heading('endereco', text='Último Endereço'); tabela_clientes.column('endereco', width=300)
    tabela_clientes.heading('pedidos', text='Total Pedidos'); tabela_clientes.column('pedidos', width=100, anchor=ttkb.E)
    
    tabela_clientes.pack(fill="both", expand=True)
    
    frame_clientes_botoes = ttkb.Frame(frame_clientes)
    frame_clientes_botoes.pack(pady=10, fill="x")
    
    ttkb.Button(frame_clientes_botoes, text="Atualizar Lista de Clientes", command=atualizar_lista_clientes, bootstyle="info").pack(side="left", padx=5)
    ttkb.Button(frame_clientes_botoes, text="Editar Cliente", command=lambda: abrir_janela_edicao_cliente(tabela_clientes), bootstyle="warning").pack(side="left", padx=5)
    
    atualizar_lista_clientes()


def _criar_aba_gerenciamento(notebook_principal):
    """Cria a aba de Gerenciamento."""
    global relatorio_label
    
    frame_gerenciamento = ttkb.Frame(notebook_principal, padding=10)
    notebook_principal.add(frame_gerenciamento, text="Gerenciamento")

    # --- Relatórios ---

    frame_relatorios = ttkb.LabelFrame(frame_gerenciamento, text=" Relatório Geral ", padding=10)
    frame_relatorios.pack(fill="x", pady=10)
    
    relatorio_label = ttkb.Label(frame_relatorios, text="Carregando resumo...", justify="left", font=('Courier', 10))
    relatorio_label.pack(fill="x", padx=5, pady=5)
    
    ttkb.Button(frame_relatorios, text="Gerar Resumo Geral (Todos os Pedidos)", command=gerar_relatorio_geral, bootstyle="info").pack(fill="x", pady=5)
    
    # Botão de Limpeza
    frame_limpeza = ttkb.Frame(frame_gerenciamento)
    frame_limpeza.pack(fill="x", pady=10)
    ttkb.Button(frame_limpeza, text="Limpar TODOS os Dados da Planilha (CUIDADO)", command=limpar_dados_planilha, bootstyle="danger").pack(fill="x", pady=5)


    # Gera o relatório inicial
    gerar_relatorio_geral()


def iniciar_aplicacao():
    """Inicia a aplicação principal."""
    global janela
    janela = ttkb.Window(themename="superhero") 
    janela.title("Sistema de Pedidos PIZZARIA v1.0")
    janela.geometry("800x600")

    notebook_principal = ttkb.Notebook(janela)
    notebook_principal.pack(pady=10, padx=10, fill="both", expand=True)

    _criar_aba_pedidos_ativos(notebook_principal)
    _criar_aba_pedidos_arquivados(notebook_principal) 
    _criar_aba_clientes(notebook_principal) 
    _criar_aba_gerenciamento(notebook_principal)

    janela.mainloop()

if __name__ == "__main__":
    iniciar_aplicacao()
