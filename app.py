# APLICATIVO QUE SE CONECTA NO BANCO DE DADOS DO EVOLUTI
# TIPOS DE MENSAGENS E OPÇOES DE MARKTING PARA SER DISPARADO PARA OS CLIENTES OU PROSPECTS POR DATA NASCIMENTO
mensagens = {
    "CLIENTE de ANIVERSÁRIO - MODELO 1": (
        ", {nome}, da empresa {empresa}, "
        "*Parabéns pelo seu dia!* "
        "Caro cliente, neste dia especial em que você comemora mais um ano da sua vida, gostaríamos de lhe prestar esta pequena, "
        "mas sincera, *homenagem*, desejando-lhe para o efeito, um *feliz aniversário*. Esperamos que passe um dia cheio de alegria, " 
        "surpresas boas, que possa desfrutar do seu dia na companhia de amigos e familiares, e que conte muitos anos. "
        "Esperamos também poder continuar a contar com a sua preferência, pois sem você não seríamos nada do que somos. "
        "Parabéns, caro cliente, muitas felicidades, hoje e sempre! Muito amor, muita paz e muita saúde para a sua vida que desejamos longa e feliz!. "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "CLIENTE de ANIVERSÁRIO - MODELO 2": (
        ", {nome}, da empresa {empresa}, "
        "Que este dia seja muito feliz e maravilho para você, no dia mais importante de sua Vida! "
        "Desejamos um feliz Aniversário, com muita disposição e principalmente ouvindo toda a festa, "
        "E também se precisar de algo, basta nos solicitar a nossa ajuda, aqui nesse Whats. "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "PROSPECTS BUSCANDO DISPERTAR INTERESSE": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *AUDIPLUS* trabalhamos com *Aparelhos Modernos e preços Imbátiveis*. "
        "Montamos o valor baseada no equipamento mais adequado para sua situação. "
        "Valores que cabem no seu bolso. Peça mais informações conosco. "
        "Se não quiser mais receber informações sobre nossos serviços, envie *SAIR*."
    ),
    "PROSPECT QUE FOI ATÉ ORÇADO, EM NEGOCIAÇÃO": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *AUDIPLUS* trabalhamos com um *APARELHOS AUDITIVOS*. "
        "E como já tinhamos conversado anteriormente, chegamos a falar um pouco sobre o eles, e até foi orçado, " 
        "e agora temos *Novidades* e *Promoções* exclusivas para você. ""Entre em contato conosco e comprove! "
        "Mas Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "EMPRESAS PARCEIRAS QUE NOS INDICAM": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *Audiplus* reafirmamos nossa parceria e temos *Novidades* e *Promoções* "
        "exclusivas para seus Clientes nesse novo ciclo. Aparelhos de ultima Geração, e com um preço muito especial. "
        "Entre em contato conosco, para mais informações."
        "Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "CLIENTES QUE USAM MUITO POUCO": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Aparelho*? "
        "Sabemos que utiliza somente em algumas ocasiões. "
        "Mas estamos prontos para ajudar, caso precisar de algo. Basta falar com o nosso suport Whats no *(55)3333-4650*. "
        "Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
    "CLIENTES INADINPLENTES": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Aparelho*, tudo certo? Algo que deseja mencionar? "
        "Nosso produto atende a sua nescessidade com oque é de mais moderno, em relação a tecnologia, E Caso precise de ajuda, na utilização do mesmo, basta falar com o nosso Whats no *(55)3333-4650*. "
        "Mas Gostaríamos de salientar aqui, a importância de manter seu pagamento em dia. "
        "Evite Transtorno no retorno para o periódico,  pois o não pagamento em DIA, ocasiona o *Bloqueio* do seu atendimento presencial. " 
        "Não deixe para depois, solicite a *2 via do boleto* aqui nesse Whats, caso não tenha recebido ainda."
    ),
    
}
# IMPORTANDOS BIBLIOTECAS UTILIZADAS
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import webbrowser
import pyautogui
import time
import threading
import re
from urllib.parse import quote
import os
import winsound
import mysql.connector
from tkinter import PhotoImage


## DEFININDO A CLASSE
class WhatsAppSenderApp:
    def __init__(self, root):
        self.root = root

        # ÍCONE PERSONALIZADO
        icon = PhotoImage(file="icone.png")
        self.root.iconphoto(False, icon)

        # CABEÇALHO APLICATIVO
        self.root.title("APLICATIVO Aniversários v_Audiplus")
        self.root.geometry("880x685+100+5")

        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica", 8, "bold"))
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))

        self.enviando = False
        self.envio_ativo = threading.Event()

        self.df = None
        self.filtered_df = None

        frame_top = ttk.Frame(root, padding=10)
        frame_top.pack(fill='x')

        #ttk.Button(frame_top, text="Conectar Banco", command=self.carregar_dados_do_mysql).pack(side='left')

        ttk.Label(
            frame_top,
            text="Disparo Automático Via WhatsApp 'Businsess' a cada 2 MINUTOS!",
            font=("TkDefaultFont", 10, "bold")
        ).pack(side='left', padx=(5, 2))
        ttk.Label(
            frame_top,
            text="Usar 1 CHIP (NUMERO) especifico,\nsomente PARA ENVIO dessas Mensagens.",
            foreground="green",
            font=("TkDefaultFont", 10, "bold")
        ).pack(padx=5, pady=2)

        ttk.Label(
            frame_top,
            text="Como não estamos usando a API do WhatsApp...\nSeu número pode ser BLOQUEADO!",
            foreground="red",
            font=("TkDefaultFont", 10, "bold")
        ).pack(padx=5, pady=2)
        
        frame_middle = ttk.Frame(root, padding=10)
        frame_middle.pack(fill='x')

        tk.Label(frame_middle, text="Selecione o Tipo de Escrita da Mensagem:", font=("TkDefaultFont", 10, "bold")).pack(side='left')
        self.combo_msg_type = ttk.Combobox(frame_middle, state='readonly', width=50)
        self.combo_msg_type.pack(side='left', padx=5)

        self.combo_msg_type['values'] = sorted(mensagens.keys())
        self.combo_msg_type.bind("<<ComboboxSelected>>", self.edit_template)

        MESES = [
            ("01", "JANEIRO"), ("02", "FEVEREIRO"), ("03", "MARÇO"), ("04", "ABRIL"),
            ("05", "MAIO"), ("06", "JUNHO"), ("07", "JULHO"), ("08", "AGOSTO"),
            ("09", "SETEMBRO"), ("10", "OUTUBRO"), ("11", "NOVEMBRO"), ("12", "DEZEMBRO")
        ]

        DIAS = [f"{i:02d}" for i in range(1, 32)]  # Dias de "01" a "31"

# Filtro de Mês
        self.label_mes = ttk.Label(root, text="MÊS do ANIVERSÁRIO:")
        self.label_mes.pack()

        self.combo_mes = ttk.Combobox(root, values=[f"{num} - {nome}" for num, nome in MESES])
        self.combo_mes.pack()
        self.combo_mes.set("Selecione") 

# Filtro de Dia
        self.label_dia = ttk.Label(root, text="DIA do ANIVERSÁRIO:")
        self.label_dia.pack()

        self.combo_dia = ttk.Combobox(root, values=["Todos os DIAS"] + DIAS)
        self.combo_dia.pack()
        self.combo_dia.current(0)  # "Todos"


        ttk.Label(frame_middle, text="Categoria:").pack(side='left', padx=(20, 5))
        self.combo_group = ttk.Combobox(frame_middle, state='readonly', width=20)
        self.combo_group.pack(side='left', padx=5)

        frame_edit = ttk.Frame(root, padding=20)
        frame_edit.pack(fill='x')

        tk.Label(
            frame_edit,
            text="Modelo da Mensagem a ser Enviada: PODE EDITAR ANTES DE CARREGAR OS CLIENTES",
            font=("Arial", 12, "bold"),
            anchor='w'
        ).pack(fill='x', padx=5, pady=(0, 5))

        self.txt_template = tk.Text(frame_edit, height=5, wrap='word')
        self.txt_template.pack(fill='x', expand=True)

        frame_bottom = ttk.Frame(root, padding=10)
        frame_bottom.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(
            frame_bottom,
            columns=('Telefone', 'Mensagem'),
            show='headings'
        )
        self.tree.heading('Telefone', text='Telefones')
        self.tree.heading('Mensagem', text='Mensagens Prontas para Serem Disparadas')
        self.tree.column('Telefone', width=85)
        self.tree.column('Mensagem', width=1200)
        self.tree.pack(fill='both', expand=True)

        self.combo_mes.bind("<<ComboboxSelected>>", self.limpar_filtro_grupo)
        self.combo_dia.bind("<<ComboboxSelected>>", self.limpar_filtro_grupo)

        self.combo_group.bind("<<ComboboxSelected>>", self.limpar_filtros_data)
       

        frame_actions = ttk.Frame(root, padding=10)
        frame_actions.pack(fill='x')

        tk.Button(
            frame_actions,
            text="EDITAR",
            command=self.edit_message,
            bg="#b34b00",
            fg="white"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="EXCLUIR - Delete",
            command=self.delete_message,
            bg="#dc3545",
            fg="white"
        ).pack(side='left', padx=5)

        frame_actions.bind_all("<Delete>", lambda event: self.delete_message())

        tk.Button(
            frame_actions,
            text="LIMPAR Mensagens",
            command=self.cancel_operation,
            bg="#ffc107",
            fg="black"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="CARREGAR Mensagens",
            command=self.carregar_e_filtrar_telefone,
            bg="#0056b3",
            fg="white"
        ).pack(side='left', padx=10)

        tk.Button(
            frame_actions,
            text="INICIAR Disparo",
            command=self.iniciar_envio_thread,
            bg="#28a745",
            fg="white"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="PAUSAR Envio",
            command=self.pausar_envio,
            bg="#ff8000",
            fg="white"
        ).pack(side='left', padx=5)

        tk.Button(
            frame_actions,
            text="Fechar",
            command=root.destroy,
            bg="#6c757d",
            fg="white"
        ).pack(side='left', padx=5)

        self.label_info = ttk.Label(root, text="Nenhuma Mensagem Carregada.")
        self.label_info.pack(anchor='w', padx=10)

        self.carregar_dados_do_mysql()

### CONECTANDO NO BANCO DE DADOS DO SISTEMA EVOLUTI
    def carregar_dados_do_mysql(self):
        try:
            conn = mysql.connector.connect(
                host='localhost',
                user='root',
                password='xxxxxxxx',
                database='xxxxxxxx'
            )
            cursor = conn.cursor(dictionary=True)
            cursor.execute("SELECT contato, nome, telefone1, dtNasc, rgIm, email FROM pessoas")
            resultados = cursor.fetchall()
            conn.close()

            self.df = pd.DataFrame(resultados)

            self.df.rename(columns={
                "nome": "empresa",
                "telefone1": "telefone",
                "dtNasc": "dtNasc",
                "email": "grupo"
            }, inplace=True)

        # --- NORMALIZAÇÃO DO CAMPO GRUPO ---
            self.df['grupo'] = self.df['grupo'].astype(str).str.strip().str.upper()

        # Debug: imprimir valores únicos na coluna grupo com repr para ver espaços invisíveis
            print("Valores únicos de grupo (com repr para ver espaços):")
            print([repr(g) for g in sorted(self.df['grupo'].unique())])

            self.label_info.config(text="🔄 Conectado ao banco de dados.")
            self.update_group_list()

        except mysql.connector.Error as e:
            print(f"Erro MySQL: {e}")

    def carregar_e_filtrar_telefone(self):
        grupo_selecionado = self.combo_group.get().strip().upper()
        mes_selecionado = self.combo_mes.get().strip()
        dia_selecionado = self.combo_dia.get()

        self.df['dtNasc'] = pd.to_datetime(self.df['dtNasc'], errors='coerce')

    # Começa com registros válidos (com datas)
        self.filtered_df = self.df[self.df['dtNasc'].notna()].copy()

        if grupo_selecionado:
        # Filtra DataFrame normalizado
            df_filtrado = self.df[self.df['grupo'] == grupo_selecionado]
            print(f"Registros encontrados com grupo '{grupo_selecionado}': {len(df_filtrado)}")
            print(df_filtrado[['grupo', 'telefone', 'empresa']])

            self.filtered_df = df_filtrado.copy()

        elif mes_selecionado:
            try:
                mes_num = int(mes_selecionado.split(" - ")[0]) if " - " in mes_selecionado else int(mes_selecionado)
                self.filtered_df = self.filtered_df[self.filtered_df['dtNasc'].dt.month == mes_num]

                if dia_selecionado != "Todos":
                    self.filtered_df = self.filtered_df[self.filtered_df['dtNasc'].dt.day == int(dia_selecionado)]

                print(f"Filtro por mês {mes_num}, dia {dia_selecionado}: {len(self.filtered_df)} registros encontrados.")
            except ValueError:
                messagebox.showwarning("Erro", "Mês selecionado inválido.")
                self.filtered_df = pd.DataFrame()

        else:
            messagebox.showwarning("Aviso", "Selecione um grupo ou mês para carregar os contatos.")
            self.filtered_df = pd.DataFrame()
            return

        if self.filtered_df.empty:
            messagebox.showwarning("Aviso", "Nenhum contato encontrado com os filtros selecionados.")
            return

        self.load_messages()

    

# listando grupos na grid de procura
    def update_group_list(self):
        if self.df is not None and 'grupo' in self.df.columns:
            grupos = self.df['grupo'].dropna().astype(str).str.strip().str.upper().unique()
            self.combo_group['values'] = sorted(grupos)
            self.combo_group.set("")

        # Garante que "Selecione" esteja entre os valores do combo_mes, se não estiver
            if "Selecione" not in self.combo_mes['values']:
                self.combo_mes['values'] = ["Selecione"] + list(self.combo_mes['values'])

            self.combo_mes.set("Selecione")
            self.combo_dia.set("Todos")

#  LIMPAR DATA AO SELECIONAR GRUPO
    def limpar_filtros_data(self, event=None):
        self.combo_mes.set("Selecione")
        self.combo_dia.set("Todos")  
### LIMPAR FILTROS DO GRUPO AO SELECIONAR MES OU DIA
    def limpar_filtro_grupo(self, event=None):
        self.combo_group.set("")      

    def edit_template(self, event=None):
        tipo = self.combo_msg_type.get()
        template = mensagens.get(tipo, "")
        self.txt_template.delete('1.0', tk.END)
        self.txt_template.insert(tk.END, template)

    def load_messages(self):
        if self.filtered_df is None or self.filtered_df.empty:
            messagebox.showwarning("Aviso", "Nenhum contato encontrado para o filtro atual.")
            return
    
        tipo = self.combo_msg_type.get()
        if not tipo:
            messagebox.showwarning("Aviso", "Selecione o Tipo de Mensagem a ser Enviada.")
            return

        template = self.txt_template.get("1.0", tk.END).strip()

        self.tree.delete(*self.tree.get_children())
        count = 0
        ## USAR AS VÁRIAVEIS OI, OLA E OPA
        cumprimentos = ["Oi", "Olá", "Opa"]
        indice_cumprimento = 0  # Começa com "Oi"


        for _, row in self.filtered_df.iterrows():
            nome = str(row.get('contato', 'contato')).strip()
            empresa = str(row.get('empresa', 'Empresa')).strip()
            telefone = str(row.get('telefone', '')).strip()
    
            if not telefone:
                continue

    # Define cumprimento intercalado
            inicio = cumprimentos[indice_cumprimento]
            indice_cumprimento = (indice_cumprimento + 1) % len(cumprimentos)

    # Monta mensagem
            msg = f"{inicio} {template.format(nome=nome, empresa=empresa)}".strip()
    
    # Insere na Treeview
            self.tree.insert('', tk.END, values=(telefone, msg))
            count += 1

        self.label_info.config(text=f"Total de mensagens carregadas a DISPARAR: {count}")

    # (demais métodos continuam iguais, sem alterações nesta parte)


    def iniciar_envio_thread(self):
        if self.enviando:
            messagebox.showinfo("Info", "Envio já está em andamento.")
            return
        if not self.tree.get_children():
            messagebox.showwarning("Aviso", "Nenhuma mensagem para enviar.")
            return

        self.enviando = True
        self.envio_ativo.set()
        threading.Thread(target=self.enviar_mensagens, daemon=True).start()

    def enviar_mensagens(self):
        # Exemplo simplificado
        for item in self.tree.get_children():
            if not self.enviando or not self.envio_ativo.is_set():
                break

            telefone, mensagem = self.tree.item(item, 'values')

            url = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
            webbrowser.open(url)
            time.sleep(10)  # esperar o whatsapp carregar

            # Simular enter para enviar
            pyautogui.press('enter')
            time.sleep(3)
            time.sleep(120)

            # Remover da treeview após envio
            self.tree.delete(item)

        self.enviando = False
        self.label_info.config(text="Envio concluído.")
        winsound.Beep(1000, 500)  # beep quando terminar

    def pausar_envio(self):
        if not self.enviando:
            messagebox.showinfo("Info", "Nenhum envio em andamento.")
            return
        if self.envio_ativo.is_set():
            self.envio_ativo.clear()
            self.label_info.config(text="Envio pausado.")
        else:
            self.envio_ativo.set()
            self.label_info.config(text="Envio retomado.")
            threading.Thread(target=self.enviar_mensagens, daemon=True).start()

    def delete_message(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Nenhuma mensagem selecionada para excluir.")
            return
        for item in selected_items:
            self.tree.delete(item)
        self.label_info.config(text=f"Mensagens selecionadas excluídas.")

    def cancel_operation(self):
        self.tree.delete(*self.tree.get_children())
        self.label_info.config(text="Mensagens limpas.")

    def edit_message(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione uma mensagem para editar.")
            return
        item = selected_items[0]
        telefone, mensagem = self.tree.item(item, 'values')

        def salvar_edicao():
            nova_msg = text_edit.get("1.0", tk.END).strip()
            self.tree.item(item, values=(telefone, nova_msg))
            edit_window.destroy()

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Editar Mensagem")
        text_edit = tk.Text(edit_window, height=10, width=50)
        text_edit.pack(padx=10, pady=10)
        text_edit.insert(tk.END, mensagem)

        btn_save = tk.Button(edit_window, text="Salvar", command=salvar_edicao)
        btn_save.pack(pady=(0, 10))


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal temporariamente

    app = WhatsAppSenderApp(root)

    root.deiconify()  # Exibe a janela depois que tudo estiver pronto
    root.mainloop()
