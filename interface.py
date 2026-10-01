import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox

class InterfaceGestaoComputadores:
    def __init__(self, root, backend):
        self.root = root
        self.excel = backend

        self.root.title("Controle de Computadores e Usuários")
        self.root.geometry("950x700")
        self.root.resizable(False, False)

        self.novos_itens = []
        self.itens_remover = []
        self.itens_substituir = {}

        self.patrimonios_existentes = set()
        self.mapa_linhas = {}
        self.dados_memoria = {}

        self.marcas_validas = []
        self.marcas_validas_lower = {}

        self.container = tb.Frame(self.root)
        self.container.pack(fill=BOTH, expand=True)

        self.carregar_dados_interface()
        self.menu_principal()

    def carregar_dados_interface(self):
        sucesso, marcas, pats, mapa, dados, erro = self.excel.carregar_dados()
        if not sucesso:
            messagebox.showerror("Erro Crítico", f"Erro ao abrir a planilha:\n{erro}")
            self.root.destroy()
            return

        self.marcas_validas = marcas
        self.marcas_validas_lower = {m.lower(): m for m in marcas}
        self.patrimonios_existentes = pats
        self.mapa_linhas = mapa
        self.dados_memoria = dados

    def limpar_tela(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def voltar_menu(self, fila_pendente):
        if len(fila_pendente) > 0:
            if not messagebox.askyesno("Aviso", "Você tem itens na fila não salvos." \
            "\nDeseja voltar e perder as alterações?"):
                return
        self.menu_principal()

    def menu_principal(self):
        self.limpar_tela()
        self.novos_itens.clear()
        self.itens_remover.clear()
        self.itens_substituir.clear()

        frame_menu = tb.Frame(self.container, padding=50)
        frame_menu.pack()

        tb.Label(
            frame_menu,
            text="Menu Principal",
            font=("Helvetica", 18, "bold")
        ).pack(pady=20)

        tb.Button(
            frame_menu,
            text="👁 Visualizar Planilha Completa",
            bootstyle=INFO,
            width=40,
            command=self.tela_visualizar
        ).pack(pady=10, ipady=10)

        
        tb.Button(
            frame_menu,
            text="➕ Adicionar Novos Itens",
            bootstyle=SUCCESS,
            width=40,
            command=self.tela_adicionar
        ).pack(pady=10, ipady=10)

        tb.Button(
            frame_menu,
            text="🔄 Substituir / Editar Valores",
            bootstyle=WARNING,
            width=40,
            command=self.tela_substituir
        ).pack(pady=10, ipady=10)

        tb.Button(
            frame_menu,
            text="🗑 Remover Itens",
            bootstyle=DANGER,
            width=40,
            command=self.tela_remover
        ).pack(pady=10, ipady=10)

        tb.Label(frame_menu, text="").pack(pady=10)
        tb.Button(
            frame_menu,
            text="Sair do Sistema",
            bootstyle=SECONDARY,
            width=40,
            command=self.root.quit
        ).pack(ipady=10)

    def tela_visualizar(self):
        self.limpar_tela()
        tb.Button(
            self.container,
            text="← Voltar ao Menu",
            bootstyle=LINK,
            command=lambda: self.voltar_menu([])
        ).pack(anchor=W, padx=10, pady=5)

        tb.Label(
            self.container,
            text="Planilha na Memória",
            font=("Helvetica", 16, "bold")
        ).pack(pady=5)

        frame_busca = tb.Frame(self.container, padding=10)
        frame_busca.pack(fill=X, padx=10)

        tb.Label(frame_busca, text="Buscar por Coluna:").grid(row=0, column=0, padx=5, sticky=W)

        opcoes_colunas = [
            "Todos", "Patrimônio", "Marca", "Armaz.",
            "RAM", "Servidor", "Usuário", "Monitores"
        ]

        self.combo_busca_col = tb.Combobox(frame_busca, values=opcoes_colunas, state="readonly", width=15)
        self.combo_busca_col.current(0)
        self.combo_busca_col.grid(row=0, column=1, padx=5)

        tb.Label(frame_busca, text="Valor:").grid(row=0, column=2, padx=5, sticky=W)
        self.entry_busca_val = tb.Entry(frame_busca, width=30)
        self.entry_busca_val.grid(row=0, column=3, padx=5)

        tb.Button(
            frame_busca,
            text="🔎 Buscar",
            bootstyle=PRIMARY,
            command=self.executar_busca
        ).grid(row=0, column=4, padx=10)
        
        tb.Button(
            frame_busca,
            text="Limpar",
            bootstyle=SECONDARY,
            command=self.limpar_busca
        ).grid(row=0, column=5)

        frame_lista = tb.Frame(self.container, padding=20)
        frame_lista.pack(fill=BOTH, expand=True)

        scroll_y = tb.Scrollbar(frame_lista, orient=VERTICAL)
        colunas = ("Patrimônio", "Marca Comp.", "Armaz.", "RAM", "Servidor", "Usuário", "Monitores")
        self.tree_vis = tb.Treeview(
            frame_lista, 
            columns=colunas,
            show="headings",
            height=15,
            bootstyle=INFO,
            yscrollcommand=scroll_y.set
        )

        scroll_y.config(command=self.tree_vis.yview)

        self.tree_vis.heading("Patrimônio", text="Patrimônio")
        self.tree_vis.heading("Marca Comp.", text="Marca Comp.")
        self.tree_vis.heading("Armaz.", text="Armaz.")
        self.tree_vis.heading("RAM", text="RAM")
        self.tree_vis.heading("Servidor", text="Sevidor")
        self.tree_vis.heading("Usuário", text="Usuário")
        self.tree_vis.heading("Monitores", text="Monitores")

        self.tree_vis.column("Patrimônio", width=80, anchor=CENTER)
        self.tree_vis.column("Marca Comp.", width=80, anchor=CENTER)
        self.tree_vis.column("Armaz.", width=100, anchor=CENTER)
        self.tree_vis.column("RAM", width=80, anchor=CENTER)
        self.tree_vis.column("Servidor", width=150, anchor=CENTER)
        self.tree_vis.column("Usuário", width=150, anchor=CENTER)
        self.tree_vis.column("Monitores", width=150, anchor=CENTER)

        scroll_y.pack(side=RIGHT, fill=Y)
        self.tree_vis.pack(side=LEFT, fill=BOTH, expand=True)

        self.lbl_linhas_totais = tb.Label(self.container, text="", font=("Helvetica", 12, "bold"))
        self.lbl_linhas_totais.pack(pady=5)

        tb.Button(
            self.container,
            text="Recarregar Dados",
            bootstyle=(INFO, OUTLINE),
            command=self.recarregar_dados
        ).pack(pady=10)

        self.preencher_tabela_visualizacao()

    def executar_busca(self):
        coluna = self.combo_busca_col.get()
        termo = self.entry_busca_val.get().strip()

        if not termo:
            self.preencher_tabela_visualizacao()
            return

        termo_num = termo.lower().replace("gb", "").replace("tb", "").strip()

        if termo_num.isdigit():
            numero = int(termo_num)
            if coluna == "Armaz.":
                if numero < 10:
                    termo = f"{numero} TB"
                elif numero >= 100:
                    termo = f"{numero} GB"
            elif coluna == "RAM":
                termo = f"{numero} GB"

        self.preencher_tabela_visualizacao(filtro_col=coluna, filtro_val=termo)

    def preencher_tabela_visualizacao(self, filtro_col=None, filtro_val=None):
        for row in self.tree_vis.get_children():
            self.tree_vis.delete(row)

        col_map = {
            "Patrimônio": 0,
            "Marca": 1,
            "Armaz.": 2,
            "RAM": 3,
            "Servidor": 4,
            "Usuário": 5,
            "Monitores": 6
        }

        resultados = 0
        pats_ordenados = sorted(self.mapa_linhas.keys(), key=lambda p: self.mapa_linhas[p])

        for pat in pats_ordenados:
            dados = self.dados_memoria[pat]
            d_limpos = [str(d) if d is not None else "" for d in dados]

            incluir = True
            if filtro_col and filtro_val:
                val_busca = filtro_val.lower().replace(" ", "")
                if filtro_col == "Patrimônio":
                    val_excel = pat.lower().replace(" ", "")
                else:
                    idx = col_map[filtro_col]
                    val_excel = d_limpos[idx].lower().replace(" ", "")
                if val_busca not in val_excel:
                    incluir = False

            if incluir:
                self.tree_vis.insert("", END, values=(
                    d_limpos[0], d_limpos[1], d_limpos[2], d_limpos[3], d_limpos[4], d_limpos[5]
                ))

                resultados += 1

        if filtro_val:
            self.lbl_linhas_totais.config(text=f"Linhas Encontradas: {resultados}")
            if resultados == 0:
                messagebox.showinfo(
                    "Busca Sem Resultados", 
                    f"Nenhuma ocorrência para '{filtro_val}' na coluna '{filtro_col}'.\n\nA tabela foi zerada."
                )
        else:
            self.lbl_linhas_totais.config(text=f"Linhas Totais: {len(self.dados_memoria)}")

    def limpar_busca(self):
        self.entry_busca_val.delete(0, END)
        self.preencher_tabela_visualizacao()

    def recarregar_dados(self):
        self.carregar_dados_interface()
        self.tela_visualizar()

    def tela_adicionar(self):
        self.limpar_tela()
        tb.Button(
            self.container,
            text="← Voltar ao Menu",
            bootstyle=LINK,
            command=lambda: self.voltar_menu(self.novos_itens)
        ).pack(anchor=W, padx=10, pady=5)

        tb.Label(self.container, text="Adicionar Novos Computadores/Usuários", font=("Helvetica", 10, "bold")).pack()

        frame_inputs = tb.Frame(self.container, padding=10)
        frame_inputs.pack(fill=X)

        tb.Label(
            frame_inputs,
            text="Patrimônio (6 dígitos):",
            font=("Helvetica", 10, "bold"),
        ).grid(row=0, column=0, sticky=W, pady=5)

        self.entry_pat = tb.Entry(frame_inputs, width=25)
        self.entry_pat.grid(row=0, column=1, padx=5, pady=5)

        tb.Label(
            frame_inputs,
            text="Marca:",
            font=("Helvetica", 10, "bold"),
        ).grid(row=0, column=2, sticky=W, pady=5, padx=(20, 0))

        self.combo_marca = tb.Combobox(frame_inputs, values=self.marcas_validas, state="readonly", width=23)
        self.combo_marca.grid(row=0, column=3, pady=5, padx=5)

        tb.Label(
            frame_inputs,
            text="Armazenamento:",
            font=("Helvetica", 10, "bold"),
        ).grid(row=1, column=0, sticky=W, pady=5)

        self.entry_arm = tb.Entry(frame_inputs, width=25)
        self.entry_arm.grid(row=1, column=1, padx=5, pady=5)

        tb.Label(
            frame_inputs,
            text="Memória RAM:",
            font=("Helvetica", 10, "bold"),
        ).grid(row=1, column=2, sticky=W, pady=5, padx=(20, 0))

        self.entry_ram = tb.Entry(frame_inputs, width=25)
        self.entry_ram.grid(row=1, column=3, padx=5, pady=5)

        tb.Label(
            frame_inputs,
            text="Servidor:",
            font=("Helvetica", 10, "bold"),
        ).grid(row=2, column=0, sticky=W, pady=5)

        self.entry_serv = tb.Entry(frame_inputs, width=25)
        self.entry_serv.grid(row=2, column=1, padx=5, pady=5)

        tb.Label(
            frame_inputs,
            text="Usuário:",
            font=("Helvetica", 10, "bold"),
        ).grid(row=2, column=2, sticky=W, pady=5, padx=(20, 0))

        self.entry_usuario = tb.Entry(frame_inputs, width=25)
        self.entry_usuario.grid(row=2, column=3, padx=5, pady=5)

        tb.Label(
            frame_inputs,
            text="Setup Monitores:",
            font=("Helvetica", 10, "bold")
        ).grid(row=3, column=0, sticky=W, pady=5)

        self.entry_monitores = tb.Entry(frame_inputs, width=65)
        self.entry_monitores.grid(row=3, column=1, columnspan=3, padx=5, pady=5, sticky=W)

        tb.Button(
            frame_inputs,
            text="Adicionar à Fila",
            bootstyle=INFO,
            command=self.adicionar_item_fila
        ).grid(row=4, column=0, columnspan=4, padx=10, ipadx=20)

        frame_lista = tb.Frame(self.container, padding=10)
        frame_lista.pack(fill=BOTH, expand=True)
        colunas = ("Patrimônio", "Marca", "Armaz.", "RAM", "Servidor", "Usuário", "Monitores")
        self.tree_add = tb.Treeview(frame_lista, columns=colunas, show="headings", height=5, bootstyle=INFO)
        for col in colunas:
            self.tree_add.heading(col, text=col)
            self.tree_add.column(col, width=90, anchor=CENTER)
        self.tree_add.pack(fill=BOTH, expand=True)

        tb.Button(
            self.container,
            text="Salvar Adições",
            bootstyle=SUCCESS,
            command=self.salvar_adicao
        ).pack(padx=10, ipadx=40, ipady=10)

    def adicionar_item_fila(self):
        pat = self.entry_pat.get().strip()
        marca = self.combo_marca.get().strip()
        arm = self.entry_arm.get().strip()
        ram = self.entry_ram.get().strip()
        serv = self.entry_serv.get().strip()
        usr = self.entry_usuario.get().strip()
        mon = self.entry_monitores.get().strip()

        if len(pat) != 6 or not pat.isdigit():
            messagebox.showwarning("Erro", "O patrimõnio deve conter exatamente 6 números.")
            return
        if pat in self.patrimonios_existentes or any(p[0] == pat for p in self.novos_itens):
            messagebox.showwarning("Erro", f"O patrimõnio '{pat}' já existe!")
            return
        if not marca or not arm or not ram or not serv or not usr or not mon:
            messagebox.showwarning("Erro", "Nenhum campo deve ser vazio para a adição.")
            return

        marca_formato = self.marcas_validas_lower.get(marca.lower(), marca)
        self.novos_itens.append([int(pat), marca_formato, arm, ram, serv, usr, mon])
        self.tree_add.insert("", "end", values=(pat, marca_formato, arm, ram, serv, usr, mon))
        
        self.entry_pat.delete(0, END)
        self.combo_marca.set('')
        self.entry_arm.delete(0, END)
        self.entry_ram.delete(0, END)
        self.entry_serv.delete(0, END)
        self.entry_usuario.delete(0, END)
        self.entry_monitores.delete(0, END)

    def salvar_adicao(self):
        if not self.novos_itens: return
        if not self.excel.travar_planilha():
            messagebox.showerror("Erro", "A planilha está aberta no Excel. Feche-a.")
            return

        sucesso, erro = self.excel.salvar_adicoes(self.novos_itens)
        if sucesso:
            self.carregar_dados_interface()
            self.novos_itens.clear()
            for row in self.tree_add.get_children(): self.tree_add.delete(row)
            messagebox.showinfo("Sucesso", "Itens Adicionados com sucesso.")
        else:
            messagebox.showerror("Erro ao Salvar", erro)

    def tela_substituir(self):
        self.limpar_tela()
        tb.Button(
            self.container,
            text="← Voltar ao Menu",
            bootstyle=LINK,
            command=lambda: self.voltar_menu(self.itens_substituir)
        ).pack(anchor=W, padx=10, pady=5)

        tb.Label(self.container, text="Substituir / Editar Itens", font=("Helvetica", 14, "bold")).pack()

        frame_busca = tb.Frame(self.container, padding=10)
        frame_busca.pack(fill=X)
        tb.Label(frame_busca, text="Chave (Patrimônio):").grid(row=0, column=0, sticky=W, padx=10)
        self.entry_subst_pat = tb.Entry(frame_busca, width=20)
        self.entry_subst_pat.grid(row=0, column=1, padx=10)

        tb.Button(
            frame_busca,
            text="Buscar Dados",
            bootstyle=PRIMARY,
            command=self.buscar_item_substituir
        ).grid(row=0, column=2, padx=1)

        self.frame_edicao = tb.Labelframe(
            self.container,
            text="Deixe em branco para manter original",
            padding=15,
            bootstyle=WARNING
        )

        self.frame_edicao.pack(fill=X, padx=20, pady=10)

        tb.Label(self.frame_edicao, text="Marca:").grid(row=0, column=0, sticky=W, pady=5)
        self.combo_subst_marca = tb.Combobox(
            self.frame_edicao, 
            values=self.marcas_validas,
            state="readonly",
            width=23
        )

        self.combo_subst_marca.grid(row=0, column=1, pady=5, padx=5)

        tb.Label(self.frame_edicao, text="Armaz.:").grid(row=0, column=2, sticky=W, pady=5)
        self.entry_subst_arm = tb.Entry(self.frame_edicao, width=25)
        self.entry_subst_arm.grid(row=0, column=3, pady=5, padx=5)

        tb.Label(self.frame_edicao, text="RAM:").grid(row=1, column=0, sticky=W, pady=5)
        self.entry_subst_ram = tb.Entry(self.frame_edicao, width=25)
        self.entry_subst_ram.grid(row=1, column=1, pady=5, padx=5)

        tb.Label(self.frame_edicao, text="Servidor:").grid(row=1, column=2, sticky=W, pady=5)
        self.entry_subst_serv = tb.Entry(self.frame_edicao, width=25)
        self.entry_subst_serv.grid(row=1, column=3, pady=5, padx=5)

        tb.Label(self.frame_edicao, text="Usuário:").grid(row=2, column=0, sticky=W, pady=5)
        self.entry_subst_usuario = tb.Entry(self.frame_edicao, width=25)
        self.entry_subst_usuario.grid(row=2, column=1, pady=5, padx=5)

        tb.Label(self.frame_edicao, text="Monitores:").grid(row=2, column=2, sticky=W, pady=5)
        self.entry_subst_monitores = tb.Entry(self.frame_edicao, width=25)
        self.entry_subst_monitores.grid(row=2, column=3, pady=5, padx=5)

        tb.Button(
            self.frame_edicao,
            text="Adicionar à Fila de Substituição",
            bootstyle=WARNING,
            command=self.adicionar_item_substituicao
        ).grid(row=3, column=0, columnspan=4, pady=10)

        frame_lista = tb.Frame(self.container, padding=10)
        frame_lista.pack(fill=BOTH, expand=True)
        colunas = ("Patrimônio", "Marca", "Armaz.", "RAM", "Servidor", "Usuário", "Monitores")
        self.tree_subst = tb.Treeview(
            frame_lista,
            columns=colunas,
            show="headings",
            height=5,
            bootstyle=WARNING
        )

        for col in colunas:
            self.tree_subst.heading(col, text=col)
            self.tree_subst.column(col, width=90, anchor=CENTER)
        self.tree_subst.pack(fill=BOTH, expand=True)

        tb.Button(
            self.container,
            text="Salvar Substituições",
            bootstyle=SUCCESS,
            command=self.salvar_substituicao
        ).pack(pady=10, ipadx=40, ipady=10)

    def buscar_item_substituir(self):
        pat = self.entry_subst_pat.get().strip()
        if not pat: return
        if pat not in self.patrimonios_existentes:
            messagebox.showwarning("Não Encontrado", f"O patrimônio '{pat}' não existe.")
            return

        dados_originais = self.dados_memoria[pat]
        d_str = [str(d) if d is not None else "" for d in dados_originais]

        self.combo_subst_marca.set(d_str[0])
        self.entry_subst_arm.delete(0, END)
        self.entry_subst_arm.insert(0, d_str[1])
        self.entry_subst_ram.delete(0, END)
        self.entry_subst_ram.insert(0, d_str[2])
        self.entry_subst_serv.delete(0, END)
        self.entry_subst_serv.insert(0, d_str[3])
        self.entry_subst_usuario.delete(0, END)
        self.entry_subst_usuario.insert(0, d_str[4])
        self.entry_subst_monitores.delete(0, END)
        self.entry_subst_monitores.insert(0, d_str[5])

    def adicionar_item_substituicao(self):
        pat = self.entry_subst_pat.get().strip()
        if not pat or pat not in self.patrimonios_existentes: return
        nova_marca = self.combo_subst_marca.get().strip()
        novo_arm = self.entry_subst_arm.get().strip()
        nova_ram = self.entry_subst_ram.get().strip()
        novo_serv = self.entry_subst_serv.get().strip()
        novo_usr = self.entry_subst_usuario.get().strip()
        novo_mon = self.entry_subst_monitores.get().strip()

        orig = self.dados_memoria[pat]
        d_orig = [str(d) if d is not None else "" for d in orig]

        final_marca = nova_marca if nova_marca else d_orig[0]
        final_arm = novo_arm if novo_arm else d_orig[1]
        final_ram = nova_ram if nova_ram else d_orig[2]
        final_serv = novo_serv if novo_serv else d_orig[3]
        final_usr = novo_usr if novo_usr else d_orig[4]
        final_mon = novo_mon if novo_mon else d_orig[5]

        marca_formato = self.marcas_validas_lower.get(final_marca.lower(), final_marca)

        self.itens_substituir[pat] = [marca_formato, final_arm, final_ram, final_serv, final_usr, final_mon]

        existe = False
        for child in self.tree_subst.get_children():
            if str(self.tree_subst.item(child)["values"][0]) == pat:
                self.tree_subst.item(child, values=(pat, marca_formato, final_arm, final_ram, final_serv, final_usr, final_mon))
                existe = True
                break
        if not existe:
            self.tree_subst.insert("", END, values=(pat, marca_formato, final_arm, final_ram, final_serv, final_usr, final_mon))

        self.entry_subst_pat.delete(0, END)
        self.combo_subst_marca.set('')
        self.entry_subst_arm.delete(0, END)
        self.entry_subst_ram.delete(0, END)
        self.entry_subst_serv.delete(0, END)
        self.entry_subst_usuario.delete(0, END)
        self.entry_subst_monitores.delete(0, END)

    def salvar_substituicao(self):
        if not self.itens_substituir: return
        if not self.excel.travar_planilha(): return

        sucesso, erro = self.excel.salvar_substituicoes(self.itens_substituir, self.mapa_linhas)
        if sucesso:
            self.carregar_dados_interface()
            self.itens_substituir.clear()
            for row in self.tree_subst.get_children(): self.tree_subst.delete(row)
            messagebox.showinfo("Sucesso", "Itens substituídos!")
        else:
            messagebox.showerror("Erro ao Salvar", erro)

    def tela_remover(self):
        self.limpar_tela()
        tb.Button(
            self.container,
            text="← Voltar ao Menu",
            bootstyle=LINK,
            command=lambda: self.voltar_menu(self.itens_remover)
        ).pack(anchor=W, padx=10, pady=10)

        tb.Label(self.container, text="Remover Itens", font=("Helvetica", 14, "bold")).pack()

        frame_rem = tb.Frame(self.container, padding=20)
        frame_rem.pack(fill=X)
        tb.Label(frame_rem, text="Patrimônio a Remover:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky=W)
        self.entry_rem_pat = tb.Entry(frame_rem, width=25)
        self.entry_rem_pat.grid(row=0, column=1, padx=10)

        frame_lista = tb.Frame(self.container, padding=20)
        frame_lista.pack(fill=BOTH, expand=True)

        self.tree_rem = tb.Treeview(
            frame_lista,
            columns=("Patrimônio", "Usuário Relacionado", "Status"),
            show="headings",
            height=8,
            bootstyle=DANGER
        )

        self.tree_rem.heading("Patrimônio", text="Patrimônio")
        self.tree_rem.heading("Usuário Relacionado", text="Usuário Relacionado")
        self.tree_rem.heading("Status", text="Status")
        self.tree_rem.pack(fill=BOTH, expand=True)

        tb.Button(
            self.container,
            text="Salvar Remoções",
            bootstyle=(DANGER, OUTLINE),
            command=self.salvar_remocao
        ).pack(pady=20, ipadx=40, ipady=10)

    def adicionar_item_remocao(self):
        pat = self.entry_rem_pat.get().strip()
        if not pat: return
        if pat not in self.patrimonios_existentes:
            messagebox.showwarning("Erro", "O patrimônio '{pat}' não existe.")
            return
        if pat in self.itens_remover: return

        usr_atual = self.dados_memoria[pat][4]
        if messagebox.askyesno("Confirmar Remoção", f"Tem certeza de que deseja remover o computador do usuário '{usr_atual}'?"):
            self.itens_remover.append(pat)
            self.tree_rem.insert("", END, values=(pat, usr_atual, "Aguardando Salvamento"))
            self.entry_rem_pat.delete(0, END)

    def salvar_remocao(self):
        if not self.itens_remover: return
        if not self.excel.travar_planilha(): return

        sucesso, erro = self.excel.salvar_remocoes(self.itens_remover, self.mapa_linhas)
        if sucesso:
            self.carregar_dados_interface()
            self.itens_remover.clear()
            for row in self.tree_rem.get_children(): self.tree_rem.delete(row)
            messagebox.showinfo("Sucesso", "Itens removidos!")
        else:
            messagebox.showerror("Erro ao Salvar", erro)