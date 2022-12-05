from cProfile import label
from sre_parse import expand_template
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import font
from turtle import right
import requests
import openpyxl
from excel_central_ferramentas import excel
from classes_principais import *
import pandas as pd


### Funções auxiliares da janela ferramenta ###
class cadastrar_ferramentas:

    def focus1(self, evento):
        self.descricao.focus_set()


    def focus2(self, evento):
        self.fabricante.focus_set()


    def focus3(self, evento):
        self.voltagem.focus_set()


    def focus4(self, evento):
        self.part_number.focus_set()


    def focus5(self, evento):
        self.tamanho.focus_set()


    def focus6(self, evento):
        self. tipo.focus_set()

    def focus7(self, evento):
        self.material.focus_set()

    def focus8(self, evento):
        self.maximo_reserva.focus_set()


    def clear(self):
        self.descricao.delete(0, END)
        self.fabricante.delete(0, END)
        self.voltagem.delete(0, END)
        self.part_number.delete(0, END)
        self.tamanho.delete(0, END)
        self.tipo.delete(0, END)
        self.material.delete(0, END)
        self.maximo_reserva.delete(0, END)


    def inserir(self):
        if (self.descricao.get() == "" or self.descricao.get() == " " and
                self.fabricante.get() == "" or self.fabricante.get() == " " and
                self.voltagem.get() == "" or self.voltagem.get() == " " and
                self.part_number.get() == "" or self.part_number.get() == " " and
                self.tamanho.get() == "" or self.tamanho.get() == " " and
                self.tipo.get() == "" or self.tipo.get() == " " and
                self.material.get() == "" or self.material.get() == " " and
                self.maximo_reserva.get() == "" or self.maximo_reserva.get() == " "):

            messagebox.showerror("Error", "Voce deve preencher todos os campos")
        elif (self.voltagem.get() != '110V' and self.voltagem.get() != "220V" and self.voltagem.get() != "360V" and self.voltagem.get() != "110~220V"
        and self.voltagem.get() != "sem voltagem"):
            messagebox.showerror("Error", "Essa voltagem não existe, favor preencher corretamente")
        elif (self.tamanho.get() != 'Polegadas' and self.tamanho.get() != "MM" and self.tamanho.get() != "CM" and self.tamanho.get() != "M"):
            messagebox.showerror("Error", "Tamanho não aceitável, favor preencher corretamente")
        else:
            resposta = messagebox.askquestion("Tem certeza?", "Cadastrar ferramenta " + self.descricao.get() + "?")
            if resposta == "yes":
                planilha = excel ('Aba Ferramentas')
                classe_ferramentas = ferramentas()
                classe_ferramentas.setDescricao_ferramenta (self.descricao.get())
                classe_ferramentas.setFabricante (self.fabricante.get())
                classe_ferramentas.setVoltagem (self.voltagem.get())
                classe_ferramentas.setPart_number (self.part_number.get())
                classe_ferramentas.setTamanho (self.tamanho.get())
                classe_ferramentas.setTipo (self.tipo.get())
                classe_ferramentas.setMaterial (self.material.get())
                classe_ferramentas.setMaximo_reserva (self.maximo_reserva.get())

                planilha.salvar_ferramentas(classe_ferramentas)

                messagebox.showinfo("Sucesso", "Ferramenta cadastrada com sucesso!")
                self.descricao.focus_set()
            else:
                messagebox.showwarning("Cancelado", "Ferramenta não cadastrado!")
            self.clear()



### JANELA FERRAMENTAS ###
    def __init__(self):
        janelaferramentas = Toplevel()

        
        frame = Frame(janelaferramentas, bd=4,bg='light blue', highlightthickness=3, highlightbackground='black')
        frame.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)

        janelaferramentas.configure(background='light blue')
        janelaferramentas.title("Cadastro de Ferramentas")
        janelaferramentas.state ('zoomed')

        vazio_1 = Label(janelaferramentas, text="", bg="light blue")
        
        titulo_1 = Label(janelaferramentas, text="Cadastro de Ferramentas", bg="light blue")
        titulo_1.configure (font=('Georgia',13), fg="blue")
      

        descricao_1 = Label(janelaferramentas, text="Descrição da Ferramenta", bg="light blue")
           
        fabricante_1 = Label(janelaferramentas, text="Fabricante", bg="light blue")
       
        voltagem_1 = Label(janelaferramentas, text="Voltagem", bg="light blue")
     
        part_number_1 = Label(janelaferramentas, text="Part Number", bg="light blue")

        tamanho_1 = Label(janelaferramentas, text="Tamanho", bg="light blue")

        unidadeMedida_1 = Label(janelaferramentas, text="Unidade de medida", bg="light blue")

        tipo_1 = Label(janelaferramentas, text="Tipo de Ferramenta", bg="light blue")

        material_1 = Label(janelaferramentas, text="Material", bg="light blue")

        maximo_reserva_1 = Label(janelaferramentas, text="Máximo de Reserva (hrs)", bg="light blue")

        vazio_1.grid(row=1, column=10)
        titulo_1.grid(row=2, column=1)
        descricao_1.grid(row=4, column=0)
        fabricante_1.grid(row=6, column=0)
        voltagem_1.grid(row=8, column=0)
        part_number_1.grid(row=10, column=0)
        tamanho_1.grid(row=12, column=0)
        unidadeMedida_1.grid(row=12, column=3)
        tipo_1.grid(row=14, column=0)
        material_1.grid(row=16, column=0)
        maximo_reserva_1.grid(row=18, column=0)

        self.descricao = Entry(janelaferramentas)
        self.fabricante = Entry(janelaferramentas)
        self.voltagem = Entry(janelaferramentas)
        self.part_number = Entry(janelaferramentas)
        self.tamanho = Entry(janelaferramentas)
        self.tipo = Entry(janelaferramentas)
        self.material = Entry(janelaferramentas)
        self.maximo_reserva = Entry(janelaferramentas)
        self.unidadeMedida = Entry(janelaferramentas)

        self.descricao.bind("<Return>", self.focus1)

        self.fabricante.bind("<Return>", self.focus2)

        self.voltagem.bind("<Return>", self.focus3)

        self.part_number.bind("<Return>", self.focus4)

        self.tamanho.bind("<Return>", self.focus5)

        self.tipo.bind("<Return>", self.focus6)

        self.material.bind("<Return>", self.focus7)

        self.maximo_reserva.bind("<Return>", self.focus8)

        self.descricao.grid(row=4, column=1, ipadx="100")
        self.fabricante.grid(row=6, column=1, ipadx="100")
        self.voltagemEscolha = ["110V", "220V", "360V", "110V~220V", "sem voltagem"]
        self.voltagem = ttk.Combobox(janelaferramentas, values=self.voltagemEscolha)
        self.voltagem.set("110V")
        self.voltagem.grid(row=8, column=1, ipadx="90")
        self.part_number.grid(row=10,column=1,ipadx="100")
        self.tamanhoEscolha = ["Polegadas", "MM", "CM", "M"]
        self.tamanho = ttk.Combobox(janelaferramentas, values=self.tamanhoEscolha)
        self.tamanho.set("Polegadas")
        self.tamanho.grid(row=12, column=4, ipadx="0")
        self.unidadeMedida.grid(row=12, column=1,ipadx="100")
        self.tipo.grid(row=14, column=1, ipadx="100")
        self.material.grid(row=16, column=1, ipadx="100")
        self.maximo_reserva.grid(row=18, column=1, ipadx="100")

        
        cadastrar = Button(janelaferramentas, text="Cadastrar", fg="White",
                        bg="Black", command=self.inserir)
        cadastrar.grid(row=20, column=1)

        class button():
            def exit1():
                janelaferramentas.destroy()
                janelaferramentas.update()
            #botao_salvar = Button(janelaferramentas, text="salvar").place(relx=0.87, rely=0.88, relwidth=0.12,relheight=0.1)-AGUARDAR IMPLEMENTACAO
            botao_sair1 = Button(janelaferramentas, text="Voltar", command=exit1).place(relx=0.85, rely=0.85, relwidth=0.08,relheight=0.05) 
def abrir_ferramentas():
    janela_ferramentas = cadastrar_ferramentas ()
       

### Funções auxiliares da janela tecnico ###

class cadastrar_tecnico:
        def focus1(self,evento):
            self.cpf.focus_set()


        def focus2(self, evento):
            self.nome_completo.focus_set()


        def focus3(self, evento):
            self.telefone.focus_set()


        def focus4(self, evento):
            self.turno.focus_set()


        def focus5(self, evento):
            self.equipe.focus_set()


        def focus6(self, evento):
            self.email.focus_set()


        def clear(self):
            self.cpf.delete(0, END)
            self.nome_completo.delete(0, END)
            self.telefone.delete(0, END)
            self.turno.delete(0, END)
            self.equipe.delete(0, END)
            self.email.delete(0, END)


        def inserir(self):
            if (self.cpf.get() == "" or self.cpf.get() == " " and
                    self.nome_completo.get() == "" or self.nome_completo.get() == " " and
                    self.telefone.get() == "" or self.telefone.get() == " " and
                    self.turno.get() == "" or self.turno.get() == " " or self.turno.get() and
                    self.equipe.get() == "" or self.equipe.get() == " " and
                    self.email.get() == "" or self.email.get() == " "):
                messagebox.showerror("Error", "Voce deve preencher todos os campos")
            elif (self.turno.get() != 'Manhã' and self.turno.get() != "Tarde" and self.turno.get() != "Noite"):
                messagebox.showerror("Error", "Este Turno não existe, favor preencher corretamente")
            else:
                resposta = messagebox.askquestion("Tem certeza?", "Cadastrar técnico " + self.nome_completo.get() + "?")
                if resposta == "yes":
                    messagebox.showinfo("Sucesso", "Técnico cadastrado com sucesso!")
                    
                    planilha = excel ('Aba Técnicos')
                    classe_tecnicos = tecnico()
                    classe_tecnicos.setCpf (self.cpf.get())
                    classe_tecnicos.setNome (self.nome_completo.get())
                    classe_tecnicos.setTelefone (self.telefone.get())
                    classe_tecnicos.setTurno (self.turno.get())
                    classe_tecnicos.setEquipe (self.equipe.get())
                    classe_tecnicos.setEmail(self.email.get())

                    planilha.salvar_tecnicos(classe_tecnicos)
                    self.cpf.focus_set()
                else:
                    messagebox.showwarning("Cancelado", "Técnico não cadastrado!")
                self.clear()

### JANELA TECNICOS ###
        def __init__(self):
            janelatecnicos = Toplevel()
   
            frame = Frame(janelatecnicos, bd=4,bg='light blue', highlightthickness=3, highlightbackground='black')
            frame.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)

            janelatecnicos.configure(background='light blue')
            janelatecnicos.title("Cadastro de Técnicos")
            janelatecnicos.state ('zoomed')

            self.vazio_1 = Label(janelatecnicos, text="", bg="light blue")

            self.titulo_1 = Label(janelatecnicos, text="Cadastro de Técnicos", bg="light blue")
            self.titulo_1.configure (font=('Georgia',13), fg="blue")

            self.cpf_1 = Label(janelatecnicos, text="CPF", bg="light blue")

            self.nome_completo_1 = Label(janelatecnicos, text="Nome Completo", bg="light blue")

            self.telefone_1 = Label(janelatecnicos, text="Telefone", bg="light blue")

            self.turno_1 = Label(janelatecnicos, text="Turno", bg="light blue")

            self.equipe_1 = Label(janelatecnicos, text="Equipe", bg="light blue")

            self.email_1 = Label(janelatecnicos, text="Email", bg="light blue")

            self.vazio_1.grid (row=0, column =10)
            self.titulo_1.grid(row=1, column=1)
            self.cpf_1.grid(row=2, column=0)
            self.nome_completo_1.grid(row=3, column=0)
            self.telefone_1.grid(row=4, column=0)
            self.turno_1.grid(row=5, column=0)
            self.equipe_1.grid(row=6, column=0)
            self.email_1.grid(row=7, column=0)

            self.cpf = Entry(janelatecnicos)
            self.nome_completo = Entry(janelatecnicos)
            self.telefone = Entry(janelatecnicos)
            self.turno = Entry(janelatecnicos)
            self.equipe = Entry(janelatecnicos)
            self.email = Entry(janelatecnicos)

            self.cpf.bind("<Return>", self.focus1)

            self.nome_completo.bind("<Return>", self.focus2)

            self.telefone.bind("<Return>", self.focus3)

            self.turno.bind("<Return>", self.focus4)

            self.equipe.bind("<Return>", self.focus5)

            self.email.bind("<Return>", self.focus6)

            self.cpf.grid(row=2, column=1, ipadx="100")
            self.nome_completo.grid(row=3, column=1, ipadx="100")
            self.telefone.grid(row=4, column=1, ipadx="100")
            self.turnoEscolha = ["Manhã", "Tarde", "Noite"]
            self.turno = ttk.Combobox(janelatecnicos, values=self.turnoEscolha)
            self.turno.set("Manhã")
            self.turno.grid(row=5,column=1,ipadx="90")
            self.equipe.grid(row=6, column=1, ipadx="100")
            self.email.grid(row=7, column=1, ipadx="100")


            cadastrar = Button(janelatecnicos, text="Cadastrar", fg="White",
                            bg="Black", command=self.inserir)
            cadastrar.grid(row=8, column=1) 
            
        
            class button():
                def exit2():
                    janelatecnicos.destroy()
                    janelatecnicos.update()
                #botao_salvar = Button(janelatecnicos, text="salvar").place(relx=0.87, rely=0.88, relwidth=0.12,relheight=0.1)-AGUARDAR IMPLEMENTACAO
                botao_sair2 = Button(janelatecnicos, text="Voltar", command=exit2).place(relx=0.85, rely=0.85, relwidth=0.08,relheight=0.05)
            
def abrir_tecnicos():
    janelatecnicos = cadastrar_tecnico ()

        


### JANELA CONSULTA ###
class consulta_geral:
    
    def excluir(self):
        entradas_selecionadas = self.visualizar_excel.selection()
        campo = self.campo_consulta.get()
        escolha = " "
        if campo == 'Ferramentas':
            escolha = 'Aba Ferramentas'
        elif campo == 'Técnicos':
            escolha = 'Aba Técnicos'
        if len(entradas_selecionadas) > 0:
            resposta = messagebox.askquestion("Confirmar exclusão","Deseja excluir a linha selecionada?")
            if resposta == "yes":
                for entrada in entradas_selecionadas:
                    item_atual = self.visualizar_excel.item(entrada)
                    lista_itens = item_atual.get("values")
                    id_itens = lista_itens[0]
                    planilha = excel(escolha)
                    planilha.excluir_dados(id_itens)
                    self.visualizar_excel.delete(entrada)
                    self.visualizar_excel.focus_set ()
        else:
            messagebox.showerror("Error", "Nenhuma linha selecionada")      

    def __init__(self):
        janelaconsulta = Toplevel()
    
        frame = Frame(janelaconsulta, bd=4,bg='light blue', highlightthickness=3, highlightbackground='black')
        frame.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)

        janelaconsulta.configure(background='light blue')
        janelaconsulta.title("Consulta Cadastros")
        janelaconsulta.state ('zoomed')
        
        self.campo_consulta1 = Label(janelaconsulta, text="Selecione uma opção e clique para consultar", bg="light blue")
        self.campo_consulta1.place(relx=0.05, rely=0.05, relwidth=0.25,relheight=0.05)
        self.campo_consulta1.configure (font=('Georgia',13), fg="blue")
         

        self.campo_consultaEscolha = ["Ferramentas", "Técnicos"]
        self.campo_consulta = ttk.Combobox(janelaconsulta, values=self.campo_consultaEscolha)
        self.campo_consulta.set("Ferramentas")
        
        self.campo_consulta.place(relx=0.05, rely=0.13, relwidth=0.10,relheight=0.04)

        self.visualizar_excel = ttk.Treeview(frame, selectmode='browse', show='headings')

        class button():
        # o programa lê do arquivo excel e exibe na tela as informações da planilha # 
            def consultar():
                campo = self.campo_consulta.get()
                escolha = " "
                if campo == 'Ferramentas':
                    escolha = 'Aba Ferramentas'
                elif campo == 'Técnicos':
                    escolha = 'Aba Técnicos'
                
                df = pd.read_excel('arquivo_central_ferramentas.xlsx', sheet_name=escolha)

                self.visualizar_excel.delete(*self.visualizar_excel.get_children())
                            
                self.visualizar_excel ["column"] = list(df.columns)
                self.visualizar_excel ["show"] = "headings"

  
                for coluna in self.visualizar_excel ["column"]:
                    self.visualizar_excel.column (coluna, width=140, stretch= NO)
                    self.visualizar_excel.heading(coluna, text=coluna)
                    
   
                df_linhas = df.to_numpy().tolist()
                for linha in df_linhas:
                    self.visualizar_excel .insert("", "end", values=linha)

                self.visualizar_excel.place(relx=0.02, rely=0.15, relwidth=0.97, relheight=0.50)

            
           
            botao_consultar = Button(janelaconsulta, text="Consultar", command=consultar).place(relx=0.15, rely=0.13, relwidth=0.10,relheight=0.04)     
            
            botao_excluir = Button(janelaconsulta, text="Excluir", command=self.excluir).place(relx=0.85, rely=0.03, relwidth=0.08,relheight=0.05)

            def exit4():
                janelaconsulta.destroy()
                janelaconsulta.update()
        
            botao_sair4 = Button(janelaconsulta, text="Voltar", command=exit4).place(relx=0.85, rely=0.85, relwidth=0.08,relheight=0.05)

def abrir_consulta():
    janelaconsulta = consulta_geral ()



### JANELA/TELA INICIAL ###
janela = Tk()
janela.title("Central de Ferramentas")
janela.state('zoomed')

frame_fundo = Frame(janela, width=200, height=400, background='#DCDCDC')
frame_fundo.pack(side=LEFT, expand=True, fill=BOTH)

frame_frente = Frame(frame_fundo, width=200, height=100, background='#A9A9A9')
frame_frente.pack(side=TOP, fill=BOTH)

### FIGURAS ###
logo = PhotoImage(file="Imagens\GF_Ferramentas_banner.png")
logo = logo.subsample(3, 3)
figura1 = Label (janela, image=logo, bg='#DCDCDC')
figura1.pack(side=LEFT, expand=True, fill=BOTH)


### BOTOES ###
botao_ferramenta=Button(frame_frente, text="Ferramentas", command=abrir_ferramentas)
botao_ferramenta.configure(background="#e6e6fa", font=('Georgia',16), fg="blue")
botao_ferramenta.pack(padx=20, pady=20, side=LEFT)

botao_tecnicos = Button(frame_frente, text="Técnicos", command=abrir_tecnicos)
botao_tecnicos.configure(background="#e6e6fa", font=('Georgia',16), fg="blue")
botao_tecnicos.pack(padx=20, pady=20, side=LEFT)

botao_pesquisa = Button(frame_frente, text="Consulta", command=abrir_consulta)
botao_pesquisa.configure(background="#e6e6fa", font=('Georgia',16), fg="blue")
botao_pesquisa.pack(padx=20, pady=20, side=LEFT)


janela.mainloop()
