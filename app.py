import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from docx import Document
import unicodedata

# Lista de 81 escolas
escolas = [
    "E E Accácio de Vasconcelos Camargo, Prof.",
    "E E Aggeo Pereira do Amaral Prof.",
    "E E Altamir Gonçalves Prof.",
    "E E Amélia Cesar Machado de Araújo Profa.",
    "E E Ana Cecília Martins Prof.",
    "E E Antônia Lucchesi",
    "E E Antônio Cordeiro Prof.",
    "E E Antônio Miguel Pereira Junior",
    "E E Antônio Padilha",
    "E.E Antonio Vieira Campos",
    "E E Arquiminio Marques da Silva",
    "E E Arthur Cyrillo Freire",
    "E E Baltazar Fernandes",
    "E E Beathris Caixeiro Del Cistia Profª",
    "E E Brigadeiro Tobias",
    "CEEJA Professor Norberto Soares Ramos",
    "E E Diógenes de Almeida Marins Prof.",
    "E E Dionysio Vieira Prof.",
    "E E Dulce Esmeralda Basile Ferreira Profª.",
    "E E Elza Salvestro Bonilha",
    "E.E Elzide Celestina Souza Pacheco Tunuchi Profª. (Bairro do Eden)",
    "E E Enéas Proença de Arruda Prof.",
    "E E Escolástica Rosa de Almeida Profª.",
    "E E Ezequiel Machado Nascimento Prof.",
    "E E Fernanda de Camargo Pires Profª",
    "E E Flavio Gagliardi Prof.",
    "E E Francisco Camargo Cesar",
    "E E Francisco Coccaro",
    "E E Francisco Euphrasio Monteiro",
    "E E Genesio Machado Prof.",
    "E E Genezia Izabel Cardoso Mencacci Profª.",
    "E.E. Geraldo do Espirito Santo Fogaça",
    "E E Gualberto Moreira Dr.",
    "E E Guiomar Camolesi Souza",
    "E E Gumercindo Gonçalves",
    "E E Hélio Del Cistia",
    "E E Humberto de Campos",
    "E E Ida Yolanda Lanzoni de Barros Prof.",
    "E E Isabel Lopes Monteiro Profª.",
    "E E Izabel Rodrigues Galvão",
    "E E João Clímaco de Camargo Pires",
    "E E João Machado de Araújo",
    "E E João Rodrigues Bueno",
    "E E João Soares Mons (Altos do Ipanema)",
    "E E Joaquim Izidoro Marins Prof.",
    "E E Jordina Amaral Arruda Profª.",
    "E E Jorge Madureira",
    "E E José Odin de Arruda Prof.",
    "E E José Quevedo Prof.",
    "E E José Reginato",
    "E E José Roque Almeida Rosa Prof.",
    "E E Julia Rios Athayde Profª",
    "E E Julio Bierrenbach Lima Prof.",
    "E E Julio Prestes de Albuquerque Dr. (Estadão) e Centro de Línguas",
    "E E Laila Galep Sacker",
    "E E Lauro Sanchez Prof.",
    "E E Luiz Gonzaga de Camargo Fleury Prof.",
    "E E Luiz Nogueira Martins Senador",
    "E E Marco Antonio Mencacci Prof.",
    "E E Maria Cândida de Barros Araújo Profª",
    "EE Maria Helena Gazzi Bonadio Profª",
    "E E Maria Ondina de Andrade Profª.",
    "E E Marina Grohmann Soares Fernandes Profª.",
    "E E Mário Guilherme Notari",
    "E E Monteiro Lobato",
    "E E Nazira Nagib Jorge Murad Rodrigues Profª.",
    "E E Ossis Salvestrini Mendes Profª.",
    "E E Ovidio Antônio de Souza Reverendo",
    "E E Porto Seguro Visconde de",
    "E E Rafael Orsi Filho Prof.",
    "E E Renato Sêneca de Sá Fleury Prof.",
    "E E Roberto Paschoalick Prof.",
    "E E Roque Conceição Martins Prof.",
    "E E Rosemary de Mello Moreira Pereira",
    "E E Sarah Salvestro Prof.",
    "E.E. Senador Vergueiro",
    "E E Waldemar de Freitas Rosa Prof.",
    "E E Wanda Costa Daher Profª.",
    "E E Wilson Ramos Brandão Prof.",
    "E E Zelia Dulce de Campos Maia Profª."
]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de Pastas de Escolas")
        self.root.geometry("600x500")
        self.root.resizable(False, False)  # Fixa o tamanho da janela

        # Diretório base fixo (Documentos do usuário)
        self.base_dir = os.path.join(os.path.expanduser("~"), "Documents")

        # Frame principal para centralizar tudo
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(expand=True, anchor="center")

        # Interface
        tk.Label(self.main_frame, text="Escolas:").pack(pady=5, anchor="center")

        # Instrução para o usuário com wraplength
        instrucao = ("Clique duas vezes no nome da escola para criar a pasta e o documento Word, "
                     "ou marque as caixas e clique em 'Criar Pastas'. Use 'Selecionar Tudo' para marcar todas.")
        tk.Label(self.main_frame, text=instrucao, wraplength=500, justify="center").pack(pady=5, anchor="center")

        # Frame para a lista de escolas
        self.school_frame = tk.Frame(self.main_frame)
        self.school_frame.pack(pady=5, fill=tk.BOTH, expand=True, anchor="center")

        # Canvas para rolagem
        self.canvas = tk.Canvas(self.school_frame, width=550, height=250)
        self.scrollbar = tk.Scrollbar(self.school_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((275, 0), window=self.scrollable_frame, anchor="center")  # Centraliza o frame no canvas
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, anchor="center")
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y, anchor="center")

        # Lista de widgets para escolas
        self.school_widgets = {}
        for escola in escolas:
            frame = tk.Frame(self.scrollable_frame, bd=1, relief=tk.SOLID)  # Linha delimitadora
            frame.pack(fill=tk.X, pady=2, anchor="center")

            # Checkbox
            var = tk.BooleanVar()
            check = tk.Checkbutton(frame, variable=var, bg="#f0f0f0")  # Cinza claro
            check.pack(side=tk.LEFT, padx=5, pady=2)

            # Nome da escola
            label = tk.Label(frame, text=escola, bg="#ffffff", width=40, anchor="w")  # Branco
            label.pack(side=tk.LEFT, padx=5, pady=2)
            label.bind("<Double-1>", lambda e, s=escola: self.criar_pasta_com_documento(s))

            # Botão de pasta
            nome_pasta = escola.replace("/", "_").replace(":", "_").replace(",", "").replace(".", "").replace("(", "").replace(")", "")
            pasta = os.path.join(self.base_dir, nome_pasta)
            btn = tk.Button(frame, text="📁", bg="#ffff99" if os.path.exists(pasta) else "#ffffff", command=lambda p=pasta, s=escola: self.abrir_pasta(p, s))
            btn.pack(side=tk.LEFT, padx=5, pady=2)

            self.school_widgets[escola] = {"var": var, "frame": frame, "btn": btn}

        # Frame para botões de ação
        self.action_frame = tk.Frame(self.main_frame)
        self.action_frame.pack(pady=5, anchor="center")

        # Botões de ação
        tk.Button(self.action_frame, text="Criar Pastas de Escolas Selecionadas", command=self.criar_pastas_selecionadas).pack(side=tk.LEFT, padx=5)
        tk.Button(self.action_frame, text="Selecionar Tudo", command=self.selecionar_tudo).pack(side=tk.LEFT, padx=5)
        tk.Button(self.action_frame, text="Desmarcar Tudo", command=self.desmarcar_tudo).pack(side=tk.LEFT, padx=5)
        tk.Button(self.action_frame, text="Apagar Pastas", command=self.apagar_pastas).pack(side=tk.LEFT, padx=5)

        # Campo de pesquisa
        tk.Label(self.main_frame, text="Pesquisar escola ou pasta:").pack(pady=5, anchor="center")
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(self.main_frame, textvariable=self.search_var)
        self.search_entry.pack(pady=5, anchor="center")
        self.search_entry.bind("<KeyRelease>", self.filtrar_escolas)  # Conecta o evento de digitação

        # Área de log
        self.log_area = ScrolledText(self.main_frame, height=5, width=70)
        self.log_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True, anchor="center")

    def remover_acentos(self, texto):
        # Normaliza o texto para a forma NFKD (decompondo acentos) e remove os caracteres de combinação
        texto_normalizado = unicodedata.normalize('NFKD', texto)
        return ''.join(c for c in texto_normalizado if not unicodedata.combining(c))

    def log(self, mensagem):
        self.log_area.insert(tk.END, mensagem + "\n")
        self.log_area.see(tk.END)
        self.root.update()

    def criar_pasta_com_documento(self, escola):
        nome_pasta = escola.replace("/", "_").replace(":", "_").replace(",", "").replace(".", "").replace("(", "").replace(")", "")
        pasta = os.path.join(self.base_dir, nome_pasta)
        os.makedirs(pasta, exist_ok=True)
        self.log(f"Pasta criada: {pasta}")

        # Criar documento Word
        doc = Document()
        doc.add_paragraph(f"Este é um documento inicial para a escola {escola}.")
        doc_path = os.path.join(pasta, f"{nome_pasta}.docx")
        doc.save(doc_path)
        self.log(f"Documento Word criado: {doc_path}")

        # Atualizar cor do botão de pasta
        if escola in self.school_widgets:
            btn = self.school_widgets[escola]["btn"]
            btn.config(bg="#ffff99")

    def abrir_pasta(self, pasta, escola):
        if os.path.exists(pasta):
            os.startfile(pasta)
            self.log(f"Pasta aberta: {pasta}")
        else:
            messagebox.showerror("Erro", f"A pasta para {escola} ainda não foi criada.")

    def criar_pastas_selecionadas(self):
        for escola, widgets in self.school_widgets.items():
            if widgets["var"].get():  # Se a checkbox estiver marcada
                self.criar_pasta_com_documento(escola)
        self.log("Pastas e documentos criados para as escolas selecionadas.")

    def selecionar_tudo(self):
        for widgets in self.school_widgets.values():
            widgets["var"].set(True)
        self.log("Todas as escolas selecionadas.")

    def desmarcar_tudo(self):
        for widgets in self.school_widgets.values():
            widgets["var"].set(False)
        self.log("Todas as escolas desmarcadas.")

    def apagar_pastas(self):
        for escola, widgets in self.school_widgets.items():
            if widgets["var"].get():  # Se a checkbox estiver marcada
                nome_pasta = escola.replace("/", "_").replace(":", "_").replace(",", "").replace(".", "").replace("(", "").replace(")", "")
                pasta = os.path.join(self.base_dir, nome_pasta)
                if os.path.exists(pasta):
                    try:
                        shutil.rmtree(pasta)
                        self.log(f"Pasta removida: {pasta}")
                        widgets["btn"].config(bg="#ffffff")  # Atualiza a cor do botão
                    except Exception as e:
                        self.log(f"Erro ao remover {pasta}: {str(e)}")
        self.log("Pastas selecionadas foram apagadas.")

    def filtrar_escolas(self, event=None):
        query = self.remover_acentos(self.search_var.get().strip().lower())
        for escola, widgets in self.school_widgets.items():
            escola_normalizada = self.remover_acentos(escola.lower())
            # Mostra todas se não houver pesquisa ou se a escola contém o texto pesquisado
            if not query or query in escola_normalizada:
                widgets["frame"].pack(fill=tk.X, pady=2, anchor="center")
            else:
                widgets["frame"].pack_forget()
        # Ajusta a área de scroll após filtrar
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()