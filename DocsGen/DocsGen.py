import tkinter as tk
import ttkbootstrap as tb
import subprocess
from tkinter import ttk
from PIL import Image, ImageTk

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("DocsGen - Vestas - v2.060525")
        self.root.geometry("1280x720")
        self.root.resizable(False, False)

        # Estilo do ttkbootstrap
        style = tb.Style(theme="darkly")

        # Canvas com imagem de fundo
        self.canvas = tk.Canvas(self.root, width=1280, height=720)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        bkg_img = Image.open(r"C:\PyProjects\DocsGen\bkg_docsgen.png")
        bkg_img = bkg_img.resize((1280, 720), Image.Resampling.LANCZOS)
        self.bkg = ImageTk.PhotoImage(bkg_img)
        self.canvas.create_image(0, 0, image=self.bkg, anchor="nw")

        # Frame principal por cima do canvas
        self.frame_principal = ttk.Frame(self.canvas, padding=20)
        self.frame_principal.place(relx=0.5, rely=0.5, anchor="center")  # Centraliza no canvas

        # Label de boas-vindas
        #self.label = ttk.Label(self.frame_principal, text="DocsGen - Ferramenta de Geração de Documentos", font=("Arial", 14, "bold"))
        #self.label.grid(row=0, column=0, pady=50)

        # Botão de saída
        #self.btn_sair = tb.Button(self.frame_principal, text="Sair", bootstyle="danger", command=self.root.quit)
        #self.btn_sair.grid(row=1, column=0, pady=10)

        # Menu
        menubar = tk.Menu(self.root)

        menu_anuencias = tk.Menu(menubar, tearoff=0)
        #menu_anuencias.add_command(label="Gerar Nova", command=self.abrir_anuencias)
        menu_anuencias.add_command(label="Tecnico de O&M", command=self.anuencias_tec_om)
        menu_anuencias.add_command(label="Supervisor de O&M", command=self.anuencias_sup_om)
        menu_anuencias.add_command(label="Tecnico de Pá", command=self.anuencias_tec_pa)

        menu_os = tk.Menu(menubar, tearoff=0)
        #menu_os.add_command(label="Gerar Nova", command=self.abrir_os)
        menu_os.add_command(label="Gerar OS", command=self.gerador_os)        

        menu_sit = tk.Menu(menubar, tearoff=0)
        menu_sit.add_command(label="Gerar Novo", command=self.abrir_sit)

        menu_assinaturas = tk.Menu(menubar, tearoff=0)
        menu_assinaturas.add_command(label="Remover Assinaturas", command=self.abrir_assinaturas)

        menubar.add_cascade(label="Anuências", menu=menu_anuencias)
        menubar.add_cascade(label="OS", menu=menu_os)
        menubar.add_cascade(label="SIT", menu=menu_sit)
        menubar.add_cascade(label="Assinaturas", menu=menu_assinaturas)
        menubar.add_command(label="Sair", command=self.root.quit)

        self.root.config(menu=menubar)

    def abrir_anuencias(self):
        print("Abrir tela de Anuências")
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DocsGen_Anuencias.py"]) 

    def anuencias_tec_om(self):
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DG_AnuenciasTecOM.py"])

    def anuencias_sup_om(self):
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DG_AnuenciasSupOM.py"])

    def anuencias_tec_pa(self):
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DG_AnuenciasTecPA.py"])                

    def abrir_os(self):
        print("Abrir tela de OS")
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DocsGen_OS.py"])

    def gerador_os(self):
        print("Abrir tela de OS")
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DG_OS.py"])        

    def abrir_sit(self):
        print("Abrir tela de SIT")
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DG_SIT.py"])

    def abrir_assinaturas(self):
        print("Abrir tela de Assinaturas")
        subprocess.Popen(["pythonw", r"C:\PyProjects\DocsGen\DG_RemoveSign.py"])

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
