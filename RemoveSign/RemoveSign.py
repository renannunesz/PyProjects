import os
import tkinter as tk
import shutil
import fitz  # PyMuPDF
from tkinter import ttk, filedialog, messagebox
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from PIL import Image, ImageTk

class PDFConverterApp:
    def __init__(self, root):
        # Crie a interface gráfica
        style = Style(theme='darkly')
        self.root = root
        root = style.master
        root.title("Remover Assinaturas Sobrepostas")
        frame = ttk.Frame(root)
        frame['padding'] = (10, 10, 10, 10)
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        self.dir_pasta_selecionada = ""

        # Carregar a imagem
        self.bg_image = Image.open(r"C:\PyProjects\RemoveSign\bkg_removesign.png")
        print(os.path.abspath(__file__))
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)
        
        # Inserir a imagem no topo do formulário
        ttk.Label(frame, image=self.bg_photo).grid(row=0, column=0, columnspan=6, pady=(0, 20))

        #LabelFrame dentro do Frame - Dados Documento
        lblframe_assinaturas = ttk.LabelFrame(frame, text="Opções:", padding=10)
        lblframe_assinaturas.grid(row=1, column=0, columnspan=6, sticky="ew", padx=10, pady=5)            

        # Labels
        ttk.Label(lblframe_assinaturas, text="Arquivo PDF: ").grid(row=0, column=0, sticky=tk.W)

        # Btn        
        self.btn_selecionar = ttk.Button(lblframe_assinaturas, text="...", command=self.selecionar_pdf).grid(row=0, column=1, sticky=tk.W, padx=(0, 10))   
        self.lbl_selecionado = tk.Label(lblframe_assinaturas, text="(Selecione o arquivo)", wraplength=350)
        self.lbl_selecionado.grid(row=0, column=2, sticky=tk.W, padx=(0,10))

        # Botão
        self.btn_gerar = ttk.Button(frame, text="Gerar PDF", command=self.converter_pdf)
        self.btn_gerar.grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=10)        

    def selecionar_pdf(self):
        # Abre a janela para usuário selecionar um PDF. 
        self.pdf_path = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if self.pdf_path:
            self.lbl_selecionado.config(text=f"{os.path.basename(self.pdf_path)}")
            self.btn_gerar.config(state=tk.NORMAL)    

    def pdf_para_imagens(self, pdf_path, output_folder):
        # Converte um PDF em imagens de alta qualidade usando PyMuPDF. 
        doc = fitz.open(pdf_path)
        imagens_paths = []

        for num_pagina in range(len(doc)):
            pagina = doc[num_pagina]

            # Aumentar a qualidade da imagem (DPI 300, escala 2x)
            matrix = fitz.Matrix(2, 2)  # Escala 2x para melhorar qualidade
            pix = pagina.get_pixmap(matrix=matrix)  # Renderiza a página com maior DPI

            img_path = os.path.join(output_folder, f"pagina_{num_pagina+1}.png")
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.save(img_path, "PNG", quality=100)  # Salva com qualidade máxima
            imagens_paths.append(img_path)

        return imagens_paths

    def imagens_para_pdf(self, imagens_paths, pdf_saida):
        # Converte imagens PNG de volta para um PDF.
        imagens = [Image.open(img).convert("RGB") for img in imagens_paths]
        imagens[0].save(pdf_saida, save_all=True, append_images=imagens[1:])

    def converter_pdf(self):
        # Converte o PDF para imagens e depois reconverte para um novo PDF na mesma pasta.
        if not self.pdf_path:
            messagebox.showerror("Erro", "Nenhum arquivo PDF selecionado.")
            return

        try:
            # Diretório do arquivo original
            pasta_origem = os.path.dirname(self.pdf_path)

            # Criar pasta temporária dentro do diretório original
            pasta_temp = os.path.join(pasta_origem, "pdf_imagens")
            os.makedirs(pasta_temp, exist_ok=True)

            # Converter PDF para imagens com alta qualidade
            #messagebox.showinfo("Processando", "Convertendo PDF para imagens em alta qualidade...")
            imagens = self.pdf_para_imagens(self.pdf_path, pasta_temp)

            # Nome do novo arquivo PDF
            novo_pdf_path = os.path.join(pasta_origem, f"{os.path.splitext(os.path.basename(self.pdf_path))[0]}_VESTAS.pdf")

            # Converter imagens de volta para PDF
            #messagebox.showinfo("Processando", "Convertendo imagens de volta para PDF...")
            self.imagens_para_pdf(imagens, novo_pdf_path)

            # Apagar pasta temporária
            shutil.rmtree(pasta_temp)

            messagebox.showinfo("Sucesso", f"PDF salvo com sucesso em:\n{novo_pdf_path}")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")        

# Criar janela do Tkinter
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()