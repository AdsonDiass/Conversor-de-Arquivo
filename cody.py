import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def pdf_to_txt(pdf_path, output_txt):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            with open(output_txt, 'w', encoding='utf-8') as txt_file:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        txt_file.write(text)
        messagebox.showinfo("Sucesso", f"PDF convertido para {output_txt} com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def pdf_to_excel(pdf_path, output_excel):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_data = []
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split("\n")
                    for line in lines:
                        all_data.append([line])
            df = pd.DataFrame(all_data, columns=["Conteúdo"])
            df.to_excel(output_excel, index=False)
        messagebox.showinfo("Sucesso", f"PDF convertido para {output_excel} com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def escolher_pdf():
    filepath = filedialog.askopenfilename(
        filetypes=[("Arquivos PDF", "*.pdf")],
        title="Selecione o arquivo PDF"
    )
    return filepath

def salvar_arquivo(extensao):
    filepath = filedialog.asksaveasfilename(
        defaultextension=extensao,
        filetypes=[("Arquivos de Texto", "*.txt")] if extensao == ".txt" else [("Arquivos Excel", "*.xlsx")],
        title="Escolha o nome do arquivo de saída"
    )
    return filepath

def main():
    # Criação da janela principal
    root = tk.Tk()
    root.title("Conversor de PDF")
    root.geometry("300x150")
    
    def converter_para_txt():
        pdf_path = escolher_pdf()
        if pdf_path:
            output_txt = salvar_arquivo(".txt")
            if output_txt:
                pdf_to_txt(pdf_path, output_txt)

    def converter_para_excel():
        pdf_path = escolher_pdf()
        if pdf_path:
            output_excel = salvar_arquivo(".xlsx")
            if output_excel:
                pdf_to_excel(pdf_path, output_excel)
    
    # Botões da interface
    btn_txt = tk.Button(root, text="Converter PDF para TXT", command=converter_para_txt)
    btn_txt.pack(pady=10)

    btn_excel = tk.Button(root, text="Converter PDF para Excel", command=converter_para_excel)
    btn_excel.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()
