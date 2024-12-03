import os
from tkinter import Tk, filedialog, Button, Label, StringVar, OptionMenu
from docx import Document
from PIL import Image
from reportlab.pdfgen import canvas

def get_download_folder():
    """Retorna o caminho da pasta Downloads do usuário."""
    if os.name == "nt":  # Windows
        return os.path.join(os.environ["USERPROFILE"], "Downloads")
    else:  # macOS e Linux
        return os.path.join(os.path.expanduser("~"), "Downloads")

def select_file():
    global selected_file
    selected_file = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Todos os Arquivos", "*.*")]
    )
    if selected_file:
        label_status.config(text=f"Arquivo selecionado: {os.path.basename(selected_file)}")
    else:
        label_status.config(text="Nenhum arquivo selecionado.")

def convert_file():
    if not selected_file:
        label_status.config(text="Por favor, selecione um arquivo primeiro!")
        return

    if not conversion_type.get():
        label_status.config(text="Por favor, escolha um formato para converter!")
        return

    try:
        ext = os.path.splitext(selected_file)[1].lower()
        download_folder = get_download_folder()
        output_file = os.path.join(download_folder, f"{os.path.basename(os.path.splitext(selected_file)[0])}_conv.{conversion_type.get()}")

        if conversion_type.get() == "pdf":
            if ext == ".docx":
                doc = Document(selected_file)
                pdf = canvas.Canvas(output_file)
                pdf.setFont("Times-Roman", 12)
                for paragraph in doc.paragraphs:
                    text = paragraph.text
                    pdf.drawString(100, pdf._pagesize[1] - 100, text)
                    pdf.showPage()
                pdf.save()
            elif ext in [".jpg", ".jpeg", ".png"]:
                image = Image.open(selected_file)
                image = image.convert("RGB")
                image.save(output_file, "PDF")
            else:
                label_status.config(text="Formato não suportado para PDF!")
                return
        elif conversion_type.get() in ["jpeg", "jpg", "png"]:
            if ext in [".jpg", ".jpeg", ".png"]:
                image = Image.open(selected_file)
                image.save(output_file, conversion_type.get().upper())
            else:
                label_status.config(text="Formato não suportado para conversão em imagem!")
                return
        elif conversion_type.get() == "docx":
            if ext == ".pdf":
                label_status.config(text="Conversão de PDF para Word não implementada.")
                return
            else:
                label_status.config(text="Formato não suportado para conversão em Word!")
                return
        else:
            label_status.config(text="Opção de conversão inválida!")
            return

        label_status.config(text=f"Convertido com Sucesso! Salvo em {download_folder}")
    except Exception as e:
        label_status.config(text=f"Erro na conversão: {str(e)}")

# Interface gráfica com tkinter
root = Tk()
root.title("Conversor de Arquivos")

selected_file = ""
conversion_type = StringVar(root)
conversion_type.set("")  # Nenhuma opção selecionada inicialmente

label_instructions = Label(root, text="Selecione um arquivo e o formato para conversão", font=("Arial", 12))
label_instructions.pack(pady=10)

btn_select_file = Button(root, text="Selecionar Arquivo", command=select_file, width=20)
btn_select_file.pack(pady=5)

label_conversion = Label(root, text="Escolha o formato para converter", font=("Arial", 10))
label_conversion.pack(pady=5)

# Dropdown para selecionar o tipo de conversão
options = ["pdf", "jpeg", "png", "jpg", "docx"]
dropdown = OptionMenu(root, conversion_type, *options)
dropdown.pack(pady=5)

btn_convert = Button(root, text="Converter", command=convert_file, width=20)
btn_convert.pack(pady=5)

label_status = Label(root, text="", font=("Arial", 14), fg="blue")
label_status.pack(pady=10)

root.mainloop()
