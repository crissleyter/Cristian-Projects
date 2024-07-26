import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import time
from tkinter import ttk
from threading import Thread

# Função para enviar e-mails
def enviar_email(destinatario, assunto, corpo, anexo=None):
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = destinatario
    email.Subject = assunto
    email.HTMLBody = corpo
    if anexo:
        email.Attachments.add(anexo)
    email.Send()

# Função para enviar e-mails em lote com intervalo de 5 segundos
def enviar_emails_thread():
    global df

    assunto = entrada_assunto.get()
    corpo = texto_corpo.get("1.0", tk.END)

    if df.empty:
        messagebox.showwarning("Aviso", "Nenhuma planilha foi importada.", parent=root)
        return

    progresso['maximum'] = len(df)
    status['text'] = "Enviando e-mails..."

    # Desabilita botões durante o envio
    btn_importar.config(state=tk.DISABLED)
    btn_enviar.config(state=tk.DISABLED)

    for index, row in df.iterrows():
        destinatario = row['Email']
        nome = row['Nome']

        corpo_personalizado = f'''
        <p>Olá {nome}, tudo bem?</p>
        <p>{corpo}</p>
        <p></p>
        <p></p>
        <p>Caso não queira mais receber nossas comunicações, basta responder a este e-mail com a mensagem "Não quero receber comunicação".
        '''

        #anexo = "caminho/para/seu/anexo.pdf"
        #enviar_email(destinatario, assunto, corpo_personalizado, anexo)
        enviar_email(destinatario, assunto, corpo_personalizado)
        status['text'] = f"Enviando e-mail para {destinatario}..."
        progresso['value'] = index + 1
        root.update()  # Atualiza a interface gráfica

        time.sleep(5)

    messagebox.showinfo("Sucesso", "Todos os e-mails foram enviados!", parent=root)

    # Limpa campos de assunto e corpo após envio
    entrada_assunto.delete(0, tk.END)
    texto_corpo.delete("1.0", tk.END)

    # Habilita botões após envio
    btn_importar.config(state=tk.NORMAL)
    btn_enviar.config(state=tk.NORMAL)

    progresso['value'] = 0
    status['text'] = ""

def enviar_emails():
    thread = Thread(target=enviar_emails_thread)
    thread.start()

# Função para importar planilha Excel
def importar_planilha():
    global df
    arquivo_excel = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx;*.xls")])
    if arquivo_excel:
        try:
            df = pd.read_excel(arquivo_excel)

            if 'Nome' in df.columns and 'Email' in df.columns:
                messagebox.showinfo("Sucesso",
                                    "Planilha importada com sucesso!\nCertifique-se de que a planilha contém as colunas 'Nome' e 'Email'.",
                                    parent=root)
            else:
                messagebox.showwarning("Aviso", "A planilha deve conter as colunas 'Nome' e 'Email'.", parent=root)

        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler a planilha:\n{str(e)}", parent=root)

# Configuração da GUI
root = tk.Tk()
root.title("Envios de e-mails | DRTECHS")
root.geometry("800x600")

cor_azul = "#0078D4"
cor_marine = "#014454"
cor_branco = "#FFFFFF"
fonte = "Aptos SemiBold"

estilo = ttk.Style()
estilo.configure('TButton', foreground=cor_azul, background=cor_azul, font=(fonte, 12), padding=10, borderwidth=110, relief="flat")
estilo.map('TButton', background=[('active', cor_azul)])
estilo.configure('TLabel', foreground=cor_branco, background=cor_marine, font=(fonte, 12))
estilo.configure('TFrame', background=cor_marine)
estilo.configure('TEntry', font=('Arial', 12), padding=5)

frame_email = ttk.Frame(root, padding="40")
frame_email.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

ttk.Label(frame_email, text="Assunto do E-mail:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
entrada_assunto = ttk.Entry(frame_email, width=50)
entrada_assunto.grid(row=0, column=1, padx=10, pady=10)

ttk.Label(frame_email, text="Corpo do E-mail:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
texto_corpo = tk.Text(frame_email, width=50, height=10)
texto_corpo.grid(row=1, column=1, padx=10, pady=10)

btn_importar = ttk.Button(frame_email, text="Importar Planilha", command=importar_planilha)
btn_importar.grid(row=2, column=0, columnspan=2, pady=10)

btn_enviar = ttk.Button(frame_email, text="Enviar E-mails", command=enviar_emails)
btn_enviar.grid(row=3, column=0, columnspan=2, pady=10)

# Ícone para o botão de enviar e-mails
img_enviar = tk.PhotoImage(file='send.png')
btn_enviar.config(image=img_enviar, compound=tk.LEFT)

# Barra de progresso e status
ttk.Label(frame_email, text="Progresso:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=10)
progresso = ttk.Progressbar(frame_email, orient=tk.HORIZONTAL, length=400, mode='determinate')
progresso.grid(row=4, column=1, padx=10, pady=10)

status = ttk.Label(frame_email, text="", anchor=tk.W)
status.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

desenvolvido = ttk.Label(frame_email, text="Desenvolvido por Cristian Duarte", anchor=tk.W)
desenvolvido.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()

