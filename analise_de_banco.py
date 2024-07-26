import pandas as pd
import re
from tkinter import filedialog
import tkinter as tk
import numpy as np

def validar_cpf_formula(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    if len(cpf) != 11:
        return False, "CPF deve ter exatamente 11 dígitos"

    total, multiplicador = 0, 10
    for digito in cpf[:9]:
        total += int(digito) * multiplicador
        multiplicador -= 1
    resto = total % 11
    digito_verif_1 = 0 if resto < 2 else 11 - resto

    total, multiplicador = 0, 11
    for digito in cpf[:10]:
        total += int(digito) * multiplicador
        multiplicador -= 1
    resto = total % 11
    digito_verif_2 = 0 if resto < 2 else 11 - resto

    if int(cpf[9]) != digito_verif_1 or int(cpf[10]) != digito_verif_2:
        return False, "Dígitos verificadores incorretos"

    return True, "CPF válido"

def validar_email(email):
    # Verifica se o valor é NaN
    if isinstance(email, float) and np.isnan(email):
        return False
    # Usando uma expressão regular simples para validar o formato do e-mail
    if re.match(r"[^@]+@[^@]+\.[^@]+", email):
        return True
    else:
        return False

def validar_telefone(telefone):
    # Verifica se o valor é NaN
    if isinstance(telefone, float) and np.isnan(telefone):
        return False
    # Verifica se o telefone tem 10 ou 11 dígitos numéricos
    telefone_str = str(telefone)
    if len(telefone_str) == 10 and telefone_str.isdigit():
        return True
    elif len(telefone_str) == 11 and telefone_str[2] == '9' and telefone_str.isdigit():
        return True
    else:
        return False

def limpar_nome_cliente(cliente):
    # Utiliza expressão regular para remover números e caracteres especiais
    return re.sub(r'[^a-zA-Z\s]', '', str(cliente))

def selecionar_planilha():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    caminho_planilha = filedialog.askopenfilename(title="Selecione a planilha a ser conferida", filetypes=[("Excel Files", "*.xlsx")])
    return caminho_planilha

# Função para escrever no relatório e fechar o arquivo
def escrever_e_fechar_relatorio(relatorio, texto):
    relatorio.write(texto)
    relatorio.close()

with open('relatorio.txt', 'w') as relatorio:
    caminho_planilha = selecionar_planilha()  # Chama a função para selecionar a planilha
    df_before = pd.read_excel(caminho_planilha)
    df_before = df_before.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    registros_alterados = 0

    def imprimir_validacao_cpf(cpf, relatorio):
        global registros_alterados
        cpf = ''.join(filter(str.isdigit, str(cpf)))
        cpf_formatado = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        valido, motivo = validar_cpf_formula(cpf)

        if not valido:
            mensagem = f"CPF: {cpf_formatado} - Inválido - Motivo: {motivo}\n"
            relatorio.write(mensagem)

            if len(cpf) != 11:
                relatorio.write(f"  - Menos de 11 dígitos\n")

            if cpf_formatado != f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}":
                relatorio.write(f"  - Formatação diferente de 000.000.000-00\n")

            # Incrementa o contador de registros alterados
            registros_alterados += 1

        return valido

    def imprimir_validacao_email(email, relatorio):
        valido = validar_email(email)

        if not valido:
            mensagem = f"E-mail: {email} - Inválido\n"
            relatorio.write(mensagem)

        return valido

    def imprimir_validacao_telefone(telefone, relatorio):
        valido = validar_telefone(telefone)

        if not valido:
            mensagem = f"Telefone: {telefone} - Inválido\n"
            relatorio.write(mensagem)

        return valido

    df_before.rename(columns={'cpf': 'CPF_ID', 'telefone': 'Telefone', 'email': 'E-mail', 'cliente': 'Cliente'}, inplace=True)

    df_before['Cliente'] = df_before['Cliente'].apply(limpar_nome_cliente)  # Remove caracteres especiais da coluna Cliente

    df_before['CPF_valido'] = df_before['CPF_ID'].apply(lambda x: imprimir_validacao_cpf(x, relatorio))
    qtd_cpfs_invalidos = df_before['CPF_valido'].eq(False).sum()

    if 'E-mail' in df_before.columns:
        df_before['E-mail_valido'] = df_before['E-mail'].apply(lambda x: imprimir_validacao_email(x, relatorio))
        qtd_emails_invalidos = df_before['E-mail_valido'].eq(False).sum()
    else:
        qtd_emails_invalidos = 0

    if 'Telefone' in df_before.columns:
        df_before['Telefone_valido'] = df_before['Telefone'].apply(lambda x: imprimir_validacao_telefone(x, relatorio))
        qtd_telefones_invalidos = df_before['Telefone_valido'].eq(False).sum()
    else:
        qtd_telefones_invalidos = 0

    celulas_vazias_por_coluna = df_before.isna().sum()

    # Realiza o processamento posterior
    df_after = df_before.copy()

    # Exportar o DataFrame para o arquivo Excel
    df_after.to_excel(r'C:\Users\c_sle\OneDrive\Área de Trabalho\banco_ajustado.xlsx', index=False)

    relatorio.write("\nRegistros alterados:\n")
    relatorio.write("------------ ANTES -------------\n")
    relatorio.write(df_before.to_string() + '\n')
    relatorio.write("------------ DEPOIS -------------\n")
    relatorio.write(df_after.to_string() + '\n')
    relatorio.write("------------------\n")
    relatorio.write(f"Registros alterados: {registros_alterados}\n")
    relatorio.write(f"CPFs inválidos ou fora de formatação: {qtd_cpfs_invalidos}\n")
    relatorio.write(f"E-mails inválidos: {qtd_emails_invalidos}\n")
    relatorio.write(f"Telefones inválidos: {qtd_telefones_invalidos}\n")
    relatorio.write("------------------\n")
    relatorio.write("Relatório de células vazias por coluna:\n")
    relatorio.write(celulas_vazias_por_coluna.to_string() + '\n')

    # Fechar o arquivo de relatório após a escrita
    escrever_e_fechar_relatorio(relatorio, "")



