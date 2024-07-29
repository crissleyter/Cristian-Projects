import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import urllib.parse
import win32com.client as win32
import pyautogui

# Ler o arquivo Excel
contato_df = pd.read_excel(r'C:\Users\c_sle\OneDrive\Área de Trabalho\Automação disparo\enviar.xlsx')

print(contato_df)

# Inicializar o navegador
options = webdriver.ChromeOptions()
options.add_argument(r"webdriver.chrome.driver=r'C:\Users\c_sle\OneDrive\Área de Trabalho\chromedriver.exe")
navegador = webdriver.Chrome(options=options)

# Abrir o WhatsApp Web
navegador.get('https://web.whatsapp.com/')

whatswait = WebDriverWait(navegador, 1000).until(
    EC.visibility_of_element_located((By.XPATH,
                                      "/html/body/div[1]/div/div/div[2]/div[3]/div/div[1]/div/div[2]/div[2]"))
)

# Iterar sobre os contatos
for i in range(len(contato_df)):
    pessoa = contato_df.loc[i, "Nome"]
    numero = contato_df.loc[i, "Telefone"]
    cpf = contato_df.loc[i, "CPF"]
    emaildest = contato_df.loc[i, "E-mail"]
    mensagem = f'''
Olá, tudo bem?

Sou Cristian, e falo em nome da 3 Corações.

Se você é o(a) sr(a){pessoa}, portador(a) do CPF final xxx.xxx.{cpf}, gostaríamos de PARABENIZÁ-LO(A), pois você foi um(a) dos(as) ganhadores(as) do prêmio de 1(um) Crédito através da carteira digital PicPay, sem função saque/transferência, no valor de R$12.000,00, na Promoção "Frisco com mais gostosura".

Pedimos que confira seu nome no site da promoção www.frisco.com.br/promo.

Enviamos um e-mail com o remetente entregadepremios@veramouraadvocacia.adv.br para o seu e-mail cadastrado na promoção. 

Neste e-mail consta o termo chamado "Recibo de Entrega do Prêmio" para assinatura, ao qual você terá que incluir um documento legível, com foto, contendo seu RG e CPF, no prazo de até 3 dias a partir do recebimento desta mensagem, para que possamos dar continuidade na entrega do seu prêmio.

Caso não tenha recebido, verifique na sua caixa de spam ou lixo eletrônico.

IMPORTANTE
Se você já possui uma conta Picpay em seu nome, você receberá o crédito nesta conta em até 30 dias. 

Caso não possua uma conta em seu nome, baixe o aplicativo Picpay, na Loja de aplicativos de seu celular e abra sua conta. 

Depois de aberta a conta, o crédito entrará automaticamente nesta conta em até 30 dias.

Após o recebimento e validação de seus documentos, basta aguardar o crédito em sua conta PICPAY nos prazos citado acima.

Caso tenha dúvidas, estamos à disposição neste mesmo canal de atendimento ou através do e-mail entregadepremios@veramouraadvocacia.adv.br.

Atenciosamente, 
Promoção "Frisco com mais gostosura".
'''

    # Codificar a mensagem para URL
    texto_codificado = urllib.parse.quote(mensagem)

    # Construir o link com a mensagem codificada
    link = f'https://web.whatsapp.com/send?phone=55{numero}&text={texto_codificado}'

    # Abrir o link no navegador
    navegador.get(link)
    #sleep(10)
    try:
        campo_texto = WebDriverWait(navegador, 30).until(
            EC.visibility_of_element_located((By.XPATH,
                                              "/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p"))
        )
        sleep(1)
        campo_texto.click()
        pyautogui.press('enter')
        sleep(2)
    except:
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = emaildest
        email.Subject = 'Promocao frisco com mais gostosura'
        email.HTMLBody = f'''
<p>Olá, tudo bem?</p>
<p></p>
<p>Sou Cristian, e falo em nome da 3 Corações.</p>
<p></p>
<p>Se você é o(a) sr(a){pessoa}, portador(a) do CPF final xxx.xxx.{cpf}, gostaríamos de PARABENIZÁ-LO(A), pois você foi um(a) dos(as) ganhadores(as) do prêmio de 1(um) Crédito através da carteira digital PicPay, sem função saque/transferência, no valor de R$12.000,00, na Promoção "Frisco com mais gostosura".</p>
<p></p>
<p>Pedimos que confira seu nome no site da promoção www.frisco.com.br/promo.</p>
<p></p>
<p>Enviamos um e-mail com o remetente entregadepremios@veramouraadvocacia.adv.br para o seu e-mail cadastrado na promoção. </p>
<p></p>
<p>Neste e-mail consta o termo chamado "Recibo de Entrega do Prêmio" para assinatura, ao qual você terá que incluir um documento legível, com foto, contendo seu RG e CPF, no prazo de até 3 dias a partir do recebimento desta mensagem, para que possamos dar continuidade na entrega do seu prêmio.</p>
<p></p>
<p>Caso não tenha recebido, verifique na sua caixa de spam ou lixo eletrônico.</p>
<p></p>
<p>IMPORTANTE</p>
<p>Se você já possui uma conta Picpay em seu nome, você receberá o crédito nesta conta em até 30 dias. </p>
<p></p>
<p>Caso não possua uma conta em seu nome, baixe o aplicativo Picpay, na Loja de aplicativos de seu celular e abra sua conta. </p>
<p></p>
<p>Depois de aberta a conta, o crédito entrará automaticamente nesta conta em até 30 dias.</p>
<p></p>
<p>Após o recebimento e validação de seus documentos, basta aguardar o crédito em sua conta PICPAY nos prazos citado acima.</p>
<p></p>
<p>Caso tenha dúvidas, estamos à disposição neste mesmo canal de atendimento ou através do e-mail entregadepremios@veramouraadvocacia.adv.br.</p>
<p></p>
<p>Atenciosamente, </p>
<p>Promoção "Frisco com mais gostosura".</p>'''
        email.Send()
        print(f'E-mail enviado para {pessoa}: {emaildest} ')