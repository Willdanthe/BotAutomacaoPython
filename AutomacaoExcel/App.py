#Bibliotecas=-=-=-=
import openpyxl # manipulação com Excel
from urllib.parse import quote #Converte a mensagem para link
import webbrowser # Abre o navegador padrão
from time import sleep
import pyautogui # Precisa instalar a biblitoca PIL, que só está disponível na versão 3.13 do python. A versão 3.14 não suporta ainda
#Bibliotecas=-=-=-=






# (NOT WORKING)seta = pyautogui.locateCenterOnScreen('seta.png', confidence=0.2) # Não esta conseguindo encontrar a imgagem :,)))))))) Está indo para o icone de conversa

workbook = openpyxl.load_workbook('TestedeAutomacao.xlsx') # Abrindo o arquivo Excel para manipular
pagina_clientes = workbook['pag1'] # Acessando a página 1

# Iterando as linhas do arquivo
for linha in pagina_clientes.iter_rows(min_row=3): # Parâmetro diz que vai começar da linha 3
    
    # Quebra se a linha estiver nula os 3 campos
    if(linha[0].value == None or linha[1].value == None or linha[2].value == None ): 
        break

    # o index de linha é as colunas
    nome_crianca = linha[0].value
    telefone = linha[1].value # O número tem que ter 55 (Brasil) e o DDD
    responsavel = linha[2].value

    # Criar link personalizados do whatsapp e enviar mensagaem 
    # para cada pessoas com base nos dados da planílha
    mensagem = f'Olá, {nome_crianca} Fulano Tal tal tal'

    # Conversão compatível com o URL
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    

    webbrowser.open(link_mensagem_whatsapp)
    sleep(10) # Tempo de espera (Em segundos) para carregar o WhatsAppWeb 

    pyautogui.click(1849,957) # Clicando no botão de enviar (configuração pode mudar de dipositivo pra dispositivo)
    sleep(1)
    pyautogui.hotkey('ctrl','w')
    sleep(1)

# # Teste

# nome = 'criança tal'
# telefone = '5599999999999'
# responsavel = 'responsável tal'
# mensagem = f'Olá, boa tarde {responsavel}!\nEu sou o *Willdanthê*, secretário do Departamento Infanfil IEQ BAND! Estou entrando em contato pois vi que a criança {nome} não vem a nossa igreja faz tempo. Estamos com programações novas para as nossas crianças! Como Balé, Street Dance, aulas de músicas, eventos para as crianças se divertirem e aprender mais sobre Cristo...'
# link_mensagem_atizap = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
# webbrowser.open(link_mensagem_atizap)
# sleep(10)
# pyautogui.click(1849,957) #Point(x=1849, y=957) Configuração pode mudar de dipositivo pra dispositivo

# # sleep(2)
# # pyautogui.click(seta.x,seta.y)
# sleep(2)
# pyautogui.hotkey('ctrl','w')
# sleep(2)

