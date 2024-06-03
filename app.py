"""
Automatizar mensagens p/ meus clientes, 
Quero pode mandar mensagem oferecendo meu produto e e servios de forma automática;

"""
# 1- Passo = Descrever os passos manuais e transformar em cod;


import openpyxl #Biblioteca para acessar planilha;
from urllib.parse import quote # Permite formatar links para envio, em links expeciais;
import webbrowser # Permite acessar o navegador;
from time import sleep #Datas, pausas, calendario, etc..
import pyautogui #
# 2- Passo = Ler planilha e guardar informações sobre nome e telefone;
workbook = openpyxl.load_workbook('PlanClientes.xlsx') #Para carregar a planilha;
pagina_clientes = workbook['Planilha1'] #Página da planilha;
for linha in pagina_clientes.iter_rows(min_row=2):#Linha minima que deve começar a ler os dados;
    nome = linha[0].value # Extrai o dado da linha/indice;
    telefone = linha[1].value
    #Mensagem que será enviada..
    mensagem = f'''
Prezado(a) {nome},

Estamos entrando em contato para oferecer uma oportunidade de otimizar a presença online de seu estabelecimento no Google. Atualizar suas informações no Google é essencial para garantir que os clientes em potencial encontrem seu negócio facilmente e tenham acesso às informações corretas.

Estamos à disposição para auxiliá-lo em todo o processo de atualização. Caso tenha interesse em saber mais detalhes ou agendar uma consulta, por favor, entre em contato conosco.

Atenciosamente,

Agradecemos pela atenção!

'''
    try:
        #Criar links personalizados e enviar mensagens p/ cada cliente com base na planilha;
        link_mensagem_Wthatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_Wthatsapp) #link que irá abrir no navegador;
        sleep(15) #intervalo
        seta = pyautogui.locateCenterOnScreen('setcort.PNG')#Localiza o centro de uma img;
        sleep(5)
        pyautogui.click(seta[0], seta[1])# Clica no centro;
        sleep(5)
        pyautogui.hotkey('ctrl', 'w') #Fecha a aba automatic; hotkey vem de atalho;
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline=' ', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}', {telefone})




