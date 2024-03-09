# Passo a passo do projeto
# Passo 1: Entrar no sistema da empresa 

import pyautogui
import time

# pyautogui.write -> escrever um texto
# pyautogui.press -> apertar 1 tecla
# pyautogui.click -> clicar em algum lugar da tela
# pyautogui.hotkey -> combinação de teclas

# -----------------------------------------------------------------------------------------------------------------------------

pyautogui.PAUSE = 1  

    # # abrir o navegador (chrome)
    # pyautogui.press("win")
    # pyautogui.write("chrome")
    # pyautogui.press("enter")

    # # entrar no link da plataforma
    # pyautogui.click(x=219, y=58)
    # pyautogui.write("https://admin.plataformabmc.com.br/")
    # pyautogui.press("enter")
    # time.sleep(6)

    # # Passo 2: Fazer login e acessar o painel da logística
    # # selecionar o campo da loja
    
    # pyautogui.click(x=833, y=413)   
    # pyautogui.write("229")
    # pyautogui.press("tab")
    # # escrever o seu email
    # pyautogui.write("logistica@bmchyundai.com.br")
    # pyautogui.press("tab") # passando pro próximo campo
    # # escrever o senha
    # pyautogui.write("123")
    # pyautogui.press("tab") # passando pro próximo campo
    # pyautogui.press("enter") # confirmando para entrar
    # time.sleep(3)
    # pyautogui.scroll(-200)
    # pyautogui.click(x=131, y=697)
    
    # pyautogui.scroll(-300)
    # pyautogui.click(x=145, y=700)
    # pyautogui.click(x=1239, y=418)
    
    # pyautogui.click(x=1236, y=478)
    # pyautogui.click(x=587, y=490)
    # pyautogui.click(x=550, y=554)
    # pyautogui.click(x=446, y=568)
    # pyautogui.scroll(-300)

    # # Posição de um clique para colocar as notas fiscais em ordem
    # pyautogui.click(x=774, y=549) 

# Passo 3: Importar a base com as informações do TE
import pandas as pd

planilha = pd.read_excel("PlataformaTE.xlsx")
colunas_para_deletar = ['Cliente DOC','Cidade','UF','Ocorrencia','DATA','1º Data PENSKE','TE','Mínimo','Máximo','Sinal','NOTA.1','14.168.536/0001-25','14168536000125']
planilha = planilha.drop(columns=colunas_para_deletar)
planilha['DATA ENTREGA'] = pd.to_datetime(planilha['DATA ENTREGA']).dt.strftime('%d/%m/%Y')
print(planilha)

# Passo 4: Finalizar as notas fiscais entregues com a data de entrega de cada nota fiscal ou inserir a observação atual de status

# Utilize a linha de código abaixo, caso a plataforma do TE esteja aberta e não tenha entrado na plataforma pelo código... 
# Caso contrário, a linha de código abaixo, deve estar comentada;
pyautogui.hotkey('alt', 'tab') 

for CadaLinha, row in planilha.iterrows():
    Observacao = row['Observação']
    
    if "MERCADORIA ENTREGUE" in Observacao:
        Nota = planilha.loc[CadaLinha, "NOTA"]
        DataEntrega = planilha.loc[CadaLinha, "DATA ENTREGA"]
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.write(str(Nota))
        pyautogui.press("enter") 
        pyautogui.press("tab") 
        pyautogui.press("tab") 
        pyautogui.press("tab") 
        pyautogui.press("enter")
        pyautogui.press("tab") 
        pyautogui.press("enter")
        pyautogui.press("enter")
        pyautogui.write(str(DataEntrega))
        pyautogui.press("tab") 
        pyautogui.press("enter")
        pyautogui.sleep(2)
        pyautogui.press("enter")
        print(DataEntrega)    
    else:
        Nota = planilha.loc[CadaLinha, "NOTA"]
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.write(str(Nota))
        pyautogui.press("enter")    
        pyautogui.press("tab") 
        pyautogui.press("tab") 
        pyautogui.press("tab") 
        pyautogui.press("enter")
        pyautogui.press("enter")
        pyautogui.press("tab") 
        pyautogui.press("tab") 
        pyautogui.hotkey('Shift', 'F10')
        pyautogui.press("down") 
        
        pyautogui.press("enter")        
        pyautogui.hotkey('ctrl', 'tab')
        
        #pyautogui para selecionar o campo "Outros"
        pyautogui.click(x=804, y=517)
        pyautogui.click(x=613, y=680)
        
        #pyautogui para selecionar a data de hoje
        pyautogui.click(x=1209, y=518)
        pyautogui.press("tab") 
        pyautogui.press("tab") 
        pyautogui.press("enter")
        
        #pyautogui para selecionar o campo de observação
        pyautogui.click(x=528, y=711)
        pyautogui.write(str(Observacao))
        
        #pyautogui para confirmar a informação inserida
        pyautogui.sleep(1)
        pyautogui.press("tab") 
        pyautogui.press("enter")
        pyautogui.click(x=853, y=189)

        
        #pyautogui para voltar a página
        pyautogui.click(x=493, y=17)
            
    # EntregaOuObservacao(Nota, Observacao, DataEntrega)
    # pyautogui.hotkey('ctrl', 'f') 

# # Passo 4: Marcar como entregue ou inserir a observação
# for linha in tabela.index:
#     # clicar no campo de código
#     pyautogui.click(x=653, y=294)
#     # pegar da tabela o valor do campo que a gente quer preencher
#     codigo = tabela.loc[linha, "codigo"]
#     # preencher o campo
#     pyautogui.write(str(codigo))
#     # passar para o proximo campo
#     pyautogui.press("tab")
#     # preencher o campo
#     pyautogui.write(str(tabela.loc[linha, "marca"]))
#     pyautogui.press("tab")
#     pyautogui.write(str(tabela.loc[linha, "tipo"]))
#     pyautogui.press("tab")
#     pyautogui.write(str(tabela.loc[linha, "categoria"]))
#     pyautogui.press("tab")
#     pyautogui.write(str(tabela.loc[linha, "preco_unitario"]))
#     pyautogui.press("tab")
#     pyautogui.write(str(tabela.loc[linha, "custo"]))
#     pyautogui.press("tab")
#     obs = tabela.loc[linha, "obs"]
#     if not pd.isna(obs):
#         pyautogui.write(str(tabela.loc[linha, "obs"]))
#     pyautogui.press("tab")
#     pyautogui.press("enter") # cadastra o produto (botao enviar)
#     # dar scroll de tudo pra cima
#     pyautogui.scroll(5000)
#     # Passo 5: Repetir o processo de cadastro até o fim
