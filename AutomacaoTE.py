import pyautogui
import time
import pandas as pd

# -----------------------------------------------------------------------------------------------------------------------------

pyautogui.PAUSE = 1  

planilha = pd.read_excel("PlataformaTE.xlsx")
colunas_para_deletar = ['Cliente DOC','Cidade','UF','Vendedor','Ocorrencia','DATA','1º Data PENSKE','TE','Mínimo','Máximo','Sinal','NOTA.1','14.168.536/0001-25','14168536000125']
planilha = planilha.drop(columns=colunas_para_deletar)
planilha['DATA ENTREGA'] = pd.to_datetime(planilha['DATA ENTREGA']).dt.strftime('%d/%m/%Y')
print(planilha)

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
