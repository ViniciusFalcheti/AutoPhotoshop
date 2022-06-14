from numpy import False_
import win32com.client
import os
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from datetime import date, datetime

def exists(navegador, elem):
    try:
        teste = navegador.find_element_by_xpath(elem)
        return True
    except NoSuchElementException:
        return False

def saveJpg(jpg = r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste", the = False):
    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")

    options.Format = 6
    options.Quality = 100

    if(rodada == 'BRASILEIRÃO SÉRIE A'):
         bra = "serieA"
    elif(rodada == 'BRASILEIRÃO SÉRIE B'):
        bra = "serieB"
    elif(rodada == 'BRASILEIRÃO SÉRIE C'):
        bra = "serieC"


    jpgfile = jpg + f"-{bra}.jpg"
    if(the == True):
        jpgfile = jpg + f"-{bra}-the.jpg" 

    doc.Export(ExportIn=jpgfile, ExportAs = 2, Options = options)

def closePSD(doc):
    doc.Save()
    doc.Application.ActiveDocument.Close(2)

def tiraLogo(jpg):
    layerLogo = doc.ArtLayers[f"LOGO"]
    layerFundo = doc.ArtLayers[f"LOGOFUNDO"]
    layerLogo.Visible = False
    layerFundo.Visible = False

    saveJpg(jpg, True)

def colocaLogo():
    layerLogo = doc.ArtLayers[f"LOGO"]
    layerFundo = doc.ArtLayers[f"LOGOFUNDO"]
    layerLogo.Visible = True
    layerFundo.Visible = True

def criaConfrontos():
    rodada = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/nav/span[2]').text
    print(rodada)
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f"brasileirão - {rodada}"

    # print(navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/ul/li[10]/div/div/div/div[1]/span[1]').text)

    for i in range(10):
        if(exists(navegador, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]')):
            str_date = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text
        else:
            str_date = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text
        
        str_date = str_date[-10:]
        data = datetime.strptime(str_date, '%d/%m/%Y').date()

        print(data)

        if (data < date.today()):
            p1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[1]').text
            p2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[5]').text
            nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[1]/span[1]')
            nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[3]/span[1]')
            local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[3]').text
            layerGols1 = doc.ArtLayers[f"RES{i+1}TIME1"]
            layerGols2 = doc.ArtLayers[f"RES{i+1}TIME2"]
            layerGols1.Visible = True
            layerGols2.Visible = True
            text_of_Gols1 = layerGols1.TextItem
            text_of_Gols2 = layerGols2.TextItem
            text_of_Gols1.contents = p1
            text_of_Gols2.contents = p2
        else:
            p1 = 0
            p2 = 0
            nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[1]/span[1]')                                         
            nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[3]/span[1]')
            local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[3]').text 
            layerGols1 = doc.ArtLayers[f"RES{i+1}TIME1"]
            layerGols2 = doc.ArtLayers[f"RES{i+1}TIME2"]
            layerGols1.Visible = False
            layerGols2.Visible = False

        layerLocal = doc.ArtLayers[f"LOCAL{i+1}"]
        text_of_Local = layerLocal.TextItem
        text_of_Local.contents = local
        nt1 = nt1.get_attribute("title")
        nt2 = nt2.get_attribute("title")
        # print(f"{nt1} {p1} x {p2} {nt2}")

        if(nt1.upper() == "BRASIL DE PELOTAS"):
            nt1 = "BRASIL-RS"
        if(nt2.upper() == "BRASIL DE PELOTAS"):
            nt2 = "BRASIL-RS"

        layerTime1 = doc.ArtLayers[f"JOGO{i+1}TIME1"]
        layerTime2 = doc.ArtLayers[f"JOGO{i+1}TIME2"]

        text_of_Time1 = layerTime1.TextItem
        text_of_Time2 = layerTime2.TextItem
        text_of_Time1.contents = nt1
        text_of_Time2.contents = nt2
    colocaLogo()

def criaClass(navegador, inicio, fim):
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = rodada
    for i in range(inicio, fim):
        classificacao = []
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[1]/tbody/tr[{i+1}]/td[2]/strong').text)
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[1]').text)
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[2]').text)
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[3]').text)
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[4]').text)
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[5]').text)
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[6]').text)
        classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[8]').text)
        print(classificacao)

        layerTime = doc.ArtLayers[f"COLOCACAO{i+1}"]
        layerPontos = doc.ArtLayers[f"PONTOS{i+1}"]
        layerJogos = doc.ArtLayers[f"JOGOS{i+1}"]
        layerVitorias = doc.ArtLayers[f"VITORIAS{i+1}"]
        layerEmpates = doc.ArtLayers[f"EMPATES{i+1}"]
        layerDerrotas = doc.ArtLayers[f"DERROTAS{i+1}"]
        layerGP = doc.ArtLayers[f"GP{i+1}"]
        layerSG = doc.ArtLayers[f"SG{i+1}"]

        text_of_Time = layerTime.TextItem
        text_of_Pontos = layerPontos.TextItem    
        text_of_Jogos = layerJogos.TextItem  
        text_of_Vitorias = layerVitorias.TextItem  
        text_of_Empates = layerEmpates.TextItem  
        text_of_Derrotas = layerDerrotas.TextItem  
        text_of_GP = layerGP.TextItem  
        text_of_SG = layerSG.TextItem            
        text_of_Time.contents = classificacao[0]
        text_of_Pontos.contents = classificacao[1]   
        text_of_Jogos.contents = classificacao[2] 
        text_of_Vitorias.contents = classificacao[3]
        text_of_Empates.contents = classificacao[4]  
        text_of_Derrotas.contents = classificacao[5]  
        text_of_GP.contents = classificacao[6]
        text_of_SG.contents = classificacao[7]

def mudaCor():
    textColor1 = win32com.client.Dispatch("Photoshop.SolidColor")
    textColor2 = win32com.client.Dispatch("Photoshop.SolidColor")
    textColor3 = win32com.client.Dispatch("Photoshop.SolidColor")
    textColor1.rgb.red = 41
    textColor1.rgb.green = 41
    textColor1.rgb.blue = 41

    textColor2.rgb.red = 0
    textColor2.rgb.green = 100
    textColor2.rgb.blue = 0

    textColor3.rgb.red = 103
    textColor3.rgb.green = 252
    textColor3.rgb.blue = 103


    if(rodada == 'BRASILEIRÃO SÉRIE A'):
        layerPos1 = doc.ArtLayers[f"POS1"]
        layerPos1.textItem.color = textColor2
        layerPos2 = doc.ArtLayers[f"POS2"]
        layerPos2.textItem.color = textColor2
        layerPos3 = doc.ArtLayers[f"POS3"]
        layerPos3.textItem.color = textColor2
        layerPos4 = doc.ArtLayers[f"POS4"]
        layerPos4.textItem.color = textColor2
        layerPos5 = doc.ArtLayers[f"POS5"]
        layerPos5.textItem.color = textColor3
        layerPos6 = doc.ArtLayers[f"POS6"]
        layerPos6.textItem.color = textColor3
        layerPos7 = doc.ArtLayers[f"POS7"]
        layerPos7.textItem.color = textColor1
        layerPos8 = doc.ArtLayers[f"POS8"]
        layerPos8.textItem.color = textColor1
    elif(rodada == 'BRASILEIRÃO SÉRIE B'):
        layerPos1 = doc.ArtLayers[f"POS1"]
        layerPos1.textItem.color = textColor2
        layerPos2 = doc.ArtLayers[f"POS2"]
        layerPos2.textItem.color = textColor2
        layerPos3 = doc.ArtLayers[f"POS3"]
        layerPos3.textItem.color = textColor2
        layerPos4 = doc.ArtLayers[f"POS4"]
        layerPos4.textItem.color = textColor2
        layerPos5 = doc.ArtLayers[f"POS5"]
        layerPos5.textItem.color = textColor1
        layerPos6 = doc.ArtLayers[f"POS6"]
        layerPos6.textItem.color = textColor1
        layerPos7 = doc.ArtLayers[f"POS7"]
        layerPos7.textItem.color = textColor1
        layerPos8 = doc.ArtLayers[f"POS8"]
        layerPos8.textItem.color = textColor1
    elif(rodada == 'BRASILEIRÃO SÉRIE C'):
        layerPos1 = doc.ArtLayers[f"POS1"]
        layerPos1.textItem.color = textColor2
        layerPos2 = doc.ArtLayers[f"POS2"]
        layerPos2.textItem.color = textColor2
        layerPos3 = doc.ArtLayers[f"POS3"]
        layerPos3.textItem.color = textColor2
        layerPos4 = doc.ArtLayers[f"POS4"]
        layerPos4.textItem.color = textColor2
        layerPos5 = doc.ArtLayers[f"POS5"]
        layerPos5.textItem.color = textColor2
        layerPos6 = doc.ArtLayers[f"POS6"]
        layerPos6.textItem.color = textColor2
        layerPos7 = doc.ArtLayers[f"POS7"]
        layerPos7.textItem.color = textColor2
        layerPos8 = doc.ArtLayers[f"POS8"]
        layerPos8.textItem.color = textColor2
    
os.system("cls")

options = webdriver.ChromeOptions()
# options.add_argument("--headless")
navegador = webdriver.Chrome(options=options)
navegador.get("https://ge.globo.com/futebol/brasileirao-serie-a/")

psApp = win32com.client.Dispatch("Photoshop.Application")

rodada = navegador.find_element(By.XPATH,'//*[@id="header-produto"]/div[2]/div/div/h1/div/a').text
print(rodada)

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\resultados-brasileirao.psd")

doc = psApp.Application.ActiveDocument

criaConfrontos()

saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\resultados-brasileirao")

tiraLogo(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\resultados-brasileirao")

navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/nav/span[3]').click()
criaConfrontos()
saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\resultados-brasileirao-proxima-rodada.jpg")

tiraLogo(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\resultados-brasileirao-proxima-rodada-the.jpg")
closePSD(doc)

# navegador.get("https://ge.globo.com/futebol/brasileirao-serie-b/")

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\tabela-brasileirao-parte1.psd")
doc = psApp.Application.ActiveDocument

criaClass(navegador, 0, 10)
mudaCor()
colocaLogo()

print(rodada)
saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\tabela-brasileirao-parte1")
tiraLogo(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\tabela-brasileirao-parte1")
closePSD(doc)

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\tabela-brasileirao-parte2.psd")
doc = psApp.Application.ActiveDocument

criaClass(navegador, 10, 20)

colocaLogo()

saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\tabela-brasileirao-parte2")
tiraLogo(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\tabela-brasileirao-parte2")
closePSD(doc)


if(rodada != "BRASILEIRÃO SÉRIE C"):
    psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\artilharia-brasileirao.psd")
    doc = psApp.Application.ActiveDocument

    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem

    if(rodada == "BRASILEIRÃO SÉRIE A"):
        text_of_layer.contents = "Artilharia - Série A"
    elif(rodada == "BRASILEIRÃO SÉRIE B"):
        text_of_layer.contents = "Artilharia - Série B"

    
    for i in range(0,5):
        jogador = []
        jogador.append(navegador.find_element(By.XPATH, f'/html/body/div[2]/main/div[2]/div/section[2]/div/div/div[2]/div[{i+1}]/div[1]').text)
        time = navegador.find_element(By.XPATH, f'/html/body/div[2]/main/div[2]/div/section[2]/div/div/div[2]/div[{i+1}]/div[2]/div[2]/img')
        print(time.get_attribute('alt'))
        jogador.append(time.get_attribute('alt'))
        jogador.append(navegador.find_element(By.XPATH, f'/html/body/div[2]/main/div[2]/div/section[2]/div/div/div[2]/div[{i+1}]/div[2]/div[3]/div[1]').text)
        jogador.append(navegador.find_element(By.XPATH, f'/html/body/div[2]/main/div[2]/div/section[2]/div/div/div[2]/div[{i+1}]/div[2]/div[4]').text)
        layerPos = doc.ArtLayers[f"POS{i+1}"]
        layerTime = doc.ArtLayers[f"TIME{i+1}"]
        layerNome = doc.ArtLayers[f"NOME{i+1}"]
        layerGols = doc.ArtLayers[f"GOLS{i+1}"]
        
        layerRet = doc.ArtLayers[f"RET{i+1}"]
        if(jogador[0] == ''):
            layerRet.Visible = False
            layerPos.Visible = False
        else:
            layerRet.Visible = True
            layerPos.Visible = True

        text_of_Pos = layerPos.TextItem
        text_of_Time = layerTime.TextItem    
        text_of_Nome = layerNome.TextItem  
        text_of_Gols = layerGols.TextItem          
        text_of_Pos.contents = str(i + 1)
        text_of_Time.contents = jogador[1]   
        text_of_Nome.contents = jogador[2] 
        text_of_Gols.contents = jogador[3]

    colocaLogo()

    saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\artilharia-brasileirao")
    tiraLogo(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\teste\artilharia-brasileirao")
    closePSD(doc)

navegador.close()