from numpy import False_
from sqlalchemy import null
import win32com.client
import os, time
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

def saveJpg(jpg = r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\2grupos", the = False):
    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")

    options.Format = 6
    options.Quality = 100

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
    # layerLogo = doc.ArtLayers[f"LOGO"]
    layerFundo = doc.ArtLayers[f"LOGOFUNDO"]
    # layerLogo.Visible = True
    layerFundo.Visible = True

def criaConfrontos(qtdJogos, qtdGrupos):  
    rodada = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[1]/section[2]/nav/span[2]').text
    print(rodada)
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f"{rodada}"
    # print(navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/ul/li[10]/div/div/div/div[1]/span[1]').text)
    for j in range(qtdGrupos):
        for i in range(qtdJogos):
            if(exists(navegador, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]')):                                                    
                str_date = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text
            else:
                str_date = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text
            
            if(str_date == None or str_date == ""):
                data = None
                layerGols2.Visible = False
            else:
                str_date = str_date[-10:]
                data = datetime.strptime(str_date, '%d/%m/%Y').date()

            print(data)

            if data != None:
                if (data < date.today()):
                    nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[1]/span[1]')
                    nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[3]/span[1]')
                    p1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[1]').text
                    p2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[5]').text                                                                                                                                                                                                                                                        
                    local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[3]').text
                    layerGols1 = doc.ArtLayers[f"GRUPO{j+1}RES{i+1}TIME1"]
                    layerGols2 = doc.ArtLayers[f"GRUPO{j+1}RES{i+1}TIME2"]
                    layerGols1.Visible = True
                    layerGols2.Visible = True
                    text_of_Gols1 = layerGols1.TextItem
                    text_of_Gols2 = layerGols2.TextItem
                    text_of_Gols1.contents = p1
                    text_of_Gols2.contents = p2
                else:
                    p1 = 0
                    p2 = 0
                    nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[1]/span[1]') 
                    nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[3]/span[1]')
                    local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[3]').text
                    layerGols1 = doc.ArtLayers[f"GRUPO{j+1}RES{i+1}TIME1"]
                    layerGols2 = doc.ArtLayers[f"GRUPO{j+1}RES{i+1}TIME2"]
                    layerGols1.Visible = False
                    layerGols2.Visible = False
                
            else:
                p1 = 0
                p2 = 0
                nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[1]/span[1]')   
                nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[3]/span[1]')
                local = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text
                layerGols1 = doc.ArtLayers[f"GRUPO{j+1}RES{i+1}TIME1"]
                layerGols2 = doc.ArtLayers[f"GRUPO{j+1}RES{i+1}TIME2"]
                layerGols1.Visible = False
                layerGols2.Visible = False
            
            layerLocal = doc.ArtLayers[f"GRUPO{j+1}LOCAL{i+1}"]
            text_of_Local = layerLocal.TextItem
            text_of_Local.contents = local
            nt1 = nt1.get_attribute("title")
            nt2 = nt2.get_attribute("title")
            print(f"{nt1} {p1} x {p2} {nt2}")

            layerTime1 = doc.ArtLayers[f"GRUPO{j+1}JOGO{i+1}TIME1"]
            layerTime2 = doc.ArtLayers[f"GRUPO{j+1}JOGO{i+1}TIME2"]

            text_of_Time1 = layerTime1.TextItem
            text_of_Time2 = layerTime2.TextItem
            text_of_Time1.contents = nt1
            text_of_Time2.contents = nt2
    colocaLogo()

def criaClass(navegador, qtdGrupos, qtdTimes):
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = rodada

    for j in range(qtdGrupos):
        for i in range(qtdTimes):
            classificacao = []
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[1]/tbody/tr[{i+1}]/td[2]/strong').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[1]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[2]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[3]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[4]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[5]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[6]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[8]').text)
            print(classificacao)

            layerTime = doc.ArtLayers[f"GRUPO{j+1}COLOCACAO{i+1}"]
            layerPontos = doc.ArtLayers[f"GRUPO{j+1}PONTOS{i+1}"]
            layerJogos = doc.ArtLayers[f"GRUPO{j+1}JOGOS{i+1}"]
            layerVitorias = doc.ArtLayers[f"GRUPO{j+1}VITORIAS{i+1}"]
            layerEmpates = doc.ArtLayers[f"GRUPO{j+1}EMPATES{i+1}"]
            layerDerrotas = doc.ArtLayers[f"GRUPO{j+1}DERROTAS{i+1}"]
            layerGP = doc.ArtLayers[f"GRUPO{j+1}GP{i+1}"]
            layerSG = doc.ArtLayers[f"GRUPO{j+1}SG{i+1}"]

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
# options.binary_location = "C:\\Program Files\\Google\\Chrome Beta\\Application\\chrome"
options.add_argument("--headless")
navegador = webdriver.Chrome(options=options)
navegador.get("https://ge.globo.com/futebol/brasileirao-serie-c/")

# time.sleep(15)

faseanterior = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/nav/span[1]')
navegador.execute_script('arguments[0].click();', faseanterior)

# proximaRodadaGrupo1 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[1]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo1)

# rodadaAnteriorGrupo1 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[1]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo1)

# proximaRodadaGrupo2 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[2]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo2)

# rodadaAnteriorGrupo2 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[2]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo2)

psApp = win32com.client.Dispatch("Photoshop.Application")

rodada = navegador.find_element(By.XPATH,'//*[@id="header-produto"]/div[2]/div/div/h1/div/a').text
print(rodada)

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\2grupos\resultados-brasileirao.psd")

doc = psApp.Application.ActiveDocument

criaConfrontos(2,2)

saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\2grupos\resultados-brasileirao")

tiraLogo(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\2grupos\resultados-brasileirao")

closePSD(doc)

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\2grupos\tabela-grupos.psd")
doc = psApp.Application.ActiveDocument

criaClass(navegador, 2, 4)
colocaLogo()

print(rodada)
saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\2grupos\tabela-grupos")
tiraLogo(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\2grupos\tabela-grupos")
closePSD(doc)

navegador.close()