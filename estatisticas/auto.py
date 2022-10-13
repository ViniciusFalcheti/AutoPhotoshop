from numpy import False_
import win32com.client
import os, time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from datetime import date, datetime
import requests

def exists(navegador, elem):
    try:
        teste = navegador.find_element_by_xpath(elem)
        return True
    except NoSuchElementException:
        return False

def closePSD(doc):
    doc.Save()
    doc.Application.ActiveDocument.Close(2)

def saveJpg(jpg = r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\estatisticas", the = False):
    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")

    options.Format = 6
    options.Quality = 100

    if(rodada == 'BRASILEIRÃO SÉRIE A'):
         bra = "serieA"
    elif(rodada == 'BRASILEIRÃO SÉRIE B'):
        bra = "serieB"
    elif(rodada == 'BRASILEIRÃO SÉRIE C'):
        bra = "serieC"
    elif(rodada == 'FUTEBOL ITALIANO'):
        bra = "italiano"
    elif(rodada == 'PREMIER LEAGUE'):
        bra = "ingles"
    elif(rodada == 'FUTEBOL ALEMÃO'):
        bra = "alemão"

    jpgfile = jpg + f"-{bra}.jpg"
    if(the == True):
        jpgfile = jpg + f"-{bra}-the.jpg" 

    doc.Export(ExportIn=jpgfile, ExportAs = 2, Options = options)

def criaConfrontos(jogos):
    camp = navegador.find_element(By.XPATH, '//*[@id="header-produto"]/div[2]/div/div/h1/div/a').text
    rodada = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/nav/span[2]').text
    print(rodada)
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f"{camp} - {rodada}"

    # print(navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/ul/li[10]/div/div/div/div[1]/span[1]').text)

    for i in range(jogos):
        if(exists(navegador, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]')):
            str_date = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text
        else:
            str_date = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text
        
        if(str_date == None or str_date == ""):
            data = None
            layerGols2.Visible = False
        else:
            str_date = str_date[-10:]
            data = datetime.strptime(str_date, '%d/%m/%Y').date()

        print(data)
        # if data != None and navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[2]/span[1]').text == '':
        if data != None:
            if (data < date.today()):
                if (exists(navegador, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[2]')):
                        img = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[1]/img')
                        img2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[3]/img')
                else:
                        img = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[1]/img')
                        img2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[3]/img')
                       
            else:
                img = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[1]/img')
                img2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[3]/img')
            
        else:
            img = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[1]/img')
            img2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[3]/img')
        
        psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\estatisticas\arte-jogo.psd")
        doc2 = psApp.Application.ActiveDocument
        img = img.get_attribute('src')
        img2 = img2.get_attribute('src')
       
        
        img_data = requests.get(img).content
        with open('img/time1.svg', 'wb') as handler:
            handler.write(img_data)

        img_data2 = requests.get(img2).content
        with open('img/time2.svg', 'wb') as handler:
            handler.write(img_data2)
            
        caminho = 'F:/OneDrive - OPIC Telecom/Área de Trabalho/auto\estatisticas/img/time1.svg'
        psApp.Load(caminho)
        psApp.ActiveDocument.Selection.SelectAll()
        psApp.ActiveDocument.Selection.Copy()
        psApp.ActiveDocument.Close()
        psApp.ActiveDocument.Paste()
        layerImg1 = doc.activeLayer

        layerImg1.Translate(200,200)
        psApp.ActiveDocument.Close()



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

        classificacao[0] = nomeTimesClass(classificacao[0])

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

os.system("cls")

options = webdriver.ChromeOptions()
# options.binary_location = "C:\\Program Files\\Google\\Chrome Beta\\Application\\chrome"
options.add_argument("--headless")
navegador = webdriver.Chrome(options=options)
navegador.get("https://ge.globo.com/futebol/brasileirao-serie-a/")
# navegador.get("https://ge.globo.com/futebol/futebol-internacional/futebol-ingles/")
# navegador.get("https://ge.globo.com/futebol/futebol-internacional/futebol-italiano/")
# navegador.get("https://ge.globo.com/futebol/futebol-internacional/futebol-alemao/")

# time.sleep(15)

# proximaRodada = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodada)

# if(exists(navegador, '//*[@id="classificacao__wrapper"]/section/nav/span[3]')):
#     proximaRodada = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/nav/span[3]')
# else:
#     proximaRodada = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodada)          

# if(exists(navegador, '//*[@id="classificacao__wrapper"]/nav/span[1]')):
#     rodadaAnterior = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/nav/span[1]')
# else:
#     rodadaAnterior = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnterior)

psApp = win32com.client.Dispatch("Photoshop.Application")

rodada = navegador.find_element(By.XPATH,'//*[@id="header-produto"]/div[2]/div/div/h1/div/a').text
print(rodada)

if (rodada != "FUTEBOL ALEMÃO"):
    criaConfrontos(10)

else:
    criaConfrontos(9)

saveJpg(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\estatisticas\resultados")

# navegador.get("https://ge.globo.com/futebol/brasileirao-serie-b/")

navegador.close()