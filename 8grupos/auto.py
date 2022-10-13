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

def saveJpg(jpg = f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/", the = False):
    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")

    options.Format = 6
    options.Quality = 100

    jpgfile = jpg + f"-{campeonato.lower()}.jpg"
    if(the == True):
        jpgfile = jpg + f"-{campeonato.lower()}-the.jpg" 

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

def nomeTimesConfrontos(nt1, nt2):
        if(nt1 == "República Tcheca"):
                nt1 = "R. Tcheca" 
        elif(nt1.upper() == "MANCHESTER UNITED"):
            nt1 = "MAN. UNITED"
        elif(nt1 == "Borussia Dortmund"):
                nt1 = "Dortmund"
        elif(nt1 == "Shakhtar Donetsk"):
                nt1 = "Shakhtar"
        elif(nt1 == "Bayern de Munique"):
                nt1 = "Bayern"
        elif(nt1 == "Atlético de Madrid"):
                nt1 = "Atlético Madrid"
        elif(nt1 == "Eintracht Frankfurt"):
                nt1 = "E. Frankfurt"
        elif(nt1 == "Olympique de Marselha"):
                nt1 = "O. Marselha"
        elif(nt1 == "Bayer Leverkusen "):
                nt1 = "Leverkusen "           
        elif(nt1 == "Paris Saint-Germain"):
                nt1 = "PSG"
        elif(nt1 == "Estrela Vermelha"):
            nt1 = "Estrela v."
            

        if(nt2 == "República Tcheca"):
                nt2 = "R. Tcheca"
        elif(nt2.upper() == "MANCHESTER UNITED"):
            nt2 = "MAN. UNITED"
        elif(nt2 == "Borussia Dortmund"):
                nt2 = "Dortmund"
        elif(nt2 == "Shakhtar Donetsk"):
                nt2 = "Shakhtar"
        elif(nt2 == "Bayern de Munique"):
                nt2 = "Bayern"
        elif(nt2 == "Atlético de Madrid"):
                nt2 = "Atlético Madrid"
        elif(nt2 == "Eintracht Frankfurt"):
                nt2 = "E. Frankfurt"
        elif(nt2 == "Olympique de Marselha"):
                nt2 = "O. Marselha"
        elif(nt2 == "Bayer Leverkusen "):
                nt2 = "Leverkusen "
        elif(nt2 == "Paris Saint-Germain"):
                nt2 = "PSG"
        elif(nt2 == "Estrela Vermelha"):
            nt2 = "Estrela v."
        

        return nt1, nt2

def colocaLogo():
    # layerLogo = doc.ArtLayers[f"LOGO"]
    layerFundo = doc.ArtLayers[f"LOGOFUNDO"]
    # layerLogo.Visible = True
    layerFundo.Visible = True

def criaConfrontos(qtdJogos, iniGrupos, fimGrupos):
    layerCampeonato = doc.ArtLayers["CAMPEONATO"]
    text_of_layerCampeonato = layerCampeonato.TextItem
    text_of_layerCampeonato.contents = f"{campeonato}"
    rodada = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[1]/section[2]/nav/span[2]').text
    print(rodada)
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f"{rodada}"

    try:
        os.remove(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}")
    except OSError:
        pass
    
    try:
        os.mkdir(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}")
    except OSError:
        pass

    # print(navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/ul/li[10]/div/div/div/div[1]/span[1]').text)
    for j in range(iniGrupos, fimGrupos):
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
                    if (exists(navegador, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[2]')):
                        nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[1]/span[1]')
                        nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[3]/span[1]')
                        p1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[1]').text
                        p2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[5]').text                                                                                                                                                                                                                                                        
                        local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/a/div[1]/div[1]/span[3]').text
                    else:
                        nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[1]/span[1]') 
                        nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[3]/span[1]')
                        p1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[2]/span[1]').text
                        p2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[2]/div[2]/span[5]').text                                                                                                                                                                                                                                                        
                        local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[2]/ul/li[{i+1}]/div/div/div/div[1]/span[3]').text

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
            
            nt1, nt2 = nomeTimesConfrontos(nt1, nt2)
                
            print(f"{nt1} {p1} x {p2} {nt2}")

            layerTime1 = doc.ArtLayers[f"GRUPO{j+1}JOGO{i+1}TIME1"]
            layerTime2 = doc.ArtLayers[f"GRUPO{j+1}JOGO{i+1}TIME2"]

            text_of_Time1 = layerTime1.TextItem
            text_of_Time2 = layerTime2.TextItem            
            
            text_of_Time1.contents = nt1
            text_of_Time2.contents = nt2
    colocaLogo()

def criaClass(navegador, iniGrupos, fimGrupos, qtdTimes):
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = campeonato

    for j in range(iniGrupos, fimGrupos):
        for i in range(qtdTimes):
            classificacao = []
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[1]/tbody/tr[{i+1}]/td[2]/strong').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[1]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[2]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[3]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[6]').text)
            classificacao.append(navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article[{j+1}]/section[1]/div/table[2]/tbody/tr[{i+1}]/td[8]').text)
            print(classificacao)

            if(classificacao[0] == "Borussia Dortmund"):
                classificacao[0] = "Dortmund"
            
            elif(classificacao[0] == "Shakhtar Donetsk"):
                classificacao[0] = "Shakhtar"
                
            elif(classificacao[0] == "Bayern de Munique"):
                classificacao[0] = "Bayern"

            elif(classificacao[0] == "Atlético de Madrid"):
                classificacao[0] = "Atlético Madrid"

            elif(classificacao[0] == "Eintracht Frankfurt"):
                classificacao[0] = "E. Frankfurt"
            
            elif(classificacao[0] == "Olympique de Marselha"):
                classificacao[0] = "Olympique"

            elif(classificacao[0] == "Bayer Leverkusen"):
                classificacao[0] = "Leverkusen"
            
            elif(classificacao[0] == "Paris Saint-Germain"):
                classificacao[0] = "PSG"

            layerTime = doc.ArtLayers[f"GRUPO{j+1}COLOCACAO{i+1}"]
            layerPontos = doc.ArtLayers[f"GRUPO{j+1}PONTOS{i+1}"]
            layerJogos = doc.ArtLayers[f"GRUPO{j+1}JOGOS{i+1}"]
            layerVitorias = doc.ArtLayers[f"GRUPO{j+1}VITORIAS{i+1}"]
            layerGP = doc.ArtLayers[f"GRUPO{j+1}GP{i+1}"]
            layerSG = doc.ArtLayers[f"GRUPO{j+1}SG{i+1}"]

            text_of_Time = layerTime.TextItem
            text_of_Pontos = layerPontos.TextItem    
            text_of_Jogos = layerJogos.TextItem  
            text_of_Vitorias = layerVitorias.TextItem  
            text_of_GP = layerGP.TextItem  
            text_of_SG = layerSG.TextItem            
            text_of_Time.contents = classificacao[0]
            text_of_Pontos.contents = classificacao[1]   
            text_of_Jogos.contents = classificacao[2] 
            text_of_Vitorias.contents = classificacao[3]
            text_of_GP.contents = classificacao[4]
            text_of_SG.contents = classificacao[5]

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

    textColor3.rgb.red = 255
    textColor3.rgb.green = 0
    textColor3.rgb.blue = 0

    if(campeonato == 'LIGA DAS NAÇÕES'):
        for i in range(4):
            print(i)
            layerPos2 = doc.ArtLayers[f"GRUPO{i+1}POS2"]
            layerPos2.textItem.color = textColor1
            layerPos4 = doc.ArtLayers[f"GRUPO{i+1}POS4"]
            layerPos4.textItem.color = textColor3
           
    else:
        for i in range(4):
            layerPos2 = doc.ArtLayers[f"GRUPO{i+1}POS2"]
            layerPos2.textItem.color = textColor2
            layerPos4 = doc.ArtLayers[f"GRUPO{i+1}POS4"]
            layerPos4.textItem.color = textColor1
    
os.system("cls")

options = webdriver.ChromeOptions()
# options.binary_location = "C:\\Program Files\\Google\\Chrome Beta\\Application\\chrome"
options.add_argument("--headless")
navegador = webdriver.Chrome(options=options)
# navegador.get("https://ge.globo.com/futebol/copa-do-mundo/2022/")
# navegador.get("https://ge.globo.com/futebol/futebol-internacional/liga-dos-campeoes/")
# navegador.get("https://ge.globo.com/futebol/futebol-internacional/liga-das-nacoes/")
navegador.get("https://ge.globo.com/futebol/futebol-internacional/liga-europa/")

# time.sleep(15)

######################### PROXIMA RODADA #########################
# proximaRodadaGrupo1 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[1]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo1)

# proximaRodadaGrupo2 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[2]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo2)

# proximaRodadaGrupo3 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[3]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo3)

# proximaRodadaGrupo4 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[4]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo4)

# proximaRodadaGrupo5 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[5]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo5)

# proximaRodadaGrupo6 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[6]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo6)

# proximaRodadaGrupo7 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[7]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo7)

# proximaRodadaGrupo8 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[8]/section[2]/nav/span[3]')
# navegador.execute_script('arguments[0].click();', proximaRodadaGrupo8)

######################### RODADA ANTERIOR #########################
# rodadaAnteriorGrupo1 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[1]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo1)

# rodadaAnteriorGrupo2 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[2]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo2)

# rodadaAnteriorGrupo3 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[3]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo3)

# rodadaAnteriorGrupo4 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[4]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo4)

# rodadaAnteriorGrupo5 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[5]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo5)

# rodadaAnteriorGrupo6 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[6]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo6)

# rodadaAnteriorGrupo7 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[7]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo7)

# rodadaAnteriorGrupo8 = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/article[8]/section[2]/nav/span[1]')
# navegador.execute_script('arguments[0].click();', rodadaAnteriorGrupo8)

psApp = win32com.client.Dispatch("Photoshop.Application")

campeonato = navegador.find_element(By.XPATH,'//*[@id="header-produto"]/div[2]/div/div/h1/div/a').text
print(campeonato)

if(campeonato == "COPA DO MUNDO DA FIFA™"):
    campeonato = "COPA DO MUNDO"

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\8grupos\resultados-grupos-parte1.psd")

doc = psApp.Application.ActiveDocument

if (campeonato == "LIGA DAS NAÇÕES"):
    criaConfrontos(2,0,4)

    saveJpg(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/resultados")

    tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/resultados")
else:
    criaConfrontos(2,0,4)

    saveJpg(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/resultados-parte1")

    tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/resultados-parte1")

    closePSD(doc)

    psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\8grupos\resultados-grupos-parte2.psd")

    doc = psApp.Application.ActiveDocument

    criaConfrontos(2,4,8)

    saveJpg(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/resultados-parte2")

    tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/resultados-parte2")

closePSD(doc)

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\8grupos\tabela-grupos-parte1.psd")
doc = psApp.Application.ActiveDocument

if (campeonato == "LIGA DAS NAÇÕES"):
    criaClass(navegador, 0, 4, 4)
    colocaLogo()

    mudaCor()

    saveJpg(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/tabela-grupos")
    tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/tabela-grupos")

else: 

    criaClass(navegador, 0, 4, 4)
    colocaLogo()

    mudaCor()

    saveJpg(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/tabela-grupos-parte1")
    tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/tabela-grupos-parte1")

    closePSD(doc)

    psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\8grupos\tabela-grupos-parte2.psd")
    doc = psApp.Application.ActiveDocument

    criaClass(navegador, 4, 8, 4)
    colocaLogo()

    print(campeonato)
    saveJpg(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/tabela-grupos-parte2")
    tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/8grupos/{campeonato}/tabela-grupos-parte2")

closePSD(doc)

navegador.close()