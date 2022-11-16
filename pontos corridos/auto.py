from asyncio.windows_events import NULL
from numpy import False_
import win32com.client
import os, time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from datetime import date, datetime
import shutil
import requests

def exists(navegador, elem):
    try:
        teste = navegador.find_element_by_xpath(elem)
        return True
    except NoSuchElementException:
        return False

def saveJpg(doc, jpg = r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos", the = False):
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

def closePSD(doc):
    doc.Save()
    doc.Application.ActiveDocument.Close(2)

def tiraLogo(jpg):
    layerLogo = doc.ArtLayers[f"LOGO"]
    layerFundo = doc.ArtLayers[f"LOGOFUNDO"]
    layerLogo.Visible = False
    layerFundo.Visible = False

    saveJpg(doc, jpg, True)

def colocaLogo():
    # layerLogo = doc.ArtLayers[f"LOGO"]
    layerFundo = doc.ArtLayers[f"LOGOFUNDO"]
    # layerLogo.Visible = True
    layerFundo.Visible = True

def nomeTimesConfrontos(nt1, nt2):
        if(nt1.upper() == "BRASIL DE PELOTAS"):
            nt1 = "BRASIL-RS"
        elif(nt1.upper() == "MANCHESTER UNITED"):
            nt1 = "MAN. UNITED"
        elif(nt1.upper() == "MANCHESTER CITY"):
            nt1 = "MAN. CITY"
        elif(nt1.upper() == "NOTTINGHAM FOREST"):
            nt1 = "NOTTINGHAM F."
        elif(nt1 == "Bayern de Munique"):
            nt1 = "Bayern"
        elif(nt1 == "Borussia Dortmund"):
            nt1 = "Dortmund"                    
        elif(nt1 == "Eintracht Frankfurt"):
            nt1 = "E. Frankfurt"
        elif(nt1 == "Bayer Leverkusen "):
            nt1 = "Leverkusen"
        elif(nt1 == "Borussia Mönchengladbach"):
            nt1 = "M'gladbach"

        if(nt2.upper() == "BRASIL DE PELOTAS"):
            nt2 = "BRASIL-RS"
        elif(nt2.upper() == "MANCHESTER UNITED"):
            nt2 = "MAN. UNITED"
        elif(nt2.upper() == "MANCHESTER CITY"):
            nt2 = "MAN. CITY"
        elif(nt2.upper() == "NOTTINGHAM FOREST"):
            nt2 = "NOTTINGHAM F."        
        elif(nt2 == "Bayern de Munique"):
            nt2 = "Bayern"
        elif(nt2 == "Borussia Dortmund"):
            nt2 = "Dortmund"
        elif(nt2 == "Eintracht Frankfurt"):
            nt2 = "E. Frankfurt"
        elif(nt2 == "Bayer Leverkusen "):
            nt2 = "Leverkusen"
        elif(nt2 == "Borussia Mönchengladbach"):
            nt2 = "M'gladbach"

        return nt1, nt2

def arteConfronto(rodada, campeonato, nt1, nt2, resTimeCasa, resTimeFora, horarioJogo, localJogo, dataJogo, ordem):

    if(resTimeCasa != NULL):
        psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\arte-jogo-resultado.psd")
        doc2 = psApp.Application.ActiveDocument
        timeCasaRes = doc2.layerSets["RESULTADO"].ArtLayers[f"RESTIMECASA"]
        timeForaRes = doc2.layerSets["RESULTADO"].ArtLayers[f"RESTIMEFORA"]

        resTimeCasaLayer = timeCasaRes.TextItem
        resTimeForaLayer = timeForaRes.TextItem

        resTimeCasaLayer.contents = f"{resTimeCasa}"
        resTimeForaLayer.contents = f"{resTimeFora}"
    else:
        psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\arte-jogo.psd")
        doc2 = psApp.Application.ActiveDocument
    
    

    logoCampeonato = doc2.layerSets["CAMPEONATO"].ArtLayers[f"{campeonato.upper()}"]
    logoCampeonato.Visible = True

    print(campeonato)
    emblemaTime1 = doc2.layerSets["EMBLEMAS"].LayerSets[f"{campeonato.upper()}"].layerSets["CASA"].ArtLayers[f"{nt1.upper()}"]
    emblemaTime2 = doc2.layerSets["EMBLEMAS"].layerSets[f"{campeonato.upper()}"].layerSets["FORA"].ArtLayers[f"{nt2.upper()}"]

    emblemaTime1.Visible = True
    emblemaTime2.Visible = True

    local = doc2.layerSets["LOCAL"].ArtLayers[f"LOCAL"]
    horario = doc2.layerSets["LOCAL"].ArtLayers[f"HORARIO"]

    localLayer = local.TextItem
    horarioLayer = horario.TextItem

    localLayer.contents = f"{localJogo}"
    horarioLayer.contents = f"{dataJogo} ÀS {horarioJogo}"
    # retColor1 = win32com.client.Dispatch("Photoshop.SolidColor") #VERDE - COR DA VITORIA
    # retColor2 = win32com.client.Dispatch("Photoshop.SolidColor") #VERMELHO - COR DA DERROTA
    # retColor3 = win32com.client.Dispatch("Photoshop.SolidColor") #CINZA - COR DO EMPATE

    # retColor1.rgb.red = 48
    # retColor1.rgb.green = 255
    # retColor1.rgb.blue = 0

    # retColor2.rgb.red = 255
    # retColor2.rgb.green = 0
    # retColor2.rgb.blue = 0

    # retColor3.rgb.red = 200
    # retColor3.rgb.green = 200
    # retColor3.rgb.blue = 200
                                               
    for i in range(20):
        ntTabela = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[1]/tbody/tr[{i+1}]/td[2]/strong').text
        if(ntTabela == nt1):
            colocacao = doc2.ArtLayers[f"COLOCACAO1"]
            colocacaoLayer = colocacao.TextItem
            colocacaoLayer.contents = f"{i+1}º"
            for j in range(5):
                resJogo = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[10]/span[{j+1}]')
                resJogo = resJogo.get_attribute("class")
                res = doc2.layerSets["RESANTERIORES"].layerSets["CASA"].layerSets[F"JOGO{j+1}"].ArtLayers[f"RES"]
                resLayer = res.TextItem
                if(resJogo == "classificacao__ultimos_jogos classificacao__ultimos_jogos--d "):   
                    resLayer.contents = "D"
                elif(resJogo == "classificacao__ultimos_jogos classificacao__ultimos_jogos--e "):
                    resLayer.contents = "E"
                elif(resJogo == "classificacao__ultimos_jogos classificacao__ultimos_jogos--v "):
                    resLayer.contents = "V"
                ret = doc2.layerSets["RESANTERIORES"].layerSets["CASA"].layerSets[F"JOGO{j+1}"].layerSets[F"RET"].ArtLayers[f"{resLayer.contents}"]
                ret.Visible = True
        elif(ntTabela == nt2):
            colocacao = doc2.ArtLayers[f"COLOCACAO2"]
            colocacaoLayer = colocacao.TextItem
            colocacaoLayer.contents = f"{i+1}º"
            for j in range(5):
                resJogo = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/article/section[1]/div/table[2]/tbody/tr[{i+1}]/td[10]/span[{j+1}]')
                resJogo = resJogo.get_attribute("class")
                res = doc2.layerSets["RESANTERIORES"].layerSets["FORA"].layerSets[F"JOGO{j+1}"].ArtLayers[f"RES"]
                resLayer = res.TextItem
                if(resJogo == "classificacao__ultimos_jogos classificacao__ultimos_jogos--d "):   
                    resLayer.contents = "D"
                elif(resJogo == "classificacao__ultimos_jogos classificacao__ultimos_jogos--e "):
                    resLayer.contents = "E"
                elif(resJogo == "classificacao__ultimos_jogos classificacao__ultimos_jogos--v "):
                    resLayer.contents = "V"
                ret = doc2.layerSets["RESANTERIORES"].layerSets["FORA"].layerSets[F"JOGO{j+1}"].layerSets[F"RET"].ArtLayers[f"{resLayer.contents}"]
                ret.Visible = True
    saveJpg(doc2, f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{campeonato}/confrontos/{rodada.lower()}/{ordem}- {nt1}x{nt2}")

    doc2.Close(2)

def miniaturaPalpite(rodada, campeonato):
    psApp.Open(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/miniatura-palpites-{campeonato}.psd")
    doc2 = psApp.Application.ActiveDocument

    layerRodada = doc2.ArtLayers["RODADA"]
    text_of_rodada = layerRodada.TextItem
    text_of_rodada.contents = f"{rodada}"

    saveJpg(doc2, f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{campeonato}/confrontos/{rodada.lower()}/miniatura-palpites-{rodada.lower()}")
    closePSD(doc2)

def criaConfrontos(jogos):
    camp = navegador.find_element(By.XPATH, '//*[@id="header-produto"]/div[2]/div/div/h1/div/a').text
    rodada = navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/nav/span[2]').text
    print(rodada)
    layerText = doc.ArtLayers["RODADA"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = f"{camp} - {rodada}"

    # print(navegador.find_element(By.XPATH, '//*[@id="classificacao__wrapper"]/section/ul/li[10]/div/div/div/div[1]/span[1]').text)
    
    if(camp == "BRASILEIRÃO SÉRIE A"):
        camp = "seriea"
        campeonatoDetalhado = "Campeonato Brasileiro Série A 2022"
    elif(camp == "BRASILEIRÃO SÉRIE B"):
        camp = "serieb"
        campeonatoDetalhado = "Campeonato Brasileiro Série B 2022"

    # try:
    #     os.remove(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}")
    # except OSError:
    #     pass
    
    try:
        os.mkdir(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}")
    except OSError:
        pass

    # try:
    #     os.remove(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/confrontos")
    # except OSError:
    #     pass

    try:
        os.mkdir(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/confrontos")
    except OSError:
        pass

    try:
        os.mkdir(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/confrontos/{rodada.lower()}")
    except OSError:
        pass

    arquivo = open(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/confrontos/{rodada.lower()}/descricao - {camp}.txt", 'w')
    texto = [f"Palpites {rodada.lower()} {campeonatoDetalhado}\n\n"]

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
                        nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[1]/span[1]')
                        nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[3]/span[1]')
                        p1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[1]').text
                        p2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[2]/div[2]/span[5]').text
                        dataJogo = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text
                        localJogo = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[2]').text
                        horarioJogo = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[3]').text
                        # local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/a/div[1]/div[1]/span[3]').text
                else:
                        nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[1]/span[1]')                                         
                        nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[3]/span[1]')
                        # local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[3]').text
                        dataJogo = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text
                        localJogo = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text
                        horarioJogo = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[3]').text
                        p1 = NULL
                        p2 = NULL
                        # p1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[2]/span[1]').text
                        # p2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[2]/span[5]').text                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           
                        

                layerGols1 = doc.ArtLayers[f"RES{i+1}TIME1"]
                layerGols2 = doc.ArtLayers[f"RES{i+1}TIME2"]
                layerGols1.Visible = True
                layerGols2.Visible = True
                text_of_Gols1 = layerGols1.TextItem
                text_of_Gols2 = layerGols2.TextItem
                text_of_Gols1.contents = p1
                text_of_Gols2.contents = p2
            else:
                p1 = NULL
                p2 = NULL
                nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[1]/span[1]')                                         
                nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[3]/span[1]')
                # local = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text + " " + navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[3]').text
                dataJogo = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[1]').text
                localJogo = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text
                horarioJogo = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[3]').text
                layerGols1 = doc.ArtLayers[f"RES{i+1}TIME1"]
                layerGols2 = doc.ArtLayers[f"RES{i+1}TIME2"]
                layerGols1.Visible = False
                layerGols2.Visible = False
            
        else:
            p1 = NULL
            p2 = NULL
            nt1 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[1]/span[1]')                                         
            nt2 = navegador.find_element(By.XPATH, f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[2]/div[3]/span[1]')
            local = navegador.find_element_by_xpath(f'//*[@id="classificacao__wrapper"]/section/ul/li[{i+1}]/div/div/div/div[1]/span[2]').text
            layerGols1 = doc.ArtLayers[f"RES{i+1}TIME1"]
            layerGols2 = doc.ArtLayers[f"RES{i+1}TIME2"]
            layerGols1.Visible = False
            layerGols2.Visible = False
                
        layerLocal = doc.ArtLayers[f"LOCAL{i+1}"]
        text_of_Local = layerLocal.TextItem
        if(dataJogo == ""):
            text_of_Local.contents = local
        else:
            text_of_Local.contents = f"{localJogo} {dataJogo} {horarioJogo}"
        nt1 = nt1.get_attribute("title")
        nt2 = nt2.get_attribute("title")
        # print(f"{nt1} {p1} x {p2} {nt2}")

        arteConfronto(rodada, camp, nt1, nt2, p1, p2, horarioJogo, localJogo, dataJogo, i+1)

        nt1, nt2 = nomeTimesConfrontos(nt1,nt2)

        texto.append(f"{nt1} x {nt2}\n")

        layerTime1 = doc.ArtLayers[f"JOGO{i+1}TIME1"]
        layerTime2 = doc.ArtLayers[f"JOGO{i+1}TIME2"]

        text_of_Time1 = layerTime1.TextItem
        text_of_Time2 = layerTime2.TextItem
        text_of_Time1.contents = nt1
        text_of_Time2.contents = nt2
    
    print(texto)
    arquivo.writelines(texto)
    arquivo.close()
    miniaturaPalpite(rodada, camp)
    colocaLogo()

def nomeTimesClass(nt):
    if(nt.upper() == "BRASIL DE PELOTAS"):
        nt = "BRASIL-RS"

    elif(nt.upper() == "MANCHESTER UNITED"):
        nt = "Man. United"
        
    elif(nt.upper() == "MANCHESTER CITY"):
        nt = "Man. City"

    elif(nt.upper() == "NOTTINGHAM FOREST"):
        nt = "Nottingham F."

    elif(nt == "Bayern de Munique"):
        nt = "Bayern"

    elif(nt == "Borussia Dortmund"):
        nt = "Dortmund"

    elif(nt == "Eintracht Frankfurt"):
        nt = "E. Frankfurt"
        
    elif(nt == "Bayer Leverkusen"):
        nt = "Leverkusen"

    elif(nt == "Borussia Mönchengladbach"):
        nt = "M'gladbach"

    return nt

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

def mudaCor1():
    textColor1 = win32com.client.Dispatch("Photoshop.SolidColor") #CINZA - COR PADRÃO
    textColor2 = win32com.client.Dispatch("Photoshop.SolidColor") #VERDÃO - PRIMEIROS COLOCADOS  
    textColor3 = win32com.client.Dispatch("Photoshop.SolidColor") #VERDE MAIS CLARO - CLASSIFICADOS PARA A PRE LIBERTA OU LIGA EUROPA
    textColor4 = win32com.client.Dispatch("Photoshop.SolidColor") #VERDE DIFERENTE - CLASSIFICADO PARA CONFERENCE LEAGUE

    textColor1.rgb.red = 41
    textColor1.rgb.green = 41
    textColor1.rgb.blue = 41

    textColor2.rgb.red = 0
    textColor2.rgb.green = 100
    textColor2.rgb.blue = 0

    textColor3.rgb.red = 0
    textColor3.rgb.green = 250
    textColor3.rgb.blue = 0

    textColor4.rgb.red = 128
    textColor4.rgb.green = 255
    textColor4.rgb.blue = 128

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
        layerPos5.textItem.color = textColor2
        layerPos6 = doc.ArtLayers[f"POS6"]
        layerPos6.textItem.color = textColor2
        layerPos7 = doc.ArtLayers[f"POS7"]
        layerPos7.textItem.color = textColor3
        layerPos8 = doc.ArtLayers[f"POS8"]
        layerPos8.textItem.color = textColor3
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
    elif(rodada == 'FUTEBOL ALEMÃO' or rodada == 'PREMIER LEAGUE'):
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
        layerPos7.textItem.color = textColor4
        layerPos8 = doc.ArtLayers[f"POS8"]
        layerPos8.textItem.color = textColor1

def mudaCor2():
    textColor1 = win32com.client.Dispatch("Photoshop.SolidColor") #CINZA - COR PADRÃO
    textColor2 = win32com.client.Dispatch("Photoshop.SolidColor") #VERMELHO - TIMES REBAIXADOS 
    textColor3 = win32com.client.Dispatch("Photoshop.SolidColor") #LARANJA - TIME QUE NÃO CAI DIRETO 

    textColor1.rgb.red = 41
    textColor1.rgb.green = 41
    textColor1.rgb.blue = 41

    textColor2.rgb.red = 255
    textColor2.rgb.green = 0
    textColor2.rgb.blue = 0

    textColor3.rgb.red = 255
    textColor3.rgb.green = 114
    textColor3.rgb.blue = 0
    
    print(rodada)
    if(rodada == 'FUTEBOL ALEMÃO'):
        layerPos16 = doc.ArtLayers[f"POS16"]
        layerPos16.textItem.color = textColor3
    elif(rodada == 'PREMIER LEAGUE' or rodada == 'FUTEBOL ITALIANO' or rodada == 'FUTEBOL FRANCÊS'): 
        layerPos16 = doc.ArtLayers[f"POS16"]
        layerPos16.textItem.color = textColor1
        layerPos17 = doc.ArtLayers[f"POS17"]
        layerPos17.textItem.color = textColor1
    else:
        layerPos16 = doc.ArtLayers[f"POS16"]
        layerPos16.textItem.color = textColor1
        layerPos17 = doc.ArtLayers[f"POS17"]
        layerPos17.textItem.color = textColor2     

def artilharia():
    if(rodada == "BRASILEIRÃO SÉRIE B" or rodada == "BRASILEIRÃO SÉRIE A"):
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

        saveJpg(doc, f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/artilharia-brasileirao")
        tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/artilharia-brasileirao")
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

if(rodada == "BRASILEIRÃO SÉRIE A"):
        camp = "seriea"
elif(rodada == "BRASILEIRÃO SÉRIE B"):
        camp = "serieb"

if (rodada != "FUTEBOL ALEMÃO"):
    psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\resultados-brasileirao.psd")
    doc = psApp.Application.ActiveDocument

    criaConfrontos(10)

else:
    psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\resultados-alemao.psd")
    doc = psApp.Application.ActiveDocument

    criaConfrontos(9)

saveJpg(doc, f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/resultados")

tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/resultados")

closePSD(doc)

# navegador.get("https://ge.globo.com/futebol/brasileirao-serie-b/")

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\tabela-brasileirao-parte1.psd")
doc = psApp.Application.ActiveDocument

criaClass(navegador, 0, 10)
mudaCor1()
colocaLogo()

print(rodada)
saveJpg(doc, f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/tabela-parte1")
tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/tabela-parte1")
closePSD(doc)

if (rodada != "FUTEBOL ALEMÃO"):
    psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\tabela-brasileirao-parte2.psd")
    doc = psApp.Application.ActiveDocument

    criaClass(navegador, 10, 20)
    mudaCor2()

else:
    psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\tabela-alemao-parte2.psd")
    doc = psApp.Application.ActiveDocument

    criaClass(navegador, 10, 18)
    mudaCor2()

colocaLogo()

saveJpg(doc,f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/tabela-parte2")
tiraLogo(f"F:/OneDrive - OPIC Telecom/Área de Trabalho/auto/pontos corridos/{camp}/tabela-parte2")
closePSD(doc)

psApp.Open(r"F:\OneDrive - OPIC Telecom\Área de Trabalho\auto\pontos corridos\artilharia-brasileirao.psd")
doc = psApp.Application.ActiveDocument

artilharia()

closePSD(doc)

navegador.close()