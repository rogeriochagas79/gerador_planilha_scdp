import asyncio
import datetime
import json
import logging
import os
import random
import re
import time
import warnings
import xml.etree.ElementTree as ET
from collections import defaultdict
from functools import partial

import pandas as pd

import httpx
from aiometer import run_all
from docx import Document
from docx.shared import Pt
from httpx import AsyncClient
from lxml import etree, html


warnings.filterwarnings("ignore")

# --- Configuração de Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- VARIÁVEIS DE ACESSO E CONFIGURAÇÃO ---
GLOBAL_COOKIE_SCDP = 'INSERIR COOKIE'
GLOBAL_COOKIE_OFC = 'INSERIR COOKIE'
ID_TABELA = '1771003390211'
BASE_URL_SCDP = 'https://www2.scdp.gov.br'
BASE_URL_OFC = 'https://ofcweb.inss.gov.br'
HOST_SCDP = 'www2.scdp.gov.br'

SCDP_DETAIL_PATH = '/novoscdp/pages/consultar_solicitacao/consultar_solicitacao_detalhes.xhtml?idPcdp={id_scdp}&scdpTabId={idTabela}'
DIARIAS_REPORT_PATH = '/novoscdp/pages/relatorio/proposto/relatorio_proposto.xhtml'

# --- VARIAVEL ABA PLANILHA ---

abas_planilha_gt = {
'Acao Intinerante': ('000313/26', '000295/26', '000372/26', '000902/26', '000324/26', '000901/26'),
'COBEN': ('001424/26', '000152/26', '000876/26', '000883/26', '000461/26', '000743/26', '000771/26', '000826/26', '000833/26'),
'GT-Zika Virus- Anapolis': ('000161/26', '000162/26', '000168/26', '000180/26', '000177/26', '000169/26', '000166/26', '000176/26', '000175/26', '000184/26', '000171/26', '000191/26', '000206/26', '000172/26', '000173/26', '000179/26', '000163/26', '000164/26', '000193/26', '000182/26', '000183/26', '000170/26', '000167/26'),
'GT-Zika Virus- Brasilia': ('001138/26', '001106/26', '001103/26', '001101/26', '001102/26', '001104/26'),
'GT-Zika Virus- Macapá': ('001111/26', '001132/26', '001108/26', '001107/26'),
'GT-B21, B41 e B80- Anapolis': ('000985/26', '000925/26', '000961/26', '000921/26', '000928/26', '000922/26', '000927/26', '000936/26', '000931/26', '000932/26', '000959/26', '000924/26', '000882/26', '000937/26', '000940/26', '000939/26', '000934/26', '000684/26', '000685/26', '000686/26', '000689/26', '000687/26', '000690/26', '000691/26', '000692/26', '000693/26', '000694/26', '000695/26', '000696/26', '000698/26', '000705/26', '000701/26', '000709/26', '000710/26', '000699/26', '000818/26', '000872/26'),
'GT-B21, B41 e B80- Manaus': ('001345/26', '001346/26', '001347/26', '001348/26', '001350/26', '001352/26', '001354/26', '001356/26', '001416/26', '001442/26', '001305/26','000688/26', '000697/26', '000700/26', '000702/26', '000703/26', '000706/26', '000707/26', '000708/26', '000733/26', '000804/26'),
'GT-B21, B41 e B80- Belem': ('000712/26', '000714/26', '000738/26', '000716/26', '000718/26', '000719/26', '000742/26', '000813/26'),
'GT-B21, B41 e B80- Ji-Parana': ('000736/26','000737/26'),
'GT-B21, B41 e B80- Brasilia': ('001488/26', '001491/26', '001349/26', '001351/26', '001353/26', '001397/26'),
'GT-B21, B41 e B80- Caceres': ('001398/26', '001401/26', '001403/26', '001405/26', '001409/26', '001411/26', '001410/26', '001408/26', '001407/26', '001406/26', '001417/26','000717/26', '000720/26', '000739/26', '000721/26', '000722/26', '000723/26', '000724/26', '000725/26', '000726/26', '000727/26', '000715/26', '000805/26'),
'GT-B21, B41 e B80- Campo Grande': ('001489/26', '004004/26', '004002/26', '004000/26', '001399/26', '000729/26', '000730/26', '000731/26', '000713/26'),
'GT-B21, B41 e B80- Colider': ('001414/26', '001378/26', '001364/26', '001367/26', '001415/26'),
'GT-B21, B41 e B80- Castanhal': ('001355/26',),
'Reabilitacao': ('001182/26', '001185/26', '001129/26', '001073/26', '001127/26', '001130/26', '001123/26'),
'Servico Social': (),
'PreviBarco': ('001469/26', '001379/26', '001376/26', '001381/26', '001387/26', '001388/26','001222/26', '001226/26', '001087/26', '001288/26', '001169/26', '001167/26', '001168/26', '001166/26', '001165/26', '001134/26', '001164/26', '001084/26', '001056/26', '001121/26', '001095/26', '001087/26', '000814/26', '000908/26', '000913/26', '000917/26', '000918/26', '000796/26', '000983/26', '000808/26', '000810/26', '000794/26', '000812/26', '000454/26', '000816/26', '000481/26', '000814/26', '000795/26', '000811/26', '000810/26', '000330/26', '000333/26', '000785/26', '000791/26', '000807/26', '000808/26', '000331/26', '000334/26', '000796/26', '000809/26', '000332/26',	'000391/26', '000790/26')}


abas_planilha_multiroes = {
'Pericia Externa': (),
'PMF-Suporte': ('001139/26', '001133/26', '001136/26', '001141/26', '001140/26', '001142/26', '001083/26', '001063/26', '001064/26', '000366/26', '000364/26', '000368/26', '000365/26', '000377/26', '000415/26', '000371/26', '000480/26', '000325/26', '000474/26', '000369/26', '000482/26', '000400/26', '000412/26', '000477/26', '000470/26', '000419/26', '000370/26', '000414/26', '000367/26'),
'Janeiro': ('000033/26', '000012/26', '000014/26', '000015/26', '000011/26', '000018/26', '000010/26')}


abas_planilha_acao_itinerante = {
'Acre (AC)': ('016398/25', '016616/25', '016810/25', '016549/25', '016496/25', '007873/25', '007945/25', '006979/25', '007848/25', '007942/25', '007853/25', '011109/25', '011184/25', '011186/25', '011514/25'),
'Manaus (AM)': ('007947/25', '007944/25', '007948/25', '008108/25', '008110/25', '007946/25', '010858/25', '010859/25', '010951/25', '011957/25', '011958/25', '012120/25', '012231/25')}


# --- VARIAVEIS NOME ARQUIVOS ---
agora = datetime.datetime.now()
nome_base_lista = "lista_PCDP"
nome_arquivo_lista = nome_base_lista + ".json"
nome_arquivo_lista_bkp = agora.strftime(nome_base_lista + "_" + "%Y%m%d_%H%M%S" + ".json")

nome_base_planilha = "planilha_dados"
nome_arquivo_planilha = nome_base_planilha + ".xlsx"
nome_arquivo_planilha_bkp = agora.strftime(nome_base_planilha + "_" + "%Y%m%d_%H%M%S" + ".xlsx")

# --- FUNÇÕES  ---

def carregar_lista_consulta(arquivo):
    if os.path.exists(arquivo):
        try:
            return ler_json(arquivo)
        except json.JSONDecodeError:
            logging.warning("Erro ao decodificar o arquivo JSON parcial. Iniciando do zero.")
            return []
    return []

def gravar_json(lista, nome_arquivo):
    try:
        with open(nome_arquivo, 'w', encoding='utf-8') as arquivo:
            json.dump(lista, arquivo, ensure_ascii=False)
        logging.info(f"Lista gravada com sucesso em '{nome_arquivo}' com codificação UTF-8.")
    except Exception as e:
        logging.warning(f"Ocorreu um erro ao gravar o arquivo JSON: {e}")

def ler_json(nome_arquivo):

    try:
        with open(nome_arquivo, 'r', encoding='utf-8') as arquivo:
            lista_lida = json.load(arquivo)
        logging.info(f"Lista lida com sucesso de '{nome_arquivo}' com codificação UTF-8.")
        return lista_lida
    except FileNotFoundError:
        logging.error(f"Erro: Arquivo JSON '{nome_arquivo}' não encontrado.")
        return []
    except json.JSONDecodeError:
        logging.error(f"Erro: Arquivo JSON '{nome_arquivo}' contém um formato inválido.")
        return []
    except Exception as e:
        logging.error(f"Ocorreu um erro ao ler o arquivo JSON: {e}")
        return []

def extrair_javaViewState(resposta):

    arvoreDados = etree.HTML(resposta.content)
    javaViewState = arvoreDados.xpath("//input[@name='jakarta.faces.ViewState']/@value")
    if javaViewState:
        return javaViewState[0]
    logging.warning("jakarta.faces.ViewState não encontrado na resposta.")
    return ''

def obter_novos_id_scdp(resposta):
  
  arvoreDados = etree.HTML(resposta.content)
  lista_id_scdp = arvoreDados.xpath("//li[contains(@class,'ui-autocomplete-item ui-autocomplete-list-item')]/@data-item-value")
    
  return lista_id_scdp

def extrair_lista_idpcdp_do_json_carregado(lista_dados):
    lista_idPCDP = []
    for count, lista in enumerate(lista_dados, start=1):
        lista_idPCDP.append(lista['ID'])
    return lista_idPCDP

def encontrar_idpcdp_unicos(listaCarregada, listaNova):
  idsConsultar = set(listaNova).difference(set(listaCarregada))
  return list(idsConsultar)

def inserir_abas_gt(lista_sem_abas):
    for indice, dicionario in enumerate(lista_sem_abas, start=0):
        aba = "Outros"
        for chave, lista in abas_planilha_gt.items():
            if str(dicionario["SCDP"])[:9] in [str(item)[:9] for item in lista]:
                aba = chave
                break
        lista_sem_abas[indice].update({"Aba": aba})
    return lista_sem_abas   

def inserir_abas_multiroes(lista_sem_abas):
    for indice, dicionario in enumerate(lista_sem_abas, start=0):
        aba = "Outros"
        for chave, lista in abas_planilha_multiroes.items():
            if str(dicionario["SCDP"])[:9] in [str(item)[:9] for item in lista]:
                aba = chave
                break
        lista_sem_abas[indice].update({"Aba_PMF": aba})
    return lista_sem_abas   

def inserir_abas_acao_itinerante(lista_sem_abas):
    for indice, dicionario in enumerate(lista_sem_abas, start=0):
        aba = "Outros"
        for chave, lista in abas_planilha_acao_itinerante.items():
            if str(dicionario["SCDP"])[:9] in [str(item)[:9] for item in lista]:
                aba = chave
                break
        lista_sem_abas[indice].update({"Aba_Acao_Itinerante": aba})
    return lista_sem_abas

def extrair_funcao_cpf_matricula(resposta):
  arvoreDados = etree.HTML(resposta.content)
  funcao = arvoreDados.xpath("//*[@id='cabecalhoViagem:j_idt657:tabViewDadosProposto:j_idt659:j_idt674:j_idt674_text:outputText']/text()")[0]
  cpf = arvoreDados.xpath("//*[@id='cabecalhoViagem:j_idt657:tabViewDadosProposto:j_idt659:j_idt668:j_idt668_text:outputText']/text()")[0]
  matricula_siape = (arvoreDados.xpath("//*[@id='cabecalhoViagem:j_idt657:tabViewDadosProposto:j_idt659:j_idt665:j_idt665_text:outputText']/text()")[0]).strip()
  matricula_siape = matricula_siape[5:]  

  pattern = r'\d+\s-\s(.*)'

  match = re.search(pattern, funcao)

  if match:
      funcao = match.group(1)
      return funcao, cpf, matricula_siape
  else:
      return None

def obter_dados_scdp(resposta):
  arvoreDados = etree.HTML(resposta.content)
  numeroSCDP = str(arvoreDados.xpath("//span[@id='cabecalhoViagem:numSolicitacao:numSolicitacao_text:outputText']/text()")[0]).strip()[:9]
  orgaoSolicitante = arvoreDados.xpath("//*[@id='cabecalhoViagem:componenteCabecalhoViagem']/div/div[3]/div[2]/span/text()") 
  nomeDoProposto = arvoreDados.xpath("//a[@id='cabecalhoViagem:nomeProposto']/text()")
  tabelaRoteiro = arvoreDados.xpath("//span[contains(@id,'roteirosViagemSolicitacao')]/text()")
  valorDiarias = arvoreDados.xpath("//*[@id='qTotalEtapaSolicitacao:valorAtualNacional:valor']/text()")
  valorRestituicao = arvoreDados.xpath("//*[@id='qTotalEtapaSolicitacao:j_idt2633:valor']/text()")
  valorPassagem = arvoreDados.xpath("//*[@id='qTotalEtapaSolicitacao:valorSomaTarifaNacional:valor']/text()")
  posicaoDaPCDPnoFluxo = arvoreDados.xpath("//*[@id='cabecalhoViagem:pnInfoPcdp']/div[1]/div[3]/div/div[1]/a/text() | //*[@id='cabecalhoViagem:pnInfoPcdp']/div[1]/div[3]/div/div[1]/span/text()")
  posicaoDaPCDPnoFluxo = posicaoDaPCDPnoFluxo[0].strip()
  print(numeroSCDP)
  dict_dados_scdp = {
    'SCDP': numeroSCDP,
    'Orgão do Solicitante': orgaoSolicitante[0],
    'Nome do Proposto': nomeDoProposto[0],
    'Situação': posicaoDaPCDPnoFluxo,
    'Tabela Roteiro': tabelaRoteiro,
    'Valor das Diarias': valorDiarias[0],
    'Valor da Restituicao': valorRestituicao[0],
    'Valor da Passagens': valorPassagem[0]
  }

  return dict_dados_scdp

def gerar_dict_dados(resposta, resposta_popup, id_scdp):
    lts_funcao_cpf_matricula = extrair_funcao_cpf_matricula(resposta_popup)
    print(lts_funcao_cpf_matricula)
    dict_dados = obter_dados_scdp(resposta)

    if len(lts_funcao_cpf_matricula) >= 3:
        dict_dados.update({"Função": lts_funcao_cpf_matricula[0]})
        dict_dados.update({"CPF": lts_funcao_cpf_matricula[1]})
        dict_dados.update({"Matrícula": lts_funcao_cpf_matricula[2]})
    else:
        print("Erro: A lista lts_funcao_cpf_matricula não possui o número esperado de elementos.")
        return None  

    if 'Tabela Roteiro' in dict_dados and isinstance(dict_dados['Tabela Roteiro'], list):
        tabela = dict_dados['Tabela Roteiro']
        if len(tabela) > 1:
            origem = tabela[1]
        else:
            origem = ""
            print("Aviso: Tabela Roteiro não possui informações de origem suficientes.")

        if len(tabela) > 3:
            dataInicio = tabela[3]
        else:
            dataInicio = ""
            print("Aviso: Tabela Roteiro não possui informações de data de início suficientes.")

        destino = ""
        for indice in range(8, len(tabela), 9):
            if indice - 6 >= 0 and indice - 6 < len(tabela) and tabela[indice] != '---':
                destino = tabela[indice - 6]
                break

        if len(tabela) > 4:
            dataFim = tabela[len(tabela) - 4]
        else:
            dataFim = ""
            print("Aviso: Tabela Roteiro não possui informações de data de fim suficientes.")

        transporte = ""
        for indice in range(7, len(tabela), 9):
            if indice < len(tabela) and tabela[indice] == 'Aéreo':
                transporte = tabela[indice]
                break

        dict_dados.pop("Tabela Roteiro")
        dict_dados.update({"ID": id_scdp})
        dict_dados.update({"Origem": origem})
        dict_dados.update({"Destino": destino})
        dict_dados.update({"Transporte": transporte})
        dict_dados.update({"Data Inicial": dataInicio})
        dict_dados.update({"Data Final": dataFim})

    else:
        print("Erro: 'Tabela Roteiro' não encontrada no dict_dados ou não é uma lista.")
        return None  

    dict_dados['Diarias'] = 0.0
    dict_dados['APWEB'] = ''
    dict_dados['Aba'] = ''
    ordem_desejada = ['ID','Orgão do Solicitante','Diarias' ,'Nome do Proposto', 'CPF', 'Matrícula', 'SCDP', 'Data Inicial', 'Data Final', 'Origem', 'Destino', 'Transporte', 'Situação', 'APWEB', 'Valor das Diarias', 'Valor da Restituicao', 'Valor da Passagens', 'Aba']
    
    dict_dados = {chave: dict_dados[chave] for chave in ordem_desejada}

    return dict_dados

def unir_listas(lista_de_dados_scdp, lista_de_dados_scdp_novos):
    lista_unida = lista_de_dados_scdp + lista_de_dados_scdp_novos
    return lista_unida

def obter_dados_scdp_atualizacao(resposta):
  arvoreDados = etree.HTML(resposta.content)
  numeroSCDP = str(arvoreDados.xpath("//span[@id='cabecalhoViagem:numSolicitacao:numSolicitacao_text:outputText']/text()")[0]).strip()[:9]
  orgaoSolicitante = arvoreDados.xpath("//*[@id='cabecalhoViagem:componenteCabecalhoViagem']/div/div[3]/div[2]/span/text()") 
  nomeDoProposto = arvoreDados.xpath("//a[@id='cabecalhoViagem:nomeProposto']/text()")
  tabelaRoteiro = arvoreDados.xpath("//span[contains(@id,'roteirosViagemSolicitacao')]/text()")
  valorDiarias = arvoreDados.xpath("//*[@id='qTotalEtapaSolicitacao:valorAtualNacional:valor']/text()")
  valorRestituicao = arvoreDados.xpath("//*[@id='qTotalEtapaSolicitacao:j_idt2633:valor']/text()")
  valorPassagem = arvoreDados.xpath("//*[@id='qTotalEtapaSolicitacao:valorSomaTarifaNacional:valor']/text()")
  posicaoDaPCDPnoFluxo = arvoreDados.xpath("//*[@id='cabecalhoViagem:pnInfoPcdp']/div[1]/div[3]/div/div[1]/a/text() | //*[@id='cabecalhoViagem:pnInfoPcdp']/div[1]/div[3]/div/div[1]/span/text()")
  posicaoDaPCDPnoFluxo = posicaoDaPCDPnoFluxo[0].strip()
  dict_dados_scdp = {
    'SCDP': numeroSCDP,
    'Orgão do Solicitante': orgaoSolicitante[0],
    'Nome do Proposto': nomeDoProposto[0],
    'Situação': posicaoDaPCDPnoFluxo,
    'Tabela Roteiro': tabelaRoteiro,
    'Valor das Diarias': valorDiarias[0],
    'Valor da Restituicao': valorRestituicao[0],
    'Valor da Passagens': valorPassagem[0]
  }

  return dict_dados_scdp

def atualizar_dict_dados(resposta):
    dict_dados = obter_dados_scdp_atualizacao(resposta)
    
    if 'Tabela Roteiro' in dict_dados and isinstance(dict_dados['Tabela Roteiro'], list):
        tabela = dict_dados['Tabela Roteiro']
        if len(tabela) > 1:
            origem = tabela[1]
        else:
            origem = ""
            print("Aviso: Tabela Roteiro não possui informações de origem suficientes.")

        if len(tabela) > 3:
            dataInicio = tabela[3]
        else:
            dataInicio = ""
            print("Aviso: Tabela Roteiro não possui informações de data de início suficientes.")

        destino = ""
        for indice in range(8, len(tabela), 9):
            if indice - 6 >= 0 and indice - 6 < len(tabela) and tabela[indice] != '---':
                destino = tabela[indice - 6]
                break

        if len(tabela) > 4:
            dataFim = tabela[len(tabela) - 4]
        else:
            dataFim = ""
            print("Aviso: Tabela Roteiro não possui informações de data de fim suficientes.")

        transporte = ""
        for indice in range(7, len(tabela), 9):
            if indice < len(tabela) and tabela[indice] == 'Aéreo':
                transporte = tabela[indice]
                break

        dict_dados.pop("Tabela Roteiro")
        dict_dados.update({"Origem": origem})
        dict_dados.update({"Destino": destino})
        dict_dados.update({"Transporte": transporte})
        dict_dados.update({"Data Inicial": dataInicio})
        dict_dados.update({"Data Final": dataFim})


    else:
        print("Erro: 'Tabela Roteiro' não encontrada no dict_dados ou não é uma lista.")
        return None  
    
    return dict_dados

def obter_dados_diarias(resposta, indice):
  arvoreDados = etree.HTML(resposta.content)
  situacao = arvoreDados.xpath(f"//*[contains(@id, 'repResultadoPcdp:{str(indice)}:repResultadoTrechos:0:j_idt') and contains(@id, ':outputText')]/text()")[3]
  diarias = arvoreDados.xpath(f"//form[@id='formConsulta']//*[contains(@id, 'repResultadoPcdp:{str(indice)}:totalQtdDiarias:valor')]/text()")[0]
  
  return situacao, diarias

def obter_tamanho_da_tabela(resposta):
  arvoreDados = etree.HTML(resposta.content)
  tamanho = arvoreDados.xpath(f"//*[contains(@id, 'formConsulta:') and contains(@class, 'scdp_impressao')]")
     
  return (tamanho)

def obter_subotal_diarias(resposta):
  arvoreDados = etree.HTML(resposta.content)
  total = arvoreDados.xpath(f"//*[contains(@id, 'totalQtdDiariasRelatorio:valor')]/text()")[0]
  
  return total

def converter_string_to_float(texto):
    numero = float(texto.replace(",", "."))

    return numero 

def obter_total_diarias(resposta, tamanho_da_tabela):
    indice = 0
    qtdDiarias = 0
    while indice < tamanho_da_tabela:
        dados = obter_dados_diarias(resposta, indice)
        if dados[0] != 'Cancelada':
            qtdDiarias = qtdDiarias + converter_string_to_float(dados[1])
        indice += 1
    return qtdDiarias

def obter_situacao_ofc(resposta):
    
    try:
        arvoreDados = html.fromstring(resposta.content)
        tabSituacao = arvoreDados.xpath("//*[@id='tbl']/tbody[1]//tr[*]/td[11]")

        if not tabSituacao:
            return '' 

        situacao_referencia = tabSituacao[0].text_content().strip() 
        
        for celSituacao in tabSituacao[1:]:
            if situacao_referencia != celSituacao.text_content().strip():
                return 'VERIFICAR' 

        return situacao_referencia
    except Exception as e:
        logging.error(f"Erro ao processar a resposta HTML: {e}")
        return 'ERRO_PROCESSAMENTO'


# --- FUNÇÕES ACESSO SÍNCRONAS  ---

def acessar_pagina_diarias(cpf, javaViewState, max_retries=7, initial_delay=1):
    
    url = f'{BASE_URL_SCDP}{DIARIAS_REPORT_PATH}'
    
    headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-encoding': 'gzip, deflate, br, zstd',
        'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'cache-control': 'max-age=0',
        'connection': 'keep-alive',
        'content-type': 'application/x-www-form-urlencoded',
        'cookie': GLOBAL_COOKIE_SCDP,
        'host': HOST_SCDP,
        'origin': f'{BASE_URL_SCDP}',
        'referer': f'{BASE_URL_SCDP}{DIARIAS_REPORT_PATH}',
        'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36'
    }
    
    data = {
        'scdpTabId': ID_TABELA,
        'formConsulta':	'formConsulta',
        'formConsulta:selOrgao:autoComplete_input':	'Instituto Nacional do Seguro Social {Ministério da Previdência Social} [1934]',
        'formConsulta:selOrgao:autoComplete_hinput': 'Instituto Nacional do Seguro Social {Ministério da Previdência Social} [1934]',
        'formConsulta:selOrgao:tree_selOrgao_selection': '',
        'formConsulta:selOrgao:tree_selOrgao_scrollState':	'0,0',
        'formConsulta:tipoViagem':'',
        'formConsulta:tipoSolicitacao_selection':'',
        'formConsulta:tipoSolicitacao_scrollState':	'0,0',
        'formConsulta:cmbMotivoViagem_selection':'',
        'formConsulta:cmbMotivoViagem_scrollState':	'0,0',
        'formConsulta:situacaoViagem_selection':'',
        'formConsulta:situacaoViagem_scrollState':	'0,0',
        'formConsulta:txtNome_input':'',
        'formConsulta:txtNome_hinput':'',
        'formConsulta:txtCPF':	cpf,
        'formConsulta:dtInicioRelatorio:inputCalendario_input':	'01/01/2026',
        'formConsulta:dtFimRelatorio:inputCalendario_input':	'31/12/2026',
        'formConsulta:chkAnoExercicio':	'on',
        'formConsulta:btnPesquisar':'',
        'pesquisado': 'true',
        'gridPageIndex': '1',
        'jakarta.faces.ViewState': javaViewState
    }

    delay = initial_delay
    for tentativa in range(1, max_retries + 1):
        try:
            resposta = httpx.post(url, headers=headers, data=data, verify=False, follow_redirects=False, timeout=10)
            resposta.raise_for_status()  
            return resposta  
        except httpx.RequestError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro de requisição: {e}")
        except httpx.HTTPStatusError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
        except Exception as e:
            logging.warning(f"[Tentativa {tentativa}] Erro inesperado: {e}")

        if tentativa < max_retries:
            logging.info(f"[{cpf}] Aguardando {delay} segundos antes de tentar novamente...")
            asyncio.sleep(delay + random.uniform(0, 0.5))
            delay *= 2  

    logging.warning("Falha ao acessar a página após múltiplas tentativas.")
    return None

def acessar_pagina_popup(javaViewState, idPcdp, max_retries=7, initial_delay=1):
    
    url = f'{BASE_URL_SCDP}/novoscdp/pages/consultar_solicitacao/consultar_solicitacao_detalhes.xhtml'

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        'Accept': 'application/xml, text/xml, */*; q=0.01',
        'Origin': f'{BASE_URL_SCDP}',
        'Referer': f"{BASE_URL_SCDP}{SCDP_DETAIL_PATH.format(id_scdp=idPcdp, idTabela=ID_TABELA)}",
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cookie': GLOBAL_COOKIE_SCDP,  
        'Connection': 'keep-alive',
        'Faces-Request': 'partial/ajax',
        'X-Requested-With': 'XMLHttpRequest',
    }

    data = {
        'jakarta.faces.partial.ajax':	'true',
        'jakarta.faces.source':	'cabecalhoViagem:j_idt657:j_idt779',
        'jakarta.faces.partial.execute':	'@all',
        'jakarta.faces.partial.render':	'cabecalhoViagem:j_idt657:tabViewDadosProposto',
        'cabecalhoViagem:j_idt657:j_idt779':	'cabecalhoViagem:j_idt657:j_idt779',
        'formSolicitacaoDetalhes':	'formSolicitacaoDetalhes',
        'cabecalhoViagem:j_idt361:tabViewJustificativasViagem_activeIndex':	'0',
        'cabecalhoViagem:popupEncaminhamentosPCDPCabecalho:encaminhamentos_rppDD':	'10',
        'cabecalhoViagem:j_idt549:tabHistoricoGeral_activeIndex':	'0',
        'cabecalhoViagem:componenteCabecalhoViagem_collapsed':	'false',
        'cabecalhoViagem:j_idt657:tabViewDadosProposto_activeIndex':	'0',
        'anexoViagem:idAnexoHistorico': '',	
        'anexoViagem:componenteAnexoViagem_collapsed':	'false',
        'roteirosViagemSolicitacao:j_idt1103:resumoListaTrecho_rowExpansionState': '',	
        'roteirosViagemSolicitacao:fsRoteiroViagem_collapsed': 	'false',
        'calculoDiariasSolicitacao:gridDiariasNacionais_collapsed':	'false',
        'passagensNacionaisPlan:tabelaroteiro:0:recursosTrecho:fieldsetRecursos_collapsed':	'false',
        'passagensNacionaisPlan:tabelaroteiro:1:recursosTrecho:fieldsetRecursos_collapsed':	'false',
        'passagensNacionaisPlan:j_idt1222_collapsed':	'false',
        'quadTotalSolicitacao:j_idt2252_collapsed':	'false',
        'quadPagamentos:gridPagamento_rowExpansionState': '',	
        'quadPagamentos:gridPagamentos_collapsed':	'false',
        'qTotalEtapaSolicitacao:j_idt2519_collapsed':	'false',
        'quadValoresInformativosBilhete:componenteAnexoViagem_collapsed':	'false',
        'fsDetalhesDemo_collapsed':	'false',
        'qConsPrestacao:j_idt2799_collapsed':	'false',
        'planejamentos_selection': '',	
        'fsHistoricoProrrogacaoComplementacao_collapsed':	'false',
        'jakarta.faces.ViewState': javaViewState,
        'scdpTabId': ID_TABELA,
    }

    logging.debug(f"Acessando URL de detalhes SCDP: {url}")

    delay = initial_delay
    for tentativa in range(1, max_retries + 1):
        try:
            resposta = httpx.post(url, headers=headers, data=data, verify=False, follow_redirects=False, timeout=10)
            resposta.raise_for_status()  
            return resposta  
        except httpx.RequestError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro de requisição: {e}")
        except httpx.HTTPStatusError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
        except Exception as e:
            logging.warning(f"[Tentativa {tentativa}] Erro inesperado: {e}")

        if tentativa < max_retries:
            logging.info(f"[{idPcdp}] Aguardando {delay} segundos antes de tentar novamente...")
            asyncio.sleep(delay + random.uniform(0, 0.5))
            delay *= 2 

    logging.warning("Falha ao acessar a página após múltiplas tentativas.")
    return None

def acessar_pagina_principal(max_retries=7, initial_delay=1):
    url = 'https://www2.scdp.gov.br/novoscdp/pages/consultar_solicitacao/consultar_situacao_solicitacao_filter.xhtml'
    headers = {
        'Cookie': GLOBAL_COOKIE_SCDP,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'
    }

    delay = initial_delay
    for tentativa in range(1, max_retries + 1):
        try:
            resposta = httpx.get(url, headers=headers, verify=False, follow_redirects=False, timeout=10)
            resposta.raise_for_status() 
            return resposta  # sucesso
        except httpx.RequestError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro de requisição: {e}")
        except httpx.HTTPStatusError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
        except Exception as e:
            logging.warning(f"[Tentativa {tentativa}] Erro inesperado: {e}")

        if tentativa < max_retries:
            logging.info(f"Aguardando {delay} segundos antes de tentar novamente...")
            time.sleep(delay)
            delay *= 2 

    logging.warning("Falha ao acessar a página após múltiplas tentativas.")
    return None

def acessar_pagina_id_scdp(javaViewState, max_retries=7, initial_delay=1):
    url = 'https://www2.scdp.gov.br/novoscdp/pages/consultar_solicitacao/consultar_situacao_solicitacao_filter.xhtml'

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36',
        'Content-Type' : 'application/x-www-form-urlencoded',
        'Cookie' : GLOBAL_COOKIE_SCDP
    }

    data = {
      'jakarta.faces.partial.ajax' : 'true',
      'jakarta.faces.source' : 'tabViewSituacao:consultaSimples:numeroPcdpCmb',
      'jakarta.faces.partial.execute' : 'tabViewSituacao:consultaSimples:numeroPcdpCmb',
      'jakarta.faces.partial.render' : 'tabViewSituacao:consultaSimples:numeroPcdpCmb',
      'jakarta.faces.behavior.event' : 'query',
      'jakarta.faces.partial.event' : 'query',
      'tabViewSituacao:consultaSimples:numeroPcdpCmb_query' : '',
      'tabViewSituacao:consultaSimples' : 'tabViewSituacao:consultaSimples',
      'tabViewSituacao:consultaSimples:seletorTipoPcdp' : 'false',
      'tabViewSituacao:consultaSimples:numeroPcdpTxt' : '',
      'tabViewSituacao:consultaSimples:numeroPcdpCmb_input' : '',
      'tabViewSituacao:consultaSimples:numeroPcdpCmb_hinput' : '',
      'jakarta.faces.ViewState' : javaViewState
    }

    logging.debug(f"Acessando URL de detalhes SCDP: {url}")

    delay = initial_delay
    for tentativa in range(1, max_retries + 1):
        try:
            resposta = httpx.post(url, headers=headers, data=data, verify=False, follow_redirects=False, timeout=10)
            resposta.raise_for_status() 
            return resposta 
        except httpx.RequestError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro de requisição: {e}")
        except httpx.HTTPStatusError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
        except Exception as e:
            logging.warning(f"[Tentativa {tentativa}] Erro inesperado: {e}")

        if tentativa < max_retries:
            logging.info(f"Aguardando {delay} segundos antes de tentar novamente...")
            time.sleep(delay)
            delay *= 2  

    logging.warning("Falha ao acessar a página após múltiplas tentativas.")
    return None

def acessar_pagina_dados_scdp(id_scdp, max_retries=7, initial_delay=1):
    url = f"{BASE_URL_SCDP}{SCDP_DETAIL_PATH.format(id_scdp=id_scdp, idTabela=ID_TABELA)}"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
        'Cookie': GLOBAL_COOKIE_SCDP,        
    }

    logging.debug(f"Acessando URL de detalhes SCDP: {url}")

    delay = initial_delay
    for tentativa in range(1, max_retries + 1):
        try:
            resposta = httpx.get(url, headers=headers, verify=False, follow_redirects=False, timeout=10)
            resposta.raise_for_status()  
            return resposta  
        except httpx.RequestError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro de requisição: {e}")
        except httpx.HTTPStatusError as e:
            logging.warning(f"[Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
        except Exception as e:
            logging.warning(f"[Tentativa {tentativa}] Erro inesperado: {e}")

        if tentativa < max_retries:
            logging.info(f"Aguardando {delay} segundos antes de tentar novamente...")
            time.sleep(delay)
            delay *= 2 

    logging.warning("Falha ao acessar a página após múltiplas tentativas.")
    return None

# --- FUNÇÕES ACESSO ASSÍNCRONAS  ---

async def acessar_pagina_dados_scdp_async(id_scdp, max_retries=7, initial_delay=1):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        'Cookie': GLOBAL_COOKIE_SCDP,
    }
    url = f"{BASE_URL_SCDP}{SCDP_DETAIL_PATH.format(id_scdp=id_scdp, idTabela=ID_TABELA)}"

    logging.debug(f"Acessando URL de detalhes SCDP: {url}")

    delay = initial_delay
    async with httpx.AsyncClient(verify=False, follow_redirects=True, timeout=10) as client:
        for tentativa in range(1, max_retries + 1):
            try:
                logging.info(f"[{id_scdp}] Tentativa {tentativa} de acesso à página.")
                resposta = await client.get(url, headers=headers)
                resposta.raise_for_status()  
                logging.info(f"[{id_scdp}] Sucesso ao acessar a página principal. Status: {resposta.status_code}")
                return resposta  # Sucesso
            except httpx.RequestError as e:
                logging.warning(f"[{id_scdp}] [Tentativa {tentativa}] Erro de requisição: {e}")
            except httpx.HTTPStatusError as e:
                logging.warning(f"[{id_scdp}] [Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
            except Exception as e:
                logging.warning(f"[{id_scdp}] [Tentativa {tentativa}] Erro inesperado: {e}")

            if tentativa < max_retries:
                logging.info(f"[{id_scdp}] Aguardando {delay:.2f} segundos antes de tentar novamente...")
                await asyncio.sleep(delay + random.uniform(0, 0.5))
                delay *= 2 

    logging.warning(f"[{id_scdp}] Falha ao acessar a página após múltiplas tentativas.")
    return None

async def acessar_pagina_diarias_async(cpf, javaViewState, max_retries=7, initial_delay=1):
    
    headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-encoding': 'gzip, deflate, br, zstd',
        'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'cache-control': 'max-age=0',
        'connection': 'keep-alive',
        'content-type': 'application/x-www-form-urlencoded',
        'cookie': GLOBAL_COOKIE_SCDP,
        'host': HOST_SCDP,
        'origin': f'https://www2.scdp.gov.br',
        'referer': f'{BASE_URL_SCDP}{DIARIAS_REPORT_PATH}?m=true&scdpTabId={ID_TABELA}',
        'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36'
    }

    data = {
        'scdpTabId': ID_TABELA,
        'formConsulta': 'formConsulta',
        'formConsulta:selOrgao:autoComplete_input': 'Instituto Nacional do Seguro Social {Ministério da Previdência Social} [1934]',
        'formConsulta:selOrgao:autoComplete_hinput': 'Instituto Nacional do Seguro Social {Ministério da Previdência Social} [1934]',
        'formConsulta:selOrgao:tree_selOrgao_selection': '',
        'formConsulta:selOrgao:tree_selOrgao_scrollState': '0,0',
        'formConsulta:tipoViagem':'',
        'formConsulta:tipoSolicitacao_selection':'',
        'formConsulta:tipoSolicitacao_scrollState': '0,0',
        'formConsulta:cmbMotivoViagem_selection':'',
        'formConsulta:cmbMotivoViagem_scrollState': '0,0',
        'formConsulta:situacaoViagem_selection':'',
        'formConsulta:situacaoViagem_scrollState': '0,0',
        'formConsulta:txtNome_input':'',
        'formConsulta:txtNome_hinput':'',
        'formConsulta:txtCPF': cpf,
        'formConsulta:dtInicioRelatorio:inputCalendario_input': '01/01/2026',
        'formConsulta:dtFimRelatorio:inputCalendario_input': '31/12/2026',
        'formConsulta:chkAnoExercicio': 'on',
        'formConsulta:btnPesquisar':'',
        'pesquisado': 'true',
        'gridPageIndex': '1',
        'jakarta.faces.ViewState': javaViewState
    }

    url = f'{BASE_URL_SCDP}{DIARIAS_REPORT_PATH}'
    
    logging.debug(f"Acessando URL de relatório de diárias: {url} com CPF: {cpf}")
    
    delay = initial_delay
    async with httpx.AsyncClient(verify=False, follow_redirects=False, timeout=10) as client:
        for tentativa in range(1, max_retries + 1):
            try:
                logging.info(f"[{cpf}] Tentativa {tentativa} de acesso à página.")
                resposta = await client.post(url, headers=headers, data=data)
                resposta.raise_for_status()  
                logging.info(f"[{cpf}] Sucesso ao acessar a página principal. Status: {resposta.status_code}")
                return resposta  # Sucesso
            except httpx.RequestError as e:
                logging.warning(f"[{cpf}] [Tentativa {tentativa}] Erro de requisição: {e}")
            except httpx.HTTPStatusError as e:
                logging.warning(f"[{cpf}] [Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
            except Exception as e:
                logging.warning(f"[{cpf}] [Tentativa {tentativa}] Erro inesperado: {e}")

            if tentativa < max_retries:
                logging.info(f"[{cpf}] Aguardando {delay} segundos antes de tentar novamente...")
                await asyncio.sleep(delay + random.uniform(0, 0.5))
                delay *= 2 

    logging.warning(f"[{cpf}] Falha ao acessar a página após múltiplas tentativas.")
    return None

async def acessar_pagina_ofc_async(pcdp, max_retries=7, initial_delay=1):
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        'authority': 'ofcweb.inss.gov.br',
        'scheme': 'https',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-encoding': 'gzip, deflate, br, zstd',
        'accept-language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'cookie': GLOBAL_COOKIE_OFC,
        'priority': 'u=0, i',
        'referer': f'{BASE_URL_OFC}/View/Pcdp_Acompanhar.php',
        'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'content-type': 'application/x-www-form-urlencoded', # Adicionado para POST
    }
    
    data = {
        'order': 'id_pcdp:DESC',
        'filter': f'tb_pcdp2026.nr_pcdp:{pcdp}'
    }

    url = f'{BASE_URL_OFC}/View/Pcdp_Acompanhar.php?&page=1&order=id_pcdp:DESC&filter=tb_pcdp2026.nr_pcdp:{pcdp}'
    logging.debug(f"Enviando POST para OFC: {url} com dados: {data}")
   
    delay = initial_delay
    async with httpx.AsyncClient(verify=False, follow_redirects=False, timeout=10) as client:
        for tentativa in range(1, max_retries + 1):
            try:
                logging.info(f"[{pcdp}] Tentativa {tentativa} de acesso à página.")
                resposta = await client.post(url, headers=headers, data=data)
                resposta.raise_for_status() 
                logging.info(f"[{pcdp}] Sucesso ao acessar a página principal. Status: {resposta.status_code}")
                return resposta  # Sucesso
            except httpx.RequestError as e:
                logging.warning(f"[{pcdp}] [Tentativa {tentativa}] Erro de requisição: {e}")
            except httpx.HTTPStatusError as e:
                logging.warning(f"[{pcdp}] [Tentativa {tentativa}] Erro HTTP: {e.response.status_code}")
            except Exception as e:
                logging.warning(f"[{pcdp}] [Tentativa {tentativa}] Erro inesperado: {e}")

            if tentativa < max_retries:
                logging.info(f"[{pcdp}] Aguardando {delay} segundos antes de tentar novamente...")
                await asyncio.sleep(delay + random.uniform(0, 0.5))
                delay *= 2  # Backoff exponencial

    logging.warning(f"[{pcdp}] Falha ao acessar a página após múltiplas tentativas.")
    return None

# --- FUNÇÕES ATUALIZAÇÕES ASSÍNCRONAS  ---

async def atualizar_dados_scdp_async(lista_de_dados_scdp):

   # --- Processamento inicial para obter a lista de IdPcdps únicos ---
    lista_idPcdp = []
    for dicionario in lista_de_dados_scdp:
        #if str(dicionario['Situação']) != 'Viagem Encerrada' and 'Viagem Cancelada':
        lista_idPcdp.append(dicionario['ID'])
    lista_idPcdp = list(set(lista_idPcdp)) 

    # --- Criando o mapeamento IdPcdp -> LISTA DE REFERÊNCIAS aos dicionários originais ---
    # Usamos defaultdict para criar listas automaticamente para cada IdPcdp
    dados_por_idPcdp_multi = defaultdict(list)
    for item in lista_de_dados_scdp:
        dados_por_idPcdp_multi[item['ID']].append(item)

    # --- Criando as tarefas para o aiometer ---
    tarefas = []
    idPcdp_para_processamento_ordenado = [] # Manterá a ordem real das tarefas

    
    for idPcdp_item in lista_idPcdp: # Itera sobre a lista de IdPcdps já processada e única
        idPcdp_para_processamento_ordenado.append(idPcdp_item) # Guarda a ordem para o zip

        tarefas.append(
            partial(
                acessar_pagina_dados_scdp_async,
                idPcdp_item,
                
            )
        )

    logging.info(f"Total de tarefas a serem executadas: {len(tarefas)}")

    results = await run_all(
        tarefas,
        max_at_once=10,
        max_per_second=20
    )
    
    # --- Processando os resultados e atualizando TODOS os registros na lista_de_dados_scdp original ---
    for i, (idPcdp_processado, res) in enumerate(zip(idPcdp_para_processamento_ordenado, results)):
        if res:
            dados_idPcdp_novos = atualizar_dict_dados(res)
            
            # Atualizar TODOS os dicionários associados a estID
            if idPcdp_processado in dados_por_idPcdp_multi:
                for item_referencia in dados_por_idPcdp_multi[idPcdp_processado]:
                    item_referencia['Situação'] = dados_idPcdp_novos['Situação']
                    item_referencia['Valor das Diarias'] = dados_idPcdp_novos['Valor das Diarias']
                    item_referencia['Valor da Restituicao'] = dados_idPcdp_novos['Valor da Restituicao']
                    item_referencia['Valor da Passagens'] = dados_idPcdp_novos['Valor da Passagens']
                    item_referencia['Transporte'] = dados_idPcdp_novos['Transporte']
                    item_referencia['Data Inicial'] = dados_idPcdp_novos['Data Inicial']
                    item_referencia['Data Final'] = dados_idPcdp_novos['Data Final']
            else:
                logging.warning(f"ID {idPcdp_processado} processado não encontrado no mapeamento original. Impossível atualizar.")
        else:
            logging.warning(f"Tarefa {i} (ID: {idPcdp_processado}): Falha ao obter resposta.")
            if idPcdp_processado in dados_por_idPcdp_multi:
                for item_referencia in dados_por_idPcdp_multi[idPcdp_processado]:
                    item_referencia['Situação'] = 'Erro' 
                    item_referencia['Valor das Diarias'] = 'Erro'
                    item_referencia['Valor da Restituicao'] = 'Erro'
                    item_referencia['Valor da Passagens'] = 'Erro'
                    item_referencia['Transporte'] = 'Erro'
                    item_referencia['Data Inicial'] = 'Erro'
                    item_referencia['Data Final'] = 'Erro'
    
    return lista_de_dados_scdp

async def atualizar_diarias_async(lista_de_dados_scdp):
   
   # --- Processamento inicial para obter a lista de CPFs únicos ---
    lista_cpf = []
    for dicionario in lista_de_dados_scdp:
        lista_cpf.append(dicionario['CPF'])
    lista_cpf = list(set(lista_cpf)) # Garante CPFs únicos para as requisições

    # --- Criando o mapeamento CPF -> LISTA DE REFERÊNCIAS aos dicionários originais ---
    # Usamos defaultdict para criar listas automaticamente para cada CPF
    dados_por_cpf_multi = defaultdict(list)
    for item in lista_de_dados_scdp:
        dados_por_cpf_multi[item['CPF']].append(item)

    # --- Criando as tarefas para o aiometer ---
    tarefas = []
    cpfs_para_processamento_ordenado = [] 

    resposta = acessar_pagina_diarias("", "")
    javaViewState = extrair_javaViewState(resposta)
    print(f"JavaViewState: {javaViewState}")

    for cpf_item in lista_cpf: 
        cpfs_para_processamento_ordenado.append(cpf_item) 

        tarefas.append(
            partial(
                acessar_pagina_diarias_async,
                cpf_item,
                javaViewState
            )
        )

    logging.info(f"Total de tarefas a serem executadas: {len(tarefas)}")

    results = await run_all(
        tarefas,
        max_at_once=10,
        max_per_second=20
    )

    logging.info("\n--- Resultados das requisições e Atualização da Lista ---")

    # --- Processando os resultados e atualizando TODOS os registros na lista_de_dados_scdp original ---
    for i, (cpf_processado, res) in enumerate(zip(cpfs_para_processamento_ordenado, results)):
        if res:
            quantidade_diarias_nova = obter_total_diarias(res, len(obter_tamanho_da_tabela(res)))
            logging.info(f"Tarefa {i} (CPF: {cpf_processado}): Sucesso - Novas Diárias: {quantidade_diarias_nova}, Status: {res.status_code}")

            # Atualizar TODOS os dicionários associados a este CPF
            if cpf_processado in dados_por_cpf_multi:
                for item_referencia in dados_por_cpf_multi[cpf_processado]:
                    item_referencia['Diarias'] = quantidade_diarias_nova
            else:
                logging.warning(f"CPF {cpf_processado} processado não encontrado no mapeamento original. Impossível atualizar.")
        else:
            logging.warning(f"Tarefa {i} (CPF: {cpf_processado}): Falha ao obter resposta.")
            if cpf_processado in dados_por_cpf_multi:
                for item_referencia in dados_por_cpf_multi[cpf_processado]:
                    item_referencia['Diarias'] = 'Erro' 

    logging.info("\n--- Lista de Dados SCDP Original ATUALIZADA (TODOS os registros) ---")
   
    return lista_de_dados_scdp

async def atualizar_situacao_ofc_async(lista_de_dados_scdp):

   # --- Processamento inicial para obter a lista de PCDPs únicos ---
    lista_pcdp = []
    for dicionario in lista_de_dados_scdp:
        if str(dicionario['Situação']) != 'Viagem Encerrada' and 'Viagem Cancelada':
            lista_pcdp.append(dicionario['SCDP'])
    lista_pcdp = list(set(lista_pcdp)) 

    # --- Criando o mapeamento CPF -> LISTA DE REFERÊNCIAS aos dicionários originais ---
    # Usamos defaultdict para criar listas automaticamente para cada CPF
    dados_por_cpf_multi = defaultdict(list)
    for item in lista_de_dados_scdp:
        dados_por_cpf_multi[item['SCDP']].append(item)

    # --- Criando as tarefas para o aiometer ---
    tarefas = []
    pcdp_para_processamento_ordenado = [] 

    for pcdp_item in lista_pcdp: # Itera sobre a lista de PCDPs já processada e única
        pcdp_para_processamento_ordenado.append(pcdp_item) 

        tarefas.append(
            partial(
                acessar_pagina_ofc_async,
                pcdp_item.lstrip('0')
            )
        )

    logging.info(f"Total de tarefas a serem executadas: {len(tarefas)}")

    results = await run_all(
        tarefas,
        max_at_once=5,
        max_per_second=10
    )

    logging.info("\n--- Resultados das requisições e Atualização da Lista ---")

   

    # --- Processando os resultados e atualizando TODOS os registros na lista_de_dados_scdp original ---
    for i, (cpf_processado, res) in enumerate(zip(pcdp_para_processamento_ordenado, results)):
        if res:
            situacao_ofc_nova = obter_situacao_ofc(res).title()
            
            if cpf_processado in dados_por_cpf_multi:
                for item_referencia in dados_por_cpf_multi[cpf_processado]:
                    item_referencia['APWEB'] = situacao_ofc_nova
            else:
                logging.warning(f"CPF {cpf_processado} processado não encontrado no mapeamento original. Impossível atualizar.")
        else:
            logging.warning(f"Tarefa {i} (CPF: {cpf_processado}): Falha ao obter resposta.")
            if cpf_processado in dados_por_cpf_multi:
                for item_referencia in dados_por_cpf_multi[cpf_processado]:
                    item_referencia['APWEB'] = 'Erro' # Ou outro valor para indicar falha
    
    return lista_de_dados_scdp

# --- FUNÇÕES GERAR PLANILHAS  ---

def gerar_planilha_pcdp(lista_de_dados_scdp):
    ids_vistos = set()
    lista_sem_duplicados = []
    lista_sem_pmf = []
    listaRelatorio = []

    for dicionario in lista_de_dados_scdp:
        if dicionario['ID'] not in ids_vistos:
            lista_sem_duplicados.append(dicionario)
            ids_vistos.add(dicionario['ID'])
    
    for dicionario in lista_sem_duplicados:
        if str(dicionario['Situação']) != 'Viagem Encerrada' and str(dicionario['Situação']) != 'Viagem Cancelada':
            dicionario['SCDP'] = dicionario['SCDP'][:9]
            lista_sem_pmf.append(dicionario)
    
    for dicionario in lista_sem_pmf:
        if str(dicionario['Aba_PMF']) == 'Outros':
            dicionario['SCDP'] = dicionario['SCDP'][:9]
            listaRelatorio.append(dicionario)

    print(f"Total de registros: {len(listaRelatorio)}")

    df = pd.DataFrame(listaRelatorio)

    ordem = ['ID','Orgão do Solicitante','Diarias' ,'Nome do Proposto', 'CPF', 'Matrícula', 'SCDP', 'Data Inicial', 'Data Final', 'Origem', 'Destino', 'Transporte', 'Situação', 'APWEB', 'Valor das Diarias', 'Valor da Restituicao', 'Valor da Passagens', 'Aba', 'Aba_PMF']
    df = df[ordem]

    colunas_moeda = ['Valor das Diarias', 'Valor da Restituicao', 'Valor da Passagens']

    # Converter colunas moeda para float, se existirem
    for col in colunas_moeda:
        if col in df.columns:
            df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
        else:
            print(f"Atenção: coluna '{col}' não encontrada no DataFrame.")

    df_ordenado = df.sort_values(by=['Aba', 'Nome do Proposto']) # Ordenação das planilhas ['Aba', 'Situação', 'Nome do Proposto']
    df = df_ordenado

    abas_unicas = df['Aba'].unique()
    
    nome_arquivo_excel = 'PCDP - 2026.xlsx'
    writer = pd.ExcelWriter(nome_arquivo_excel, engine='xlsxwriter')
    workbook = writer.book

    # Dicionário de cores por Situação
    cores_situacao = {
        'Prestação de Contas': '#D8E4BC',
        'Execução Financeira': '#FCD5B4',
        'Aprovação da Autoridade Superior': '#DCE6F1',
        'Aprovação da Autoridade Superior da Prestação de Contas': '#DCE6F1',
        'Aprovação de Despesas': '#DCE6F1',
        'Aprovação do Proponente da Prestação de Contas': '#DCE6F1',
        'Cadastramento da Viagem': '#DCE6F1',
        'Aprovação de Despesas da Prestação de Contas': '#D8E4BC',
        'Devolução de Valores': '#D8E4BC',
        'Emissão': '#DCE6F1',
        'Reserva de Passagem': '#DCE6F1',
    }

    # Formatos
    formato_cabecalho = workbook.add_format({
        'bold': True, 'bg_color': 'black', 'font_color': 'white',
        'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': 'white'
    })

    formato_default = workbook.add_format({
        'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': 'white'
    })

    formato_moeda_default = workbook.add_format({
        'num_format': '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-',
        'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': 'white'
    })

    # Função para criar formatos coloridos
    def criar_formatos_cor(cor_hex):
        return {
            'texto': workbook.add_format({
                'bg_color': cor_hex,
                'align': 'center', 'valign': 'vcenter',
                'border': 1, 'border_color': 'white'
            }),
            'moeda': workbook.add_format({
                'bg_color': cor_hex,
                'num_format': '_-"R$" * #,##0.00_-;-"R$" * #,##0.00_-;_-"R$" * "-"??_-;_-@_-',
                'align': 'center', 'valign': 'vcenter',
                'border': 1, 'border_color': 'white'
            })
        }

    # Criar todos os formatos coloridos
    formatos_personalizados = {
        situacao: criar_formatos_cor(cor)
        for situacao, cor in cores_situacao.items()
    }

    for aba in abas_unicas:
        df_aba = df[df['Aba'] == aba].drop(columns=['Aba', 'ID', 'Orgão do Solicitante', 'Aba_PMF'], errors='ignore').reset_index(drop=True)
        df_aba.to_excel(writer, sheet_name=aba, index=False)

        worksheet = writer.sheets[aba]

        # Cabeçalho
        for col_num, value in enumerate(df_aba.columns.values):
            worksheet.write(0, col_num, value, formato_cabecalho)

        # Ajuste de largura
        for i, col_name in enumerate(df_aba.columns):
            largura_coluna = max(
                df_aba[col_name].astype(str).str.len().max(),
                len(col_name) + 2
            )
            worksheet.set_column(i, i, largura_coluna)

        num_linhas = df_aba.shape[0]
        num_colunas = df_aba.shape[1]

        for linha in range(num_linhas):
            situacao = df_aba.at[linha, 'Situação']
            formatos = formatos_personalizados.get(situacao, {'texto': formato_default, 'moeda': formato_moeda_default})

            for col_idx in range(num_colunas):
                valor = df_aba.iat[linha, col_idx]
                if pd.isna(valor):
                    valor = ''
                nome_coluna = df_aba.columns[col_idx]

                fmt = formatos['moeda'] if nome_coluna in colunas_moeda else formatos['texto']
                worksheet.write(linha + 1, col_idx, valor, fmt)

    writer.close()
    print(f'Arquivo Excel \"{nome_arquivo_excel}\" exportado com sucesso!')

# --- FUNÇÕES PRINCIPAL ---
async def main():
    
    # --- Carregando Dados ---
    
    lista_de_dados_scdp = ler_json('lista_PCDP.json')
    print(len(lista_de_dados_scdp))

    # --- Gerando Lista IDs Novos ---
    resposta = acessar_pagina_principal()
    javaViewState = extrair_javaViewState (resposta)
    resposta = acessar_pagina_id_scdp(javaViewState)
    listaNovosIds = obter_novos_id_scdp (resposta)
    listaIdsCarregagos = extrair_lista_idpcdp_do_json_carregado (lista_de_dados_scdp)
    listaIdsNovosToConsultar = encontrar_idpcdp_unicos(listaIdsCarregagos, listaNovosIds)
    print(len(listaIdsNovosToConsultar))
    
    # --- Acessando dados IDs Novos ---
    
    lista_de_dados_scdp_novos = []
    for count, id_scdp in enumerate(listaIdsNovosToConsultar, start=1):
        logging.info(f'{count} de {len(listaIdsNovosToConsultar)}')
        resposta = acessar_pagina_dados_scdp(id_scdp)
        javaViewState = extrair_javaViewState(resposta) 
        resposta_popup = acessar_pagina_popup(javaViewState, id_scdp)
        dict_dado = gerar_dict_dados(resposta, resposta_popup, id_scdp)
        lista_de_dados_scdp_novos.append(dict_dado)
    
    
    lista_de_dados_scdp = unir_listas(lista_de_dados_scdp, lista_de_dados_scdp_novos)
    lista_de_dados_scdp = await atualizar_dados_scdp_async(lista_de_dados_scdp)
    lista_de_dados_scdp = await atualizar_diarias_async(lista_de_dados_scdp)
    lista_de_dados_scdp = await atualizar_situacao_ofc_async(lista_de_dados_scdp)
    
    #lista_de_dados_scdp = ler_json('lista_PCDP.json')   
    lista_de_dados_scdp = inserir_abas_gt(lista_de_dados_scdp)
    lista_de_dados_scdp = inserir_abas_multiroes(lista_de_dados_scdp)
    lista_de_dados_scdp = inserir_abas_acao_itinerante(lista_de_dados_scdp)

    logging.info("\n--- Gravando Atualização ---")
    gravar_json(lista_de_dados_scdp, nome_arquivo_lista)
    gravar_json(lista_de_dados_scdp, nome_arquivo_lista_bkp)
    logging.info("\n--- Atualização Finalizada ---")
    logging.info("\n--- Gerando Planilhas ---")
    gerar_planilha_pcdp(lista_de_dados_scdp)
    logging.info("\n--- Geração de Planilhas Finalizada ---")
    
    
if __name__ == "__main__":
    asyncio.run(main())