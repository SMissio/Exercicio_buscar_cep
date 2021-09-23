#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from openpyxl import load_workbook
import os

nome_arquivo_cep = "C:\\Users\\User1\\Desktop\\RPA1\\buscacep\\PesquisaEndereco.xlsx"
planDadosend = load_workbook(nome_arquivo_cep)

sheet_selecionada = planDadosend['CEP']

from selenium import webdriver as opselenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempesp
navegador = opselenium.Chrome()
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

navegador.find_element_by_name("endereco").send_keys("13087000")

tempesp.sleep(2)

navegador.find_element_by_name("btn_pesquisar").click()

tempesp.sleep(4)

for i in range(2,len(sheet_selecionada['A']) + 1):
    tempesp.sleep(4)
    
    navegador.find_element_by_name('btn_voltar').click()
    cepPesquisa = sheet_selecionada['A%s' % i].value
    tempesp.sleep(6)
    
    navegador.find_element_by_name("endereco").send_keys(cepPesquisa)

    tempesp.sleep(2)

    navegador.find_element_by_name("btn_pesquisar").click()

    tempesp.sleep(6)
    
    rua = navegador.find_element_by_xpath('//*[@id="resultado-DNEC"]/tbody/tr/td[1]').text
    print("Rua:",rua)

    bairro = navegador.find_element_by_xpath('//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text
    print("Bairro:",bairro)

    cidade = navegador.find_element_by_xpath('//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text
    print("Cidade:",cidade)

    cep = navegador.find_element_by_xpath('//*[@id="resultado-DNEC"]/tbody/tr/td[4]').text
    print("Cep:",cep)

    sheet_Dados = planDadosend['Dados']

    linha = len(sheet_Dados['A']) + 1
    colunaA = "A" + str(linha)
    colunaB = "B" + str(linha)
    colunaC = "C" + str(linha)
    colunaD = "D" + str(linha)

    sheet_Dados[colunaA] = rua
    sheet_Dados[colunaB] = bairro
    sheet_Dados[colunaC] = cidade
    sheet_Dados[colunaD] = cep

planDadosend.save(filename=nome_arquivo_cep)

os.startfile(nome_arquivo_cep)   

