#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from selenium import webdriver as opselenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempesp
navegador = opselenium.Chrome()
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

navegador.find_element_by_name("endereco").send_keys("13087000")

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
#____________________________________________________

from openpyxl import load_workbook
import os

nome_arquivo = "C:\\Users\\User1\\Desktop\\RPA1\\buscacep\\PesquisaEndereco.xlsx"
planDadosend = load_workbook(nome_arquivo)

sheet_selecionada = planDadosend['Dados']

linha = len(sheet_selecionada['A']) + 1
colunaA = "A" + str(linha)
colunaB = "B" + str(linha)
colunaC = "C" + str(linha)
colunaD = "D" + str(linha)

sheet_selecionada[colunaA] = rua
sheet_selecionada[colunaB] = bairro
sheet_selecionada[colunaC] = cidade
sheet_selecionada[colunaD] = cep

planDadosend.save(filename=nome_arquivo)

os.startfile(nome_arquivo)

