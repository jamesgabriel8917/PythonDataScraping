import json
import os
from bs4 import BeautifulSoup
from selenium import webdriver
from time import sleep
from datetime import datetime, timedelta
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
import html_to_json
import csv
import unicodedata

# caminho relartorios e destinos de pastas
CaminhoRelatorioCSV = 'ListaRelatorios.csv'

#PathBase = 'C:\\Users\\james_fontana\\PycharmProjects\\BuscaDadosFluid\\Relatorios\\'

# Imprime Json *Arquivo tem que estar em String
def imprimeArquivo(PathArquivo, TipoEscrita, Arquivo):
    fp = open(PathArquivo, TipoEscrita)
    fp.write(Arquivo)
    fp.close()

# Efetua a leitura do arquivo CSV, nao passar delimitador para obter ; como padrão
def LeArquivoCSV(PathRelatorio, delimitador=";"):
    file = open(PathRelatorio)
    csvreader = csv.reader(file, delimiter=delimitador)

    return csvreader

# Efetua a consulta, retorna dados para a pasta designada e retorna status **Varivel navegador passado como objeto instanciado
def EfetuaConsulta(RowNumRelatorio, RowQtdResutados, RowInfoRelatorio, RowQtdDiasRetorno, navegador):
    # formatação de link
    data = datetime
    dia = str(data.today().day)
    mes = str(data.today().month)
    ano = str(data.today().year)

    CalculoRetroativo = data.today() + timedelta(int(RowQtdDiasRetorno))
    diaAnt = str(CalculoRetroativo.day)
    mesAnt = str(CalculoRetroativo.month)
    anoAnt = str(CalculoRetroativo.year)

    # LINK TESTE
    url = "https://[...]sso/id/{0}?dt_ini={1}%2F{2}%2F{3}&dt_fim={4}%2F{5}%2F{6}&situacao=0&itens_pagina={7}&go=true".format(
        str(RowNumRelatorio), diaAnt, mesAnt, anoAnt, dia, mes, ano, str(RowQtdResutados))

    navegador.get(url=url)

    # definição dos paths das pastas e relatorios
    dirPasta = PathBase + str(RowNumRelatorio)
    dirArquivo = PathBase + str(
        RowNumRelatorio) + '\\Processos.json'
    dirArquivoInfo = PathBase + str(
        RowNumRelatorio) + "\\" + RowInfoRelatorio.replace("/", "").replace(":", "") + ".txt"

    # cria as pastas dos relatorios
    if not os.path.exists(dirPasta):
        try:
            os.makedirs(dirPasta)
        except:
            print("Não foi possivel a criação da pasta " + str(dirPasta))
    try:
        if int(RowQtdResutados) <= 500:
            sleep(2)
        elif int(RowQtdResutados) > 500 and int(RowQtdResutados) <= 1000:
            sleep(4)
        elif int(RowQtdResutados) > 1001 and int(RowQtdResutados) <= 2500:
            sleep(20)
        elif int(RowQtdResutados) > 2501 and int(RowQtdResutados) <= 5000:
            sleep(40)
        elif int(RowQtdResutados) > 5001 and int(RowQtdResutados) <= 11000:
            sleep(60)
        # Encontra o elemento da tabela pelo Xpath completo
        # elemento = navegador.find_element_by_xpath("/html/body/div[2]/div[1]/div/div[2]/div[1]/table")
        elemento = navegador.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div[2]/div[1]/table")

        elemento_html = elemento.get_attribute('outerHTML')
        match RowNumRelatorio:
            case "877":
                strElement = str(elemento_html)
                #debug
                #imprimeArquivo("C:\\Users\\james_fontana\\Desktop\\Arquivotxtteste.txt", 'w', strElement)
                #print(strElement)
                strElement = strElement.replace(
                    "Limite Cheque Especial</th><th>Taxa Cheque Especial</th><th>Taxa do Cheque Especial</th><th>Número da Conta</th><th>Forma de Aprovação - Cheque Especial</th><th>Limite Cheque Especial",
                    "Limite Cheque Especial 2</th><th>Taxa Cheque Especial</th><th>Taxa do Cheque Especial</th><th>Número da Conta</th><th>Forma de Aprovação - Cheque Especial</th><th>Limite Cheque Especial 3").replace(
                    "Limite Cartão de Crédito</th><th>Limite cheque especial</th><th>Limites Aprovados?</th><th>Limites já Aprovados?</th><th>Limite Aprovado Cartão?</th><th>Limite Aprovado Cheque?</th><th>Limite Aprovado do Cheque?</th><th>Limite Cartão de Crédito",
                    "Limite Cartão de Crédito</th><th>Limite cheque especial 4</th><th>Limites Aprovados?</th><th>Limites já Aprovados?</th><th>Limite Aprovado Cartão?</th><th>Limite Aprovado Cheque?3</th><th>Limite Aprovado do Cheque?</th><th>Limite Cartão de Crédito 2").replace(
                    "<th>Taxa Cheque Especial</th><th>Taxa do Cheque Especial",
                    "<th>Taxa Cheque Especial 2</th><th>Taxa do Cheque Especial 3").replace(
                    "<th>Cheque Especial</th><th>Cheque Especial</th>",
                    "<th>Cheque Especial</th><th>Cheque Especial 2</th>").replace(
                    "Majorar para (Ch Especial):</th><th>Limite de Cheque Especial</th>",
                    "Majorar para (Ch Especial):</th><th>Limite de Cheque Especial 2</th>").replace(
                    "<th>Empresa é MEI?</th><th>Empresa é MEI?</th>",
                    "<th>Empresa é MEI?</th><th>Empresa é MEI 2</th>").replace(
                    "<th>Nº da Conta</th><th>Nº da Conta</th><th>Nº da Conta</th>",
                    "<th>Nº da Conta 1</th><th>Nº da Conta 2</th><th>Nº da Conta 3</th>").replace(
                    "Anuidade PJ</th><th>Qtd de Parcelas Anuidade</th>",
                    "Anuidade PJ</th><th>Qtd de Parcelas Anuidade 2</th>").replace(
                    "<th>Nome para Cartão</th><th>Email p/ Contato",
                    "<th>Nome para Cartão 2</th><th>Email p/ Contato").replace(
                    "<th>Nome para Cartão</th><th>Novo limite", "<th>Nome para Cartão 3</th><th>Novo limite").replace(
                    "<th>CPF do Associado</th><th>CNPJ/ CPF Empregador</th><th>CPF Associado</th><th>CPF*</th><th>CPF</th><th>CPF</th><th>CPF</th><th>CPF</th><th>CPF</th><th>CPF | Primeiro Titular</th><th>CPF do Associado</th><th>CPF</th><th>CPF</th><th>CNPJ</th><th>CNPJ</th><th>CNPJ Proponente</th>",
                    "<th>CPF do Associado</th><th>CNPJ/ CPF Empregador</th><th>CPF Associado</th><th>CPF*</th><th>CPF</th><th>CPF 2</th><th>CPF 3</th><th>CPF 4</th><th>CPF 5</th><th>CPF | Primeiro Titular</th><th>CPF do Associado 2</th><th>CPF 6</th><th>CPF 7</th><th>CNPJ</th><th>CNPJ 2</th><th>CNPJ Proponente</th>").replace(
                    "h><th>Cartão Solicitado</th><th>Nome para o Cartão</th>",
                    "h><th>Cartão Solicitado</th><th>Nome para o Cartão2</th> ").replace("<th>Celular Portador 4</th><th>CPF Portador 1</th><th>CPF Portador 2</th><th>CPF Portador 3</th><th>CPF Portador 4</th>","<th>Celular Portador 4</th><th>CPF Portador N 1</th><th>CPF Portador N 2</th><th>CPF Portador N 3</th><th>CPF Portador N 4</th>").replace("<th>Divisão de Limite entre Portadores</th><th>Nome Embossamento Portador 1</th><th>Nome Embossamento Portador 2</th><th>Nome Embossamento Portador 3</th>","<th>Divisão de Limite entre Portadores</th><th>Nome Embossamento P 1</th><th>Nome Embossamento P 2</th><th>Nome Embossamento P 3</th>").replace(
                    "<th>Cheque especial majorado para</th><th>Risco do Associado</th>",
                    "<th>Cheque especial majorado para</th><th>Risco do Associado2</th>"
                ).replace("</th><th>Forma de Pagamento</th><th>Nome para Cartão</th><th>Prestamista Cartão</th><th>Valor do Limite</th><th>Vencimento da Fatura Master</th><th>Vencimento Fatura Visa</th><th>Cartão</th><th>Já Possui Cartão?</th>",
                          "</th><th>Forma de Pagamento 2</th><th>Nome para Cartão 4</th><th>Prestamista Cartão 2</th><th>Valor do Limite</th><th>Vencimento da Fatura Master 2</th><th>Vencimento Fatura Visa 2</th><th>Cartão</th><th>Já Possui Cartão?</th>"
                          ).replace(
                    "</th><th>Vencimento Fatura Cartão Empresarial</th><th>Vencimento Fatura Visa</th>",
                    "</th><th>Vencimento Fatura Cartão Empresarial</th><th>Vencimento Fatura Visa 3</th>"
                ).replace(
                    "</th><th>Tipo do Cartão (Múltiplo)</th><th>Vencimento da Fatura Master</th>",
                    "</th><th>Tipo do Cartão (Múltiplo)</th><th>Vencimento da Fatura Master 3</th>"
                )
                soup = BeautifulSoup(strElement, 'html.parser')
            case "887":
                strElement = str(elemento_html)
                #imprimeArquivo("C:\\Users\\james_fontana\\Desktop\\Arquivotxt.txt", 'w', strElement)
                strElement = strElement.replace("<th>CPF do Cônjuge - Representante 1</th><th>CPF - Representante 1</th><th>CPF do Cônjuge - Representante 2</th><th>CPF - Representante 2</th><th>CPF do Cônjuge - Representante 3</th><th>CPF - Representante 3</th><th>CPF do Cônjuge - Representante 4</th><th>CPF - Representante 4</th><th>CPF do Cônjuge - Sócio 1</th><th>CPF - Sócio 1</th><th>CPF do Cônjuge - Sócio 2</th><th>CPF - Sócio 2</th><th>CPF do Cônjuge - Sócio 3</th><th>CPF - Sócio 3</th><th>CPF do Cônjuge - Sócio 4</th><th>CPF - Sócio 4</th><th>Atualização de Renda?</th><th>Comprovação de Renda</th><th>CNPJ</th><th>CPF do Associado</th><th>Alteração de Renda?</th><th>CPF Associado</th><th>CPF Portador (Cartão Adicional)</th><th>CPF*</th><th>CPF Cônjuge</th><th>CPF Cônjuge</th><th>CPF do Cônjuge</th><th>CPF</th><th>CPF</th><th>CPF</th><th>CNPJ</th><th>CPF do Associado</th><th>CPF Representante</th><th>CPF Representante 1</th><th>CPF Representante 2</th><th>CPF Representante 3</th><th>Nº da Conta</th><th>CNPJ Proponente</th><th>CPF</th><th>Tipo de Pessoa</th><th>CPF do Cônjuge - Sócio Empresário</th><th>CPF - Sócio/Empresário</th><th>CPF Portador 1</th><th>CPF Portador 2</th><th>CPF Portador 3</th><th>Comprovação de Renda</th><th>CPF | Primeiro Titular</th><th>CPF | Segundo Titular</th><th>CPF - Sócio 1</th><th>CPF - Sócio 2</th><th>CPF - Sócio 3</th><th>CPF - Sócio 4</th><th>CNPJ</th><th>CPF | Beneficiário 1</th><th>CPF | Beneficiário 2</th><th>CPF | Beneficiário 3</th><th>CPF Cônjuge | Primeiro Titular</th><th>CPF do Cônjuge | Primeiro Titular</th><th>CPF Associado 2 (BNDES Coletivo)</th><th>CPF cônjuge | Segundo Titular</th><th>CPF - Cônjuge | Segundo Titular</th><th>CPF do Cônjuge | Segundo Titular</th><th>CPF do Cônjuge</th><th>CPF do Cônjuge</th><th>CPF Cônjuge | Segundo Titular</th><th>CPF</th><th>CPF</th><th>CPF</th><th>CPF</th><th>CPF Cônjuge</th><th>CPF - Representante 1</th><th>CPF - Representante 2</th><th>CPF - Representante 3</th><th>CPF - Representante 4</th><th>CPF | Sócio/Empresário</th>",
                                                "<th>1</th><th>2</th><th>3</th><th>4</th><th>5</th><th>6</th><th>7</th><th>8</th><th>9</th><th>10</th><th>11</th><th>12</th><th>13</th><th>14</th><th>15</th><th>16</th><th>17</th><th>18</th><th>19</th><th>20</th><th>21</th><th>22</th><th>23</th><th>24</th><th>25</th><th>26</th><th>27</th><th>28</th><th>29</th><th>30</th><th>31</th><th>32</th><th>33</th><th>34</th><th>35</th><th>36</th><th>37</th><th>38</th><th>39</th><th>40</th><th>41</th><th>42</th><th>43</th><th>44</th><th>45</th><th>46</th><th>47</th><th>48</th><th>49</th><th>50</th><th>51</th><th>52</th><th>53</th><th>54</th><th>55</th><th>56</th><th>57</th><th>58</th><th>59</th><th>60</th><th>61</th><th>62</th><th>63</th><th>64</th><th>65</th><th>66</th><th>67</th><th>68</th><th>69</th><th>70</th><th>71</th><th>72</th><th>73</th><th>74</th><th>75</th>")
                soup = BeautifulSoup(strElement, 'html.parser')
            case "956":
                strElement = str(elemento_html)
                # imprimeArquivo("C:\\Users\\james_fontana\\Desktop\\Arquivotxt.txt", 'w', strElement)
                strElement = strElement.replace(
                    "<th>CPF | Primeiro Titular</th><th>CPF Representante</th><th>CPF</th><th>CPF Procurador</th><th>CPF Procurador 2</th><th>CPF Procurador 3</th><th>CPF - Sócio/Empresário</th><th>CPF do Cônjuge - Sócio Empresário</th><th>CPF | Segundo Titular</th><th>CPF Cônjuge | Primeiro Titular</th><th>CPF do Cônjuge | Primeiro Titular</th><th>CPF cônjuge | Segundo Titular</th><th>CPF - Cônjuge | Segundo Titular</th><th>CPF Cônjuge | Segundo Titular</th><th>CPF do Cônjuge | Segundo Titular</th><th>CPF - Representante 1</th><th>CPF do Cônjuge - Representante 1</th><th>CPF - Representante 2</th><th>CPF do Cônjuge - Representante 2</th><th>CPF - Representante 3</th><th>CPF do Cônjuge - Representante 3</th><th>CPF - Representante 4</th><th>CPF do Cônjuge - Representante 4</th><th>CPF - Sócio 1</th><th>CPF do Cônjuge - Sócio 1</th><th>CPF - Sócio 2</th><th>CPF do Cônjuge - Sócio 2</th><th>CPF - Sócio 3</th><th>CPF do Cônjuge - Sócio 3</th><th>CPF - Sócio 4</th><th>CPF do Cônjuge - Sócio 4</th><th>CPF - Representante 1</th><th>CPF - Representante 2</th><th>CPF - Representante 3</th><th>CPF - Representante 4</th><th>CPF - Representante 1</th><th>CPF - Representante 2</th><th>CPF - Representante 3</th><th>CPF - Representante 4</th><th>CPF - Sócio 1</th><th>CPF - Sócio 2</th><th>CPF - Sócio 3</th><th>CPF - Sócio 4</th><th>CNPJ Proponente</th>",
                    "<th>1</th><th>2</th><th>3</th><th>4</th><th>5</th><th>6</th><th>7</th><th>8</th><th>9</th><th>10</th><th>11</th><th>12</th><th>13</th><th>14</th><th>15</th><th>16</th><th>17</th><th>18</th><th>19</th><th>20</th><th>21</th><th>22</th><th>23</th><th>24</th><th>25</th><th>26</th><th>27</th><th>28</th><th>29</th><th>30</th><th>31</th><th>32</th><th>33</th><th>34</th><th>35</th><th>36</th><th>37</th><th>38</th><th>39</th><th>40</th><th>41</th><th>42</th><th>43</th><th>44</th>")
                soup = BeautifulSoup(strElement, 'html.parser')
            case "974":
                strElement = str(elemento_html)
                # imprimeArquivo("C:\\Users\\james_fontana\\Desktop\\Arquivotxt.txt", 'w', strElement)
                strElement = strElement.replace(
                    "<th>Do total investido, quanto gostaria de ter disponível para resgate a qualquer momento?</th><th>Nos últimos 2 anos, com que frequência você investiu parte da sua renda?</th><th>Nos últimos 2 anos, você investiu em Fundos Imobiliários, ETF ou BDR?</th><th>Nos últimos 2 anos, você investiu em Fundos Multimercado, Inflação ou Cambial?</th><th>Nos últimos 2 anos, você investiu em Poupança, Tesouro Direto, CDB, Fundos de Investimento Renda Fixa, LCA ou LCI?</th><th>Nos últimos 2 anos, você investiu em Ações, Fundos de Ações ou Debentures?</th><th>O quanto a sua formação ou experiência profissional ajudam você a entender sobre investimentos?</th><th>Por quanto tempo pretende deixar o dinheiro investido com a gente?</th><th>Qual o seu principal objetivo ao investir?</th><th>Qual o valor total da sua renda mensal?</th><th>Qual valor aproximadamente você investiu nos últimos 2 anos?</th><th>Qual valor aproximadamente você tem de patrimônio, considerando moveis, imóveis e investimentos financeiros?</th><th>Do total investido, quanto gostaria de ter disponível para resgate a qualquer momento?</th><th>Nos últimos 2 anos, a empresa investiu em Ações, Fundos de Ações ou Debentures?</th><th>Nos últimos 2 anos, a empresa investiu em CDB ou Fundos de Investimentos Renda Fixa?</th><th>Nos últimos 2 anos, a empresa investiu em Fundos Imobiliários, ETF ou BDR?</th><th>Nos últimos 2 anos, a empresa investiu em Fundos Multimercado, Inflação ou Cambial?</th><th>Nos últimos 2 anos, com que frequência a empresa realizou aplicações em produtos de investimentos?</th><th>Por quanto tempo pretende deixar o dinheiro investido com a gente?</th>",
                    "<th>Do total investido, quanto gostaria de ter disponível para resgate a qualquer momento?</th><th>Nos últimos 2 anos, com que frequência você investiu parte da sua renda?</th><th>Nos últimos 2 anos, você investiu em Fundos Imobiliários, ETF ou BDR?</th><th>Nos últimos 2 anos, você investiu em Fundos Multimercado, Inflação ou Cambial?</th><th>Nos últimos 2 anos, você investiu em Poupança, Tesouro Direto, CDB, Fundos de Investimento Renda Fixa, LCA ou LCI?</th><th>Nos últimos 2 anos, você investiu em Ações, Fundos de Ações ou Debentures?</th><th>O quanto a sua formação ou experiência profissional ajudam você a entender sobre investimentos?</th><th>Por quanto tempo pretende deixar o dinheiro investido com a gente?</th><th>Qual o seu principal objetivo ao investir?</th><th>Qual o valor total da sua renda mensal?</th><th>Qual valor aproximadamente você investiu nos últimos 2 anos?</th><th>Qual valor aproximadamente você tem de patrimônio, considerando moveis, imóveis e investimentos financeiros?</th><th>10</th><th>Nos últimos 2 anos, a empresa investiu em Ações, Fundos de Ações ou Debentures?</th><th>Nos últimos 2 anos, a empresa investiu em CDB ou Fundos de Investimentos Renda Fixa?</th><th>Nos últimos 2 anos, a empresa investiu em Fundos Imobiliários, ETF ou BDR?</th><th>Nos últimos 2 anos, a empresa investiu em Fundos Multimercado, Inflação ou Cambial?</th><th>Nos últimos 2 anos, com que frequência a empresa realizou aplicações em produtos de investimentos?</th><th>11</th>")
                soup = BeautifulSoup(strElement, 'html.parser')
            case "993":
                strElement = str(elemento_html)
                #imprimeArquivo("C:\\Users\\james_fontana\\Desktop\\ArquivotxtNelore.txt", 'w', strElement)
                strElement = strElement.replace("<th>Tipo de Alteração PF</th><th>Tipo de Alteração PJ</th><th>Motivo da Redução de Cheque Especial</th>",
                                                "<th>Tipo de Alteração PF2</th><th>Tipo de Alteração PJ2</th><th>Motivo da Redução de Cheque Especial</th>")
                soup = BeautifulSoup(strElement, 'html.parser')

            case _:
                soup = BeautifulSoup(elemento_html, 'html.parser')

        table = soup.find(name='table')
        # table = soup.find(name='table')

        # imprimeArquivo("testeretornotabela.txt", "w", str(table))
        outputJson = html_to_json.convert_tables(str(table))
        '''
        #formata o Json para o padrão seguido antes (lowerCase _ no lugar dos espaços)
        #
        #***********************************************************parte critica, atenção ao alterar******************************************************************
        #
        if RowNumRelatorio == "703":
            print(outputJson[0])
        '''
        for i in outputJson[0]:
            for j in i:
                aux = j
                aux = aux.replace("º", "")
                aux = unicodedata.normalize("NFKD", aux).encode('cp1252', 'ignore').decode("cp1252")
                aux = aux.replace(" ", "_").replace("-", "").replace("%", "").replace("?", "").replace("/", "").replace(
                    "|", "").lower()
                outputJson[0] = str(outputJson[0]).lower().replace(str(j).lower(), aux)
            break

        outputJson[0] = str(outputJson[0]).replace("[{", "{'processos':[{").replace("}]", "}]}")
        outputJson[0] = json.loads(json.dumps(outputJson[0]))

        # Importante! manter outputJson[0] para manter o index de impressão dos dados
        imprimeArquivo(dirArquivo, 'w', json.dumps(outputJson[0]).replace('"', ""))
        imprimeArquivo(dirArquivoInfo, 'w', "...")
        '''
        #
        #**********************************************************************************************************************************************************
        #
        '''
        return ('Imprimiu os arquivos do relatorio ' + str(RowNumRelatorio) + ' ' + str(datetime.now()))

    except Exception as err:
        imprimeArquivo(dirArquivo, 'w', "{'processos':[]}")
        imprimeArquivo(dirArquivoInfo, 'w', "...")
        #print(err)
        return 'Não localizou dados ao imprimir arquivos do relatorio, imprimindo Json Vazio ' + str(
            RowNumRelatorio) + ' ' + str(datetime.now())

# faz a busca no fluid, *passar como parâmetro: o id e a quantía máxima de resultados esperados
def BuscaDadosFluid():
    # Option define se será visível a extração
    option = Options()
    option.headless = True
    navegador = webdriver.Edge(executable_path=("msedgedriver.exe"), options=option)
    navegador.maximize_window()
    url = 'https://0731.fluid.prd.sicredi.cloud/usuario'
    navegador.get(url=url)

    email = ""
    senha =""

    # -----------Faz login-----------
    # Preenche o campo de email
    campo_email = navegador.find_element(By.XPATH, '/html/body/div/div[1]/section/div[1]/div/form/div[1]/div/input')
    campo_email.send_keys(email)
    # Preenche o campo de senha
    campo_senha = navegador.find_element(By.XPATH, '/html/body/div/div[1]/section/div[1]/div/form/div[2]/div/input')
    sleep(1)
    campo_senha.send_keys(senha)
    sleep(1)
    # Click no login
    navegador.find_element(By.XPATH, '/html/body/div/div[1]/section/div[1]/div/form/div[3]/div/button').click()
    sleep(10)
    Consulta = True

    while Consulta:
        # laço de repetição das consultas dos processos
        for row in LeArquivoCSV(CaminhoRelatorioCSV):
            hora = datetime
            if int(hora.now().hour) == 23 and int(hora.now().minute) > 50:
                sleep(900)

            # row[0] refere se ao numero do relatorio e row[1] a quantia de linhas retornadas
            # Efetua a consulta e retorna o resultado
            print(EfetuaConsulta(str(row[0]), str(row[1]), str(row[2]), str(row[3]), navegador))
    navegador.quit()

# Laço de repetição infinito

while True:
    try:
        BuscaDadosFluid()
    except:
        print("Erro. Reiniciando processo")
