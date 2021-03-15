import requests
import datetime as dt
import openpyxl

hoje = dt.date.today()

if hoje.weekday == 0:
    delta = dt.timedelta(days=3)
elif hoje.weekday== 6:
    delta = dt.timedelta(days=2)
else:
    delta = dt.timedelta(days=1)

data = hoje - delta
data_string = data.strftime('%m-%d-%Y')

# https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/aplicacao#!/recursos
url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='{}'&$top=100&$format=json".format(data_string)

req = requests.get(url)

if req.status_code == 200:
    cotacao = req.json().get('value')[0].get('cotacaoCompra')
    print(cotacao)
else:
    print('erro na requisição, tentar de novo')
'''
arquivo = 'dolar.xlsx'

excel = openpyxl.load_workbook(arquivo)
aba = excel.worksheets[0]
aba.insert_rows(0)
aba['A1'] = data.strftime('%d/%m/%Y')
aba['B1'] = cotacao
excel.save(arquivo)
'''
