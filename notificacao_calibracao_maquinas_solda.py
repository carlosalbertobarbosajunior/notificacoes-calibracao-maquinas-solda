#Importando bibliotecas
import win32com.client
import datetime
import time
import sys
from dateutil.relativedelta import relativedelta

hoje = datetime.datetime.now()
ano = hoje.year
mes = hoje.month
dia = hoje.day
hora = hoje.hour
minuto = hoje.minute

#importando o pandas
import pandas as pd

#lendo o arquivo
listacalibracaosolda = pd.read_excel(
    "Q:/8. Soldagem/4- CALIBRAÇÃO MÁQ. SOLDA - ESTUFAS/2022/Controle de Máquinas de Solda.xlsx", sheet_name="Plan1")

# #renomeando o cabeçalho e removendo a primeira linha
# listacalibracaosolda = listacalibracaosolda.rename(columns=listacalibracaosolda.iloc[0]).drop([0], axis=0)

#filtrando máquinas que estejam com calibração vencida
maquinasatrasadas = listacalibracaosolda[(listacalibracaosolda["PRÓXIMA CALIBRAÇÃO"]<hoje)]


#Colocando os patrimônios em um vetor
maquinasatrasadas_patrimonios = maquinasatrasadas["TAG"].values
maquinasatrasadas_locais = maquinasatrasadas["LOCAL"].values
maquinasatrasadas_proxcalibracao = maquinasatrasadas["PRÓXIMA CALIBRAÇÃO"].values

tabelaatrasadas = {'TAG': maquinasatrasadas_patrimonios,
       'LOCAL': maquinasatrasadas_locais,
        'PRÓXIMA CALIBRAÇÃO': maquinasatrasadas_proxcalibracao
       }

dfatrasadas = pd.DataFrame(tabelaatrasadas)

bodyatrasadas = '<html><body>' + dfatrasadas.to_html() + '</body></html>'

# coluna_maquinasatrasadas_patrimonios = ('\n\t'.join(maquinasatrasadas_patrimonios))
# coluna_maquinasatrasadas_locais = ('\n\t'.join(maquinasatrasadas_locais))

numero_maquinasatrasadas = len(maquinasatrasadas_patrimonios)


#--------------------------------------------------------------------------------------------------------
#filtrando máquinas que faltam um mês para vencer
maquinasparavencer = listacalibracaosolda[(listacalibracaosolda["PRÓXIMA CALIBRAÇÃO"]>=hoje) & (listacalibracaosolda["PRÓXIMA CALIBRAÇÃO"]<(hoje+relativedelta(months=1)))]

#Colocando os patrimônios em um vetor
maquinasparavencer_patrimonios = maquinasparavencer["TAG"].values
maquinasparavencer_locais = maquinasparavencer["LOCAL"].values
maquinasparavencer_proxcalibracao = maquinasparavencer["PRÓXIMA CALIBRAÇÃO"].values

tabelaparavencer = {'TAG': maquinasparavencer_patrimonios,
       'LOCAL': maquinasparavencer_locais,
        'PRÓXIMA CALIBRAÇÃO': maquinasparavencer_proxcalibracao
       }

dfparavencer = pd.DataFrame(tabelaparavencer)

bodyparavencer = '<html><body>' + dfparavencer.to_html() + '</body></html>'

# coluna_maquinasparavencer_patrimonios = ('\n\t'.join(maquinasparavencer_patrimonios))
# coluna_maquinasparavencer_locais = ('\n\t'.join(maquinasparavencer_patrimonios))

numero_maquinasparavencer = len(maquinasparavencer_patrimonios)

#---------------------------------------------------------------------------------------------------------
#Mandando e-mail
if (numero_maquinasatrasadas != 0 ) or (numero_maquinasparavencer != 0):
    outlook = win32com.client.Dispatch("Outlook.Application")
    Msg = outlook.CreateItem(0)
    Msg.To = "carlos.junior@hkm.ind.br;angelo.silva@hkm.ind.br;ramon.novaes@hkm.ind.br;amaury.rodrigues@hkm.ind.br"
    Msg.Subject = "Resumo da Lista de Calibração de Máquinas de Solda"
    Msg.HTMLBody = f'''
    Bom dia!
    
    Existem algumas máquinas de solda que precisam de atenção na calibração.
    
    Há um total de {numero_maquinasatrasadas} máquinas com calibração vencida. São elas:
    
{bodyatrasadas}
    
    
    Há também {numero_maquinasparavencer} máquinas cuja calibração vencerá dentro de um mês. São elas:
    
{bodyparavencer}
    
    Este é um e-mail automático, mas sinta-se livre para respondê-lo.
    '''
    Msg.Send()
    


