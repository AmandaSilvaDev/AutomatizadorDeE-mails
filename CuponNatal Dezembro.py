import pandas as pd
import random
import math
df = pd.read_excel('C:\\Users\\vitor.souza\OneDrive - Sicoob\\Documentos - Sicoob UniRondônia\\27. Atualizações Python\\Natal Premiado\\Cupons Natal Premiado Novembro.xlsx', sheet_name='Tabela1')
dfAssociados = pd.read_excel('C:\\Users\\vitor.souza\\OneDrive - Sicoob\\Documentos - Sicoob UniRondônia\\27. Atualizações Python\\Natal Premiado\\Dados Associados.xlsx',dtype={'Número CPF/CNPJ':str})
dfAssociados = dfAssociados.drop_duplicates(subset='Nome Cliente')

dfBaseNova = pd.read_excel('C:\\Users\\vitor.souza\\OneDrive - Sicoob\\Documentos - Sicoob UniRondônia\\27. Atualizações Python\\Natal Premiado\\Cupons Natal Premiado Dezembro.xlsx',sheet_name='Valor Dezembro')
arquivo_columns = ['Nome do Cooperado', 'Valor Integralizado', 'Qtd Cupons', 'Código', 'Index','Valor Gasto','Valor Acumulado']

arquivo = pd.DataFrame(columns=arquivo_columns)
# Lista para armazenar códigos gerados
codigos_gerados = set()

def recalcular():
    for i in range(len(dfBaseNova)):
        nome = dfBaseNova.at[i, 'Nome do Cooperado']
        # Verifica se o nome está no outro DataFrame
        if nome in df['Nome do Cooperado'].values:
            # Encontra a linha correspondente no DataFrame df
            linha_baseAntiga = df[df['Nome do Cooperado'] == nome]
            # Obtém os valores da linha correspondente no DataFrame df
            valor_Integralizado_df = linha_baseAntiga['Valor Integralizado'].values[0]
            qtd_gasta = linha_baseAntiga['Qtd Cupons'].values[0] * 250
            # Atualiza o DataFrame dfBaseNova
            dfBaseNova.at[i, 'Valor Acumulado'] = dfBaseNova.at[i, 'Valor Integralizado'] + valor_Integralizado_df
            dfBaseNova.at[i, 'Valor Resto'] = dfBaseNova.at[i, 'Valor Acumulado']
            dfBaseNova.at[i, 'Valor Resto'] -= qtd_gasta
            dfBaseNova.at[i, 'Valor Integralizado Anterior'] = valor_Integralizado_df
            dfBaseNova.at[i, 'Valor Gasto'] = qtd_gasta
            dfBaseNova.at[i, 'Qtd Cupons'] = math.floor(dfBaseNova.at[i, 'Valor Resto'] / 250)
        else: 
            dfBaseNova.at[i, 'Qtd Cupons'] = math.floor(dfBaseNova.at[i, 'Valor Integralizado'] / 250)
            dfBaseNova.at[i, 'Valor Gasto'] = 0
            dfBaseNova.at[i, 'Valor Acumulado'] = dfBaseNova.at[i, 'Valor Integralizado']



        
def gerar_codigo2():
    
    for i in range(len(dfBaseNova)):
        qtd = dfBaseNova.at[i, 'Qtd Cupons']
        for cod in range(int(qtd)):
            while True:
                codigo_aleatorio = random.randint(50000, 99999)
                if codigo_aleatorio not in codigos_gerados:
                    # Adicionar código à lista de códigos gerados
                    codigos_gerados.add(codigo_aleatorio)
                    
                    # Criar uma nova linha no DataFrame arquivo
                    arquivo.loc[len(arquivo)] = [dfBaseNova.at[i, 'Nome do Cooperado'], 
                                                 dfBaseNova.at[i, 'Valor Integralizado'], 
                                                 dfBaseNova.at[i, 'Qtd Cupons'],
                                                 str(codigo_aleatorio),
                                                 str(cod),
                                                 dfBaseNova.at[i, 'Valor Gasto'],
                                                 dfBaseNova.at[i, 'Valor Acumulado']]
                    break
recalcular()

gerar_codigo2()
nova =  pd.merge(arquivo,dfAssociados, left_on= 'Nome do Cooperado', right_on= 'Nome Cliente', how="left")
nova.to_excel('C:\\Users\\vitor.souza\\OneDrive - Sicoob\\Documentos - Sicoob UniRondônia\\27. Atualizações Python\\Natal Premiado\\Base Dezembro Cupons.xlsx', index=False)
