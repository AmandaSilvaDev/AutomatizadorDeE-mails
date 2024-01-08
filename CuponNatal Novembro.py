import pandas as pd
import random

df = pd.read_excel('C:/Users/vitor.souza/OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Natal Premiado/Cupons Natal Premiado.xlsx')
dfAssociados = pd.read_excel('C:/Users/vitor.souza/OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Natal Premiado/Dados Associados.xlsx',dtype={'Número CPF/CNPJ':str})
dfAssociados = dfAssociados.drop_duplicates(subset='Nome Cliente')
df = df[['Nome do Cooperado', 'Valor integralizado', 'Qtd Cupons']]
arquivo_columns = ['Nome do Cooperado', 'Valor integralizado', 'Qtd Cupons', 'Código', 'Index']
arquivo = pd.DataFrame(columns=arquivo_columns)
# Lista para armazenar códigos gerados
codigos_gerados = set()

def gerar_codigo():
    for i in range(len(df)):
        qtd = df.at[i, 'Qtd Cupons']
        for cod in range(int(qtd)):
            while True:
                codigo_aleatorio = random.randint(10000, 49999)# outro vai ser de 49991 até 99999
                if codigo_aleatorio not in codigos_gerados:
                    # Adicionar código à lista de códigos gerados
                    codigos_gerados.add(codigo_aleatorio)
                    
                    # Criar uma nova linha no DataFrame arquivo
                    arquivo.loc[len(arquivo)] = [df.at[i, 'Nome do Cooperado'], 
                                                 df.at[i, 'Valor integralizado'], 
                                                 df.at[i, 'Qtd Cupons'],
                                                 str(codigo_aleatorio),
                                                 str(cod)]
                    break
            

gerar_codigo()
nova =  pd.merge(arquivo,dfAssociados, left_on= 'Nome do Cooperado', right_on= 'Nome Cliente', how="left")
nova.to_excel('C:/Users/vitor.souza/OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Natal Premiado/Cupons.xlsx', index=False)
