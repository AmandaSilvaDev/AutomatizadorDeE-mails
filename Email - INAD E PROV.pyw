from tkinter import *
import win32com.client as win32
import pandas as pd
import os
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
import pyautogui as pg

#Instancia o caminho da pasta para não precisar corrigir os acessos.
lista = os.getcwd().split("\\")
lista_dow = os.path.expanduser("~").split("\\")
path = ""
path_user = ""


#formata os diretorios
for item in lista:
    path += item + '/'
    if 'Automatização' in item:
        break

for item in lista_dow:
    path_user += item + '/'
    if 'Automatização' in item:
        break

#Programa de Envio de E-mails
def enviar_email(para,copia,copia_oculta,assunto,corpoHTML):
    outlook = win32.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = para
    message.Subject = assunto
    message.HTMLBody = corpoHTML
    message.cc = copia
    message.bcc = copia_oculta

    From = None
    for myEmailAddress in outlook.Session.Accounts:
        if "5018.bi" in str(myEmailAddress):
            From = myEmailAddress
            break

    if From != None:
        # This line basically calls the "mail.SendUsingAccount = xyz@email.com" outlook VBA command
        message._oleobj_.Invoke(*(64209, 0, 8, 0, From))

        message.Send()

diretorioPROV05 = path_user + '/OneDrive - Sicoob\Documentos - Sicoob UniRondônia/08.Prov 05/00.Dados/'

def encontrar_arquivo_mais_recente(diretorioPROV05):
    lista_arquivos = os.listdir(diretorioPROV05)

    if not lista_arquivos:
        print(f"Nenhum arquivo encontrado no diretório {diretorioPROV05}.")
        return None

    # Inicializa a data do arquivo mais recente
    data_mais_recente = None
    arquivo_mais_recente = None

    for nome_arquivo in lista_arquivos:
        if nome_arquivo.startswith("PROV05-") and nome_arquivo.endswith(".xlsx"):
            partes_nome = nome_arquivo.split('-')

            if len(partes_nome) == 4:
                data_arquivo = f"{partes_nome[3]}-{partes_nome[2]}-{partes_nome[1]}"

                if data_mais_recente is None or data_arquivo > data_mais_recente:
                    data_mais_recente = data_arquivo
                    arquivo_mais_recente = nome_arquivo

    if arquivo_mais_recente is not None:
        return diretorioPROV05+arquivo_mais_recente
    else:
        print("Nenhum arquivo válido encontrado.")


#Programa que executa o envio
def distribuicao(enviarsemanal):
    #-------------------- Inicia as bases -----------------------------------
    #Base Cobrança

    
    try:
        pasta = pd.DataFrame(os.listdir(path+"/filacobranca"), columns=['items'])
        arquivo = pasta[pasta['items'].str.contains('relatorioFilas')]['items'].values[0]
        colunas = ['Fila', 'Nome', ' CPF/CNPJ', 'Cooperativa Origem Dívida', 'PAC', 'Acionado', 'A Cobrar', 'Produto', 'Carteira', 'Risco', 'Contrato', 'Qtd. Dias Atraso', 'Valor Operação', 'Valor Atualizado']
        bs_cobrança = pd.read_csv(path+"/filacobranca/"+arquivo).reset_index()
        bs_cobrança.columns = colunas #bs_cobrança é a antiga bs_cobrança
    except:
        pg.alert('Erro na base de cobrança administrativa (Base Inadimplentes)')
    #Base Gerentes
    try:
        path_dCarteiras = path_user+"OneDrive - Sicoob\Documentos - Sicoob UniRondônia\_Dimensões Globais\dCarteiras.xlsx"
        df_dCarteiras = pd.read_excel(path_dCarteiras)
        df_dCarteiras = df_dCarteiras.iloc[:,[0,2,1,5,3]]
    except:
        pg.alert('Erro na base de Carteira')
    try:
        df_prov05 = pd.read_excel(encontrar_arquivo_mais_recente(diretorioPROV05))
        df_prov05 = df_prov05.iloc[:,[3,4,6,10,11,17,21,22,24,27,30,32,33,34,23]]
    except:
        pg.alert('Erro na base de Carteira')
    #Base E-mails
    
    try:

        path_Email =  path_user+"OneDrive - Sicoob\Documentos - Sicoob UniRondônia\_Dimensões Globais\dbaseEmail.xlsx"
        bs_emailGerente = pd.read_excel(path_Email, sheet_name='Gerente', usecols=['Agências','E-mails','Código Carteira']).dropna(subset='Agências')
        bs_emailGerentePA = pd.read_excel(path_Email, sheet_name='GerentePA', usecols=['Agências','E-mails']).dropna(subset='Agências')
        
    except:
        pg.alert('Erro na base de e-mails - Lista E-mails')
    #--------------------- Formatação de Arquivos ------------------------------
    # Cobrança
    print("iniciando a database")
    bs_cobrança['PA'] = bs_cobrança['PAC'].apply(lambda x: int(x.split(' - ')[0]))
    bs_cobrança['CPF/CNPJ Limpo'] = bs_cobrança[' CPF/CNPJ'].apply(lambda x: int(str(x).replace("-","").replace("/","").replace(".","")))
    bs_cobrança['Valor Atualizado'] = bs_cobrança['Valor Atualizado'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    bs_cobrança['Valor Operação'] = bs_cobrança['Valor Operação'].apply(lambda x: float(x.replace(".","").replace(",",".")))
    df_cobrança = bs_cobrança.merge(df_dCarteiras,how='inner',left_on='CPF/CNPJ Limpo',right_on='CPF/CNPJ').drop_duplicates()
    
    bs_geral = df_cobrança.iloc[:,[14,4,17,16,2,1,8,10,9,11,12,13,18,20]].sort_values('PA') # Base já formatada.
   
    #---------------------- Base de Gerentes 50000 -----------------------------------
    df_gerente = bs_geral[(bs_geral['Qtd. Dias Atraso'] <= 45)].sort_values(by=['Qtd. Dias Atraso',"Valor Operação"], ascending=[True,False]) 
    df_ProvGerente = df_prov05[df_prov05['Situação do Nível de Risco'] == "PIORA"]
    df_ProvGerente.loc[:, 'Dias em Atraso'] = df_ProvGerente['Dias em Atraso'].fillna(0).astype(int)
    df_ProvGerente.loc[:, 'Atraso Projetado Final do Mês'] = df_ProvGerente['Atraso Projetado Final do Mês'].fillna(0).astype(int)

    #---------------------- Base de GerentesPa 50000 E 200000-----------------------------------
    df_gerentePA = bs_geral[(bs_geral['Qtd. Dias Atraso'] <= 45) & (bs_geral['Valor Operação'] > 50000)].sort_values(by=['Qtd. Dias Atraso',"Valor Operação"], ascending=[True,False]) 
    df_ProvGerentePA = df_prov05[df_prov05['Situação do Nível de Risco'] == "PIORA"]
    df_ProvGerentePA.loc[:, 'Dias em Atraso'] = df_ProvGerentePA['Dias em Atraso'].fillna(0).astype(int)
    df_ProvGerentePA.loc[:, 'Atraso Projetado Final do Mês'] = df_ProvGerentePA['Atraso Projetado Final do Mês'].fillna(0).astype(int)


    #----------------------- Base Diretoria +200000-------------------------------------
    bs_diretoria = bs_geral[(bs_geral['Valor Operação'] >= 200000) & (bs_geral['Carteira'] != "PREJUÍZO")].sort_values(by=['Qtd. Dias Atraso',"Valor Operação"], ascending=[True,False]) 
    df_ordenadoDiretoria = df_prov05.sort_values(by='Variação de Provisão', ascending=False)
    bs_diretoriaPROV = df_ordenadoDiretoria.head(20)
    bs_diretoriaPROV.loc[:, 'Dias em Atraso'] = bs_diretoriaPROV['Dias em Atraso'].fillna(0).astype(int)
    bs_diretoriaPROV.loc[:, 'Atraso Projetado Final do Mês'] = bs_diretoriaPROV['Atraso Projetado Final do Mês'].fillna(0).astype(int)

  
  
    #------------------------ Conteúdo base dos e-mails --------------------------------
    valor_total = 0
    inicio_email =  f"""<body style="width: 100%">
        <table style="width: 100%; text-align: center">
        <tr>
            <td style="width: 35%"><img src="https://raw.githubusercontent.com/AmandaSilvaDev/Templete/a1c82edbf95e8fcab713ca0a23d582d8c3a9240a/logoSicoob.png" width="250"/></td> 
            <td style="width: 60%; text-align: center"><table><tr><h1 style="color: #003641;">PROV05 E INAD DE 1 A 45 DIAS DIÁRIO</h1></tr></table></td> 
        </tr>
        </table>
        <h2 style="width: 100%"><span style="color: #003641;"><strong>Ol&aacute;, Gerente!</strong></span></h2>
        <h3 style="width: 100%"><span style="color: #003641;">baixo segue listagem de contratos inadimplentes até 45 dias da sua Carteira.</span></h3>
        <h3 style="width: 100%"><span style="color: #003641;">Favor realizar tratativas e caso haja d&uacute;vidas, entrar em contato com o departamento de Recupera&ccedil;&atilde;o de Cr&eacute;dito para maiores informa&ccedil;&otilde;es.</span></h3>"""
   
    fim_email = """</tbody></table><h2><span style="color: #003641;"><strong>Este &eacute; um e-mail enviado de forma autom&aacute;tica. N&atilde;o responda. Para mais informações acesse o COBRANÇA ADMINISTRATIVA 2.0</strong></span></h2>
    
        </tbody></table><h2><span style="color: #003641;"><strong>Para mais informações acesse o COBRANÇA ADMINISTRATIVA 2.0</strong></span></h2>
        <blockquote>
        <p style="text-align: center;"><span style="color: #003641;"><img src="https://raw.githubusercontent.com/AmandaSilvaDev/Templete/a196701ab3437e88f0fb705cb94c2562f2fc44f2/logoTimeBi.png" alt="" width="250" height="200" /></span></p>
        """
    inicio_email_GerentePA = f"""<body style="width: 100%">
    <table style="width: 100%; text-align: center">
        <tr>
            <td style="width: 35%"><img src="https://raw.githubusercontent.com/AmandaSilvaDev/Templete/a1c82edbf95e8fcab713ca0a23d582d8c3a9240a/logoSicoob.png" width="250"/></td> 
            <td style="width: 60%; text-align: center"><table><tr><h1 style="color: #003641;"> PROV05 E INAD DE 1 A 45 DIAS DIÁRIO</h1></tr></table></td> 
        </tr>
    </table>
    <h2 style="width: 100%"><span style="color: #003641;"><strong>Ol&aacute;, Gerente!</strong></span></h2>
    <h3 style="width: 100%"><span style="color: #003641;">Abaixo segue listagem de contratos inadimplentes até 45 dias acima R$ 50 mil do seu PA.</span></h3>"""
    inicio_email_Diretoria = f"""<body style="width: 100%">
    <table style="width: 100%; text-align: center">
        <tr>
            <td style="width: 35%"><img src="https://raw.githubusercontent.com/AmandaSilvaDev/Templete/a1c82edbf95e8fcab713ca0a23d582d8c3a9240a/logoSicoob.png" width="250" alt="Logo Sicoob"/></td>
            <td style="width: 60%; text-align: center">
                <table>
                    <tr>
                        <h1 style="color: #003641;">DIRETORIA - PROV05 E INAD ACIMA DE R$ 200 MIL</h1>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <h2 style="width: 100%"><span style="color: #003641;"><strong>Ol&aacute;, Diretores!</strong></span></h2>
    <h3 style="width: 100%"><span style="color: #003641;">Abaixo segue listagem de contratos inadimplentes de toda a cooperativa acima de R$ 200 MIL.</span></h3>"""
   

    #E-mail Gerente ↓↓
    if(enviarsemanal == False):
        for ger in bs_emailGerente['Código Carteira']:
        
            if ger == "Código Carteira":
                pass
            if ger =="":
                pass
            
            else:
                df_email_gerente = bs_emailGerente[bs_emailGerente['Código Carteira'] == ger]['E-mails']
                email = (str(df_email_gerente).split())
                base_email_gerente= df_gerente[df_gerente['Código Carteira'] == ger] #base geral filtrada por gerente
                mailto =email[1]      

                bs_base = base_email_gerente
                pa = bs_base["PA"].max()
                nome_car = df_gerente[df_gerente['Código Carteira'] == ger]['Nome Carteira']
                nome_car = base_email_gerente['Nome Carteira'].max()
                #EMAIL GERENTE PA
                

                baseGerenteProv = df_ProvGerente[df_ProvGerente['Código Carteira'] == ger].sort_values(by='Variação de Provisão', ascending=False)
                baseGerenteProv = baseGerenteProv.head(10)

                valor_total = locale.currency(bs_base['Valor Atualizado'].sum(), symbol=True, grouping=True)
                valor_operacao = locale.currency(bs_base['Valor Operação'].sum(), symbol=True, grouping=True)
                num_linhas = len(bs_base)
                if(num_linhas == 0 ):
                    pass
                else:
                    # Formata o corpo do e-mail
                    corpo1 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Com o total de <span style="font-weight: 1000"><strong>{num_linhas}</strong></span> contratos, o valor total da inadimplência é de <span style="font-weight: 1000"><strong>{valor_total}</strong></span> e o valor total de contratos é de <span style="font-weight: 1000"><strong>{valor_operacao}</strong></span>.</p>
                                <table style="border-color: #003641; border-collapse: collapse; width: 100%" border="1">
                                    <tbody>
                                        <tr style="text-align: center; background-color: #003641; color: #ffffff">
                                            <th>AGÊNCIA</th>
                                            <th>NOME CARTEIRA</th>
                                            <th>CPF/CNPJ</th>
                                            <th>COOPERADO</th>
                                            <th>PRODUTO</th>
                                            <th>CONTRATO</th>
                                            <th>RISCO</th>
                                            <th>DIAS EM ATRASO</th>
                                            <th>VALOR OPERAÇÃO</th>
                                            <th>VALOR INADIMPLENTE</th>
                                        </tr>"""
                    for n in range(len(bs_base)):
                        col1 = f'<td style ="text-align: left;">{bs_base["PA"].values[n]}</td>'
                        col2 = f'<td style ="text-align: left;">{bs_base["Nome Carteira"].values[n]}</td>'
                        col3 = f'<td style ="text-align: left;">{bs_base[" CPF/CNPJ"].values[n]}</td>'
                        col4 = f'<td style ="text-align: left;">{bs_base["Nome"].values[n].upper()}</td>'
                        col5 = f'<td style ="text-align: center;">{bs_base["Carteira"].values[n].upper()}</td>'
                        col6 = f'<td style ="text-align: center;">{bs_base["Contrato"].values[n]}</td>'
                        col7 = f'<td style ="text-align: center;">{bs_base["Risco"].values[n]}</td>'
                        col8 = f'<td style ="text-align: center;">{bs_base["Qtd. Dias Atraso"].values[n]}</td>'
                        col9 = f'<td style ="text-align: right;">{locale.currency(bs_base["Valor Operação"].values[n], symbol=False, grouping=True)}</td>'
                        col10 = f'<td style ="text-align: right;">{locale.currency(bs_base["Valor Atualizado"].values[n], symbol=False, grouping=True)}</td>'
                        corpo1 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}</tr>'


                    valor_totalVariação = locale.currency(baseGerenteProv["Variação de Provisão"].sum(), symbol=True, grouping=True)
                    valor_totalSaldo = locale.currency(baseGerenteProv['Saldo Devedor'].sum(), symbol=True, grouping=True)
                    # Formata o corpo do e-mail PROV
                    corpo2 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Os 10 contratos em piora com maior variação da provisão do dia.</p>
                    <p style="color: #003641; text-align: center; font-size: 20px">A soma da variação é de <span style="font-weight: 1000"><strong>{valor_totalVariação}</strong></span> e o valor total do saldo devedor é <span style="font-weight: 1000"><strong>{valor_totalSaldo}</strong></span>.</p>
                        <table style="border-color: #003641; border-collapse: collapse; width: 100%; table-layout: auto" border="1">
                            <tbody>
                                <tr style="text-align: center; background-color: #003641; color: #ffffff">
                                    <th>AGÊNCIA</th>
                                    <th>NOME CARTEIRA</th>
                                    <th>CPF/CNPJ</th>
                                    <th>COOPERADO</th>
                                    <th>CONTRATO</th>
                                    <th>RISCO</th>
                                    <th>RISCO ATUAL</th>
                                    <th>RISCO PROJETADO</th>
                                    <th>DIAS EM ATRASO</th>
                                    <th>ATRASO PROJETADO</th>
                                    <th>MOTIVO DA ALTERAÇÃO</th>
                                    <th>VARIAÇÃO DA PROVISÃO</th>
                                    <th>SALDO DEVEDOR</th>                            
                                </tr>"""
                    for n in range(len(baseGerenteProv)):
                        col1 = f'<td style ="text-align: left;">{baseGerenteProv["Número PA Carteira"].values[n].astype(int)}</td>'
                        col2 = f'<td style ="text-align: left;">{baseGerenteProv["Nome Carteira"].values[n]}</td>'
                        col3 = f'<td style ="text-align: left;">{baseGerenteProv["Número CPF/CNPJ"].values[n]}</td>'
                        col4 = f'<td style ="text-align: left;">{baseGerenteProv["Nome Cliente"].values[n].upper()}</td>'
                        col5 = f'<td style ="text-align: center;">{baseGerenteProv["Contrato"].values[n].upper()}</td>'
                        col6 = f'<td style ="text-align: center;">{baseGerenteProv["Risco CRL"].values[n]}</td>'
                        col7 = f'<td style ="text-align: center;">{baseGerenteProv["Nivel Risco COP ou Atual"].values[n]}</td>'
                        col8 = f'<td style ="text-align: center;">{baseGerenteProv["Nível Risco Projetado"].values[n]}</td>'
                        col9 = f'<td style ="text-align: center;">{baseGerenteProv["Dias em Atraso"].values[n].astype(int)}</td>'
                        col10 = f'<td style ="text-align: center;">{baseGerenteProv["Atraso Projetado Final do Mês"].values[n].astype(int)}</td>'
                        col11 = f'<td style ="text-align: center;">{baseGerenteProv["Motivo da Alteração"].values[n]}</td>'
                        col12 = f'<td style ="text-align: right;">{locale.currency(baseGerenteProv["Variação de Provisão"].values[n], symbol=False, grouping=True)}</td>'
                        col13 = f'<td style ="text-align: right;">{locale.currency(baseGerenteProv["Saldo Devedor"].values[n], symbol=False, grouping=True)}</td>'
                        corpo2 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}{col11}{col12}{col13}</tr>'
                    valor_total = baseGerenteProv["Variação de Provisão"].sum()

                    corpo_email = f"{corpo1}</tbody></table>{corpo2}</tbody></table>"

                    tabela_html = inicio_email + corpo_email + fim_email
                    assunto = 'CARTEIRA '+ str(nome_car) + ' - PROV05 E CONTRATOS EM ATRASO DE ATÉ 45 DIAS' 
            
                    corpoHTML = tabela_html
                    para = str(mailto) #"amandal.silva@sicoob.com.br"
                    copia = ""
                    copia_oculta = ''#"unicng.bi@sicoob.com.br" #Cópia Oculta
                    
                    enviar_email(para,copia,copia_oculta,assunto,corpoHTML)
                    print("Enviou Gerente Carteira")
               
    if (enviarsemanal == True):
        for ger in bs_emailGerente['Código Carteira']:
        
            if ger == "Código Carteira":
                pass
            if ger =="":
                pass
            
            else:
                df_email_gerente = bs_emailGerente[bs_emailGerente['Código Carteira'] == ger]['E-mails']
                email = (str(df_email_gerente).split())
                base_email_gerente= df_gerente[df_gerente['Código Carteira'] == ger] #base geral filtrada por gerente
                mailto =email[1]      

                bs_base = base_email_gerente
                pa = bs_base["PA"].max()
                nome_car = df_gerente[df_gerente['Código Carteira'] == ger]['Nome Carteira']
                nome_car = base_email_gerente['Nome Carteira'].max()
                #EMAIL GERENTE PA
                df_email_gerentePA = bs_emailGerentePA[bs_emailGerentePA['Agências'] == pa]['E-mails']
                email2 = (str(df_email_gerentePA).split())
                emailGerentePA =email2[1]

                baseGerenteProv = df_ProvGerente[df_ProvGerente['Código Carteira'] == ger].sort_values(by='Variação de Provisão', ascending=False)
                baseGerenteProv = baseGerenteProv.head(10)

                valor_total = locale.currency(bs_base['Valor Atualizado'].sum(), symbol=True, grouping=True)
                valor_operacao = locale.currency(bs_base['Valor Operação'].sum(), symbol=True, grouping=True)
                num_linhas = len(bs_base)
                if(num_linhas == 0 ):
                    pass
                else:
                    # Formata o corpo do e-mail
                    corpo1 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Com o total de <span style="font-weight: 1000"><strong>{num_linhas}</strong></span> contratos, o valor total da inadimplência é de <span style="font-weight: 1000"><strong>{valor_total}</strong></span> e o valor total de contratos é de <span style="font-weight: 1000"><strong>{valor_operacao}</strong></span>.</p>
                                <table style="border-color: #003641; border-collapse: collapse; width: 100%" border="1">
                                    <tbody>
                                        <tr style="text-align: center; background-color: #003641; color: #ffffff">
                                            <th>AGÊNCIA</th>
                                            <th>NOME CARTEIRA</th>
                                            <th>CPF/CNPJ</th>
                                            <th>COOPERADO</th>
                                            <th>PRODUTO</th>
                                            <th>CONTRATO</th>
                                            <th>RISCO</th>
                                            <th>DIAS EM ATRASO</th>
                                            <th>VALOR OPERAÇÃO</th>
                                            <th>VALOR INADIMPLENTE</th>
                                        </tr>"""
                    for n in range(len(bs_base)):
                        col1 = f'<td style ="text-align: left;">{bs_base["PA"].values[n]}</td>'
                        col2 = f'<td style ="text-align: left;">{bs_base["Nome Carteira"].values[n]}</td>'
                        col3 = f'<td style ="text-align: left;">{bs_base[" CPF/CNPJ"].values[n]}</td>'
                        col4 = f'<td style ="text-align: left;">{bs_base["Nome"].values[n].upper()}</td>'
                        col5 = f'<td style ="text-align: center;">{bs_base["Carteira"].values[n].upper()}</td>'
                        col6 = f'<td style ="text-align: center;">{bs_base["Contrato"].values[n]}</td>'
                        col7 = f'<td style ="text-align: center;">{bs_base["Risco"].values[n]}</td>'
                        col8 = f'<td style ="text-align: center;">{bs_base["Qtd. Dias Atraso"].values[n]}</td>'
                        col9 = f'<td style ="text-align: right;">{locale.currency(bs_base["Valor Operação"].values[n], symbol=False, grouping=True)}</td>'
                        col10 = f'<td style ="text-align: right;">{locale.currency(bs_base["Valor Atualizado"].values[n], symbol=False, grouping=True)}</td>'
                        corpo1 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}</tr>'


                    valor_totalVariação = locale.currency(baseGerenteProv["Variação de Provisão"].sum(), symbol=True, grouping=True)
                    valor_totalSaldo = locale.currency(baseGerenteProv['Saldo Devedor'].sum(), symbol=True, grouping=True)
                    # Formata o corpo do e-mail PROV
                    corpo2 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Os 10 contratos em piora com maior variação da provisão do dia.</p>
                    <p style="color: #003641; text-align: center; font-size: 20px">A soma da variação é de <span style="font-weight: 1000"><strong>{valor_totalVariação}</strong></span> e o valor total do saldo devedor é <span style="font-weight: 1000"><strong>{valor_totalSaldo}</strong></span>.</p>
                        <table style="border-color: #003641; border-collapse: collapse; width: 100%; table-layout: auto" border="1">
                            <tbody>
                                <tr style="text-align: center; background-color: #003641; color: #ffffff">
                                    <th>AGÊNCIA</th>
                                    <th>NOME CARTEIRA</th>
                                    <th>CPF/CNPJ</th>
                                    <th>COOPERADO</th>
                                    <th>CONTRATO</th>
                                    <th>RISCO</th>
                                    <th>RISCO ATUAL</th>
                                    <th>RISCO PROJETADO</th>
                                    <th>DIAS EM ATRASO</th>
                                    <th>ATRASO PROJETADO</th>
                                    <th>MOTIVO DA ALTERAÇÃO</th>
                                    <th>VARIAÇÃO DA PROVISÃO</th>
                                    <th>SALDO DEVEDOR</th>                            
                                </tr>"""
                    for n in range(len(baseGerenteProv)):
                        col1 = f'<td style ="text-align: left;">{baseGerenteProv["Número PA Carteira"].values[n].astype(int)}</td>'
                        col2 = f'<td style ="text-align: left;">{baseGerenteProv["Nome Carteira"].values[n]}</td>'
                        col3 = f'<td style ="text-align: left;">{baseGerenteProv["Número CPF/CNPJ"].values[n]}</td>'
                        col4 = f'<td style ="text-align: left;">{baseGerenteProv["Nome Cliente"].values[n].upper()}</td>'
                        col5 = f'<td style ="text-align: center;">{baseGerenteProv["Contrato"].values[n].upper()}</td>'
                        col6 = f'<td style ="text-align: center;">{baseGerenteProv["Risco CRL"].values[n]}</td>'
                        col7 = f'<td style ="text-align: center;">{baseGerenteProv["Nivel Risco COP ou Atual"].values[n]}</td>'
                        col8 = f'<td style ="text-align: center;">{baseGerenteProv["Nível Risco Projetado"].values[n]}</td>'
                        col9 = f'<td style ="text-align: center;">{baseGerenteProv["Dias em Atraso"].values[n].astype(int)}</td>'
                        col10 = f'<td style ="text-align: center;">{baseGerenteProv["Atraso Projetado Final do Mês"].values[n].astype(int)}</td>'
                        col11 = f'<td style ="text-align: center;">{baseGerenteProv["Motivo da Alteração"].values[n]}</td>'
                        col12 = f'<td style ="text-align: right;">{locale.currency(baseGerenteProv["Variação de Provisão"].values[n], symbol=False, grouping=True)}</td>'
                        col13 = f'<td style ="text-align: right;">{locale.currency(baseGerenteProv["Saldo Devedor"].values[n], symbol=False, grouping=True)}</td>'
                        corpo2 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}{col11}{col12}{col13}</tr>'
                    valor_total = baseGerenteProv["Variação de Provisão"].sum()

                    corpo_email = f"{corpo1}</tbody></table>{corpo2}</tbody></table>"

                    tabela_html = inicio_email + corpo_email + fim_email
                    assunto = 'CARTEIRA '+ str(nome_car) + ' - RELATÓRIO SEMANAL PROV05 E CONTRATOS EM ATRASO DE ATÉ 45 DIAS' 
            
                    corpoHTML = tabela_html
                    para = str(mailto) #"amandal.silva@sicoob.com.br"
                    copia = str(emailGerentePA)
                    copia_oculta = ''#"unicng.bi@sicoob.com.br" #Cópia Oculta
                    
                    enviar_email(para,copia,copia_oculta,assunto,corpoHTML)
                    print("Enviou Gerente Carteira Semanal com copia para Gerente do PA")
        else:
            pass   

        #E-mail Gerente PA ↓↓
    for sup in bs_emailGerentePA['Agências']:
        if sup == "Agências":
            pass
        if sup =="":
            pass
        
        else:
            
            df_email_gerentePA = bs_emailGerentePA[bs_emailGerentePA['Agências'] == sup]['E-mails']
            email = (str(df_email_gerentePA).split())
            base_email_gerentePA= df_gerentePA[df_gerentePA['PA'] ==sup]
            mailto =email[1]
            bs_base = base_email_gerentePA
            valor_total = locale.currency(bs_base['Valor Atualizado'].sum(), symbol=True, grouping=True)
            valor_operacao = locale.currency(bs_base['Valor Operação'].sum(), symbol=True, grouping=True)
            num_linhas = len(bs_base)

            baseGerentePAProv = df_ProvGerentePA[df_ProvGerentePA['Número PA Carteira'] ==sup].sort_values(by='Variação de Provisão', ascending=False)
            baseGerentePAProv = baseGerentePAProv.head(20)


            if(num_linhas == 0 ):
                pass
            else:
            # Formata o corpo do e-mail
                corpo1 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Com o total de <span style="font-weight: 1000"><strong>{num_linhas}</strong></span> contratos, o valor total da inadimplência é de <span style="font-weight: 1000"><strong>{valor_total}</strong></span> e o valor total de contratos é de <span style="font-weight: 1000"><strong>{valor_operacao}</strong></span>.</p>
                            <table style="border-color: #003641; border-collapse: collapse; width: 100%" border="1">
                                <tbody>
                                    <tr style="text-align: center; background-color: #003641; color: #ffffff">
                                        <th>AGÊNCIA</th>
                                        <th>NOME CARTEIRA</th>
                                        <th>CPF/CNPJ</th>
                                        <th>COOPERADO</th>
                                        <th>PRODUTO</th>
                                        <th>CONTRATO</th>
                                        <th>RISCO</th>
                                        <th>DIAS EM ATRASO</th>
                                        <th>VALOR OPERAÇÃO</th>
                                        <th>VALOR INADIMPLENTE</th>
                                    </tr>"""
                for n in range(len(bs_base)):
                    col1 = f'<td style ="text-align: left;">{bs_base["PA"].values[n]}</td>'
                    col2 = f'<td style ="text-align: left;">{bs_base["Nome Carteira"].values[n]}</td>'
                    col3 = f'<td style ="text-align: left;">{bs_base[" CPF/CNPJ"].values[n]}</td>'
                    col4 = f'<td style ="text-align: left;">{bs_base["Nome"].values[n].upper()}</td>'
                    col5 = f'<td style ="text-align: center;">{bs_base["Carteira"].values[n].upper()}</td>'
                    col6 = f'<td style ="text-align: center;">{bs_base["Contrato"].values[n]}</td>'
                    col7 = f'<td style ="text-align: center;">{bs_base["Risco"].values[n]}</td>'
                    col8 = f'<td style ="text-align: center;">{bs_base["Qtd. Dias Atraso"].values[n]}</td>'
                    col9 = f'<td style ="text-align: right;">{locale.currency(bs_base["Valor Operação"].values[n], symbol=False, grouping=True)}</td>'
                    col10 = f'<td style ="text-align: right;">{locale.currency(bs_base["Valor Atualizado"].values[n], symbol=False, grouping=True)}</td>'
                    corpo1 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}</tr>'
                
                valor_totalVariação = locale.currency(baseGerentePAProv["Variação de Provisão"].sum(), symbol=True, grouping=True)
                valor_totalSaldo = locale.currency(baseGerentePAProv['Saldo Devedor'].sum(), symbol=True, grouping=True)
                # Formata o corpo do e-mail PROV
                corpo2 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Os 20 contratos em piora com maior variação da provisão do dia.</p>
                <p style="color: #003641; text-align: center; font-size: 20px">A soma da variação é de <span style="font-weight: 1000"><strong>{valor_totalVariação}</strong></span> e o valor total do saldo devedor é <span style="font-weight: 1000"><strong>{valor_totalSaldo}</strong></span>.</p>
                    <table style="border-color: #003641; border-collapse: collapse; width: 100%; table-layout: auto" border="1">
                        <tbody>
                            <tr style="text-align: center; background-color: #003641; color: #ffffff">
                                <th>AGÊNCIA</th>
                                <th>NOME CARTEIRA</th>
                                <th>CPF/CNPJ</th>
                                <th>COOPERADO</th>
                                <th>CONTRATO</th>
                                <th>RISCO</th>
                                <th>RISCO ATUAL</th>
                                <th>RISCO PROJETADO</th>
                                <th>DIAS EM ATRASO</th>
                                <th>ATRASO PROJETADO</th>
                                <th>MOTIVO DA ALTERAÇÃO</th>
                                <th>VARIAÇÃO DA PROVISÃO</th>
                                <th>SALDO DEVEDOR</th>                            
                            </tr>"""
                for n in range(len(baseGerentePAProv)):
                    col1 = f'<td style ="text-align: left;">{baseGerentePAProv["Número PA Carteira"].values[n].astype(int)}</td>'
                    col2 = f'<td style ="text-align: left;">{baseGerentePAProv["Nome Carteira"].values[n]}</td>'
                    col3 = f'<td style ="text-align: left;">{baseGerentePAProv["Número CPF/CNPJ"].values[n]}</td>'
                    col4 = f'<td style ="text-align: left;">{baseGerentePAProv["Nome Cliente"].values[n].upper()}</td>'
                    col5 = f'<td style ="text-align: center;">{baseGerentePAProv["Contrato"].values[n].upper()}</td>'
                    col6 = f'<td style ="text-align: center;">{baseGerentePAProv["Risco CRL"].values[n]}</td>'
                    col7 = f'<td style ="text-align: center;">{baseGerentePAProv["Nivel Risco COP ou Atual"].values[n]}</td>'
                    col8 = f'<td style ="text-align: center;">{baseGerentePAProv["Nível Risco Projetado"].values[n]}</td>'
                    col9 = f'<td style ="text-align: center;">{baseGerentePAProv["Dias em Atraso"].values[n].astype(int)}</td>'
                    col10 = f'<td style ="text-align: center;">{baseGerentePAProv["Atraso Projetado Final do Mês"].values[n].astype(int)}</td>'
                    col11 = f'<td style ="text-align: center;">{baseGerentePAProv["Motivo da Alteração"].values[n]}</td>'
                    col12 = f'<td style ="text-align: right;">{locale.currency(baseGerentePAProv["Variação de Provisão"].values[n], symbol=False, grouping=True)}</td>'
                    col13 = f'<td style ="text-align: right;">{locale.currency(baseGerentePAProv["Saldo Devedor"].values[n], symbol=False, grouping=True)}</td>'
                    corpo2 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}{col11}{col12}{col13}</tr>'
                valor_total = baseGerentePAProv["Variação de Provisão"].sum()

                corpo_email = f"{corpo1}</tbody></table>{corpo2}</tbody></table>"

                tabela_html = inicio_email_GerentePA + corpo_email + fim_email
                if(valor_total=='0,00'):
                    pass
                #Variáveis dos E-mails
                para = str(mailto) #"amandal.silva@sicoob.com.br"
                copia = "" #Cópiann
                copia_oculta = "" #Cópia Oculta
                assunto = 'GERENTE - CONTRATOS EM ATRASO DE ATÉ 45 DIAS ACIMA DE R$ 50 MIL DO SEU PA.'
                corpoHTML = tabela_html
                
                enviar_email(para,copia,copia_oculta,assunto,corpoHTML)    
                print("Envio Diario Gerente PA")  

    #E-mail Diretoria ↓↓
    if len(bs_diretoria.index) > 0:
        
        bs_base_dir = bs_diretoria
        valor_total = locale.currency(bs_base_dir['Valor Atualizado'].sum(), symbol=True, grouping=True)
        valor_operacao = locale.currency(bs_base_dir['Valor Operação'].sum(), symbol=True, grouping=True)
        num_linhas = len(bs_base_dir)
        if(num_linhas == 0 ):
            pass
        else:
            # Formata o corpo do e-mail
            corpo1 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Com o total de <span style="font-weight: 1000"><strong>{num_linhas}</strong></span> contratos, o valor total da inadimplência é de <span style="font-weight: 1000"><strong>{valor_total}</strong></span> e o valor total de contratos é de <span style="font-weight: 1000"><strong>{valor_operacao}</strong></span>.</p>
                <table style="border-color: #003641; border-collapse: collapse; width: 100%; table-layout: auto" border="1">
                    <tbody>
                        <tr style="text-align: center; background-color: #003641; color: #ffffff">
                            <th>AGÊNCIA</th>
                            <th>NOME CARTEIRA</th>
                            <th>CPF/CNPJ</th>
                            <th>COOPERADO</th>
                            <th>PRODUTO</th>
                            <th>CONTRATO</th>
                            <th>RISCO</th>
                            <th>DIAS EM ATRASO</th>
                            <th>VALOR OPERAÇÃO</th>
                            <th>VALOR INADIMPLENTE</th>
                        </tr>"""
            for n in range(len(bs_base_dir)):
                col1 = f'<td style ="text-align: left;">{bs_base_dir["PA"].values[n]}</td>'
                col2 = f'<td style ="text-align: left;">{bs_base_dir["Nome Carteira"].values[n]}</td>'
                col3 = f'<td style ="text-align: left;">{bs_base_dir[" CPF/CNPJ"].values[n]}</td>'
                col4 = f'<td style ="text-align: left;">{bs_base_dir["Nome"].values[n].upper()}</td>'
                col5 = f'<td style ="text-align: center;">{bs_base_dir["Carteira"].values[n].upper()}</td>'
                col6 = f'<td style ="text-align: center;">{bs_base_dir["Contrato"].values[n]}</td>'
                col7 = f'<td style ="text-align: center;">{bs_base_dir["Risco"].values[n]}</td>'
                col8 = f'<td style ="text-align: center;">{bs_base_dir["Qtd. Dias Atraso"].values[n]}</td>'
                col9 = f'<td style ="text-align: right;">{locale.currency(bs_base_dir["Valor Operação"].values[n], symbol=False, grouping=True)}</td>'
                col10 = f'<td style ="text-align: right;">{locale.currency(bs_base_dir["Valor Atualizado"].values[n], symbol=False, grouping=True)}</td>'
                corpo1 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}</tr>'
                valor_total = bs_base_dir['Valor Atualizado'].sum()


            valor_totalVariação = locale.currency(bs_diretoriaPROV["Variação de Provisão"].sum(), symbol=True, grouping=True)
            valor_totalSaldo = locale.currency(bs_diretoriaPROV['Saldo Devedor'].sum(), symbol=True, grouping=True)
            # Formata o corpo do e-mail PROV
            corpo2 = f"""<p style="color: #003641; text-align: center; font-size: 20px">Os 20 contratos com maior variação da provisão do dia.</p>
            <p style="color: #003641; text-align: center; font-size: 20px">A soma da variação é de <span style="font-weight: 1000"><strong>{valor_totalVariação}</strong></span> e o valor total do saldo devedor é <span style="font-weight: 1000"><strong>{valor_totalSaldo}</strong></span>.</p>
                <table style="border-color: #003641; border-collapse: collapse; width: 100%; table-layout: auto" border="1">
                    <tbody>
                        <tr style="text-align: center; background-color: #003641; color: #ffffff">
                            <th>AGÊNCIA</th>
                            <th>NOME CARTEIRA</th>
                            <th>CPF/CNPJ</th>
                            <th>COOPERADO</th>
                            <th>CONTRATO</th>
                            <th>RISCO</th>
                            <th>RISCO ATUAL</th>
                            <th>RISCO PROJETADO</th>
                            <th>DIAS EM ATRASO</th>
                            <th>ATRASO PROJETADO</th>
                            <th>MOTIVO DA ALTERAÇÃO</th>
                            <th>VARIAÇÃO DA PROVISÃO</th>
                            <th>SALDO DEVEDOR</th>                            
                        </tr>"""
            for n in range(len(bs_diretoriaPROV)):
                col1 = f'<td style ="text-align: left;">{bs_diretoriaPROV["Número PA Carteira"].values[n].astype(int)}</td>'
                col2 = f'<td style ="text-align: left;">{bs_diretoriaPROV["Nome Carteira"].values[n]}</td>'
                col3 = f'<td style ="text-align: left;">{bs_diretoriaPROV["Número CPF/CNPJ"].values[n]}</td>'
                col4 = f'<td style ="text-align: left;">{bs_diretoriaPROV["Nome Cliente"].values[n].upper()}</td>'
                col5 = f'<td style ="text-align: center;">{bs_diretoriaPROV["Contrato"].values[n].upper()}</td>'
                col6 = f'<td style ="text-align: center;">{bs_diretoriaPROV["Risco CRL"].values[n]}</td>'
                col7 = f'<td style ="text-align: center;">{bs_diretoriaPROV["Nivel Risco COP ou Atual"].values[n]}</td>'
                col8 = f'<td style ="text-align: center;">{bs_diretoriaPROV["Nível Risco Projetado"].values[n]}</td>'
                col9 = f'<td style ="text-align: center;">{bs_diretoriaPROV["Dias em Atraso"].values[n].astype(int)}</td>'
                col10 = f'<td style ="text-align: center;">{bs_diretoriaPROV["Atraso Projetado Final do Mês"].values[n].astype(int)}</td>'
                col11 = f'<td style ="text-align: center;">{bs_diretoriaPROV["Motivo da Alteração"].values[n]}</td>'
                col12 = f'<td style ="text-align: right;">{locale.currency(bs_diretoriaPROV["Variação de Provisão"].values[n], symbol=False, grouping=True)}</td>'
                col13 = f'<td style ="text-align: right;">{locale.currency(bs_diretoriaPROV["Saldo Devedor"].values[n], symbol=False, grouping=True)}</td>'
                
                
                
                corpo2 += f'<tr>{col1}{col2}{col3}{col4}{col5}{col6}{col7}{col8}{col9}{col10}{col11}{col12}{col13}</tr>'
                valor_total = bs_diretoriaPROV["Variação de Provisão"].sum()

            corpo_email = f"{corpo1}</tbody></table>{corpo2}</tbody></table>"

            tabela_html = inicio_email_Diretoria + corpo_email + fim_email
            
            #Variáveis dos E-mails
            para = "priscila.dasilva@sicoob.com.br;mario.schutz@sicoob.com.br;andrezza.ribeiro@sicoob.com.br"#"amandal.silva@sicoob.com.br"
            copia = "recuperauniro@sicoob.onmicrosoft.com"
            copia_oculta = "" 
            assunto = 'DIRETORIA -  PROV05 E CONTRATOS EM ATRASO ACIMA DE R$ 200 MIL'
            corpoHTML = tabela_html
            enviar_email(para,copia,copia_oculta,assunto,corpoHTML)
            print("Envio Diario DIRETORIA")   
    else:
        pass
    os.remove(path+"/filacobranca/"+arquivo)
    enviarsemanal = False
    envio.set("Finalizado")
    app.update_idletasks()
    print("finalizado")
   
# Programa que valida se o arquivo existe
def arquivo_existe(enviarsemanal):
    pasta = pd.DataFrame(os.listdir(path+"/filacobranca"), columns=['items'])
    arquivo = pasta[pasta['items'].str.contains('Filas')]['items'].values[0]
    if arquivo == "":
        envio.set("Não há arquivo")
        app.update_idletasks()
        pg.confirm(text= "O arquivo não se encontra na pasta, favor, baixá-lo e então iniciar.")

    else:
        envio.set("Em andamento")
        app.update_idletasks()
        print("Entrando no algoritimo para enviar...") 
        distribuicao(enviarsemanal)  
        envio.set("E-mails semanais enviados"if enviarsemanal else "E-mails diários enviados" )
        

# ----------------Inicializa o Aplicativo-------------------------------------------------------------------------------------------------------------------------------------------------------
# Criação da janela principal
app = Tk()
app.title("Automação E-Mails INAD e PROV - Business Intelligence")
app.geometry('500x300')
app.configure(background="#003641", highlightbackground="#00ae9d", padx=5, pady=5)
#Propriedades da janela:
largura_janela = 600
altura_janela = 300

# Obtenha a largura e altura da tela
largura_tela = app.winfo_screenwidth()
altura_tela = app.winfo_screenheight()
# coordenadas x e y para centralizar a janela
x = (largura_tela - largura_janela) // 2
y = (altura_tela - altura_janela) // 2
#dimensões e a posição da janela
app.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")
app.resizable(width=False, height=False)
app.configure(bg='#f0f3f7')


envio = StringVar()

titulo1 = Label(font=('Arial', '16', 'bold'), fg='#003641', text='Envio de PROV05 E INAD')
titulo1.pack()
# Botão e Label para e-mails diários
lb_diario = Label(app, text="Enviar os e-mails diários:", font=("Arial", 13,'bold'))
lb_diario.place(x=10, y=70)

btn_diario = Button(app, text="Clique aqui para executar", bd='10',font=("Arial", 10,'bold'), command=lambda: arquivo_existe(False))
btn_diario.place(x=355, y=64)

# Botão e Label para e-mails semanais
lb_semanal = Label(app, text="Enviar os e-mails semanais:",  font=("Arial", 13,'bold'))
lb_semanal.place(x=10, y=150)

btn_semanal = Button(app, text="Clique aqui para executar", bd='10',font=("Arial", 10,'bold'), command=lambda: arquivo_existe(True))
btn_semanal.place(x=355, y=150)

# Label para exibir status/envio
data_base = Label(app, textvariable=envio,font=('Arial', '10', 'bold'), fg='#003641')
data_base.pack()

app.mainloop()