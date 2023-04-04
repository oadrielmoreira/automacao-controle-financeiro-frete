import pandas as pd
from datetime import datetime
import openpyxl

#Importação das Bases
nf = pd.read_excel('## CAMINHO DO ARQUIVO ##')
vpeso = pd.read_excel('## CAMINHO DO ARQUIVO ##')
loc = pd.read_excel('## CAMINHO DO ARQUIVO ##')

#Tratamento das Variáveis
nf['DATA'] = pd.to_datetime(nf['DATA'])
nf['DATA'] = nf['DATA'].dt.strftime('%d/%m/%Y')
nf['SERIE'] = nf['SERIE'].astype(str)
nf['PEDIDOV'] = nf['PEDIDOV'].astype(str)


#Dividindo duas planilhas em Série 001 e 003
nf1 = nf[nf['SERIE'] == '1']
nf3 = nf[nf['SERIE'] == '3']

#Somando os pesoss e trazendo para a nf
vpeso['COD_PEDIDOV'] = vpeso['COD_PEDIDOV'].astype(str)
vpeso['PESO_KG'] = vpeso['PESO_KG'].astype('float64')
pesos_por_pedido = vpeso.groupby('COD_PEDIDOV')['PESO_KG'].sum().reset_index()
peso = pd.DataFrame({
    'PEDIDOV': pesos_por_pedido['COD_PEDIDOV'],
    'Peso_total': pesos_por_pedido['PESO_KG']
})
nf1 = pd.merge(nf1, peso, on='PEDIDOV')
nf3 = pd.merge(nf3, peso, on='PEDIDOV')


#Definindo quais serão as substituições do código das transportadoras
substituicoes = {
    "00508 - Transportadora": "Transportadora",
    "00508 - Transportadora": "Transportadora",
    "00508 - Transportadora": "Transportadora",
    "00508 - Transportadora": "Transportadora",
    "00508 - Transportadora": "Transportadora",
    "00508 - Transportadora": "Transportadora",
    "00508 - Transportadora": "Transportadora",
    "00508 - Transportadora": "Transportadora",
}

#fazendo agrupamento
agrupado = vpeso.groupby('COD_PEDIDOV')['DESC_TRANSPORTADORA'].first().reset_index()
agrupado.rename(columns={'DESC_TRANSPORTADORA_PEDIDO': 'TRANSPORTADORA'}, inplace=True)

agrupado.to_excel('agrp.xlsx', index = False)


#fazendo substituições
workbook = openpyxl.load_workbook('agrp.xlsx')

sheet = workbook['Sheet1']

for index, row in agrupado.iterrows():
    # obtém o valor da célula atual
    texto_original = row['DESC_TRANSPORTADORA']
    
    # verifica se o texto precisa ser substituído
    for chave, valor in substituicoes.items():
        if chave in texto_original:
            # realiza a substituição
            texto_substituido = texto_original.replace(chave, valor)
            agrupado.at[index, 'DESC_TRANSPORTADORA'] = texto_substituido
            # interrompe o loop para não realizar outras substituições desnecessárias
            break
            
            
            
workbook = openpyxl.load_workbook('## CAMINHO DO ARQUIVO ##')

sheet = workbook['Sheet1']

for index, row in vpeso.iterrows():
    # obtém o valor da célula atual
    texto_original = row['DESC_TRANSPORTADORA']
    
    # verifica se o texto precisa ser substituído
    for chave, valor in substituicoes.items():
        if chave in texto_original:
            # realiza a substituição
            texto_substituido = texto_original.replace(chave, valor)
            vpeso.at[index, 'DESC_TRANSPORTADORA'] = texto_substituido
            # interrompe o loop para não realizar outras substituições desnecessárias
            break
vpeso.to_excel('## CAMINHO DO ARQUIVO ##', index = False)


#Exportando planilha agrp com transportadoras
agrupado['COD_PEDIDOV'] = agrupado['COD_PEDIDOV'].astype('object')
agrupado = agrupado.rename(columns={'COD_PEDIDOV': 'PEDIDOV'})


vpeso['PEDIDOV'] = vpeso['PEDIDOV'].astype('object')
agrupado['PEDIDOV'] = agrupado['PEDIDOV'].astype('object')

agrupado.to_excel('agrp.xlsx', index = False)

#trazendo a transportadora para o NF
nf1 = pd.merge(nf1, agrupado[['PEDIDOV', 'DESC_TRANSPORTADORA']], left_on='PEDIDOV', right_on='PEDIDOV', how='outer')
nf3 = pd.merge(nf3, agrupado[['PEDIDOV', 'DESC_TRANSPORTADORA']], left_on='PEDIDOV', right_on='PEDIDOV', how='outer')


#Fazendo a contagem dos pedidos e definindo quantos valores
counts = vpeso['COD_PEDIDOV'].value_counts()
vqtd = pd.DataFrame({'PEDIDOV': counts.index, 'count': counts.values})
vqtd['count'] = vqtd['count'].astype(int)
vqtd.to_excel('qtd.xlsx', index = False)

nf1 = pd.merge(nf1, vqtd[['PEDIDOV', 'count']], on='PEDIDOV', how='left')
nf3 = pd.merge(nf3, vqtd[['PEDIDOV', 'count']], on='PEDIDOV', how='left')

agrupadoloc = loc.groupby('Nfs')['Cidade nf','Estado nf'].first().reset_index()

nf['NOTA'] = nf['NOTA'].astype(str)
agrupadoloc['Nfs'] = agrupadoloc['Nfs'].astype(str)

agrupadoloc['Nfs'] = agrupadoloc['Nfs'].astype('float64')

nf1 = pd.merge(nf1, agrupadoloc[['Nfs', 'Cidade nf']], left_on='NOTA', right_on='Nfs', how='left')
nf1 = pd.merge(nf1, agrupadoloc[['Nfs', 'Estado nf']], left_on='NOTA', right_on='Nfs', how='left')

nf3 = pd.merge(nf3, agrupadoloc[['Nfs', 'Cidade nf']], left_on='NOTA', right_on='Nfs', how='left')
nf3 = pd.merge(nf3, agrupadoloc[['Nfs', 'Estado nf']], left_on='NOTA', right_on='Nfs', how='left')


plan1 = pd.DataFrame({
    'Co+A2:Q2ntrole' :'', #preenchido
    'Status': nf1['STATUS'],
    'Natureza' : nf1['DESC_EVENTO'],
    'Nf' : nf1['NOTA'],
    #'Pedido' : nf['PEDIDOV'],
    'Valor': nf1['VALOR'],
    'Cliente': nf1['NOME_DESTINATARIO'],
    'UF': nf1['Estado nf'],
    'Cidade': nf1['Cidade nf'],
    'Emissão': nf1['DATA'],
    'Entrega': '', #preenchido
    'Transportadora': nf1['DESC_TRANSPORTADORA'],
    'Cx': nf1['count'],
    'Kg': nf1['Peso_total'],
    'CT': '', #preenchido
    'Cif': '', #preenchido
    '%': '', #preenchido
    'Fob Cobrança': '', #preenchido
    'Fob orçamento': '', #preenchido
    'Valor Total': '', #preenchido
    '%': '', #preenchido
})

plan2 = pd.DataFrame({
    'Co+A2:Q2ntrole' :'', #preenchido
    'Status': nf3['STATUS'],
    'Natureza' : nf3['DESC_EVENTO'],
    'Nf' : nf3['NOTA'],
    #'Pedido' : nf['PEDIDOV'],
    'Valor': nf3['VALOR'],
    'Cliente': nf3['NOME_DESTINATARIO'],
    'UF': nf3['Estado nf'],
    'Cidade': nf3['Cidade nf'],
    'Emissão': nf3['DATA'],
    'Entrega': '', #preenchido
    'Transportadora': nf3['DESC_TRANSPORTADORA'],
    'Cx': nf3['count'],
    'Kg': nf3['Peso_total'],
    'CT': '', #preenchido
    'Cif': '', #preenchido
    '%': '', #preenchido
    'Fob Cobrança': '', #preenchido
    'Fob orçamento': '', #preenchido
    'Valor Total': '', #preenchido
    '%': '', #preenchido
})


with pd.ExcelWriter('Planilha Financeira.xlsx') as writer:
    plan1.to_excel(writer, sheet_name='Série 001')
    plan2.to_excel(writer, sheet_name='Série 003')
