# -*- coding: utf-8 -*-
"""
Created on Thu Jan 30 12:05:16 2025

@author: tiago.piccoli
"""

import pandas as pd
import os
import shutil
from openpyxl import load_workbook
import time
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

#funções------------------------------------------------------------------------
def mover_arquivos(mover_des, pasta_fabrica, pasta_obsoletos, pasta_novos): #desenhos existentes
    for arquivo in subst_des:
        # Procurar arquivos com o mesmo nome (considerar quaisquer extensões)
        arquivo_str = str(arquivo)
        for extensao in ['.dxf', '.PDF', '.x_t']:
            arquivo_fabrica = os.path.join(pasta_fabrica, arquivo_str + extensao)
            arquivo_novos = os.path.join(pasta_novos, arquivo_str + extensao)
            arquivo_obsoletos = os.path.join(pasta_obsoletos, arquivo_str + extensao)
            
            # Mover da pasta desenhos fabrica para pasta obsoletos
            if os.path.exists(arquivo_fabrica):
                shutil.move(arquivo_fabrica, arquivo_obsoletos)
            
            # Copiar da pasta novos desenhos para pasta desenhos fabrica
            if os.path.exists(arquivo_novos):
                shutil.copy2(arquivo_novos, arquivo_fabrica)
                print(arquivo," copiado para Desenhos Fabrica")
                os.remove(arquivo_novos)# Deletar da pasta novos desenhos
                
            else:
                pass

def novo_arquivos(subst_des, pasta_fabrica, pasta_novos): #desenhos novos
    for arquivo in mover_des:
        # Procurar arquivos com o mesmo nome (considerar quaisquer extensões)
        arquivo_str = str(arquivo)
        for extensao in ['.dxf', '.PDF', '.x_t']:
            arquivo_fabrica = os.path.join(pasta_fabrica, arquivo_str + extensao)
            arquivo_novos = os.path.join(pasta_novos, arquivo_str + extensao)
            
            # Copiar da pasta novos desenhos para pasta desenhos fabrica
            if os.path.exists(arquivo_novos):
                shutil.copy2(arquivo_novos, arquivo_fabrica)
                print(arquivo, 'Movido para pasta Desenhos Fábrica')
                os.remove(arquivo_novos)
            else:
                pass
# Função para salvar lista em arquivo .txt
def emails_lista(lista, nome_arquivo):
    with open(f'{nome_arquivo}.txt', 'w') as f:
        for item in lista:
            f.write(f"{item}\n")

def listar_des_nv(pasta_novos):
    arquivos_novos = []
    for arquivo in os.listdir(pasta_novos):
        if arquivo.endswith(('.pdf','.PDF')):
            arquivos_novos.append(arquivo)
    return arquivos_novos

def des_nv_ant(arquivos_pdf, pasta_fabrica): #verifica se o desenho novo já existe na pasta.
    arquivos_existentes = []
    for arquivo in arquivos_novos:
        caminho_arquivo = os.path.join(pasta_fabrica, arquivo)
        if os.path.exists(caminho_arquivo):
            arquivos_existentes.append(arquivo)
            arquivos_novos.remove(arquivo)
        else:
            pass
        
    return arquivos_existentes
#---------------------------------------------------

print("PROGRAMA FLUXO PARA LIBERAÇÃO DE DESENHOS, Versão 0")
input("Tecle 'Enter' para iniciar o programa")

pasta_fabrica = 'P:\\Útil\\Desenhos Fábrica'
pasta_obsoletos = 'P:\\Útil\\Desenhos Tecnicos\\OBSOLETOS'
pasta_novos = 'P:\\Útil\\Desenhos Tecnicos\\Novos Desenhos'

# pasta_fabrica = r'C:\Users\tiago.piccoli\Desktop\Automacao_desenhos\des_fabrica'
# pasta_obsoletos = r'C:\Users\tiago.piccoli\Desktop\Automacao_desenhos\des_obsletos'
# pasta_novos = r'C:\Users\tiago.piccoli\Desktop\Automacao_desenhos\des_nv'

dia=date.today()
data_hj=dia.strftime('%d/%m/%Y')
arquivos_novos = listar_des_nv(pasta_novos) #gera lista com os arquivos novos
arquivos_existentes = des_nv_ant(arquivos_novos, pasta_fabrica) #verifica se é desenho novo

# Para os arquivos revisados (REVISADO)
novos_dados_rev = pd.DataFrame([{
    'CODIGO': str(arquivo)[:-4], 
    'Data Liberação': data_hj,
    'PROCESSOS':'',    
    'VERSAO': 'REVISADO',
    'FLUXO': 0,
    'OBS':''
} for arquivo in arquivos_existentes])

# Para os arquivos novos (DESENHO NOVO)
novos_dados_nv = pd.DataFrame([{
    'CODIGO': str(arquivo)[:-4], 
    'Data Liberação': data_hj,
    'PROCESSOS':'',    
    'VERSAO': 'DESENHO NOVO',
    'FLUXO': 0,
    'OBS':''
} for arquivo in arquivos_novos])

# Concatenar os novos dados com o DataFrame existente
df = pd.concat([novos_dados_nv, novos_dados_rev], ignore_index=True)

# Salvar o DataFrame atualizado de volta no arquivo Excel
df.to_excel('FLUXO DESENHOS.xlsx', index=False, engine='openpyxl')

# Carregar o workbook e a sheet
wb = load_workbook('FLUXO DESENHOS.xlsx')
ws = wb.active

# Definir as opções da lista suspensa
opcoes_lista = [
    "LASER", "LASER, DOBRA", "USINAGEM", "SERRA", "MONTAGEM", "PRE MONTAGEM",
    "TERCEIROS", "ALIMENTADORES", "SOLDA", "FONTES", "COMPRADO", "CENTRO",
    "OUTROS", "SERRA, CENTRO", "ROSQUEAMENTO", "GRAVADOR", "SERRA, TORNO",
    "SERRA, TORNO, CENTRO", "DOBRA", "SEM ROTEIRO"
]

# Criar a validação de dados (drop-down)
dv = DataValidation(
    type="list",
    formula1=f'"{",".join(opcoes_lista)}"',
    allow_blank=True,
    showDropDown=True
)

# Adicionar a validação de dados à coluna "PROCESSOS"
coluna_processos = ws['C']
for cell in coluna_processos[1:]:
    dv.add(cell)

# Adicionar a validação à worksheet
ws.add_data_validation(dv)

# Salvar o workbook atualizado
wb.save('FLUXO DESENHOS.xlsx')

print('Relatório para preenchimento de PROCESSO/OPERAÇÃO')

input('Por favor, abra a planilha, preencha a coluna "PROCESSOS" a partir do roteiro do item, salve, feche e tecle Enter quando tiver concluído.')

df=pd.read_excel('FLUXO DESENHOS.XLSX')

cod_qualidade = []
cod_compras = []
cod_laser = []
ver_impressao_cq = []
atencao=[]

mover_des=[] #lista para função de operação de desenhos novos
subst_des=[] #lista para função de operação de desenhos antigos

for index, row in df.iterrows(): #Algoritmo
    fluxo = row['FLUXO']
    codigo = row['CODIGO']
    versao = row['VERSAO']
    processos = row['PROCESSOS']
    obs=row['OBS']
    
    #verificar se é desenho novo, vendo se há existência do arquivo na pasta.
    
    if versao=="DESENHO NOVO":
        mover_des.append(str(codigo))
        
        if processos in ('COMPRADO','TERCEIROS'):
            cod_compras.append(str(codigo))
            df.at[index, 'FLUXO'] = 1
            df.at[index, 'OBS']="Desenho Movido"
            df.at[index, 'Data Liberação']=data_hj
            
        if processos in ('MONTAGEM', 'PRE MONTAGEM', 'ALIMENTADORES', 'SERRA','FONTES','GRAVADOR'):
            df.at[index, 'FLUXO'] = 1
            df.at[index, 'OBS']="Desenho Movido"
            df.at[index, 'Data Liberação']=data_hj
            
        if processos in ('LASER','LASER, DOBRA'):
            cod_laser.append(str(codigo))
            ver_impressao_cq.append(str(codigo))
            df.at[index, 'FLUXO'] = 1
            df.at[index, 'OBS']="Desenho Movido, Verificar CQ e Imprimir"
            df.at[index, 'Data Liberação']=data_hj
            
        if processos in ('USINAGEM','SOLDA','SERRA,CENTRO','SERRA, TORNO, CENTRO','SERRA, TORNO'):
            ver_impressao_cq.append(str(codigo))
            df.at[index, 'FLUXO'] = 1      
            df.at[index, 'OBS']="Desenho Movido, Verificar CQ e Imprimir"
            df.at[index, 'Data Liberação']=data_hj
        else:
            atencao.append(codigo)
            mover_des.remove(codigo)
            
    if versao=="REVISADO":
        subst_des.append(str(codigo))
        
        if processos in ("LASER","LASER, DOBRA"):
            cod_laser.append(codigo)
            ver_impressao_cq.append(codigo)
            df.at[index, 'FLUXO'] = 1
            df.at[index, 'OBS']="Desenho Substituído, Verificar/Atualizar CQ caso haja e Imprimir"
            df.at[index, 'Data Liberação']=data_hj
            
        if processos in ("TERCEIROS","COMPRADO"):
            cod_qualidade.append(codigo)
            df.at[index, 'FLUXO'] = 1
            df.at[index, 'OBS']="Desenho Substituído, E-mail Suprimentos e Qualidade"
            df.at[index, 'Data Liberação']=data_hj
            
        if processos in ('USINAGEM','SOLDA'):
            ver_impressao_cq.append(codigo)
            df.at[index, 'FLUXO'] = 1
            df.at[index, 'OBS']="Desenho Substituído, Verificar/Atualizar CQ caso haja e Imprimir"
            df.at[index, 'Data Liberação']=data_hj
            
        if processos in ('MONTAGEM', 'PRE MONTAGEM', 'ALIMENTADORES', 'SERRA','FONTES','GRAVADOR'):
            df.at[index, 'FLUXO'] = 1
            df.at[index, 'OBS']="Desenho Substituído"
            df.at[index, 'Data Liberação']=data_hj

        else:
            atencao.append(codigo)
            subst_des.remove(codigo)
    else:
        atencao.append(codigo)

print('------------------------------------------------')
novo_arquivos(mover_des, pasta_fabrica, pasta_novos)
mover_arquivos(subst_des, pasta_fabrica, pasta_obsoletos, pasta_novos)

df.to_excel('FLUXO DESENHOS.xlsx',index=False, engine='openpyxl')
print('--------------RESUMO DAS OPERAÇÕES--------------')
print(df)

emails_lista(cod_compras, 'cod_compras')
emails_lista(cod_laser, 'cod_laser')
emails_lista(cod_qualidade, 'cod_qualidade')
emails_lista(ver_impressao_cq, 'ver_impressao_cq')
emails_lista(atencao, 'VERIFICAR')

input()