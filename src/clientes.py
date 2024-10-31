import pandas as pd
import os 
import glob
import re
import datetime

local_cadastrocc = 'src\\data\\raw'

# Ler cadastro Clientes
cd_clientes = glob.glob(os.path.join(local_cadastrocc, 'Cadastro Clientes.xlsx'))


# Confirma se tem excel
if not cd_clientes:
    print("Não encontrados arquivos no formato")
else:

 #Define um DataFrame
    dfcc = []

# Vai tentar ler os arquivos 
for excel_file in cd_clientes:
    try: 
        dfcc_temp = pd.read_excel(excel_file)

       # Define gênero
        dfcc_temp['Genero'] = dfcc_temp['Genero'].replace('M','Masculino')
        dfcc_temp['Genero'] = dfcc_temp['Genero'].replace('F', 'Feminino')

       # Converte formato data, data de nascimento
        
        dfcc_temp['Data Nascimento'] = pd.to_datetime(dfcc_temp['Data Nascimento'], format='%m/%d/%Y')

        # Define estado civil

        dfcc_temp['Estado Civil'] = dfcc_temp['Estado Civil'].replace('C' , 'Casado')
        dfcc_temp['Estado Civil'] = dfcc_temp['Estado Civil'].replace('S' , 'Solteiro')
       
        # Define formato CPF

        
        
        dfcc.append(dfcc_temp)

    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file} : {e}")

if dfcc:

    #Garante que todas as tabelas estejam concatenadas
    result = pd.concat(dfcc, ignore_index=True)

    #Caminho de saída
    output_file = os.path.join('src', 'data', 'ready', 'bd_clientes.xlsx')

    #Configuração motor de escrita
    writer = pd.ExcelWriter(output_file, engine= 'xlsxwriter')

    #Leva os dados a serem escritos em um motor de excel configurado
    result.to_excel(writer, index=False)

    #Salva o arquivo de Excel
    writer._save()
else:
  print("Nenhum dado para ser salvo")

