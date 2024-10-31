import pandas as pd
import os 
import glob
import re
import datetime

local_loc = 'src\\data\\raw'

# Ler cadastro Clientes
cd_loc = glob.glob(os.path.join(local_loc, 'Cadastro Localidades.xlsx'))

# Confirma se tem excel
if not cd_loc:
    print("Não encontrados arquivos no formato")
else:

# Define um DataFrame
    dflc = []

for excel_file in cd_loc:
    try: 
        dflc_temp = pd.read_excel(excel_file) 
        
        dflc.append(dflc_temp)

    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file} : {e}")

if dflc:

    #Garante que todas as tabelas estejam concatenadas
    result = pd.concat(dflc, ignore_index=True)

    #Caminho de saída
    output_file = os.path.join('src', 'data', 'ready', 'bd_localidades.xlsx')

    #Configuração motor de escrita
    writer = pd.ExcelWriter(output_file, engine= 'xlsxwriter')

    #Leva os dados a serem escritos em um motor de excel configurado
    result.to_excel(writer, index=False)

    #Salva o arquivo de Excel
    writer._save()
else:
  print("Nenhum dado para ser salvo")

