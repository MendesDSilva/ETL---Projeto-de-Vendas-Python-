import pandas as pd
import os 
import glob
import re
import datetime

local_devol = 'src\\data\\raw'

# Ler cadastro Clientes
cd_devol = glob.glob(os.path.join(local_devol, 'Base Devoluções.xlsx'))

# Confirma se tem excel
if not cd_devol:
    print("Não encontrados arquivos no formato")
else:

 #Define um DataFrame
    dfdv = []

for excel_file in cd_devol:
    try: 
        dfdv_temp = pd.read_excel(excel_file) 
        
        dfdv.append(dfdv_temp)

    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file} : {e}")

if dfdv:

    #Garante que todas as tabelas estejam concatenadas
    result = pd.concat(dfdv, ignore_index=True)

    #Caminho de saída
    output_file = os.path.join('src', 'data', 'ready', 'bd_devolucoes.xlsx')

    #Configuração motor de escrita
    writer = pd.ExcelWriter(output_file, engine= 'xlsxwriter')

    #Leva os dados a serem escritos em um motor de excel configurado
    result.to_excel(writer, index=False)

    #Salva o arquivo de Excel
    writer._save()
else:
  print("Nenhum dado para ser salvo")

