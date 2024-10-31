import pandas as pd
import os
import glob

# Definir caminho 
local_loja = 'src\\data\\raw'

# Ler Excel Cadastro produtos 
cd_loja = glob.glob(os.path.join(local_loja, 'Cadastro Lojas.xlsx'))

# Confirma se tem excel
if not cd_loja:
    print("Não encontrados arquivos no formato")
else:

# Define um DataFrame
    dflj = []


for excel_file in cd_loja:
    try: 
        dflj_temp = pd.read_excel(excel_file) 

        dflj_temp[['Sobrenome','Nome']] = dflj_temp['Gerente Loja'].str.split(', ',expand=True)
        dflj_temp['Gerente'] = dflj_temp['Nome'] + " " + dflj_temp['Sobrenome']
        dflj_temp = dflj_temp[['ID Loja', 'Nome da Loja', 'Quantidade Colaboradores', 'Tipo','id Localidade', 'Gerente']]
        
        dflj.append(dflj_temp)  

    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file} : {e}")

if dflj:

    #Garante que todas as tabelas estejam concatenadas
    result = pd.concat(dflj, ignore_index=True)

    #Caminho de saída
    output_file = os.path.join('src', 'data', 'ready', 'bd_lojas.xlsx')

    #Configuração motor de escrita
    writer = pd.ExcelWriter(output_file, engine= 'xlsxwriter')

    #Leva os dados a serem escritos em um motor de excel configurado
    result.to_excel(writer, index=False)

    #Salva o arquivo de Excel
    writer._save()
else:
  print("Nenhum dado para ser salvo")
