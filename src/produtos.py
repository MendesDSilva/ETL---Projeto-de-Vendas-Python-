import pandas as pd
import os
import glob

# Definir caminho 
folder_path = 'src\\data\\raw'

# Ler Excel Cadastro produtos 
cd_produto = glob.glob(os.path.join(folder_path, 'Cadastro Produtos.xlsx'))

# Confirma se tem excel
if not cd_produto:
    print("Não encontrados arquivos no formato")
else:

# Define um DataFrame
    dfpd = []

# Vai tentar ler os arquivos 

for excel_file in cd_produto:
    try: 
        dfpd_temp = pd.read_excel(excel_file)
        dfpd_temp = dfpd_temp.drop("Observação", axis=1)
        
        # Formatação R$ custo e preços unitários
        def format(valor):
            return "{:.2f}".format(valor)
        
        dfpd_temp['Preço Unitario'] = dfpd_temp['Preço Unitario'].apply(format)
        dfpd_temp['Custo Unitario'] = dfpd_temp['Custo Unitario'].apply(format)
        
        dfpd.append(dfpd_temp)
    
    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file} : {e}")

if dfpd:

    #Garante que todas as tabelas estejam concatenadas
    result = pd.concat(dfpd, ignore_index=True)

    #Caminho de saída
    output_file = os.path.join('src', 'data', 'ready', 'bd_produtos.xlsx')

    #Configuração motor de escrita
    writer = pd.ExcelWriter(output_file, engine= 'xlsxwriter')

    #Leva os dados a serem escritos em um motor de excel configurado
    result.to_excel(writer, index=False)

    #Salva o arquivo de Excel
    writer._save()
else:
  print("Nenhum dado para ser salvo")

