import pandas as pd
import os
import glob
import re
import datetime

# Definir caminho 
folder_path = 'src\\data\\raw'

# Ler Excel Cadastro produtos 
cd_vendas = glob.glob(os.path.join(folder_path, 'Base Vendas*'))
cd_prod = glob.glob(os.path.join(folder_path, 'Cadastro Produtos.xlsx'))
cd_cli = glob.glob(os.path.join(folder_path, 'Cadastro Clientes.xlsx'))

# Confirma se tem excel
if not cd_vendas:
    print("Não encontrados arquivos no formato")
else:

# Define um DataFrame final de vendas

    dfvd = []

# Ler dataframe de produtos
for excel_file1 in cd_prod:
    try:
        dfpr_temp = pd.read_excel(excel_file1)
        

        # dfvd.append(dfvd_temp)

    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file1} : {e}")

# Ler dataframe de clientes
for excel_file2 in cd_cli:
    try:
        dfcl_temp = pd.read_excel(excel_file2)

        # dfvd.append(dfvd_temp)

    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file2} : {e}")

# Ler dataframe de vendas
for excel_file in cd_vendas:
    try:
        dfvd_temp = pd.read_excel(excel_file)
        # Puxa informações da tabela de produtos para a de vendas
        dfvd_temp = pd.merge(dfvd_temp,dfpr_temp, on='SKU', how='inner')
        # Apaga coluna de Observação
        dfvd_temp = dfvd_temp.drop("Observação", axis=1)

        # Puxa informações dos clientes
        dfvd_temp = pd.merge(dfvd_temp, dfcl_temp, on='ID Cliente', how='inner')
        dfvd_temp['Genero'] = dfvd_temp['Genero'].replace('M','Masculino')
        dfvd_temp['Genero'] = dfvd_temp['Genero'].replace('F', 'Feminino')
        dfvd_temp['Data Nascimento'] = pd.to_datetime(dfvd_temp['Data Nascimento'], format='%m/%d/%Y')
        dfvd_temp['Estado Civil'] = dfvd_temp['Estado Civil'].replace('C' , 'Casado')
        dfvd_temp['Estado Civil'] = dfvd_temp['Estado Civil'].replace('S' , 'Solteiro')
        dfvd_temp['Cliente'] = dfvd_temp['Primeiro Nome'] + ' ' + dfvd_temp['Sobrenome']
        dfvd_temp = dfvd_temp.drop("Primeiro Nome", axis=1)
        dfvd_temp = dfvd_temp.drop("Sobrenome", axis=1)
        
        # Separa cor de produto
        dfvd_temp[['Nome Produto','Cor']] = dfvd_temp['Produto'].str.split(' - ',expand=True)
        dfvd_temp = dfvd_temp.drop("Produto", axis=1)
        
        # Cálculo faturamento por ítem
        dfvd_temp['Faturamento total'] = dfvd_temp['Preço Unitario'] * dfvd_temp['Qtd Vendida']
        dfvd_temp['Custo Total'] = dfvd_temp['Custo Unitario'] * dfvd_temp['Qtd Vendida']
        dfvd_temp['Margem %'] = dfvd_temp['Custo Total']/dfvd_temp['Faturamento total']
        dfvd_temp['Lucro total'] = dfvd_temp['Faturamento total'] - dfvd_temp['Custo Total']
       
       # Formatação R$ custo, preços unitários e faturamento
        
        dfvd_temp['Preço Unitario'] = dfvd_temp['Preço Unitario'].replace(".",",").replace(",",".")
        dfvd_temp['Custo Unitario'] = dfvd_temp['Custo Unitario'].replace(".",",").replace(",",".")
        dfvd_temp['Faturamento total'] = dfvd_temp['Faturamento total'].replace(".",",").replace(",",".")
        dfvd_temp['Custo Total'] = dfvd_temp['Custo Total'].replace(".",",").replace(",",".")
        dfvd_temp['Lucro total'] = dfvd_temp['Lucro total'].replace(".",",").replace(",",".")
        
        # Formatação Margem %
        def format(valor1):
            return "{:.2%}".format(valor1)
        
        dfvd_temp['Margem %'] = dfvd_temp['Margem %'].apply(format)
  

        dfvd_temp['Ordem de Compra'] = dfvd_temp['Ordem de Compra'].str.upper()
        dfvd_temp['SKU'] = dfvd_temp['SKU'].str.upper()
        dfvd_temp['Marca'] = dfvd_temp['Marca'].str.upper()
        dfvd_temp['Categoria'] = dfvd_temp['Categoria'].str.upper()
        dfvd_temp['Nome Produto'] = dfvd_temp['Nome Produto'].str.upper()
        dfvd_temp['Cor'] = dfvd_temp['Cor'].str.upper()
        dfvd_temp['Cliente'] = dfvd_temp['Cliente'].str.upper()
        dfvd_temp['Email'] = dfvd_temp['Email'].str.upper()
        dfvd_temp['Genero'] = dfvd_temp['Genero'].str.upper()
        dfvd_temp['Estado Civil'] = dfvd_temp['Estado Civil'].str.upper()
        dfvd_temp['Nivel Escolar'] = dfvd_temp['Nivel Escolar'].str.upper()

        # Ordena colunas
        dfvd_temp = dfvd_temp[['Data da Venda','Ordem de Compra','SKU','Marca','Categoria','Nome Produto','Cor','Custo Unitario','Preço Unitario','Qtd Vendida','Custo Total','Faturamento total','Lucro total','Margem %','ID Cliente','Cliente','Documento','Data Nascimento','Email','Genero','Estado Civil','Num Filhos','Nivel Escolar']]
    



        dfvd.append(dfvd_temp)

    except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file} : {e}")



if dfvd:

    #Garante que todas as tabelas estejam concatenadas
    result = pd.concat(dfvd, ignore_index=True)

    #Caminho de saída
    output_file = os.path.join('src', 'data', 'ready', 'bd_vendas.xlsx')

    #Configuração motor de escrita
    writer = pd.ExcelWriter(output_file, engine= 'xlsxwriter')

    #Leva os dados a serem escritos em um motor de excel configurado
    result.to_excel(writer, index=False)

    #Salva o arquivo de Excel
    writer._save()
else:
  print("Nenhum dado para ser salvo")


    

