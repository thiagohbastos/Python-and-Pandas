#%%
import pandas as pd

def adaptando_modelo(modelo_pd: str) -> pd.DataFrame:
    '''
    modelo_pd: é a string com os dados do DataFrame informados no LedCode
    
    Retorna o dataframe equivalente à string separada por '|'
    '''

    #Separando a tabela por linha, removendo linhas vazias
    tabela = modelo_pd.split('\n')
    tabela = [x for x in tabela if x != '']

    if '-' in tabela[1]:
        tabela.pop(1)

    #Definindo o cabeçalho
    cabecalho = [x.strip() for x in tabela[0].split('|') if x.strip() != '']

    # Manipulando tuplas da tabela
    itens_list = tabela
        #removendo cabeçalho
    itens_list.pop(0)
        #criando cada linha após split, removendo linhas vazias
    itens_list = [x.split('|') for x in itens_list]
    itens_list = [[y.strip() for y in x if y != ''] for x in itens_list]

    #Criando o dicionário base para o DataFrame
    dicionario = {cabecalho[i]: [x[i] for x in itens_list] 
                  for i in range(len(cabecalho))}

    df = pd.DataFrame(dicionario)

    for coluna in df:
        try:
            df[coluna] = df[coluna].astype(float)
        except:
            pass

    return df
    
