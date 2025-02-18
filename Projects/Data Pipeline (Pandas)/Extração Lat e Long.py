#%%
from geopy.geocoders import Nominatim
import pandas as pd
import os
import pyodbc
import warnings
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL
warnings.filterwarnings('ignore')
warnings.simplefilter(action='ignore', category=FutureWarning)
import os
import datetime


# %%
proxy_url = "http://proxy.mercantil.com.br:3128"
os.environ['HTTP_PROXY'] = proxy_url
os.environ['HTTPS_PROXY'] = proxy_url

#%%
conn_mbcorp = pyodbc.connect(
 "Driver={SQL Server Native Client 11.0};"
  "Server=SQLMBCORP;"
  "Database=MBCORP;"
  "Trusted_Connection=yes;")

query_enderecos = """
    SELECT  distinct ISNULL(A.NUM_DND_PA,A.NUM_DND) NUM_DND
        ,E.NOM_UND_ADN AS NOME_DEPENDENCIA
        ,I.SGL_UF      AS UF
        ,E.COD_CEP
        ,I.NOM_LOC AS MUNICIPIO
        ,E.DES_LGD
        ,REPLACE(trim(E.DES_LGD),',','') + ', ' + trim(I.NOM_LOC) + ', ' + trim(I.SGL_UF) + ', Brazil' AS GEOLOC
    FROM dbo.TML_AGENCIA AS A
    INNER JOIN dbo.AGENCIA AS C
    ON A.IDT_EMP = C.IDT_EMP AND ISNULL
    (A.NUM_DND_PA, A.NUM_DND) = C.NUM_DND
    INNER JOIN dbo.DEPENDENCIA AS D
    ON C.IDT_EMP = D.IDT_EMP AND C.NUM_DND = D.NUM_DND
    INNER JOIN dbo.UNIDADE_ORGANIZ AS E
    ON D.NUM_SEQ_UND_ADN = E.NUM_SEQ_UND_ADN
    INNER JOIN dbo.LOCALIDADE AS I
    ON E.IDT_LOC = I.IDT_LOC
    WHERE A.IDC_TML_ATV = 'S'
"""
df = pd.read_sql_query(query_enderecos, conn_mbcorp)


# %%
df['CONFIANCA'] = 1
df['END_REF'] = None
df['LATITUDE'] = None
df['LONGITUDE'] = None
df['LAT_LONG'] = None

geolocator = Nominatim(user_agent="geoapiExercises")

for x in df['GEOLOC'].values:
    linha = df[df['GEOLOC'] == x].index[0]
    print(linha)

    try:
        local = geolocator.geocode(f"{x}")
    except:
        continue

    if local == None:
        df.at[linha, 'CONFIANCA'] = 0
    else:
        df.at[linha, 'LATITUDE'] = local.latitude
        df.at[linha, 'LONGITUDE'] = local.longitude
        df.at[linha, 'LAT_LONG'] = f'{local.latitude}, {local.longitude}'
        df.at[linha, 'END_REF'] = local


#%%
nao_encontrados = df[df['CONFIANCA'] == 0]

geolocator = Nominatim(user_agent="myGeocoder")

for x in nao_encontrados['COD_CEP'].values:
    linha = df[df['COD_CEP'] == x].index[0]
    print(linha)

    try:
        local = geolocator.geocode(f"{x}, Brasil")
    except:
        continue

    if local == None:
        df.at[linha, 'CONFIANCA'] = 0
    else:
        df.at[linha, 'CONFIANCA'] = 1
        df.at[linha, 'LATITUDE'] = local.latitude
        df.at[linha, 'LONGITUDE'] = local.longitude
        df.at[linha, 'LAT_LONG'] = f'{local.latitude}, {local.longitude}'
        df.at[linha, 'END_REF'] = local


#%%
df['TESTE'] = [x.upper().replace('LJ', ' && ').replace('QD', '&&').replace('QDR', '&&').replace('LOJA', ' && ') for x in df['GEOLOC']]
result = []
for x in df['TESTE'].values:
    partes = x.split('&&')

    if len(partes) >= 2:
        part1 = partes[0]
        part2 = partes[1].split(',')[1:]
        part2 = ','.join(part2)
        previa = [part1, part2]
        previa = ','.join(previa)
        previa = previa.replace(' ,', ',').replace('.', ' ').replace(' s/n', '').replace(' -', '').replace('PRACA', 'PRAÇA').replace('AV', 'AVENIDA')
        result.append(previa)

    else:
        result.append(x.replace('.', ' ').replace(' S/N', '').replace(' -', '').replace('PRACA', 'PRAÇA'))
df['TESTE'] = result


# %%
nao_encontrados = df[df['CONFIANCA'] == 0]

for x in nao_encontrados['TESTE'].values:
    linha = df[df['TESTE'] == x].index[0]
    print(linha)

    try:
        local = geolocator.geocode(f"{x}")
    except Exception as erro:
        print(erro, '\n', x)

    if local == None:
        df.at[linha, 'CONFIANCA'] = 0
    else:
        df.at[linha, 'CONFIANCA'] = 1
        df.at[linha, 'LATITUDE'] = local.latitude
        df.at[linha, 'LONGITUDE'] = local.longitude
        df.at[linha, 'LAT_LONG'] = f'{local.latitude}, {local.longitude}'
        df.at[linha, 'END_REF'] = local


#%%
# ------------- CRIANDO CONEXÃO SQLALCHEMY -------------
connection_string = (
  r"Driver=SQL Server Native Client 11.0;"
  r"Server=SQLGDNP;" #
  r"Database=GNU;"
  r"Trusted_Connection=yes;"
)
connection_url = URL.create(
  "mssql+pyodbc", 
  query={"odbc_connect": connection_string}
)
engine = create_engine(connection_url, fast_executemany=True, connect_args={'connect_timeout': 10}, echo=False)
conn_alchemy_gnu = engine.connect()

#%%
'''df.drop(columns=['TESTE', 'CONFIANCA'], inplace=True)
with conn_alchemy_gnu as conn:
    with conn.begin() as beg:
        df.to_sql(name="TC_PROGRAMACAO_ENVIADA_K7_HST"
                , con=conn
                , if_exists='append'
                , index=False
                , chunksize=500)
        beg.commit()'''

#%%
df.drop(columns=['TESTE', 'CONFIANCA'], inplace=True)

#%%
arquivo = pd.ExcelWriter(path=r'K:\GSAS\09 - Coordenacao Gestao Numerario\09 Prototipos SSIS\B042786\__PYTHON__\LATITUDE E LONGITUDE PONTOS\ LATITUDE e LONGITUDE PONTOS.xlsx',engine='xlsxwriter')
df.to_excel(arquivo, sheet_name="Detalhes", index=False)
arquivo.save()
arquivo.close()

# %%
