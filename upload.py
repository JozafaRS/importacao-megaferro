import pandas as pd
from datetime import datetime
from dotenv import load_dotenv, find_dotenv
from os import environ
import sqlalchemy

load_dotenv(find_dotenv())

URL_DB = environ.get('URL_DB')

def validar_planilha(data_frame: pd.DataFrame) -> None:
    colunas_esperadas = ['Cod Parc', 'Unico', 'N Nota', 'Perfil', 'Tipo Negociacao',
       'Grupo Economico', 'Cód Tipo', 'Perfil.1', 'Data', 'Telefone',
       'Vendedor', 'Empresa', 'Tipo', 'Endereco', 'Numero', 'CGC_CPF',
       'Bairro', 'Cidade', 'Regiao', 'Razao', 'Cod Produto', 'Qtd Neg',
       'Desc Produto', 'Grupo Produto', 'LUCRO', 'PERC_LUCRO',
       'Grupo Pai de Prod', 'Peso Bruto', 'Valor', 'Retirada', 'Status',
       'Entrega', 'Lim de Credito Mensal', 'Limite de Credito',
       'Ultima Negociacao', 'INADIMPLENCIA', 'AD_WHATSAPP', 'ANO',
       'CODIGO_TIPO_VENDA', 'MES']
    
    if data_frame.empty:
        raise ValueError('Planilha Vazia')

    if list(data_frame.columns) != colunas_esperadas:
        raise TypeError('Colunas Incompatíveis')
    
    # if not pd.api.types.is_datetime64_any_dtype(data_frame['Data']):
    #     raise TypeError('A coluna Data não é do tipo datetime.')

    if not pd.api.types.is_datetime64_any_dtype(data_frame['Ultima Negociacao']):
        raise TypeError('A coluna Ultima Negociacao não é do tipo datetime.')
    
def formatar_data_frame(df_original: pd.DataFrame) -> pd.DataFrame:
    data_frame = df_original.copy()
    
    data_frame.columns = data_frame.columns.to_series().apply(lambda x: x.lower().replace(" ", "_").replace("ó", "o").replace(".", "_"))
    data_frame['data'] = pd.to_datetime(data_frame['data'], dayfirst=True)
    data_frame['data_ajustada'] = data_frame['data']
    data_frame['ficticio'] = 1
    
    return data_frame

def filtrar_novos_dados(data_frame: pd.DataFrame, tabela: str, campo: str = "unico"):
    conn = sqlalchemy.create_engine(URL_DB)
    query = f'SELECT {campo} FROM {tabela};'
    pedidos_registrados = pd.read_sql_query(query, conn)

    df_filtrado = data_frame[~data_frame[campo].isin(pedidos_registrados[campo])]
    
    return df_filtrado


def adicionar_registros(data_frame: pd.DataFrame, nome_tabela: str):
    conn = sqlalchemy.create_engine(URL_DB)  
    data_frame.to_sql(nome_tabela, conn, if_exists='append', index=False)