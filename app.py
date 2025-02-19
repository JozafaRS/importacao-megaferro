import streamlit as st
import upload as up
import pandas as pd
import os

def processar_arquivo(arquivo):
    st.markdown(f'### -> {arquivo.name}')
    progresso = st.progress(0, 'Lendo arquivo...')

    try: 
        data_frame = pd.read_excel(arquivo, header=2, skipfooter=1)
    except:
        progresso.empty()
        st.error('Erro ao abrir arquivo')
        return
    
    progresso.progress(1/5, 'Validando planilha...')

    try:
        up.validar_planilha(data_frame)
    except Exception as e:
        progresso.empty()
        st.error(f'Planilha inválida. Erro: {e}')
        return

    progresso.progress(2/5, 'Formatando Planilha...')
    try: 
        df_formatado = up.formatar_data_frame(data_frame)
    except Exception as e:
        progresso.empty()
        st.error(f'Erro ao formatar planilha: {e}')
        return

    progresso.progress(3/5, 'Filtrando Registros...')

    try:
        novos_dados = up.filtrar_novos_dados(df_formatado, 'base_de_vendas_megaferro')
    except Exception as e:
        progresso.empty()
        st.error(f'Erro ao se conectar ao banco de dados. Error: {e}')
        return

    if novos_dados.empty:
        progresso.empty()
        st.warning('Não há dados novos na planilha')
        return
    
    progresso.progress(4/5, 'Enviando Dados...')

    try:
        up.adicionar_registros(novos_dados, 'base_de_vendas_megaferro')
        progresso.progress(5/5, 'Finalizado!')
        progresso.empty()
        st.success(f'Base de dados atualizada com sucesso! {len(novos_dados)} registros adicionados')
    except Exception as e:
        progresso.empty()
        st.warning(f'Houve um erro ao enviar dados. Erro: {e}')
        return

def page_upload():
    st.write('# MEGAFERRO')
    col1, col2 = st.columns(2, gap='large')

    with col1:
        st.header('Enviar para o banco de dados')
        st.divider()
        arquivo = st.file_uploader('**Planilha**', ['xlsx', 'xls'])
        botao = st.button('Enviar')

    with col2:
        st.header('Logs')
        st.divider()

        if arquivo and botao:
            processar_arquivo(arquivo)
            

                

def main():
    st.set_page_config(layout='wide', page_title="Frigosol")
    page_upload()

if __name__ == "__main__":
    main()