import streamlit as st
import upload as up
import pandas as pd
import os
import bitrix as btx

def processar_arquivo(arquivo):
    st.markdown(f'### -> {arquivo.name}')
    #progresso = st.progress(0, 'Lendo arquivo...')

    with st.status("Processando...") as status:
        try: 
            data_frame = pd.read_excel(arquivo, header=2, skipfooter=1)
        except:
            status.update(label='Falha na Execução!', state='error')
            st.error('Erro ao abrir arquivo')
            return
        
        st.write('Validando planilha...')
        try:
            up.validar_planilha(data_frame)
        except Exception as e:
            status.update(label='Falha na Execução!', state='error')
            st.error(f'Planilha inválida. Erro: {e}')
            return

        st.write('Formatando Planilha...')
        try: 
            df_formatado = up.formatar_data_frame(data_frame)
        except Exception as e:
            status.update(label='Falha na Execução!', state='error')
            st.error(f'Erro ao formatar planilha: {e}')
            return

        st.write('Buscando cards...')
        cards = btx.deal_list(
            {
                "CATEGORY_ID": ["10", "16"],
                "STAGE_ID": ["C10:WON", "C16:WON"],
                "UF_CRM_1748271797931": None
            },
            [
                "TITLE", "CONTACT_ID", "UF_CRM_1747310398835",
                "UF_CRM_1740679069771", "UF_CRM_1740679233003",
                "UF_CRM_1740679239002", "UF_CRM_1740679274315",
                "OPPORTUNITY", "CLOSEDATE", "UF_CRM_1738004398103",
                "CATEGORY_ID", "STAGE_ID"
            ]
        )

        st.write('Processando Cards...')
        for card in cards:
            id_card = card.pop('ID')
            data_venda = pd.to_datetime(card.get('CLOSEDATE')).replace(tzinfo = None)
            valor_card = float(card.get('OPPORTUNITY'))

            codigos = card.get("UF_CRM_1747310398835", '')
            
            if not codigos:
                continue
            
            codigos = [codigo.strip() for codigo in codigos.split(";")]

            codigos_encontrados = []
            codigos_nao_encontrados = []

            for codigo in codigos:
                matches = df_formatado[df_formatado['unico'].astype('str') == codigo]
                
                if not matches.empty:
                    codigos_encontrados.append({
                        "codigos_associados": "; ".join(codigos),
                        "numero_nota": matches['n_nota'].max().item(),
                        "valor_faturado": matches['valor_total'].sum().item(),
                        "data_faturamento": matches['data'].max(),
                        "codigo_venda": matches['unico'].max().item()
                    })
                else:
                    codigos_nao_encontrados.append(codigo)
        
            if not codigos_encontrados:
                continue

            if codigos_nao_encontrados:
                valor_encontrado = 0

                for codigo in codigos_encontrados:
                    codigos_associados = codigo.get('codigos_associados')
                    numero_nota = codigo.get('numero_nota')
                    valor_faturado = codigo.get('valor_faturado')
                    data_faturamento = codigo.get('data_faturamento')
                    codigo_venda = codigo.get('codigo_venda')

                    retroatividade = "236" if data_venda.month < data_faturamento.month else "238"

                    valor_encontrado += valor_faturado


                    st.write(f"Criando Card - Codigo venda: {codigo_venda}...")
                    try:
                        response = btx.deal_add({ 
                            **card,            
                            **{                       
                                "UF_CRM_1748271817019": codigos_associados,
                                "UF_CRM_1748271797931": numero_nota,
                                "UF_CRM_1748271650340": retroatividade,
                                "UF_CRM_1748271599059": valor_faturado,
                                "UF_CRM_1748271579108": data_faturamento.isoformat(),
                                "UF_CRM_1747310398835": codigo_venda,
                                "OPPORTUNITY": valor_faturado
                            }
                        })
                        st.success(f"Card criado com sucesso! - ID: {response.get('result')}")
                    except Exception as e:
                        status.update(label='Falha na Execução!', state='error')
                        st.error(f'Erro ao criar card: {e}')
                        return

                diferenca = valor_card - valor_encontrado

                st.write(f"Atualizando Card - ID: {id_card}...")
                try:
                    response = btx.deal_update(
                        id_card,
                        {   
                            "OPPORTUNITY": diferenca,
                            "UF_CRM_1747310398835": "; ".join(codigos_nao_encontrados),
                            "UF_CRM_1748271817019": codigos_associados
                        }
                    )
                    st.success(f"Card Atualizado com sucesso! - ID: {response.get('result')}")
                except Exception as e:
                    status.update(label='Falha na Execução!', state='error')
                    st.error(f'Erro ao atualizar card: {e}')
                    return
            else:
                for codigo in codigos_encontrados[:-1]:
                    codigos_associados = codigo.get('codigos_associados')
                    numero_nota = codigo.get('numero_nota')
                    valor_faturado = codigo.get('valor_faturado')
                    data_faturamento = codigo.get('data_faturamento')
                    codigo_venda = codigo.get('codigo_venda')

                    retroatividade = "236" if data_venda.month < data_faturamento.month else "238"

                    st.write(f"Criando Card - Codigo venda: {codigo_venda}...")
                    try:
                        response = btx.deal_add({
                            **card,            
                            **{
                                "UF_CRM_1748271817019": codigos_associados,
                                "UF_CRM_1748271797931": numero_nota,
                                "UF_CRM_1748271650340": retroatividade,
                                "UF_CRM_1748271599059": valor_faturado,
                                "UF_CRM_1748271579108": data_faturamento.isoformat(),
                                "UF_CRM_1747310398835": codigo_venda,
                                "OPPORTUNITY": valor_faturado
                            }
                        })
                        st.success(f"Card criado com sucesso! - ID: {response.get('result')}")
                    except Exception as e:
                        status.update(label='Falha na Execução!', state='error')
                        st.error(f'Erro ao criar card: {e}')
                        return

                codigo = codigos_encontrados[-1]

                codigos_associados = codigo.get('codigos_associados')
                numero_nota = codigo.get('numero_nota')
                valor_faturado = codigo.get('valor_faturado')
                data_faturamento = codigo.get('data_faturamento')
                codigo_venda = codigo.get('codigo_venda')

                retroatividade = "236" if data_venda.month < data_faturamento.month else "238"

                st.write(f"Atualizando Card - ID: {id_card}...")
                try:
                    response = btx.deal_update(
                        id_card,
                        {
                            "UF_CRM_1748271817019": codigos_associados,
                            "UF_CRM_1748271797931": numero_nota,
                            "UF_CRM_1748271650340": retroatividade,
                            "UF_CRM_1748271599059": valor_faturado,
                            "UF_CRM_1748271579108": data_faturamento.isoformat(),
                            "UF_CRM_1747310398835": codigo_venda,
                            "OPPORTUNITY": valor_faturado,
                        }
                    )
                    st.success(f"Card Atualizado com sucesso! - ID: {response.get('result')}")
                except Exception as e:
                    status.update(label='Falha na Execução!', state='error')
                    st.error(f'Erro ao atualizar card: {e}')
                    return
        
        status.update(label='Execução Completa!', state='complete')
    
    with st.status("Enviando para a base de dados") as status:
        st.write('Filtrando Registros...')

        try:
            novos_dados = up.filtrar_novos_dados(df_formatado, 'base_de_vendas_megaferro')
        except Exception as e:
            status.update(label='Falha na Execução!', state='error')
            st.error(f'Erro ao se conectar ao banco de dados. Error: {e}')
            return

        if novos_dados.empty:
            status.update(label='Falha na Execução!', state='error')
            st.warning('Não há dados novos na planilha')
            return
        
        st.write('Enviando Dados...')

        try:
            up.adicionar_registros(novos_dados, 'base_de_vendas_megaferro')
            st.success(f'Base de dados atualizada com sucesso! {len(novos_dados)} registros adicionados')
            status.update(label='Envio concluido com sucesso!', state='complete')
        except Exception as e:
            status.update(label='Falha na Execução!', state='error')
            st.error(f'Houve um erro ao enviar dados. Erro: {e}')
            return

def processar_nao_faturado(arquivo):
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
        up.validar_planilha_nao_faturado(data_frame)
    except Exception as e:
        progresso.empty()
        st.error(f'Planilha inválida. Erro: {e}')
        return

    progresso.progress(2/5, 'Formatando Planilha...')
    try: 
        df_formatado = up.formatar_df_nao_faturado(data_frame)
    except Exception as e:
        progresso.empty()
        st.error(f'Erro ao formatar planilha: {e}')
        return

    st.write(df_formatado)
    
    progresso.progress(4/5, 'Enviando Dados...')

    try:
        up.adicionar_registros(df_formatado, 'nao_faturado_megaferro', replace=True)
        progresso.progress(5/5, 'Finalizado!')
        progresso.empty()
        st.success(f'Base de dados atualizada com sucesso! {len(df_formatado)} registros adicionados')
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
        arquivo = st.file_uploader('**Faturado**', ['xlsx', 'xls'])
        nao_faturado = st.file_uploader("**Não Faturado**", ['xlsx', 'xls'])
        botao = st.button('Enviar')
    with col2:
        st.header('Logs')
        st.divider()

        if (arquivo or nao_faturado) and botao:
            if arquivo:
                processar_arquivo(arquivo)
            if nao_faturado:
                processar_nao_faturado(nao_faturado)

def main():
    st.set_page_config(layout='wide', page_title="Megaferro")
    page_upload()

if __name__ == "__main__":
    main()