import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
from datetime import datetime
from collections import OrderedDict
from PyPDF2 import PdfMerger
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import io
import unicodedata
import re
from io import BytesIO

def downloadSharePoint(file_url):
    # Autenticação no SharePoint
    ctx_auth = AuthenticationContext(url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url, ctx_auth)
        
        # Buffer para armazenar o conteúdo do arquivo
        file_buffer = BytesIO()
        
        # Download do arquivo para o buffer
        ctx.web.get_file_by_server_relative_url(file_url).download(file_buffer).execute_query()
        
        # Posicionar o cursor no início do buffer
        file_buffer.seek(0)
        
        # Ler o CSV diretamente do buffer
        try:
            df = pd.read_parquet(file_buffer)
        except:
            try:
                df = pd.read_excel(file_buffer)
            except Exception as e:
                print(f"Erro ao ler o arquivo: {e}")
                exit()
        return df
        
    else:
        print("Autenticação falhou.")
        exit()

def remover_acentos(texto):
    if isinstance(texto, str):
        texto = unicodedata.normalize('NFKD', texto)
        texto = ''.join(c for c in texto if not unicodedata.combining(c))
        return texto
    return texto

def tratar_dados(df, colunas=None):
    if colunas is None:
        colunas = df.columns  # Se nenhuma coluna for especificada, trata todas
    
    for coluna in colunas:
        df[coluna] = (
            df[coluna]
            .astype(str)  # Garante que os valores são strings
            .str.strip()  # Remove espaços extras no início e fim
            .apply(remover_acentos)  # Remove acentos
            .str.replace(r'[^a-zA-Z0-9\s]', '', regex=True)  # Remove caracteres especiais
            .str.replace(r'\s+', ' ', regex=True)  # Substitui múltiplos espaços por um único
            .str.upper()  # Converte para maiúsculas
        )
    return df

# Função para gerar PDF a partir de um DataFrame
def dataframe_to_pdf(dataframe_dict, buffer):
    pdf = SimpleDocTemplate(buffer, pagesize=landscape(A4), topMargin=6, bottomMargin=6)
    elements = []
    c = Canvas("temp.pdf")
    styles = getSampleStyleSheet()
    title_style = styles["Heading2"]
    subtitle_style = styles["Heading2"]
    centered_subtitle_style = ParagraphStyle(name="CenteredSubtitle", parent=subtitle_style, alignment=1)
    centered_title_style = ParagraphStyle(name="CenteredTitle", parent=title_style, alignment=1)

    zona_entrega_mais_comum = df['Almoxarifado'].mode()[0]
    zona_entrega_mais_comum = df['Almoxarifado'].mode()[0]
    total_nfs = df['Nota'].nunique()
    peso_por_status = df.groupby('Status')['KG'].sum().round(2).to_dict()
    status_peso_text = ' \\\\ '.join([f"{status}: {peso} KG" for status, (peso) in peso_por_status.items()])
    
    title_text_line1 = f"Relatório de Carregamento - Zona de Entrega: {zona_entrega_mais_comum}"
    title_text_line2 = f"Peso Total: {subtotal_total} KG - NFs Totais: {total_nfs}"
    title_text_line3 = f"{status_peso_text}"
    
    header_elements = [
        Paragraph(title_text_line1, centered_title_style),
        Spacer(1, -0.3 * inch),
        Paragraph(title_text_line2, centered_title_style),
        Spacer(1, -0.3 * inch),
        Paragraph(title_text_line3, centered_title_style),
        Spacer(1, 0.00005 * inch)
    ]
    
    header_table = Table([[header_elements]], colWidths=[landscape(A4)[0]])
    header_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.white)]))
    
    subtitle_line1 = f"Notas Dentro do Prazo"
    subtitle_elements = [
        Paragraph(subtitle_line1,title_style),
        Spacer(1, -0.3 * inch)
    ]
    
    subheader_table = Table([[subtitle_elements]], colWidths=[landscape(A4)[0]])
    
    first_pdf = "primeiro"
    
    for almoxarifado, dataframe in dataframe_dict.items():
        if first_pdf == "primeiro":
            elements.append(header_table)
            first_pdf = "segundo"
        
        subtotal_peso = round(dataframe['KG'].sum(), 2)
        subtotal_notas = dataframe['Nota'].nunique()
        subtitle_text = f"{almoxarifado} - Peso: {subtotal_peso} KG - Notas: {subtotal_notas}"
        elements.append(Paragraph(subtitle_text, subtitle_style))
        elements.append(Spacer(1, -0.1 * inch))
        
        cols_para_impressao =   ['Chegada', 'Nota', 'Etq. Unica', 'CTE', 'KG', 'Vol', 'Prior', 'ONU', 'Remetente',
                              'Almoxarifado', 'Mercadoria', 'Data Entrega', 'Lim. Embarque', 'Status', 'End. WMS']
    
        data = dataframe[cols_para_impressao].values.tolist()
        headers = dataframe[cols_para_impressao].columns.tolist()
        
        col_widths = [50, 50, 40, 45, 30, 10, 20, 20, 120, 120, 90, 50, 70, 50, 60]
        table = Table([headers] + data, colWidths=col_widths)
        
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
            ('FONTSIZE', (0, 1), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 0.7),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0.7),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.gray)
        ])
        
        table.setStyle(style)
        
        elements.append(table)
        elements.append(Spacer(1, 0.0001 * inch))
    
    pdf.build(elements)

# Função para ordenar DataFrame por almoxarifado ou região WMS
from collections import OrderedDict

from collections import OrderedDict

def ordenar_dataframe(df):
    ordem_regiao_wms = ['Quimico', 'Cantilever', 'Praça Externa 1', 'Praça Externa 2', 'Praça - Rua 00 - Posição 01']
    ordem_prioridade = {'P': 0, '0': 1, '1': 2, '2': 3, '3': 4, '4': 5, '5': 6, 'N': 7}

    def ordenar_por_prioridade(df):
        """Ordena o DataFrame pela coluna 'Prioridade' conforme a ordem definida."""
        df['Prioridade_Ordenada'] = df['Prior'].map(ordem_prioridade).fillna(99)
        df = df.sort_values(by='Prioridade_Ordenada', ascending=True).drop(columns=['Prioridade_Ordenada'])
        return df

    if 'Região WMS' in df.columns:
        dataframes_por_regiao_wms = dict(tuple(df.groupby('Região WMS')))
        dataframes_ordenados_regiao_wms = OrderedDict()

        for regiao in ordem_regiao_wms:
            if regiao in dataframes_por_regiao_wms:
                dataframes_ordenados_regiao_wms[regiao] = ordenar_por_prioridade(dataframes_por_regiao_wms[regiao])

        for regiao, df_outros in dataframes_por_regiao_wms.items():
            if regiao not in ordem_regiao_wms:
                dataframes_ordenados_regiao_wms[regiao] = ordenar_por_prioridade(df_outros)

        return dataframes_ordenados_regiao_wms
    
    else:
        dataframes_por_almoxarifado = dict(tuple(df.groupby('Almoxarifado')))
        return {almox: ordenar_por_prioridade(df) for almox, df in dataframes_por_almoxarifado.items()}

# Interface do Streamlit
st.title("Gerador de PDF de Carregamento - SAP")
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Verifica se a coluna 'Região WMS' existe no DataFrame
    if 'Região WMS' in df.columns:
        colunas_para_selecao = ['Chegada', 'Nota Numero', 'Etiqueta Unica', 'Chave Conhecimento', 'Peso Nota', 'Nota Volumes',
                                'Tp. Solicitacao Coleta', 'Nº ONU', 'Razao Remetente', 'Almox. Destino', 'Mercadoria Descricao',
                                'Limite Entregar (Definitivo)', 'Endereco WMS', 'Data Limite Embarque', 'Zona Entrega', 'Região WMS']
    else:
        colunas_para_selecao = ['Data de Entrada da Nota', 'Nota', 'Número da Etiqueta Única', 'CTRC', 'Peso', 'Quantidade de volumes', 'Prioridade',
                               'ONU', 'Fornecedor', 'Almoxarifado', 'Mercadoria', 'Dt Final Entrega', 'Dt Lim Embarque', 'End. WMS']
        
        #['Chegada', 'Nota Numero', 'Etiqueta Unica', 'Chave Conhecimento', 'Peso Nota', 'Nota Volumes',
         #                       'Tp. Solicitacao Coleta', 'Nº ONU', 'Razao Remetente', 'Almox. Destino', 'Mercadoria Descricao',
          #                      'Limite Entregar (Definitivo)', 'Endereco WMS', 'Data Limite Embarque', 'Zona Entrega']

    # Seleciona as colunas no DataFrame
    df = df[colunas_para_selecao]
    df = remover_acentos(df)
    df = tratar_dados(df, ['Fornecedor'])
    df['Fornecedor'] = df['Fornecedor'].str.replace(' ','')
    
    # Realiza as manipulações no DataFrame
    df = df.iloc[:-2] # precisava retirar as duas linhas quando o arquivo vem do bi, vindo do sap nao tem uma linha embaixo especificando os filtros
    df['CTRC'] = df['CTRC'].fillna('')
    df['Prioridade'] = df['Prioridade'].astype(str)
    df['Prioridade'] = df['Prioridade'].str.slice(0, 1)
    df['Mercadoria'] = df['Mercadoria'].str.slice(0, 15)
    df['Almoxarifado'] = df['Almoxarifado'].str.slice(0, 25)
    df['Fornecedor'] = df['Fornecedor'].str.slice(0, 25)
    df['Data Entrega'] = pd.to_datetime(df['Data Entrega'], errors='coerce')
    df['Data Entrega'] = df['Data Entrega'].dt.strftime('%d/%m/%y')
    df['Data de Entrada da Nota'] = pd.to_datetime(df['Data de Entrada da Nota'], errors='coerce')
    df['Data de Entrada da Nota'] = df['Data de Entrada da Nota'].dt.strftime('%d/%m/%y %H:%M')
    df['Data Embarque'] = pd.to_datetime(df['Data Embarque'], errors='coerce')
    df['Data Embarque'] = df['Data Embarque'].dt.strftime('%d/%m/%y')
    df['Nota'] = df['Nota'].astype(int)
    df['Quantidade de volumes'] = df['Quantidade de volumes'].astype(int)
    df['Número da Etiqueta Única'] = pd.to_numeric(df['Número da Etiqueta Única'], errors='coerce').fillna('').astype('string')
    df['Número da Etiqueta Única'] = df['Número da Etiqueta Única'].apply(lambda x: x[:-2] if x.endswith('.0') else x).replace('nan', '')
    df['ONU'] = df['ONU'].astype(str)
    df['ONU'] = df['ONU'].where(df['ONU'].notna() & df['ONU'].str.strip().ne('')).str.slice(0, 5)
    df['ONU'] = pd.to_numeric(df['ONU'], errors='coerce').fillna('').astype('string')
    df['ONU'] = df['ONU'].apply(lambda x: x[:-2] if x.endswith('.0') else x).replace('nan', '')
    # df['Endereco WMS'] = df['Endereco WMS'].fillna('')
    #df['Peso'] = df['Peso'].replace({'.': '', ',': '.'}, regex=True)
    #df['Peso'] = df['Peso'].replace('', pd.NA)
    #df['Peso'] = df['Peso'].fillna(0)
    #df['Peso'] = df['Peso'].astype(float).astype(int)
    #df['Peso'] = df['Peso'].round(2)

    # df['Peso'] = df['Peso'].str.replace(",",".")
    df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').round(2)

    url = "https://dellavolpecombr.sharepoint.com/sites/DellaVolpe/"
    username = "alexandre.sousa@dellavolpe.com.br"
    password = "TDV<10203040acs"

    # dfwms = "/sites/DellaVolpe/Documentos%20Compartilhados/Planejamentos/Dados_PVD/WMS.parquet"
    # dfwms = downloadSharePoint(dfwms) # pd.read_parquet(r"\\tableau\Central_de_Performance\PVD\Base\WMS.parquet")
    
    # df_aux = dfwms[['NOTA_CNPJREMETENTE', 'NOTA_RAZAOREMETENTE']].drop_duplicates().reset_index(drop=True)

    # dfwms_filtrado = dfwms[dfwms['NOTA_CHAVESEFAZ'].str.len() == 44]

    # dfwms_filtrado['CNPJ do Sefaz'] = dfwms_filtrado['NOTA_CHAVESEFAZ'].apply(lambda x: x[6:20] if len(x) == 44 else None)
    # dfwms_filtrado['Numero da NF Sefaz'] = dfwms_filtrado['NOTA_CHAVESEFAZ'].str[25:34].apply(lambda x: str(int(x)))
    # dfwms_filtrado['Numero da NF Sefaz'] = pd.to_numeric(dfwms_filtrado['Numero da NF Sefaz'], errors='coerce').astype('Int64')
    
    # df_mergeado = pd.merge(dfwms_filtrado, df_aux, left_on='CNPJ do Sefaz', right_on='NOTA_CNPJREMETENTE', how='left')
    # df_mergeado = df_mergeado.drop_duplicates(subset=['NOTA_CHAVESEFAZ']) # Aqui eu retiro as duplicatas para impedir que as linham apareçam mais de uma vez quando tiverem mais de uma posiçao wms (por volumes)

    # df_mergeado = remover_acentos(df_mergeado)
    # df_mergeado = tratar_dados(df_mergeado, ['NOTA_RAZAOREMETENTE_y'])
    # df_mergeado['NOTA_RAZAOREMETENTE_y'] = df_mergeado['NOTA_RAZAOREMETENTE_y'].str.replace(' ','')
    
    # df = pd.merge(df, df_mergeado, left_on=['Nota', 'Fornecedor'], right_on=['Numero da NF Sefaz', 'NOTA_RAZAOREMETENTE_y'], how='left')
    
    novos_nomes = {
    'Data de Entrada da Nota': 'Chegada',
    'Nota': 'Nota',
    'Número da Etiqueta Única': 'Etq. Unica',
    'CTRC': 'CTE',
    'Peso': 'KG',
    'Quantidade de volumes': 'Vol',
    'Prioridade': 'Prior',
    'ONU': 'ONU',
    'Fornecedor': 'Remetente',
    'Almoxarifado': 'Almoxarifado',
    'Mercadoria': 'Mercadoria',
    'Dt Final Entrega': 'Data Entrega',
    'Dt Lim Embarque': 'Lim. Embarque',
    'WMS_ENDERECO': 'End. WMS'
    # 'Endereco WMS': 'End. WMS'
    }
    
    df = df.rename(columns=novos_nomes)
    df['Lim. Embarque'] = pd.to_datetime(df['Lim. Embarque'], format='%d/%m/%y', errors='coerce')
    df = df.sort_values(by=['Lim. Embarque', 'Etq. Unica', 'Almoxarifado'])
    subtotal_total = round(df['KG'].sum(), 2)
    df.loc[:, 'Status'] = ''
    today = datetime.today().date()
    
    for index, row in df.iterrows():
        try:
            data_embarque = pd.to_datetime(row['Lim. Embarque'], format='%d/%m/%y').date()
        except ValueError:
            data_embarque = None
        try:
            data_entrega = pd.to_datetime(row['Data Entrega'], format='%d/%m/%y').date()
        except ValueError:
            data_entrega = None
        if data_embarque == today:
            df.at[index, 'Status'] = 'Embarque Hoje'
        elif pd.isna(row['Data Entrega']):
            df.at[index, 'Status'] = 'S/ Data Entrega'
        elif data_embarque < today:
            df.at[index, 'Status'] = 'Embarque Atrasado'
        else:
            df.at[index, 'Status'] = 'Dentro do Prazo'
    
    df['Lim. Embarque'] = df['Lim. Embarque'].dt.strftime('%d/%m/%Y')
    
    # Processamento dos DataFrames e geração dos PDFs
    df_hoje = df.loc[df['Status'] == 'Dentro do Prazo']
    df_atrasado = df.loc[df['Status'] != 'Dentro do Prazo']
    df_atrasado = ordenar_dataframe(df_atrasado)
    df_hoje = ordenar_dataframe(df_hoje)
    
    # Buffer para armazenar os PDFs
    buffer_atrasado = io.BytesIO()
    buffer_hoje = io.BytesIO()
    
    # Gerar os PDFs
    dataframe_to_pdf(df_atrasado, buffer_atrasado)
    dataframe_to_pdf(df_hoje, buffer_hoje)
    
    # Mesclar os PDFs
    buffer_atrasado.seek(0)
    buffer_hoje.seek(0)
    merger = PdfMerger()
    merger.append(buffer_atrasado)
    merger.append(buffer_hoje)
    
    # Salvar o PDF final em um buffer
    output_buffer = io.BytesIO()
    merger.write(output_buffer)
    merger.close()

    data_hoje = datetime.now().strftime('%d%b%y').lower()
    zona_entrega_mais_comum = df['Almoxarifado'].mode()[0]
    zona_almoxarifado = df['Almoxarifado'].mode()[0] if 'Região WMS' in df.columns else df['Almoxarifado'].mode()[0]
    
    # Botão para download do PDF
    st.download_button(
        label="Baixar PDF",
        data=output_buffer.getvalue(),
        file_name=f"{zona_almoxarifado}_{data_hoje}.pdf",
        mime="application/pdf"
    )
