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
import io

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

    zona_entrega_mais_comum = df['Zona Entrega'].mode()[0]
    total_nfs = df['Nota'].nunique()
    peso_por_status = df.groupby('Status')['KG'].sum().round(2).to_dict()
    status_peso_text = ' \\ '.join([f"{status}: {peso} KG" for status, (peso) in peso_por_status.items()])
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
        #elif first_pdf == "segundo":
         #   elementssubheader_table
        
        subtotal_peso = round(dataframe['KG'].sum(), 2)
        subtotal_notas = dataframe['Nota'].nunique()
        subtitle_text = f"{almoxarifado} - Peso: {subtotal_peso} KG - Notas: {subtotal_notas}"
        elements.append(Paragraph(subtitle_text, subtitle_style))
        elements.append(Spacer(1, -0.1 * inch))
        
        cols_para_impressao = ['Chegada', 'Nota', 'Etq. Unica', 'CTE', 'KG', 'Vol', 'Prior', 'ONU', 'Remetente',
                               'Almoxarifado', 'Mercadoria', 'Data Entrega', 'End. WMS', 'Lim. Embarque', 'Status']
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

# Função para ordenar DataFrame por almoxarifado
def teste(df):
    ordem_regiao_wms = ['Quimico','Cantilever', 'Praça Externa 1', 'Praça Externa 2', 'Praça - Rua 00 - Posição 01',
                        'Praça - Rua 01 - Posição 05', 'Praça - Rua 01 - Posição 06','Praça - Rua 01 - Posição 07',
                        'Praça - Rua 01 - Posição 08','Praça - Rua 01 - Posição 09','Praça - Rua 02 - Posição 01',
                        'Praça - Rua 02 - Posição 02','Praça - Rua 02 - Posição 03','Praça - Rua 02 - Posição 04',
                        'Praça - Rua 02 - Posição 05','Praça - Rua 02 - Posição 06','Praça - Rua 02 - Posição 07',
                        'Praça - Rua 02 - Posição 08','Praça - Rua 02 - Posição 09','Praça - Rua 03 - Posição 01',
                        'Praça - Rua 03 - Posição 02','Praça - Rua 03 - Posição 03','Praça - Rua 03 - Posição 04',
                        'Praça - Rua 03 - Posição 05','Praça - Rua 03 - Posição 06','Praça - Rua 03 - Posição 07',
                        'Praça - Rua 03 - Posição 08','Praça - Rua 03 - Posição 09','Praça - Rua 04 - Posição 01',
                        'Praça - Rua 04 - Posição 02','Praça - Rua 04 - Posição 03','Praça - Rua 04 - Posição 04',
                        'Praça - Rua 04 - Posição 05','Praça - Rua 04 - Posição 06','Praça - Rua 04 - Posição 07',
                        'Praça - Rua 04 - Posição 08','Praça - Rua 04 - Posição 09','Praça - Rua 05 - Posição 01',
                        'Praça - Rua 05 - Posição 02','Praça - Rua 05 - Posição 03','Praça - Rua 05 - Posição 04',
                        'Praça - Rua 05 - Posição 05','Praça - Rua 05 - Posição 06','Praça - Rua 05 - Posição 07',
                        'Praça - Rua 05 - Posição 08','Praça - Rua 05 - Posição 09','Praça - Rua 06 - Posição 01',
                        'Praça - Rua 06 - Posição 02','Praça - Rua 06 - Posição 03','Praça - Rua 06 - Posição 04',
                        'Praça - Rua 06 - Posição 05','Praça - Rua 06 - Posição 06','Praça - Rua 06 - Posição 07',
                        'Praça - Rua 06 - Posição 08','Praça - Rua 06 - Posição 09','Praça - Rua 07 - Posição 01',
                        'Praça - Rua 07 - Posição 02','Praça - Rua 07 - Posição 03','Praça - Rua 07 - Posição 04',
                        'Praça - Rua 07 - Posição 05','Praça - Rua 07 - Posição 06','Praça - Rua 07 - Posição 07',
                        'Praça - Rua 07 - Posição 08','Praça - Rua 07 - Posição 09','Praça - Rua 08 - Posição 01',
                        'Praça - Rua 08 - Posição 02','Praça - Rua 08 - Posição 03','Praça - Rua 08 - Posição 04',
                        'Praça - Rua 08 - Posição 05','Praça - Rua 08 - Posição 06','Praça - Rua 08 - Posição 07',
                        'Praça - Rua 08 - Posição 08','Praça - Rua 08 - Posição 09','Praça - Rua 09 - Posição 01',
                        'Praça - Rua 09 - Posição 02','Praça - Rua 09 - Posição 03','Praça - Rua 09 - Posição 04',
                        'Praça - Rua 09 - Posição 05','Praça - Rua 09 - Posição 06','Praça - Rua 09 - Posição 07',
                        'Praça - Rua 09 - Posição 08','Praça - Rua 09 - Posição 09','Praça - Rua 10 - Posição 01',
                        'Praça - Rua 10 - Posição 02','Praça - Rua 10 - Posição 03','Praça - Rua 10 - Posição 04',
                        'Praça - Rua 10 - Posição 05','Praça - Rua 10 - Posição 06','Praça - Rua 10 - Posição 07',
                        'Praça - Rua 10 - Posição 08','Praça - Rua 10 - Posição 09','Praça - Rua 11 - Posição 01',
                        'Praça - Rua 11 - Posição 02','Praça - Rua 12 - Posição 01','Praça - Rua 12 - Posição 02',
                        'Praça - Rua 12 - Posição 03','Praça - Rua 12 - Posição 04','Praça - Rua 12 - Posição 05',
                        'Praça - Rua 12 - Posição 06','Praça - Rua 12 - Posição 07','Praça - Rua 12 - Posição 08',
                        'Praça - Rua 12 - Posição 09','Praça - Rua 13 - Posição 01','Praça - Rua 13 - Posição 02',
                        'Praça - Rua 13 - Posição 03','Praça - Rua 13 - Posição 04','Praça - Rua 13 - Posição 05',
                        'Praça - Rua 13 - Posição 06','Praça - Rua 13 - Posição 07','Praça - Rua 13 - Posição 08',
                        'Praça - Rua 13 - Posição 09','Praça - Rua 14 - Posição 01','Praça - Rua 14 - Posição 02',
                        'Praça - Rua 14 - Posição 03','Praça - Rua 14 - Posição 04','Praça - Rua 14 - Posição 05',
                        'Praça - Rua 14 - Posição 06','Praça - Rua 14 - Posição 07','Praça - Rua 14 - Posição 08',
                        'Praça - Rua 14 - Posição 09','Praça - Rua 15 - Posição 01','Praça - Rua 15 - Posição 02',
                        'Praça - Rua 15 - Posição 03','Praça - Rua 15 - Posição 04','Praça - Rua 15 - Posição 05',
                        'Praça - Rua 15 - Posição 06','Praça - Rua 15 - Posição 07','Praça - Rua 15 - Posição 08',
                        'Praça - Rua 15 - Posição 09','Praça - Rua 16 - Posição 01','Praça - Rua 16 - Posição 02',
                        'Praça - Rua 16 - Posição 03','Praça - Rua 16 - Posição 04','Praça - Rua 16 - Posição 06',
                        'Praça - Rua 16 - Posição 07','Praça - Rua 16 - Posição 09','Porta Palete A', 'Porta Palete B',
                        'Porta Palete C', 'Porta Palete D', 'Porta Palete E', 'Porta Palete F', 'Porta Palete G',
                        'Porta Palete H', 'Porta Palete I', 'Porta Palete J', 'Porta Palete K', 'Praça de Consolidação',
                        'Rua 08', 'Rua 12', 'Rua 13', 'Rua 14', 'Rua 15', 'Rua 16', 'Stage', 'Sem Posição']
    
    dataframes_por_almoxarifado = dict(tuple(df.groupby('Região WMS')))
    
    dataframes_ordenados = OrderedDict({
        regiao: dataframes_por_almoxarifado[regiao]
        for regiao in ordem_regiao_wms if regiao in dataframes_por_almoxarifado
    })
    
    dataframes_ordenados.update({
        regiao: df for regiao, df in dataframes_por_almoxarifado.items()
        if regiao not in ordem_regiao_wms
    })
    
    return dataframes_ordenados

# Interface do Streamlit
st.title("Gerador de PDF de Carregamento")

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Realiza as manipulações no DataFrame
    df = df.iloc[:-2]
    df = df[['Chegada', 'Nota Numero', 'Etiqueta Unica', 'Chave Conhecimento', 'Peso Nota', 'Nota Volumes', 
             'Tp. Solicitacao Coleta', 'Nº ONU', 'Razao Remetente', 'Almox. Destino', 'Mercadoria Descricao', 
             'Limite Entregar (Definitivo)', 'Endereco WMS', 'Data Limite Embarque', 'Zona Entrega', 'Região WMS']]
    df['Chave Conhecimento'] = df['Chave Conhecimento'].fillna('')
    df['Tp. Solicitacao Coleta'] = df['Tp. Solicitacao Coleta'].str.slice(0, 1)
    df['Mercadoria Descricao'] = df['Mercadoria Descricao'].str.slice(0, 15)
    df['Almox. Destino'] = df['Almox. Destino'].str.slice(0, 25)
    df['Razao Remetente'] = df['Razao Remetente'].str.slice(0, 25)
    df['Limite Entregar (Definitivo)'] = pd.to_datetime(df['Limite Entregar (Definitivo)'], errors='coerce')
    df['Limite Entregar (Definitivo)'] = df['Limite Entregar (Definitivo)'].dt.strftime('%d/%m/%y')
    df['Chegada'] = pd.to_datetime(df['Chegada'], errors='coerce')
    df['Chegada'] = df['Chegada'].dt.strftime('%d/%m/%y %H:%M')
    df['Data Limite Embarque'] = pd.to_datetime(df['Data Limite Embarque'], errors='coerce')
    df['Data Limite Embarque'] = df['Data Limite Embarque'].dt.strftime('%d/%m/%y')
    df['Nota Numero'] = df['Nota Numero'].astype(int)
    df['Nota Volumes'] = df['Nota Volumes'].astype(int)
    df['Etiqueta Unica'] = pd.to_numeric(df['Etiqueta Unica'], errors='coerce').fillna('')
    df['Etiqueta Unica'] = df['Etiqueta Unica'].astype('string')
    df['Etiqueta Unica'] = df['Etiqueta Unica'].apply(lambda x: x[:-2] if x.endswith('.0') else x).replace('nan', '')
    df['Nº ONU'] = pd.to_numeric(df['Nº ONU'], errors='coerce').fillna('')
    df['Nº ONU'] = df['Nº ONU'].astype('string')
    df['Nº ONU'] = df['Nº ONU'].apply(lambda x: x[:-2] if x.endswith('.0') else x).replace('nan', '')
    df['Endereco WMS'] = df['Endereco WMS'].fillna('')
    df['Peso Nota'] = df['Peso Nota'].round(2)
    
    novos_nomes = {
        'Chegada': 'Chegada',
        'Nota Numero': 'Nota',
        'Etiqueta Unica': 'Etq. Unica',
        'Chave Conhecimento': 'CTE',
        'Peso Nota': 'KG',
        'Nota Volumes': 'Vol',
        'Tp. Solicitacao Coleta': 'Prior',
        'Nº ONU': 'ONU',
        'Razao Remetente': 'Remetente',
        'Almox. Destino': 'Almoxarifado',
        'Mercadoria Descricao': 'Mercadoria',
        'Limite Entregar (Definitivo)': 'Data Entrega',
        'Data Limite Embarque': 'Lim. Embarque',
        'Endereco WMS': 'End. WMS'
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

    df_atrasado = teste(df_atrasado)
    df_hoje = teste(df_hoje)

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

# Botão para download do PDF
    st.download_button(
    label="Baixar PDF",
    data=output_buffer.getvalue(),
    file_name="MapaDeCarregamento.pdf",
    mime="application/pdf"
)