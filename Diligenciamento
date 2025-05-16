import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
import io

def format_brazilian(number):
    return f"{number:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def create_pdf(of_coleta, placa, group_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=6, bottomMargin=6)
    elements = []

    styles = getSampleStyleSheet()
    header = Paragraph(f"OF_COLETA: {of_coleta} - Placa: {placa}", styles['Heading2'])
    elements.append(header)

    table_data = [['OEF', 'Fornecedor', 'Cliente Pai', 'Peso']]
    total_peso = 0
    num_oef_set = set()
    for index, row in group_df.iterrows():
        peso = float(row['Peso']) if pd.notnull(row['Peso']) else 0
        table_data.append([
            row['OEF'],
            row['Fornecedor'],
            row['Cliente Pai'],
            f"{peso:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        ])
        total_peso += peso
        num_oef_set.add(row['OEF'])

    table_data.append(['', '', 'Total Peso', f"{total_peso:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])
    table_data.append(['', '', 'Qtd. Distinta OEF', str(len(num_oef_set))])
    
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('BACKGROUND', (-2, -2), (-1, -2), colors.lightgrey),
        ('BACKGROUND', (-2, -1), (-1, -1), colors.lightgrey),
        ('TEXTCOLOR', (-2, -2), (-1, -2), colors.black),
        ('TEXTCOLOR', (-2, -1), (-1, -1), colors.black),
        ('ALIGN', (-2, -2), (-1, -2), 'RIGHT'),
        ('ALIGN', (-2, -1), (-1, -1), 'RIGHT'),
        ('FONTNAME', (-2, -2), (-1, -2), 'Helvetica-Bold'),
        ('FONTNAME', (-2, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (-2, -2), (-1, -2), 8),
        ('FONTSIZE', (-2, -1), (-1, -1), 8),
        ('BOTTOMPADDING', (-2, -2), (-1, -2), 4),
        ('BOTTOMPADDING', (-2, -1), (-1, -1), 4),
    ]))
    elements.append(table)

    doc.build(elements)
    
    return buffer

def create_all_pdfs(grouped, df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=6, bottomMargin=6)
    elements = []
    styles = getSampleStyleSheet()

    # --- COVER PAGE ---
    cover_header = Paragraph("<b>Resumo das Placas, Motoristas e OF</b>", styles['Title'])
    elements.append(cover_header)
    cover_table_data = [['Placa', 'Motorista(s)', 'OF(s)', 'Total Peso', 'Qtd. Distinta OEF']]
    for placa, placa_df in df.groupby('Placa'):
        ofs = sorted(placa_df['OF'].unique())
        ofs_str = ', '.join(str(ofc) for ofc in ofs)
        motoristas = sorted(placa_df['Motorista'].dropna().unique())
        motoristas_str = ', '.join(str(m) for m in motoristas)
        total_peso = placa_df['Peso'].fillna(0).sum()
        num_oef_count = placa_df['OEF'].nunique()
        cover_table_data.append([
            placa,
            motoristas_str,
            ofs_str,
            f"{total_peso:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            str(num_oef_count)
        ])
    cover_table = Table(cover_table_data)
    cover_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
    ]))
    elements.append(cover_table)
    elements.append(PageBreak())
    # --- END COVER PAGE ---

    for (of, placa), group_df in grouped:
        motoristas = sorted(group_df['Motorista'].dropna().unique())
        motoristas_str = ', '.join(str(m) for m in motoristas)
        header = Paragraph(f"OF: {of} - Placa: {placa} - Motorista: {motoristas_str}", styles['Heading2'])
        elements.append(header)
        table_data = [['OEF', 'Fornecedor', 'Cliente Pai', 'Peso']]
        total_peso = 0
        num_oef_set = set()
        for index, row in group_df.iterrows():
            peso = float(row['Peso']) if pd.notnull(row['Peso']) else 0
            table_data.append([
                row['OEF'],
                row['Fornecedor'],
                row['Cliente Pai'],
                f"{peso:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            ])
            total_peso += peso
            num_oef_set.add(row['OEF'])
        table_data.append(['', '', 'Total Peso', f"{total_peso:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])
        table_data.append(['', '', 'Qtd. Distinta OEF', str(len(num_oef_set))])
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('BACKGROUND', (-2, -2), (-1, -2), colors.lightgrey),
            ('BACKGROUND', (-2, -1), (-1, -1), colors.lightgrey),
            ('TEXTCOLOR', (-2, -2), (-1, -2), colors.black),
            ('TEXTCOLOR', (-2, -1), (-1, -1), colors.black),
            ('ALIGN', (-2, -2), (-1, -2), 'RIGHT'),
            ('ALIGN', (-2, -1), (-1, -1), 'RIGHT'),
            ('FONTNAME', (-2, -2), (-1, -2), 'Helvetica-Bold'),
            ('FONTNAME', (-2, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (-2, -2), (-1, -2), 8),
            ('FONTSIZE', (-2, -1), (-1, -1), 8),
            ('BOTTOMPADDING', (-2, -2), (-1, -2), 4),
            ('BOTTOMPADDING', (-2, -1), (-1, -1), 4),
        ]))
        elements.append(table)
        elements.append(Paragraph('<br/><br/>', styles['Normal']))  # Add space/new page
        elements.append(PageBreak())
    if elements and isinstance(elements[-1], PageBreak):
        elements = elements[:-1]  # Remove last page break
    doc.build(elements)
    return buffer

st.title("Gerador de PDF por Vendedor")
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    grouped = df.groupby(['OF', 'Placa'])

    # Move the combined PDF button to the top
    all_pdf_buffer = create_all_pdfs(grouped, df)
    all_pdf_buffer.seek(0)
    st.download_button(
        label="Baixar PDF Único com Todos",
        data=all_pdf_buffer.getvalue(),
        file_name="Todos_OF_Placa.pdf",
        mime="application/pdf"
    )

    for (of, placa), group_df in grouped:
        pdf_buffer = create_pdf(of, placa, group_df)
        
        pdf_buffer.seek(0)
        
        st.download_button(
            label=f"Baixar PDF - {of} - {placa}",
            data=pdf_buffer.getvalue(),
            file_name=f"{of}_{placa}.pdf",
            mime="application/pdf"
        )
