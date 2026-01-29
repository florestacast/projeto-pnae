import warnings
warnings.filterwarnings('ignore', category=UserWarning)

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib import colors
from datetime import datetime
from PIL import Image as PILImage

# Definir data atual
data_atual = datetime.now()

# Função auxiliar para converter valores monetários
def converter_preco(valor):
    if valor is None or valor == '':
        return 0.0
    try:
        # Converter para string, remover 'R$', espaços e substituir vírgula por ponto
        valor_str = str(valor).replace('R$', '').replace('\xa0', '').replace(' ', '').replace(',', '.')
        return float(valor_str)
    except:
        return 0.0

# Função auxiliar para converter quantidades
def converter_quantidade(valor):
    if valor is None or valor == '':
        return 0.0
    try:
        valor_str = str(valor).replace('\xa0', '').replace(' ', '').replace(',', '.')
        return float(valor_str)
    except:
        return 0.0

# Ler arquivo Excel com nomes corretos das abas
arquivo_excel = "projeto_venda.xlsx"
df_administracao = pd.read_excel(arquivo_excel, sheet_name="administracao")
df_produtores = pd.read_excel(arquivo_excel, sheet_name="produtor")
df_edital = pd.read_excel(arquivo_excel, sheet_name="edital")
df_estoque = pd.read_excel(arquivo_excel, sheet_name="estoque")
df_envelope = pd.read_excel(arquivo_excel, sheet_name="envelope")
df_capa = pd.read_excel(arquivo_excel, sheet_name="capa")
df_alimentos = pd.read_excel(arquivo_excel, sheet_name="alimentos")

# Criar PDF com marca d'água
pdf_path = "Projeto_Venda_Escola.pdf"

class PDFComMarcaDagua(SimpleDocTemplate):
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        self.marca_dagua_path = "marca_dagua.png"
    
    def afterPage(self):
        try:
            from reportlab.pdfgen import canvas
            self.canv.saveState()
            
            img = PILImage.open(self.marca_dagua_path)
            img_width = 4*inch      # Aumentar para ficar maior
            img_height = 4*inch     # Aumentar para ficar maior
            
            # Posicionar marca d'água no centro da página
            x = (A4[0] - img_width) / 2
            y = (A4[1] - img_height) / 2
            
            # Ajustar opacidade (0.1=10% a 0.5=50%)
            self.canv.drawImage(self.marca_dagua_path, x, y, width=img_width, height=img_height, preserveAspectRatio=True, opacity=0.2)
            self.canv.restoreState()
        except:
            pass

doc = PDFComMarcaDagua(
    pdf_path,
    pagesize=A4,
    leftMargin=2.5*inch/2.54,
    rightMargin=2.5*inch/2.54,
    topMargin=2.5*inch/2.54,
    bottomMargin=2.5*inch/2.54
)
styles = getSampleStyleSheet()
story = []

# Estilos customizados
titulo_style = ParagraphStyle(
    'titulo',
    parent=styles['Normal'],
    fontSize=12,
    textColor=colors.black,
    alignment=1,  # Centro
    fontName='Helvetica-Bold',
    leading=12 * 1.5
)

titulo_capa_style = ParagraphStyle(
    'titulo_capa',
    parent=styles['Normal'],
    fontSize=16,
    textColor=colors.black,
    alignment=1,  # Centro
    fontName='Helvetica-Bold',
    leading=16 * 1.5
)

titulo_secao_style = ParagraphStyle(
    'titulo_secao',
    parent=styles['Normal'],
    fontSize=9,
    alignment=0,  # Esquerda
    fontName='Helvetica-Bold',
    leading=9 * 1.5
)

centro_style = ParagraphStyle(
    'centro',
    parent=styles['Normal'],
    alignment=1,  # Centro
    leading=10 * 1.5
)

esquerda_style = ParagraphStyle(
    'esquerda',
    parent=styles['Normal'],
    alignment=0,  # Esquerda
    leading=10 * 1.5
)

justificado_style = ParagraphStyle(
    'justificado',
    parent=styles['Normal'],
    alignment=4,  # Justificado
    leading=10 * 1.5
)

# Adicionar espaçamento ao estilo Normal
normal_style = ParagraphStyle(
    'normal_espacado',
    parent=styles['Normal'],
    fontSize=9,
    leading=9 * 1.5
)

# Cabeçalho
try:
    img = Image("cabecalho.png", width=7*inch, height=1*inch)
    story.append(img)
except:
    story.append(Paragraph("CABEÇALHO", titulo_style))

story.append(Spacer(1, 0.3*inch))

# Títulos
story.append(Paragraph("<b>PROJETO DE VENDA DE GÊNEROS ALIMENTÍCIOS DA AGRICULTURA FAMILIAR</b>", titulo_style))
story.append(Paragraph("<b>PARA ALIMENTAÇÃO ESCOLAR/PNAE</b>", titulo_style))
story.append(Spacer(1, 0.10*inch))

# Dados da primeira linha de cada tabela
admin = df_administracao[df_administracao['status_representante'] == 'Ativo'].iloc[0]
edital = df_edital.iloc[0]

def obter_data_assinatura(valor_data):
    data = pd.to_datetime(valor_data, dayfirst=True, errors='coerce')
    if pd.isna(data):
        return data_atual
    return data.to_pydatetime()

def formatar_data_extenso(data):
    meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
             'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
    return f"{data.day} de {meses[data.month-1]} de {data.year}"

data_assinatura = obter_data_assinatura(edital.get('fim_edital', None))

story.append(Paragraph(f"<b>IDENTIFICAÇÃO DA PROPOSTA DE ATENDIMENTO AO EDITAL/CHAMADA PÚBLICA {edital['chamada_publica']}</b>", titulo_style))
story.append(Spacer(1, 0.2*inch))

# Seção I - IDENTIFICAÇÃO DOS FORNECEDORES
secao1_header = [
    [Paragraph("<b>I. IDENTIFICAÇÃO DOS FORNECEDORES</b>", ParagraphStyle('header', parent=titulo_secao_style, alignment=1))]
]
secao1_grupo = [
    [Paragraph("<b>GRUPO FORMAL</b>", ParagraphStyle('header', parent=titulo_secao_style, alignment=1))]
]
secao1_row1 = [
    [Paragraph("<b>1. Nome do Proponente</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>2. CNPJ</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{admin['proponente']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{admin['cnpj_proponente']}", ParagraphStyle('cell', fontSize=7))]
]
secao1_row2 = [
    [Paragraph("<b>3. Endereço</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>4. Município/UF</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{admin['endereco_proponente']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{admin['municipio_proponente']}/{admin['uf_proponente']}", ParagraphStyle('cell', fontSize=7))]
]
secao1_row3 = [
    [Paragraph("<b>5. E-mail</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>6. DDD/Telefone</b>", ParagraphStyle('cell', fontSize=7)),
     Paragraph("<b>7. CEP</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{admin['e-mailp']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{admin['celular_proponente']}", ParagraphStyle('cell', fontSize=7)),
     Paragraph(f"{admin['cep_proponente']}", ParagraphStyle('cell', fontSize=7))]
]
secao1_row4 = [
    [Paragraph("<b>8. N. DAP/CAF Jurídica ou NIS</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>9. Banco</b>", ParagraphStyle('cell', fontSize=7)),
     Paragraph("<b>10. Agência</b>", ParagraphStyle('cell', fontSize=7)),
     Paragraph("<b>11. Conta Corrente</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{admin['caf_juridica']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{admin['banco_proponente']}", ParagraphStyle('cell', fontSize=7)),
     Paragraph(f"{admin['agencia_proponente']}", ParagraphStyle('cell', fontSize=7)),
     Paragraph(f"{admin['conta_proponente']}", ParagraphStyle('cell', fontSize=7))]
]
secao1_row5 = [
    [Paragraph("<b>12. N. Total de Associados</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>13. N. de Associados sem DAP/CAF Física<br/>ou NIS</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
     Paragraph("<b>14. N. de Associados com<br/>DAP/CAF Física ou NIS</b>", ParagraphStyle('cell', fontSize=7, alignment=1))],
    [Paragraph("31", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("3", ParagraphStyle('cell', fontSize=7)),
     Paragraph("28", ParagraphStyle('cell', fontSize=7))]
]
secao1_row6 = [
    [Paragraph("<b>15. Nome do Representante Legal</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>16. CPF</b>", ParagraphStyle('cell', fontSize=7)),
     Paragraph("<b>17. DDD/Telefone</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{admin['representante_proponente']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{admin['cpf_proponente']}", ParagraphStyle('cell', fontSize=7)),
     Paragraph(f"{admin['celular_proponente']}", ParagraphStyle('cell', fontSize=7))]
]
secao1_row7 = [
    [Paragraph("<b>18. Endereço</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>19. Município/UF</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{admin['endereco_proponente']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{admin['municipio_proponente']}/{admin['uf_proponente']}", ParagraphStyle('cell', fontSize=7))]
]

table_secao1_header = Table(secao1_header, colWidths=[7*inch])
table_secao1_header.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#87CEEB')),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
]))

table_secao1_grupo = Table(secao1_grupo, colWidths=[7*inch])
table_secao1_grupo.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 8),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
]))

table_secao1_row1 = Table(secao1_row1, colWidths=[4*inch, 3*inch])
table_secao1_row1.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao1_row2 = Table(secao1_row2, colWidths=[4*inch, 3*inch])
table_secao1_row2.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao1_row3 = Table(secao1_row3, colWidths=[2.5*inch, 2.5*inch, 2*inch])
table_secao1_row3.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao1_row4 = Table(secao1_row4, colWidths=[2*inch, 1.5*inch, 1.5*inch, 2*inch])
table_secao1_row4.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao1_row5 = Table(secao1_row5, colWidths=[2*inch, 2.5*inch, 2.5*inch])
table_secao1_row5.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao1_row6 = Table(secao1_row6, colWidths=[3*inch, 2*inch, 2*inch])
table_secao1_row6.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao1_row7 = Table(secao1_row7, colWidths=[4*inch, 3*inch])
table_secao1_row7.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

story.append(table_secao1_header)
story.append(table_secao1_grupo)
story.append(table_secao1_row1)
story.append(table_secao1_row2)
story.append(table_secao1_row3)
story.append(table_secao1_row4)
story.append(table_secao1_row5)
story.append(table_secao1_row6)
story.append(table_secao1_row7)
story.append(Spacer(1, 0.2*inch))

# Seção II - IDENTIFICAÇÃO DA UNIDADE EXECUTORA
secao2_header = [
    [Paragraph("<b>II. IDENTIFICAÇÃO DA UNIDADE EXECUTORA DO PNAE/FNDE/MEC</b>", ParagraphStyle('header', parent=titulo_secao_style, alignment=1))]
]
table_secao2_header = Table(secao2_header, colWidths=[7*inch])
table_secao2_header.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#87CEEB')),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
]))

secao2_row1 = [
    [Paragraph("<b>1. Nome da Entidade</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>2. CNPJ</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>3. Município/UF</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{edital['nome_executora']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{edital['cnpj_executora']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{edital['municipio_executora']}/{edital.get('uf_executora', '')}", ParagraphStyle('cell', fontSize=7))]
]
secao2_row2 = [
    [Paragraph("<b>4. Endereço</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>5. DDD/Telefone</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{edital['endereco_executora']}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("", ParagraphStyle('cell', fontSize=7))]
]
secao2_row3 = [
    [Paragraph("<b>6. Nome do Representante e E-mail</b>", ParagraphStyle('cell', fontSize=7)), 
     Paragraph("<b>7. CPF</b>", ParagraphStyle('cell', fontSize=7))],
    [Paragraph(f"{edital['gestor_executora']} / {edital.get('e-mail_r_ex', '')}", ParagraphStyle('cell', fontSize=7)), 
     Paragraph(f"{edital['cpf_executora']}", ParagraphStyle('cell', fontSize=7))]
]

table_secao2_row1 = Table(secao2_row1, colWidths=[3*inch, 2*inch, 2*inch])
table_secao2_row1.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao2_row2 = Table(secao2_row2, colWidths=[5*inch, 2*inch])
table_secao2_row2.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

table_secao2_row3 = Table(secao2_row3, colWidths=[5*inch, 2*inch])
table_secao2_row3.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 7),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
]))

story.append(table_secao2_header)
story.append(table_secao2_row1)
story.append(table_secao2_row2)
story.append(table_secao2_row3)
story.append(Spacer(1, 0.3*inch))

# LOOP - MÚLTIPLAS PÁGINAS DE ENVELOPE
envelopes = df_envelope[df_envelope['status_envelope'] == 'SIM']

for idx, envelope in envelopes.iterrows():
    story.append(PageBreak())
    
    # Cabeçalho
    try:
        img = Image("cabecalho.png", width=7*inch, height=1*inch)
        story.append(img)
    except:
        story.append(Paragraph("CABEÇALHO", titulo_style))
    
    story.append(Spacer(1, 0.3*inch))
    
    # À executora (alinhado à esquerda)
    story.append(Paragraph(f"À {edital['nome_executora']}.", esquerda_style))
    story.append(Spacer(1, 0.2*inch))
    
    story.append(Paragraph("COMISSÃO PERMANENTE DE LICITAÇÕES.", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    # Chamada Pública e Anexo (centralizado)
    story.append(Paragraph(f"<b>CHAMADA PÚBLICA {edital['chamada_publica']}</b>", titulo_style))
    story.append(Paragraph(f"<b>ANEXO {envelope['anexo_envelope']}</b>", titulo_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Assunto
    story.append(Paragraph(f"<b>Assunto:</b> {envelope['assunto']}.", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    # Declaração (justificado)
    texto_declaracao = f"A {admin['proponente']} com sede na {admin['endereco_proponente']}, na cidade de {admin['municipio_proponente']}/{admin['uf_proponente']}, inscrita no CNPJ de nº {admin['cnpj_proponente']}, representada neste ato, por sua Presidente a Sra. {admin['representante_proponente']}, brasileira, solteira, agricultora, portadora do Registro de Identidade nº {admin['rg_proponente']}, expedido pela SSP-BA, devidamente inscrita no Cadastro de Pessoas Físicas do Ministério da Fazenda, sob o nº {admin['cpf_proponente']}, {envelope['declaracao']}."
    story.append(Paragraph(texto_declaracao, justificado_style))
    story.append(Spacer(1, 0.4*inch))
    
    # Data (alinhado à esquerda)
    data_str = f"{admin['municipio_proponente']}/{admin['uf_proponente']}, {formatar_data_extenso(data_assinatura)}."
    story.append(Paragraph(data_str, esquerda_style))
    story.append(Spacer(1, 0.5*inch))
    
    # Assinaturas (centralizadas)
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph("--------------------------------------------------", centro_style))
    story.append(Paragraph(admin['proponente'], centro_style))
    story.append(Spacer(1, 0.05*inch))
    story.append(Paragraph(admin['representante_proponente'], centro_style))
    story.append(Spacer(1, 0.05*inch))
    story.append(Paragraph(admin['cpf_proponente'], centro_style))

# ANEXO I - ALIMENTOS (RELAÇÃO DE PRODUTOS)
story.append(PageBreak())

# Cabeçalho
try:
    img = Image("cabecalho.png", width=7*inch, height=1*inch)
    story.append(img)
except:
    story.append(Paragraph("CABEÇALHO", titulo_style))

story.append(Spacer(1, 0.3*inch))

# Cabeçalho do ANEXO 1 - RELAÇÃO DE PRODUTOS
anexo1_header = [
    [Paragraph("<b>III. RELAÇÃO DE PRODUTOS</b>", ParagraphStyle('header', parent=titulo_secao_style, alignment=1))]
]
table_anexo1_header = Table(anexo1_header, colWidths=[7*inch])
table_anexo1_header.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#87CEEB')),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 9),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
]))
story.append(table_anexo1_header)
story.append(Spacer(1, 0.1*inch))

# Filtrar alimentos com status_alimentos = 'Ativo'
alimentos_ativos = df_alimentos[df_alimentos['status_alimentos'] == 'Ativo']

if len(alimentos_ativos) > 0:
    # Cabeçalho da tabela
    tabela_alimentos = [[
        Paragraph("<b>1. Produto</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("<b>2. Unidade</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("<b>3. Quantidade</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("<b>4.1. Preço de<br/>Aquisição<br/>Unitário (R$)</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("<b>4.2. Preço de<br/>Aquisição Total*<br/>(R$)</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("<b>5. Cronograma de<br/>Entrega dos<br/>Alimentos</b>", ParagraphStyle('cell', fontSize=7, alignment=1))
    ]]
    
    item_num = 1
    total_geral_alimentos = 0
    
    for _, alimento in alimentos_ativos.iterrows():
        desc = str(alimento.get('produto', ''))
        unid = str(alimento.get('unidade', ''))
        qtd = converter_quantidade(alimento.get('quantidade', 0))
        
        if pd.isna(qtd):
            qtd = 0.0
        preco = converter_preco(alimento.get('preco', 0))
        
        if pd.isna(preco):
            preco = 0.0
        total_item = qtd * preco
        total_geral_alimentos += total_item
        
        # Obter sazonalidade da guia alimentos
        sazonalidade = str(alimento.get('sazonalidade', 'Ano todo'))
        
        tabela_alimentos.append([
            Paragraph(f"{item_num}. {desc}", ParagraphStyle('cell', fontSize=7)),
            Paragraph(unid, ParagraphStyle('cell', fontSize=7, alignment=1)),
            Paragraph(str(int(qtd)) if qtd == int(qtd) else f"{qtd:.1f}", ParagraphStyle('cell', fontSize=7, alignment=1)),
            Paragraph(f"R$ {preco:.2f}", ParagraphStyle('cell', fontSize=7, alignment=1)),
            Paragraph(f"R$ {total_item:.2f}", ParagraphStyle('cell', fontSize=7, alignment=1)),
            Paragraph(sazonalidade, ParagraphStyle('cell', fontSize=6, alignment=1))
        ])
        item_num += 1
    
    # Adicionar linha de total
    tabela_alimentos.append([
        Paragraph(f"<b>6. Total do Projeto (R$): {total_geral_alimentos:.2f}</b>", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7))
    ])
    
    # Nota sobre preço
    tabela_alimentos.append([
        Paragraph(f"<i>* Preço publicado no Edital N. {edital['chamada_publica']} (o mesmo que consta na chamada pública).</i>", 
                  ParagraphStyle('cell', fontSize=6, alignment=1)),
        Paragraph("", ParagraphStyle('cell', fontSize=6)),
        Paragraph("", ParagraphStyle('cell', fontSize=6)),
        Paragraph("", ParagraphStyle('cell', fontSize=6)),
        Paragraph("", ParagraphStyle('cell', fontSize=6)),
        Paragraph("", ParagraphStyle('cell', fontSize=6))
    ])
    
    # Declaração
    tabela_alimentos.append([
        Paragraph("Declaro estar de acordo com as condições estabelecidas neste projeto e que as informações acima conferem com as condições de fornecimento.", 
                  ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7))
    ])
    
    # Rodapé com assinatura
    tabela_alimentos.append([
        Paragraph("<b>Local e Data</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("<b>Assinatura do Representante do Grupo Formal</b>", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph(f"<b>Telefone/E-mail:</b> {admin['celular_proponente']} / {admin['e-mailp']}", ParagraphStyle('cell', fontSize=7, alignment=1)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7)),
        Paragraph("", ParagraphStyle('cell', fontSize=7))
    ])
    
    # Calcular número de linhas de produtos
    num_produtos = len(alimentos_ativos)
    table_alimentos = Table(tabela_alimentos, colWidths=[1.5*inch, 1.0*inch, 1.0*inch, 1.2*inch, 1.2*inch, 1.1*inch])
    table_alimentos.setStyle(TableStyle([
        # Cabeçalho
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        
        # Linhas de dados (produtos)
        ('ALIGN', (0, 1), (0, num_produtos), 'LEFT'),
        ('ALIGN', (1, 1), (-1, num_produtos), 'CENTER'),
        ('VALIGN', (0, 1), (-1, num_produtos), 'MIDDLE'),
        ('FONTSIZE', (0, 1), (-1, num_produtos), 7),
        
        # Linha Total (mesclar células)
        ('SPAN', (0, num_produtos+1), (-1, num_produtos+1)),
        ('ALIGN', (0, num_produtos+1), (-1, num_produtos+1), 'LEFT'),
        ('FONTNAME', (0, num_produtos+1), (-1, num_produtos+1), 'Helvetica-Bold'),
        
        # Linha da nota (mesclar células)
        ('SPAN', (0, num_produtos+2), (-1, num_produtos+2)),
        ('ALIGN', (0, num_produtos+2), (-1, num_produtos+2), 'CENTER'),
        
        # Linha da declaração (mesclar células)
        ('SPAN', (0, num_produtos+3), (-1, num_produtos+3)),
        ('ALIGN', (0, num_produtos+3), (-1, num_produtos+3), 'CENTER'),
        
        # Linha de assinatura (mesclar colunas)
        ('SPAN', (0, num_produtos+4), (0, num_produtos+4)),
        ('SPAN', (1, num_produtos+4), (3, num_produtos+4)),
        ('SPAN', (4, num_produtos+4), (5, num_produtos+4)),
        ('ALIGN', (0, num_produtos+4), (-1, num_produtos+4), 'CENTER'),
        
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    
    story.append(table_alimentos)
else:
    story.append(Paragraph("Nenhum alimento cadastrado.", normal_style))

story.append(Spacer(1, 0.3*inch))

# LOOP - MÚLTIPLAS PÁGINAS DE CAPA
if 'status_capa' in df_capa.columns:
    capas_ativas = df_capa[df_capa['status_capa'] == 'Ativo']
else:
    capas_ativas = df_capa

for _, capa in capas_ativas.iterrows():
    story.append(PageBreak())

    # Cabeçalho
    try:
        img = Image("cabecalho.png", width=7*inch, height=1*inch)
        story.append(img)
    except:
        story.append(Paragraph("CABEÇALHO", titulo_style))

    story.append(Spacer(1, 0.3*inch))

    # À executora (alinhado à esquerda)
    story.append(Paragraph(f"À {edital['nome_executora']}.", esquerda_style))
    story.append(Spacer(1, 0.15*inch))

    story.append(Paragraph("COMISSÃO PERMANENTE DE LICITAÇÕES.", styles['Normal']))
    story.append(Spacer(1, 0.2*inch))

    # Credenciamento/Chamada Pública (alinhado à esquerda)
    credenciamento = capa.get('credenciamento', 'CREDENCIAMENTO/CHAMADA PÚBLICA')
    story.append(Paragraph(f"CREDENCIAMENTO/CHAMADA PÚBLICA {edital['chamada_publica']}", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))

    # Título da Capa (centralizado - tamanho maior)
    capa_titulo = capa.get('capa', 'CAPA')
    story.append(Paragraph(f"<b>{capa_titulo}</b>", titulo_capa_style))
    story.append(Spacer(1, 0.2*inch))

    titulo_capa = capa.get('titulo_capa', '')
    if titulo_capa:
        story.append(Paragraph(f"<b>{titulo_capa}</b>", titulo_capa_style))
        story.append(Spacer(1, 0.3*inch))

    # Dados da empresa (alinhado à esquerda)
    story.append(Paragraph(f"EMPRESA: {admin['proponente']}", styles['Normal']))
    story.append(Paragraph(
        f"ENDEREÇO: {admin['endereco_proponente']}, {admin['municipio_proponente']}/{admin['uf_proponente']}",
        styles['Normal']
    ))
    story.append(Paragraph(f"CNPJ: {admin['cnpj_proponente']}", styles['Normal']))
    story.append(Spacer(1, 0.5*inch))

    # Assinaturas (centralizadas)
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph(" ", centro_style))
    story.append(Paragraph("--------------------------------------------------", centro_style))
    story.append(Paragraph(admin['proponente'], centro_style))
    story.append(Spacer(1, 0.05*inch))
    story.append(Paragraph(admin['representante_proponente'], centro_style))
    story.append(Spacer(1, 0.05*inch))
    story.append(Paragraph(admin['cpf_proponente'], centro_style))

# Construir PDF
doc.build(story)
print(f"PDF gerado com sucesso: {pdf_path}")