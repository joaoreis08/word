from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn  
import pandas as pd
from docx.shared import Pt, Inches
from docx.shared import RGBColor
from datetime import datetime


cores_por_tema = {
    "CONHECIMENTO E INOVAÇÃO": "#4400FF",  # Azul
    "SAÚDE E QUALIDADE DE VIDA": "#ED282C",  # Vermelho
    "SEGURANÇA E CIDADANIA": "#FFB000",  # Amarelo
    "DESENVOLVIMENTO SUSTENTÁVEL": "#87D200",  # Verde
    "Gestão, Transparência e Participação": "#002060"  # Azul escuro
}

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)



def set_paragraph_background(paragraph, color):
    """
    Define a cor de fundo (background) para um parágrafo.
    :param paragraph: Objeto do parágrafo (docx.paragraph.Paragraph)
    :param color: Código hexadecimal para a cor de fundo (ex.: 'FFFF00' para amarelo)
    """
    # Obtém o elemento XML subjacente do parágrafo
    p = paragraph._p
    pPr = p.get_or_add_pPr()  # Adiciona ou obtém as propriedades do parágrafo

    # Cria um elemento <w:shd> para aplicar a cor de fundo
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')  # Define o preenchimento como "clear"
    shd.set(qn('w:color'), 'auto')  # Define a cor do texto como padrão (automática)
    shd.set(qn('w:fill'), color)  # Configura o preenchimento do fundo com a cor escolhida (hexadecimal)

    # Adiciona o elemento <w:shd> nas propriedades do parágrafo
    pPr.append(shd)

# Carregar o arquivo Excel
df = pd.read_excel('Iniciativas - RGS 2025.1 - Extração Painel de Controle copy.xlsx',skiprows=1)

# Selecionar e renomear as colunas
colunas = ['Órgão', 'Iniciativa', 'Status Informado', 'Ação', 'Programa',
           'Início Realizado', 'Término Realizado', 'RGS 2025.1 - GGGE', 'Localização Geográfica','Objetivo Estratégico']
df2 = df[colunas]

df2.rename(columns={
    'Órgão': 'Orgao',
    'Iniciativa': 'Iniciativa',
    'Status Informado': 'Status_Informado',
    'Ação': 'Acao',
    'Programa': 'Programa',
    'Início Realizado': 'Inicio_Realizado',
    'Término Realizado': 'Termino_Realizado',
    'RGS 2025.1 - GGGE': 'RGS_2025_GGGE',
    'Localização Geográfica': 'Localizacao_Geografica',
    'Objetivo Estratégico':'Objetivo_Estrategico'
}, inplace=True)

# Converter as colunas de datas
df2[['Inicio_Realizado', 'Termino_Realizado']] = df2[['Inicio_Realizado', 'Termino_Realizado']].apply(
    lambda x: pd.to_datetime(x, errors='coerce', dayfirst=True)
)

# Criação do documento
doc = Document()


for idx, row in enumerate(df2.itertuples()):
    if idx > 0:
        doc.add_paragraph('\n')  # Espaço entre blocos

    cor = cores_por_tema[row.Objetivo_Estrategico]

    # --- ORGAO ---
    p_orgao = doc.add_paragraph()
    p_orgao.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_orgao.paragraph_format.space_before = Pt(0)
    p_orgao.paragraph_format.space_after = Pt(0)
    run = p_orgao.add_run(f'{row.Orgao}')
    run.font.name = 'Gilroy ExtraBold'
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 32, 96)
    set_paragraph_background(p_orgao, 'D3D3D3')

    # --- PROGRAMA ---
    p_programa = doc.add_paragraph()
    p_programa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_programa.paragraph_format.space_before = Pt(0)
    p_programa.paragraph_format.space_after = Pt(0)
    run = p_programa.add_run(f'{row.Programa}')
    run.font.name = 'Gilroy Light'
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_paragraph_background(p_programa, cor)

    # --- ACAO ---
    p_acao = doc.add_paragraph()
    p_acao.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_acao.paragraph_format.space_before = Pt(0)
    p_acao.paragraph_format.space_after = Pt(0)
    run = p_acao.add_run(f'{row.Acao}')
    run.font.name = 'Gilroy Light'
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_paragraph_background(p_acao, cor)

    # Espaço menor que uma linha entre Acao e Iniciativa
    doc.add_paragraph()

    status = 'imagens\concluído.png' if row.Status_Informado == 'CONCLUÍDO' else 'imagens\em_excecucao.png'
    status_texto = 'Data de Entrega:' if row.Status_Informado == 'CONCLUÍDO' else 'Data de Início:'
    prazo = row.Termino_Realizado if row.Status_Informado == 'CONCLUÍDO' else row.Inicio_Realizado
    localizacao = 'imagens\localização.png'
    calendario = 'imagens\calendário.png'
    
    #criar tabela
    table = doc.add_table(rows=4, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    #1 linha
    cell = table.cell(0, 0)
    cell_merge = cell.merge(table.cell(0, 4))
    paragraph = cell_merge.paragraphs[0]
    paragraph.alignment = 1  # Center
    run = paragraph.add_run(str(row.Iniciativa))
    run.font.name = 'Gilroy ExtraBold'
    run.font.size = Pt(10)
    run.bold = True
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_background(cell, 'D3D3D3')


    #2 linha
    #linha 2, coluna 1
    cell_status = table.cell(1, 0)
    cell_status_merge = cell_status.merge(table.cell(1, 1))
    par_status = cell_status_merge.paragraphs[0]
    run_status = par_status.add_run()
    run_status.add_picture(status, width=Inches(0.17))
    run_status.add_text(f'  Status:  {row.Status_Informado}')
    run_status.font.name = 'Gilroy Light'
    run_status.font.size = Pt(9)
    run_status.font.color.rgb = RGBColor(0, 0, 0)

    #linha 2,coluna 3

    if pd.notnull(prazo):
        data_texto = prazo.strftime('%d/%m/%Y')
    else:
        data_texto = ''

    cell_status2 = table.cell(1,2)
    cell_status_merge2 = cell_status2.merge(table.cell(1, 3))
    par_status2 = cell_status_merge2.paragraphs[0]
    run_status2 = par_status2.add_run()
    run_status2.add_picture(calendario, width=Inches(0.17))
    run_status2.add_text(f'  {status_texto}')

    cell_status3 = table.cell(1,4)
    par_status3 = cell_status3.paragraphs[0]
    run_status3 = par_status3.add_run()
    run_status3.add_text(data_texto)
    




    #linha 3,coluna 1
    icone_localizacao = table.cell(2,0).paragraphs[0]
    icone_localizacao = icone_localizacao.add_run()
    icone_localizacao.add_picture(localizacao,width=Inches(0.17))
    #linha3, coluna2
    table.cell(2,1).text = 'Municípios Atendidos: '

    #linha3, coluna 3
    cell_loc = table.cell(2,2)
    cell_loc_merged = cell_loc.merge(table.cell(2,4))
    par_loc = cell_loc_merged.paragraphs[0]
    par_loc.add_run("" if pd.isnull(row.Localizacao_Geografica) else str(row.Localizacao_Geografica))

    #linha 4, coluna 1
    cell2 = table.cell(3, 0)
    cell_merge2 = cell2.merge(table.cell(3, 4))
    paragraph2 = cell_merge2.paragraphs[0]
    run2 = paragraph2.add_run(str(row.RGS_2025_GGGE)) 


doc.save('teste.docx')