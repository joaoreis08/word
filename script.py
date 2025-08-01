from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn  
import pandas as pd
from docx.shared import RGBColor


cores_por_tema = {
    "CONHECIMENTO E INOVAÇÃO": "#4400FF",  # Azul
    "SAÚDE E QUALIDADE DE VIDA": "#ED282C",  # Vermelho
    "SEGURANÇA E CIDADANIA": "#FFB000",  # Amarelo
    "DESENVOLVIMENTO SUSTENTÁVEL": "#87D200",  # Verde
    "Gestão, Transparência e Participação": "#002060"  # Azul escuro
}



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
df = pd.read_excel('Iniciativas - RGS 2025.1 - Extração Painel de Controle.xlsx')

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
df2.loc[:, ['Inicio_Realizado', 'Termino_Realizado']] = df2[['Inicio_Realizado', 'Termino_Realizado']].apply(
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

    # --- INICIATIVA ---
    p_iniciativa = doc.add_paragraph()
    p_iniciativa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_iniciativa.add_run(f'{row.Iniciativa}')
    run.font.name = 'Gilroy ExtraBold'
    run.font.size = Pt(10)
    run.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    set_paragraph_background(p_iniciativa, 'D3D3D3')

    # --- STATUS E DATAS ---
    p_status = doc.add_paragraph()
    p_status.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Tabulação
    tab_stop = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), '8000')
    tab_stop.append(tab)
    p_status._p.get_or_add_pPr().append(tab_stop)

    font_name_status = 'Neutro Thin'

    if row.Status_Informado == 'CONCLUÍDO':
        run_status = p_status.add_run(f"✅Status: {row.Status_Informado}")
        run_status.font.name = font_name_status
        run_status.font.size = Pt(9)
        p_status.add_run("\t")
        termino_formatado = row.Termino_Realizado.strftime('%d/%m/%Y') if pd.notna(row.Termino_Realizado) else ''
        run_date = p_status.add_run(f"📅 Data de Término: {termino_formatado}")
        run_date.font.name = font_name_status
        run_date.font.size = Pt(9)
    else:
        run_status = p_status.add_run(f"🔄 Status: {row.Status_Informado}")
        run_status.font.name = font_name_status
        run_status.font.size = Pt(9)
        p_status.add_run("\t")
        inicio_formatado = row.Inicio_Realizado.strftime('%d/%m/%Y') if pd.notna(row.Inicio_Realizado) else ''
        run_date = p_status.add_run(f"📅 Data de Início: {inicio_formatado}")
        run_date.font.name = font_name_status
        run_date.font.size = Pt(9)

    # --- MUNICÍPIOS ATENDIDOS ---
    p_municipios = doc.add_paragraph(f'📍 Municípios Atendidos:\t \t{row.Localizacao_Geografica}')
    run = p_municipios.runs[0]
    run.font.name = 'Neutro'
    run.font.size = Pt(10)

    # --- RGS 2025 GGGE ---
    p_rgs = doc.add_paragraph()
    run = p_rgs.add_run(f'{row.RGS_2025_GGGE}')
    run.font.name = 'Neutro'
    run.font.size = Pt(9)


# Salvar o documento
doc.save("teste.docx")  