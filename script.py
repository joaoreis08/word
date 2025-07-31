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
df = pd.read_excel('Iniciativas - RGS 2025.1 - Extração Painel de Controle.xlsx', skiprows=1)

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
        doc.add_paragraph('\n')  # Adiciona espaçamento entre os órgãos

    cor = cores_por_tema[row.Objetivo_Estrategico]
    # Adicionar título com fundo
    title = doc.add_heading(f'{row.Orgao}', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_background(title, cor)  # Fundo cinza
    style = doc.styles['Heading 1']  # Aqui, passe o nome do estilo como string
    font = style.font
    font.name = 'Gilroy Extrabold'  # Não é chamada, é uma atribuição
    font.size = Pt(12)
    font.color.rgb = RGBColor(255, 255, 255)


    # Adicionar um parágrafo para Programa e Ação juntos
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Alinhamento centralizado
    run = p.add_run(f'{row.Programa}\n{row.Acao}')  # Ambos os textos no mesmo parágrafo
    font = run.font
    font.name = 'Gilroy Light'
    font.size = Pt(12)
    set_paragraph_background(p, cor)  # Fundo vermelho para o parágrafo inteiro
    
    # Evita fundo em linhas extras
    doc.add_paragraph()  # Adiciona linha em branco sem fundo

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p3.add_run(f'{row.Iniciativa}')
    font = run.font
    font.name = 'Gilroy-ExtraBold'
    font.size = Pt(10)
    font.color.rgb = RGBColor(0, 0, 0)  # Cor preta

    set_paragraph_background(p3, 'D3D3D3')

    # Parágrafo onde "Status" e "Data de Início/Término" aparecem
    p4 = doc.add_paragraph()

    # Configurar tabulação
    tab_stop = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')  # Define como alinhamento à direita
    tab.set(qn('w:pos'), '8000')  # Define a posição da tabulação (8000 twips)
    tab_stop.append(tab)
    p4._p.get_or_add_pPr().append(tab_stop)

    # Adicionar datas com o formato dd/mm/aaaa
    if row.Status_Informado == 'CONCLUÍDO':
        # Adicionar "Status:"
        run_status = p4.add_run(f"✅Status: {row.Status_Informado}")
        run_status.font.size = Pt(9)
        p4.add_run("\t")

        termino_formatado = row.Termino_Realizado.strftime('%d/%m/%Y') if pd.notna(row.Termino_Realizado) else ''
        run_date = p4.add_run(f" 📅 Data de Término: {termino_formatado}")
        run_date.font.size = Pt(9)
    else:
        # Adicionar "Status:"
        run_status = p4.add_run(f"🔄 Status: {row.Status_Informado}")
        run_status.font.size = Pt(9)
        p4.add_run("\t")
        inicio_formatado = row.Inicio_Realizado.strftime('%d/%m/%Y') if pd.notna(row.Inicio_Realizado) else ''
        run_date = p4.add_run(f" 📅 Data de Início: {inicio_formatado}")
        run_date.font.size = Pt(9)

    # Adicionar outros textos
    p5 = doc.add_paragraph(f'📍 Municípios Atendidos:\t \t{row.Localizacao_Geografica}')
    p6 = doc.add_paragraph()
    run = p6.add_run(f'{row.RGS_2025_GGGE}')
    run.font.size = Pt(9)

# Salvar o documento
doc.save("teste.docx")  