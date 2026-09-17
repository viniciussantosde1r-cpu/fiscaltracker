from io import BytesIO
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    KeepTogether,
    Flowable,
)

from pypdf import PdfReader, PdfWriter
from PIL import Image as PILImage

from database import (
    obter_dados_flash_report,
    calcular_status_prazo_tarefa,
    listar_aprovacoes_encerramento,
    obter_status_encerramento,
    ler_conteudo_documento,
)


# ============================================================
# IDENTIDADE VISUAL
# ============================================================

COR_VERDE = colors.HexColor('#006B3F')
COR_VERDE_ESCURO = colors.HexColor('#005533')
COR_AZUL = colors.HexColor('#17365D')
COR_AZUL_ESCURO = colors.HexColor('#102A43')
COR_TEXTO = colors.HexColor('#1D2924')
COR_TEXTO_SUAVE = colors.HexColor('#66736D')
COR_BORDA = colors.HexColor('#000000')
COR_FUNDO = colors.HexColor('#F4F6F5')
COR_FUNDO_SUAVE = colors.HexColor('#EEF2F0')
COR_OK = colors.HexColor('#006B3F')
COR_ALERTA = colors.HexColor('#B45309')
COR_CRITICO = colors.HexColor('#B91C1C')


# ============================================================
# AUXILIARES
# ============================================================

def texto_seguro(valor):
    if valor is None:
        return ''
    return escape(str(valor))


def texto_ou_traco(valor):
    if valor is None:
        return '-'
    valor = str(valor).strip()
    return valor if valor else '-'


def formatar_moeda(valor):
    valor = valor or 0
    valor_formatado = (
        f'{valor:,.2f}'
        .replace(',', 'X')
        .replace('.', ',')
        .replace('X', '.')
    )
    return f'R$ {valor_formatado}'


def formatar_data(data_texto):
    if not data_texto:
        return '-'
    try:
        ano, mes, dia = str(data_texto)[:10].split('-')
        return f'{dia}/{mes}/{ano}'
    except Exception:
        return str(data_texto)


def formatar_data_hora(data_texto):
    if not data_texto:
        return '-'
    try:
        obj = datetime.fromisoformat(str(data_texto))
        return obj.strftime('%d/%m/%Y %H:%M')
    except Exception:
        return str(data_texto)


def texto_dias_restantes(dias):
    if dias is None:
        return '-'
    if dias < 0:
        atraso = abs(dias)
        return '1 dia em atraso' if atraso == 1 else f'{atraso} dias em atraso'
    if dias == 0:
        return 'Vence hoje'
    if dias == 1:
        return '1 dia restante'
    return f'{dias} dias restantes'


def _nome_seguro_zip(nome):
    nome = Path(str(nome or 'arquivo')).name
    proibidos = '<>:"/\\|?*'
    for caractere in proibidos:
        nome = nome.replace(caractere, '_')
    return nome or 'arquivo'


def _extensao(nome):
    return Path(str(nome or '')).suffix.lower()


# ============================================================
# ESTILOS
# ============================================================

def criar_estilos():
    base = getSampleStyleSheet()
    return {
        'titulo': ParagraphStyle(
            'TituloFlash',
            parent=base['Title'],
            fontName='Helvetica-Bold',
            fontSize=18,
            leading=22,
            textColor=COR_AZUL_ESCURO,
            alignment=TA_LEFT,
            spaceAfter=4,
        ),
        'codigo': ParagraphStyle(
            'CodigoFlash',
            parent=base['Normal'],
            fontName='Helvetica-Bold',
            fontSize=8.5,
            leading=11,
            textColor=COR_VERDE,
            spaceAfter=4,
        ),
        'subtitulo': ParagraphStyle(
            'SubtituloFlash',
            parent=base['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            textColor=COR_TEXTO_SUAVE,
            spaceAfter=8,
        ),
        'secao': ParagraphStyle(
            'SecaoFlash',
            parent=base['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=11.5,
            leading=15,
            textColor=colors.white,
            spaceAfter=0,
        ),
        'subsecao': ParagraphStyle(
            'SubsecaoFlash',
            parent=base['Heading3'],
            fontName='Helvetica-Bold',
            fontSize=10,
            leading=13,
            textColor=COR_AZUL,
            spaceBefore=7,
            spaceAfter=4,
        ),
        'corpo': ParagraphStyle(
            'CorpoFlash',
            parent=base['BodyText'],
            fontName='Helvetica',
            fontSize=9.2,
            leading=13.5,
            textColor=COR_TEXTO,
            alignment=TA_LEFT,
            spaceAfter=6,
        ),
        'pequeno': ParagraphStyle(
            'PequenoFlash',
            parent=base['BodyText'],
            fontName='Helvetica',
            fontSize=7.8,
            leading=10.5,
            textColor=COR_TEXTO,
        ),
        'pequeno_cinza': ParagraphStyle(
            'PequenoCinzaFlash',
            parent=base['BodyText'],
            fontName='Helvetica',
            fontSize=7.5,
            leading=10,
            textColor=COR_TEXTO_SUAVE,
        ),
        'tabela_cabecalho': ParagraphStyle(
            'TabelaCabecalho',
            parent=base['Normal'],
            fontName='Helvetica-Bold',
            fontSize=7.5,
            leading=9.5,
            textColor=colors.white,
            alignment=TA_CENTER,
        ),
        'tabela': ParagraphStyle(
            'TabelaCorpo',
            parent=base['Normal'],
            fontName='Helvetica',
            fontSize=7.5,
            leading=9.5,
            textColor=COR_TEXTO,
        ),
        'tabela_centro': ParagraphStyle(
            'TabelaCentro',
            parent=base['Normal'],
            fontName='Helvetica',
            fontSize=7.5,
            leading=9.5,
            textColor=COR_TEXTO,
            alignment=TA_CENTER,
        ),
        'rodape': ParagraphStyle(
            'RodapeFlash',
            parent=base['Normal'],
            fontName='Helvetica',
            fontSize=7,
            textColor=COR_TEXTO_SUAVE,
            alignment=TA_CENTER,
        ),
    }


def criar_paragrafo_texto(texto, estilo):
    texto = texto_seguro(texto_ou_traco(texto)).replace('\n', '<br/>')
    return Paragraph(texto, estilo)


# ============================================================
# MARCA VETORIAL DE APROVACAO
# ============================================================

class MarcaAprovacao(Flowable):
    def __init__(self, aprovado=True, tamanho=12):
        super().__init__()
        self.aprovado = aprovado
        self.width = tamanho
        self.height = tamanho
        self.tamanho = tamanho

    def draw(self):
        c = self.canv
        t = self.tamanho
        c.saveState()
        c.setLineWidth(1.2)
        c.setStrokeColor(COR_OK if self.aprovado else COR_TEXTO_SUAVE)
        c.circle(t / 2, t / 2, t * 0.42, stroke=1, fill=0)
        if self.aprovado:
            c.setLineWidth(1.6)
            c.line(t * 0.25, t * 0.50, t * 0.43, t * 0.31)
            c.line(t * 0.43, t * 0.31, t * 0.76, t * 0.70)
        c.restoreState()


# ============================================================
# CABECALHO / RODAPE
# ============================================================

def desenhar_cabecalho_rodape(canvas, doc):
    canvas.saveState()
    largura, altura = A4

    canvas.setStrokeColor(COR_AZUL_ESCURO)
    canvas.setLineWidth(0.8)
    canvas.line(18 * mm, altura - 14 * mm, largura - 18 * mm, altura - 14 * mm)

    canvas.setFont('Helvetica-Bold', 7)
    canvas.setFillColor(COR_AZUL_ESCURO)
    canvas.drawString(18 * mm, altura - 11 * mm, 'FISCAL TRACKER - REPORT FISCALIZACOES')

    canvas.setStrokeColor(colors.HexColor('#B8C2BD'))
    canvas.setLineWidth(0.5)
    canvas.line(18 * mm, 13 * mm, largura - 18 * mm, 13 * mm)

    canvas.setFont('Helvetica', 7)
    canvas.setFillColor(COR_TEXTO_SUAVE)
    canvas.drawString(18 * mm, 8 * mm, 'Syngenta Brasil - Tax Controversy')
    canvas.drawRightString(largura - 18 * mm, 8 * mm, f'Pagina {doc.page}')
    canvas.restoreState()


# ============================================================
# COMPONENTES VISUAIS
# ============================================================

def adicionar_secao(elementos, numero, titulo, estilos):
    elementos.append(Spacer(1, 3 * mm))
    barra = Table(
        [[Paragraph(f'{numero}. {texto_seguro(titulo)}', estilos['secao'])]],
        colWidths=[174 * mm],
    )
    barra.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), COR_AZUL_ESCURO),
        ('BOX', (0, 0), (-1, -1), 0.8, COR_BORDA),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    elementos.append(barra)
    elementos.append(Spacer(1, 2.5 * mm))


def tabela_info(linhas, larguras, estilos):
    dados = []
    for linha in linhas:
        row = []
        for indice, valor in enumerate(linha):
            if indice % 2 == 0:
                row.append(Paragraph(f'<b>{texto_seguro(valor)}</b>', estilos['pequeno']))
            else:
                row.append(Paragraph(texto_seguro(texto_ou_traco(valor)), estilos['pequeno']))
        dados.append(row)
    tabela = Table(dados, colWidths=larguras)
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), COR_FUNDO_SUAVE),
        ('BACKGROUND', (2, 0), (2, -1), COR_FUNDO_SUAVE),
        ('BOX', (0, 0), (-1, -1), 0.7, COR_BORDA),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, COR_BORDA),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    return tabela


def _tabela_generica(dados, larguras, estilos, repetir_cabecalho=True):
    tabela = Table(dados, colWidths=larguras, repeatRows=1 if repetir_cabecalho else 0)
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COR_AZUL),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BOX', (0, 0), (-1, -1), 0.7, COR_BORDA),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, COR_BORDA),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, COR_FUNDO]),
    ]))
    return tabela


# ============================================================
# PDF BASE DO REPORT FISCALIZACOES
# ============================================================

def _gerar_pdf_base(fiscalizacao_id):
    dados = obter_dados_flash_report(fiscalizacao_id)
    if not dados:
        raise ValueError('Fiscalizacao nao encontrada.')

    fiscalizacao = dados['fiscalizacao']
    complementares = dados['dados_complementares']
    tributos = dados['tributos']
    areas = dados['areas']
    historico = dados['historico']
    tarefas = dados['tarefas']
    flash = dados['flash_report']
    status_prazo = dados['status_prazo']
    documentos = dados.get('documentos') or []

    aprovacoes = listar_aprovacoes_encerramento(fiscalizacao_id)
    status_encerramento = obter_status_encerramento(fiscalizacao_id)

    (
        id_, codigo, titulo, objeto, empresa, filial, orgao,
        exposicao, probabilidade, fase, data_entrada_fase,
        criado_em, atualizado_em,
    ) = fiscalizacao

    responsavel = complementares.get('responsavel_principal')
    data_recebimento = complementares.get('data_recebimento')
    prazo_resposta = complementares.get('prazo_resposta')
    tarefa_estracta = complementares.get('tarefa_estracta')

    buffer = BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=20 * mm,
        bottomMargin=18 * mm,
        title=f'Report Fiscalizacoes - {codigo}',
        author='Fiscal Tracker',
        subject='Relatorio executivo de fiscalizacao tributaria',
    )
    estilos = criar_estilos()
    elementos = []

    # Capa / cabecalho do documento
    elementos.append(Paragraph('SYNGENTA BRASIL - TAX CONTROVERSY', estilos['codigo']))
    elementos.append(Paragraph('REPORT FISCALIZACOES', estilos['titulo']))
    elementos.append(Paragraph(texto_seguro(codigo), estilos['subsecao']))
    elementos.append(Paragraph(texto_seguro(titulo), estilos['subtitulo']))

    status_documento = flash['status'] if flash else 'Em elaboracao'
    tabela_status = Table([
        [
            Paragraph('<b>Status do relatorio</b>', estilos['pequeno']),
            Paragraph(texto_seguro(status_documento), estilos['pequeno']),
            Paragraph('<b>Fase atual</b>', estilos['pequeno']),
            Paragraph(texto_seguro(fase), estilos['pequeno']),
        ]
    ], colWidths=[34 * mm, 50 * mm, 28 * mm, 62 * mm])
    tabela_status.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), COR_FUNDO),
        ('BOX', (0, 0), (-1, -1), 0.7, COR_BORDA),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, COR_BORDA),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    elementos.append(tabela_status)

    # 1 Identificacao
    adicionar_secao(elementos, 1, 'Identificacao', estilos)
    tributos_texto = ', '.join(tributos) if tributos else '-'
    elementos.append(tabela_info([
        ['Empresa', empresa, 'Filial', filial],
        ['Orgao fiscalizador', orgao, 'Responsavel', responsavel],
        ['Tributos', tributos_texto, 'Probabilidade', probabilidade],
        ['Exposicao', formatar_moeda(exposicao), 'Fase atual', fase],
        ['Recebimento', formatar_data(data_recebimento), 'Prazo', formatar_data(prazo_resposta)],
        ['Tarefa e-Stracta', tarefa_estracta, 'Entrada na fase', formatar_data(data_entrada_fase)],
    ], [31 * mm, 56 * mm, 31 * mm, 56 * mm], estilos))

    # 2 Resumo executivo
    adicionar_secao(elementos, 2, 'Resumo executivo', estilos)
    elementos.append(criar_paragrafo_texto(flash['resumo_executivo'] if flash else '', estilos['corpo']))

    # 3 Objeto e controversia
    adicionar_secao(elementos, 3, 'Objeto e controversia', estilos)
    elementos.append(Paragraph('<b>Objeto / escopo</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(objeto, estilos['corpo']))
    elementos.append(Paragraph('<b>Problema / controversia</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(flash['problema'] if flash else '', estilos['corpo']))
    elementos.append(Paragraph('<b>Causa / fundamento</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(flash['causa'] if flash else '', estilos['corpo']))

    # 4 Acoes
    adicionar_secao(elementos, 4, 'Acoes e encaminhamentos', estilos)
    elementos.append(Paragraph('<b>Acoes necessarias</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(flash['necessarias'] if flash else '', estilos['corpo']))
    elementos.append(Paragraph('<b>Acoes implementadas</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(flash['implementadas'] if flash else '', estilos['corpo']))
    elementos.append(Paragraph('<b>Proximos passos</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(flash['proximos_passos'] if flash else '', estilos['corpo']))

    # 5 Riscos
    adicionar_secao(elementos, 5, 'Riscos, impactos e posicao atual', estilos)
    elementos.append(Paragraph('<b>Riscos e impactos</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(flash['riscos_impactos'] if flash else '', estilos['corpo']))
    elementos.append(Paragraph('<b>Conclusao / posicao atual</b>', estilos['subsecao']))
    elementos.append(criar_paragrafo_texto(flash['conclusao'] if flash else '', estilos['corpo']))

    if status_prazo:
        prazo_quadro = Table([[
            Paragraph('<b>Situacao do prazo</b>', estilos['pequeno']),
            Paragraph(texto_seguro(status_prazo.get('status')), estilos['pequeno']),
            Paragraph('<b>Contagem</b>', estilos['pequeno']),
            Paragraph(texto_seguro(texto_dias_restantes(status_prazo.get('dias_restantes'))), estilos['pequeno']),
        ]], colWidths=[34 * mm, 50 * mm, 28 * mm, 62 * mm])
        prazo_quadro.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), COR_FUNDO),
            ('BOX', (0, 0), (-1, -1), 0.7, COR_BORDA),
            ('INNERGRID', (0, 0), (-1, -1), 0.35, COR_BORDA),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        elementos.append(prazo_quadro)

    # 6 Areas envolvidas
    adicionar_secao(elementos, 6, 'Areas envolvidas', estilos)
    if areas:
        dados_areas = [[
            Paragraph('Area', estilos['tabela_cabecalho']),
            Paragraph('Papel', estilos['tabela_cabecalho']),
            Paragraph('Status', estilos['tabela_cabecalho']),
        ]]
        for area in areas:
            nome_area, tipo, status_area = area
            dados_areas.append([
                Paragraph(texto_seguro(nome_area), estilos['tabela']),
                Paragraph(texto_seguro(tipo), estilos['tabela']),
                Paragraph(texto_seguro(status_area), estilos['tabela_centro']),
            ])
        elementos.append(_tabela_generica(dados_areas, [76 * mm, 58 * mm, 40 * mm], estilos))
    else:
        elementos.append(Paragraph('Nenhuma area vinculada.', estilos['corpo']))

    # 7 Tarefas
    adicionar_secao(elementos, 7, 'Tarefas e pendencias', estilos)
    if tarefas:
        dados_tarefas = [[
            Paragraph('Tarefa', estilos['tabela_cabecalho']),
            Paragraph('Responsavel', estilos['tabela_cabecalho']),
            Paragraph('Prazo', estilos['tabela_cabecalho']),
            Paragraph('Status', estilos['tabela_cabecalho']),
        ]]
        for tarefa in tarefas:
            tarefa_id, tarefa_titulo, descricao, resp, prazo, status, concluida_em, criado_t, atualizado_t = tarefa
            prazo_status = calcular_status_prazo_tarefa(prazo, status)
            titulo_tarefa = texto_seguro(tarefa_titulo)
            if descricao:
                titulo_tarefa += f'<br/><font size="7" color="#66736D">{texto_seguro(descricao)}</font>'
            dados_tarefas.append([
                Paragraph(titulo_tarefa, estilos['tabela']),
                Paragraph(texto_seguro(texto_ou_traco(resp)), estilos['tabela']),
                Paragraph(texto_seguro(formatar_data(prazo)), estilos['tabela_centro']),
                Paragraph(texto_seguro(f"{status} - {prazo_status['status_prazo']}"), estilos['tabela']),
            ])
        elementos.append(_tabela_generica(dados_tarefas, [72 * mm, 42 * mm, 26 * mm, 34 * mm], estilos))
    else:
        elementos.append(Paragraph('Nenhuma tarefa cadastrada.', estilos['corpo']))

    # 8 Linha do tempo
    adicionar_secao(elementos, 8, 'Linha do tempo', estilos)
    if historico:
        dados_hist = [[
            Paragraph('Data', estilos['tabela_cabecalho']),
            Paragraph('Fase', estilos['tabela_cabecalho']),
            Paragraph('Evento / observacao', estilos['tabela_cabecalho']),
        ]]
        # O banco retorna mais recente primeiro; no relatorio usamos cronologia crescente.
        for item in reversed(historico):
            hist_id, hist_fase, hist_data, hist_obs, hist_criado = item
            dados_hist.append([
                Paragraph(texto_seguro(formatar_data(hist_data)), estilos['tabela_centro']),
                Paragraph(texto_seguro(hist_fase), estilos['tabela']),
                Paragraph(texto_seguro(texto_ou_traco(hist_obs)), estilos['tabela']),
            ])
        elementos.append(_tabela_generica(dados_hist, [28 * mm, 47 * mm, 99 * mm], estilos))
    else:
        elementos.append(Paragraph('Nenhum evento registrado.', estilos['corpo']))

    # 9 Governanca e aprovadores
    adicionar_secao(elementos, 9, 'Aprovacoes e encerramento', estilos)
    if not aprovacoes:
        elementos.append(Paragraph('Nenhum aprovador de encerramento cadastrado.', estilos['corpo']))
    else:
        dados_aprov = [[
            '',
            Paragraph('Area / papel', estilos['tabela_cabecalho']),
            Paragraph('Aprovador', estilos['tabela_cabecalho']),
            Paragraph('Obrigatoriedade', estilos['tabela_cabecalho']),
            Paragraph('Ciencia', estilos['tabela_cabecalho']),
        ]]
        for item in aprovacoes:
            ciente = item['status'] == 'Ciente'
            area_papel = item['area']
            if item.get('papel'):
                area_papel += f" / {item['papel']}"
            ciencia = formatar_data_hora(item.get('ciente_em')) if ciente else 'Pendente'
            dados_aprov.append([
                MarcaAprovacao(ciente, 12),
                Paragraph(texto_seguro(area_papel), estilos['tabela']),
                Paragraph(texto_seguro(item['nome']), estilos['tabela']),
                Paragraph('Essencial' if item['essencial'] else 'Aprovador', estilos['tabela_centro']),
                Paragraph(texto_seguro(ciencia), estilos['tabela_centro']),
            ])
        tabela_aprov = _tabela_generica(dados_aprov, [14 * mm, 50 * mm, 43 * mm, 29 * mm, 38 * mm], estilos)
        tabela_aprov.setStyle(TableStyle([
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ]))
        elementos.append(tabela_aprov)

        elementos.append(Spacer(1, 2 * mm))
        if fase == 'Encerrada':
            elementos.append(Paragraph(
                '<b>Status de encerramento:</b> Fiscalizacao encerrada com a trilha de ciencia registrada.',
                estilos['corpo'],
            ))
        elif status_encerramento.get('pode_encerrar'):
            elementos.append(Paragraph(
                '<b>Status de encerramento:</b> Todos os aprovadores essenciais registraram ciencia; encerramento liberado.',
                estilos['corpo'],
            ))
        else:
            elementos.append(Paragraph(
                '<b>Status de encerramento:</b> Encerramento ainda bloqueado por aprovacao essencial pendente.',
                estilos['corpo'],
            ))

    # 10 Indice documental
    adicionar_secao(elementos, 10, 'Documentos e evidencias', estilos)
    if documentos:
        dados_docs = [[
            Paragraph('No.', estilos['tabela_cabecalho']),
            Paragraph('Tipo', estilos['tabela_cabecalho']),
            Paragraph('Arquivo', estilos['tabela_cabecalho']),
            Paragraph('Data', estilos['tabela_cabecalho']),
            Paragraph('Integridade SHA-256', estilos['tabela_cabecalho']),
        ]]
        for indice, doc in enumerate(documentos, start=1):
            hash_curto = doc.get('hash_sha256') or '-'
            dados_docs.append([
                Paragraph(f'{indice:02d}', estilos['tabela_centro']),
                Paragraph(texto_seguro(doc.get('tipo_documento')), estilos['tabela']),
                Paragraph(texto_seguro(doc.get('nome_original')), estilos['tabela']),
                Paragraph(texto_seguro(formatar_data(doc.get('data_documento') or doc.get('criado_em'))), estilos['tabela_centro']),
                Paragraph(texto_seguro(hash_curto), estilos['pequeno_cinza']),
            ])
        elementos.append(_tabela_generica(dados_docs, [12 * mm, 35 * mm, 55 * mm, 25 * mm, 47 * mm], estilos))
        elementos.append(Spacer(1, 2 * mm))
        elementos.append(Paragraph(
            'Os arquivos originais permanecem preservados no Dossie Final. PDFs e imagens compativeis podem ser incorporados ao final deste documento consolidado.',
            estilos['pequeno_cinza'],
        ))
    else:
        elementos.append(Paragraph('Nenhum documento ou evidencia anexado ao caso.', estilos['corpo']))

    elementos.append(Spacer(1, 5 * mm))
    elementos.append(Paragraph(
        f'Documento gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")}.',
        estilos['pequeno_cinza'],
    ))

    documento.build(
        elementos,
        onFirstPage=desenhar_cabecalho_rodape,
        onLaterPages=desenhar_cabecalho_rodape,
    )
    return buffer.getvalue(), dados, aprovacoes


# ============================================================
# INCORPORACAO DE ANEXOS COMPATIVEIS
# ============================================================

def _imagem_para_pdf(conteudo, nome='imagem'):
    imagem = PILImage.open(BytesIO(conteudo))
    if imagem.mode not in ('RGB', 'L'):
        fundo = PILImage.new('RGB', imagem.size, 'white')
        if imagem.mode == 'RGBA':
            fundo.paste(imagem, mask=imagem.split()[-1])
        else:
            fundo.paste(imagem.convert('RGB'))
        imagem = fundo
    elif imagem.mode == 'L':
        imagem = imagem.convert('RGB')

    saida = BytesIO()
    imagem.save(saida, format='PDF', resolution=150.0)
    return saida.getvalue()


def _pagina_separadora_anexo(indice, documento):
    estilos = criar_estilos()
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=20 * mm,
        bottomMargin=18 * mm,
    )
    elementos = [
        Spacer(1, 45 * mm),
        Paragraph('ANEXO DOCUMENTAL', estilos['codigo']),
        Paragraph(f'Anexo {indice:02d}', estilos['titulo']),
        Spacer(1, 4 * mm),
        tabela_info([
            ['Tipo', documento.get('tipo_documento'), 'Data', formatar_data(documento.get('data_documento') or documento.get('criado_em'))],
            ['Arquivo', documento.get('nome_original'), 'Tamanho', f"{(documento.get('tamanho_bytes') or 0) / 1024:.1f} KB"],
        ], [26 * mm, 70 * mm, 24 * mm, 54 * mm], estilos),
        Spacer(1, 4 * mm),
        Paragraph(
            '<b>SHA-256</b><br/>' + texto_seguro(documento.get('hash_sha256') or '-'),
            estilos['pequeno'],
        ),
    ]
    doc.build(elementos, onFirstPage=desenhar_cabecalho_rodape)
    return buffer.getvalue()


def _consolidar_com_anexos(pdf_base, documentos):
    writer = PdfWriter()
    base_reader = PdfReader(BytesIO(pdf_base))
    for pagina in base_reader.pages:
        writer.add_page(pagina)

    anexos_incorporados = []
    anexos_nao_incorporados = []

    for indice, documento in enumerate(documentos, start=1):
        ext = _extensao(documento.get('nome_original'))
        if ext not in {'.pdf', '.png', '.jpg', '.jpeg'}:
            anexos_nao_incorporados.append(documento)
            continue

        try:
            conteudo = ler_conteudo_documento(documento['id'])
            separador = PdfReader(BytesIO(_pagina_separadora_anexo(indice, documento)))
            for pagina in separador.pages:
                writer.add_page(pagina)

            if ext == '.pdf':
                leitor_anexo = PdfReader(BytesIO(conteudo))
            else:
                leitor_anexo = PdfReader(BytesIO(_imagem_para_pdf(conteudo, documento.get('nome_original'))))

            for pagina in leitor_anexo.pages:
                writer.add_page(pagina)

            anexos_incorporados.append(documento)

        except Exception:
            # O Report Fiscalizacoes principal continua valido. O arquivo original
            # permanece no ZIP e no indice documental.
            anexos_nao_incorporados.append(documento)

    saida = BytesIO()
    writer.write(saida)
    return saida.getvalue(), anexos_incorporados, anexos_nao_incorporados


# ============================================================
# API PUBLICA
# ============================================================

def gerar_flash_report_pdf(fiscalizacao_id, incorporar_anexos=True):
    """Gera o Report Fiscalizacoes consolidado em memoria e retorna bytes."""
    pdf_base, dados, aprovacoes = _gerar_pdf_base(fiscalizacao_id)
    documentos = dados.get('documentos') or []

    if not incorporar_anexos or not documentos:
        return pdf_base

    pdf_final, _, _ = _consolidar_com_anexos(pdf_base, documentos)
    return pdf_final


def gerar_dossie_final_zip(fiscalizacao_id):
    """
    Gera um pacote ZIP contendo:
      - Report Fiscalizacoes consolidado;
      - todos os documentos originais que passam na verificacao de integridade;
      - manifesto textual com hashes e eventuais falhas de leitura.
    """
    pdf_base, dados, aprovacoes = _gerar_pdf_base(fiscalizacao_id)
    documentos = dados.get('documentos') or []
    pdf_final, incorporados, nao_incorporados = _consolidar_com_anexos(pdf_base, documentos)

    fiscalizacao = dados['fiscalizacao']
    codigo = fiscalizacao[1]
    titulo = fiscalizacao[2]

    buffer_zip = BytesIO()
    manifest = [
        'FISCAL TRACKER - DOSSIÊ FINAL',
        f'Fiscalizacao: {codigo} - {titulo}',
        f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}',
        '',
        'DOCUMENTOS:',
    ]

    with ZipFile(buffer_zip, 'w', ZIP_DEFLATED) as zf:
        zf.writestr(f'Report_Fiscalizacao_Final_{codigo}.pdf', pdf_final)

        nomes_usados = set()
        for indice, documento in enumerate(documentos, start=1):
            nome_base = _nome_seguro_zip(documento.get('nome_original'))
            nome_zip = f'Evidencias/{indice:02d}_{nome_base}'
            if nome_zip.lower() in nomes_usados:
                nome_zip = f'Evidencias/{indice:02d}_{documento["id"]}_{nome_base}'
            nomes_usados.add(nome_zip.lower())

            try:
                conteudo = ler_conteudo_documento(documento['id'])
                zf.writestr(nome_zip, conteudo)
                situacao = 'OK'
            except Exception as erro:
                situacao = f'ERRO AO LER: {erro}'

            manifest.append(
                f"{indice:02d} | {documento.get('tipo_documento')} | "
                f"{documento.get('nome_original')} | SHA-256: {documento.get('hash_sha256')} | {situacao}"
            )

        manifest.extend([
            '',
            f'Anexos incorporados ao PDF consolidado: {len(incorporados)}',
            f'Anexos apenas preservados no ZIP/indice: {len(nao_incorporados)}',
            '',
            'APROVACOES:',
        ])
        for item in aprovacoes:
            manifest.append(
                f"{item['status']} | {item['area']} | {item['papel']} | "
                f"{item['nome']} | Essencial: {'Sim' if item['essencial'] else 'Nao'} | "
                f"Ciencia: {formatar_data_hora(item.get('ciente_em'))}"
            )

        zf.writestr('Manifesto_Dossie.txt', '\n'.join(manifest).encode('utf-8'))

    return buffer_zip.getvalue()
