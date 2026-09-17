import hashlib
import mimetypes
import os
import sqlite3
import uuid
from pathlib import Path
from datetime import date, datetime
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
UPLOAD_DIR = BASE_DIR / 'uploads'
DB_PATH = DATA_DIR / 'fiscal_tracker.db'
FASES_FISCALIZACAO = ['Intimação recebida', 'Em análise interna', 'Resposta protocolada', 'Auto de infração', 'Recurso', 'Julgamento', 'Encerrada']
STATUS_TAREFAS = ['Pendente', 'Em andamento', 'Concluída']
STATUS_FLASH_REPORT = ['Em elaboração', 'Em revisão', 'Finalizado']
PROBABILIDADES = ['Provável', 'Possível', 'Remota']
TIPOS_DOCUMENTO = ['Termo de Fiscalização', 'Intimação', 'Resposta Protocolada', 'Auto de Infração', 'Recurso', 'Parecer', 'E-mail', 'Evidência', 'Planilha', 'Comprovante', 'Outros']
EXTENSOES_DOCUMENTOS_PERMITIDAS = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.csv', '.txt', '.msg', '.eml', '.png', '.jpg', '.jpeg'}
TAMANHO_MAXIMO_DOCUMENTO = 25 * 1024 * 1024

def conectar():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.execute('PRAGMA foreign_keys = ON')
    conn.execute('PRAGMA busy_timeout = 10000')
    conn.execute('PRAGMA journal_mode = WAL')
    conn.execute('PRAGMA synchronous = NORMAL')
    return conn

def _agora():
    return datetime.now().isoformat(timespec='seconds')

def _hoje():
    return date.today().isoformat()

def _texto(valor):
    if valor is None:
        return ''
    return str(valor).strip()

def validar_data_iso(valor):
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    valor = str(valor).strip()
    if not valor:
        return None
    try:
        date.fromisoformat(valor)
    except ValueError:
        raise ValueError('Data inválida. Utilize o formato AAAA-MM-DD.')
    return valor

def _validar_texto_obrigatorio(valor, campo):
    valor = _texto(valor)
    if not valor:
        raise ValueError(f'Informe {campo}.')
    return valor

def _validar_exposicao(valor):
    try:
        valor = float(valor or 0)
    except (TypeError, ValueError):
        raise ValueError('A exposição deve ser numérica.')
    if valor < 0:
        raise ValueError('A exposição não pode ser negativa.')
    return valor

def _validar_probabilidade(valor):
    if valor not in PROBABILIDADES:
        raise ValueError('Probabilidade inválida.')
    return valor

def _validar_fase(valor):
    if valor not in FASES_FISCALIZACAO:
        raise ValueError('Fase inválida.')
    return valor

def _validar_status_tarefa(valor):
    if valor not in STATUS_TAREFAS:
        raise ValueError('Status de tarefa inválido.')
    return valor

def _validar_status_flash(valor):
    if valor not in STATUS_FLASH_REPORT:
        raise ValueError('Status do Flash Report inválido.')
    return valor

def _validar_intervalo_datas(data_inicial, data_final):
    if data_inicial and data_final and (data_final < data_inicial):
        raise ValueError('A data final não pode ser anterior à data inicial.')

def _buscar_estado_fiscalizacao(cursor, fiscalizacao_id):
    cursor.execute('\n        SELECT\n            id,\n            codigo,\n            fase,\n            COALESCE(\n                arquivada,\n                0\n            )\n\n        FROM fiscalizacoes\n\n        WHERE id = ?\n        ', (fiscalizacao_id,))
    resultado = cursor.fetchone()
    if not resultado:
        raise ValueError('Fiscalização não encontrada.')
    return {'id': resultado[0], 'codigo': resultado[1], 'fase': resultado[2], 'arquivada': bool(resultado[3])}

def _garantir_nao_arquivada(cursor, fiscalizacao_id):
    estado = _buscar_estado_fiscalizacao(cursor, fiscalizacao_id)
    if estado['arquivada']:
        raise ValueError('A fiscalização está arquivada. Reative o registro antes de realizar esta operação.')
    return estado

def _garantir_operacional(cursor, fiscalizacao_id):
    estado = _garantir_nao_arquivada(cursor, fiscalizacao_id)
    if estado['fase'] == 'Encerrada':
        raise ValueError('A fiscalização está encerrada.')
    return estado

def _registrar_historico(cursor, fiscalizacao_id, fase, observacao, data_evento=None):
    data_evento = validar_data_iso(data_evento) or _hoje()
    cursor.execute('\n        INSERT INTO historico_fiscalizacao (\n            fiscalizacao_id,\n            fase,\n            data,\n            observacao\n        )\n\n        VALUES (?, ?, ?, ?)\n        ', (fiscalizacao_id, fase, data_evento, _texto(observacao)))

def _gerar_codigo_cursor(cursor):
    cursor.execute("\n        SELECT codigo\n\n        FROM fiscalizacoes\n\n        WHERE codigo LIKE 'FT-%'\n\n        ORDER BY id DESC\n\n        LIMIT 1\n        ")
    resultado = cursor.fetchone()
    if not resultado:
        return 'FT-0001'
    ultimo_codigo = resultado[0]
    try:
        numero = int(ultimo_codigo.replace('FT-', ''))
    except Exception:
        cursor.execute('\n            SELECT MAX(id)\n            FROM fiscalizacoes\n            ')
        resultado_id = cursor.fetchone()
        numero = resultado_id[0] or 0
    return f'FT-{numero + 1:04d}'

def _validar_tipo_documento(tipo_documento):
    if tipo_documento not in TIPOS_DOCUMENTO:
        raise ValueError('Tipo de documento inválido.')
    return tipo_documento

def _validar_nome_documento(nome_original):
    nome_original = _validar_texto_obrigatorio(nome_original, 'o nome do arquivo')
    if '/' in nome_original or '\\' in nome_original or '\x00' in nome_original:
        raise ValueError('Nome de arquivo inválido.')
    nome_limpo = Path(nome_original).name
    if nome_limpo != nome_original:
        raise ValueError('Nome de arquivo inválido.')
    if nome_limpo in {'.', '..'}:
        raise ValueError('Nome de arquivo inválido.')
    extensao = Path(nome_limpo).suffix.lower()
    if extensao not in EXTENSOES_DOCUMENTOS_PERMITIDAS:
        formatos = ', '.join(sorted(EXTENSOES_DOCUMENTOS_PERMITIDAS))
        raise ValueError(f'Formato de arquivo não permitido. Formatos aceitos: {formatos}.')
    return (nome_limpo, extensao)

def _validar_conteudo_documento(conteudo):
    if conteudo is None:
        raise ValueError('Nenhum arquivo foi informado.')
    if isinstance(conteudo, bytearray):
        conteudo = bytes(conteudo)
    if not isinstance(conteudo, bytes):
        raise ValueError('Conteúdo de arquivo inválido.')
    tamanho = len(conteudo)
    if tamanho == 0:
        raise ValueError('O arquivo está vazio.')
    if tamanho > TAMANHO_MAXIMO_DOCUMENTO:
        raise ValueError('O arquivo excede o limite de 25 MB.')
    return (conteudo, tamanho)

def _calcular_sha256(conteudo):
    return hashlib.sha256(conteudo).hexdigest()

def _mime_type_seguro(nome_original, mime_type_informado=None):
    mime_detectado = mimetypes.guess_type(nome_original)[0]
    mime_informado = _texto(mime_type_informado)
    if mime_informado:
        return mime_informado
    return mime_detectado or 'application/octet-stream'

def _obter_caminho_absoluto(caminho_relativo):
    caminho_relativo = _validar_texto_obrigatorio(caminho_relativo, 'o caminho do documento')
    rel = Path(caminho_relativo)
    if rel.is_absolute():
        raise ValueError('Caminho de documento inválido.')
    caminho = (BASE_DIR / rel).resolve()
    upload_base = UPLOAD_DIR.resolve()
    try:
        caminho.relative_to(upload_base)
    except ValueError:
        raise ValueError('Caminho de documento inválido.')
    return caminho

def _validar_pasta_codigo(codigo):
    codigo = _validar_texto_obrigatorio(codigo, 'o código da fiscalização')
    caracteres_permitidos = set('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_')
    if not set(codigo).issubset(caracteres_permitidos):
        raise ValueError('Código da fiscalização inválido para armazenamento.')
    return codigo

def _arquivo_duplicado_cursor(cursor, fiscalizacao_id, hash_sha256):
    cursor.execute('\n        SELECT\n            id,\n            nome_original\n\n        FROM documentos_fiscalizacao\n\n        WHERE\n            fiscalizacao_id = ?\n            AND hash_sha256 = ?\n\n        LIMIT 1\n        ', (fiscalizacao_id, hash_sha256))
    resultado = cursor.fetchone()
    if not resultado:
        return None
    return {'id': resultado[0], 'nome_original': resultado[1]}

def _gravar_arquivo_atomicamente(destino_final, conteudo):
    destino_final = Path(destino_final)
    destino_final.parent.mkdir(parents=True, exist_ok=True)
    temporario = destino_final.parent / f'.{destino_final.name}.{uuid.uuid4().hex}.tmp'
    try:
        with open(temporario, 'xb') as arquivo:
            arquivo.write(conteudo)
            arquivo.flush()
            os.fsync(arquivo.fileno())
        if destino_final.exists():
            raise FileExistsError('O nome interno do arquivo já existe.')
        os.replace(temporario, destino_final)
    except Exception:
        if temporario.exists():
            try:
                temporario.unlink()
            except OSError:
                pass
        raise

def _verificar_arquivo_fisico(documento):
    caminho = _obter_caminho_absoluto(documento['caminho_relativo'])
    if not caminho.exists():
        return {'ok': False, 'motivo': 'Arquivo físico ausente.', 'caminho': caminho, 'conteudo': None}
    if not caminho.is_file():
        return {'ok': False, 'motivo': 'O caminho armazenado não é um arquivo.', 'caminho': caminho, 'conteudo': None}
    try:
        tamanho_fisico = caminho.stat().st_size
    except OSError:
        return {'ok': False, 'motivo': 'Não foi possível consultar o arquivo físico.', 'caminho': caminho, 'conteudo': None}
    tamanho_esperado = documento.get('tamanho_bytes')
    if tamanho_esperado is not None and tamanho_fisico != tamanho_esperado:
        return {'ok': False, 'motivo': 'Tamanho do arquivo divergente.', 'caminho': caminho, 'conteudo': None}
    try:
        conteudo = caminho.read_bytes()
    except PermissionError:
        return {'ok': False, 'motivo': 'Sem permissão para ler o arquivo.', 'caminho': caminho, 'conteudo': None}
    except OSError:
        return {'ok': False, 'motivo': 'Não foi possível ler o arquivo.', 'caminho': caminho, 'conteudo': None}
    hash_atual = _calcular_sha256(conteudo)
    if hash_atual != documento.get('hash_sha256'):
        return {'ok': False, 'motivo': 'Hash SHA-256 divergente.', 'caminho': caminho, 'conteudo': None}
    return {'ok': True, 'motivo': 'OK', 'caminho': caminho, 'conteudo': conteudo}

def coluna_existe(tabela, coluna):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute(f'PRAGMA table_info({tabela})')
        return coluna in [item[1] for item in cursor.fetchall()]
    finally:
        conn.close()

def criar_tabelas():
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("\n            CREATE TABLE IF NOT EXISTS fiscalizacoes (\n\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n\n                codigo TEXT UNIQUE NOT NULL,\n\n                titulo TEXT NOT NULL,\n\n                objeto TEXT,\n\n                empresa TEXT NOT NULL,\n\n                filial TEXT,\n\n                orgao TEXT,\n\n                exposicao REAL\n                    NOT NULL\n                    DEFAULT 0,\n\n                probabilidade TEXT\n                    NOT NULL\n                    DEFAULT 'Possível',\n\n                fase TEXT\n                    NOT NULL\n                    DEFAULT 'Intimação recebida',\n\n                data_entrada_fase TEXT,\n\n                data_recebimento TEXT,\n\n                prazo_resposta TEXT,\n\n                responsavel_principal TEXT,\n\n                tarefa_estracta TEXT,\n\n                arquivada INTEGER\n                    NOT NULL\n                    DEFAULT 0,\n\n                arquivada_em TEXT,\n\n                criado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP,\n\n                atualizado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP\n            )\n            ")
        cursor.execute('\n            CREATE TABLE IF NOT EXISTS tributos (\n\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n\n                nome TEXT\n                    UNIQUE\n                    NOT NULL\n            )\n            ')
        cursor.execute('\n            CREATE TABLE IF NOT EXISTS fiscalizacao_tributos (\n\n                fiscalizacao_id INTEGER\n                    NOT NULL,\n\n                tributo_id INTEGER\n                    NOT NULL,\n\n                PRIMARY KEY (\n                    fiscalizacao_id,\n                    tributo_id\n                ),\n\n                FOREIGN KEY (\n                    fiscalizacao_id\n                )\n                REFERENCES fiscalizacoes(id)\n                ON DELETE CASCADE,\n\n                FOREIGN KEY (\n                    tributo_id\n                )\n                REFERENCES tributos(id)\n                ON DELETE CASCADE\n            )\n            ')
        cursor.execute('\n            CREATE TABLE IF NOT EXISTS areas (\n\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n\n                nome TEXT\n                    UNIQUE\n                    NOT NULL\n            )\n            ')
        cursor.execute("\n            CREATE TABLE IF NOT EXISTS fiscalizacao_areas (\n\n                fiscalizacao_id INTEGER\n                    NOT NULL,\n\n                area_id INTEGER\n                    NOT NULL,\n\n                tipo TEXT\n                    NOT NULL\n                    DEFAULT 'Suporte técnico',\n\n                status TEXT\n                    NOT NULL\n                    DEFAULT 'Pendente',\n\n                PRIMARY KEY (\n                    fiscalizacao_id,\n                    area_id\n                ),\n\n                FOREIGN KEY (\n                    fiscalizacao_id\n                )\n                REFERENCES fiscalizacoes(id)\n                ON DELETE CASCADE,\n\n                FOREIGN KEY (\n                    area_id\n                )\n                REFERENCES areas(id)\n                ON DELETE CASCADE\n            )\n            ")
        cursor.execute('\n            CREATE TABLE IF NOT EXISTS historico_fiscalizacao (\n\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n\n                fiscalizacao_id INTEGER\n                    NOT NULL,\n\n                fase TEXT\n                    NOT NULL,\n\n                data TEXT\n                    NOT NULL,\n\n                observacao TEXT,\n\n                criado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP,\n\n                FOREIGN KEY (\n                    fiscalizacao_id\n                )\n                REFERENCES fiscalizacoes(id)\n                ON DELETE CASCADE\n            )\n            ')
        cursor.execute("\n            CREATE TABLE IF NOT EXISTS tarefas (\n\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n\n                fiscalizacao_id INTEGER\n                    NOT NULL,\n\n                titulo TEXT NOT NULL,\n\n                descricao TEXT,\n\n                responsavel TEXT,\n\n                prazo TEXT,\n\n                status TEXT\n                    NOT NULL\n                    DEFAULT 'Pendente',\n\n                concluida_em TEXT,\n\n                criado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP,\n\n                atualizado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP,\n\n                FOREIGN KEY (\n                    fiscalizacao_id\n                )\n                REFERENCES fiscalizacoes(id)\n                ON DELETE CASCADE\n            )\n            ")
        cursor.execute("\n            CREATE TABLE IF NOT EXISTS flash_reports (\n\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n\n                fiscalizacao_id INTEGER\n                    UNIQUE\n                    NOT NULL,\n\n                resumo_executivo TEXT,\n\n                problema TEXT,\n\n                causa TEXT,\n\n                necessarias TEXT,\n\n                implementadas TEXT,\n\n                riscos_impactos TEXT,\n\n                proximos_passos TEXT,\n\n                conclusao TEXT,\n\n                status TEXT\n                    NOT NULL\n                    DEFAULT 'Em elaboração',\n\n                criado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP,\n\n                atualizado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP,\n\n                FOREIGN KEY (\n                    fiscalizacao_id\n                )\n                REFERENCES fiscalizacoes(id)\n                ON DELETE CASCADE\n            )\n            ")
        cursor.execute('\n            CREATE TABLE IF NOT EXISTS documentos_fiscalizacao (\n\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n\n                fiscalizacao_id INTEGER\n                    NOT NULL,\n\n                historico_id INTEGER,\n\n                tipo_documento TEXT\n                    NOT NULL,\n\n                descricao TEXT,\n\n                data_documento TEXT,\n\n                nome_original TEXT\n                    NOT NULL,\n\n                nome_armazenado TEXT\n                    NOT NULL,\n\n                caminho_relativo TEXT\n                    NOT NULL\n                    UNIQUE,\n\n                mime_type TEXT,\n\n                tamanho_bytes INTEGER\n                    NOT NULL,\n\n                hash_sha256 TEXT\n                    NOT NULL,\n\n                criado_em TEXT\n                    NOT NULL\n                    DEFAULT CURRENT_TIMESTAMP,\n\n                FOREIGN KEY (\n                    fiscalizacao_id\n                )\n                REFERENCES fiscalizacoes(id)\n                ON DELETE CASCADE,\n\n                FOREIGN KEY (\n                    historico_id\n                )\n                REFERENCES historico_fiscalizacao(id)\n                ON DELETE SET NULL\n            )\n            ')
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def executar_migracoes():
    migracoes = [('fiscalizacoes', 'data_recebimento', 'TEXT'), ('fiscalizacoes', 'prazo_resposta', 'TEXT'), ('fiscalizacoes', 'responsavel_principal', 'TEXT'), ('fiscalizacoes', 'tarefa_estracta', 'TEXT'), ('fiscalizacoes', 'arquivada', 'INTEGER NOT NULL DEFAULT 0'), ('fiscalizacoes', 'arquivada_em', 'TEXT'), ('flash_reports', 'resumo_executivo', 'TEXT'), ('flash_reports', 'riscos_impactos', 'TEXT'), ('flash_reports', 'proximos_passos', 'TEXT'), ('flash_reports', 'conclusao', 'TEXT')]
    for tabela, coluna, definicao in migracoes:
        if not coluna_existe(tabela, coluna):
            conn = conectar()
            try:
                conn.execute(f'\n                    ALTER TABLE {tabela}\n                    ADD COLUMN {coluna}\n                    {definicao}\n                    ')
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()

def criar_indices():
    conn = conectar()
    try:
        cursor = conn.cursor()
        comandos = ['\n            CREATE INDEX IF NOT EXISTS\n            idx_fiscalizacoes_arquivada\n            ON fiscalizacoes(arquivada)\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_fiscalizacoes_fase\n            ON fiscalizacoes(fase)\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_fiscalizacoes_prazo\n            ON fiscalizacoes(prazo_resposta)\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_historico_fiscalizacao\n            ON historico_fiscalizacao(\n                fiscalizacao_id,\n                data\n            )\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_tarefas_fiscalizacao\n            ON tarefas(fiscalizacao_id)\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_documentos_fiscalizacao\n            ON documentos_fiscalizacao(\n                fiscalizacao_id\n            )\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_documentos_tipo\n            ON documentos_fiscalizacao(\n                tipo_documento\n            )\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_documentos_data\n            ON documentos_fiscalizacao(\n                data_documento\n            )\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_documentos_hash\n            ON documentos_fiscalizacao(\n                hash_sha256\n            )\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_documentos_fiscalizacao_hash\n            ON documentos_fiscalizacao(\n                fiscalizacao_id,\n                hash_sha256\n            )\n            ']
        for comando in comandos:
            cursor.execute(comando)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def inserir_dados_mestres():
    tributos = ['ICMS', 'ISS', 'IRPJ', 'CSLL', 'PIS', 'COFINS', 'INSS', 'IPI']
    areas = ['Jurídico Tributário', 'Financeiro', 'Controladoria', 'Auditoria Interna', 'Diretoria']
    conn = conectar()
    try:
        cursor = conn.cursor()
        for tributo in tributos:
            cursor.execute('\n                INSERT OR IGNORE INTO tributos (\n                    nome\n                )\n\n                VALUES (?)\n                ', (tributo,))
        for area in areas:
            cursor.execute('\n                INSERT OR IGNORE INTO areas (\n                    nome\n                )\n\n                VALUES (?)\n                ', (area,))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def garantir_flash_reports():
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute("\n            INSERT OR IGNORE\n            INTO flash_reports (\n                fiscalizacao_id,\n                resumo_executivo,\n                problema,\n                causa,\n                necessarias,\n                implementadas,\n                riscos_impactos,\n                proximos_passos,\n                conclusao,\n                status\n            )\n\n            SELECT\n                id,\n                '',\n                '',\n                '',\n                '',\n                '',\n                '',\n                '',\n                '',\n                'Em elaboração'\n\n            FROM fiscalizacoes\n            ")
        conn.commit()
    finally:
        conn.close()

def inicializar_banco():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    criar_tabelas()
    executar_migracoes()
    criar_indices()
    inserir_dados_mestres()
    garantir_flash_reports()
    garantir_estrutura_comunicacoes()

def cadastrar_fiscalizacao(titulo, objeto, empresa, filial, orgao, exposicao, probabilidade, tributos=None, areas=None, data_recebimento=None, prazo_resposta=None, responsavel_principal=None, tarefa_estracta=None):
    titulo = _validar_texto_obrigatorio(titulo, 'o título da fiscalização')
    empresa = _validar_texto_obrigatorio(empresa, 'a empresa')
    exposicao = _validar_exposicao(exposicao)
    probabilidade = _validar_probabilidade(probabilidade)
    data_recebimento = validar_data_iso(data_recebimento) or _hoje()
    prazo_resposta = validar_data_iso(prazo_resposta)
    _validar_intervalo_datas(data_recebimento, prazo_resposta)
    tributos = list(dict.fromkeys(tributos or []))
    areas = list(dict.fromkeys(areas or []))
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        codigo = _gerar_codigo_cursor(cursor)
        agora = _agora()
        hoje = _hoje()
        cursor.execute("\n            INSERT INTO fiscalizacoes (\n                codigo,\n                titulo,\n                objeto,\n                empresa,\n                filial,\n                orgao,\n                exposicao,\n                probabilidade,\n                fase,\n                data_entrada_fase,\n                data_recebimento,\n                prazo_resposta,\n                responsavel_principal,\n                tarefa_estracta,\n                arquivada,\n                criado_em,\n                atualizado_em\n            )\n\n            VALUES (\n                ?, ?, ?, ?, ?, ?, ?, ?,\n                'Intimação recebida',\n                ?, ?, ?, ?, ?,\n                0, ?, ?\n            )\n            ", (codigo, titulo, _texto(objeto), empresa, _texto(filial), _texto(orgao), exposicao, probabilidade, hoje, data_recebimento, prazo_resposta, _texto(responsavel_principal), _texto(tarefa_estracta), agora, agora))
        fiscalizacao_id = cursor.lastrowid
        for tributo in tributos:
            cursor.execute('\n                SELECT id\n                FROM tributos\n                WHERE nome = ?\n                ', (tributo,))
            resultado = cursor.fetchone()
            if not resultado:
                raise ValueError(f'Tributo não cadastrado: {tributo}')
            cursor.execute('\n                INSERT INTO fiscalizacao_tributos (\n                    fiscalizacao_id,\n                    tributo_id\n                )\n\n                VALUES (?, ?)\n                ', (fiscalizacao_id, resultado[0]))
        for area in areas:
            cursor.execute('\n                SELECT id\n                FROM areas\n                WHERE nome = ?\n                ', (area,))
            resultado = cursor.fetchone()
            if not resultado:
                raise ValueError(f'Área não cadastrada: {area}')
            cursor.execute('\n                INSERT INTO fiscalizacao_areas (\n                    fiscalizacao_id,\n                    area_id,\n                    tipo,\n                    status\n                )\n\n                VALUES (?, ?, ?, ?)\n                ', (fiscalizacao_id, resultado[0], 'Suporte técnico', 'Pendente'))
        _registrar_historico(cursor, fiscalizacao_id, 'Intimação recebida', 'Registro criado.', hoje)
        cursor.execute("\n            INSERT INTO flash_reports (\n                fiscalizacao_id,\n                status,\n                criado_em,\n                atualizado_em\n            )\n\n            VALUES (\n                ?,\n                'Em elaboração',\n                ?,\n                ?\n            )\n            ", (fiscalizacao_id, agora, agora))
        conn.commit()
        return codigo
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def listar_fiscalizacoes(situacao='Ativas'):
    filtros = {'Ativas': 'WHERE COALESCE(arquivada, 0) = 0', 'Arquivadas': 'WHERE COALESCE(arquivada, 0) = 1', 'Todas': ''}
    if situacao not in filtros:
        raise ValueError('Situação de listagem inválida.')
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute(f'\n            SELECT\n                id,\n                codigo,\n                titulo,\n                empresa,\n                filial,\n                orgao,\n                exposicao,\n                probabilidade,\n                fase,\n                data_entrada_fase\n\n            FROM fiscalizacoes\n\n            {filtros[situacao]}\n\n            ORDER BY id DESC\n            ')
        return cursor.fetchall()
    finally:
        conn.close()

def buscar_fiscalizacao_por_id(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                id,\n                codigo,\n                titulo,\n                objeto,\n                empresa,\n                filial,\n                orgao,\n                exposicao,\n                probabilidade,\n                fase,\n                data_entrada_fase,\n                criado_em,\n                atualizado_em\n\n            FROM fiscalizacoes\n\n            WHERE id = ?\n            ', (fiscalizacao_id,))
        return cursor.fetchone()
    finally:
        conn.close()

def buscar_fiscalizacao_por_codigo(codigo):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                id,\n                codigo,\n                titulo,\n                objeto,\n                empresa,\n                filial,\n                orgao,\n                exposicao,\n                probabilidade,\n                fase,\n                data_entrada_fase,\n                criado_em,\n                atualizado_em\n\n            FROM fiscalizacoes\n\n            WHERE codigo = ?\n            ', (codigo,))
        return cursor.fetchone()
    finally:
        conn.close()

def buscar_dados_complementares_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                data_recebimento,\n                prazo_resposta,\n                responsavel_principal,\n                tarefa_estracta\n\n            FROM fiscalizacoes\n\n            WHERE id = ?\n            ', (fiscalizacao_id,))
        resultado = cursor.fetchone()
        if not resultado:
            return None
        return {'data_recebimento': resultado[0], 'prazo_resposta': resultado[1], 'responsavel_principal': resultado[2], 'tarefa_estracta': resultado[3]}
    finally:
        conn.close()

def buscar_status_arquivamento(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                COALESCE(\n                    arquivada,\n                    0\n                ),\n                arquivada_em\n\n            FROM fiscalizacoes\n\n            WHERE id = ?\n            ', (fiscalizacao_id,))
        resultado = cursor.fetchone()
        if not resultado:
            return None
        return {'arquivada': bool(resultado[0]), 'arquivada_em': resultado[1]}
    finally:
        conn.close()

def arquivar_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        estado = _buscar_estado_fiscalizacao(cursor, fiscalizacao_id)
        if estado['arquivada']:
            raise ValueError('A fiscalização já está arquivada.')
        agora = _agora()
        cursor.execute('\n            UPDATE fiscalizacoes\n\n            SET\n                arquivada = 1,\n                arquivada_em = ?,\n                atualizado_em = ?\n\n            WHERE id = ?\n            ', (agora, agora, fiscalizacao_id))
        _registrar_historico(cursor, fiscalizacao_id, estado['fase'], 'Fiscalização arquivada.')
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def reativar_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        estado = _buscar_estado_fiscalizacao(cursor, fiscalizacao_id)
        if not estado['arquivada']:
            raise ValueError('A fiscalização já está ativa.')
        cursor.execute('\n            UPDATE fiscalizacoes\n\n            SET\n                arquivada = 0,\n                arquivada_em = NULL,\n                atualizado_em = ?\n\n            WHERE id = ?\n            ', (_agora(), fiscalizacao_id))
        _registrar_historico(cursor, fiscalizacao_id, estado['fase'], 'Fiscalização reativada.')
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def atualizar_dados_fiscalizacao(fiscalizacao_id, titulo, objeto, empresa, filial, orgao, exposicao, probabilidade, tributos=None, areas=None, responsavel_principal=None, tarefa_estracta=None):
    titulo = _validar_texto_obrigatorio(titulo, 'o título')
    empresa = _validar_texto_obrigatorio(empresa, 'a empresa')
    exposicao = _validar_exposicao(exposicao)
    probabilidade = _validar_probabilidade(probabilidade)
    tributos = list(dict.fromkeys(tributos or []))
    areas = list(dict.fromkeys(areas or []))
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        _garantir_nao_arquivada(cursor, fiscalizacao_id)
        cursor.execute('\n            SELECT\n                a.nome,\n                fa.tipo,\n                fa.status\n\n            FROM fiscalizacao_areas fa\n\n            INNER JOIN areas a\n                ON a.id = fa.area_id\n\n            WHERE fa.fiscalizacao_id = ?\n            ', (fiscalizacao_id,))
        areas_existentes = {item[0]: {'tipo': item[1], 'status': item[2]} for item in cursor.fetchall()}
        agora = _agora()
        cursor.execute('\n            UPDATE fiscalizacoes\n\n            SET\n                titulo = ?,\n                objeto = ?,\n                empresa = ?,\n                filial = ?,\n                orgao = ?,\n                exposicao = ?,\n                probabilidade = ?,\n                responsavel_principal = ?,\n                tarefa_estracta = ?,\n                atualizado_em = ?\n\n            WHERE id = ?\n            ', (titulo, _texto(objeto), empresa, _texto(filial), _texto(orgao), exposicao, probabilidade, _texto(responsavel_principal), _texto(tarefa_estracta), agora, fiscalizacao_id))
        cursor.execute('\n            DELETE FROM fiscalizacao_tributos\n            WHERE fiscalizacao_id = ?\n            ', (fiscalizacao_id,))
        for tributo in tributos:
            cursor.execute('\n                SELECT id\n                FROM tributos\n                WHERE nome = ?\n                ', (tributo,))
            resultado = cursor.fetchone()
            if not resultado:
                raise ValueError(f'Tributo inválido: {tributo}')
            cursor.execute('\n                INSERT INTO fiscalizacao_tributos (\n                    fiscalizacao_id,\n                    tributo_id\n                )\n\n                VALUES (?, ?)\n                ', (fiscalizacao_id, resultado[0]))
        cursor.execute('\n            DELETE FROM fiscalizacao_areas\n            WHERE fiscalizacao_id = ?\n            ', (fiscalizacao_id,))
        for area in areas:
            cursor.execute('\n                SELECT id\n                FROM areas\n                WHERE nome = ?\n                ', (area,))
            resultado = cursor.fetchone()
            if not resultado:
                raise ValueError(f'Área inválida: {area}')
            anterior = areas_existentes.get(area)
            tipo = anterior['tipo'] if anterior else 'Suporte técnico'
            status = anterior['status'] if anterior else 'Pendente'
            cursor.execute('\n                INSERT INTO fiscalizacao_areas (\n                    fiscalizacao_id,\n                    area_id,\n                    tipo,\n                    status\n                )\n\n                VALUES (?, ?, ?, ?)\n                ', (fiscalizacao_id, resultado[0], tipo, status))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def atualizar_prazos_fiscalizacao(fiscalizacao_id, data_recebimento=None, prazo_resposta=None):
    data_recebimento = validar_data_iso(data_recebimento)
    prazo_resposta = validar_data_iso(prazo_resposta)
    _validar_intervalo_datas(data_recebimento, prazo_resposta)
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        _garantir_operacional(cursor, fiscalizacao_id)
        cursor.execute('\n            UPDATE fiscalizacoes\n\n            SET\n                data_recebimento = ?,\n                prazo_resposta = ?,\n                atualizado_em = ?\n\n            WHERE id = ?\n            ', (data_recebimento, prazo_resposta, _agora(), fiscalizacao_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def calcular_status_prazo(prazo_resposta, fase=None):
    if fase == 'Encerrada':
        return {'status': 'Encerrada', 'dias_restantes': None}
    if not prazo_resposta:
        return {'status': 'Sem prazo', 'dias_restantes': None}
    prazo = date.fromisoformat(prazo_resposta)
    dias = (prazo - date.today()).days
    if dias < 0:
        status = 'Vencido'
    elif dias == 0:
        status = 'Vence hoje'
    elif dias <= 3:
        status = 'Crítico'
    elif dias <= 7:
        status = 'Atenção'
    else:
        status = 'No prazo'
    return {'status': status, 'dias_restantes': dias}

def listar_prazos_fiscalizacoes(incluir_arquivadas=False):
    conn = conectar()
    try:
        cursor = conn.cursor()
        filtro = '' if incluir_arquivadas else 'WHERE COALESCE(arquivada, 0) = 0'
        cursor.execute(f'\n            SELECT\n                id,\n                codigo,\n                titulo,\n                empresa,\n                orgao,\n                fase,\n                data_recebimento,\n                prazo_resposta\n\n            FROM fiscalizacoes\n\n            {filtro}\n\n            ORDER BY\n                CASE\n                    WHEN prazo_resposta IS NULL\n                    THEN 1\n                    ELSE 0\n                END,\n                prazo_resposta ASC\n            ')
        resultados = cursor.fetchall()
    finally:
        conn.close()
    saida = []
    for item in resultados:
        fiscalizacao_id, codigo, titulo, empresa, orgao, fase, recebimento, prazo = item
        status = calcular_status_prazo(prazo, fase)
        saida.append({'id': fiscalizacao_id, 'codigo': codigo, 'titulo': titulo, 'empresa': empresa, 'orgao': orgao, 'fase': fase, 'data_recebimento': recebimento, 'prazo_resposta': prazo, 'status_prazo': status['status'], 'dias_restantes': status['dias_restantes']})
    return saida

def obter_resumo_prazos():
    resumo = {'vencidos': 0, 'vence_hoje': 0, 'criticos': 0, 'atencao': 0, 'no_prazo': 0, 'sem_prazo': 0, 'encerradas': 0}
    for item in listar_prazos_fiscalizacoes():
        status = item['status_prazo']
        if status == 'Vencido':
            resumo['vencidos'] += 1
        elif status == 'Vence hoje':
            resumo['vence_hoje'] += 1
        elif status == 'Crítico':
            resumo['criticos'] += 1
        elif status == 'Atenção':
            resumo['atencao'] += 1
        elif status == 'No prazo':
            resumo['no_prazo'] += 1
        elif status == 'Sem prazo':
            resumo['sem_prazo'] += 1
        elif status == 'Encerrada':
            resumo['encerradas'] += 1
    return resumo

def listar_tributos_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT t.nome\n\n            FROM fiscalizacao_tributos ft\n\n            JOIN tributos t\n                ON t.id = ft.tributo_id\n\n            WHERE ft.fiscalizacao_id = ?\n\n            ORDER BY t.nome\n            ', (fiscalizacao_id,))
        return [item[0] for item in cursor.fetchall()]
    finally:
        conn.close()

def listar_areas_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                a.nome,\n                fa.tipo,\n                fa.status\n\n            FROM fiscalizacao_areas fa\n\n            JOIN areas a\n                ON a.id = fa.area_id\n\n            WHERE fa.fiscalizacao_id = ?\n\n            ORDER BY a.nome\n            ', (fiscalizacao_id,))
        return cursor.fetchall()
    finally:
        conn.close()

def listar_historico_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                id,\n                fase,\n                data,\n                observacao,\n                criado_em\n\n            FROM historico_fiscalizacao\n\n            WHERE fiscalizacao_id = ?\n\n            ORDER BY\n                data DESC,\n                id DESC\n            ', (fiscalizacao_id,))
        return cursor.fetchall()
    finally:
        conn.close()

def alterar_fase_fiscalizacao(fiscalizacao_id, nova_fase, observacao=''):
    nova_fase = _validar_fase(nova_fase)
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        estado = _garantir_nao_arquivada(cursor, fiscalizacao_id)
        fase_anterior = estado['fase']
        if fase_anterior == 'Encerrada':
            raise ValueError('A fiscalização já está encerrada.')
        if nova_fase == fase_anterior:
            raise ValueError('A fiscalização já está nesta fase.')
        agora = _agora()
        hoje = _hoje()
        cursor.execute('\n            UPDATE fiscalizacoes\n\n            SET\n                fase = ?,\n                data_entrada_fase = ?,\n                atualizado_em = ?\n\n            WHERE id = ?\n            ', (nova_fase, hoje, agora, fiscalizacao_id))
        observacao_final = _texto(observacao)
        if not observacao_final:
            observacao_final = f"Fase alterada de '{fase_anterior}' para '{nova_fase}'."
        _registrar_historico(cursor, fiscalizacao_id, nova_fase, observacao_final, hoje)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def cadastrar_tarefa(fiscalizacao_id, titulo, descricao='', responsavel='', prazo=None, status='Pendente'):
    titulo = _validar_texto_obrigatorio(titulo, 'o título da tarefa')
    status = _validar_status_tarefa(status)
    prazo = validar_data_iso(prazo)
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        _garantir_operacional(cursor, fiscalizacao_id)
        agora = _agora()
        cursor.execute('\n            INSERT INTO tarefas (\n                fiscalizacao_id,\n                titulo,\n                descricao,\n                responsavel,\n                prazo,\n                status,\n                concluida_em,\n                criado_em,\n                atualizado_em\n            )\n\n            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)\n            ', (fiscalizacao_id, titulo, _texto(descricao), _texto(responsavel), prazo, status, agora if status == 'Concluída' else None, agora, agora))
        tarefa_id = cursor.lastrowid
        conn.commit()
        return tarefa_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def listar_tarefas_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute("\n            SELECT\n                id,\n                titulo,\n                descricao,\n                responsavel,\n                prazo,\n                status,\n                concluida_em,\n                criado_em,\n                atualizado_em\n\n            FROM tarefas\n\n            WHERE fiscalizacao_id = ?\n\n            ORDER BY\n                CASE\n                    WHEN status = 'Concluída'\n                    THEN 1\n                    ELSE 0\n                END,\n                prazo ASC,\n                id DESC\n            ", (fiscalizacao_id,))
        return cursor.fetchall()
    finally:
        conn.close()

def atualizar_status_tarefa(tarefa_id, novo_status):
    novo_status = _validar_status_tarefa(novo_status)
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                t.fiscalizacao_id,\n                f.arquivada\n\n            FROM tarefas t\n\n            JOIN fiscalizacoes f\n                ON f.id = t.fiscalizacao_id\n\n            WHERE t.id = ?\n            ', (tarefa_id,))
        resultado = cursor.fetchone()
        if not resultado:
            raise ValueError('Tarefa não encontrada.')
        if resultado[1]:
            raise ValueError('A fiscalização está arquivada.')
        agora = _agora()
        cursor.execute('\n            UPDATE tarefas\n\n            SET\n                status = ?,\n                concluida_em = ?,\n                atualizado_em = ?\n\n            WHERE id = ?\n            ', (novo_status, agora if novo_status == 'Concluída' else None, agora, tarefa_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def calcular_status_prazo_tarefa(prazo, status):
    if status == 'Concluída':
        return {'status_prazo': 'Concluída', 'dias_restantes': None}
    if not prazo:
        return {'status_prazo': 'Sem prazo', 'dias_restantes': None}
    prazo_data = date.fromisoformat(prazo)
    dias = (prazo_data - date.today()).days
    if dias < 0:
        situacao = 'Vencida'
    elif dias == 0:
        situacao = 'Vence hoje'
    elif dias <= 3:
        situacao = 'Crítica'
    elif dias <= 7:
        situacao = 'Atenção'
    else:
        situacao = 'No prazo'
    return {'status_prazo': situacao, 'dias_restantes': dias}

def obter_resumo_tarefas():
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                t.prazo,\n                t.status\n\n            FROM tarefas t\n\n            JOIN fiscalizacoes f\n                ON f.id = t.fiscalizacao_id\n\n            WHERE COALESCE(\n                f.arquivada,\n                0\n            ) = 0\n            ')
        resultados = cursor.fetchall()
    finally:
        conn.close()
    resumo = {'pendentes': 0, 'em_andamento': 0, 'concluidas': 0, 'vencidas': 0, 'vence_hoje': 0, 'criticas': 0}
    for prazo, status in resultados:
        if status == 'Pendente':
            resumo['pendentes'] += 1
        elif status == 'Em andamento':
            resumo['em_andamento'] += 1
        elif status == 'Concluída':
            resumo['concluidas'] += 1
        situacao = calcular_status_prazo_tarefa(prazo, status)['status_prazo']
        if situacao == 'Vencida':
            resumo['vencidas'] += 1
        elif situacao == 'Vence hoje':
            resumo['vence_hoje'] += 1
        elif situacao == 'Crítica':
            resumo['criticas'] += 1
    return resumo

def buscar_flash_report(fiscalizacao_id):
    garantir_flash_reports()
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                id,\n                fiscalizacao_id,\n                resumo_executivo,\n                problema,\n                causa,\n                necessarias,\n                implementadas,\n                riscos_impactos,\n                proximos_passos,\n                conclusao,\n                status,\n                criado_em,\n                atualizado_em\n\n            FROM flash_reports\n\n            WHERE fiscalizacao_id = ?\n            ', (fiscalizacao_id,))
        resultado = cursor.fetchone()
        if not resultado:
            return None
        return {'id': resultado[0], 'fiscalizacao_id': resultado[1], 'resumo_executivo': resultado[2] or '', 'problema': resultado[3] or '', 'causa': resultado[4] or '', 'necessarias': resultado[5] or '', 'implementadas': resultado[6] or '', 'riscos_impactos': resultado[7] or '', 'proximos_passos': resultado[8] or '', 'conclusao': resultado[9] or '', 'status': resultado[10] or 'Em elaboração', 'criado_em': resultado[11], 'atualizado_em': resultado[12]}
    finally:
        conn.close()

def atualizar_flash_report(fiscalizacao_id, resumo_executivo='', problema='', causa='', necessarias='', implementadas='', riscos_impactos='', proximos_passos='', conclusao='', status='Em elaboração'):
    status = _validar_status_flash(status)
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        _garantir_nao_arquivada(cursor, fiscalizacao_id)
        cursor.execute('\n            UPDATE flash_reports\n\n            SET\n                resumo_executivo = ?,\n                problema = ?,\n                causa = ?,\n                necessarias = ?,\n                implementadas = ?,\n                riscos_impactos = ?,\n                proximos_passos = ?,\n                conclusao = ?,\n                status = ?,\n                atualizado_em = ?\n\n            WHERE fiscalizacao_id = ?\n            ', (_texto(resumo_executivo), _texto(problema), _texto(causa), _texto(necessarias), _texto(implementadas), _texto(riscos_impactos), _texto(proximos_passos), _texto(conclusao), status, _agora(), fiscalizacao_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def adicionar_documento_fiscalizacao(fiscalizacao_id, tipo_documento, nome_original, conteudo, descricao='', data_documento=None, mime_type=None, historico_id=None):
    tipo_documento = _validar_tipo_documento(tipo_documento)
    nome_original, extensao = _validar_nome_documento(nome_original)
    conteudo, tamanho = _validar_conteudo_documento(conteudo)
    data_documento = validar_data_iso(data_documento)
    hash_sha256 = _calcular_sha256(conteudo)
    mime_type_final = _mime_type_seguro(nome_original, mime_type)
    conn = conectar()
    caminho_final = None
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        estado = _garantir_nao_arquivada(cursor, fiscalizacao_id)
        duplicado = _arquivo_duplicado_cursor(cursor, fiscalizacao_id, hash_sha256)
        if duplicado:
            raise ValueError(f"Este arquivo já está cadastrado nesta fiscalização como '{duplicado['nome_original']}'.")
        if historico_id is not None:
            cursor.execute('\n                SELECT id\n\n                FROM historico_fiscalizacao\n\n                WHERE\n                    id = ?\n                    AND fiscalizacao_id = ?\n                ', (historico_id, fiscalizacao_id))
            if not cursor.fetchone():
                raise ValueError('Evento de Timeline inválido.')
        codigo = _validar_pasta_codigo(estado['codigo'])
        pasta_fiscalizacao = UPLOAD_DIR / codigo
        pasta_fiscalizacao.mkdir(parents=True, exist_ok=True)
        nome_armazenado = f'{uuid.uuid4().hex}{extensao}'
        caminho_final = pasta_fiscalizacao / nome_armazenado
        caminho_relativo = caminho_final.relative_to(BASE_DIR).as_posix()
        _gravar_arquivo_atomicamente(caminho_final, conteudo)
        hash_gravado = _calcular_sha256(caminho_final.read_bytes())
        if hash_gravado != hash_sha256:
            raise IOError('Falha de integridade ao gravar o documento.')
        if caminho_final.stat().st_size != tamanho:
            raise IOError('Falha de integridade no tamanho do documento gravado.')
        cursor.execute('\n            INSERT INTO documentos_fiscalizacao (\n                fiscalizacao_id,\n                historico_id,\n                tipo_documento,\n                descricao,\n                data_documento,\n                nome_original,\n                nome_armazenado,\n                caminho_relativo,\n                mime_type,\n                tamanho_bytes,\n                hash_sha256,\n                criado_em\n            )\n\n            VALUES (\n                ?, ?, ?, ?, ?, ?, ?,\n                ?, ?, ?, ?, ?\n            )\n            ', (fiscalizacao_id, historico_id, tipo_documento, _texto(descricao), data_documento, nome_original, nome_armazenado, caminho_relativo, mime_type_final, tamanho, hash_sha256, _agora()))
        documento_id = cursor.lastrowid
        conn.commit()
        return documento_id
    except Exception:
        conn.rollback()
        if caminho_final and caminho_final.exists():
            try:
                caminho_final.unlink()
            except OSError:
                pass
        raise
    finally:
        conn.close()

def listar_documentos_fiscalizacao(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        _buscar_estado_fiscalizacao(cursor, fiscalizacao_id)
        cursor.execute('\n            SELECT\n                id,\n                fiscalizacao_id,\n                historico_id,\n                tipo_documento,\n                descricao,\n                data_documento,\n                nome_original,\n                nome_armazenado,\n                caminho_relativo,\n                mime_type,\n                tamanho_bytes,\n                hash_sha256,\n                criado_em\n\n            FROM documentos_fiscalizacao\n\n            WHERE fiscalizacao_id = ?\n\n            ORDER BY\n                COALESCE(\n                    data_documento,\n                    criado_em\n                ) DESC,\n                id DESC\n            ', (fiscalizacao_id,))
        resultados = cursor.fetchall()
        documentos = []
        for item in resultados:
            documentos.append({'id': item[0], 'fiscalizacao_id': item[1], 'historico_id': item[2], 'tipo_documento': item[3], 'descricao': item[4], 'data_documento': item[5], 'nome_original': item[6], 'nome_armazenado': item[7], 'caminho_relativo': item[8], 'mime_type': item[9], 'tamanho_bytes': item[10], 'hash_sha256': item[11], 'criado_em': item[12]})
        return documentos
    finally:
        conn.close()

def obter_documento_fiscalizacao(documento_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                id,\n                fiscalizacao_id,\n                historico_id,\n                tipo_documento,\n                descricao,\n                data_documento,\n                nome_original,\n                nome_armazenado,\n                caminho_relativo,\n                mime_type,\n                tamanho_bytes,\n                hash_sha256,\n                criado_em\n\n            FROM documentos_fiscalizacao\n\n            WHERE id = ?\n            ', (documento_id,))
        item = cursor.fetchone()
        if not item:
            return None
        return {'id': item[0], 'fiscalizacao_id': item[1], 'historico_id': item[2], 'tipo_documento': item[3], 'descricao': item[4], 'data_documento': item[5], 'nome_original': item[6], 'nome_armazenado': item[7], 'caminho_relativo': item[8], 'mime_type': item[9], 'tamanho_bytes': item[10], 'hash_sha256': item[11], 'criado_em': item[12]}
    finally:
        conn.close()

def ler_conteudo_documento(documento_id):
    documento = obter_documento_fiscalizacao(documento_id)
    if not documento:
        raise ValueError('Documento não encontrado.')
    verificacao = _verificar_arquivo_fisico(documento)
    if not verificacao['ok']:
        motivo = verificacao['motivo']
        if motivo == 'Arquivo físico ausente.':
            raise FileNotFoundError('O arquivo físico não foi encontrado.')
        if motivo == 'Sem permissão para ler o arquivo.':
            raise PermissionError('Sem permissão para ler o arquivo.')
        raise ValueError(f'O documento falhou na verificação de integridade: {motivo}')
    return verificacao['conteudo']

def excluir_documento_fiscalizacao(documento_id):
    conn = conectar()
    caminho_original = None
    caminho_quarentena = None
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('\n            SELECT\n                fiscalizacao_id,\n                caminho_relativo,\n                nome_original\n\n            FROM documentos_fiscalizacao\n\n            WHERE id = ?\n            ', (documento_id,))
        resultado = cursor.fetchone()
        if not resultado:
            raise ValueError('Documento não encontrado.')
        fiscalizacao_id = resultado[0]
        _garantir_nao_arquivada(cursor, fiscalizacao_id)
        caminho_original = _obter_caminho_absoluto(resultado[1])
        if caminho_original.exists():
            if not caminho_original.is_file():
                raise ValueError('O caminho físico do documento é inválido.')
            caminho_quarentena = caminho_original.parent / f'.{caminho_original.name}.{uuid.uuid4().hex}.removing'
            os.replace(caminho_original, caminho_quarentena)
        cursor.execute('\n            DELETE FROM documentos_fiscalizacao\n\n            WHERE id = ?\n            ', (documento_id,))
        if cursor.rowcount != 1:
            raise RuntimeError('A exclusão do registro do documento não foi confirmada.')
        conn.commit()
        if caminho_quarentena and caminho_quarentena.exists():
            try:
                caminho_quarentena.unlink()
            except OSError:
                pass
        return True
    except Exception:
        conn.rollback()
        if caminho_quarentena and caminho_quarentena.exists() and caminho_original and (not caminho_original.exists()):
            try:
                os.replace(caminho_quarentena, caminho_original)
            except OSError:
                pass
        raise
    finally:
        conn.close()

def verificar_integridade_documento(documento_id):
    documento = obter_documento_fiscalizacao(documento_id)
    if not documento:
        return {'ok': False, 'motivo': 'Documento não encontrado.'}
    verificacao = _verificar_arquivo_fisico(documento)
    return {'ok': verificacao['ok'], 'motivo': verificacao['motivo']}

def verificar_integridade_documentos_fiscalizacao(fiscalizacao_id):
    documentos = listar_documentos_fiscalizacao(fiscalizacao_id)
    resultados = []
    for documento in documentos:
        verificacao = verificar_integridade_documento(documento['id'])
        resultados.append({'id': documento['id'], 'nome_original': documento['nome_original'], 'ok': verificacao['ok'], 'motivo': verificacao['motivo']})
    return resultados

def diagnosticar_repositorio_documentos():
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            SELECT\n                id,\n                caminho_relativo\n\n            FROM documentos_fiscalizacao\n            ')
        registros = cursor.fetchall()
    finally:
        conn.close()
    caminhos_registrados = set()
    registros_ausentes = []
    for documento_id, caminho_relativo in registros:
        try:
            caminho = _obter_caminho_absoluto(caminho_relativo)
            caminhos_registrados.add(caminho.resolve())
            if not caminho.exists():
                registros_ausentes.append({'id': documento_id, 'caminho_relativo': caminho_relativo})
        except Exception:
            registros_ausentes.append({'id': documento_id, 'caminho_relativo': caminho_relativo})
    arquivos_orfaos = []
    if UPLOAD_DIR.exists():
        for caminho in UPLOAD_DIR.rglob('*'):
            if not caminho.is_file():
                continue
            if caminho.resolve() not in caminhos_registrados:
                arquivos_orfaos.append(str(caminho.relative_to(BASE_DIR)))
    return {'registros_ausentes': registros_ausentes, 'arquivos_orfaos': arquivos_orfaos}

def obter_detalhes_fiscalizacao(fiscalizacao_id):
    fiscalizacao = buscar_fiscalizacao_por_id(fiscalizacao_id)
    if not fiscalizacao:
        return None
    complementares = buscar_dados_complementares_fiscalizacao(fiscalizacao_id)
    return {'fiscalizacao': fiscalizacao, 'tributos': listar_tributos_fiscalizacao(fiscalizacao_id), 'areas': listar_areas_fiscalizacao(fiscalizacao_id), 'historico': listar_historico_fiscalizacao(fiscalizacao_id), 'prazos': {'data_recebimento': complementares['data_recebimento'], 'prazo_resposta': complementares['prazo_resposta']}, 'responsavel_principal': complementares['responsavel_principal'], 'tarefa_estracta': complementares['tarefa_estracta'], 'tarefas': listar_tarefas_fiscalizacao(fiscalizacao_id), 'flash_report': buscar_flash_report(fiscalizacao_id), 'arquivamento': buscar_status_arquivamento(fiscalizacao_id), 'documentos': listar_documentos_fiscalizacao(fiscalizacao_id)}

def obter_dados_flash_report(fiscalizacao_id):
    fiscalizacao = buscar_fiscalizacao_por_id(fiscalizacao_id)
    if not fiscalizacao:
        return None
    complementares = buscar_dados_complementares_fiscalizacao(fiscalizacao_id)
    status_prazo = calcular_status_prazo(complementares['prazo_resposta'], fiscalizacao[9])
    return {'fiscalizacao': fiscalizacao, 'dados_complementares': complementares, 'tributos': listar_tributos_fiscalizacao(fiscalizacao_id), 'areas': listar_areas_fiscalizacao(fiscalizacao_id), 'historico': listar_historico_fiscalizacao(fiscalizacao_id), 'tarefas': listar_tarefas_fiscalizacao(fiscalizacao_id), 'flash_report': buscar_flash_report(fiscalizacao_id), 'status_prazo': status_prazo, 'arquivamento': buscar_status_arquivamento(fiscalizacao_id), 'documentos': listar_documentos_fiscalizacao(fiscalizacao_id)}

def diagnostico_banco():
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('PRAGMA integrity_check')
        integridade = cursor.fetchone()[0]
        cursor.execute('\n            SELECT COUNT(*)\n            FROM fiscalizacoes\n            ')
        fiscalizacoes = cursor.fetchone()[0]
        cursor.execute('\n            SELECT COUNT(*)\n            FROM tarefas\n            ')
        tarefas = cursor.fetchone()[0]
        cursor.execute('\n            SELECT COUNT(*)\n            FROM flash_reports\n            ')
        flash_reports = cursor.fetchone()[0]
        cursor.execute('\n            SELECT COUNT(*)\n            FROM documentos_fiscalizacao\n            ')
        documentos = cursor.fetchone()[0]
    finally:
        conn.close()
    repositorio = diagnosticar_repositorio_documentos()
    return {'integridade': integridade, 'fiscalizacoes': fiscalizacoes, 'tarefas': tarefas, 'flash_reports': flash_reports, 'documentos': documentos, 'registros_documentos_ausentes': len(repositorio['registros_ausentes']), 'arquivos_orfaos': len(repositorio['arquivos_orfaos'])}
PAPEIS_ENVOLVIMENTO = ['Responsável', 'Suporte técnico', 'Consultado', 'Aprovador', 'Informado']

def _validar_email_basico(valor, obrigatorio=False, campo='o e-mail'):
    email = _texto(valor)
    if not email:
        if obrigatorio:
            raise ValueError(f'Informe {campo}.')
        return ''
    if ' ' in email or '@' not in email:
        raise ValueError(f'Informe um valor válido para {campo}.')
    local, dominio = email.rsplit('@', 1)
    if not local or not dominio or '.' not in dominio or dominio.startswith('.') or dominio.endswith('.'):
        raise ValueError(f'Informe um valor válido para {campo}.')
    return email.lower()
_executar_migracoes_base_1_9d_b2 = executar_migracoes

def executar_migracoes():
    _executar_migracoes_base_1_9d_b2()
    novas_colunas_areas = [('descricao', 'TEXT'), ('email_area', 'TEXT'), ('ativa', 'INTEGER NOT NULL DEFAULT 1'), ('criado_em', 'TEXT'), ('atualizado_em', 'TEXT')]
    for coluna, definicao in novas_colunas_areas:
        if not coluna_existe('areas', coluna):
            conn = conectar()
            try:
                conn.execute(f'\n                    ALTER TABLE areas\n                    ADD COLUMN {coluna} {definicao}\n                    ')
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('\n            CREATE TABLE IF NOT EXISTS pessoas (\n                id INTEGER PRIMARY KEY AUTOINCREMENT,\n                nome TEXT NOT NULL,\n                email TEXT UNIQUE NOT NULL,\n                cargo TEXT,\n                ativa INTEGER NOT NULL DEFAULT 1,\n                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,\n                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP\n            )\n            ')
        cursor.execute("\n            CREATE TABLE IF NOT EXISTS area_pessoas (\n                area_id INTEGER NOT NULL,\n                pessoa_id INTEGER NOT NULL,\n                papel TEXT NOT NULL DEFAULT 'Informado',\n                principal INTEGER NOT NULL DEFAULT 0,\n                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,\n                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,\n                PRIMARY KEY (\n                    area_id,\n                    pessoa_id\n                ),\n                FOREIGN KEY (\n                    area_id\n                )\n                REFERENCES areas(id)\n                ON DELETE CASCADE,\n                FOREIGN KEY (\n                    pessoa_id\n                )\n                REFERENCES pessoas(id)\n                ON DELETE CASCADE\n            )\n            ")
        agora = _agora()
        cursor.execute('\n            UPDATE areas\n            SET ativa = 1\n            WHERE ativa IS NULL\n            ')
        cursor.execute("\n            UPDATE areas\n            SET criado_em = ?\n            WHERE criado_em IS NULL\n               OR TRIM(criado_em) = ''\n            ", (agora,))
        cursor.execute("\n            UPDATE areas\n            SET atualizado_em = ?\n            WHERE atualizado_em IS NULL\n               OR TRIM(atualizado_em) = ''\n            ", (agora,))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
_criar_indices_base_1_9d_b2 = criar_indices

def criar_indices():
    _criar_indices_base_1_9d_b2()
    conn = conectar()
    try:
        cursor = conn.cursor()
        comandos = ['\n            CREATE INDEX IF NOT EXISTS\n            idx_areas_ativa\n            ON areas(ativa)\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_pessoas_ativa\n            ON pessoas(ativa)\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_area_pessoas_area\n            ON area_pessoas(area_id)\n            ', '\n            CREATE INDEX IF NOT EXISTS\n            idx_area_pessoas_pessoa\n            ON area_pessoas(pessoa_id)\n            ']
        for comando in comandos:
            cursor.execute(comando)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def inicializar_banco():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    criar_tabelas()
    executar_migracoes()
    criar_indices()
    inserir_dados_mestres()
    garantir_flash_reports()
    garantir_estrutura_comunicacoes()
    garantir_estrutura_aprovacoes_encerramento()

def listar_areas_cadastradas(incluir_inativas=True):
    conn = conectar()
    try:
        cursor = conn.cursor()
        filtro = '' if incluir_inativas else 'WHERE COALESCE(ativa, 1) = 1'
        cursor.execute(f"\n            SELECT\n                id,\n                nome,\n                COALESCE(descricao, ''),\n                COALESCE(email_area, ''),\n                COALESCE(ativa, 1),\n                criado_em,\n                atualizado_em\n            FROM areas\n            {filtro}\n            ORDER BY\n                COALESCE(ativa, 1) DESC,\n                nome COLLATE NOCASE\n            ")
        return [{'id': item[0], 'nome': item[1], 'descricao': item[2], 'email_area': item[3], 'ativa': bool(item[4]), 'criado_em': item[5], 'atualizado_em': item[6]} for item in cursor.fetchall()]
    finally:
        conn.close()

def obter_area(area_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute("\n            SELECT\n                id,\n                nome,\n                COALESCE(descricao, ''),\n                COALESCE(email_area, ''),\n                COALESCE(ativa, 1),\n                criado_em,\n                atualizado_em\n            FROM areas\n            WHERE id = ?\n            ", (area_id,))
        item = cursor.fetchone()
        if not item:
            return None
        return {'id': item[0], 'nome': item[1], 'descricao': item[2], 'email_area': item[3], 'ativa': bool(item[4]), 'criado_em': item[5], 'atualizado_em': item[6]}
    finally:
        conn.close()

def cadastrar_area(nome, descricao='', email_area=''):
    nome = _validar_texto_obrigatorio(nome, 'o nome da área')
    email_area = _validar_email_basico(email_area, obrigatorio=False, campo='o e-mail da área')
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('\n            SELECT id\n            FROM areas\n            WHERE LOWER(TRIM(nome)) = LOWER(TRIM(?))\n            ', (nome,))
        if cursor.fetchone():
            raise ValueError('Já existe uma área cadastrada com este nome.')
        agora = _agora()
        cursor.execute('\n            INSERT INTO areas (\n                nome,\n                descricao,\n                email_area,\n                ativa,\n                criado_em,\n                atualizado_em\n            )\n            VALUES (?, ?, ?, 1, ?, ?)\n            ', (nome, _texto(descricao), email_area, agora, agora))
        area_id = cursor.lastrowid
        conn.commit()
        return area_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def atualizar_area(area_id, nome, descricao='', email_area=''):
    nome = _validar_texto_obrigatorio(nome, 'o nome da área')
    email_area = _validar_email_basico(email_area, obrigatorio=False, campo='o e-mail da área')
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('SELECT id FROM areas WHERE id = ?', (area_id,))
        if not cursor.fetchone():
            raise ValueError('Área não encontrada.')
        cursor.execute('\n            SELECT id\n            FROM areas\n            WHERE LOWER(TRIM(nome)) = LOWER(TRIM(?))\n              AND id <> ?\n            ', (nome, area_id))
        if cursor.fetchone():
            raise ValueError('Já existe outra área cadastrada com este nome.')
        cursor.execute('\n            UPDATE areas\n            SET\n                nome = ?,\n                descricao = ?,\n                email_area = ?,\n                atualizado_em = ?\n            WHERE id = ?\n            ', (nome, _texto(descricao), email_area, _agora(), area_id))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def alterar_status_area(area_id, ativa):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('SELECT id FROM areas WHERE id = ?', (area_id,))
        if not cursor.fetchone():
            raise ValueError('Área não encontrada.')
        cursor.execute('\n            UPDATE areas\n            SET\n                ativa = ?,\n                atualizado_em = ?\n            WHERE id = ?\n            ', (1 if ativa else 0, _agora(), area_id))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def listar_pessoas_cadastradas(incluir_inativas=True):
    conn = conectar()
    try:
        cursor = conn.cursor()
        filtro = '' if incluir_inativas else 'WHERE COALESCE(ativa, 1) = 1'
        cursor.execute(f"\n            SELECT\n                id,\n                nome,\n                email,\n                COALESCE(cargo, ''),\n                COALESCE(ativa, 1),\n                criado_em,\n                atualizado_em\n            FROM pessoas\n            {filtro}\n            ORDER BY\n                COALESCE(ativa, 1) DESC,\n                nome COLLATE NOCASE\n            ")
        return [{'id': item[0], 'nome': item[1], 'email': item[2], 'cargo': item[3], 'ativa': bool(item[4]), 'criado_em': item[5], 'atualizado_em': item[6]} for item in cursor.fetchall()]
    finally:
        conn.close()

def obter_pessoa(pessoa_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute("\n            SELECT\n                id,\n                nome,\n                email,\n                COALESCE(cargo, ''),\n                COALESCE(ativa, 1),\n                criado_em,\n                atualizado_em\n            FROM pessoas\n            WHERE id = ?\n            ", (pessoa_id,))
        item = cursor.fetchone()
        if not item:
            return None
        return {'id': item[0], 'nome': item[1], 'email': item[2], 'cargo': item[3], 'ativa': bool(item[4]), 'criado_em': item[5], 'atualizado_em': item[6]}
    finally:
        conn.close()

def cadastrar_pessoa(nome, email, cargo=''):
    nome = _validar_texto_obrigatorio(nome, 'o nome da pessoa')
    email = _validar_email_basico(email, obrigatorio=True, campo='o e-mail da pessoa')
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('\n            SELECT id\n            FROM pessoas\n            WHERE LOWER(TRIM(email)) = LOWER(TRIM(?))\n            ', (email,))
        if cursor.fetchone():
            raise ValueError('Já existe uma pessoa cadastrada com este e-mail.')
        agora = _agora()
        cursor.execute('\n            INSERT INTO pessoas (\n                nome,\n                email,\n                cargo,\n                ativa,\n                criado_em,\n                atualizado_em\n            )\n            VALUES (?, ?, ?, 1, ?, ?)\n            ', (nome, email, _texto(cargo), agora, agora))
        pessoa_id = cursor.lastrowid
        conn.commit()
        return pessoa_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def atualizar_pessoa(pessoa_id, nome, email, cargo=''):
    nome = _validar_texto_obrigatorio(nome, 'o nome da pessoa')
    email = _validar_email_basico(email, obrigatorio=True, campo='o e-mail da pessoa')
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('SELECT id FROM pessoas WHERE id = ?', (pessoa_id,))
        if not cursor.fetchone():
            raise ValueError('Pessoa não encontrada.')
        cursor.execute('\n            SELECT id\n            FROM pessoas\n            WHERE LOWER(TRIM(email)) = LOWER(TRIM(?))\n              AND id <> ?\n            ', (email, pessoa_id))
        if cursor.fetchone():
            raise ValueError('Já existe outra pessoa cadastrada com este e-mail.')
        cursor.execute('\n            UPDATE pessoas\n            SET\n                nome = ?,\n                email = ?,\n                cargo = ?,\n                atualizado_em = ?\n            WHERE id = ?\n            ', (nome, email, _texto(cargo), _agora(), pessoa_id))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def alterar_status_pessoa(pessoa_id, ativa):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('SELECT id FROM pessoas WHERE id = ?', (pessoa_id,))
        if not cursor.fetchone():
            raise ValueError('Pessoa não encontrada.')
        cursor.execute('\n            UPDATE pessoas\n            SET\n                ativa = ?,\n                atualizado_em = ?\n            WHERE id = ?\n            ', (1 if ativa else 0, _agora(), pessoa_id))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def vincular_pessoa_area(area_id, pessoa_id, papel='Informado', principal=False):
    if papel not in PAPEIS_ENVOLVIMENTO:
        raise ValueError('Papel de envolvimento inválido.')
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('\n            SELECT COALESCE(ativa, 1)\n            FROM areas\n            WHERE id = ?\n            ', (area_id,))
        area = cursor.fetchone()
        if not area:
            raise ValueError('Área não encontrada.')
        if not bool(area[0]):
            raise ValueError('A área está inativa.')
        cursor.execute('\n            SELECT COALESCE(ativa, 1)\n            FROM pessoas\n            WHERE id = ?\n            ', (pessoa_id,))
        pessoa = cursor.fetchone()
        if not pessoa:
            raise ValueError('Pessoa não encontrada.')
        if not bool(pessoa[0]):
            raise ValueError('A pessoa está inativa.')
        agora = _agora()
        if principal:
            cursor.execute('\n                UPDATE area_pessoas\n                SET\n                    principal = 0,\n                    atualizado_em = ?\n                WHERE area_id = ?\n                ', (agora, area_id))
        cursor.execute('\n            INSERT INTO area_pessoas (\n                area_id,\n                pessoa_id,\n                papel,\n                principal,\n                criado_em,\n                atualizado_em\n            )\n            VALUES (?, ?, ?, ?, ?, ?)\n            ON CONFLICT(area_id, pessoa_id)\n            DO UPDATE SET\n                papel = excluded.papel,\n                principal = excluded.principal,\n                atualizado_em = excluded.atualizado_em\n            ', (area_id, pessoa_id, papel, 1 if principal else 0, agora, agora))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def remover_vinculo_pessoa_area(area_id, pessoa_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute('\n            DELETE FROM area_pessoas\n            WHERE area_id = ?\n              AND pessoa_id = ?\n            ', (area_id, pessoa_id))
        if cursor.rowcount == 0:
            raise ValueError('Vínculo entre área e pessoa não encontrado.')
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def listar_pessoas_area(area_id, somente_ativas=False):
    conn = conectar()
    try:
        cursor = conn.cursor()
        filtro_ativo = 'AND COALESCE(p.ativa, 1) = 1' if somente_ativas else ''
        cursor.execute(f"\n            SELECT\n                p.id,\n                p.nome,\n                p.email,\n                COALESCE(p.cargo, ''),\n                COALESCE(p.ativa, 1),\n                ap.papel,\n                COALESCE(ap.principal, 0),\n                ap.criado_em,\n                ap.atualizado_em\n            FROM area_pessoas ap\n            INNER JOIN pessoas p\n                ON p.id = ap.pessoa_id\n            WHERE ap.area_id = ?\n            {filtro_ativo}\n            ORDER BY\n                COALESCE(ap.principal, 0) DESC,\n                p.nome COLLATE NOCASE\n            ", (area_id,))
        return [{'id': item[0], 'nome': item[1], 'email': item[2], 'cargo': item[3], 'ativa': bool(item[4]), 'papel': item[5], 'principal': bool(item[6]), 'criado_em': item[7], 'atualizado_em': item[8]} for item in cursor.fetchall()]
    finally:
        conn.close()

def listar_areas_pessoa(pessoa_id, somente_ativas=False):
    conn = conectar()
    try:
        cursor = conn.cursor()
        filtro_ativo = 'AND COALESCE(a.ativa, 1) = 1' if somente_ativas else ''
        cursor.execute(f"\n            SELECT\n                a.id,\n                a.nome,\n                COALESCE(a.descricao, ''),\n                COALESCE(a.email_area, ''),\n                COALESCE(a.ativa, 1),\n                ap.papel,\n                COALESCE(ap.principal, 0)\n            FROM area_pessoas ap\n            INNER JOIN areas a\n                ON a.id = ap.area_id\n            WHERE ap.pessoa_id = ?\n            {filtro_ativo}\n            ORDER BY\n                a.nome COLLATE NOCASE\n            ", (pessoa_id,))
        return [{'id': item[0], 'nome': item[1], 'descricao': item[2], 'email_area': item[3], 'ativa': bool(item[4]), 'papel': item[5], 'principal': bool(item[6])} for item in cursor.fetchall()]
    finally:
        conn.close()

def obter_destinatarios_area(area_id):
    area = obter_area(area_id)
    if not area:
        raise ValueError('Área não encontrada.')
    destinatarios = []
    if area['ativa'] and area['email_area']:
        destinatarios.append({'nome': area['nome'], 'email': area['email_area'], 'origem': 'Área', 'papel': 'E-mail da área', 'principal': False})
    for pessoa in listar_pessoas_area(area_id, somente_ativas=True):
        destinatarios.append({'nome': pessoa['nome'], 'email': pessoa['email'], 'origem': 'Pessoa', 'papel': pessoa['papel'], 'principal': pessoa['principal']})
    unicos = []
    emails_vistos = set()
    for item in destinatarios:
        email = item['email'].strip().lower()
        if not email or email in emails_vistos:
            continue
        emails_vistos.add(email)
        unicos.append(item)
    return unicos
_diagnostico_banco_base_1_9d_b2 = diagnostico_banco

def diagnostico_banco():
    resultado = _diagnostico_banco_base_1_9d_b2()
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM areas')
        resultado['areas'] = cursor.fetchone()[0]
        cursor.execute('SELECT COUNT(*) FROM pessoas')
        resultado['pessoas'] = cursor.fetchone()[0]
        cursor.execute('SELECT COUNT(*) FROM area_pessoas')
        resultado['vinculos_area_pessoas'] = cursor.fetchone()[0]
    finally:
        conn.close()
    return resultado

# ============================================================
# COMUNICAÇÕES
# FASE 2.0B1
# ============================================================

STATUS_COMUNICACAO = [
    'Rascunho',
    'Preparada',
    'Enviando',
    'Enviada',
    'Falhou'
]

TIPOS_COMUNICACAO = [
    'Atualização',
    'Prazo',
    'Solicitação de informação',
    'Nova intimação',
    'Resposta protocolada',
    'Encerramento',
    'Outros'
]


def garantir_estrutura_comunicacoes():
    """
    Cria somente as estruturas da fase 2.0B1.

    Esta função é propositalmente independente de criar_tabelas(),
    executar_migracoes() e criar_indices() para não alterar a lógica
    já estabilizada das fases anteriores.
    """
    conn = conectar()

    try:
        cursor = conn.cursor()

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS comunicacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fiscalizacao_id INTEGER NOT NULL,
                tipo TEXT NOT NULL DEFAULT 'Atualização',
                assunto TEXT NOT NULL,
                mensagem TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'Rascunho',
                erro_envio TEXT,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                enviado_em TEXT,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (
                    fiscalizacao_id
                )
                REFERENCES fiscalizacoes(id)
                ON DELETE CASCADE
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS comunicacao_destinatarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                comunicacao_id INTEGER NOT NULL,
                nome TEXT,
                email TEXT NOT NULL,
                origem TEXT,
                papel TEXT,
                status_envio TEXT NOT NULL DEFAULT 'Pendente',
                erro_envio TEXT,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (
                    comunicacao_id
                )
                REFERENCES comunicacoes(id)
                ON DELETE CASCADE
            )
            """
        )

        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS
            idx_comunicacoes_fiscalizacao
            ON comunicacoes(fiscalizacao_id)
            """
        )

        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS
            idx_comunicacoes_status
            ON comunicacoes(status)
            """
        )

        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS
            idx_comunicacao_destinatarios_comunicacao
            ON comunicacao_destinatarios(comunicacao_id)
            """
        )

        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS
            idx_comunicacao_destinatarios_email
            ON comunicacao_destinatarios(email)
            """
        )

        conn.commit()

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


def _validar_status_comunicacao(status):
    if status not in STATUS_COMUNICACAO:
        raise ValueError('Status de comunicação inválido.')
    return status


def _validar_tipo_comunicacao(tipo):
    if tipo not in TIPOS_COMUNICACAO:
        raise ValueError('Tipo de comunicação inválido.')
    return tipo


def _normalizar_destinatarios(destinatarios):
    if not destinatarios:
        raise ValueError('Informe pelo menos um destinatário.')

    normalizados = []
    emails_vistos = set()

    for item in destinatarios:
        if isinstance(item, str):
            nome = ''
            email = item
            origem = 'Manual'
            papel = ''

        elif isinstance(item, dict):
            nome = _texto(item.get('nome'))
            email = item.get('email')
            origem = _texto(item.get('origem') or 'Manual')
            papel = _texto(item.get('papel'))

        else:
            raise ValueError('Destinatário inválido.')

        email = _validar_email_basico(
            email,
            obrigatorio=True,
            campo='o e-mail do destinatário'
        )

        chave = email.lower()

        if chave in emails_vistos:
            continue

        emails_vistos.add(chave)

        normalizados.append(
            {
                'nome': nome,
                'email': email,
                'origem': origem,
                'papel': papel
            }
        )

    if not normalizados:
        raise ValueError('Nenhum destinatário válido foi informado.')

    return normalizados


def obter_destinatarios_fiscalizacao(fiscalizacao_id):
    """
    Retorna os e-mails das áreas vinculadas à fiscalização e das
    pessoas ativas vinculadas a essas áreas, sem duplicidade.
    """
    conn = conectar()

    try:
        cursor = conn.cursor()

        _buscar_estado_fiscalizacao(
            cursor,
            fiscalizacao_id
        )

        cursor.execute(
            """
            SELECT DISTINCT
                a.id
            FROM fiscalizacao_areas fa
            INNER JOIN areas a
                ON a.id = fa.area_id
            WHERE
                fa.fiscalizacao_id = ?
                AND COALESCE(a.ativa, 1) = 1
            ORDER BY a.nome
            """,
            (fiscalizacao_id,)
        )

        area_ids = [
            item[0]
            for item in cursor.fetchall()
        ]

    finally:
        conn.close()

    destinatarios = []
    emails_vistos = set()

    for area_id in area_ids:
        for item in obter_destinatarios_area(area_id):
            email = item['email'].strip().lower()

            if not email or email in emails_vistos:
                continue

            emails_vistos.add(email)
            destinatarios.append(item)

    return destinatarios


def criar_comunicacao(
    fiscalizacao_id,
    assunto,
    mensagem,
    destinatarios,
    tipo='Atualização',
    status='Rascunho'
):
    """
    Registra uma comunicação e congela a lista de destinatários
    utilizada naquele momento. Ainda não envia e-mail.
    """
    assunto = _validar_texto_obrigatorio(
        assunto,
        'o assunto da comunicação'
    )

    mensagem = _validar_texto_obrigatorio(
        mensagem,
        'a mensagem da comunicação'
    )

    tipo = _validar_tipo_comunicacao(tipo)
    status = _validar_status_comunicacao(status)
    destinatarios = _normalizar_destinatarios(destinatarios)

    conn = conectar()

    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')

        _garantir_nao_arquivada(
            cursor,
            fiscalizacao_id
        )

        agora = _agora()

        cursor.execute(
            """
            INSERT INTO comunicacoes (
                fiscalizacao_id,
                tipo,
                assunto,
                mensagem,
                status,
                erro_envio,
                criado_em,
                enviado_em,
                atualizado_em
            )
            VALUES (?, ?, ?, ?, ?, NULL, ?, NULL, ?)
            """,
            (
                fiscalizacao_id,
                tipo,
                assunto,
                mensagem,
                status,
                agora,
                agora
            )
        )

        comunicacao_id = cursor.lastrowid

        for destinatario in destinatarios:
            cursor.execute(
                """
                INSERT INTO comunicacao_destinatarios (
                    comunicacao_id,
                    nome,
                    email,
                    origem,
                    papel,
                    status_envio,
                    erro_envio,
                    criado_em,
                    atualizado_em
                )
                VALUES (?, ?, ?, ?, ?, 'Pendente', NULL, ?, ?)
                """,
                (
                    comunicacao_id,
                    destinatario['nome'],
                    destinatario['email'],
                    destinatario['origem'],
                    destinatario['papel'],
                    agora,
                    agora
                )
            )

        conn.commit()
        return comunicacao_id

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


def obter_comunicacao(comunicacao_id):
    conn = conectar()

    try:
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT
                id,
                fiscalizacao_id,
                tipo,
                assunto,
                mensagem,
                status,
                COALESCE(erro_envio, ''),
                criado_em,
                enviado_em,
                atualizado_em
            FROM comunicacoes
            WHERE id = ?
            """,
            (comunicacao_id,)
        )

        item = cursor.fetchone()

        if not item:
            return None

        cursor.execute(
            """
            SELECT
                id,
                nome,
                email,
                COALESCE(origem, ''),
                COALESCE(papel, ''),
                status_envio,
                COALESCE(erro_envio, ''),
                criado_em,
                atualizado_em
            FROM comunicacao_destinatarios
            WHERE comunicacao_id = ?
            ORDER BY
                nome COLLATE NOCASE,
                email COLLATE NOCASE
            """,
            (comunicacao_id,)
        )

        destinatarios = [
            {
                'id': dest[0],
                'nome': dest[1] or '',
                'email': dest[2],
                'origem': dest[3],
                'papel': dest[4],
                'status_envio': dest[5],
                'erro_envio': dest[6],
                'criado_em': dest[7],
                'atualizado_em': dest[8]
            }
            for dest in cursor.fetchall()
        ]

        return {
            'id': item[0],
            'fiscalizacao_id': item[1],
            'tipo': item[2],
            'assunto': item[3],
            'mensagem': item[4],
            'status': item[5],
            'erro_envio': item[6],
            'criado_em': item[7],
            'enviado_em': item[8],
            'atualizado_em': item[9],
            'destinatarios': destinatarios
        }

    finally:
        conn.close()


def listar_comunicacoes_fiscalizacao(fiscalizacao_id):
    conn = conectar()

    try:
        cursor = conn.cursor()

        _buscar_estado_fiscalizacao(
            cursor,
            fiscalizacao_id
        )

        cursor.execute(
            """
            SELECT
                c.id,
                c.tipo,
                c.assunto,
                c.status,
                COALESCE(c.erro_envio, ''),
                c.criado_em,
                c.enviado_em,
                c.atualizado_em,
                COUNT(cd.id)
            FROM comunicacoes c
            LEFT JOIN comunicacao_destinatarios cd
                ON cd.comunicacao_id = c.id
            WHERE c.fiscalizacao_id = ?
            GROUP BY
                c.id,
                c.tipo,
                c.assunto,
                c.status,
                c.erro_envio,
                c.criado_em,
                c.enviado_em,
                c.atualizado_em
            ORDER BY c.id DESC
            """,
            (fiscalizacao_id,)
        )

        return [
            {
                'id': item[0],
                'tipo': item[1],
                'assunto': item[2],
                'status': item[3],
                'erro_envio': item[4],
                'criado_em': item[5],
                'enviado_em': item[6],
                'atualizado_em': item[7],
                'quantidade_destinatarios': item[8]
            }
            for item in cursor.fetchall()
        ]

    finally:
        conn.close()


def atualizar_status_comunicacao(
    comunicacao_id,
    status,
    erro_envio='',
    enviado_em=None
):
    status = _validar_status_comunicacao(status)

    if status == 'Enviada' and not enviado_em:
        enviado_em = _agora()

    conn = conectar()

    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')

        cursor.execute(
            """
            SELECT id
            FROM comunicacoes
            WHERE id = ?
            """,
            (comunicacao_id,)
        )

        if not cursor.fetchone():
            raise ValueError('Comunicação não encontrada.')

        cursor.execute(
            """
            UPDATE comunicacoes
            SET
                status = ?,
                erro_envio = ?,
                enviado_em = ?,
                atualizado_em = ?
            WHERE id = ?
            """,
            (
                status,
                _texto(erro_envio),
                enviado_em,
                _agora(),
                comunicacao_id
            )
        )

        conn.commit()
        return True

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


def atualizar_status_destinatario(
    destinatario_id,
    status_envio,
    erro_envio=''
):
    permitidos = [
        'Pendente',
        'Enviado',
        'Falhou'
    ]

    if status_envio not in permitidos:
        raise ValueError(
            'Status de envio do destinatário inválido.'
        )

    conn = conectar()

    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')

        cursor.execute(
            """
            SELECT id
            FROM comunicacao_destinatarios
            WHERE id = ?
            """,
            (destinatario_id,)
        )

        if not cursor.fetchone():
            raise ValueError(
                'Destinatário da comunicação não encontrado.'
            )

        cursor.execute(
            """
            UPDATE comunicacao_destinatarios
            SET
                status_envio = ?,
                erro_envio = ?,
                atualizado_em = ?
            WHERE id = ?
            """,
            (
                status_envio,
                _texto(erro_envio),
                _agora(),
                destinatario_id
            )
        )

        conn.commit()
        return True

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


def excluir_comunicacao_rascunho(comunicacao_id):
    conn = conectar()

    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')

        cursor.execute(
            """
            SELECT status
            FROM comunicacoes
            WHERE id = ?
            """,
            (comunicacao_id,)
        )

        resultado = cursor.fetchone()

        if not resultado:
            raise ValueError('Comunicação não encontrada.')

        if resultado[0] not in [
            'Rascunho',
            'Preparada'
        ]:
            raise ValueError(
                'Somente comunicações ainda não enviadas '
                'podem ser excluídas.'
            )

        cursor.execute(
            """
            DELETE FROM comunicacoes
            WHERE id = ?
            """,
            (comunicacao_id,)
        )

        conn.commit()
        return True

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


def obter_resumo_comunicacoes():
    """
    Diagnóstico simples da estrutura 2.0B1 sem modificar
    o diagnostico_banco() já estável.
    """
    conn = conectar()

    try:
        cursor = conn.cursor()

        cursor.execute(
            'SELECT COUNT(*) FROM comunicacoes'
        )
        comunicacoes = cursor.fetchone()[0]

        cursor.execute(
            'SELECT COUNT(*) FROM comunicacao_destinatarios'
        )
        destinatarios = cursor.fetchone()[0]

        return {
            'comunicacoes': comunicacoes,
            'destinatarios': destinatarios
        }

    finally:
        conn.close()



# ============================================================
# APROVACOES DE ENCERRAMENTO
# FASE 2.0D1
# ============================================================

STATUS_APROVACAO_ENCERRAMENTO = [
    'Pendente',
    'Ciente',
]


def garantir_estrutura_aprovacoes_encerramento():
    """Cria as estruturas de governanca para encerramento."""
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS encerramento_aprovadores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fiscalizacao_id INTEGER NOT NULL,
                pessoa_id INTEGER NOT NULL,
                area_id INTEGER NOT NULL,
                papel TEXT,
                essencial INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'Pendente',
                observacao_ciencia TEXT,
                ciente_em TEXT,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (fiscalizacao_id, pessoa_id, area_id),
                FOREIGN KEY (fiscalizacao_id)
                    REFERENCES fiscalizacoes(id) ON DELETE CASCADE,
                FOREIGN KEY (pessoa_id)
                    REFERENCES pessoas(id) ON DELETE RESTRICT,
                FOREIGN KEY (area_id)
                    REFERENCES areas(id) ON DELETE RESTRICT
            )
            """
        )
        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_encerramento_aprovadores_fiscalizacao
            ON encerramento_aprovadores(fiscalizacao_id)
            """
        )
        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_encerramento_aprovadores_status
            ON encerramento_aprovadores(status, essencial)
            """
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def listar_candidatos_aprovacao_encerramento(fiscalizacao_id):
    """Lista pessoas ativas das areas atualmente envolvidas no caso."""
    conn = conectar()
    try:
        cursor = conn.cursor()
        _buscar_estado_fiscalizacao(cursor, fiscalizacao_id)
        cursor.execute(
            """
            SELECT DISTINCT
                p.id,
                p.nome,
                p.email,
                COALESCE(p.cargo, ''),
                a.id,
                a.nome,
                ap.papel,
                COALESCE(ap.principal, 0)
            FROM fiscalizacao_areas fa
            INNER JOIN areas a ON a.id = fa.area_id
            INNER JOIN area_pessoas ap ON ap.area_id = a.id
            INNER JOIN pessoas p ON p.id = ap.pessoa_id
            WHERE fa.fiscalizacao_id = ?
              AND COALESCE(a.ativa, 1) = 1
              AND COALESCE(p.ativa, 1) = 1
            ORDER BY a.nome COLLATE NOCASE,
                     COALESCE(ap.principal, 0) DESC,
                     p.nome COLLATE NOCASE
            """,
            (fiscalizacao_id,),
        )
        return [
            {
                'pessoa_id': item[0],
                'nome': item[1],
                'email': item[2],
                'cargo': item[3],
                'area_id': item[4],
                'area': item[5],
                'papel': item[6],
                'principal': bool(item[7]),
            }
            for item in cursor.fetchall()
        ]
    finally:
        conn.close()


def definir_aprovador_encerramento(
    fiscalizacao_id,
    pessoa_id,
    area_id,
    essencial=False,
    papel=None,
):
    """Inclui ou atualiza um envolvido como aprovador do encerramento."""
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        estado = _garantir_nao_arquivada(cursor, fiscalizacao_id)
        if estado['fase'] == 'Encerrada':
            raise ValueError('A fiscalizacao ja esta encerrada.')

        cursor.execute(
            """
            SELECT p.nome, a.nome, ap.papel
            FROM fiscalizacao_areas fa
            INNER JOIN areas a ON a.id = fa.area_id
            INNER JOIN area_pessoas ap ON ap.area_id = a.id
            INNER JOIN pessoas p ON p.id = ap.pessoa_id
            WHERE fa.fiscalizacao_id = ?
              AND p.id = ?
              AND a.id = ?
              AND COALESCE(a.ativa, 1) = 1
              AND COALESCE(p.ativa, 1) = 1
            """,
            (fiscalizacao_id, pessoa_id, area_id),
        )
        vinculo = cursor.fetchone()
        if not vinculo:
            raise ValueError(
                'A pessoa precisa estar ativa e vinculada a uma area envolvida nesta fiscalizacao.'
            )

        papel_final = _texto(papel) or _texto(vinculo[2]) or 'Aprovador'
        agora = _agora()
        cursor.execute(
            """
            INSERT INTO encerramento_aprovadores (
                fiscalizacao_id, pessoa_id, area_id, papel, essencial,
                status, observacao_ciencia, ciente_em, criado_em, atualizado_em
            )
            VALUES (?, ?, ?, ?, ?, 'Pendente', '', NULL, ?, ?)
            ON CONFLICT(fiscalizacao_id, pessoa_id, area_id)
            DO UPDATE SET
                papel = excluded.papel,
                essencial = excluded.essencial,
                atualizado_em = excluded.atualizado_em
            """,
            (
                fiscalizacao_id,
                pessoa_id,
                area_id,
                papel_final,
                1 if essencial else 0,
                agora,
                agora,
            ),
        )
        cursor.execute(
            """
            SELECT id
            FROM encerramento_aprovadores
            WHERE fiscalizacao_id = ? AND pessoa_id = ? AND area_id = ?
            """,
            (fiscalizacao_id, pessoa_id, area_id),
        )
        aprovador_id = cursor.fetchone()[0]
        conn.commit()
        return aprovador_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def remover_aprovador_encerramento(aprovador_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute(
            """
            SELECT ea.fiscalizacao_id, ea.status, f.fase, COALESCE(f.arquivada, 0)
            FROM encerramento_aprovadores ea
            INNER JOIN fiscalizacoes f ON f.id = ea.fiscalizacao_id
            WHERE ea.id = ?
            """,
            (aprovador_id,),
        )
        item = cursor.fetchone()
        if not item:
            raise ValueError('Aprovador nao encontrado.')
        if item[3]:
            raise ValueError('A fiscalizacao esta arquivada.')
        if item[2] == 'Encerrada':
            raise ValueError('Nao e possivel alterar aprovadores apos o encerramento.')
        if item[1] == 'Ciente':
            raise ValueError('Revogue a ciencia antes de remover este aprovador.')
        cursor.execute(
            'DELETE FROM encerramento_aprovadores WHERE id = ?',
            (aprovador_id,),
        )
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def listar_aprovacoes_encerramento(fiscalizacao_id):
    conn = conectar()
    try:
        cursor = conn.cursor()
        _buscar_estado_fiscalizacao(cursor, fiscalizacao_id)
        cursor.execute(
            """
            SELECT
                ea.id,
                ea.pessoa_id,
                p.nome,
                p.email,
                COALESCE(p.cargo, ''),
                ea.area_id,
                a.nome,
                COALESCE(ea.papel, ''),
                COALESCE(ea.essencial, 0),
                ea.status,
                COALESCE(ea.observacao_ciencia, ''),
                ea.ciente_em,
                ea.criado_em,
                ea.atualizado_em
            FROM encerramento_aprovadores ea
            INNER JOIN pessoas p ON p.id = ea.pessoa_id
            INNER JOIN areas a ON a.id = ea.area_id
            WHERE ea.fiscalizacao_id = ?
            ORDER BY COALESCE(ea.essencial, 0) DESC,
                     a.nome COLLATE NOCASE,
                     p.nome COLLATE NOCASE
            """,
            (fiscalizacao_id,),
        )
        return [
            {
                'id': item[0],
                'pessoa_id': item[1],
                'nome': item[2],
                'email': item[3],
                'cargo': item[4],
                'area_id': item[5],
                'area': item[6],
                'papel': item[7],
                'essencial': bool(item[8]),
                'status': item[9],
                'observacao_ciencia': item[10],
                'ciente_em': item[11],
                'criado_em': item[12],
                'atualizado_em': item[13],
            }
            for item in cursor.fetchall()
        ]
    finally:
        conn.close()


def registrar_ciencia_encerramento(aprovador_id, observacao=''):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute(
            """
            SELECT ea.fiscalizacao_id, ea.status, p.nome, a.nome, ea.papel
            FROM encerramento_aprovadores ea
            INNER JOIN pessoas p ON p.id = ea.pessoa_id
            INNER JOIN areas a ON a.id = ea.area_id
            WHERE ea.id = ?
            """,
            (aprovador_id,),
        )
        item = cursor.fetchone()
        if not item:
            raise ValueError('Aprovador nao encontrado.')

        fiscalizacao_id, status_atual, nome, area, papel = item
        estado = _garantir_nao_arquivada(cursor, fiscalizacao_id)
        if estado['fase'] == 'Encerrada':
            raise ValueError('A fiscalizacao ja esta encerrada.')
        if status_atual == 'Ciente':
            raise ValueError('A ciencia deste aprovador ja foi registrada.')

        agora = _agora()
        cursor.execute(
            """
            UPDATE encerramento_aprovadores
            SET status = 'Ciente',
                observacao_ciencia = ?,
                ciente_em = ?,
                atualizado_em = ?
            WHERE id = ?
            """,
            (_texto(observacao), agora, agora, aprovador_id),
        )
        _registrar_historico(
            cursor,
            fiscalizacao_id,
            estado['fase'],
            f'Ciencia de encerramento registrada por {nome} - {area} ({papel or "Aprovador"}).',
        )
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def revogar_ciencia_encerramento(aprovador_id, observacao=''):
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('BEGIN IMMEDIATE')
        cursor.execute(
            """
            SELECT ea.fiscalizacao_id, ea.status, p.nome, a.nome
            FROM encerramento_aprovadores ea
            INNER JOIN pessoas p ON p.id = ea.pessoa_id
            INNER JOIN areas a ON a.id = ea.area_id
            WHERE ea.id = ?
            """,
            (aprovador_id,),
        )
        item = cursor.fetchone()
        if not item:
            raise ValueError('Aprovador nao encontrado.')

        fiscalizacao_id, status_atual, nome, area = item
        estado = _garantir_nao_arquivada(cursor, fiscalizacao_id)
        if estado['fase'] == 'Encerrada':
            raise ValueError('Nao e possivel revogar ciencia apos o encerramento.')
        if status_atual != 'Ciente':
            raise ValueError('Este aprovador ainda nao registrou ciencia.')

        agora = _agora()
        cursor.execute(
            """
            UPDATE encerramento_aprovadores
            SET status = 'Pendente',
                observacao_ciencia = '',
                ciente_em = NULL,
                atualizado_em = ?
            WHERE id = ?
            """,
            (agora, aprovador_id),
        )
        texto = f'Ciencia de encerramento revogada para {nome} - {area}.'
        if _texto(observacao):
            texto += f' Motivo: {_texto(observacao)}'
        _registrar_historico(
            cursor,
            fiscalizacao_id,
            estado['fase'],
            texto,
        )
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def obter_status_encerramento(fiscalizacao_id):
    aprovacoes = listar_aprovacoes_encerramento(fiscalizacao_id)
    essenciais = [item for item in aprovacoes if item['essencial']]
    pendentes = [item for item in essenciais if item['status'] != 'Ciente']
    cientes = [item for item in essenciais if item['status'] == 'Ciente']
    return {
        'total_aprovadores': len(aprovacoes),
        'total_essenciais': len(essenciais),
        'essenciais_cientes': len(cientes),
        'essenciais_pendentes': len(pendentes),
        'pode_encerrar': bool(essenciais) and not pendentes,
        'pendentes': pendentes,
        'aprovacoes': aprovacoes,
    }


def validar_encerramento_fiscalizacao(fiscalizacao_id):
    status = obter_status_encerramento(fiscalizacao_id)
    if status['total_essenciais'] == 0:
        raise ValueError(
            'O encerramento exige pelo menos um aprovador essencial cadastrado.'
        )
    if status['essenciais_pendentes'] > 0:
        nomes = ', '.join(
            f"{item['nome']} ({item['area']})"
            for item in status['pendentes']
        )
        raise ValueError(
            'O encerramento esta bloqueado. Falta ciencia dos aprovadores essenciais: '
            + nomes
        )
    return True


_alterar_fase_fiscalizacao_base_2_0b1 = alterar_fase_fiscalizacao


def alterar_fase_fiscalizacao(fiscalizacao_id, nova_fase, observacao=''):
    """Bloqueia a fase Encerrada sem a governanca exigida pela 2.0D1."""
    nova_fase = _validar_fase(nova_fase)
    if nova_fase == 'Encerrada':
        validar_encerramento_fiscalizacao(fiscalizacao_id)
    return _alterar_fase_fiscalizacao_base_2_0b1(
        fiscalizacao_id,
        nova_fase,
        observacao,
    )


def obter_resumo_aprovacoes_encerramento():
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM encerramento_aprovadores')
        total = cursor.fetchone()[0]
        cursor.execute(
            """
            SELECT COUNT(*)
            FROM encerramento_aprovadores
            WHERE essencial = 1
            """
        )
        essenciais = cursor.fetchone()[0]
        cursor.execute(
            """
            SELECT COUNT(*)
            FROM encerramento_aprovadores
            WHERE essencial = 1 AND status = 'Ciente'
            """
        )
        essenciais_cientes = cursor.fetchone()[0]
        return {
            'total': total,
            'essenciais': essenciais,
            'essenciais_cientes': essenciais_cientes,
        }
    finally:
        conn.close()



# ============================================================
# FLASH REPORT EXECUTIVO
# FASE 2.1A
# ============================================================

PRIORIDADES_EXECUTIVAS = ["Alta", "Média", "Baixa"]


def _validar_prioridade_executiva(valor):
    valor = _texto(valor) or "Média"
    if valor not in PRIORIDADES_EXECUTIVAS:
        raise ValueError("Prioridade executiva inválida.")
    return valor


def garantir_estrutura_flash_report_executivo():
    novas_colunas = [
        ("incluir_flash_report", "INTEGER NOT NULL DEFAULT 0"),
        ("prioridade_executiva", "TEXT NOT NULL DEFAULT 'Média'"),
        ("atualizacao_executiva", "TEXT"),
    ]
    for coluna, definicao in novas_colunas:
        if not coluna_existe("fiscalizacoes", coluna):
            conn = conectar()
            try:
                conn.execute(
                    f"ALTER TABLE fiscalizacoes ADD COLUMN {coluna} {definicao}"
                )
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()

    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            UPDATE fiscalizacoes
            SET incluir_flash_report = 0
            WHERE incluir_flash_report IS NULL
            """
        )
        cursor.execute(
            """
            UPDATE fiscalizacoes
            SET prioridade_executiva = 'Média'
            WHERE prioridade_executiva IS NULL
               OR TRIM(prioridade_executiva) = ''
            """
        )
        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_fiscalizacoes_flash_report
            ON fiscalizacoes(incluir_flash_report, prioridade_executiva)
            """
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def obter_configuracao_flash_report_executivo(fiscalizacao_id):
    garantir_estrutura_flash_report_executivo()
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT
                COALESCE(incluir_flash_report, 0),
                COALESCE(prioridade_executiva, 'Média'),
                COALESCE(atualizacao_executiva, '')
            FROM fiscalizacoes
            WHERE id = ?
            """,
            (fiscalizacao_id,),
        )
        resultado = cursor.fetchone()
        if not resultado:
            return None
        return {
            "incluir_flash_report": bool(resultado[0]),
            "prioridade_executiva": resultado[1],
            "atualizacao_executiva": resultado[2],
        }
    finally:
        conn.close()


def atualizar_configuracao_flash_report_executivo(
    fiscalizacao_id,
    incluir_flash_report=False,
    prioridade_executiva="Média",
    atualizacao_executiva="",
):
    garantir_estrutura_flash_report_executivo()
    prioridade_executiva = _validar_prioridade_executiva(
        prioridade_executiva
    )
    atualizacao_executiva = _texto(atualizacao_executiva)

    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute("BEGIN IMMEDIATE")
        _garantir_nao_arquivada(cursor, fiscalizacao_id)
        cursor.execute(
            """
            UPDATE fiscalizacoes
            SET
                incluir_flash_report = ?,
                prioridade_executiva = ?,
                atualizacao_executiva = ?,
                atualizado_em = ?
            WHERE id = ?
            """,
            (
                1 if incluir_flash_report else 0,
                prioridade_executiva,
                atualizacao_executiva,
                _agora(),
                fiscalizacao_id,
            ),
        )
        if cursor.rowcount != 1:
            raise ValueError("Fiscalização não encontrada.")
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def listar_flash_report_executivo(incluir_arquivadas=False):
    garantir_estrutura_flash_report_executivo()
    garantir_flash_reports()
    filtro_arquivada = (
        ""
        if incluir_arquivadas
        else "AND COALESCE(f.arquivada, 0) = 0"
    )

    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute(
            f"""
            SELECT
                f.id,
                f.codigo,
                f.titulo,
                f.empresa,
                f.orgao,
                f.exposicao,
                f.probabilidade,
                f.fase,
                f.prazo_resposta,
                COALESCE(f.prioridade_executiva, 'Média'),
                COALESCE(f.atualizacao_executiva, ''),
                COALESCE(fr.resumo_executivo, ''),
                COALESCE(fr.proximos_passos, ''),
                f.atualizado_em,
                (
                    SELECT GROUP_CONCAT(nome, ', ')
                    FROM (
                        SELECT t.nome AS nome
                        FROM fiscalizacao_tributos ft
                        INNER JOIN tributos t ON t.id = ft.tributo_id
                        WHERE ft.fiscalizacao_id = f.id
                        ORDER BY t.nome COLLATE NOCASE
                    )
                ) AS tributos,
                (
                    SELECT COUNT(*)
                    FROM documentos_fiscalizacao d
                    WHERE d.fiscalizacao_id = f.id
                ) AS quantidade_documentos
            FROM fiscalizacoes f
            LEFT JOIN flash_reports fr
                ON fr.fiscalizacao_id = f.id
            WHERE COALESCE(f.incluir_flash_report, 0) = 1
            {filtro_arquivada}
            ORDER BY
                CASE COALESCE(f.prioridade_executiva, 'Média')
                    WHEN 'Alta' THEN 1
                    WHEN 'Média' THEN 2
                    WHEN 'Baixa' THEN 3
                    ELSE 4
                END,
                CASE WHEN f.prazo_resposta IS NULL THEN 1 ELSE 0 END,
                f.prazo_resposta ASC,
                f.exposicao DESC,
                f.id DESC
            """
        )
        resultados = cursor.fetchall()
    finally:
        conn.close()

    saida = []
    for item in resultados:
        status_prazo = calcular_status_prazo(item[8], item[7])
        atualizacao = item[10].strip() if item[10] else ""
        resumo_executivo = item[11].strip() if item[11] else ""
        saida.append(
            {
                "id": item[0],
                "codigo": item[1],
                "titulo": item[2],
                "empresa": item[3],
                "orgao": item[4],
                "exposicao": item[5] or 0,
                "probabilidade": item[6],
                "fase": item[7],
                "prazo_resposta": item[8],
                "prioridade_executiva": item[9],
                "atualizacao_executiva": atualizacao,
                "resumo_executivo": resumo_executivo,
                "texto_executivo": atualizacao or resumo_executivo,
                "proximos_passos": item[12] or "",
                "atualizado_em": item[13],
                "tributos": item[14] or "",
                "quantidade_documentos": item[15] or 0,
                "status_prazo": status_prazo["status"],
                "dias_restantes": status_prazo["dias_restantes"],
            }
        )
    return saida


_cadastrar_fiscalizacao_base_2_0d4 = cadastrar_fiscalizacao


def cadastrar_fiscalizacao(
    titulo, objeto, empresa, filial, orgao, exposicao, probabilidade,
    tributos=None, areas=None, data_recebimento=None, prazo_resposta=None,
    responsavel_principal=None, tarefa_estracta=None,
    incluir_flash_report=False, prioridade_executiva="Média",
    atualizacao_executiva="",
):
    prioridade_executiva = _validar_prioridade_executiva(
        prioridade_executiva
    )
    atualizacao_executiva = _texto(atualizacao_executiva)
    codigo = _cadastrar_fiscalizacao_base_2_0d4(
        titulo, objeto, empresa, filial, orgao, exposicao, probabilidade,
        tributos=tributos, areas=areas,
        data_recebimento=data_recebimento, prazo_resposta=prazo_resposta,
        responsavel_principal=responsavel_principal,
        tarefa_estracta=tarefa_estracta,
    )
    fiscalizacao = buscar_fiscalizacao_por_codigo(codigo)
    if not fiscalizacao:
        raise RuntimeError(
            "Fiscalização criada, mas não foi possível recuperar o registro."
        )
    atualizar_configuracao_flash_report_executivo(
        fiscalizacao[0],
        incluir_flash_report=incluir_flash_report,
        prioridade_executiva=prioridade_executiva,
        atualizacao_executiva=atualizacao_executiva,
    )
    return codigo


_atualizar_dados_fiscalizacao_base_2_0d4 = atualizar_dados_fiscalizacao


def atualizar_dados_fiscalizacao(
    fiscalizacao_id, titulo, objeto, empresa, filial, orgao, exposicao,
    probabilidade, tributos=None, areas=None, responsavel_principal=None,
    tarefa_estracta=None, incluir_flash_report=None,
    prioridade_executiva=None, atualizacao_executiva=None,
):
    resultado = _atualizar_dados_fiscalizacao_base_2_0d4(
        fiscalizacao_id, titulo, objeto, empresa, filial, orgao, exposicao,
        probabilidade, tributos=tributos, areas=areas,
        responsavel_principal=responsavel_principal,
        tarefa_estracta=tarefa_estracta,
    )
    if (
        incluir_flash_report is not None
        or prioridade_executiva is not None
        or atualizacao_executiva is not None
    ):
        atual = obter_configuracao_flash_report_executivo(fiscalizacao_id)
        if not atual:
            raise ValueError("Fiscalização não encontrada.")
        atualizar_configuracao_flash_report_executivo(
            fiscalizacao_id,
            incluir_flash_report=(
                atual["incluir_flash_report"]
                if incluir_flash_report is None
                else bool(incluir_flash_report)
            ),
            prioridade_executiva=(
                atual["prioridade_executiva"]
                if prioridade_executiva is None
                else prioridade_executiva
            ),
            atualizacao_executiva=(
                atual["atualizacao_executiva"]
                if atualizacao_executiva is None
                else atualizacao_executiva
            ),
        )
    return resultado


_inicializar_banco_base_2_0d4 = inicializar_banco


def inicializar_banco():
    _inicializar_banco_base_2_0d4()
    garantir_estrutura_flash_report_executivo()


_obter_detalhes_fiscalizacao_base_2_0d4 = obter_detalhes_fiscalizacao


def obter_detalhes_fiscalizacao(fiscalizacao_id):
    resultado = _obter_detalhes_fiscalizacao_base_2_0d4(fiscalizacao_id)
    if not resultado:
        return None
    resultado["flash_report_executivo"] = (
        obter_configuracao_flash_report_executivo(fiscalizacao_id)
    )
    return resultado


_obter_dados_flash_report_base_2_0d4 = obter_dados_flash_report


def obter_dados_flash_report(fiscalizacao_id):
    resultado = _obter_dados_flash_report_base_2_0d4(fiscalizacao_id)
    if not resultado:
        return None
    resultado["flash_report_executivo"] = (
        obter_configuracao_flash_report_executivo(fiscalizacao_id)
    )
    return resultado


if __name__ == '__main__':
    inicializar_banco()
    resultado = diagnostico_banco()
    print()
    print('========================================')
    print(' FISCAL TRACKER')
    print(' AREAS, ENVOLVIDOS & COMUNICACOES')
    print(' FASE 2.0D1')
    print('========================================')
    print()
    print('Integridade SQLite:', resultado['integridade'])
    print('Fiscalizacoes:', resultado['fiscalizacoes'])
    print('Tarefas:', resultado['tarefas'])
    print('Flash Reports:', resultado['flash_reports'])
    print('Documentos cadastrados:', resultado['documentos'])
    print('Areas cadastradas:', resultado['areas'])
    print('Pessoas cadastradas:', resultado['pessoas'])
    print('Vinculos area-pessoas:', resultado['vinculos_area_pessoas'])
    resumo_comunicacoes = obter_resumo_comunicacoes()
    print('Comunicacoes:', resumo_comunicacoes['comunicacoes'])
    print('Destinatarios de comunicacoes:', resumo_comunicacoes['destinatarios'])
    resumo_aprovacoes = obter_resumo_aprovacoes_encerramento()
    print('Aprovadores de encerramento:', resumo_aprovacoes['total'])
    print('Aprovadores essenciais:', resumo_aprovacoes['essenciais'])
    print('Essenciais com ciencia:', resumo_aprovacoes['essenciais_cientes'])
    print('Registros com arquivo ausente:', resultado['registros_documentos_ausentes'])
    print('Arquivos orfaos no repositorio:', resultado['arquivos_orfaos'])
    print()
    if resultado['integridade'] == 'ok':
        print('BANCO INTEGRO: OK')
    print()
    print('Nenhum registro operacional foi alterado pelo diagnostico.')
    print()
