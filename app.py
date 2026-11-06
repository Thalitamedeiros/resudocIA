import base64
import io
import json
import os
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor

import anthropic
from docx import Document
from flask import Flask, jsonify, render_template, request
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from pypdf import PdfReader, PdfWriter
from werkzeug.exceptions import HTTPException

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB

MODELO = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5-5")
PAGINAS_POR_PARTE = 20
LIMITE_PARTE_BYTES = 18 * 1024 * 1024   # limite da API é 32 MB já em base64
CARACTERES_POR_PARTE = 60_000
MAX_PARTES = int(os.getenv("MAX_PARTES", "60"))
EXT_IMAGEM = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff")
TAMANHOS = {
    "curto": "um parágrafo de 2 a 3 frases",
    "medio": "um parágrafo de 4 a 6 frases",
    "longo": "dois ou três parágrafos detalhados",
}

cliente = anthropic.Anthropic()  # lê ANTHROPIC_API_KEY do ambiente
JOBS = {}  # andamento de cada resumo (fica na memória do servidor)


# ---------- leitura dos arquivos ----------
def b64(b):
    return base64.standard_b64encode(b).decode()


def bloco_imagem(dados):
    """Normaliza qualquer imagem para JPEG de até 1568 px. Ignora ícones minúsculos."""
    img = Image.open(io.BytesIO(dados)).convert("RGBA")
    if img.width < 80 or img.height < 80:
        return None
    fundo = Image.new("RGB", img.size, "white")
    fundo.paste(img, mask=img.split()[3])
    fundo.thumbnail((1568, 1568))
    buf = io.BytesIO()
    fundo.save(buf, "JPEG", quality=85)
    return {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64(buf.getvalue())}}


def partes_pdf(dados):
    """Divide o PDF em pedaços pequenos o bastante para a API (texto, imagens e páginas escaneadas)."""
    leitor = PdfReader(io.BytesIO(dados))
    if leitor.is_encrypted:
        leitor.decrypt("")
    n = len(leitor.pages)
    if n <= PAGINAS_POR_PARTE and len(dados) <= LIMITE_PARTE_BYTES:
        return [(dados, 1, n)]

    def pedaco(ini, fim):
        w = PdfWriter()
        for i in range(ini, fim):
            w.add_page(leitor.pages[i])
        buf = io.BytesIO()
        w.write(buf)
        b = buf.getvalue()
        if len(b) > LIMITE_PARTE_BYTES and fim - ini > 1:
            meio = (ini + fim) // 2
            return pedaco(ini, meio) + pedaco(meio, fim)
        return [(b, ini + 1, fim)]

    saida = []
    for ini in range(0, n, PAGINAS_POR_PARTE):
        saida += pedaco(ini, min(n, ini + PAGINAS_POR_PARTE))
    return saida


def imagens_do_zip(dados, pasta, limite=24):
    blocos = []
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        for nome in sorted(n for n in z.namelist() if n.startswith(pasta) and n.lower().endswith(EXT_IMAGEM)):
            try:
                b = bloco_imagem(z.read(nome))
            except Exception:
                continue
            if b:
                blocos.append(b)
            if len(blocos) >= limite:
                break
    return blocos


def texto_docx(dados):
    doc = Document(io.BytesIO(dados))
    linhas = [p.text for p in doc.paragraphs]
    for tabela in doc.tables:
        for linha in tabela.rows:
            linhas.append(" | ".join(c.text for c in linha.cells))
    return "\n".join(linhas)


def texto_pptx(dados):
    linhas = []
    for i, slide in enumerate(Presentation(io.BytesIO(dados)).slides, 1):
        linhas.append(f"--- Slide {i} ---")
        for forma in slide.shapes:
            if forma.has_text_frame:
                linhas.append(forma.text_frame.text)
        if slide.has_notes_slide:
            linhas.append("Notas: " + slide.notes_slide.notes_text_frame.text)
    return "\n".join(linhas)


def texto_xlsx(dados):
    linhas = []
    for aba in load_workbook(io.BytesIO(dados), read_only=True, data_only=True):
        linhas.append(f"--- Planilha {aba.title} ---")
        for i, linha in enumerate(aba.iter_rows(values_only=True)):
            if i >= 5000:
                break
            linhas.append(" | ".join("" if c is None else str(c) for c in linha))
    return "\n".join(linhas)


def decodificar(dados):
    if b"\x00" in dados[:4096]:
        raise ValueError("Não consegui ler este formato. Converta para PDF, DOCX ou imagem e tente de novo.")
    try:
        return dados.decode("utf-8")
    except UnicodeDecodeError:
        return dados.decode("latin-1")


def montar_partes(nome, dados):
    """Devolve uma lista de (blocos_para_a_IA, rótulo)."""
    ext = "." + nome.lower().rsplit(".", 1)[-1] if "." in nome else ""
    imagens = []
    if ext == ".pdf":
        return [([{"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64(b)}}],
                 f"páginas {a} a {z}") for b, a, z in partes_pdf(dados)]
    if ext in EXT_IMAGEM:
        bloco = bloco_imagem(dados)
        if not bloco:
            raise ValueError("A imagem é pequena demais para ser lida.")
        return [([bloco], "imagem")]
    if ext == ".docx":
        texto, imagens = texto_docx(dados), imagens_do_zip(dados, "word/media/")
    elif ext == ".pptx":
        texto, imagens = texto_pptx(dados), imagens_do_zip(dados, "ppt/media/")
    elif ext == ".xlsx":
        texto = texto_xlsx(dados)
    else:
        texto = decodificar(dados)

    partes = []
    texto = texto.strip()
    for i in range(0, len(texto), CARACTERES_POR_PARTE):
        trecho = texto[i:i + CARACTERES_POR_PARTE]
        partes.append(([{"type": "text", "text": f"<documento>\n{trecho}\n</documento>"}], "texto"))
    for i in range(0, len(imagens), 8):
        partes.append((imagens[i:i + 8], "imagens do documento"))
    if not partes:
        raise ValueError("Não encontrei texto nem imagens neste documento.")
    return partes


# ---------- IA ----------
def pedir(blocos, instrucao, max_tokens):
    resp = cliente.messages.create(
        model=MODELO,
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": blocos + [{"type": "text", "text": instrucao}]}],
    )
    return resp.content[0].text.strip()


def prompt_final(tamanho, unico=False):
    origem = "o documento acima" if unico else "os resumos parciais acima, que cobrem o documento inteiro"
    return (
        f"Com base em {origem} (inclusive o que aparece em imagens, gráficos e tabelas), responda SOMENTE "
        "com um JSON válido, no idioma do documento, com as chaves:\n"
        f'"resumo": {tamanho};\n'
        '"pontos_chave": lista de 3 a 6 frases curtas;\n'
        '"topicos": lista de 3 a 6 palavras ou expressões curtas.'
    )


def processar(job_id, nome, dados, tamanho):
    job = JOBS[job_id]
    try:
        partes = montar_partes(nome, dados)
        total = len(partes)
        if total > MAX_PARTES:
            raise ValueError(f"Documento grande demais ({total} partes; o limite é {MAX_PARTES}).")
        job.update(estado="resumindo", total=total, feito=0)

        if total == 1:
            bruto = pedir(partes[0][0], prompt_final(tamanho, unico=True), 1500)
        else:
            def resumir_parte(i):
                blocos, rotulo = partes[i]
                r = pedir(blocos,
                          f"Esta é a parte {i + 1} de {total} de um documento ({rotulo}). Resuma o conteúdo "
                          "com fidelidade, em tópicos curtos, incluindo o que aparece em imagens, gráficos, "
                          "tabelas, fórmulas e páginas escaneadas. Escreva no idioma do documento. "
                          "Responda só com o resumo.", 1200)
                job["feito"] += 1
                return f"Parte {i + 1} ({rotulo}):\n{r}"

            with ThreadPoolExecutor(max_workers=3) as pool:
                parciais = list(pool.map(resumir_parte, range(total)))

            job["estado"] = "finalizando"
            while len(parciais) > 1 and sum(map(len, parciais)) > 90_000:  # condensa em etapas
                grupos = [parciais[i:i + 20] for i in range(0, len(parciais), 20)]
                parciais = [pedir([{"type": "text", "text": "\n\n".join(g)}],
                                  "Condense estes resumos parciais em um único resumo em tópicos, "
                                  "sem perder as informações principais.", 2000) for g in grupos]
            bruto = pedir([{"type": "text", "text": "<resumos_parciais>\n" + "\n\n".join(parciais) + "\n</resumos_parciais>"}],
                          prompt_final(tamanho), 1500)

        ini, fim = bruto.find("{"), bruto.rfind("}")
        try:
            resultado = json.loads(bruto[ini:fim + 1])
        except (ValueError, json.JSONDecodeError):
            resultado = {"resumo": bruto, "pontos_chave": [], "topicos": []}
        job.update(estado="pronto", resultado=resultado)

    except anthropic.AuthenticationError:
        job.update(estado="erro", erro="Chave da API inválida (ANTHROPIC_API_KEY).")
    except anthropic.RateLimitError:
        job.update(estado="erro", erro="Limite de uso da API atingido. Tente de novo em instantes.")
    except anthropic.BadRequestError as e:
        app.logger.error(e)
        job.update(estado="erro", erro="A IA não conseguiu ler este arquivo (protegido, corrompido ou grande demais).")
    except anthropic.APIError:
        job.update(estado="erro", erro="A IA não respondeu. Tente novamente.")
    except ValueError as e:
        job.update(estado="erro", erro=str(e))
    except Exception as e:
        app.logger.exception(e)
        job.update(estado="erro", erro="Não foi possível ler o arquivo.")


# ---------- rotas ----------
@app.errorhandler(413)
def arquivo_grande(_):
    return jsonify(erro="O arquivo passa de 500 MB."), 413


@app.errorhandler(Exception)
def erro_geral(e):
    if isinstance(e, HTTPException):
        return jsonify(erro=e.description), e.code
    app.logger.exception(e)
    return jsonify(erro="Erro interno no servidor. Veja o terminal para detalhes."), 500


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/resumir")
def resumir():
    if not os.getenv("ANTHROPIC_API_KEY"):
        return jsonify(erro="Chave da API não configurada. Defina ANTHROPIC_API_KEY e reinicie o servidor."), 500
    arquivo = request.files.get("arquivo")
    if not arquivo or not arquivo.filename:
        return jsonify(erro="Selecione um documento."), 400

    job_id = uuid.uuid4().hex
    JOBS[job_id] = {"estado": "lendo", "feito": 0, "total": 0}
    tamanho = TAMANHOS.get(request.form.get("tamanho"), TAMANHOS["medio"])
    threading.Thread(target=processar, args=(job_id, arquivo.filename, arquivo.read(), tamanho), daemon=True).start()
    return jsonify(job=job_id)


@app.get("/status/<job_id>")
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify(erro="Processo não encontrado (o servidor foi reiniciado?)."), 404
    return jsonify(job)


if __name__ == "__main__":
    app.run(debug=True)

