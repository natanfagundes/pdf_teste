import sys
import io
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


def checar_dependencias():
    faltando = []
    libs = [
        ("fitz",        "PyMuPDF"),
        ("PIL",         "Pillow"),
        ("cv2",         "opencv-python-headless"),
        ("imagehash",   "imagehash"),
        ("skimage",     "scikit-image"),
        ("numpy",       "numpy"),
        ("pytesseract", "pytesseract"),
    ]
    for modulo, nome_pip in libs:
        try:
            __import__(modulo)
        except ImportError:
            faltando.append(nome_pip)
    return faltando


faltando = checar_dependencias()
if faltando:
    janela_erro = tk.Tk()
    janela_erro.withdraw()
    messagebox.showerror(
        "Bibliotecas nao instaladas",
        "Execute no terminal e tente novamente:\n\n"
        f"pip install {' '.join(faltando)}"
    )
    janela_erro.destroy()
    sys.exit(1)


import fitz
import numpy as np
import imagehash
import cv2
from PIL import Image, ImageTk
from skimage.metrics import structural_similarity as ssim
import pytesseract
from difflib import SequenceMatcher

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


CORES = {
    "fundo":       "#0f1117",
    "painel":      "#1a1d27",
    "painel2":     "#22263a",
    "painel3":     "#2a2f47",
    "borda":       "#2e3250",
    "azul":        "#4f8ef7",
    "roxo":        "#7c6af7",
    "verde":       "#34d399",
    "vermelho":    "#f87171",
    "amarelo":     "#fbbf24",
    "ciano":       "#22d3ee",
    "lilas":       "#a78bfa",
    "texto":       "#e2e8f0",
    "texto2":      "#8892b0",
    "texto3":      "#4a5568",
    "barra_fundo": "#1e2235",
}

FONTE_LABEL   = ("Segoe UI", 10)
FONTE_PEQUENA = ("Segoe UI", 9)
FONTE_BADGE   = ("Segoe UI", 8, "bold")
FONTE_MONO_P  = ("Consolas", 8)

FORMATOS_JORNAL = {
    "Personalizado":                    None,
    "Standard / Broadsheet (56x32 cm)": (56.0, 32.0),
    "Tabloide (28x32 cm)":              (28.0, 32.0),
    "Berliner (47x31.5 cm)":            (47.0, 31.5),
    "A4 (29.7x21 cm)":                  (29.7, 21.0),
    "A3 (42x29.7 cm)":                  (42.0, 29.7),
    "Folha de S.Paulo (55x32.5 cm)":    (55.0, 32.5),
    "O Globo Tabloide (29x27 cm)":      (29.0, 27.0),
}

# Cada jornal cadastrado tem:
#   largura_cm, altura_cm, num_colunas, margem_cm (espaco entre colunas)
# A largura de 1 coluna = (largura_cm - margem_cm * (num_colunas-1)) / num_colunas
JORNAIS_CADASTRADOS = {
    "Meu Jornal (28.5x52cm, 9col)": {
        "largura": 28.5, "altura": 52.0,
        "colunas": 9,    "margem": 0.3,
    },
    "A Tarde (56x32cm, 9col)": {
        "largura": 56.0, "altura": 32.0,
        "colunas": 9,    "margem": 0.4,
    },
    "Massa (25.5x34cm, 6col)": {
        "largura": 25.5, "altura": 34.0,
        "colunas": 6,    "margem": 0.35,
    },
    "Correio (56x32cm, 8col)": {
        "largura": 56.0, "altura": 32.0,
        "colunas": 8,    "margem": 0.4,
    },
}

# Formatos nomeados por proporcao de pagina
FORMATOS_NOMEADOS = [
    ("Pagina dupla",  1.80, float("inf")),
    ("1 Pagina",      0.90, 1.80),
    ("1/2 Pagina",    0.45, 0.90),
    ("1/3 Pagina",    0.28, 0.45),
    ("1/4 Pagina",    0.20, 0.28),
    ("1/8 Pagina",    0.10, 0.20),
    ("Rodape",        0.05, 0.10),
    ("Faixa",         0.00, 0.05),
]


def largura_coluna(jornal):
    area_util = jornal["largura"] - jornal["margem"] * (jornal["colunas"] - 1)
    return area_util / jornal["colunas"]


def calcular_colunas(w_cm, jornal):
    lc = largura_coluna(jornal)
    # Testa de 1 ate num_colunas e acha o que mais se aproxima
    melhor = 1
    menor_diff = float("inf")
    for n in range(1, jornal["colunas"] + 1):
        largura_n = lc * n + jornal["margem"] * (n - 1)
        diff = abs(w_cm - largura_n)
        if diff < menor_diff:
            menor_diff = diff
            melhor = n
    return melhor


def largura_de_n_colunas(n, jornal):
    lc = largura_coluna(jornal)
    return lc * n + jornal["margem"] * (n - 1)


def identificar_formato(w_cm, h_cm, pag_w, pag_h):
    if pag_w <= 0 or pag_h <= 0:
        return "Desconhecido"
    proporcao = (w_cm * h_cm) / (pag_w * pag_h)
    for nome, minimo, maximo in FORMATOS_NOMEADOS:
        if minimo <= proporcao < maximo:
            return nome
    return "Faixa"

PT_PARA_CM      = 0.03528
TAMANHO_PHASH   = 16
TAMANHO_SSIM    = (256, 256)
MIN_MATCHES_ORB = 12
DPI_RENDER      = 150


def pontos_para_cm(pt):
    return pt * PT_PARA_CM


def pegar_tamanho_pagina_cm(pagina):
    return pontos_para_cm(pagina.rect.width), pontos_para_cm(pagina.rect.height)


def pegar_posicao_imagem_cm(pagina, xref):
    try:
        rects = pagina.get_image_rects(xref)
        if rects:
            r = rects[0]
            return {
                "x_cm": pontos_para_cm(r.x0),
                "y_cm": pontos_para_cm(r.y0),
                "w_cm": pontos_para_cm(r.width),
                "h_cm": pontos_para_cm(r.height),
            }
    except Exception:
        pass
    return None


def calcular_proporcao(ad_w, ad_h, pag_w, pag_h):
    if pag_w <= 0 or pag_h <= 0:
        return {"area_pct": 0.0, "w_pct": 0.0, "h_pct": 0.0}
    return {
        "w_pct":    (ad_w / pag_w) * 100,
        "h_pct":    (ad_h / pag_h) * 100,
        "area_pct": (ad_w * ad_h) / (pag_w * pag_h) * 100,
    }


def escalar_para_jornal(ad_w, ad_h, pdf_pag_w, pdf_pag_h, papel_w, papel_h):
    if pdf_pag_w <= 0 or pdf_pag_h <= 0:
        return {"w_cm": ad_w, "h_cm": ad_h}
    return {
        "w_cm": ad_w * (papel_w / pdf_pag_w),
        "h_cm": ad_h * (papel_h / pdf_pag_h),
    }


def pil_para_cv2(img):
    return cv2.cvtColor(np.array(img.convert("RGB")), cv2.COLOR_RGB2BGR)


def comparar_phash(img_a, img_b):
    hash_a = imagehash.phash(img_a, hash_size=TAMANHO_PHASH)
    hash_b = imagehash.phash(img_b, hash_size=TAMANHO_PHASH)
    return 1.0 - ((hash_a - hash_b) / (TAMANHO_PHASH ** 2))


def comparar_orb(img_a, img_b):
    cinza_a = cv2.cvtColor(pil_para_cv2(img_a), cv2.COLOR_BGR2GRAY)
    cinza_b = cv2.cvtColor(pil_para_cv2(img_b), cv2.COLOR_BGR2GRAY)
    orb = cv2.ORB_create(nfeatures=1000)
    kp_a, desc_a = orb.detectAndCompute(cinza_a, None)
    kp_b, desc_b = orb.detectAndCompute(cinza_b, None)
    if desc_a is None or desc_b is None or len(kp_a) < 5 or len(kp_b) < 5:
        return 0.0
    matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    bons = [m for m in matcher.match(desc_a, desc_b) if m.distance < 64]
    if len(bons) < MIN_MATCHES_ORB:
        return 0.0
    return min(1.0, len(bons) / min(len(kp_a), len(kp_b)))


def comparar_ssim(img_a, img_b):
    arr_a = np.array(img_a.resize(TAMANHO_SSIM).convert("L"))
    arr_b = np.array(img_b.resize(TAMANHO_SSIM).convert("L"))
    score, _ = ssim(arr_a, arr_b, full=True)
    return max(0.0, float(score))


def comparar_tudo(img_a, img_b):
    return (
        comparar_phash(img_a, img_b) * 0.40 +
        comparar_orb(img_a, img_b)   * 0.35 +
        comparar_ssim(img_a, img_b)  * 0.25
    )


METODOS = {
    "phash": comparar_phash,
    "orb":   comparar_orb,
    "ssim":  comparar_ssim,
    "all":   comparar_tudo,
}


def renderizar_pagina(pagina):
    matriz = fitz.Matrix(DPI_RENDER / 72, DPI_RENDER / 72)
    pixmap = pagina.get_pixmap(matrix=matriz, alpha=False)
    return Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)


def gerar_preview_anuncio(pdf_path, num_pagina, x_cm, y_cm, w_cm, h_cm,
                           largura_thumb=340):
    try:
        import PIL.ImageDraw as ImageDraw

        doc = fitz.open(pdf_path)
        pagina = doc[num_pagina - 1]
        escala = DPI_RENDER / 72
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(escala, escala), alpha=False)
        img_pagina = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
        doc.close()

        px_por_cm = DPI_RENDER / 2.54
        px = int(x_cm * px_por_cm)
        py = int(y_cm * px_por_cm)
        pw = int(w_cm * px_por_cm)
        ph = int(h_cm * px_por_cm)

        px = max(0, px); py = max(0, py)
        pw = min(pw, img_pagina.width  - px)
        ph = min(ph, img_pagina.height - py)

        if pw < 5 or ph < 5:
            return None

        # Desenha retangulo vermelho GROSSO na pagina inteira
        draw = ImageDraw.Draw(img_pagina)
        espessura = max(6, int(px_por_cm * 0.18))
        for i in range(espessura):
            draw.rectangle(
                [px - i, py - i, px + pw + i, py + ph + i],
                outline=(220, 30, 30)
            )

        # Redimensiona a pagina inteira para caber no card
        ratio = largura_thumb / img_pagina.width
        nova_h = int(img_pagina.height * ratio)
        return img_pagina.resize((largura_thumb, nova_h), Image.LANCZOS)

    except Exception:
        return None


def gerar_pdf_marcado(pdf_path, marcacoes, destino_path):
    doc = fitz.open(pdf_path)
    for m in marcacoes:
        num_pagina = m["pagina"]
        x_cm = m.get("x_cm"); y_cm = m.get("y_cm")
        w_cm = m.get("w_cm"); h_cm = m.get("h_cm")
        if x_cm is None or w_cm is None:
            continue
        pagina = doc[num_pagina - 1]
        x0 = x_cm / PT_PARA_CM
        y0 = y_cm / PT_PARA_CM
        x1 = (x_cm + w_cm) / PT_PARA_CM
        y1 = (y_cm + h_cm) / PT_PARA_CM
        rect = fitz.Rect(x0, y0, x1, y1)
        # Retangulo vermelho sem preenchimento, borda grossa
        pagina.draw_rect(rect, color=(0.86, 0.12, 0.12), width=3.0)
    doc.save(destino_path)
    doc.close()


def montar_candidatos(doc):
    candidatos = []
    for num_pagina in range(len(doc)):
        pagina = doc[num_pagina]
        pag_w, pag_h = pegar_tamanho_pagina_cm(pagina)

        for info_img in pagina.get_images(full=True):
            xref = info_img[0]
            try:
                dados_brutos = doc.extract_image(xref)
                img = Image.open(io.BytesIO(dados_brutos["image"])).convert("RGB")
                pos = pegar_posicao_imagem_cm(pagina, xref)
                candidatos.append({
                    "pagina": num_pagina + 1,
                    "fonte":  "embutida",
                    "imagem": img,
                    "pag_w":  pag_w,
                    "pag_h":  pag_h,
                    "w_cm":   pos["w_cm"] if pos else pontos_para_cm(img.width),
                    "h_cm":   pos["h_cm"] if pos else pontos_para_cm(img.height),
                    "x_cm":   pos["x_cm"] if pos else None,
                    "y_cm":   pos["y_cm"] if pos else None,
                })
            except Exception:
                pass

        try:
            img_pagina = renderizar_pagina(pagina)
            candidatos.append({
                "pagina": num_pagina + 1,
                "fonte":  "renderizada",
                "imagem": img_pagina,
                "pag_w":  pag_w,
                "pag_h":  pag_h,
                "w_cm":   pag_w,
                "h_cm":   pag_h,
                "x_cm":   0.0,
                "y_cm":   0.0,
            })
        except Exception:
            pass

    return candidatos


def buscar_imagem_no_pdf(pdf_path, imagem_busca, funcao_comparacao,
                         limiar, cb_progresso, cb_resultado):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        cb_resultado(erro=str(e))
        return

    candidatos = montar_candidatos(doc)
    doc.close()

    if not candidatos:
        cb_resultado(erro="Nao foi possivel extrair nenhuma imagem do PDF.")
        return

    total = len(candidatos)
    matches = []
    melhor_score = 0.0

    for i, cand in enumerate(candidatos):
        try:
            score = funcao_comparacao(imagem_busca, cand["imagem"])
        except Exception:
            score = 0.0

        if score > melhor_score:
            melhor_score = score

        if score >= limiar:
            thumb = cand["imagem"].copy()
            thumb.thumbnail((120, 120), Image.LANCZOS)

            # Tenta gerar preview recortado da pagina (com borda vermelha)
            # Se nao tiver coordenadas, usa a propria imagem embutida redimensionada
            preview = None
            if cand.get("x_cm") is not None and cand.get("w_cm"):
                preview = gerar_preview_anuncio(
                    pdf_path, cand["pagina"],
                    cand["x_cm"], cand["y_cm"],
                    cand["w_cm"], cand["h_cm"])

            if preview is None:
                # Fallback: usa a propria imagem embutida como preview
                preview = cand["imagem"].copy()
                preview.thumbnail((200, 400), Image.LANCZOS)

            matches.append({
                **cand,
                "score":   score,
                "thumb":   thumb,
                "preview": preview,
                "w_px":    cand["imagem"].width,
                "h_px":    cand["imagem"].height,
            })

        cb_progresso(i + 1, total, cand["pagina"])

    cb_resultado(matches=matches, melhor_score=melhor_score, total=total, erro=None)


def listar_todos_anuncios(pdf_path, cb_progresso, cb_resultado):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        cb_resultado(erro=str(e))
        return

    candidatos = montar_candidatos(doc)
    doc.close()

    if not candidatos:
        cb_resultado(erro="Nao foi possivel extrair nenhuma imagem do PDF.")
        return

    total = len(candidatos)
    resultados = []

    for i, cand in enumerate(candidatos):
        thumb = cand["imagem"].copy()
        thumb.thumbnail((120, 120), Image.LANCZOS)

        preview = None
        if cand.get("x_cm") is not None and cand.get("w_cm"):
            preview = gerar_preview_anuncio(
                pdf_path, cand["pagina"],
                cand["x_cm"], cand["y_cm"],
                cand["w_cm"], cand["h_cm"])

        if preview is None:
            preview = cand["imagem"].copy()
            preview.thumbnail((200, 400), Image.LANCZOS)

        resultados.append({
            **cand,
            "numero":  i + 1,
            "thumb":   thumb,
            "preview": preview,
            "w_px":    cand["imagem"].width,
            "h_px":    cand["imagem"].height,
        })
        cb_progresso(i + 1, total, cand["pagina"])

    cb_resultado(imagens=resultados, total=total, erro=None)



import unicodedata
import re as _re


def normalizar(texto):
    # Remove acentos, coloca tudo minusculo, separa palavras coladas por CamelCase
    # e remove pontuacao — resolve o problema do OCR juntar palavras
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = texto.lower()
    # Separa palavras grudadas por letra maiuscula (ex: "MunicípiodeSanta" -> "municipio de santa")
    texto = _re.sub(r"([a-z])([A-Z])", r"\1 \2", texto)
    # Separa letras de numeros colados (ex: "N002" -> "N 002")
    texto = _re.sub(r"([a-zA-Z])([0-9])", r"\1 \2", texto)
    texto = _re.sub(r"([0-9])([a-zA-Z])", r"\1 \2", texto)
    # Remove pontuacao e caracteres especiais
    texto = _re.sub(r"[^a-z0-9 ]", " ", texto)
    # Colapsa espacos multiplos
    texto = " ".join(texto.split())
    return texto


def palavras_relevantes(texto, tamanho_minimo=3):
    # Retorna apenas palavras com tamanho minimo — ignora artigos, preposicoes
    return set(p for p in normalizar(texto).split() if len(p) >= tamanho_minimo)


def score_por_palavras(texto_busca, texto_pdf):
    # Calcula quantas palavras do anuncio aparecem no texto do PDF
    # Mais robusto que SequenceMatcher quando o OCR embaralha a ordem
    palavras_busca = palavras_relevantes(texto_busca)
    palavras_pdf   = palavras_relevantes(texto_pdf)

    if not palavras_busca:
        return 0.0

    encontradas = palavras_busca & palavras_pdf
    return len(encontradas) / len(palavras_busca)


def score_sequencia(texto_busca, texto_pdf):
    # Compara a sequencia de caracteres apos normalizar
    a = normalizar(texto_busca)
    b = normalizar(texto_pdf)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def calcular_score_texto(texto_busca, texto_pdf):
    # Combina os dois metodos:
    # - palavras (peso maior): robusto ao OCR embaralhar ou juntar palavras
    # - sequencia (peso menor): garante que a ordem importa um pouco
    sp = score_por_palavras(texto_busca, texto_pdf)
    ss = score_sequencia(texto_busca, texto_pdf)
    return sp * 0.70 + ss * 0.30


def trecho_representativo(texto_pdf, max_chars=400):
    # Retorna os primeiros max_chars do texto do bloco, limpo
    limpo = " ".join(texto_pdf.split())
    return limpo[:max_chars]


def ocr_da_imagem(caminho_img):
    img = Image.open(caminho_img).convert("RGB")
    # Usa PSM 6 (bloco uniforme de texto) — melhor para anuncios
    config = "--psm 6 -l por+eng"
    texto = pytesseract.image_to_string(img, config=config)
    return texto.strip()


def extrair_blocos_por_pagina(pdf_path):
    doc = fitz.open(pdf_path)
    paginas = {}
    for num_pagina in range(len(doc)):
        pagina = doc[num_pagina]
        pag_w, pag_h = pegar_tamanho_pagina_cm(pagina)
        blocos = []
        for bloco in pagina.get_text("blocks"):
            x0, y0, x1, y1, texto, *_ = bloco
            texto = texto.strip()
            if not texto or len(texto) < 5:
                continue
            blocos.append({
                "texto": texto,
                "x0": x0, "y0": y0, "x1": x1, "y1": y1,
            })
        paginas[num_pagina + 1] = {
            "blocos":  blocos,
            "pag_w":   pag_w,
            "pag_h":   pag_h,
        }
    doc.close()
    return paginas


def bbox_expandido(bloco_match, todos_blocos, tolerancia_pt=10):
    x0f = bloco_match["x0"]; y0f = bloco_match["y0"]
    x1f = bloco_match["x1"]; y1f = bloco_match["y1"]

    # Expande iterativamente: enquanto houver blocos tocando o bbox atual, inclui
    # Isso captura o anuncio inteiro (titulo + corpo + assinatura) sem pegar
    # blocos de outros anuncios que ficam separados por espaco em branco
    mudou = True
    while mudou:
        mudou = False
        for b in todos_blocos:
            nao_toca = (b["x1"] < x0f - tolerancia_pt or
                        b["x0"] > x1f + tolerancia_pt or
                        b["y1"] < y0f - tolerancia_pt or
                        b["y0"] > y1f + tolerancia_pt)
            if not nao_toca:
                nx0 = min(x0f, b["x0"]); ny0 = min(y0f, b["y0"])
                nx1 = max(x1f, b["x1"]); ny1 = max(y1f, b["y1"])
                if (nx0, ny0, nx1, ny1) != (x0f, y0f, x1f, y1f):
                    x0f, y0f, x1f, y1f = nx0, ny0, nx1, ny1
                    mudou = True

    return x0f, y0f, x1f, y1f


def buscar_texto_no_pdf(pdf_path, texto_busca, limiar, cb_progresso, cb_resultado):
    try:
        paginas = extrair_blocos_por_pagina(pdf_path)
    except Exception as e:
        cb_resultado(erro=str(e))
        return

    if not paginas:
        cb_resultado(erro="Nenhum texto encontrado no PDF.")
        return

    # Monta lista de candidatos: blocos individuais + grupos de blocos vizinhos
    candidatos = []
    for num_pagina, info in paginas.items():
        blocos = info["blocos"]
        pag_w  = info["pag_w"]
        pag_h  = info["pag_h"]

        for b in blocos:
            candidatos.append({
                "pagina":      num_pagina,
                "texto":       b["texto"],
                "bloco_orig":  b,
                "pag_w":       pag_w,
                "pag_h":       pag_h,
                "todos_blocos": blocos,
            })

        for tamanho in range(2, min(12, len(blocos) + 1)):
            for inicio in range(len(blocos) - tamanho + 1):
                grupo = blocos[inicio:inicio + tamanho]
                texto_grupo = " ".join(b["texto"] for b in grupo)
                bloco_virtual = {
                    "x0": min(b["x0"] for b in grupo),
                    "y0": min(b["y0"] for b in grupo),
                    "x1": max(b["x1"] for b in grupo),
                    "y1": max(b["y1"] for b in grupo),
                }
                candidatos.append({
                    "pagina":       num_pagina,
                    "texto":        texto_grupo,
                    "bloco_orig":   bloco_virtual,
                    "pag_w":        pag_w,
                    "pag_h":        pag_h,
                    "todos_blocos": blocos,
                })

    total = len(candidatos)
    scored = []
    melhor_score = 0.0

    for i, cand in enumerate(candidatos):
        score = calcular_score_texto(texto_busca, cand["texto"])
        if score > melhor_score:
            melhor_score = score
        if score >= limiar:
            scored.append((score, cand))
        cb_progresso(i + 1, total, cand["pagina"])

    # Para cada match, expande o bounding box para incluir blocos vizinhos
    matches_unicos = []
    regioes_vistas = []

    for score, cand in sorted(scored, key=lambda x: x[0], reverse=True):
        bloco = cand["bloco_orig"]
        x0, y0, x1, y1 = bbox_expandido(bloco, cand["todos_blocos"], tolerancia_pt=10)

        x_cm = pontos_para_cm(x0)
        y_cm = pontos_para_cm(y0)
        w_cm = pontos_para_cm(x1 - x0)
        h_cm = pontos_para_cm(y1 - y0)

        sobrepoe = False
        for r in regioes_vistas:
            if (cand["pagina"] == r["pagina"]
                    and abs(x_cm - r["x_cm"]) < 1.5
                    and abs(y_cm - r["y_cm"]) < 1.5):
                sobrepoe = True
                break

        if not sobrepoe:
            regioes_vistas.append({"pagina": cand["pagina"], "x_cm": x_cm, "y_cm": y_cm})
            preview = gerar_preview_anuncio(
                pdf_path, cand["pagina"], x_cm, y_cm, w_cm, h_cm)
            matches_unicos.append({
                "pagina":  cand["pagina"],
                "score":   score,
                "trecho":  trecho_representativo(cand["texto"]),
                "x_cm":    x_cm,
                "y_cm":    y_cm,
                "w_cm":    w_cm,
                "h_cm":    h_cm,
                "pag_w":   cand["pag_w"],
                "pag_h":   cand["pag_h"],
                "preview": preview,
            })

    cb_resultado(matches=matches_unicos, melhor_score=melhor_score, total=total, erro=None)


class BotaoArquivo(tk.Label):
    def __init__(self, parent, texto, icone, ao_clicar):
        super().__init__(
            parent,
            text=f"{icone}\n{texto}",
            font=("Segoe UI", 10),
            fg=CORES["texto2"],
            bg=CORES["painel"],
            cursor="hand2",
            justify="center",
            padx=16,
            pady=14,
            relief="flat",
        )
        self._ao_clicar = ao_clicar
        self._selecionado = False
        self.bind("<Button-1>", lambda e: self._ao_clicar())
        self.bind("<Enter>",    lambda e: self.config(fg=CORES["azul"]))
        self.bind("<Leave>",    lambda e: self.config(
            fg=CORES["verde"] if self._selecionado else CORES["texto2"]))

    def mostrar_arquivo(self, nome):
        self._selecionado = True
        nome_curto = nome if len(nome) <= 34 else "..." + nome[-31:]
        self.config(text=f"OK\n{nome_curto}", fg=CORES["verde"])


class BarraProgresso(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, height=8, bg=CORES["fundo"],
                         highlightthickness=0, **kw)
        self._porcentagem = 0.0
        self.bind("<Configure>", self._desenhar)

    def atualizar(self, valor):
        self._porcentagem = max(0.0, min(1.0, valor))
        self._desenhar()

    def _desenhar(self, event=None):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 2:
            return
        self.create_rectangle(0, 0, w, h, fill=CORES["barra_fundo"], outline="")
        preenchido = int(w * self._porcentagem)
        if preenchido > 0:
            passos = max(1, preenchido // 4)
            for i in range(passos):
                x0 = int(preenchido * i / passos)
                x1 = int(preenchido * (i + 1) / passos) + 1
                t  = i / max(passos - 1, 1)
                r  = int(0x4f + (0x7c - 0x4f) * t)
                g  = int(0x8e + (0x6a - 0x8e) * t)
                self.create_rectangle(x0, 0, x1, h,
                                      fill=f"#{r:02x}{g:02x}f7", outline="")


def criar_barra_score(parent, score):
    cor = (CORES["verde"] if score >= 0.85
           else CORES["amarelo"] if score >= 0.65
           else CORES["vermelho"])
    linha = tk.Frame(parent, bg=CORES["painel2"])
    tk.Label(linha, text=f"{score:.1%}", font=FONTE_BADGE,
             fg=cor, bg=CORES["painel2"], width=6).pack(side="left")
    fundo = tk.Frame(linha, bg=CORES["barra_fundo"], height=6, width=110)
    fundo.pack(side="left", padx=(4, 0), pady=6)
    fundo.pack_propagate(False)
    tk.Frame(fundo, bg=cor, height=6, width=max(2, int(110 * score))).place(x=0, y=0)
    return linha


def criar_bloco_metricas(parent, w_px, h_px, w_cm_pdf, h_cm_pdf,
                          x_cm, y_cm, pag_w_cm, pag_h_cm,
                          var_formato, var_dpi):
    container = tk.Frame(parent, bg=CORES["painel2"])
    _vivo = [True]

    def cm_via_dpi():
        try:
            dpi = float(var_dpi.get())
            if dpi <= 0:
                raise ValueError
        except (ValueError, tk.TclError):
            dpi = 96.0
        return (w_px / dpi) * 2.54, (h_px / dpi) * 2.54

    def atualizar(*_):
        if not _vivo[0]:
            return
        try:
            if not container.winfo_exists():
                return
        except Exception:
            return
        try:
            papel = FORMATOS_JORNAL.get(var_formato.get())
            usar_w = w_cm_pdf if w_cm_pdf else cm_via_dpi()[0]
            usar_h = h_cm_pdf if h_cm_pdf else cm_via_dpi()[1]

            lbl_pdf.config(text=f"L {usar_w:.2f} cm  x  A {usar_h:.2f} cm")

            if papel and pag_w_cm and pag_h_cm:
                escala = escalar_para_jornal(usar_w, usar_h,
                                             pag_w_cm, pag_h_cm,
                                             papel[0], papel[1])
                lbl_print.config(
                    text=f"L {escala['w_cm']:.2f} cm  x  A {escala['h_cm']:.2f} cm",
                    fg=CORES["verde"])
            else:
                lbl_print.config(text="-- selecione um formato --", fg=CORES["texto3"])

            if pag_w_cm and pag_h_cm:
                prop = calcular_proporcao(usar_w, usar_h, pag_w_cm, pag_h_cm)
                lbl_prop.config(
                    text=f"Larg: {prop['w_pct']:.1f}%  "
                         f"Alt: {prop['h_pct']:.1f}%  "
                         f"Area: {prop['area_pct']:.1f}%")
            else:
                lbl_prop.config(text="--")
        except tk.TclError:
            pass

    def ao_destruir(event):
        if event.widget is not container:
            return
        _vivo[0] = False
        try:
            var_dpi.trace_remove("write", id_trace_dpi)
        except Exception:
            pass
        try:
            var_formato.trace_remove("write", id_trace_formato)
        except Exception:
            pass

    def nova_linha(tag, cor_tag):
        linha = tk.Frame(container, bg=CORES["painel2"])
        linha.pack(fill="x", pady=2)
        tk.Label(linha, text=tag, font=FONTE_BADGE,
                 fg=cor_tag, bg=CORES["painel3"], padx=2).pack(side="left")
        return linha

    linha_px = nova_linha("  px  ", CORES["azul"])
    tk.Label(linha_px, text=f"L {w_px} px  x  A {h_px} px",
             font=FONTE_MONO_P, fg=CORES["texto2"],
             bg=CORES["painel2"], padx=6).pack(side="left")

    linha_pdf = nova_linha(" PDF  ", CORES["ciano"])
    lbl_pdf = tk.Label(linha_pdf, text="...", font=FONTE_MONO_P,
                       fg=CORES["texto"], bg=CORES["painel2"], padx=6)
    lbl_pdf.pack(side="left")

    linha_print = nova_linha("PRINT ", CORES["lilas"])
    lbl_print = tk.Label(linha_print, text="--", font=FONTE_MONO_P,
                         fg=CORES["texto3"], bg=CORES["painel2"], padx=6)
    lbl_print.pack(side="left")

    if x_cm is not None and y_cm is not None:
        linha_pos = nova_linha(" POS  ", CORES["amarelo"])
        tk.Label(linha_pos, text=f"X {x_cm:.2f} cm  Y {y_cm:.2f} cm",
                 font=FONTE_MONO_P, fg=CORES["texto2"],
                 bg=CORES["painel2"], padx=6).pack(side="left")

    if pag_w_cm and pag_h_cm:
        linha_pag = nova_linha(" PAG  ", CORES["texto3"])
        tk.Label(linha_pag, text=f"L {pag_w_cm:.2f} cm  x  A {pag_h_cm:.2f} cm",
                 font=FONTE_MONO_P, fg=CORES["texto3"],
                 bg=CORES["painel2"], padx=6).pack(side="left")

    linha_prop = nova_linha(" PROP ", CORES["verde"])
    lbl_prop = tk.Label(linha_prop, text="...", font=FONTE_MONO_P,
                        fg=CORES["texto2"], bg=CORES["painel2"], padx=6)
    lbl_prop.pack(side="left")

    id_trace_dpi     = var_dpi.trace_add("write", atualizar)
    id_trace_formato = var_formato.trace_add("write", atualizar)
    container.bind("<Destroy>", ao_destruir)
    atualizar()

    return container


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("PDF Image Finder - Jornais")
        self.geometry("1020x760")
        self.minsize(820, 600)
        self.configure(bg=CORES["fundo"])

        self.caminho_pdf    = tk.StringVar()
        self.caminho_img    = tk.StringVar()
        self.metodo_busca   = tk.StringVar(value="phash")
        self.limiar         = tk.DoubleVar(value=0.85)
        self.dpi            = tk.DoubleVar(value=96.0)
        self.formato_jornal = tk.StringVar(value="Personalizado")
        self.limiar_texto    = tk.DoubleVar(value=0.50)
        self.caminho_img_ocr = ""
        self.jornal_ativo    = tk.StringVar(value="Meu Jornal (28.5x52cm, 9col)")
        self._buscando      = False
        self._thumbnails    = []

        self._painel_esq = None
        self._montar_interface()
        self._centralizar_janela()

    def _montar_interface(self):
        cabecalho = tk.Frame(self, bg=CORES["fundo"])
        cabecalho.pack(fill="x", padx=24, pady=(18, 0))
        tk.Label(cabecalho, text="PDF Image Finder",
                 font=("Segoe UI", 18, "bold"),
                 fg=CORES["texto"], bg=CORES["fundo"]).pack(side="left")
        tk.Label(cabecalho, text="Medicao de anuncios para jornais impressos",
                 font=FONTE_PEQUENA, fg=CORES["texto2"],
                 bg=CORES["fundo"]).pack(side="left", padx=(12, 0), pady=(4, 0))

        tk.Frame(self, bg=CORES["borda"], height=1).pack(fill="x", padx=24, pady=(10, 0))

        corpo = tk.Frame(self, bg=CORES["fundo"])
        corpo.pack(fill="both", expand=True, padx=24, pady=10)

        painel_esq = tk.Frame(corpo, bg=CORES["fundo"], width=390)
        painel_esq.pack(side="left", fill="y")
        painel_esq.pack_propagate(False)
        self._painel_esq = painel_esq
        self._montar_painel_esquerdo(painel_esq)

        tk.Frame(corpo, bg=CORES["borda"], width=1).pack(side="left", fill="y", padx=14)

        painel_dir = tk.Frame(corpo, bg=CORES["fundo"])
        painel_dir.pack(side="left", fill="both", expand=True)
        self._montar_painel_direito(painel_dir)

    def _montar_painel_esquerdo(self, parent):
        self._titulo_secao(parent, "ARQUIVOS")

        frame_pdf = tk.Frame(parent, bg=CORES["painel"],
                             highlightbackground=CORES["borda"], highlightthickness=1)
        frame_pdf.pack(fill="x", pady=(0, 6))
        self._btn_pdf = BotaoArquivo(frame_pdf, texto="Clique para selecionar o PDF",
                                     icone="PDF", ao_clicar=self._selecionar_pdf)
        self._btn_pdf.pack(fill="x")

        frame_img = tk.Frame(parent, bg=CORES["painel"],
                             highlightbackground=CORES["borda"], highlightthickness=1)
        frame_img.pack(fill="x", pady=(0, 6))
        self._btn_img = BotaoArquivo(frame_img,
                                     texto="Clique para selecionar a imagem de busca",
                                     icone="IMG", ao_clicar=self._selecionar_imagem)
        self._btn_img.pack(fill="x")

        self._lbl_preview = tk.Label(parent, bg=CORES["fundo"])
        self._lbl_preview.pack(pady=(0, 2))

        tk.Frame(parent, bg=CORES["borda"], height=1).pack(fill="x", pady=(4, 8))
        self._titulo_secao(parent, "FORMATO DO JORNAL")

        linha_formato = tk.Frame(parent, bg=CORES["fundo"])
        linha_formato.pack(fill="x", pady=(0, 4))
        tk.Label(linha_formato, text="Formato:", font=FONTE_LABEL,
                 fg=CORES["texto2"], bg=CORES["fundo"],
                 width=10, anchor="w").pack(side="left")

        estilo = ttk.Style()
        estilo.theme_use("clam")
        estilo.configure("Escuro.TCombobox",
            fieldbackground=CORES["painel2"], background=CORES["painel2"],
            foreground=CORES["texto"], arrowcolor=CORES["azul"],
            selectbackground=CORES["painel3"], selectforeground=CORES["texto"])

        ttk.Combobox(linha_formato, textvariable=self.formato_jornal,
                     values=list(FORMATOS_JORNAL.keys()),
                     state="readonly", style="Escuro.TCombobox",
                     width=28).pack(side="left", padx=(4, 0))

        self._lbl_formato_info = tk.Label(parent, text="", font=FONTE_PEQUENA,
                                          fg=CORES["lilas"], bg=CORES["fundo"], anchor="w")
        self._lbl_formato_info.pack(fill="x", pady=(2, 0))
        self.formato_jornal.trace_add("write", self._ao_mudar_formato)
        self._ao_mudar_formato()

        tk.Frame(parent, bg=CORES["borda"], height=1).pack(fill="x", pady=(8, 8))
        self._titulo_secao(parent, "JORNAL / COLUNAGEM")

        linha_jornal = tk.Frame(parent, bg=CORES["fundo"])
        linha_jornal.pack(fill="x", pady=(0, 4))
        tk.Label(linha_jornal, text="Jornal:", font=FONTE_LABEL,
                 fg=CORES["texto2"], bg=CORES["fundo"],
                 width=10, anchor="w").pack(side="left")

        opcoes_jornal = list(JORNAIS_CADASTRADOS.keys()) + ["Personalizado"]
        cb_jornal = ttk.Combobox(linha_jornal, textvariable=self.jornal_ativo,
                     values=opcoes_jornal, state="readonly",
                     style="Escuro.TCombobox", width=26)
        cb_jornal.pack(side="left", padx=(4, 0))

        self._lbl_jornal_info = tk.Label(parent, text="", font=FONTE_PEQUENA,
                                          fg=CORES["ciano"], bg=CORES["fundo"],
                                          anchor="w")
        self._lbl_jornal_info.pack(fill="x", pady=(2, 0))
        self.jornal_ativo.trace_add("write", self._ao_mudar_jornal)
        self._ao_mudar_jornal()

        tk.Button(parent, text="+ Cadastrar novo jornal",
            font=("Segoe UI", 8), fg=CORES["texto3"], bg=CORES["fundo"],
            relief="flat", cursor="hand2", anchor="w",
            activeforeground=CORES["azul"],
            command=self._abrir_cadastro_jornal
        ).pack(anchor="w", pady=(2, 0))

        tk.Frame(parent, bg=CORES["borda"], height=1).pack(fill="x", pady=(8, 8))
        self._titulo_secao(parent, "DPI")

        linha_dpi = tk.Frame(parent, bg=CORES["fundo"])
        linha_dpi.pack(fill="x", pady=(0, 2))
        tk.Label(linha_dpi, text="DPI:", font=FONTE_LABEL,
                 fg=CORES["texto2"], bg=CORES["fundo"],
                 width=10, anchor="w").pack(side="left")
        tk.Entry(linha_dpi, textvariable=self.dpi,
                 font=("Segoe UI", 10, "bold"), width=6,
                 fg=CORES["ciano"], bg=CORES["painel2"],
                 insertbackground=CORES["ciano"],
                 relief="flat", justify="center").pack(side="left", padx=(4, 4))
        tk.Label(linha_dpi, text="DPI", font=FONTE_PEQUENA,
                 fg=CORES["texto3"], bg=CORES["fundo"]).pack(side="left")

        linha_presets = tk.Frame(parent, bg=CORES["fundo"])
        linha_presets.pack(fill="x", pady=(2, 0))
        tk.Label(linha_presets, text="Preset:", font=FONTE_PEQUENA,
                 fg=CORES["texto3"], bg=CORES["fundo"]).pack(side="left")
        for valor, rotulo in [(72, "72 - PDF"), (96, "96 - Tela"),
                               (150, "150"), (300, "300 - Imp.")]:
            tk.Button(linha_presets, text=rotulo,
                      font=("Segoe UI", 7, "bold"),
                      fg=CORES["texto2"], bg=CORES["painel2"],
                      relief="flat", cursor="hand2", padx=5, pady=2,
                      activebackground=CORES["painel3"],
                      command=lambda v=valor: self.dpi.set(float(v))
                      ).pack(side="left", padx=2)

        tk.Frame(parent, bg=CORES["borda"], height=1).pack(fill="x", pady=(8, 8))
        self._titulo_secao(parent, "OPCOES DE BUSCA")

        linha_metodo = tk.Frame(parent, bg=CORES["fundo"])
        linha_metodo.pack(fill="x", pady=(0, 8))
        tk.Label(linha_metodo, text="Metodo:", font=FONTE_LABEL,
                 fg=CORES["texto2"], bg=CORES["fundo"],
                 width=10, anchor="w").pack(side="left")

        frame_radios = tk.Frame(linha_metodo, bg=CORES["fundo"])
        frame_radios.pack(side="left")
        for valor, rotulo in [("phash", "pHash"), ("orb", "ORB"),
                               ("ssim", "SSIM"), ("all", "Todos")]:
            rb = tk.Radiobutton(frame_radios, text=rotulo,
                variable=self.metodo_busca, value=valor,
                font=FONTE_PEQUENA, fg=CORES["texto2"], bg=CORES["fundo"],
                selectcolor=CORES["painel2"], activebackground=CORES["fundo"],
                activeforeground=CORES["azul"], indicatoron=0,
                relief="flat", padx=8, pady=4, cursor="hand2")
            rb.pack(side="left", padx=2)
            self._estilizar_radio(rb)

        linha_limiar = tk.Frame(parent, bg=CORES["fundo"])
        linha_limiar.pack(fill="x", pady=(0, 4))
        tk.Label(linha_limiar, text="Sensib.:", font=FONTE_LABEL,
                 fg=CORES["texto2"], bg=CORES["fundo"],
                 width=10, anchor="w").pack(side="left")
        self._lbl_limiar = tk.Label(linha_limiar,
                                    text=f"{self.limiar.get():.0%}",
                                    font=("Segoe UI", 10, "bold"),
                                    fg=CORES["azul"], bg=CORES["fundo"], width=5)
        self._lbl_limiar.pack(side="right")
        estilo.configure("TScale", background=CORES["fundo"],
                         troughcolor=CORES["barra_fundo"],
                         sliderlength=14, sliderrelief="flat")
        ttk.Scale(linha_limiar, from_=0.40, to=0.99,
                  variable=self.limiar, orient="horizontal",
                  command=lambda v: self._lbl_limiar.config(text=f"{float(v):.0%}")
                  ).pack(side="left", fill="x", expand=True, padx=(0, 4))

        tk.Frame(parent, bg=CORES["borda"], height=1).pack(fill="x", pady=(8, 8))

        self._btn_buscar = tk.Button(parent, text="Iniciar Busca",
            font=("Segoe UI", 11, "bold"), fg="#ffffff", bg=CORES["azul"],
            activebackground=CORES["roxo"], activeforeground="#ffffff",
            relief="flat", cursor="hand2", pady=9,
            command=self._iniciar_busca)
        self._btn_buscar.pack(fill="x", pady=(0, 5))

        self._btn_listar = tk.Button(parent, text="Listar Todos os Anuncios",
            font=("Segoe UI", 9, "bold"), fg=CORES["texto2"], bg=CORES["painel2"],
            activebackground=CORES["painel3"], activeforeground=CORES["ciano"],
            relief="flat", cursor="hand2", pady=7,
            command=self._iniciar_listagem)
        self._btn_listar.pack(fill="x")

    def _montar_painel_direito(self, parent):
        self._lbl_status = tk.Label(parent, text="Aguardando...",
            font=FONTE_LABEL, fg=CORES["texto3"], bg=CORES["fundo"], anchor="w")
        self._lbl_status.pack(fill="x", pady=(0, 4))

        self._barra = BarraProgresso(parent)
        self._barra.pack(fill="x", pady=(0, 2))

        self._lbl_progresso = tk.Label(parent, text="", font=FONTE_PEQUENA,
                                        fg=CORES["texto3"], bg=CORES["fundo"], anchor="w")
        self._lbl_progresso.pack(fill="x", pady=(0, 6))

        estilo_abas = ttk.Style()
        estilo_abas.configure("Escuro.TNotebook",
            background=CORES["fundo"], borderwidth=0)
        estilo_abas.configure("Escuro.TNotebook.Tab",
            background=CORES["painel2"], foreground=CORES["texto2"],
            padding=(12, 6), font=("Segoe UI", 9, "bold"))
        estilo_abas.map("Escuro.TNotebook.Tab",
            background=[("selected", CORES["painel3"])],
            foreground=[("selected", CORES["azul"])])

        self._abas = ttk.Notebook(parent, style="Escuro.TNotebook")
        self._abas.pack(fill="both", expand=True)

        self._aba_busca = tk.Frame(self._abas, bg=CORES["painel"])
        self._abas.add(self._aba_busca, text="Resultado da Busca")
        self._canvas_busca, self._frame_busca = self._criar_area_rolavel(self._aba_busca)

        self._aba_lista = tk.Frame(self._abas, bg=CORES["painel"])
        self._abas.add(self._aba_lista, text="Todos os Anuncios")
        self._canvas_lista, self._frame_lista = self._criar_area_rolavel(self._aba_lista)

        self._aba_texto = tk.Frame(self._abas, bg=CORES["painel"])
        self._abas.add(self._aba_texto, text="Busca por Texto (OCR)")
        self._montar_aba_texto(self._aba_texto)

        self._mostrar_placeholder(self._frame_busca, "busca")
        self._mostrar_placeholder(self._frame_lista, "lista")

    def _montar_aba_texto(self, parent):
        topo = tk.Frame(parent, bg=CORES["fundo"])
        topo.pack(fill="x", padx=16, pady=(14, 6))

        tk.Label(topo, text="Imagem do anuncio (cliente):",
                 font=FONTE_LABEL, fg=CORES["texto2"],
                 bg=CORES["fundo"]).pack(anchor="w")

        linha_ocr = tk.Frame(topo, bg=CORES["fundo"])
        linha_ocr.pack(fill="x", pady=(4, 0))

        self._lbl_img_ocr = tk.Label(linha_ocr,
            text="Nenhuma imagem selecionada",
            font=FONTE_PEQUENA, fg=CORES["texto3"],
            bg=CORES["painel2"], anchor="w",
            padx=8, pady=6, width=40)
        self._lbl_img_ocr.pack(side="left", fill="x", expand=True)

        tk.Button(linha_ocr, text="Selecionar imagem",
            font=FONTE_PEQUENA, fg=CORES["texto2"],
            bg=CORES["painel3"], relief="flat",
            cursor="hand2", padx=8, pady=6,
            activebackground=CORES["azul"],
            activeforeground="#fff",
            command=self._selecionar_img_ocr
        ).pack(side="left", padx=(6, 0))

        tk.Label(topo, text="Ou cole o texto do anuncio diretamente:",
                 font=FONTE_LABEL, fg=CORES["texto2"],
                 bg=CORES["fundo"]).pack(anchor="w", pady=(12, 4))

        self._txt_ocr = tk.Text(topo, height=5,
            font=FONTE_MONO_P,
            fg=CORES["texto"], bg=CORES["painel2"],
            insertbackground=CORES["azul"],
            relief="flat", padx=8, pady=6,
            wrap="word")
        self._txt_ocr.pack(fill="x")

        linha_ocr_info = tk.Frame(topo, bg=CORES["fundo"])
        linha_ocr_info.pack(fill="x", pady=(4, 0))
        self._lbl_ocr_status = tk.Label(linha_ocr_info, text="",
            font=FONTE_PEQUENA, fg=CORES["texto3"], bg=CORES["fundo"], anchor="w")
        self._lbl_ocr_status.pack(side="left", fill="x", expand=True)

        tk.Button(linha_ocr_info, text="Rodar OCR na imagem",
            font=FONTE_PEQUENA, fg=CORES["ciano"],
            bg=CORES["painel3"], relief="flat",
            cursor="hand2", padx=8, pady=4,
            activebackground=CORES["painel2"],
            command=self._rodar_ocr
        ).pack(side="right")

        linha_limiar = tk.Frame(topo, bg=CORES["fundo"])
        linha_limiar.pack(fill="x", pady=(10, 0))
        tk.Label(linha_limiar, text="Sensibilidade texto:",
                 font=FONTE_LABEL, fg=CORES["texto2"],
                 bg=CORES["fundo"]).pack(side="left")
        self._lbl_limiar_texto = tk.Label(linha_limiar,
            text=f"{self.limiar_texto.get():.0%}",
            font=("Segoe UI", 10, "bold"),
            fg=CORES["azul"], bg=CORES["fundo"], width=5)
        self._lbl_limiar_texto.pack(side="right")
        ttk.Scale(linha_limiar, from_=0.20, to=0.99,
                  variable=self.limiar_texto, orient="horizontal",
                  command=lambda v: self._lbl_limiar_texto.config(
                      text=f"{float(v):.0%}")
                  ).pack(side="left", fill="x", expand=True, padx=(8, 4))

        tk.Button(topo, text="Buscar texto no PDF",
            font=("Segoe UI", 11, "bold"),
            fg="#ffffff", bg=CORES["azul"],
            activebackground=CORES["roxo"],
            activeforeground="#ffffff",
            relief="flat", cursor="hand2", pady=9,
            command=self._iniciar_busca_texto
        ).pack(fill="x", pady=(12, 0))

        tk.Frame(parent, bg=CORES["borda"], height=1).pack(fill="x", padx=0, pady=(8, 0))

        self._canvas_texto, self._frame_texto = self._criar_area_rolavel(parent)

        self._lbl_status_texto = tk.Label(parent, text="",
            font=FONTE_PEQUENA, fg=CORES["texto3"],
            bg=CORES["fundo"], anchor="w")

    def _selecionar_img_ocr(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar imagem do anuncio",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.tiff *.webp"),
                       ("Todos", "*.*")])
        if caminho:
            self.caminho_img_ocr = caminho
            nome = Path(caminho).name
            nome_curto = nome if len(nome) <= 40 else "..." + nome[-37:]
            self._lbl_img_ocr.config(text=nome_curto, fg=CORES["verde"])

    def _rodar_ocr(self):
        if not hasattr(self, "caminho_img_ocr") or not self.caminho_img_ocr:
            messagebox.showwarning("Atencao", "Selecione uma imagem primeiro.")
            return
        self._lbl_ocr_status.config(text="Rodando OCR...", fg=CORES["amarelo"])
        self.update_idletasks()
        try:
            texto = ocr_da_imagem(self.caminho_img_ocr)
            self._txt_ocr.delete("1.0", tk.END)
            self._txt_ocr.insert("1.0", texto)
            self._lbl_ocr_status.config(
                text=f"OCR concluido - {len(texto.split())} palavras extraidas",
                fg=CORES["verde"])
        except Exception as e:
            self._lbl_ocr_status.config(text=f"Erro no OCR: {e}", fg=CORES["vermelho"])

    def _iniciar_busca_texto(self):
        if self._buscando:
            return

        pdf = self.caminho_pdf.get()
        if not pdf:
            messagebox.showwarning("Atencao", "Selecione um PDF no painel esquerdo.")
            return
        if not Path(pdf).exists():
            messagebox.showerror("Erro", f"PDF nao encontrado:\n{pdf}")
            return

        texto_busca = self._txt_ocr.get("1.0", tk.END).strip()
        if not texto_busca:
            messagebox.showwarning("Atencao",
                "Cole o texto do anuncio ou rode o OCR primeiro.")
            return

        self._buscando = True
        self._limpar_frame(self._frame_texto)
        self._barra.atualizar(0)
        self._lbl_status.config(text="Buscando texto...", fg=CORES["texto2"])
        self._abas.select(self._aba_texto)

        threading.Thread(
            target=buscar_texto_no_pdf,
            args=(pdf, texto_busca, self.limiar_texto.get(),
                  self._cb_progresso, self._cb_resultado_texto),
            daemon=True
        ).start()

    def _cb_resultado_texto(self, matches=None, melhor_score=0.0, total=0, erro=None):
        def atualizar():
            self._buscando = False
            self._barra.atualizar(1.0 if not erro else 0.0)

            if erro:
                self._lbl_status.config(text=f"Erro: {erro}", fg=CORES["vermelho"])
                self._mostrar_placeholder(self._frame_texto, "texto", erro)
                return

            self._lbl_progresso.config(
                text=f"{total} bloco(s) de texto analisado(s)", fg=CORES["texto3"])

            if matches:
                paginas = sorted(set(m["pagina"] for m in matches))
                self._lbl_status.config(
                    text=f"Texto encontrado em {len(paginas)} pagina(s): "
                         f"{', '.join(map(str, paginas))}",
                    fg=CORES["verde"])
                self._mostrar_matches_texto(matches)
            else:
                self._lbl_status.config(
                    text=f"Texto nao encontrado  "
                         f"(melhor: {melhor_score:.1%}  "
                         f"limiar: {self.limiar_texto.get():.0%})",
                    fg=CORES["vermelho"])
                self._mostrar_placeholder(self._frame_texto, "texto",
                    f"Texto nao encontrado. Maior similaridade: {melhor_score:.1%}")

        self.after(0, atualizar)

    def _mostrar_matches_texto(self, matches):
        self._limpar_frame(self._frame_texto)
        cab = tk.Frame(self._frame_texto, bg=CORES["painel2"])
        cab.pack(fill="x")
        tk.Label(cab, text=f"  {len(matches)} trecho(s) encontrado(s)",
                 font=("Segoe UI", 9, "bold"),
                 fg=CORES["verde"], bg=CORES["painel2"], pady=8).pack(side="left")

        for m in matches:
            self._card_texto(m)

    def _card_texto(self, dados):
        card = tk.Frame(self._frame_texto, bg=CORES["painel2"],
                        highlightbackground=CORES["borda"], highlightthickness=1)
        card.pack(fill="x", padx=8, pady=4)

        # Linha topo: titulo + score
        topo = tk.Frame(card, bg=CORES["painel2"])
        topo.pack(fill="x", padx=10, pady=(8, 4))

        tk.Label(topo, text=f"Pagina {dados['pagina']}",
                 font=("Segoe UI", 11, "bold"),
                 fg=CORES["texto"], bg=CORES["painel2"]).pack(side="left")

        score = dados["score"]
        cor_badge = (CORES["verde"] if score >= 0.75
                     else CORES["amarelo"] if score >= 0.50
                     else CORES["vermelho"])
        tk.Label(topo, text=f" {score:.1%} ",
                 font=FONTE_BADGE, fg="#0f1117",
                 bg=cor_badge, padx=4, pady=2).pack(side="right")

        criar_barra_score(card, score).pack(fill="x", padx=10, pady=(0, 4))

        tk.Frame(card, bg=CORES["borda"], height=1).pack(fill="x", padx=10)

        # Preview do anuncio (imagem recortada da pagina com borda vermelha)
        preview = dados.get("preview")
        if preview:
            try:
                foto_preview = ImageTk.PhotoImage(preview)
                self._thumbnails.append(foto_preview)
                lbl_prev = tk.Label(card, image=foto_preview,
                                    bg=CORES["painel2"], cursor="hand2")
                lbl_prev.pack(padx=10, pady=(8, 4))
            except Exception:
                pass

        tk.Label(card, text="Trecho encontrado:",
                 font=FONTE_BADGE, fg=CORES["texto3"],
                 bg=CORES["painel2"]).pack(anchor="w", padx=10, pady=(4, 2))

        txt_frame = tk.Frame(card, bg=CORES["painel3"],
                             highlightbackground=CORES["borda"], highlightthickness=1)
        txt_frame.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(txt_frame, text=dados["trecho"],
                 font=FONTE_MONO_P, fg=CORES["texto"],
                 bg=CORES["painel3"], anchor="w",
                 justify="left", wraplength=480,
                 padx=8, pady=6).pack(fill="x")

        tk.Frame(card, bg=CORES["borda"], height=1).pack(fill="x", padx=10)

        metricas = tk.Frame(card, bg=CORES["painel2"])
        metricas.pack(fill="x", padx=10, pady=(4, 8))

        def nova_linha_m(tag, cor_tag, valor):
            linha = tk.Frame(metricas, bg=CORES["painel2"])
            linha.pack(fill="x", pady=1)
            tk.Label(linha, text=tag, font=FONTE_BADGE,
                     fg=cor_tag, bg=CORES["painel3"], padx=2).pack(side="left")
            tk.Label(linha, text=valor, font=FONTE_MONO_P,
                     fg=CORES["texto2"], bg=CORES["painel2"], padx=6).pack(side="left")

        nova_linha_m(" POS  ", CORES["amarelo"],
            f"X {dados['x_cm']:.2f} cm  Y {dados['y_cm']:.2f} cm")
        nova_linha_m(" TAM  ", CORES["ciano"],
            f"L {dados['w_cm']:.2f} cm  x  A {dados['h_cm']:.2f} cm")

        # Colunagem do jornal ativo
        nome_jornal = self.jornal_ativo.get()
        jornal = JORNAIS_CADASTRADOS.get(nome_jornal)
        if jornal:
            num_col = calcular_colunas(dados["w_cm"], jornal)
            larg_col = largura_de_n_colunas(num_col, jornal)
            formato_nome = identificar_formato(
                dados["w_cm"], dados["h_cm"],
                jornal["largura"], jornal["altura"])
            nova_linha_m(" COL  ", CORES["lilas"],
                f"{num_col} col  ({larg_col:.2f} cm)  x  {dados['h_cm']:.1f} cm alt")
            nova_linha_m(" FMT  ", CORES["roxo"], formato_nome)

        if dados.get("pag_w") and dados.get("pag_h"):
            prop = calcular_proporcao(dados["w_cm"], dados["h_cm"],
                                      dados["pag_w"], dados["pag_h"])
            nova_linha_m(" PROP ", CORES["verde"],
                f"Larg: {prop['w_pct']:.1f}%  Alt: {prop['h_pct']:.1f}%  "
                f"Area: {prop['area_pct']:.1f}%")
            nova_linha_m(" PAG  ", CORES["texto3"],
                f"L {dados['pag_w']:.2f} cm  x  A {dados['pag_h']:.2f} cm")

        tk.Button(card,
            text="Baixar PDF com marcacao",
            font=("Segoe UI", 9, "bold"),
            fg="#ffffff", bg=CORES["verde"],
            activebackground="#22a870",
            activeforeground="#ffffff",
            relief="flat", cursor="hand2",
            padx=10, pady=7,
            command=lambda d=dados: self._baixar_pdf_marcado([d])
        ).pack(fill="x", padx=10, pady=(6, 10))

    def _baixar_pdf_marcado(self, lista_marcacoes):
        pdf = self.caminho_pdf.get()
        if not pdf or not Path(pdf).exists():
            messagebox.showerror("Erro", "PDF original nao encontrado.")
            return

        marcacoes_validas = [m for m in lista_marcacoes
                             if m.get("x_cm") is not None and m.get("w_cm")]
        if not marcacoes_validas:
            messagebox.showwarning("Atencao",
                "Nao ha coordenadas disponiveis para marcar neste anuncio.")
            return

        nome_sugerido = Path(pdf).stem + "_marcado.pdf"
        destino = filedialog.asksaveasfilename(
            title="Salvar PDF marcado",
            initialfile=nome_sugerido,
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])

        if not destino:
            return

        try:
            gerar_pdf_marcado(pdf, marcacoes_validas, destino)
            messagebox.showinfo("Sucesso",
                f"PDF salvo com marcacao em:\n{destino}")
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel salvar o PDF:\n{e}")

    def _criar_area_rolavel(self, parent):
        container = tk.Frame(parent, bg=CORES["fundo"])
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, bg=CORES["painel"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        frame_interno = tk.Frame(canvas, bg=CORES["painel"])
        janela = canvas.create_window((0, 0), window=frame_interno, anchor="nw")
        frame_interno.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
            lambda e: canvas.itemconfig(janela, width=e.width))
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        return canvas, frame_interno

    def _titulo_secao(self, parent, texto):
        tk.Label(parent, text=texto, font=("Segoe UI", 8, "bold"),
                 fg=CORES["azul"], bg=CORES["fundo"],
                 anchor="w").pack(fill="x", pady=(2, 5))

    def _estilizar_radio(self, botao):
        def atualizar(*_):
            for irmao in botao.master.winfo_children():
                selecionado = irmao["value"] == self.metodo_busca.get()
                irmao.config(
                    fg=CORES["azul"] if selecionado else CORES["texto2"],
                    bg=CORES["painel2"] if selecionado else CORES["fundo"])
        self.metodo_busca.trace_add("write", atualizar)
        botao.bind("<Enter>", lambda e: botao.config(fg=CORES["azul"]))
        botao.bind("<Leave>", lambda e: atualizar())

    def _centralizar_janela(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _ao_mudar_formato(self, *_):
        papel = FORMATOS_JORNAL.get(self.formato_jornal.get())
        if papel:
            self._lbl_formato_info.config(
                text=f"  {papel[0]:.1f} cm x {papel[1]:.1f} cm"
                     f"  (area: {papel[0] * papel[1]:.0f} cm2)")
        else:
            self._lbl_formato_info.config(text="  Use o DPI para calcular cm")

    def _ao_mudar_jornal(self, *_):
        nome = self.jornal_ativo.get()
        j = JORNAIS_CADASTRADOS.get(nome)
        if j:
            lc = largura_coluna(j)
            self._lbl_jornal_info.config(
                text=f"  {j['largura']}x{j['altura']} cm  "
                     f"{j['colunas']} col  "
                     f"1col = {lc:.2f} cm")
        else:
            self._lbl_jornal_info.config(text="  Jornal personalizado")

    def _abrir_cadastro_jornal(self):
        janela = tk.Toplevel(self)
        janela.title("Cadastrar Jornal")
        janela.configure(bg=CORES["fundo"])
        janela.geometry("360x300")
        janela.resizable(False, False)

        campos = {}

        def campo(label, valor_padrao):
            f = tk.Frame(janela, bg=CORES["fundo"])
            f.pack(fill="x", padx=20, pady=6)
            tk.Label(f, text=label, font=FONTE_LABEL,
                     fg=CORES["texto2"], bg=CORES["fundo"],
                     width=16, anchor="w").pack(side="left")
            var = tk.StringVar(value=str(valor_padrao))
            tk.Entry(f, textvariable=var, font=FONTE_LABEL,
                     fg=CORES["texto"], bg=CORES["painel2"],
                     insertbackground=CORES["azul"],
                     relief="flat", width=10).pack(side="left", padx=(4, 0))
            campos[label] = var

        tk.Label(janela, text="Novo Jornal", font=("Segoe UI", 12, "bold"),
                 fg=CORES["texto"], bg=CORES["fundo"]).pack(pady=(16, 8))

        campo("Nome do jornal:", "Meu Jornal")
        campo("Largura (cm):", "28.5")
        campo("Altura (cm):", "52.0")
        campo("Num. colunas:", "9")
        campo("Margem col (cm):", "0.3")

        def salvar():
            try:
                nome = campos["Nome do jornal:"].get().strip()
                if not nome:
                    raise ValueError("Nome vazio")
                novo = {
                    "largura": float(campos["Largura (cm):"].get()),
                    "altura":  float(campos["Altura (cm):"].get()),
                    "colunas": int(campos["Num. colunas:"].get()),
                    "margem":  float(campos["Margem col (cm):"].get()),
                }
                chave = f"{nome} ({novo['largura']}x{novo['altura']}cm, {novo['colunas']}col)"
                JORNAIS_CADASTRADOS[chave] = novo

                opcoes = list(JORNAIS_CADASTRADOS.keys()) + ["Personalizado"]
                for widget in self._painel_esq.winfo_children():
                    if isinstance(widget, ttk.Combobox):
                        widget["values"] = opcoes
                        break

                self.jornal_ativo.set(chave)
                janela.destroy()
            except ValueError as e:
                messagebox.showerror("Erro", f"Valor invalido: {e}", parent=janela)

        tk.Button(janela, text="Salvar",
            font=("Segoe UI", 10, "bold"),
            fg="#fff", bg=CORES["azul"],
            relief="flat", cursor="hand2", pady=8,
            command=salvar).pack(fill="x", padx=20, pady=(12, 0))

    def _selecionar_pdf(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar PDF",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        if caminho:
            self.caminho_pdf.set(caminho)
            self._btn_pdf.mostrar_arquivo(Path(caminho).name)

    def _selecionar_imagem(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar imagem de busca",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.tiff *.webp"),
                       ("Todos", "*.*")])
        if caminho:
            self.caminho_img.set(caminho)
            self._btn_img.mostrar_arquivo(Path(caminho).name)
            try:
                img = Image.open(caminho).convert("RGB")
                img.thumbnail((200, 72), Image.LANCZOS)
                foto = ImageTk.PhotoImage(img)
                self._lbl_preview.config(image=foto)
                self._lbl_preview._foto = foto
            except Exception:
                pass

    def _iniciar_busca(self):
        if self._buscando:
            return
        pdf = self.caminho_pdf.get()
        img = self.caminho_img.get()
        if not pdf:
            messagebox.showwarning("Atencao", "Selecione um arquivo PDF primeiro.")
            return
        if not img:
            messagebox.showwarning("Atencao", "Selecione a imagem de busca.")
            return
        if not Path(pdf).exists():
            messagebox.showerror("Erro", f"Arquivo nao encontrado:\n{pdf}")
            return
        if not Path(img).exists():
            messagebox.showerror("Erro", f"Imagem nao encontrada:\n{img}")
            return
        try:
            imagem_busca = Image.open(img).convert("RGB")
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel abrir a imagem:\n{e}")
            return
        self._buscando = True
        self._btn_buscar.config(text="Analisando...", state="disabled",
                                bg=CORES["painel2"], fg=CORES["texto2"])
        self._limpar_frame(self._frame_busca)
        self._barra.atualizar(0)
        self._lbl_status.config(text="Analisando...", fg=CORES["texto2"])
        self._abas.select(self._aba_busca)
        threading.Thread(
            target=buscar_imagem_no_pdf,
            args=(pdf, imagem_busca, METODOS[self.metodo_busca.get()],
                  self.limiar.get(), self._cb_progresso, self._cb_resultado_busca),
            daemon=True
        ).start()

    def _iniciar_listagem(self):
        if self._buscando:
            return
        pdf = self.caminho_pdf.get()
        if not pdf:
            messagebox.showwarning("Atencao", "Selecione um arquivo PDF primeiro.")
            return
        if not Path(pdf).exists():
            messagebox.showerror("Erro", f"Arquivo nao encontrado:\n{pdf}")
            return
        self._buscando = True
        self._btn_listar.config(text="Escaneando...", state="disabled",
                                bg=CORES["painel2"], fg=CORES["texto3"])
        self._limpar_frame(self._frame_lista)
        self._barra.atualizar(0)
        self._lbl_status.config(text="Escaneando o PDF...", fg=CORES["texto2"])
        self._abas.select(self._aba_lista)
        threading.Thread(
            target=listar_todos_anuncios,
            args=(pdf, self._cb_progresso, self._cb_resultado_listagem),
            daemon=True
        ).start()

    def _cb_progresso(self, atual, total, pagina):
        def atualizar():
            self._barra.atualizar(atual / total)
            self._lbl_progresso.config(
                text=f"Item {atual}/{total}  Pagina {pagina}",
                fg=CORES["texto3"])
            self._lbl_status.config(
                text=f"Processando... {atual / total:.0%}",
                fg=CORES["texto2"])
        self.after(0, atualizar)

    def _cb_resultado_busca(self, matches=None, melhor_score=0.0, total=0, erro=None):
        def atualizar():
            self._buscando = False
            self._btn_buscar.config(text="Iniciar Busca", state="normal",
                                    bg=CORES["azul"], fg="#ffffff")
            self._barra.atualizar(1.0 if not erro else 0.0)
            if erro:
                self._lbl_status.config(text=f"Erro: {erro}", fg=CORES["vermelho"])
                self._mostrar_placeholder(self._frame_busca, "busca", erro)
                return
            self._lbl_progresso.config(
                text=f"{total} item(ns) analisado(s)", fg=CORES["texto3"])
            if matches:
                paginas = sorted(set(m["pagina"] for m in matches))
                self._lbl_status.config(
                    text=f"Encontrado em {len(paginas)} pagina(s): "
                         f"{', '.join(map(str, paginas))}",
                    fg=CORES["verde"])
                self._mostrar_matches(matches)
            else:
                self._lbl_status.config(
                    text=f"Nao encontrado  "
                         f"(melhor: {melhor_score:.1%}  "
                         f"limiar: {self.limiar.get():.0%})",
                    fg=CORES["vermelho"])
                self._mostrar_nao_encontrado(melhor_score)
        self.after(0, atualizar)

    def _cb_resultado_listagem(self, imagens=None, total=0, erro=None):
        def atualizar():
            self._buscando = False
            self._btn_listar.config(text="Listar Todos os Anuncios",
                                    state="normal",
                                    bg=CORES["painel2"], fg=CORES["texto2"])
            self._barra.atualizar(1.0 if not erro else 0.0)
            if erro:
                self._lbl_status.config(text=f"Erro: {erro}", fg=CORES["vermelho"])
                self._mostrar_placeholder(self._frame_lista, "lista", erro)
                return
            self._lbl_progresso.config(
                text=f"{total} item(ns) encontrado(s)", fg=CORES["texto3"])
            self._lbl_status.config(
                text=f"{total} anuncio(s) listado(s)", fg=CORES["ciano"])
            self._mostrar_todos(imagens)
        self.after(0, atualizar)

    def _limpar_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()
        frame.update_idletasks()
        self._thumbnails.clear()

    def _mostrar_placeholder(self, frame, aba, mensagem=None):
        self._limpar_frame(frame)
        if mensagem:
            texto = mensagem
        elif aba == "busca":
            texto = "Selecione PDF e imagem, depois clique em Iniciar Busca"
        elif aba == "texto":
            texto = "Cole o texto ou rode o OCR, depois clique em Buscar texto no PDF"
        else:
            texto = "Selecione um PDF e clique em Listar Todos os Anuncios"
        tk.Label(frame, text=texto, font=FONTE_LABEL,
                 fg=CORES["texto3"], bg=CORES["painel"],
                 justify="center", pady=40).pack(expand=True, fill="both")

    def _mostrar_nao_encontrado(self, melhor_score):
        self._limpar_frame(self._frame_busca)
        f = tk.Frame(self._frame_busca, bg=CORES["painel"])
        f.pack(fill="both", expand=True, padx=16, pady=24)
        tk.Label(f, text="X", font=("Segoe UI", 36),
                 fg=CORES["vermelho"], bg=CORES["painel"]).pack()
        tk.Label(f, text="Anuncio nao encontrado",
                 font=("Segoe UI", 13, "bold"),
                 fg=CORES["texto"], bg=CORES["painel"]).pack(pady=(4, 0))
        tk.Label(f, text=f"Maior similaridade: {melhor_score:.1%}\n"
                          "Tente reduzir o limiar de sensibilidade.",
                 font=FONTE_PEQUENA, fg=CORES["texto2"],
                 bg=CORES["painel"], justify="center").pack(pady=(6, 0))

    def _mostrar_matches(self, matches):
        self._limpar_frame(self._frame_busca)
        cab = tk.Frame(self._frame_busca, bg=CORES["painel2"])
        cab.pack(fill="x")
        tk.Label(cab, text=f"  {len(matches)} correspondencia(s) encontrada(s)",
                 font=("Segoe UI", 9, "bold"),
                 fg=CORES["verde"], bg=CORES["painel2"], pady=8).pack(side="left")
        for match in sorted(matches, key=lambda x: x["score"], reverse=True):
            self._criar_card(self._frame_busca, match, eh_match=True)

    def _mostrar_todos(self, imagens):
        self._limpar_frame(self._frame_lista)
        cab = tk.Frame(self._frame_lista, bg=CORES["painel2"])
        cab.pack(fill="x")
        tk.Label(cab, text=f"  {len(imagens)} anuncio(s) no PDF",
                 font=("Segoe UI", 9, "bold"),
                 fg=CORES["ciano"], bg=CORES["painel2"], pady=8).pack(side="left")
        for info_img in imagens:
            self._criar_card(self._frame_lista, info_img, eh_match=False)

    def _criar_card(self, parent, dados, eh_match):
        card = tk.Frame(parent, bg=CORES["painel2"],
                        highlightbackground=CORES["borda"], highlightthickness=1)
        card.pack(fill="x", padx=8, pady=4)

        # Preview do anuncio: recorte da pagina com borda vermelha,
        # ou a propria imagem embutida se nao houver coordenadas
        preview = dados.get("preview")
        if preview:
            try:
                foto_prev = ImageTk.PhotoImage(preview)
                self._thumbnails.append(foto_prev)
                tk.Label(card, image=foto_prev,
                         bg=CORES["painel2"]).pack(padx=10, pady=(8, 4))
            except Exception:
                pass

        # Thumb pequeno na coluna esquerda (sempre mostra)
        area_thumb = tk.Frame(card, bg=CORES["painel2"], width=90)
        area_thumb.pack(side="left", padx=(10, 0), pady=8)
        area_thumb.pack_propagate(False)

        img_thumb = dados.get("thumb")
        if img_thumb:
            try:
                foto = ImageTk.PhotoImage(img_thumb)
                self._thumbnails.append(foto)
                tk.Label(area_thumb, image=foto, bg=CORES["painel2"]).pack(expand=True)
            except Exception:
                tk.Label(area_thumb, text="IMG", font=("Segoe UI", 22),
                         fg=CORES["texto3"], bg=CORES["painel2"]).pack(expand=True)
        else:
            tk.Label(area_thumb, text="IMG", font=("Segoe UI", 22),
                     fg=CORES["texto3"], bg=CORES["painel2"]).pack(expand=True)

        area_info = tk.Frame(card, bg=CORES["painel2"])
        area_info.pack(side="left", fill="both", expand=True, padx=10, pady=8)

        linha_titulo = tk.Frame(area_info, bg=CORES["painel2"])
        linha_titulo.pack(fill="x")

        if eh_match:
            tk.Label(linha_titulo, text=f"Pagina {dados['pagina']}",
                     font=("Segoe UI", 11, "bold"),
                     fg=CORES["texto"], bg=CORES["painel2"]).pack(side="left")
            score = dados["score"]
            cor_badge = (CORES["verde"] if score >= 0.85
                         else CORES["amarelo"] if score >= 0.65
                         else CORES["vermelho"])
            tk.Label(linha_titulo, text=f" {score:.1%} ",
                     font=FONTE_BADGE, fg="#0f1117",
                     bg=cor_badge, padx=4, pady=2).pack(side="right")
        else:
            tk.Label(linha_titulo,
                     text=f"#{dados['numero']}  Pag. {dados['pagina']}",
                     font=("Segoe UI", 11, "bold"),
                     fg=CORES["texto"], bg=CORES["painel2"]).pack(side="left")

        fonte = dados.get("fonte", "renderizada")
        txt_fonte = "Imagem embutida" if fonte == "embutida" else "Texto / Vetorial"
        cor_fonte = CORES["ciano"] if fonte == "embutida" else CORES["amarelo"]
        tk.Label(area_info, text=txt_fonte, font=FONTE_BADGE,
                 fg=cor_fonte, bg=CORES["painel2"]).pack(anchor="w", pady=(2, 0))

        if eh_match:
            criar_barra_score(area_info, dados["score"]).pack(fill="x", pady=(3, 4))

        tk.Frame(area_info, bg=CORES["borda"], height=1).pack(fill="x", pady=(2, 4))

        bloco = criar_bloco_metricas(
            area_info,
            w_px        = dados.get("w_px", 0),
            h_px        = dados.get("h_px", 0),
            w_cm_pdf    = dados.get("w_cm"),
            h_cm_pdf    = dados.get("h_cm"),
            x_cm        = dados.get("x_cm"),
            y_cm        = dados.get("y_cm"),
            pag_w_cm    = dados.get("pag_w"),
            pag_h_cm    = dados.get("pag_h"),
            var_formato = self.formato_jornal,
            var_dpi     = self.dpi,
        )
        bloco.pack(anchor="w", fill="x", pady=(0, 2))

        nome_jornal = self.jornal_ativo.get()
        jornal = JORNAIS_CADASTRADOS.get(nome_jornal)
        w_cm_real = dados.get("w_cm") or 0
        h_cm_real = dados.get("h_cm") or 0
        if jornal and w_cm_real and h_cm_real:
            num_col    = calcular_colunas(w_cm_real, jornal)
            larg_col   = largura_de_n_colunas(num_col, jornal)
            fmt_nome   = identificar_formato(
                w_cm_real, h_cm_real, jornal["largura"], jornal["altura"])

            frame_col = tk.Frame(area_info, bg=CORES["painel2"])
            frame_col.pack(fill="x", pady=(2, 0))

            def linha_col(tag, cor, val):
                l = tk.Frame(frame_col, bg=CORES["painel2"])
                l.pack(fill="x", pady=1)
                tk.Label(l, text=tag, font=FONTE_BADGE,
                         fg=cor, bg=CORES["painel3"], padx=2).pack(side="left")
                tk.Label(l, text=val, font=FONTE_MONO_P,
                         fg=CORES["texto2"], bg=CORES["painel2"],
                         padx=6).pack(side="left")

            linha_col(" COL  ", CORES["lilas"],
                f"{num_col} col  ({larg_col:.2f} cm)  x  {h_cm_real:.1f} cm alt")
            linha_col(" FMT  ", CORES["roxo"], fmt_nome)

        tk.Button(area_info,
            text="Baixar PDF com marcacao", 
            font=("Segoe UI", 9, "bold"),
            fg="#ffffff", bg=CORES["verde"],
            activebackground="#22a870",
            activeforeground="#ffffff",
            relief="flat", cursor="hand2",
            padx=10, pady=7,
            command=lambda d=dados: self._baixar_pdf_marcado([d])
        ).pack(fill="x", pady=(8, 2))


if __name__ == "__main__":
    App().mainloop()
