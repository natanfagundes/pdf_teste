import sys
import io
import os
import hashlib
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


def checar_dependencias():
    faltando = []
    # Dependências obrigatórias — app não inicia sem elas
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
    # spaCy é opcional — NER usa fallback se não estiver disponível
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
from PIL import Image, ImageTk, ImageEnhance, ImageFilter
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
    "Meu Jornal (28.5x52cm, 6col)": {
        "largura": 28.5, "altura": 52.0,
        "colunas": 6,    "margem": 0.3,
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


def identificar_coluna_inicial(x_cm, jornal):
    """Retorna o numero da coluna (1-based) onde o item comeca."""
    lc = largura_coluna(jornal)
    passo = lc + jornal["margem"]
    if passo <= 0:
        return 1
    col = int(x_cm / passo) + 1
    return max(1, min(col, jornal["colunas"]))


def calcular_faixa_colunas(x_cm, w_cm, jornal):
    """Retorna (col_inicial, col_final, num_colunas_largura)."""
    col_ini = identificar_coluna_inicial(x_cm or 0.0, jornal)
    num_col = calcular_colunas(w_cm, jornal)
    col_fim = min(col_ini + num_col - 1, jornal["colunas"])
    return col_ini, col_fim, num_col


def calcular_info_colunas(w_cm, x_cm, jornal):
    """Calcula colunagem com maxima precisao.
    Retorna dict:
      num_col    - numero inteiro de colunas (melhor ajuste)
      cols_exato - colunas exatas como float
      larg_padrao- largura teorica para num_col colunas (cm)
      larg_1col  - largura de 1 coluna (cm)
      margem     - margem entre colunas (cm)
      desvio     - diferenca entre medido e padrao (cm)
      col_ini    - coluna inicial (1-based)
      col_fim    - coluna final (1-based)
    """
    lc    = largura_coluna(jornal)
    marg  = jornal["margem"]
    passo = lc + marg   # largura de 1 col + 1 margem

    # Formula exata: w = lc*n + marg*(n-1)  =>  n = (w + marg) / (lc + marg)
    cols_exato  = (w_cm + marg) / passo if passo > 0 else 1.0
    # Arredonda para o numero de colunas mais proximo dentro do limite do jornal
    num_col     = max(1, min(round(cols_exato), jornal["colunas"]))
    larg_padrao = lc * num_col + marg * (num_col - 1)
    desvio      = w_cm - larg_padrao

    col_ini = identificar_coluna_inicial(x_cm or 0.0, jornal)
    col_fim = min(col_ini + num_col - 1, jornal["colunas"])

    return {
        "num_col":     num_col,
        "cols_exato":  cols_exato,
        "larg_padrao": larg_padrao,
        "larg_1col":   lc,
        "margem":      marg,
        "desvio":      desvio,
        "col_ini":     col_ini,
        "col_fim":     col_fim,
    }


def identificar_formato(w_cm, h_cm, pag_w, pag_h):
    if pag_w <= 0 or pag_h <= 0:
        return "Desconhecido"
    proporcao = (w_cm * h_cm) / (pag_w * pag_h)
    for nome, minimo, maximo in FORMATOS_NOMEADOS:
        if minimo <= proporcao < maximo:
            return nome
    return "Faixa"

PT_PARA_CM      = 2.54 / 72   # = 0.03527̄ cm/pt  (1" = 72pt = 2.54cm)
TAMANHO_PHASH   = 32
TAMANHO_SSIM    = (512, 512)
MIN_MATCHES_ORB = 20
DPI_RENDER      = 200
DPI_SCAN        = 72   # DPI para varredura rapida por template matching (grayscale)

# ── Cache de paginas renderizadas ─────────────────────────────────────────────
# Evita re-abrir/re-renderizar o PDF a cada chamada de gerar_preview_anuncio
# e a cada nova busca no mesmo arquivo. Armazena em ~/.dimensao_cache/
_CACHE_DIR = Path.home() / ".dimensao_cache"


def _hash_pdf(pdf_path: str) -> str:
    """Hash do PDF baseado em caminho absoluto + tamanho + mtime."""
    p = Path(pdf_path)
    try:
        stat = p.stat()
        chave = f"{p.resolve()}|{stat.st_size}|{stat.st_mtime}"
    except Exception:
        chave = str(p.resolve())
    return hashlib.md5(chave.encode()).hexdigest()[:16]


def _cache_np_path(pdf_hash: str, num_pagina: int, dpi: int, gray: bool) -> Path:
    suf = "g" if gray else ""
    return _CACHE_DIR / f"{pdf_hash}_p{num_pagina:04d}_d{dpi}{suf}.npy"


def _arr_do_cache(pdf_hash: str, num_pagina: int, dpi: int, gray: bool = False):
    """Retorna array numpy lido do cache em disco, ou None se nao existir."""
    p = _cache_np_path(pdf_hash, num_pagina, dpi, gray)
    if p.exists():
        try:
            return np.load(str(p))
        except Exception:
            pass
    return None


def _salvar_no_cache(arr: np.ndarray, pdf_hash: str, num_pagina: int,
                     dpi: int, gray: bool = False) -> None:
    """Persiste array no cache. Remove os mais antigos se ultrapassar 400 arquivos."""
    try:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        p = _cache_np_path(pdf_hash, num_pagina, dpi, gray)
        np.save(str(p), arr)
        arquivos = sorted(_CACHE_DIR.glob("*.npy"), key=lambda f: f.stat().st_atime)
        for f in arquivos[:-400]:
            try:
                f.unlink()
            except Exception:
                pass
    except Exception:
        pass


def _get_pagina_arr(pdf_path: str, num_pagina: int, dpi: int) -> np.ndarray:
    """Array RGB (H x W x 3) uint8 da pagina (1-based), com cache em disco."""
    pdf_hash = _hash_pdf(pdf_path)
    arr = _arr_do_cache(pdf_hash, num_pagina, dpi)
    if arr is None:
        doc = fitz.open(pdf_path)
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = doc[num_pagina - 1].get_pixmap(matrix=mat, alpha=False)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, 3).copy()
        doc.close()
        _salvar_no_cache(arr, pdf_hash, num_pagina, dpi)
    return arr
# ──────────────────────────────────────────────────────────────────────────────


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
    orb = cv2.ORB_create(nfeatures=2000)
    kp_a, desc_a = orb.detectAndCompute(cinza_a, None)
    kp_b, desc_b = orb.detectAndCompute(cinza_b, None)
    if desc_a is None or desc_b is None or len(kp_a) < 5 or len(kp_b) < 5:
        return 0.0
    matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    bons = [m for m in matcher.match(desc_a, desc_b) if m.distance < 55]
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
        comparar_phash(img_a, img_b) * 0.45 +
        comparar_orb(img_a, img_b)   * 0.30 +
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

        # Usa cache: nao reabre nem re-renderiza o PDF se a pagina ja foi processada
        arr = _get_pagina_arr(pdf_path, num_pagina, DPI_RENDER)
        img_pagina = Image.fromarray(arr, "RGB")

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

        # Borda azul fina ao redor do anuncio
        draw = ImageDraw.Draw(img_pagina)
        espessura = max(3, int(px_por_cm * 0.10))
        for i in range(espessura):
            draw.rectangle(
                [px - i, py - i, px + pw + i, py + ph + i],
                outline=(69, 161, 237)
            )

        # Recorta somente a regiao do anuncio com margem de contexto
        margem_px = int(px_por_cm * 0.5)
        crop_x0 = max(0, px - margem_px)
        crop_y0 = max(0, py - margem_px)
        crop_x1 = min(img_pagina.width,  px + pw + margem_px)
        crop_y1 = min(img_pagina.height, py + ph + margem_px)
        img_crop = img_pagina.crop((crop_x0, crop_y0, crop_x1, crop_y1))
        if img_crop.width < 5 or img_crop.height < 5:
            return None

        ratio = largura_thumb / img_crop.width
        nova_h = int(img_crop.height * ratio)
        return img_crop.resize((largura_thumb, nova_h), Image.LANCZOS)

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


def _query_para_cinza(img_pil):
    """Converte PIL Image para numpy uint8 grayscale."""
    return cv2.cvtColor(np.array(img_pil.convert("RGB")), cv2.COLOR_RGB2GRAY)


def _template_match_piramide(arr_query, arr_pagina):
    """Template matching multi-escala com piramide de imagem (coarse-to-fine).

    Estrategia:
    1. Busca grossa: 11 proporcoes testadas em nivel baixo da piramide (~1/4 res)
       -> muito rapido porque as imagens sao minusculas
    2. Refinamento: top-3 candidatos da busca grossa sao refinados em sub-janela
       da resolucao original -> preciso sem varrer a imagem inteira

    Retorna (best_score, x_px, y_px, w_px, h_px) em coords da imagem original.
    """
    qh, qw = arr_query.shape[:2]
    ph, pw = arr_pagina.shape[:2]
    if qw < 8 or qh < 8 or pw < 20 or ph < 20:
        return 0.0, 0, 0, qw, qh

    # Escolhe nivel da piramide para busca grossa:
    # pagina com ~120-200px de largura e o sweet-spot (rapido e ainda legivel)
    nivel = 0
    while (pw >> (nivel + 1)) >= 150 and nivel < 3:
        nivel += 1
    fator = 1 << nivel  # 2^nivel

    pag_coarse = arr_pagina
    for _ in range(nivel):
        if pag_coarse.shape[1] < 30 or pag_coarse.shape[0] < 30:
            break
        pag_coarse = cv2.pyrDown(pag_coarse)
    pch, pcw = pag_coarse.shape[:2]

    # Busca grossa: testa 11 proporcoes da largura da pagina coarse
    candidatos_grossos = []
    for pct in [0.06, 0.10, 0.15, 0.21, 0.28, 0.37, 0.47, 0.58, 0.70, 0.82, 0.93]:
        nw = int(pct * pcw)
        nh = int(nw * qh / max(qw, 1))
        if nw < 8 or nh < 8 or nw >= pcw or nh >= pch:
            continue
        q_c = cv2.resize(arr_query, (nw, nh), interpolation=cv2.INTER_AREA)
        try:
            res = cv2.matchTemplate(pag_coarse, q_c, cv2.TM_CCOEFF_NORMED)
            _, val, _, loc = cv2.minMaxLoc(res)
        except cv2.error:
            continue
        candidatos_grossos.append((val, pct, loc))

    if not candidatos_grossos:
        return 0.0, 0, 0, qw, qh

    # Refinamento: top-3 candidatos na resolucao original, em sub-janela
    candidatos_grossos.sort(reverse=True)
    best_score = 0.0
    best_loc   = (0, 0, qw, qh)

    for val_c, pct, loc_c in candidatos_grossos[:3]:
        nw_full = int(pct * pw)
        nh_full = int(nw_full * qh / max(qw, 1))
        if nw_full < 8 or nh_full < 8 or nw_full >= pw or nh_full >= ph:
            continue

        # Janela de refinamento ao redor do match grosseiro
        cx = loc_c[0] * fator
        cy = loc_c[1] * fator
        mg = max(fator * 6, 12)
        rx0 = max(0, cx - mg)
        ry0 = max(0, cy - mg)
        rx1 = min(pw - nw_full, cx + nw_full + mg)
        ry1 = min(ph - nh_full, cy + nh_full + mg)

        if rx1 > rx0 and ry1 > ry0:
            patch    = arr_pagina[ry0:ry1 + nh_full, rx0:rx1 + nw_full]
            offset_x = rx0
            offset_y = ry0
        else:
            patch    = arr_pagina
            offset_x = offset_y = 0

        q_full = cv2.resize(arr_query, (nw_full, nh_full), interpolation=cv2.INTER_LINEAR)
        try:
            res = cv2.matchTemplate(patch, q_full, cv2.TM_CCOEFF_NORMED)
            _, val_f, _, loc_f = cv2.minMaxLoc(res)
        except cv2.error:
            continue

        if val_f > best_score:
            best_score = val_f
            best_loc   = (loc_f[0] + offset_x, loc_f[1] + offset_y, nw_full, nh_full)

    return max(0.0, float(best_score)), best_loc[0], best_loc[1], best_loc[2], best_loc[3]


def buscar_imagem_no_pdf(pdf_path, imagens_busca, funcao_comparacao,
                         limiar, cb_progresso, cb_resultado):
    """imagens_busca: PIL.Image ou lista de PIL.Images.

    Estrategia multi-camada:
    1. Imagens embutidas no PDF: comparacao por funcao_comparacao (phash/orb/ssim/all)
    2. Template matching com piramide coarse-to-fine + cache de paginas em disco
    3. OCR hibrido: fallback textual quando o visual falha
       - Extrai texto OCR das imagens de busca uma unica vez
       - Compara contra texto nativo do PDF (rapido, sem re-OCR)
       - Se texto bate, localiza o bloco de texto na pagina para ter coordenadas
    """
    if not isinstance(imagens_busca, list):
        imagens_busca = [imagens_busca]

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        cb_resultado(erro=str(e))
        return

    num_paginas = len(doc)
    if num_paginas == 0:
        cb_resultado(erro="PDF vazio.")
        return

    # Pre-processa queries: grayscale numpy para template matching
    queries_cinza = [_query_para_cinza(img) for img in imagens_busca]

    # OCR nas imagens de busca (feito uma unica vez, antes do loop de paginas)
    textos_query_ocr = []
    for img_q in imagens_busca:
        try:
            textos_query_ocr.append(ocr_de_pil(img_q))
        except Exception:
            textos_query_ocr.append("")
    # Liga o fallback OCR apenas se pelo menos uma query tem texto suficiente
    tem_ocr = any(len(t.split()) >= 5 for t in textos_query_ocr)

    pdf_hash       = _hash_pdf(pdf_path)
    px_por_cm_scan = DPI_SCAN / 2.54
    mat_scan       = fitz.Matrix(DPI_SCAN / 72, DPI_SCAN / 72)

    matches_raw  = []
    melhor_score = 0.0

    for num_pagina in range(num_paginas):
        pagina = doc[num_pagina]
        pag_w, pag_h = pegar_tamanho_pagina_cm(pagina)

        # ── 1. Imagens embutidas ──────────────────────────────────────────
        for info_img in pagina.get_images(full=True):
            xref = info_img[0]
            try:
                dados   = doc.extract_image(xref)
                img_emb = Image.open(io.BytesIO(dados["image"])).convert("RGB")
                if img_emb.width < 40 or img_emb.height < 40:
                    continue
                best_s = 0.0; best_qi = 0
                for qi, img_q in enumerate(imagens_busca):
                    try:
                        s = funcao_comparacao(img_q, img_emb)
                    except Exception:
                        s = 0.0
                    if s > best_s:
                        best_s = s; best_qi = qi
                if best_s > melhor_score:
                    melhor_score = best_s
                if best_s >= limiar:
                    pos = pegar_posicao_imagem_cm(pagina, xref)
                    matches_raw.append({
                        "pagina": num_pagina + 1, "score": best_s,
                        "x_cm":   pos["x_cm"] if pos else 0.0,
                        "y_cm":   pos["y_cm"] if pos else 0.0,
                        "w_cm":   pos["w_cm"] if pos else pontos_para_cm(img_emb.width),
                        "h_cm":   pos["h_cm"] if pos else pontos_para_cm(img_emb.height),
                        "pag_w":  pag_w, "pag_h": pag_h,
                        "qi":     best_qi, "fonte": "embutida",
                    })
            except Exception:
                pass

        # ── 2. Template matching com piramide + cache ─────────────────────
        # Usa cache: na 2a busca no mesmo PDF, carrega do disco (sem re-renderizar)
        arr_scan = _arr_do_cache(pdf_hash, num_pagina + 1, DPI_SCAN, gray=True)
        if arr_scan is None:
            try:
                pix = pagina.get_pixmap(matrix=mat_scan, alpha=False,
                                        colorspace=fitz.csGRAY)
                arr_scan = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                    pix.height, pix.width).copy()
                _salvar_no_cache(arr_scan, pdf_hash, num_pagina + 1, DPI_SCAN, gray=True)
            except Exception:
                arr_scan = None

        best_template = None
        if arr_scan is not None:
            for qi, arr_q in enumerate(queries_cinza):
                try:
                    s, tx, ty, tw, th = _template_match_piramide(arr_q, arr_scan)
                except Exception:
                    s = 0.0; tx = ty = tw = th = 0
                if s > melhor_score:
                    melhor_score = s
                if best_template is None or s > best_template["score"]:
                    best_template = {
                        "score": s, "qi": qi,
                        "x_cm":  tx / px_por_cm_scan,
                        "y_cm":  ty / px_por_cm_scan,
                        "w_cm":  tw / px_por_cm_scan,
                        "h_cm":  th / px_por_cm_scan,
                    }

        if best_template and best_template["score"] >= limiar:
            matches_raw.append({
                "pagina": num_pagina + 1, "pag_w": pag_w, "pag_h": pag_h,
                "fonte": "template",
                **{k: best_template[k] for k in ("score","x_cm","y_cm","w_cm","h_cm","qi")},
            })

        # ── 3. OCR hibrido: fallback quando template nao encontrou ────────
        # Compara o texto OCR das imagens de busca com o texto nativo do PDF.
        # Nao exige re-OCR do PDF (usa get_text do fitz, muito rapido).
        template_achou = best_template and best_template["score"] >= limiar
        if tem_ocr and not template_achou:
            texto_pag = pagina.get_text("text").strip()
            if texto_pag:
                for qi, texto_q in enumerate(textos_query_ocr):
                    if not texto_q or len(texto_q.split()) < 5:
                        continue
                    ts = calcular_score_texto(texto_q, texto_pag)
                    score_ocr = ts * 0.82  # desconto: localizacao menos precisa
                    if score_ocr > melhor_score:
                        melhor_score = score_ocr
                    if score_ocr >= limiar:
                        # Encontra o bloco com maior similaridade para ter coordenadas
                        blocos_pag = [b for b in pagina.get_text("blocks")
                                      if len(b[4].strip()) >= 10]
                        melhor_b = max(
                            blocos_pag,
                            key=lambda b: calcular_score_texto(texto_q, b[4]),
                            default=None
                        )
                        if melhor_b:
                            bx = pontos_para_cm(melhor_b[0])
                            by = pontos_para_cm(melhor_b[1])
                            bw = pontos_para_cm(melhor_b[2] - melhor_b[0])
                            bh = pontos_para_cm(melhor_b[3] - melhor_b[1])
                        else:
                            bx, by, bw, bh = 0.0, 0.0, pag_w, pag_h
                        matches_raw.append({
                            "pagina": num_pagina + 1, "score": score_ocr,
                            "x_cm": bx, "y_cm": by, "w_cm": bw, "h_cm": bh,
                            "pag_w": pag_w, "pag_h": pag_h,
                            "qi": qi, "fonte": "ocr",
                        })

        cb_progresso(num_pagina + 1, num_paginas, num_pagina + 1)

    doc.close()

    # Deduplica: mantem apenas o melhor score por regiao proxima
    matches_raw.sort(key=lambda x: x["score"], reverse=True)
    regioes_vistas = []
    matches_finais = []
    for m in matches_raw:
        sobrepoe = any(
            m["pagina"] == r["pagina"]
            and abs(m["x_cm"] - r["x_cm"]) < 2.0
            and abs(m["y_cm"] - r["y_cm"]) < 2.0
            for r in regioes_vistas
        )
        if sobrepoe:
            continue
        regioes_vistas.append({"pagina": m["pagina"], "x_cm": m["x_cm"], "y_cm": m["y_cm"]})

        preview = gerar_preview_anuncio(
            pdf_path, m["pagina"], m["x_cm"], m["y_cm"], m["w_cm"], m["h_cm"])
        if preview is None:
            preview = imagens_busca[m["qi"]].copy()
            preview.thumbnail((200, 400), Image.LANCZOS)

        img_q = imagens_busca[m["qi"]]
        matches_finais.append({
            **m,
            "preview":       preview,
            "thumb":         img_q.copy(),
            "w_px":          img_q.width,
            "h_px":          img_q.height,
            "img_match_idx": m["qi"] + 1,
        })

    cb_resultado(matches=matches_finais, melhor_score=melhor_score,
                 total=num_paginas, erro=None)


def _agrupar_regioes(regioes_pt, tolerancia=8):
    """Une retangulos proximos iterativamente ate estabilizar.
    Cada regiao: (x0, y0, x1, y1, meta) onde meta e dict."""
    grupos = list(regioes_pt)
    mudou = True
    while mudou:
        mudou = False
        novos = []
        usado = [False] * len(grupos)
        for i, g in enumerate(grupos):
            if usado[i]:
                continue
            x0, y0, x1, y1, meta = g
            usado[i] = True
            for j in range(i + 1, len(grupos)):
                if usado[j]:
                    continue
                gx0, gy0, gx1, gy1, _ = grupos[j]
                # Verifica sobreposicao com tolerancia
                if (gx0 <= x1 + tolerancia and gx1 >= x0 - tolerancia and
                        gy0 <= y1 + tolerancia and gy1 >= y0 - tolerancia):
                    x0 = min(x0, gx0); y0 = min(y0, gy0)
                    x1 = max(x1, gx1); y1 = max(y1, gy1)
                    usado[j] = True
                    mudou = True
            novos.append((x0, y0, x1, y1, meta))
        grupos = novos
    return grupos


def listar_todos_anuncios(pdf_path, cb_progresso, cb_resultado):
    """Identifica todos os anuncios de TODAS as paginas do PDF.
    Combina imagens embutidas e blocos de texto por regiao de pagina,
    agrupando elementos proximos em um unico anuncio."""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        cb_resultado(erro=str(e))
        return

    num_paginas = len(doc)
    candidatos_por_pagina = []  # lista de listas

    for num_pagina in range(num_paginas):
        pg = num_pagina + 1
        pagina = doc[num_pagina]
        pag_w, pag_h = pegar_tamanho_pagina_cm(pagina)

        regioes = []  # (x0_pt, y0_pt, x1_pt, y1_pt, meta)

        # 1) Imagens embutidas
        for info_img in pagina.get_images(full=True):
            xref = info_img[0]
            try:
                rects = pagina.get_image_rects(xref)
                if not rects:
                    continue
                r = rects[0]
                if r.width < 10 or r.height < 10:
                    continue
                regioes.append((r.x0, r.y0, r.x1, r.y1,
                                {"fonte": "embutida", "xref": xref}))
            except Exception:
                pass

        # 2) Blocos de texto com texto significativo
        for bloco in pagina.get_text("blocks"):
            x0, y0, x1, y1, texto = bloco[0], bloco[1], bloco[2], bloco[3], bloco[4]
            texto = texto.strip()
            if not texto or len(texto) < 10:
                continue
            w_pt = x1 - x0; h_pt = y1 - y0
            if w_pt < 20 or h_pt < 8:
                continue
            # Ignora numeros de pagina isolados
            if len(texto) <= 4 and texto.isdigit():
                continue
            regioes.append((x0, y0, x1, y1, {"fonte": "texto", "texto": texto}))

        if not regioes:
            candidatos_por_pagina.append([])
            continue

        # Agrupa regioes proximas em anuncios compostos
        grupos = _agrupar_regioes(regioes, tolerancia=14)

        # Filtra grupos muito pequenos (menos de ~4 cm2)
        grupos_validos = []
        for gx0, gy0, gx1, gy1, meta in grupos:
            w_cm = pontos_para_cm(gx1 - gx0)
            h_cm = pontos_para_cm(gy1 - gy0)
            if w_cm * h_cm < 3.0:
                continue
            grupos_validos.append((gx0, gy0, gx1, gy1, meta, w_cm, h_cm, pag_w, pag_h))

        candidatos_por_pagina.append((pg, grupos_validos))

    doc.close()

    # Monta lista plana de candidatos
    candidatos = []
    for entrada in candidatos_por_pagina:
        if not entrada:
            continue
        pg, grupos_validos = entrada
        for gx0, gy0, gx1, gy1, meta, w_cm, h_cm, pag_w, pag_h in grupos_validos:
            candidatos.append({
                "pagina": pg,
                "fonte":  meta.get("fonte", "misto"),
                "imagem": None,
                "pag_w":  pag_w,
                "pag_h":  pag_h,
                "w_cm":   w_cm,
                "h_cm":   h_cm,
                "x_cm":   pontos_para_cm(gx0),
                "y_cm":   pontos_para_cm(gy0),
                "x0_pt":  gx0, "y0_pt": gy0,
                "x1_pt":  gx1, "y1_pt": gy1,
            })

    if not candidatos:
        cb_resultado(erro="Nao foi possivel identificar anuncios no PDF.")
        return

    total = len(candidatos)
    resultados = []

    for i, cand in enumerate(candidatos):
        # Extrai imagem embutida real se a regiao tem exatamente uma imagem
        imagem_real = None
        try:
            doc2 = fitz.open(pdf_path)
            pagina2 = doc2[cand["pagina"] - 1]
            for info_img in pagina2.get_images(full=True):
                xref = info_img[0]
                try:
                    rects = pagina2.get_image_rects(xref)
                    if not rects:
                        continue
                    r = rects[0]
                    # Verifica se esta dentro da regiao do grupo
                    if (r.x0 >= cand["x0_pt"] - 5 and r.y0 >= cand["y0_pt"] - 5 and
                            r.x1 <= cand["x1_pt"] + 5 and r.y1 <= cand["y1_pt"] + 5):
                        dados_brutos = doc2.extract_image(xref)
                        imagem_real = Image.open(
                            io.BytesIO(dados_brutos["image"])).convert("RGB")
                        break
                except Exception:
                    pass
            doc2.close()
        except Exception:
            pass

        thumb = None
        if imagem_real is not None:
            thumb = imagem_real.copy()
            thumb.thumbnail((120, 120), Image.LANCZOS)

        preview = gerar_preview_anuncio(
            pdf_path, cand["pagina"],
            cand["x_cm"], cand["y_cm"],
            cand["w_cm"], cand["h_cm"])

        if preview is None and imagem_real is not None:
            preview = imagem_real.copy()
            preview.thumbnail((300, 500), Image.LANCZOS)

        if thumb is None and preview is not None:
            thumb = preview.copy()
            thumb.thumbnail((120, 120), Image.LANCZOS)

        resultados.append({
            **cand,
            "numero":  i + 1,
            "thumb":   thumb,
            "preview": preview,
            "w_px":    imagem_real.width  if imagem_real else 0,
            "h_px":    imagem_real.height if imagem_real else 0,
            "autor_info": extrair_autor_de_regiao(
                pdf_path, cand["pagina"],
                cand["x0_pt"], cand["y0_pt"],
                cand["x1_pt"], cand["y1_pt"]),
        })
        cb_progresso(i + 1, total, cand["pagina"])

    cb_resultado(imagens=resultados, total=total, erro=None)


# ═══════════════════════════════════════════════════════════════════
#  EXTRAÇÃO DE AUTOR / COLABORADOR DE MATÉRIAS JORNALÍSTICAS
# ═══════════════════════════════════════════════════════════════════

import re as _re_autor
import unicodedata as _uc_autor

# ══════════════════════════════════════════════════════════════════════════════
#  MÓDULO DE EXTRAÇÃO DE AUTOR — v4 (arquitetura modular + híbrida)
#
#  Pipeline de 6 etapas:
#    1. Pré-processamento      → linhas limpas
#    2. Localização do título  → título + posição
#    3. Delimitação do corpo   → início e fim do corpo
#    4. Extração de candidatos → zonas "antes" E "depois" do corpo
#    5. Pontuação + NER        → heurísticas + spaCy
#    6. Normalização + saída   → resultado estruturado
#
#  Funções públicas:
#    extrair_autores_completo(texto) → dict com título + lista de autores
#    extrair_autor(texto)            → dict compat. com versão anterior
# ══════════════════════════════════════════════════════════════════════════════


# ──────────────────────────────────────────────────────────────────────────────
#  CONFIGURAÇÃO — pesos e regras centralizados e ajustáveis
# ──────────────────────────────────────────────────────────────────────────────

class AutorConfig:
    """Configuração centralizada do extrator de autores.

    Ajuste os atributos desta classe para calibrar o comportamento
    sem precisar alterar a lógica central.
    """

    # ── Pesos positivos ────────────────────────────────────────────────────
    SCORE_NOME_PROPRIO  = 4   # is_human_name(): 2-5 palavras, iniciais maiúsculas
    SCORE_CAPS          = 2   # linha toda em MAIÚSCULAS (byline clássico de jornal)
    SCORE_TITLE_CASE    = 3   # Title Case sem CAPS (Marta Rodrigues)
    SCORE_LINHA_CURTA   = 2   # 2–4 palavras → bylines são concisos
    SCORE_ISOLADA       = 2   # linha curta entre linhas longas (isolação visual)
    SCORE_PERTO_TITULO  = 4   # próximo ao título detectado (posição de byline)
    SCORE_NER_PESSOA    = 3   # NER reforça PERSON — bônus, nunca veto
    SCORE_REDACAO       = 5   # contém "DA REDAÇÃO"
    SCORE_ASTERISCO     = 3   # asterisco de byline editorial (* antes/depois)
    SCORE_APOS_TEXTO    = 2   # zona após o corpo (assinatura de artigo)
    SCORE_DESC_CARGO    = 3   # padrão "Nome\nCargo" → próxima linha é cargo

    # ── Penalidades (reduzem o score; não vetam sozinhas) ─────────────────
    PENALTY_LONGA           = -3   # linha > 50 chars → bylines são curtos
    PENALTY_SECAO_PROXIMA   = -2   # próxima linha é seção → provável dateline
    PENALTY_PONTUACAO_FINAL = -2   # termina com . ou , → frase normal

    # ── Threshold e janelas ────────────────────────────────────────────────
    THRESHOLD           = 4    # score mínimo para aceitar candidato
    JANELA_ANTES        = 10   # linhas antes do corpo a escanear
    JANELA_DEPOIS       = 8    # linhas após o corpo a escanear
    DISTANCIA_TITULO    = 5    # dist. máx. (linhas) do título para bônus

    # ── Comprimento máximo de uma linha de autor (chars) ──────────────
    MAX_LEN_AUTOR       = 80

    # ── Padrões de "DA REDAÇÃO" (sem acento, uppercase) ───────────────
    PADROES_REDACAO     = ["DA REDACAO", "DA REDAÇAO"]

    # ── Regex de colaboração ("COM Fulano e Ciclano") ─────────────────
    PADRAO_COLABORACAO  = r'\bCOM\b\s+(.+)'

    # ── Indicadores de cargo/função (usados em _is_description_line) ──
    INDICADORES_CARGO   = [
        "jornalista", "reporter", "repórter", "colunista",
        "articulista", "presidente", "diretor", "diretora",
        "vereador", "vereadora", "deputad", "senador", "senadora",
        "professor", "professora", "advogad", "médic", "medic",
        "escritor", "escritora", "especialista", "analista",
        "secretári", "secretari", "administrad", "economista",
        "sociólog", "sociolog", "historiador", "filósof", "filosof",
        "psicólog", "psicolog", "editor", "editora",
        "correspondente", "enviado especial", "colaborador",
        "colaboradora", "especial para", "é jornalista",
    ]


# Instância padrão (usada quando nenhuma config é passada)
_cfg_padrao = AutorConfig() 


# ──────────────────────────────────────────────────────────────────────────────
#  NER (spaCy) — carregamento lazy
# ──────────────────────────────────────────────────────────────────────────────

_nlp_ner = None
_nlp_ner_disponivel = None


def _carregar_ner():
    """Carrega o modelo spaCy (lazy, uma única vez)."""
    global _nlp_ner, _nlp_ner_disponivel
    if _nlp_ner_disponivel is not None:
        return _nlp_ner_disponivel
    try:
        import spacy
        for modelo in ("pt_core_news_lg", "pt_core_news_sm"):
            try:
                _nlp_ner = spacy.load(modelo)
                _nlp_ner_disponivel = True
                return True
            except OSError:
                continue
        _nlp_ner_disponivel = False
    except ImportError:
        _nlp_ner_disponivel = False
    return _nlp_ner_disponivel


def is_person_name_ner(text):
    """NER desativado — retorna None (sem dependência externa)."""
    return None


# ──────────────────────────────────────────────────────────────────────────────
#  LISTAS DE REJEIÇÃO
# ──────────────────────────────────────────────────────────────────────────────

_SECOES_PROIBIDAS = frozenset({
    "SALVADOR", "OPINIÃO", "OPINIAO", "BAHIA", "BRASIL", "MUNDO",
    "ESPORTES", "ESPORTE", "ECONOMIA", "CULTURA", "POLITICA", "POLÍTICA",
    "CLASSIFICADOS", "CADERNO", "CIDADES", "EDITORIAL", "OBITUÁRIO",
    "OBITUARIO", "POLICIA", "POLÍCIA", "SAÚDE", "SAUDE", "EDUCAÇÃO",
    "EDUCACAO", "TODOS", "FOTO", "FOTOS", "INFOGRÁFICO", "INFOGRAFICO",
    "PÁGINA", "PAGINA", "PUBLICIDADE", "ANÚNCIO", "ANUNCIO",
    "COTIDIANO", "LAZER", "VARIEDADES", "ESPECIAL", "NACIONAL","SALVADOR",
    "INTERNACIONAL", "TECNOLOGIA", "CIÊNCIA", "CIENCIA", "MEIO AMBIENTE",
})

_BLACKLIST_NOMES = frozenset({
    "CLUBE", "BAHIA", "ENSINO", "SUPERIOR", "NÚCLEO", "NUCLEO",
    "IMPRESSOS", "DIGITAL", "SECRETARIA", "PREFEITURA", "UNIVERSIDADE",
    "FACULDADE", "HOSPITAL", "CENTRO", "PROGRAMA", "PROJETO",
    "GOVERNO", "MINISTÉRIO", "MINISTERIO", "CONSELHO", "EMPRESA",
    "AGÊNCIA", "AGENCIA", "ASSOCIAÇÃO", "ASSOCIACAO",
    "ESPORTE", "ESPORTES", "SISTEMA", "PLANO", "INSTITUTO",
    "FUNDAÇÃO", "FUNDACAO", "COMISSÃO", "COMISSAO", "TRIBUNAL",
    "CÂMARA", "CAMARA", "ASSEMBLEIA", "ASSEMBLEIA", "FEDERAÇÃO",
    "FEDERACAO", "SINDICATO", "COOPERATIVA", "COMPANHIA",
    "DEPARTAMENTO", "DIRETORIA", "GERÊNCIA", "GERENCIA",
    "COORDENAÇÃO", "COORDENACAO", "SUPERINTENDÊNCIA", "SUPERINTENDENCIA",
    "EDITORIA", "COLUNA", "SEÇÃO", "SECAO", "SUPLEMENTO",
    "CADERNO", "JORNAL", "REVISTA", "GAZETA", "CORREIO",
    "GRUPO", "REDE", "ORGANIZAÇÃO", "ORGANIZACAO",
    "LTDA", "S/A", "EIRELI", "MEI", "CNPJ",
})

_PALAVRAS_NAO_NOME = frozenset({
    "CLUBE", "ENSINO", "SUPERIOR", "DIGITAL", "SISTEMA", "PLANO",
    "ESPORTE", "RESULTADO", "RELATÓRIO", "RELATORIO", "PROCESSO",
    "SERVIÇO", "SERVICO", "PRODUTO", "MERCADO", "VENDA", "COMPRA",
    "OFERTA", "DEMANDA", "VALOR", "PREÇO", "PRECO", "CUSTO",
    "TAXA", "ÍNDICE", "INDICE", "NÚMERO", "NUMERO", "TOTAL",
    "PARCIAL", "FINAL", "INICIAL", "GERAL", "LOCAL", "REGIONAL",
    "ESTADUAL", "FEDERAL", "MUNICIPAL", "PÚBLICO", "PUBLICO",
    "PRIVADO", "SOCIAL", "CIVIL", "MILITAR", "URBANO", "RURAL",
    "NOVO", "NOVA", "ANTIGO", "ANTIGA", "GRANDE", "PEQUENO",
    "ALTO", "BAIXO", "BOM", "MAU", "PRIMEIRO", "SEGUNDO",
    "TERCEIRO", "QUARTO", "ÚLTIMO", "ULTIMO",
})

_PREPOSICOES_NOME = frozenset({"DE", "DA", "DO", "DOS", "DAS", "E"})

_VERBOS_COMUNS = frozenset({
    "É", "FOI", "SÃO", "SAO", "TEM", "VAI", "VEM", "FAZ", "DIZ",
    "DEU", "DEZ", "FEZ", "VER", "SER", "TER", "PODE", "DEVE",
    "QUER", "SERÁ", "SERA", "ESTÁ", "ESTA", "ESTÃO", "ESTAO",
    "SOBRE", "CONTRA", "PARA", "COMO", "MAIS", "APÓS", "APOS",
    "ANTES", "ENTRE", "AINDA", "TAMBÉM", "TAMBEM", "ONDE",
    "PORQUE", "QUANDO", "QUAL", "QUAIS",
})


def _remover_acentos(texto):
    nfkd = _uc_autor.normalize("NFD", texto)
    return "".join(c for c in nfkd if _uc_autor.category(c) != "Mn")


def _palavra_so_letras(palavra):
    """Retorna True se a palavra contém apenas letras (incluindo acentuadas)."""
    return all(c.isalpha() for c in palavra)


def _normalizar_nome(nome):
    """Normaliza capitalização e extrai primeiro + último nome quando há mais de 2.

    Regras:
    - JACKSON SOUZA          → Jackson Souza
    - João Carlos da Silva   → João Silva  (primeiro + último; partículas descartadas)
    - Marta Rodrigues        → Marta Rodrigues  (inalterado)
    - Jackson de Souza       → Jackson Souza  (preposição descartada)
    """
    nome = nome.strip().strip("*").strip()
    if not nome:
        return nome

    # Passo 1: normaliza capitalização
    partes = []
    for p in nome.split():
        if p.upper() in _PREPOSICOES_NOME:
            # mantém minúsculo (da / de / do)
            partes.append(p.lower())
        elif p == p.upper():
            # tudo maiúsculo → capitaliza
            partes.append(p.capitalize())
        else:
            partes.append(p)

    # Passo 2: se mais de 2 palavras, reduz a Primeiro + Último (ignora partículas)
    _PREP = {"de", "da", "do", "dos", "das", "e"}
    significativas = [p for p in partes if p.lower() not in _PREP]
    if len(significativas) > 2:
        partes = [significativas[0], significativas[-1]]
    elif len(significativas) == 2 and len(partes) > 2:
        # Ex: "Jackson de Souza" → ["Jackson", "de", "Souza"] → drop "de"
        partes = significativas

    return " ".join(partes)


# Regex determinista de nome de pessoa: cada palavra começa com maiúscula +
# minúsculas, separadas por espaço.  Aceita até 4 palavras (inclui partícula).
# Ex: "Jackson Souza" ✔  "João da Silva" ✔  "JACKSON SOUZA" ✘  "Economia" ✘
_RE_NOME_PROPRIO = _re_autor.compile(
    r'^'
    r'[A-ZÀ-ÖØ-Þ][a-zà-öø-ÿ]+'           # primeira palavra: Mai+min
    r'(?:'                                  # grupo repetível (1-4 vezes):
        r'\s(?:de|da|do|dos|das|e)'         #   partícula sem capitalização
        r'|'
        r'\s[A-ZÀ-ÖØ-Þ][a-zà-öø-ÿ]+'      #   palavra com Mai+min
    r'){1,4}'                               # até 4 partes adicionais (suporta até 5 tokens)
    r'$'
)


def is_valid_person_name(text: str) -> bool:
    """Validação determinista de nome de pessoa. Sem IA, apenas regex + regras.

    Critérios obrigatórios (todos devem passar):
    1. 2–4 palavras totais (preposicões contam)
    2. Cada palavra significativa: inicial maiúscula + restante minúsculo
    3. Sem dígitos, sem pontuação fora do nome, sem @
    4. Nenhuma palavra na blocklist institucional ou de seção
    5. Nenhum verbo comum
    6. Linha não é inteiramente em CAPS (não-normalizada) → já deve ter passado
       por _normalizar_nome antes de ser testada
    """
    if not text:
        return False
    t = text.strip().strip("*").strip()
    if not t or len(t) > 80:
        return False
    # Sem dígitos
    if any(c.isdigit() for c in t):
        return False
    # Sem pontuação forte ou arroba (endereços de e-mail não são nomes)
    if _re_autor.search(r'[.:;!?()\[\]{}"\'/\\@#$%&+=<>,]', t):
        return False
    palavras = t.split()
    if len(palavras) < 2 or len(palavras) > 4:
        return False
    # Valida estrutura com a regex principal
    if not _RE_NOME_PROPRIO.match(t):
        return False
    # Blocklist: nenhuma palavra pode ser institucional ou de seção
    t_upper = _remover_acentos(t.upper())
    for bl in _BLACKLIST_NOMES:
        if _re_autor.search(r'\b' + _re_autor.escape(bl) + r'\b', t_upper):
            return False
    if t_upper in _SECOES_PROIBIDAS:
        return False
    for p in palavras:
        pu = _remover_acentos(p.upper())
        if pu in _PALAVRAS_NAO_NOME:
            return False
        if pu in _VERBOS_COMUNS:
            return False
    # Não pode ser "DA REDAÇÃO"
    if "REDACAO" in _remover_acentos(t.upper()):
        return False
    return True


# Mantém o nome antigo como alias para não quebrar chamadas externas
is_human_name = is_valid_person_name


# ──────────────────────────────────────────────────────────────────────────────
#  FUNÇÕES AUXILIARES
# ──────────────────────────────────────────────────────────────────────────────

def _linha_tem_data(linha):
    """Detecta se a linha contém uma data."""
    return bool(_re_autor.search(
        r'\b\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}\b|\b\d{1,2}\s+de\s+\w+\b',
        linha, _re_autor.IGNORECASE))


def _linha_tem_numeros_precos(linha):
    """Detecta se a linha contém números, preços ou localizações."""
    return bool(
        _re_autor.search(r'R\$\s*[\d.,]+', linha) or
        _re_autor.search(r'\b\d{3,}\b', linha) or
        _re_autor.search(r'\b\d+[.,]\d+\b', linha)
    )


def _is_description_line(linha, config=None):
    """Linha parece descrição de cargo ou função."""
    cfg = config or _cfg_padrao
    linha = linha.strip()
    if not linha or len(linha.split()) < 2 or len(linha) > 120:
        return False
    if linha == linha.upper() and len(linha.split()) <= 5:
        return False
    return any(ind in linha.lower() for ind in cfg.INDICADORES_CARGO)


def _extrair_colaboradores(linha, config=None):
    """Extrai nomes de colaboradores de 'DA REDAÇÃO, COM NOME1 E NOME2'."""
    cfg = config or _cfg_padrao
    colaboradores = []
    match = _re_autor.search(cfg.PADRAO_COLABORACAO, linha, _re_autor.IGNORECASE)
    if not match:
        return colaboradores
    trecho = match.group(1).strip().rstrip(".")
    for parte in _re_autor.split(r'\s+E\s+|,\s*', trecho):
        parte = parte.strip()
        if not parte:
            continue
        # Normaliza CAPS para TitleCase antes de validar
        if parte == parte.upper():
            parte = parte.title()
        if is_human_name(parte):
            colaboradores.append(_normalizar_nome(parte))
    return colaboradores


def _dividir_multiplos_nomes(linha):
    """Tenta dividir uma linha com múltiplos nomes em partes individuais.

    Ex: "PAULO LEANDRO E MIRIAM HERMES" → ["PAULO LEANDRO", "MIRIAM HERMES"]
    Só divide quando TODOS os fragmentos passam em is_human_name(), evitando
    quebrar linhas de corpo de texto comuns.
    Retorna lista com a linha original se não conseguir dividir com segurança.
    """
    limpa = linha.strip()
    # Divide por " E " (maiúsculo/minúsculo) ou ", "
    partes = _re_autor.split(r'\s+[Ee]\s+|,\s+', limpa)
    if len(partes) <= 1:
        return [limpa]
    partes = [p.strip() for p in partes if p.strip()]
    # Só aceita a divisão se cada fragmento parecer um nome humano
    if len(partes) >= 2 and all(is_human_name(p) for p in partes):
        return partes
    return [limpa]


# ──────────────────────────────────────────────────────────────────────────────
#  ETAPA 1 — PRÉ-PROCESSAMENTO
# ──────────────────────────────────────────────────────────────────────────────

def _preprocessar_linhas(texto):
    """Normaliza e divide o texto em linhas não-vazias."""
    linhas = []
    for l in texto.splitlines():
        l_strip = " ".join(l.split())
        if l_strip:
            linhas.append(l_strip)
    return linhas


# ──────────────────────────────────────────────────────────────────────────────
#  ETAPA 2 — LOCALIZAÇÃO DO TÍTULO
# ──────────────────────────────────────────────────────────────────────────────

def _extrair_titulo(linhas, max_pos=8):
    """Detecta o título da matéria e sua posição.

    Título típico: primeira linha proeminente com 2–15 palavras antes do corpo.
    """
    for i, l in enumerate(linhas[:max_pos]):
        l_strip = l.strip()
        palavras = l_strip.split()
        # Títulos têm pelo menos 3 palavras (nomes próprios têm 2 → ignorar)
        if len(palavras) < 3 or len(palavras) > 15:
            continue
        if len(l_strip) < 8 or len(l_strip) > 150:
            continue
        if not l_strip[0].isupper():
            continue
        # Seção de 1–3 palavras todo maiúsculo → não é título
        if l_strip == l_strip.upper() and len(palavras) <= 3:
            continue
        # Não é linha de cargo/função
        if _is_description_line(l_strip):
            continue
        # Não é DA REDAÇÃO
        if any(p in _remover_acentos(l_strip.upper()) for p in ["DA REDACAO", "REDACAO"]):
            continue
        if _linha_tem_data(l_strip):
            continue
        return l_strip, i
    return None, -1


# ──────────────────────────────────────────────────────────────────────────────
#  ETAPA 3 — DELIMITAÇÃO DO CORPO
# ──────────────────────────────────────────────────────────────────────────────

def _encontrar_inicio_corpo(linhas):
    """Índice da primeira linha de parágrafo do corpo."""
    for i, linha in enumerate(linhas):
        l = linha.strip()
        if not l:
            continue
        if len(l) > 60 or (len(l) > 20 and l.endswith(".")):
            return i
    return len(linhas)


def _encontrar_fim_corpo(linhas, inicio_corpo):
    """Índice logo após a última linha longa do corpo."""
    for i in range(len(linhas) - 1, inicio_corpo - 1, -1):
        if len(linhas[i].strip()) > 40:
            return i + 1
    return len(linhas)


# ──────────────────────────────────────────────────────────────────────────────
#  ETAPA 4 — EXTRAÇÃO DE CANDIDATOS (zonas antes e depois)
# ──────────────────────────────────────────────────────────────────────────────

def _candidatos_antes(linhas, inicio_corpo, config):
    ini = max(0, inicio_corpo - config.JANELA_ANTES)
    return [{"linha": linhas[i], "indice": i, "zona": "antes"}
            for i in range(ini, inicio_corpo)]


def _candidatos_depois(linhas, fim_corpo, config):
    fim = min(len(linhas), fim_corpo + config.JANELA_DEPOIS)
    return [{"linha": linhas[i], "indice": i, "zona": "depois"}
            for i in range(fim_corpo, fim)]


# ──────────────────────────────────────────────────────────────────────────────
#  ETAPA 5 — PONTUAÇÃO DE CANDIDATOS
# ──────────────────────────────────────────────────────────────────────────────

def _classificar_candidato(cand, linhas, titulo_idx, config):
    """Classifica se uma linha candidata é autor com score balanceado.

    Arquitetura em duas fases:

    FASE A — Desqualificação imediata (-999)
        Rejeita casos que nunca poderão ser nomes de autores, independente
        de qualquer contexto. Sem exceções.

    FASE B — Pontuação balanceada
        Calcula score somando sinais POSITIVOS e PENALIDADES.
        Aceito apenas se score >= config.THRESHOLD.
        Isso permite que um nome em contexto fraco ainda seja aceito
        se tiver múltiplos sinais positivos, e rejeita linhas suspeitas
        mesmo que passem na Fase A.

    Retorna (score: int, motivos: list[str]).
    score == -999 → descartado na Fase A.
    """
    linha = cand["linha"]
    idx   = cand["indice"]
    zona  = cand["zona"]

    linha_s       = linha.strip()
    has_asterisco = linha_s.startswith("*") or linha_s.endswith("*")
    limpa         = linha_s.strip("*").strip()
    limpa_upper   = _remover_acentos(limpa.upper())
    palavras      = limpa.split()
    n_palavras    = len(palavras)

    # Contexto: linhas vizinhas (string vazia se inexistente)
    ant  = linhas[idx - 1].strip() if idx > 0 else ""
    prox = linhas[idx + 1].strip() if idx + 1 < len(linhas) else ""

    # ════════════════════════════════════════════════════════════════════
    # FASE A — DESQUALIFICAÇÃO IMEDIATA
    # Estes casos são estruturalmente impossíveis de ser nomes de autores.
    # ════════════════════════════════════════════════════════════════════

    if not limpa:
        return -999, ["vazio"]

    # Linha absurdamente longa → nenhum byline ultrapassa MAX_LEN_AUTOR chars
    if len(limpa) > config.MAX_LEN_AUTOR:
        return -999, ["muito_longa"]

    # Seção jornalística / dateline de cidade (match exato)
    if limpa_upper in _SECOES_PROIBIDAS:
        return -999, ["secao_proibida"]

    # Linha que **começa** com seção/cidade mas tem sufixo (ex: "SALVADOR —",
    # "BAHIA, 15/03"). Só rejeita se não for um nome humano válido
    # (evita rejeitar "Salvador Nascimento" ou "Bahia Santos").
    primeiro_upper = _remover_acentos(palavras[0].upper()) if palavras else ""
    if primeiro_upper in _SECOES_PROIBIDAS and not is_human_name(limpa):
        return -999, [f"inicio_secao:{primeiro_upper}"]

    # Palavra institucional (busca por palavra inteira com \b)
    for bl in _BLACKLIST_NOMES:
        if _re_autor.search(r'\b' + _re_autor.escape(bl) + r'\b', limpa_upper):
            return -999, [f"blacklist:{bl}"]

    # A linha É o próprio título → o título não é o autor
    if idx == titulo_idx:
        return -999, ["eh_titulo"]

    # Contém data → é dateline, não byline
    if _linha_tem_data(linha):
        return -999, ["data"]

    # Contém números ou preços → não é nome
    if _linha_tem_numeros_precos(linha):
        return -999, ["numeros_precos"]

    # A linha EM SI é um cargo/função (ex: "Vereadora de Salvador").
    # O cargo aparece ABAIXO do nome; a linha de cargo nunca é o autor.
    if _is_description_line(linha, config):
        return -999, ["eh_cargo"]

    # ════════════════════════════════════════════════════════════════════
    # FASE B — PONTUAÇÃO BALANCEADA
    # Sinais positivos aumentam a confiança; penalidades diminuem.
    # Candidato aceito somente se score final >= config.THRESHOLD.
    # ════════════════════════════════════════════════════════════════════
    score  = 0
    motivos = []

    # ── Sinais positivos ───────────────────────────────────────────────

    # [+4] Heurística determinista de nome de pessoa (regex + blocklist)
    nome_valido = is_valid_person_name(_normalizar_nome(limpa))
    if nome_valido:
        score += config.SCORE_NOME_PROPRIO
        motivos.append(f"+{config.SCORE_NOME_PROPRIO}:nome_proprio")

    # [+2] Todo em CAPS e composto só de letras (byline clássico de jornal)
    is_upper = (limpa == limpa.upper() and n_palavras >= 2
                and all(p.isalpha() for p in palavras))
    if is_upper:
        score += config.SCORE_CAPS
        motivos.append(f"+{config.SCORE_CAPS}:CAPS")

    # [+3] Title Case: iniciais maiúsculas mas não tudo caps (Marta Rodrigues)
    if (not is_upper) and nome_valido:
        score += config.SCORE_TITLE_CASE
        motivos.append(f"+{config.SCORE_TITLE_CASE}:title_case")

    # [+2] Linha curta: bylines têm tipicamente 2–4 palavras
    if 2 <= n_palavras <= 4:
        score += config.SCORE_LINHA_CURTA
        motivos.append(f"+{config.SCORE_LINHA_CURTA}:linha_curta")

    # [+4] Próximo ao título detectado (posição característica de byline)
    if titulo_idx >= 0 and abs(idx - titulo_idx) <= config.DISTANCIA_TITULO:
        score += config.SCORE_PERTO_TITULO
        motivos.append(f"+{config.SCORE_PERTO_TITULO}:perto_titulo")

    # [+2] Linha isolada entre linhas longas (padrão visual de byline)
    if len(limpa) < 50 and (len(ant) > 55 or len(prox) > 55):
        score += config.SCORE_ISOLADA
        motivos.append(f"+{config.SCORE_ISOLADA}:isolada")

    # [+3] Padrão "Nome\nCargo": linha seguinte é descrição de cargo
    if _is_description_line(prox, config):
        score += config.SCORE_DESC_CARGO
        motivos.append(f"+{config.SCORE_DESC_CARGO}:seguido_de_cargo")

    # [+2] Assinatura no final do texto (artigos de opinião)
    if zona == "depois":
        score += config.SCORE_APOS_TEXTO
        motivos.append(f"+{config.SCORE_APOS_TEXTO}:zona_depois")

    # [+3] Asterisco editorial (* antes ou depois) — marcador explícito de byline
    if has_asterisco:
        score += config.SCORE_ASTERISCO
        motivos.append(f"+{config.SCORE_ASTERISCO}:asterisco")

    # NER removido: sistema é 100% determinístico

    # ── Penalidades ────────────────────────────────────────────────────

    # [-3] Linha longa para um byline (50–80 chars) → provavelmente frase
    if len(limpa) > 50:
        score += config.PENALTY_LONGA
        motivos.append(f"{config.PENALTY_LONGA}:linha_longa")

    # [-2] Próxima linha é seção proibida → atual provavelmente é dateline
    if _remover_acentos(prox.upper()) in _SECOES_PROIBIDAS:
        score += config.PENALTY_SECAO_PROXIMA
        motivos.append(f"{config.PENALTY_SECAO_PROXIMA}:precede_secao")

    # [-2] Termina com ponto ou vírgula → frase normal, não byline
    if limpa.endswith(".") or limpa.endswith(","):
        score += config.PENALTY_PONTUACAO_FINAL
        motivos.append(f"{config.PENALTY_PONTUACAO_FINAL}:termina_pontuacao")

    return score, motivos


# ──────────────────────────────────────────────────────────────────────────────
#  FUNÇÃO PRINCIPAL — extrair_autores_completo
# ──────────────────────────────────────────────────────────────────────────────

def extrair_autores_completo(texto, config=None, debug=False):
    """Extrai autores de um bloco de texto jornalístico.

    Escaneia antes E depois do corpo para cobrir bylines e assinaturas.

    Retorna dict:
        titulo       : str | None      — título detectado
        autores      : list[str]       — nomes normalizados de todos os autores
        tipo         : str             — tipo dominante (reporter/redacao/articulista)
        colaboradores: list[str]       — nomes de colaboradores (COM ...)
        confianca    : float           — 0.0 – 1.0
        raw_linhas   : list[str]       — linhas originais detectadas
        log          : list[str]       — log de decisões (se debug=True)
    """
    cfg = config or _cfg_padrao
    log = []

    def _log(msg):
        if debug:
            log.append(msg)

    resultado_vazio = {
        "titulo": None, "autores": [], "tipo": "desconhecido",
        "colaboradores": [], "confianca": 0.0, "raw_linhas": [], "log": log,
    }

    if not texto or not texto.strip():
        return resultado_vazio

    # ── Etapa 1: Pré-processamento ───────────────────────────────
    linhas = _preprocessar_linhas(texto)
    if not linhas:
        return resultado_vazio
    _log(f"[PRE] {len(linhas)} linhas")

    # ── Etapa 2: Localização do título ───────────────────────────
    titulo, titulo_idx = _extrair_titulo(linhas)
    _log(f"[TITULO] '{titulo}' @ idx={titulo_idx}")

    # ── Etapa 3: Delimitação do corpo ────────────────────────────
    inicio_corpo = _encontrar_inicio_corpo(linhas)
    fim_corpo    = _encontrar_fim_corpo(linhas, inicio_corpo)
    _log(f"[CORPO] início={inicio_corpo}  fim={fim_corpo}")

    # ── Detecção imediata de "DA REDAÇÃO" nas zonas candidatas ───
    zonas_check = list(range(max(0, inicio_corpo - cfg.JANELA_ANTES), inicio_corpo))
    zonas_check += list(range(fim_corpo, min(len(linhas), fim_corpo + cfg.JANELA_DEPOIS)))
    for i in zonas_check:
        linha = linhas[i]
        lu = _remover_acentos(linha.upper())
        for padrao in cfg.PADROES_REDACAO:
            if padrao in lu:
                colaboradores = _extrair_colaboradores(linha, cfg)
                _log(f"[REDACAO] '{linha}'  colab={colaboradores}")
                # Inclui colaboradores em autores:
                # "DA REDAÇÃO, COM PAULO E MIRIAM" → autores=[Redação, Paulo, Miriam]
                autores_redacao = ["Redação"] + colaboradores
                return {
                    "titulo": titulo,
                    "autores": autores_redacao,
                    "tipo": "redacao",
                    "colaboradores": colaboradores,
                    "confianca": 0.95 if colaboradores else 0.90,
                    "raw_linhas": [linha],
                    "log": log,
                }

    # ── Etapa 4: Candidatos nas duas zonas ───────────────────────
    todos_cands = (
        _candidatos_antes(linhas, inicio_corpo, cfg) +
        _candidatos_depois(linhas, fim_corpo, cfg)
    )
    _log(f"[CANDS] total={len(todos_cands)}")

    # ── Etapa 5: Classificação (score balanceado) ────────────────────────
    scored = []
    for cand in todos_cands:
        score, motivos = _classificar_candidato(cand, linhas, titulo_idx, cfg)
        if score == -999:
            _log(f"[FASE-A] rejeitado '{cand['linha']}'  ← {motivos[0] if motivos else '?'}")
            continue
        _log(f"[SCORE={score:+d}] '{cand['linha']}'  → {motivos}")
        scored.append((score, cand, motivos))

    scored.sort(key=lambda x: x[0], reverse=True)
    aceitos = [(s, c, m) for s, c, m in scored if s >= cfg.THRESHOLD]

    if not aceitos:
        _log(f"[RESULTADO] nenhum candidato alcançou threshold={cfg.THRESHOLD}")
        if debug:
            for s, c, m in scored[:3]:
                _log(f"  top_descartado score={s:+d}: '{c['linha']}'  {m}")
        return resultado_vazio

    # ── Etapa 6: Montar resultado (deduplicar por proximidade) ───
    autores_aceitos = []
    ultimos_idx = []
    for score, cand, motivos in aceitos:
        if any(abs(cand["indice"] - ui) < 2 for ui in ultimos_idx):
            continue
        autores_aceitos.append((score, cand))
        ultimos_idx.append(cand["indice"])

    autores_nomes  = []
    raw_linhas     = []
    tipos          = []
    confiancas     = []

    for score, cand in autores_aceitos:
        linha_s_c = cand["linha"].strip()
        tem_ast   = linha_s_c.startswith("*") or linha_s_c.endswith("*")
        limpa_raw = linha_s_c.strip("*").strip()
        prox      = linhas[cand["indice"] + 1] if cand["indice"] + 1 < len(linhas) else ""

        # Determina tipo do autor para este candidato
        if tem_ast:
            tipo_cand = "reporter"      # asterisco = byline editorial explícito
        elif cand["zona"] == "depois" or _is_description_line(prox, cfg):
            tipo_cand = "articulista"   # assinatura ao final ou seguida de cargo
        else:
            tipo_cand = "reporter"

        confianca_cand = min(0.95, 0.45 + score * 0.05)

        # Separa múltiplos nomes em uma linha: "FULANO E BELTRANO" → nomes individuais
        # Só divide quando ambas as partes passam em is_human_name() — sem falsos positivos
        partes_nome = _dividir_multiplos_nomes(limpa_raw)
        for parte in partes_nome:
            autores_nomes.append(_normalizar_nome(parte))
            raw_linhas.append(cand["linha"])
            tipos.append(tipo_cand)
            confiancas.append(confianca_cand)

    # ── Deduplicação final por nome normalizado ───────────────────────
    # Remove repetições causadas por candidatos de zonas sobrepostas
    vistos        = set()
    autores_nomes_d  = []
    raw_linhas_d     = []
    tipos_d          = []
    confiancas_d     = []
    for n, r, t, c in zip(autores_nomes, raw_linhas, tipos, confiancas):
        chave = _remover_acentos(n.upper())
        if chave not in vistos:
            vistos.add(chave)
            autores_nomes_d.append(n)
            raw_linhas_d.append(r)
            tipos_d.append(t)
            confiancas_d.append(c)
    autores_nomes = autores_nomes_d
    raw_linhas    = raw_linhas_d
    tipos         = tipos_d
    confiancas    = confiancas_d

    tipo_dominante = max(set(tipos), key=tipos.count) if tipos else "desconhecido"
    confianca_max  = max(confiancas) if confiancas else 0.0

    _log(f"[FINAL] autores={autores_nomes}  tipo={tipo_dominante}  confianca={confianca_max:.2f}")

    return {
        "titulo":        titulo,
        "autores":       autores_nomes,
        "tipo":          tipo_dominante,
        "colaboradores": [],
        "confianca":     confianca_max,
        "raw_linhas":    raw_linhas,
        "log":           log,
    }


# ──────────────────────────────────────────────────────────────────────────────
#  WRAPPER BACKWARD-COMPATIBLE — extrair_autor
# ──────────────────────────────────────────────────────────────────────────────

def extrair_autor(texto, debug=False):
    """Wrapper retrocompatível que chama extrair_autores_completo().

    Retorna dict:
        autor:            str | None
        tipo:             'reporter' | 'redacao' | 'articulista' | 'desconhecido'
        colaboradores:    list[str]
        raw_autor_linha:  str
        confianca:        float
        log:              list[str]
    """
    r = extrair_autores_completo(texto, debug=debug)
    return {
        "autor":           r["autores"][0] if r["autores"] else None,
        "tipo":            r["tipo"],
        "colaboradores":   r["colaboradores"],
        "raw_autor_linha": r["raw_linhas"][0] if r["raw_linhas"] else "",
        "confianca":       r["confianca"],
        "log":             r["log"],
    }


def extrair_autor_de_regiao(pdf_path, num_pagina, x0_pt, y0_pt, x1_pt, y1_pt):
    """Extrai texto de uma região do PDF e tenta identificar o autor."""
    try:
        doc = fitz.open(pdf_path)
        pagina = doc[num_pagina - 1]
        rect = fitz.Rect(x0_pt, y0_pt, x1_pt, y1_pt)
        texto = pagina.get_text("text", clip=rect)
        doc.close()
        return extrair_autor(texto)
    except Exception:
        return extrair_autor("")


def extrair_todos_autores_pdf(pdf_path, cb_progresso, cb_resultado):
    """Varre TODAS as páginas do PDF, identifica matérias/regiões e extrai autor de cada uma.
    Retorna lista de dicts com info do autor + coordenadas."""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        cb_resultado(erro=str(e))
        return

    num_paginas = len(doc)
    todas_materias = []

    for num_pagina in range(num_paginas):
        pg = num_pagina + 1
        pagina = doc[num_pagina]
        pag_w, pag_h = pegar_tamanho_pagina_cm(pagina)

        regioes = []

        # 1) Blocos de texto
        for bloco in pagina.get_text("blocks"):
            x0, y0, x1, y1, texto = bloco[0], bloco[1], bloco[2], bloco[3], bloco[4]
            texto = texto.strip()
            if not texto or len(texto) < 10:
                continue
            w_pt = x1 - x0; h_pt = y1 - y0
            if w_pt < 20 or h_pt < 8:
                continue
            if len(texto) <= 4 and texto.isdigit():
                continue
            regioes.append((x0, y0, x1, y1, {"fonte": "texto", "texto": texto}))

        # 2) Imagens embutidas
        for info_img in pagina.get_images(full=True):
            xref = info_img[0]
            try:
                rects = pagina.get_image_rects(xref)
                if not rects:
                    continue
                r = rects[0]
                if r.width < 10 or r.height < 10:
                    continue
                regioes.append((r.x0, r.y0, r.x1, r.y1,
                                {"fonte": "embutida", "xref": xref}))
            except Exception:
                pass

        if not regioes:
            cb_progresso(pg, num_paginas, pg)
            continue

        # Agrupa regiões próximas
        grupos = _agrupar_regioes(regioes, tolerancia=14)

        # Filtra muito pequenos
        for gx0, gy0, gx1, gy1, meta in grupos:
            w_cm = pontos_para_cm(gx1 - gx0)
            h_cm = pontos_para_cm(gy1 - gy0)
            if w_cm * h_cm < 5.0:
                continue

            # Extrai texto da região
            rect = fitz.Rect(gx0, gy0, gx1, gy1)
            texto_regiao = pagina.get_text("text", clip=rect)
            if not texto_regiao or len(texto_regiao.strip()) < 15:
                continue

            info_autor = extrair_autor(texto_regiao)

            # Gera trecho representativo
            trecho = " ".join(texto_regiao.split())[:200]

            todas_materias.append({
                "pagina":  pg,
                "x_cm":    pontos_para_cm(gx0),
                "y_cm":    pontos_para_cm(gy0),
                "w_cm":    w_cm,
                "h_cm":    h_cm,
                "x0_pt":   gx0, "y0_pt": gy0,
                "x1_pt":   gx1, "y1_pt": gy1,
                "pag_w":   pag_w,
                "pag_h":   pag_h,
                "trecho":  trecho,
                "autor_info": info_autor,
            })

        cb_progresso(pg, num_paginas, pg)

    doc.close()

    if not todas_materias:
        cb_resultado(erro="Nenhuma matéria/autor encontrado no PDF.")
        return

    cb_resultado(materias=todas_materias, total=len(todas_materias), erro=None)


def normalizar(texto):
    # Remove acentos, coloca tudo minusculo, separa palavras coladas por CamelCase
    # e remove pontuacao — resolve o problema do OCR juntar palavras
    texto = _uc_autor.normalize("NFD", texto)
    texto = "".join(c for c in texto if _uc_autor.category(c) != "Mn")
    texto = texto.lower()
    # Separa palavras grudadas por letra maiuscula (ex: "MunicípiodeSanta" -> "municipio de santa")
    texto = _re_autor.sub(r"([a-z])([A-Z])", r"\1 \2", texto)
    # Separa letras de numeros colados (ex: "N002" -> "N 002")
    texto = _re_autor.sub(r"([a-zA-Z])([0-9])", r"\1 \2", texto)
    texto = _re_autor.sub(r"([0-9])([a-zA-Z])", r"\1 \2", texto)
    # Remove pontuacao e caracteres especiais
    texto = _re_autor.sub(r"[^a-z0-9 ]", " ", texto)
    # Colapsa espacos multiplos
    texto = " ".join(texto.split())
    return texto


def _remover_acentos_simples(texto: str) -> str:
    """Remove acentos para busca case-insensitive robusta."""
    import unicodedata
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def _search_exact_name(name: str, text: str) -> list:
    """Retorna lista de re.Match do nome exato em text.

    Normaliza ambos (sem acentos, minúsculas) e usa \\b em torno do
    nome completo — "Luiz Teles" só casa com "Luiz Teles", nunca com
    "Luiz" ou "Luis Teles" isolados.
    """
    import re as _re_exact
    name_norm = _remover_acentos_simples(name.strip()).lower()
    # Espaços internos do nome podem virar 1+ espaços ou quebras de linha no PDF
    partes = [_re_exact.escape(p) for p in name_norm.split()]
    if not partes:
        return []
    # Une as partes com \s+ para tolerar espaços múltiplos/quebras de linha
    padrao = _re_exact.compile(r"\b" + r"\s+".join(partes) + r"\b",
                                _re_exact.IGNORECASE)
    text_norm = _remover_acentos_simples(text).lower()
    return list(padrao.finditer(text_norm))


def buscar_autor_no_pdf(pdf_path: str, nome_busca: str) -> list:
    """Busca um nome de autor em todas as páginas do PDF.

    Retorna lista de dicts:
        [{pagina: int, trechos: [str]}, ...]
    onde cada trecho é um fragmento de ~120 chars ao redor da ocorrência.
    A busca é case-insensitive, ignora acentos e é EXATA:
    "Luiz Teles" só encontra "Luiz Teles", não "Luiz" isolado.
    """
    import fitz  # PyMuPDF

    if not nome_busca or not nome_busca.strip():
        return []

    resultados = []
    try:
        doc = fitz.open(pdf_path)
        for num_pag in range(len(doc)):
            pagina = doc[num_pag]
            texto_pag = pagina.get_text("text") or ""

            matches = _search_exact_name(nome_busca, texto_pag)
            if not matches:
                continue

            trechos = []
            CONTEXTO = 60  # chars antes e depois
            # Busca nos offsets do texto normalizado mas exibe texto original
            texto_norm = _remover_acentos_simples(texto_pag).lower()
            for m in matches:
                inicio = max(0, m.start() - CONTEXTO)
                fim    = min(len(texto_pag), m.end() + CONTEXTO)
                trecho = texto_pag[inicio:fim].replace("\n", " ").strip()
                trecho = " ".join(trecho.split())
                if trecho and trecho not in trechos:
                    trechos.append(trecho)

            if trechos:
                resultados.append({"pagina": num_pag + 1, "trechos": trechos})

        doc.close()
    except Exception as exc:
        return [{"pagina": 0, "trechos": [f"Erro ao abrir PDF: {exc}"]}]

    return resultados


def load_names_from_txt(file_path: str) -> list:
    """Lê um arquivo .txt e retorna lista de nomes (um por linha).
    Ignora linhas vazias e linhas que começam com '#'.
    """
    nomes = []
    try:
        with open(file_path, encoding="utf-8", errors="replace") as f:
            for linha in f:
                nome = linha.strip()
                if nome and not nome.startswith("#"):
                    nomes.append(nome)
    except Exception as exc:
        raise IOError(f"Não foi possível ler o arquivo: {exc}")
    return nomes


def search_multiple_names(names: list, pdf_path: str) -> dict:
    """Busca vários nomes no PDF e retorna um dict:
        {nome: [pagina1, pagina2, ...], ...}
    Páginas sem resultado ficam com lista vazia.
    """
    resultado = {}
    for nome in names:
        ocorrencias = buscar_autor_no_pdf(pdf_path, nome)
        # Extrai só os números de página, ordenados e sem repetição
        paginas = sorted({r["pagina"] for r in ocorrencias if r["pagina"] > 0})
        resultado[nome] = paginas
    return resultado


def extrair_autores_ataarde(texto: str) -> dict:
    """Extrai autores do jornal A TARDE a partir do texto completo do PDF.

    Padrões detectados (da análise real do PDF):
      1. NOME_CAPS*  — byline de jornalista com asterisco (sempre Jornalista)
      2. NOME_CAPS   — byline em ALLCAPS sozinho na linha (com filtros rígidos)
      3. NOME_CAPS email@grupoatarde.com.br — jornalista + email editorial
      4. Nome / Cargo  ou  Nome - Cargo  — cargo explícito
      5. Por: Nome  /  Texto: Nome  /  Reportagem: Nome
      6. Charge/Ilustração de Nome
      7. Email pessoal (@gmail etc.) com nome antes na mesma linha (leitores)
    """
    import re
    import unicodedata
    from collections import defaultdict

    if not texto or len(texto.strip()) < 100:
        return {"total_autores": 0, "autores_unicos": [], "por_categoria": {}}

    def _norm(s: str) -> str:
        return "".join(
            c for c in unicodedata.normalize("NFD", s)
            if unicodedata.category(c) != "Mn"
        ).lower()

    # ── Blacklist: palavras (normalizadas) que nunca fazem parte de nomes ──
    _BL = {
        # Esporte
        "copa", "liga", "esporte", "clube", "esportes", "basquete", "futebol",
        "placar", "giramundo", "telinha", "real", "madrid", "mundial",
        "brasileiro", "brasileira", "campeonato", "rodada", "feminino", "masculino",
        "tenis", "rugby", "volei", "natacao", "atletismo", "jogos", "olimpico",
        # Tempo / previsão
        "sol", "chuva", "nuvens", "vento", "forte", "tempo", "clima", "previsao",
        # Seções e temas jornalísticos
        "metas", "distintas", "distintos", "ensino", "superior", "habitacao",
        "crise", "politica", "investigacao", "homenagem", "tradicao",
        "cidadania", "seguranca", "economia", "saude", "educacao", "cultura",
        "tecnologia", "policia", "celebracao", "tragedia", "cotidiano",
        "nacional", "internacional", "regional", "especial",
        "noticias", "noticia", "novidade", "novidades",
        "assembleia", "extraordinaria", "extraordinario",
        "relacoes", "exteriores", "judiciario", "judiciaria",
        # Automóveis (seção Auto)
        "auto", "motor", "turbo", "autonomia", "autos", "automovel",
        # Institucional
        "prefeitura", "municipal", "governo", "estado", "federal",
        "ministerio", "secretaria", "tribunal", "congresso", "camara", "senado",
        "complexo", "hospitalar", "ufba", "uesc", "banco", "poder",
        # Vocabulário de jornal / mídia
        "imprensa", "fundador", "jornais", "verificador", "portal",
        "foto", "fotografia", "fatos", "causos", "fotos", "agencia", "presse",
        "poucas", "boas", "charge", "astrologia",
        # Cemitérios / morte
        "bosque", "jardim", "saudade", "paz", "campo", "santo", "cemiterio",
        "falecimento", "obito", "obituario",
        # Dias / tempo
        "domingo", "segunda", "terca", "quarta", "quinta", "sexta", "sabado",
        "hoje", "ontem", "amanha", "semana", "mes", "ano",
        # Lugares comuns (não são sobrenomes de jornalistas)
        "brasil", "bahia", "nordeste", "regiao", "cidade", "capital",
        "rio", "preto", "lapa", "salvador", "territorio", "municipio",
        "distrito", "estado", "pais", "mundo", "france", "sao",
        # Artigos / preposições sozinhos em CAPS
        "de", "da", "do", "dos", "das", "a", "o", "os", "as",
        "em", "no", "na", "nos", "nas", "e", "ou", "um", "uma",
        "for", "the", "and",
        # Outros falsos positivos observados no PDF
        "dia", "outras", "outros", "hora", "horas", "anos", "dias",
        "meses", "semanas", "digital", "eletronico", "eletronica",
        "publicas", "publicos", "novo", "nova", "grande", "pequeno",
        "presas", "setor", "hectares", "serrinha",
        "musica", "heranca", "eleicoes", "audiencia", "combustiveis",
        "turismo", "sustentavel", "extrema", "direita", "conflitos",
        "recentes", "exercicio", "financeiro", "energetica", "segura",
        # Frases / topicos adicionais detectados em ALLCAPS
        "artistas", "consagrados", "conjunta", "conjunto", "acao",
        "outro", "outra", "lado", "frente", "fundo", "nota",
        "byd", "song", "plus", "pro", "max",  # marcas em inglês / automóveis
        "norte", "sul","leste", "oeste",
        "cena", "video", "texto",
        "alerta", "perigo", "risco", "cuidado", "atencao",
        "mais", "menos", "cada", "toda", "todo", "todos", "todas",
        "nova", "velho", "velha", "melhor", "pior",
    }

    # Preposições/artigos em CAPS (nunca iniciam nem terminam um nome)
    _PREP_CAPS = {"DE", "DA", "DO", "DOS", "DAS", "EM", "NO", "NA",
                  "NOS", "NAS", "UM", "UMA", "FOR", "OU", "E", "A", "O",
                  "OS", "AS", "AO", "AOS"}

    # Palavras que aparecem em fragmentos de texto, nunca em nomes de pessoas (TC)
    _NOMES_NUNCA_TC = {
        "que", "se", "mas", "por", "para", "sobre", "entre", "com",
        "seus", "suas", "nosso", "nossos", "nossa", "nossas",
        "isso", "esta", "este", "esse", "essa", "aquela", "aquele",
        "apenas", "ainda", "tambem", "cada", "dentro", "fora",
        "antes", "depois", "sempre", "nunca", "muito", "pouco",
        "mais", "menos", "bem", "mal", "quando", "onde", "como",
        "recente", "recentes", "coletiva", "coletivo",
        "disponivel", "disponivel", "importante", "preocupante",
        "sustentavel", "digital", "shows", "turismo",
        "faz", "vai", "vao", "tem", "sao", "sao", "eram",
        "estao", "esta", "pelo", "pela", "pelos", "pelas",
        "era", "vice", "assim", "neste", "nesta", "nessa", "nesse",
    }

    # Terminações de verbos conjugados (nunca estão em nomes de pessoas)
    _RE_VERBO_ENDING = re.compile(
        r'(?:ou|aram|eram|iram|ando|endo|indo|arão|erão|irão'
        r'|ificar|izar|ecer|armos|ermos|irmos'
        r'|aram$|iram$|avam$|ivam$)$',
        re.IGNORECASE
    )

    # ── Regexes ───────────────────────────────────────────────────────────
    _RE_CAPS_ONLY = re.compile(
        r'^[A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-ZÁÉÍÓÚÂÊÔÃÕÇ\s\.\-]{2,60}$')

    _RE_ASTERISCO = re.compile(
        r'([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]{3,50})\s*\*')

    _RE_CARGO = re.compile(
        r'([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-záéíóúâêôãõçÁÉÍÓÚÂÊÔÃÕÇ\s]{3,40})'
        r'\s*(?:/|\||–|-)\s*'
        r'(Jornalista|Redator[a]?|Colunista|Editor[a]?|Rep[oó]rter|'
        r'Leitor[a]?|Deputad[ao]|Prefeito|Prefeita|Vereador[a]?|'
        r'Senador[a]?|Governador[a]?|Secret[aá]ri[oa]|Correspondente)',
        re.IGNORECASE)

    _RE_PREFIXO = re.compile(
        r'^\s*(?:por|texto de|texto|reportagem de|reportagem'
        r'|enviado por|enviado)\s*:?\s+'
        r'([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-záéíóúâêôãõçÁÉÍÓÚÂÊÔÃÕÇ\s]{3,40}?)'
        r'(?:\s*[\*,\.]|\s*$)',
        re.IGNORECASE)

    _RE_CHARGE = re.compile(
        r'\b(?:charge|ilustra[cç][aã]o|chargista)\b[^:]*?(?:de|por|:)\s*'
        r'([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-záéíóúâêôãõçÁÉÍÓÚÂÊÔÃÕÇ\s]{3,40?})'
        r'(?:\s*[\*,\.]|\s*$)',
        re.IGNORECASE)

    # Nome ALLCAPS + email editorial na mesma linha (ex: NÚBIA CRISTINA email@at...)
    _RE_CAPS_EMAIL = re.compile(
        r'^([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]{3,40})\s+'
        r'[\w\.\-]+@[\w\.\-]+\.[a-z]{2,}$')

    # Email pessoal com nome em TC antes (ex: João Silva joao@gmail.com)
    _RE_EMAIL_NOME_TC = re.compile(
        r'([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-záéíóúâêôãõçÁÉÍÓÚÂÊÔÃÕÇ]{2,20}'
        r'(?:\s+[A-Za-záéíóúâêôãõçÁÉÍÓÚÂÊÔÃÕÇ]{2,20}){1,3})'
        r'\s+[\w\.\-]+@[\w\.\-]+\.\w+')

    _CARGOS_POLITICO = re.compile(
        r'\b(deputad[ao]|prefeito|prefeita|vereador[a]?|senador[a]?'
        r'|governador[a]?|secretári[oa]|ministro|ministra)\b', re.IGNORECASE)

    _EMAIL_ANY = re.compile(r'@[\w\.\-]+\.\w+')

    # Cabeçalho de página do A TARDE (reset de seção)
    _RE_CABECALHO = re.compile(
        r'SALVADOR\s+(?:SEGUNDA|TERCA|QUARTA|QUINTA|SEXTA|SABADO|DOMINGO)'
        r'[\-\s]FEIRA', re.IGNORECASE)

    # ── Prioridade de categoria (nunca rebaixa uma classificação maior) ───
    _PRIO = {
        'Jornalistas / Redatores': 3,
        'Colunistas':              2,
        'Políticos / Opinião':     2,
        'Leitores':                1,
    }

    # ── Helpers ───────────────────────────────────────────────────────────
    def _caps_e_nome_pessoa(linha: str) -> bool:
        """True se a linha ALLCAPS parece nome de pessoa."""
        palavras = linha.split()
        if len(palavras) < 2 or len(palavras) > 4:
            return False
        for p in palavras:
            if not re.match(r'^[A-ZÁÉÍÓÚÂÊÔÃÕÇ\-]{2,20}$', p):
                return False
            if _norm(p) in _BL:
                return False
        if palavras[0] in _PREP_CAPS or palavras[-1] in _PREP_CAPS:
            return False
        return True

    def _nome_valido_tc(nome: str) -> bool:
        """True se 'nome' em Title Case parece um nome de pessoa real."""
        palavras = nome.split()
        if len(palavras) < 2 or len(palavras) > 4:
            return False
        for p in palavras:
            p_norm = _norm(p)
            if p_norm in _BL:
                return False
            if p_norm in _NOMES_NUNCA_TC:
                return False
            # Rejeita palavras com terminações típicas de verbo conjugado
            if _RE_VERBO_ENDING.search(p_norm):
                return False
            # Preposição/artigo sozinho
            if p.upper() in _PREP_CAPS:
                return False
            # Cada palavra deve começar com maiúscula e ter 2+ chars
            if len(p) < 2 or not p[0].isupper():
                return False
        # Rejeita se há 2+ palavras inteiramente em CAPS (já cobre ALLCAPS path)
        caps_words = sum(1 for p in palavras if p.isupper() and len(p) > 2)
        if caps_words >= 2:
            return False
        return True

    def _classificar(linha_lower: str) -> str:
        if _EMAIL_ANY.search(linha_lower) or any(
                x in linha_lower for x in ['@gmail', '@yahoo', '@hotmail']):
            return 'Leitores'
        if _CARGOS_POLITICO.search(linha_lower):
            return 'Políticos / Opinião'
        if any(x in linha_lower for x in ['colunista', 'articulista', 'cronista']):
            return 'Colunistas'
        return 'Jornalistas / Redatores'

    def _reg(nome: str, cat: str):
        nome = nome.strip().rstrip('*., ').title()
        if not _nome_valido_tc(nome):
            return
        # Nunca rebaixa categoria já registrada
        if nome not in autores or _PRIO.get(cat, 0) > _PRIO.get(autores[nome], 0):
            autores[nome] = cat

    # ── Varredura ─────────────────────────────────────────────────────────
    autores: dict[str, str] = {}
    linhas = texto.split('\n')
    secao_leitor = False

    for linha_raw in linhas:
        linha = linha_raw.strip()
        if len(linha) < 3:
            continue
        linha_lower = _norm(linha)

        # ── Reset de seção a cada cabeçalho de página ─────────────────────
        if _RE_CABECALHO.search(linha):
            secao_leitor = False
            continue

        # ── Detecta entrada na seção "Espaço do Leitor" ───────────────────
        if any(x in linha_lower for x in
               ['espaco do leitor', 'cartas do leitor', 'espaco leitor',
                'leitor escreve']):
            secao_leitor = True

        # ── ALLCAPS: asterisco → sempre Jornalista ────────────────────────
        m = _RE_ASTERISCO.search(linha)
        if m and _RE_CAPS_ONLY.match(linha.rstrip('* ')):
            nome_caps = m.group(1).strip()
            if _caps_e_nome_pessoa(nome_caps):
                # Asterisco indica supervisão editorial → sempre Jornalista
                if nome_caps.title() not in autores or autores[nome_caps.title()] == 'Leitores':
                    autores[nome_caps.title()] = 'Jornalistas / Redatores'
            continue

        # ── ALLCAPS: nome + email editorial (ex: NÚBIA CRISTINA email@at...) ─
        m = _RE_CAPS_EMAIL.match(linha)
        if m:
            nome_caps = m.group(1).strip()
            if _caps_e_nome_pessoa(nome_caps):
                n = nome_caps.title()
                # A presença do email editorial sinaliza jornalista
                if n not in autores or _PRIO.get('Jornalistas / Redatores', 0) > _PRIO.get(autores[n], 0):
                    autores[n] = 'Jornalistas / Redatores'
            continue

        # ── ALLCAPS puro: byline de jornalista ────────────────────────────
        # ALLCAPS nomes são sempre tratados como jornalistas/colaboradores;
        # leitores são identificados via email ou secao_leitor em TC.
        if _RE_CAPS_ONLY.match(linha):
            if _caps_e_nome_pessoa(linha):
                n = linha.title()
                if n not in autores:
                    autores[n] = 'Jornalistas / Redatores'
                # Nunca rebaixa de Jornalista para Leitor via ALLCAPS
            continue   # linha CAPS: não aplica regras TC abaixo

        # A partir daqui: linhas que NÃO são todo em ALLCAPS

        if linha.endswith(':') or linha.endswith('...'):
            continue

        # ── Cargo explícito: Nome / Jornalista ───────────────────────────
        m = _RE_CARGO.search(linha)
        if m:
            _reg(m.group(1), _classificar(_norm(linha)))

        # ── Email pessoal com nome em TC antes ────────────────────────────
        m = _RE_EMAIL_NOME_TC.search(linha)
        if m:
            _reg(m.group(1), 'Leitores')

        # ── Prefixo "Por:", "Reportagem de:" ─────────────────────────────
        m = _RE_PREFIXO.match(linha)
        if m:
            cat = 'Leitores' if secao_leitor else _classificar(_norm(linha))
            _reg(m.group(1), cat)

        # ── Charge / ilustração de Nome ───────────────────────────────────
        m = _RE_CHARGE.search(linha)
        if m:
            _reg(m.group(1), 'Colunistas')

    # ── Resultado ─────────────────────────────────────────────────────────
    categorias: dict[str, list] = defaultdict(list)
    for nome, cat in autores.items():
        categorias[cat].append(nome)

    return {
        "total_autores": len(autores),
        "autores_unicos": sorted(autores),
        "por_categoria": {cat: sorted(nomes) for cat, nomes in categorias.items()},
    }


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


def ocr_de_pil(img_pil: Image.Image) -> str:
    """OCR em uma PIL Image diretamente (sem arquivo temporario).
    Usado tanto pelo botao de OCR quanto pelo fallback hibrido de busca por imagem."""
    LARGURA_MIN = 1800
    img = img_pil.convert("RGB")
    if img.width < LARGURA_MIN:
        fator = LARGURA_MIN / img.width
        img = img.resize((LARGURA_MIN, int(img.height * fator)), Image.LANCZOS)
    cinza = img.convert("L")
    cinza = ImageEnhance.Contrast(cinza).enhance(2.0)
    cinza = cinza.filter(ImageFilter.SHARPEN)
    config = "--oem 1 --psm 3 -l por+eng"
    return pytesseract.image_to_string(cinza, config=config).strip()


def ocr_da_imagem(caminho_img: str) -> str:
    """OCR em arquivo de imagem. Delegates para ocr_de_pil."""
    return ocr_de_pil(Image.open(caminho_img))


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
    """Busca texto_busca no PDF usando indice invertido por palavra.

    Indice invertido:
    - Constroi mapa palavra->blocos para cada pagina
    - Filtra blocos candidatos: so avalia blocos que compartilham >= 1 palavra
      com a busca — elimina 90%+ das comparacoes
    - Janela deslizante apenas ao redor dos blocos candidatos (nao O(n^2))
    """
    try:
        paginas = extrair_blocos_por_pagina(pdf_path)
    except Exception as e:
        cb_resultado(erro=str(e))
        return

    if not paginas:
        cb_resultado(erro="Nenhum texto encontrado no PDF.")
        return

    palavras_q   = palavras_relevantes(texto_busca)
    total_paginas = len(paginas)
    scored       = []
    melhor_score = 0.0

    for i, (num_pagina, info) in enumerate(paginas.items()):
        blocos = info["blocos"]
        pag_w  = info["pag_w"]
        pag_h  = info["pag_h"]

        if not blocos:
            cb_progresso(i + 1, total_paginas, num_pagina)
            continue

        # ── Indice invertido desta pagina: palavra -> set(indices de blocos) ────
        # Permite filtrar 90%+ dos blocos sem nenhum calculo de score
        idx_bloco: dict = {}
        for bi, bloco in enumerate(blocos):
            for p in palavras_relevantes(bloco["texto"]):
                idx_bloco.setdefault(p, set()).add(bi)

        # Blocos que compartilham pelo menos 1 palavra com a busca
        candidatos_idx = set()
        for p in palavras_q:
            candidatos_idx |= idx_bloco.get(p, set())

        if not candidatos_idx:
            cb_progresso(i + 1, total_paginas, num_pagina)
            continue

        # Para cada bloco candidato: avalia individualmente e em janelas com vizinhos
        vistos: set = set()
        for bi in sorted(candidatos_idx):
            # Janela: do bloco (bi-2) ao (bi+10)
            ini     = max(0, bi - 2)
            fim_max = min(len(blocos), bi + 11)

            for tam in range(1, fim_max - ini + 1):
                fim = ini + tam
                chave = (ini, fim)
                if chave in vistos:
                    continue
                vistos.add(chave)
                grupo = blocos[ini:fim]
                texto_grupo = " ".join(b["texto"] for b in grupo)
                bloco_virtual = {
                    "x0": min(b["x0"] for b in grupo),
                    "y0": min(b["y0"] for b in grupo),
                    "x1": max(b["x1"] for b in grupo),
                    "y1": max(b["y1"] for b in grupo),
                }
                score = calcular_score_texto(texto_busca, texto_grupo)
                if score > melhor_score:
                    melhor_score = score
                if score >= limiar:
                    scored.append((score, {
                        "pagina":       num_pagina,
                        "texto":        texto_grupo,
                        "bloco_orig":   bloco_virtual,
                        "pag_w":        pag_w,
                        "pag_h":        pag_h,
                        "todos_blocos": blocos,
                    }))

        cb_progresso(i + 1, total_paginas, num_pagina)

    # Expande bbox e deduplica
    matches_unicos = []
    regioes_vistas = []

    for score, cand in sorted(scored, key=lambda x: x[0], reverse=True):
        bloco = cand["bloco_orig"]
        x0, y0, x1, y1 = bbox_expandido(bloco, cand["todos_blocos"], tolerancia_pt=10)

        x_cm = pontos_para_cm(x0)
        y_cm = pontos_para_cm(y0)
        w_cm = pontos_para_cm(x1 - x0)
        h_cm = pontos_para_cm(y1 - y0)

        sobrepoe = any(
            cand["pagina"] == r["pagina"]
            and abs(x_cm - r["x_cm"]) < 1.5
            and abs(y_cm - r["y_cm"]) < 1.5
            for r in regioes_vistas
        )
        if sobrepoe:
            continue

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

    cb_resultado(matches=matches_unicos, melhor_score=melhor_score,
                 total=total_paginas, erro=None)


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
        super().__init__(parent, height=6, bg=CORES["fundo"],
                         highlightthickness=0, **kw)
        self._porcentagem = 0.0
        self._animando    = False
        self._pulso_x     = 0
        self._job         = None
        self.bind("<Configure>", self._desenhar)

    def atualizar(self, valor):
        self._porcentagem = max(0.0, min(1.0, valor))
        if valor <= 0.0 or valor >= 1.0:
            self._animando = False
            if self._job:
                self.after_cancel(self._job)
                self._job = None
        elif not self._animando:
            self._animando = True
            self._animar()
        self._desenhar()

    def _animar(self):
        if not self._animando:
            return
        self._pulso_x = (self._pulso_x + 3) % max(1, self.winfo_width())
        self._desenhar()
        self._job = self.after(30, self._animar)

    def _desenhar(self, event=None):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 2:
            return
        # Fundo com bordas arredondadas simuladas
        self.create_rectangle(0, 0, w, h, fill=CORES["barra_fundo"], outline="")
        preenchido = int(w * self._porcentagem)
        if preenchido > 0:
            # Gradiente azul → roxo
            passos = max(1, preenchido // 2)
            for i in range(passos):
                x0 = int(preenchido * i / passos)
                x1 = int(preenchido * (i + 1) / passos) + 1
                t  = i / max(passos - 1, 1)
                r  = int(0x4f + (0x7c - 0x4f) * t)
                g  = int(0x8e + (0x6a - 0x8e) * t)
                self.create_rectangle(x0, 0, x1, h,
                                      fill=f"#{r:02x}{g:02x}f7", outline="")
            # Pulso brilhante (shimmer) enquanto animando
            if self._animando and preenchido > 10:
                px = self._pulso_x % preenchido
                brilho_w = max(20, preenchido // 5)
                for di in range(brilho_w):
                    alpha = 1.0 - abs(di - brilho_w // 2) / (brilho_w // 2)
                    bx = (px + di) % preenchido
                    bval = int(255 * alpha * 0.35)
                    self.create_rectangle(
                        bx, 0, bx + 1, h,
                        fill=f"#{min(255,0x4f+bval):02x}"
                             f"{min(255,0x8e+bval):02x}"
                             f"{min(255,0xf7+bval):02x}",
                        outline="")


def _btn_hover(btn, cor_normal, cor_hover, fg_normal="#ffffff", fg_hover="#ffffff"):
    """Adiciona efeito hover suave a um botao tk.Button."""
    btn.bind("<Enter>",    lambda e: btn.config(bg=cor_hover,  fg=fg_hover))
    btn.bind("<Leave>",    lambda e: btn.config(bg=cor_normal, fg=fg_normal))
    btn.bind("<Button-1>", lambda e: btn.config(bg=CORES["roxo"]))
    btn.bind("<ButtonRelease-1>", lambda e: btn.config(bg=cor_hover))


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

            lbl_pdf.config(text=f"L {usar_w:.3f} cm  x  A {usar_h:.3f} cm")

            if papel and pag_w_cm and pag_h_cm:
                escala = escalar_para_jornal(usar_w, usar_h,
                                             pag_w_cm, pag_h_cm,
                                             papel[0], papel[1])
                lbl_print.config(
                    text=f"L {escala['w_cm']:.3f} cm  x  A {escala['h_cm']:.3f} cm",
                    fg=CORES["verde"])
            else:
                lbl_print.config(text="-- selecione um formato --", fg=CORES["texto3"])

            if pag_w_cm and pag_h_cm:
                prop = calcular_proporcao(usar_w, usar_h, pag_w_cm, pag_h_cm)
                lbl_prop.config(
                    text=f"Larg: {prop['w_pct']:.2f}%  "
                         f"Alt: {prop['h_pct']:.2f}%  "
                         f"Area: {prop['area_pct']:.2f}%")
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
        tk.Label(linha_pos, text=f"X {x_cm:.3f} cm  Y {y_cm:.3f} cm",
                 font=FONTE_MONO_P, fg=CORES["texto2"],
                 bg=CORES["painel2"], padx=6).pack(side="left")

    if pag_w_cm and pag_h_cm:
        linha_pag = nova_linha(" PAG  ", CORES["texto3"])
        tk.Label(linha_pag, text=f"L {pag_w_cm:.3f} cm  x  A {pag_h_cm:.3f} cm",
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
        self.title("Dimensao Anuncio")
        self.geometry("1080x780")
        self.minsize(860, 620)
        self.configure(bg=CORES["fundo"])

        self.caminho_pdf    = tk.StringVar()
        self.caminho_img    = tk.StringVar()
        self._lista_imagens = []   # lista de caminhos (strings)
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
        self._toast_job   = None
        self._montar_interface()
        self._centralizar_janela()

    def _montar_interface(self):
        # ── Header premium ─────────────────────────────────────────────────
        header = tk.Frame(self, bg=CORES["painel"],
                          highlightbackground=CORES["borda"], highlightthickness=1)
        header.pack(fill="x")

        inner_h = tk.Frame(header, bg=CORES["painel"])
        inner_h.pack(fill="x", padx=24, pady=12)

        # Ícone + título
        col_ico = tk.Frame(inner_h, bg=CORES["painel"])
        col_ico.pack(side="left")
        tk.Label(col_ico,
                 text="▣",
                 font=("Segoe UI", 26),
                 fg=CORES["azul"], bg=CORES["painel"]).pack(side="left", padx=(0, 10))

        col_txt = tk.Frame(inner_h, bg=CORES["painel"])
        col_txt.pack(side="left")
        tk.Label(col_txt,
                 text="Dimensão Anuncio",
                 font=("Segoe UI", 17, "bold"),
                 fg=CORES["texto"], bg=CORES["painel"]).pack(anchor="w")
        tk.Label(col_txt,
                 text="Medição de anúncios para jornais impressos",
                 font=("Segoe UI", 9),
                 fg=CORES["texto2"], bg=CORES["painel"]).pack(anchor="w")

        # Badge de versão
        tk.Label(inner_h,
                 text=" v2.0 ",
                 font=("Segoe UI", 8, "bold"),
                 fg=CORES["azul"], bg=CORES["painel3"],
                 padx=4, pady=2).pack(side="right", anchor="n")

        # ── Toast de notificação (overlay flutuante) ────────────────────────
        self._toast = tk.Label(self,
            text="", font=("Segoe UI", 10, "bold"),
            fg="#ffffff", bg=CORES["roxo"],
            padx=18, pady=8)
        # Não faz pack — é posicionado via place quando necessário

        corpo = tk.Frame(self, bg=CORES["fundo"])
        corpo.pack(fill="both", expand=True, padx=20, pady=10)

        # Painel esquerdo com scroll de mouse
        _esq_outer = tk.Frame(corpo, bg=CORES["fundo"], width=390)
        _esq_outer.pack(side="left", fill="y")
        _esq_outer.pack_propagate(False)

        _esq_canvas = tk.Canvas(_esq_outer, bg=CORES["fundo"],
                                width=390, highlightthickness=0)
        _esq_scroll = tk.Scrollbar(_esq_outer, orient="vertical",
                                   command=_esq_canvas.yview)
        _esq_canvas.configure(yscrollcommand=_esq_scroll.set)
        _esq_canvas.pack(side="left", fill="both", expand=True)
        # Scrollbar só aparece se necessário (não ocupa espaço fixo)
        _esq_scroll.pack(side="right", fill="y")

        painel_esq = tk.Frame(_esq_canvas, bg=CORES["fundo"])
        _esq_win = _esq_canvas.create_window((0, 0), window=painel_esq, anchor="nw")

        def _ajustar_scroll(event=None):
            _esq_canvas.configure(scrollregion=_esq_canvas.bbox("all"))
            _esq_canvas.itemconfig(_esq_win, width=_esq_canvas.winfo_width())

        painel_esq.bind("<Configure>", _ajustar_scroll)
        _esq_canvas.bind("<Configure>",
                         lambda e: _esq_canvas.itemconfig(_esq_win, width=e.width))

        def _scroll_mouse(event):
            _esq_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Vincula scroll do mouse ao canvas e a todos os filhos dinamicamente
        def _bind_scroll(widget):
            widget.bind("<MouseWheel>", _scroll_mouse)
            for child in widget.winfo_children():
                _bind_scroll(child)

        painel_esq.bind("<MouseWheel>", _scroll_mouse)
        _esq_canvas.bind("<MouseWheel>", _scroll_mouse)
        # Re-vincula sempre que novos widgets forem adicionados ao painel
        painel_esq.bind("<Configure>", lambda e: (_ajustar_scroll(), _bind_scroll(painel_esq)))

        self._painel_esq = painel_esq
        self._esq_canvas = _esq_canvas
        self._bind_scroll_esq = _bind_scroll
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

        # ---- Bloco de multiplas imagens ----
        tk.Label(parent, text="IMAGENS DE BUSCA", font=("Segoe UI", 8, "bold"),
                 fg=CORES["azul"], bg=CORES["fundo"], anchor="w").pack(
                     fill="x", pady=(6, 2))

        self._frame_imgs = tk.Frame(parent, bg=CORES["fundo"])
        self._frame_imgs.pack(fill="x")

        linha_btn_imgs = tk.Frame(parent, bg=CORES["fundo"])
        linha_btn_imgs.pack(fill="x", pady=(4, 2))
        tk.Button(linha_btn_imgs, text="+ Adicionar imagem",
            font=FONTE_PEQUENA, fg=CORES["texto2"], bg=CORES["painel2"],
            relief="flat", cursor="hand2", padx=8, pady=5,
            activebackground=CORES["painel3"],
            command=self._adicionar_imagem
        ).pack(side="left")
        tk.Button(linha_btn_imgs, text="Limpar",
            font=FONTE_PEQUENA, fg=CORES["texto3"], bg=CORES["fundo"],
            relief="flat", cursor="hand2", padx=6, pady=5,
            command=self._limpar_imagens
        ).pack(side="left", padx=(4, 0))

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

        self._btn_buscar = tk.Button(parent, text="▶  Iniciar Busca",
            font=("Segoe UI", 11, "bold"), fg="#ffffff", bg=CORES["azul"],
            activebackground=CORES["roxo"], activeforeground="#ffffff",
            relief="flat", cursor="hand2", pady=10,
            command=self._iniciar_busca)
        self._btn_buscar.pack(fill="x", pady=(0, 5))
        _btn_hover(self._btn_buscar, CORES["azul"], CORES["roxo"])

        self._btn_listar = tk.Button(parent, text="☰  Listar Todos os Anuncios",
            font=("Segoe UI", 9, "bold"), fg=CORES["texto2"], bg=CORES["painel2"],
            activebackground=CORES["painel3"], activeforeground=CORES["ciano"],
            relief="flat", cursor="hand2", pady=7,
            command=self._iniciar_listagem)
        self._btn_listar.pack(fill="x", pady=(0, 5))
        _btn_hover(self._btn_listar, CORES["painel2"], CORES["painel3"],
                   CORES["texto2"], CORES["ciano"])

        self._btn_autores = tk.Button(parent, text="✎  Extrair Autores",
            font=("Segoe UI", 9, "bold"), fg=CORES["texto2"], bg=CORES["painel2"],
            activebackground=CORES["painel3"], activeforeground=CORES["lilas"],
            relief="flat", cursor="hand2", pady=7,
            command=self._iniciar_extracao_autores)
        self._btn_autores.pack(fill="x", pady=(0, 5))
        _btn_hover(self._btn_autores, CORES["painel2"], CORES["painel3"],
                   CORES["texto2"], CORES["lilas"])

        self._btn_autores_atarde = tk.Button(parent, text="⬛  Autores A TARDE",
            font=("Segoe UI", 9, "bold"), fg=CORES["texto2"], bg=CORES["painel2"],
            activebackground=CORES["painel3"], activeforeground=CORES["ciano"],
            relief="flat", cursor="hand2", pady=7,
            command=self._iniciar_extracao_atarde)
        self._btn_autores_atarde.pack(fill="x")
        _btn_hover(self._btn_autores_atarde, CORES["painel2"], CORES["painel3"],
                   CORES["texto2"], CORES["ciano"])

    def _montar_painel_direito(self, parent):
        # ── Status bar ─────────────────────────────────────────────────────
        status_bar = tk.Frame(parent, bg=CORES["painel2"],
                              highlightbackground=CORES["borda"], highlightthickness=1)
        status_bar.pack(fill="x", pady=(0, 6))

        self._lbl_status_ico = tk.Label(status_bar,
            text="◌", font=("Segoe UI", 11),
            fg=CORES["texto3"], bg=CORES["painel2"], padx=8)
        self._lbl_status_ico.pack(side="left")

        self._lbl_status = tk.Label(status_bar,
            text="Aguardando...",
            font=("Segoe UI", 9), fg=CORES["texto3"],
            bg=CORES["painel2"], anchor="w")
        self._lbl_status.pack(side="left", fill="x", expand=True, pady=6)

        self._lbl_progresso = tk.Label(status_bar,
            text="", font=FONTE_PEQUENA,
            fg=CORES["texto3"], bg=CORES["painel2"],
            anchor="e", padx=8)
        self._lbl_progresso.pack(side="right")

        self._barra = BarraProgresso(parent)
        self._barra.pack(fill="x", pady=(0, 8))

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

        self._aba_autores = tk.Frame(self._abas, bg=CORES["painel"])
        self._abas.add(self._aba_autores, text="Autores")
        self._canvas_autores, self._frame_autores = self._criar_area_rolavel(self._aba_autores)

        self._aba_buscar_autor = tk.Frame(self._abas, bg=CORES["painel"])
        self._abas.add(self._aba_buscar_autor, text="Buscar Autor")
        self._montar_aba_buscar_autor(self._aba_buscar_autor)

        self._aba_autores_atarde = tk.Frame(self._abas, bg=CORES["painel"])
        self._abas.add(self._aba_autores_atarde, text="Autores A TARDE")
        self._canvas_autores_atarde, self._frame_autores_atarde = \
            self._criar_area_rolavel(self._aba_autores_atarde)

        self._mostrar_placeholder(self._frame_busca, "busca")
        self._mostrar_placeholder(self._frame_lista, "lista")
        self._mostrar_placeholder(self._frame_autores, "autores")

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

        tk.Button(topo, text="▶  Buscar texto no PDF",
            font=("Segoe UI", 11, "bold"),
            fg="#ffffff", bg=CORES["azul"],
            activebackground=CORES["roxo"],
            activeforeground="#ffffff",
            relief="flat", cursor="hand2", pady=10,
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
            f"X {dados['x_cm']:.3f} cm  Y {dados['y_cm']:.3f} cm")
        nova_linha_m(" TAM  ", CORES["ciano"],
            f"L {dados['w_cm']:.3f} cm  x  A {dados['h_cm']:.3f} cm")

        # Colunagem do jornal ativo
        nome_jornal = self.jornal_ativo.get()
        jornal = JORNAIS_CADASTRADOS.get(nome_jornal)
        if jornal:
            info = calcular_info_colunas(
                dados["w_cm"], dados.get("x_cm") or 0.0, jornal)
            formato_nome = identificar_formato(
                dados["w_cm"], dados["h_cm"],
                jornal["largura"], jornal["altura"])
            sinal = "+" if info["desvio"] >= 0 else ""
            nova_linha_m(" COL  ", CORES["lilas"],
                f"Col {info['col_ini']} a {info['col_fim']}  "
                f"|  {info['num_col']} col ({info['cols_exato']:.3f} exato)  "
                f"|  alt {dados['h_cm']:.3f} cm")
            nova_linha_m(" DIM  ", CORES["ciano"],
                f"Padrao {info['larg_padrao']:.3f} cm  "
                f"|  Medido {dados['w_cm']:.3f} cm  "
                f"|  Δ {sinal}{info['desvio']:.3f} cm")
            nova_linha_m(" 1COL ", CORES["texto2"],
                f"1 col = {info['larg_1col']:.3f} cm  "
                f"|  margem = {info['margem']:.3f} cm")
            nova_linha_m(" FMT  ", CORES["roxo"], formato_nome)

        if dados.get("pag_w") and dados.get("pag_h"):
            prop = calcular_proporcao(dados["w_cm"], dados["h_cm"],
                                      dados["pag_w"], dados["pag_h"])
            nova_linha_m(" PROP ", CORES["verde"],
                f"Larg: {prop['w_pct']:.2f}%  Alt: {prop['h_pct']:.2f}%  "
                f"Area: {prop['area_pct']:.2f}%")
            nova_linha_m(" PAG  ", CORES["texto3"],
                f"L {dados['pag_w']:.3f} cm  x  A {dados['pag_h']:.3f} cm")

        # ── Autor / Colaborador (busca por texto) ──
        pdf = self.caminho_pdf.get()
        if pdf and Path(pdf).exists():
            try:
                doc_temp = fitz.open(pdf)
                pag_temp = doc_temp[dados["pagina"] - 1]
                # Usa coordenadas do match para extrair texto
                x0_pt = dados.get("x_cm", 0) / PT_PARA_CM
                y0_pt = dados.get("y_cm", 0) / PT_PARA_CM
                w_pt  = dados.get("w_cm", 0) / PT_PARA_CM
                h_pt  = dados.get("h_cm", 0) / PT_PARA_CM
                rect = fitz.Rect(x0_pt, y0_pt, x0_pt + w_pt, y0_pt + h_pt)
                txt_regiao = pag_temp.get_text("text", clip=rect)
                doc_temp.close()
                ai = extrair_autor(txt_regiao)
                if ai and ai.get("autor"):
                    tipo_map = {"reporter": "Repórter", "redacao": "Redação",
                                "articulista": "Articulista", "desconhecido": ""}
                    tipo_label = tipo_map.get(ai["tipo"], "")
                    cor_tipo = (CORES["verde"] if ai["confianca"] >= 0.85
                                else CORES["amarelo"] if ai["confianca"] >= 0.65
                                else CORES["texto2"])
                    nova_linha_m(" AUTOR", CORES["verde"],
                        f"{ai['autor']}  ({tipo_label}  {ai['confianca']:.0%})")
                    if ai.get("colaboradores"):
                        nova_linha_m(" COLAB", CORES["ciano"],
                            ", ".join(ai["colaboradores"]))
            except Exception:
                pass

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
            self._mostrar_toast("✓  PDF salvo com marcação!", CORES["verde"], 3000)
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

    def _mostrar_toast(self, msg, cor=None, duracao=2500):
        """Exibe notificacao flutuante no canto inferior direito por `duracao` ms."""
        if self._toast_job:
            self.after_cancel(self._toast_job)
        bg = cor or CORES["roxo"]
        self._toast.config(text=f"  {msg}  ", bg=bg)
        self._toast.place(relx=1.0, rely=1.0, anchor="se", x=-16, y=-16)
        self._toast.lift()
        self._toast_job = self.after(duracao, self._esconder_toast)

    def _esconder_toast(self):
        self._toast.place_forget()
        self._toast_job = None

    def _titulo_secao(self, parent, texto):
        f = tk.Frame(parent, bg=CORES["fundo"])
        f.pack(fill="x", pady=(8, 4))
        tk.Label(f, text=texto,
                 font=("Segoe UI", 7, "bold"),
                 fg=CORES["azul"], bg=CORES["fundo"],
                 padx=0).pack(side="left")
        tk.Frame(f, bg=CORES["borda"], height=1).pack(
            side="left", fill="x", expand=True, padx=(6, 0), pady=4)

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
            area = j["largura"] - j["margem"] * (j["colunas"] - 1)
            self._lbl_jornal_info.config(
                text=f"  {j['largura']}x{j['altura']} cm  "
                     f"{j['colunas']} col  "
                     f"1col={lc:.3f} cm  "
                     f"marg={j['margem']:.3f} cm  "
                     f"util={area:.3f} cm")
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

    def _adicionar_imagem(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecionar imagem(ns) de busca",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.tiff *.webp"),
                       ("Todos", "*.*")])
        if caminhos:
            for c in caminhos:
                if c not in self._lista_imagens:
                    self._lista_imagens.append(c)
            self._atualizar_lista_imagens()

    def _limpar_imagens(self):
        self._lista_imagens.clear()
        self._atualizar_lista_imagens()
        self._lbl_preview.config(image="")
        try:
            del self._lbl_preview._foto
        except AttributeError:
            pass

    def _atualizar_lista_imagens(self):
        for w in self._frame_imgs.winfo_children():
            w.destroy()

        if not self._lista_imagens:
            tk.Label(self._frame_imgs,
                     text="Nenhuma imagem adicionada",
                     font=FONTE_PEQUENA, fg=CORES["texto3"],
                     bg=CORES["fundo"], anchor="w").pack(fill="x")
            return

        for idx, caminho in enumerate(self._lista_imagens):
            linha = tk.Frame(self._frame_imgs, bg=CORES["painel2"],
                             highlightbackground=CORES["borda"], highlightthickness=1)
            linha.pack(fill="x", pady=1)

            # Mini-thumb
            try:
                img_t = Image.open(caminho).convert("RGB")
                img_t.thumbnail((36, 36), Image.LANCZOS)
                foto_t = ImageTk.PhotoImage(img_t)
                lbl_t = tk.Label(linha, image=foto_t, bg=CORES["painel2"])
                lbl_t.image = foto_t   # evita GC
                lbl_t.pack(side="left", padx=4, pady=3)
            except Exception:
                tk.Label(linha, text="IMG", font=FONTE_BADGE,
                         fg=CORES["texto3"], bg=CORES["painel2"],
                         width=5).pack(side="left", padx=4)

            nome = Path(caminho).name
            nome_curto = nome if len(nome) <= 26 else "..." + nome[-23:]
            tk.Label(linha, text=f"{idx+1}. {nome_curto}",
                     font=FONTE_PEQUENA, fg=CORES["texto"],
                     bg=CORES["painel2"], anchor="w").pack(side="left", fill="x", expand=True)

            idx_local = idx
            tk.Button(linha, text="x", font=("Segoe UI", 8),
                      fg=CORES["vermelho"], bg=CORES["painel2"],
                      relief="flat", cursor="hand2", padx=4,
                      command=lambda i=idx_local: self._remover_imagem(i)
                      ).pack(side="right", padx=4)

        # Atualiza preview com a primeira imagem
        if self._lista_imagens:
            try:
                img = Image.open(self._lista_imagens[0]).convert("RGB")
                img.thumbnail((200, 72), Image.LANCZOS)
                foto = ImageTk.PhotoImage(img)
                self._lbl_preview.config(image=foto)
                self._lbl_preview._foto = foto
            except Exception:
                pass

    def _remover_imagem(self, idx):
        if 0 <= idx < len(self._lista_imagens):
            self._lista_imagens.pop(idx)
            self._atualizar_lista_imagens()

    # Mantido para compatibilidade com BotaoArquivo antigo (nao usado mais)
    def _selecionar_imagem(self):
        self._adicionar_imagem()

    def _iniciar_busca(self):
        if self._buscando:
            return
        pdf = self.caminho_pdf.get()
        if not pdf:
            messagebox.showwarning("Atencao", "Selecione um arquivo PDF primeiro.")
            return
        if not self._lista_imagens:
            messagebox.showwarning("Atencao", "Adicione ao menos uma imagem de busca.")
            return
        if not Path(pdf).exists():
            messagebox.showerror("Erro", f"Arquivo nao encontrado:\n{pdf}")
            return

        imagens_busca = []
        for caminho in self._lista_imagens:
            if not Path(caminho).exists():
                messagebox.showerror("Erro", f"Imagem nao encontrada:\n{caminho}")
                return
            try:
                imagens_busca.append(Image.open(caminho).convert("RGB"))
            except Exception as e:
                messagebox.showerror("Erro",
                    f"Nao foi possivel abrir:\n{caminho}\n{e}")
                return

        self._buscando = True
        self._btn_buscar.config(text="⏳  Analisando...", state="disabled",
                                bg=CORES["painel2"], fg=CORES["texto2"])
        self._limpar_frame(self._frame_busca)
        self._barra.atualizar(0)
        self._lbl_status.config(text="Analisando...", fg=CORES["texto2"])
        self._lbl_status_ico.config(text="⟳", fg=CORES["amarelo"])
        self._abas.select(self._aba_busca)
        threading.Thread(
            target=buscar_imagem_no_pdf,
            args=(pdf, imagens_busca, METODOS[self.metodo_busca.get()],
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
        self._btn_listar.config(text="⏳  Escaneando...", state="disabled",
                                bg=CORES["painel2"], fg=CORES["texto3"])
        self._limpar_frame(self._frame_lista)
        self._barra.atualizar(0)
        self._lbl_status.config(text="Escaneando o PDF...", fg=CORES["texto2"])
        self._lbl_status_ico.config(text="⟳", fg=CORES["amarelo"])
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
                text=f"Pág {pagina}  ·  {atual}/{total}",
                fg=CORES["texto3"])
            self._lbl_status.config(
                text=f"Processando...  {atual / total:.0%}",
                fg=CORES["texto2"])
            self._lbl_status_ico.config(text="⟳", fg=CORES["amarelo"])
        self.after(0, atualizar)

    def _cb_resultado_busca(self, matches=None, melhor_score=0.0, total=0, erro=None):
        def atualizar():
            self._buscando = False
            self._btn_buscar.config(text="▶  Iniciar Busca", state="normal",
                                    bg=CORES["azul"], fg="#ffffff")
            self._barra.atualizar(1.0 if not erro else 0.0)
            if erro:
                self._lbl_status.config(text=f"Erro: {erro}", fg=CORES["vermelho"])
                self._lbl_status_ico.config(text="✕", fg=CORES["vermelho"])
                self._mostrar_placeholder(self._frame_busca, "busca", erro)
                self._mostrar_toast(f"Erro: {erro}", CORES["vermelho"])
                return
            self._lbl_progresso.config(
                text=f"{total} página(s) analisada(s)", fg=CORES["texto3"])
            if matches:
                paginas = sorted(set(m["pagina"] for m in matches))
                self._lbl_status.config(
                    text=f"Encontrado em {len(paginas)} página(s): "
                         f"{', '.join(map(str, paginas))}",
                    fg=CORES["verde"])
                self._lbl_status_ico.config(text="✔", fg=CORES["verde"])
                self._mostrar_matches(matches)
                self._mostrar_toast(
                    f"✔  Encontrado em {len(paginas)} página(s)",
                    CORES["verde"], 3000)
            else:
                self._lbl_status.config(
                    text=f"Não encontrado  "
                         f"(melhor: {melhor_score:.1%}  "
                         f"limiar: {self.limiar.get():.0%})",
                    fg=CORES["vermelho"])
                self._lbl_status_ico.config(text="✕", fg=CORES["vermelho"])
                self._mostrar_nao_encontrado(melhor_score)
                self._mostrar_toast(
                    f"Não encontrado  ·  melhor: {melhor_score:.1%}",
                    CORES["vermelho"], 3500)
        self.after(0, atualizar)

    def _cb_resultado_listagem(self, imagens=None, total=0, erro=None):
        def atualizar():
            self._buscando = False
            self._btn_listar.config(text="☰  Listar Todos os Anuncios",
                                    state="normal",
                                    bg=CORES["painel2"], fg=CORES["texto2"])
            self._barra.atualizar(1.0 if not erro else 0.0)
            if erro:
                self._lbl_status.config(text=f"Erro: {erro}", fg=CORES["vermelho"])
                self._lbl_status_ico.config(text="✕", fg=CORES["vermelho"])
                self._mostrar_placeholder(self._frame_lista, "lista", erro)
                self._mostrar_toast(f"Erro: {erro}", CORES["vermelho"])
                return
            self._lbl_progresso.config(
                text=f"{total} anúncio(s) encontrado(s)", fg=CORES["texto3"])
            self._lbl_status.config(
                text=f"{total} anúncio(s) listado(s)", fg=CORES["ciano"])
            self._lbl_status_ico.config(text="✔", fg=CORES["ciano"])
            self._mostrar_todos(imagens)
            self._mostrar_toast(f"☰  {total} anúncio(s) listado(s)",
                                CORES["ciano"], 2500)
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
        elif aba == "autores":
            texto = "Selecione um PDF e clique em Extrair Autores"
        else:
            texto = "Selecione um PDF e clique em Listar Todos os Anuncios"
        tk.Label(frame, text=texto, font=FONTE_LABEL,
                 fg=CORES["texto3"], bg=CORES["painel"],
                 justify="center", pady=40).pack(expand=True, fill="both")

    def _mostrar_nao_encontrado(self, melhor_score):
        self._limpar_frame(self._frame_busca)
        f = tk.Frame(self._frame_busca, bg=CORES["painel"])
        f.pack(fill="both", expand=True, padx=16, pady=32)
        tk.Label(f, text="✕",
                 font=("Segoe UI", 42),
                 fg=CORES["vermelho"], bg=CORES["painel"]).pack()
        tk.Label(f, text="Anúncio não encontrado",
                 font=("Segoe UI", 14, "bold"),
                 fg=CORES["texto"], bg=CORES["painel"]).pack(pady=(4, 0))
        barra_pct = tk.Frame(f, bg=CORES["painel"])
        barra_pct.pack(pady=(16, 0))
        cor_m = (CORES["verde"] if melhor_score >= 0.75
                 else CORES["amarelo"] if melhor_score >= 0.50
                 else CORES["vermelho"])
        tk.Label(barra_pct,
                 text=f"Maior similaridade encontrada: ",
                 font=FONTE_PEQUENA, fg=CORES["texto2"],
                 bg=CORES["painel"]).pack(side="left")
        tk.Label(barra_pct,
                 text=f"{melhor_score:.1%}",
                 font=("Segoe UI", 11, "bold"),
                 fg=cor_m, bg=CORES["painel"]).pack(side="left")
        tk.Label(f,
                 text="Tente reduzir o limiar de sensibilidade ou use outra imagem.",
                 font=FONTE_PEQUENA, fg=CORES["texto2"],
                 bg=CORES["painel"], justify="center").pack(pady=(8, 0))

    def _mostrar_matches(self, matches):
        self._limpar_frame(self._frame_busca)
        cab = tk.Frame(self._frame_busca, bg=CORES["painel3"])
        cab.pack(fill="x")
        tk.Label(cab,
                 text=f"  ✔  {len(matches)} correspondência(s) encontrada(s)",
                 font=("Segoe UI", 9, "bold"),
                 fg=CORES["verde"], bg=CORES["painel3"], pady=9).pack(side="left")
        for match in sorted(matches, key=lambda x: x["score"], reverse=True):
            self._criar_card(self._frame_busca, match, eh_match=True)

    def _mostrar_todos(self, imagens):
        self._limpar_frame(self._frame_lista)
        cab = tk.Frame(self._frame_lista, bg=CORES["painel3"])
        cab.pack(fill="x")
        tk.Label(cab,
                 text=f"  ☰  {len(imagens)} anúncio(s) no PDF",
                 font=("Segoe UI", 9, "bold"),
                 fg=CORES["ciano"], bg=CORES["painel3"], pady=9).pack(side="left")
        for info_img in imagens:
            self._criar_card(self._frame_lista, info_img, eh_match=False)

    def _criar_card(self, parent, dados, eh_match):
        # Card com borda esquerda colorida e hover highlight
        cor_borda = (CORES["verde"] if eh_match and dados.get("score", 0) >= 0.85
                     else CORES["amarelo"] if eh_match and dados.get("score", 0) >= 0.65
                     else CORES["azul"] if eh_match
                     else CORES["borda"])

        outer = tk.Frame(parent, bg=CORES["painel2"],
                         highlightbackground=CORES["borda"], highlightthickness=1)
        outer.pack(fill="x", padx=8, pady=4)

        # Barra lateral colorida
        tk.Frame(outer, bg=cor_borda, width=4).pack(side="left", fill="y")

        card = tk.Frame(outer, bg=CORES["painel2"])
        card.pack(side="left", fill="both", expand=True)

        # Hover highlight
        def _on_enter(e):
            outer.config(highlightbackground=cor_borda)
            card.config(bg=CORES["painel3"])
            for w in card.winfo_children():
                try:
                    w.config(bg=CORES["painel3"])
                except Exception:
                    pass
        def _on_leave(e):
            outer.config(highlightbackground=CORES["borda"])
            card.config(bg=CORES["painel2"])
            for w in card.winfo_children():
                try:
                    w.config(bg=CORES["painel2"])
                except Exception:
                    pass
        outer.bind("<Enter>", _on_enter)
        outer.bind("<Leave>", _on_leave)
        card.bind("<Enter>",  _on_enter)
        card.bind("<Leave>",  _on_leave)

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

            img_idx = dados.get("img_match_idx")
            if img_idx and img_idx > 1:
                tk.Label(linha_titulo, text=f"Img #{img_idx}",
                         font=FONTE_BADGE, fg=CORES["ciano"],
                         bg=CORES["painel3"], padx=4).pack(side="left", padx=(6, 0))
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
        x_cm_real = dados.get("x_cm") or 0.0
        if jornal and w_cm_real and h_cm_real:
            info     = calcular_info_colunas(w_cm_real, x_cm_real, jornal)
            fmt_nome = identificar_formato(
                w_cm_real, h_cm_real, jornal["largura"], jornal["altura"])
            sinal    = "+" if info["desvio"] >= 0 else ""

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
                f"Col {info['col_ini']} a {info['col_fim']}  "
                f"|  {info['num_col']} col ({info['cols_exato']:.3f} exato)  "
                f"|  alt {h_cm_real:.3f} cm")
            linha_col(" DIM  ", CORES["ciano"],
                f"Padrao {info['larg_padrao']:.3f} cm  "
                f"|  Medido {w_cm_real:.3f} cm  "
                f"|  Δ {sinal}{info['desvio']:.3f} cm")
            linha_col(" 1COL ", CORES["texto2"],
                f"1 col = {info['larg_1col']:.3f} cm  "
                f"|  margem = {info['margem']:.3f} cm")
            linha_col(" FMT  ", CORES["roxo"], fmt_nome)

        # ── Autor / Colaborador ──
        autor_info = dados.get("autor_info")
        if autor_info and autor_info.get("autor"):
            frame_autor = tk.Frame(area_info, bg=CORES["painel2"])
            frame_autor.pack(fill="x", pady=(4, 0))
            tipo_map = {"reporter": "Repórter", "redacao": "Redação",
                        "articulista": "Articulista", "desconhecido": ""}
            tipo_label = tipo_map.get(autor_info["tipo"], "")
            cor_tipo = (CORES["verde"] if autor_info["confianca"] >= 0.85
                        else CORES["amarelo"] if autor_info["confianca"] >= 0.65
                        else CORES["texto2"])

            la = tk.Frame(frame_autor, bg=CORES["painel2"])
            la.pack(fill="x", pady=1)
            tk.Label(la, text=" AUTOR", font=FONTE_BADGE,
                     fg=CORES["verde"], bg=CORES["painel3"],
                     padx=2).pack(side="left")
            tk.Label(la, text=f"{autor_info['autor']}  ({tipo_label}  {autor_info['confianca']:.0%})",
                     font=("Consolas", 9, "bold"),
                     fg=cor_tipo, bg=CORES["painel2"],
                     padx=6).pack(side="left")

            if autor_info.get("colaboradores"):
                lc = tk.Frame(frame_autor, bg=CORES["painel2"])
                lc.pack(fill="x", pady=1)
                tk.Label(lc, text=" COLAB", font=FONTE_BADGE,
                         fg=CORES["ciano"], bg=CORES["painel3"],
                         padx=2).pack(side="left")
                tk.Label(lc, text=", ".join(autor_info["colaboradores"]),
                         font=FONTE_MONO_P, fg=CORES["texto2"],
                         bg=CORES["painel2"], padx=6).pack(side="left")

        tk.Button(area_info,
            text="⬇  Baixar PDF com marcação",
            font=("Segoe UI", 9, "bold"),
            fg="#ffffff", bg=CORES["verde"],
            activebackground="#22a870",
            activeforeground="#ffffff",
            relief="flat", cursor="hand2",
            padx=10, pady=7,
            command=lambda d=dados: self._baixar_pdf_marcado([d])
        ).pack(fill="x", pady=(8, 2))

    # ══════════════════════════════════════════════════════════════
    #  ABA AUTORES — Extração de todos os autores do PDF
    # ══════════════════════════════════════════════════════════════

    def _iniciar_extracao_autores(self):
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
        self._btn_autores.config(text="⏳  Extraindo...", state="disabled",
                                 bg=CORES["painel2"], fg=CORES["texto3"])
        self._limpar_frame(self._frame_autores)
        self._barra.atualizar(0)
        self._lbl_status.config(text="Extraindo autores...", fg=CORES["texto2"])
        self._lbl_status_ico.config(text="⟳", fg=CORES["amarelo"])
        self._abas.select(self._aba_autores)
        threading.Thread(
            target=extrair_todos_autores_pdf,
            args=(pdf, self._cb_progresso_autores, self._cb_resultado_autores),
            daemon=True
        ).start()

    def _cb_progresso_autores(self, atual, total, pagina):
        def atualizar():
            self._barra.atualizar(atual / total)
            self._lbl_progresso.config(
                text=f"Pág {atual}/{total}", fg=CORES["texto3"])
            self._lbl_status.config(
                text=f"Extraindo autores...  {atual / total:.0%}",
                fg=CORES["texto2"])
            self._lbl_status_ico.config(text="⟳", fg=CORES["amarelo"])
        self.after(0, atualizar)

    def _cb_resultado_autores(self, materias=None, total=0, erro=None):
        def atualizar():
            self._buscando = False
            self._btn_autores.config(text="✎  Extrair Autores", state="normal",
                                     bg=CORES["painel2"], fg=CORES["texto2"])
            self._barra.atualizar(1.0 if not erro else 0.0)
            if erro:
                self._lbl_status.config(text=f"Erro: {erro}", fg=CORES["vermelho"])
                self._mostrar_placeholder(self._frame_autores, "autores", erro)
                return
            self._lbl_progresso.config(
                text=f"{total} materia(s) analisada(s)", fg=CORES["texto3"])

            # Separa encontrados vs desconhecidos
            com_autor = [m for m in materias if m["autor_info"].get("autor")]
            sem_autor = [m for m in materias if not m["autor_info"].get("autor")]

            self._lbl_status.config(
                text=f"{len(com_autor)} autor(es) encontrado(s)  |  "
                     f"{len(sem_autor)} sem autor  |  {total} materia(s)",
                fg=CORES["verde"] if com_autor else CORES["amarelo"])
            self._lbl_status_ico.config(
                text="✔" if com_autor else "◌",
                fg=CORES["verde"] if com_autor else CORES["amarelo"])
            self._mostrar_autores(materias)
            self._mostrar_toast(
                f"✎  {len(com_autor)} autor(es) encontrado(s)",
                CORES["lilas"], 2500)
        self.after(0, atualizar)

    # ── Extração A TARDE ──────────────────────────────────────────────────

    def _iniciar_extracao_atarde(self):
        """Lê o texto completo do PDF e extrai autores via extrair_autores_ataarde()."""
        import threading

        pdf = self.caminho_pdf.get() if hasattr(self, "caminho_pdf") else ""
        if not pdf:
            from tkinter import messagebox
            messagebox.showwarning("Atenção", "Selecione um arquivo PDF primeiro.")
            return

        self._btn_autores_atarde.config(text="⏳  Extraindo...", state="disabled")
        self._lbl_status.config(text="Extraindo autores A TARDE...", fg=CORES["texto2"])
        self._lbl_status_ico.config(text="⟳", fg=CORES["amarelo"])
        # Navega para a aba de resultado
        self._abas.select(self._aba_autores_atarde)

        def _run():
            try:
                import fitz
                doc = fitz.open(pdf)
                texto_completo = "\n".join(
                    doc[i].get_text("text") or "" for i in range(len(doc))
                )
                doc.close()
                resultado = extrair_autores_ataarde(texto_completo)
            except Exception as exc:
                resultado = {"total_autores": 0, "autores_unicos": [],
                             "por_categoria": {}, "_erro": str(exc)}
            self.after(0, lambda: self._mostrar_autores_atarde(resultado))

        threading.Thread(target=_run, daemon=True).start()

    def _mostrar_autores_atarde(self, resultado):
        """Renderiza os resultados de extrair_autores_ataarde() na aba dedicada."""
        self._btn_autores_atarde.config(text="⬛  Autores A TARDE", state="normal")
        self._limpar_frame(self._frame_autores_atarde)

        erro = resultado.get("_erro")
        if erro:
            self._lbl_status.config(text=f"Erro: {erro}", fg=CORES["vermelho"])
            self._lbl_status_ico.config(text="✕", fg=CORES["vermelho"])
            self._mostrar_toast(f"Erro: {erro}", CORES["vermelho"])
            tk.Label(self._frame_autores_atarde,
                     text=f"Erro: {erro}",
                     font=FONTE_LABEL, fg=CORES["vermelho"],
                     bg=CORES["painel"]).pack(pady=20)
            return

        total = resultado.get("total_autores", 0)
        self._lbl_status.config(
            text=f"{total} autor(es) A TARDE encontrado(s)",
            fg=CORES["ciano"])
        self._lbl_status_ico.config(text="✔", fg=CORES["ciano"])
        self._mostrar_toast(f"⬛  {total} autor(es) A TARDE", CORES["ciano"], 2500)

        autores_unicos = resultado.get("autores_unicos", [])
        por_categoria = resultado.get("por_categoria", {})
        paginas_por_autor = resultado.get("paginas_por_autor", {})

        # ── Cabeçalho resumo ─────────────────────────────────────────────
        cab = tk.Frame(self._frame_autores_atarde, bg=CORES["painel2"],
                       pady=8, padx=12)
        cab.pack(fill="x", padx=8, pady=(0, 8))
        tk.Label(cab, text=f"Total de autores identificados: {total}",
                 font=("Segoe UI", 10, "bold"),
                 fg=CORES["verde"], bg=CORES["painel2"]).pack(side="left", anchor="w")

        def _copiar_todos():
            texto = "\n".join(autores_unicos)
            self.clipboard_clear()
            self.clipboard_append(texto)
            self._mostrar_toast("✓  Lista copiada!", CORES["verde"], 2000)

        btn_copiar = tk.Button(cab, text="Copiar lista",
                               font=FONTE_PEQUENA, fg=CORES["texto2"],
                               bg=CORES["painel3"], relief="flat", cursor="hand2",
                               padx=10, pady=3, activebackground=CORES["painel2"],
                               command=_copiar_todos)
        btn_copiar.pack(side="right", anchor="e")
        _btn_hover(btn_copiar, CORES["painel3"], CORES["azul"])

        # ── Seções por categoria ─────────────────────────────────────────
        COR_CAT = {
            "Jornalistas / Redatores": CORES["azul"],
            "Colunistas":              CORES["lilas"],
            "Políticos / Opinião":     CORES["amarelo"],
            "Leitores":                CORES["ciano"],
        }

        for cat, nomes in por_categoria.items():
            if not nomes:
                continue
            sec = tk.Frame(self._frame_autores_atarde, bg=CORES["painel2"],
                           pady=8, padx=12)
            sec.pack(fill="x", padx=8, pady=(0, 6))

            cor = COR_CAT.get(cat, CORES["texto2"])
            tk.Label(sec, text=f"{cat}  ({len(nomes)})",
                     font=("Segoe UI", 9, "bold"),
                     fg=cor, bg=CORES["painel2"]).pack(anchor="w", pady=(0, 4))

            for nome in nomes:
                pags = paginas_por_autor.get(nome, [])
                sufixo = "  (p. " + ", ".join(str(p) for p in pags) + ")" if pags else ""
                linha_nome = tk.Frame(sec, bg=CORES["painel2"])
                linha_nome.pack(fill="x")
                tk.Label(linha_nome, text=f"  • {nome}",
                         font=FONTE_LABEL, fg=CORES["texto"],
                         bg=CORES["painel2"], anchor="w"
                         ).pack(side="left")
                if sufixo:
                    tk.Label(linha_nome, text=sufixo,
                             font=FONTE_PEQUENA, fg=CORES["texto2"],
                             bg=CORES["painel2"], anchor="w"
                             ).pack(side="left")

        # ── Lista completa (alfabética) ──────────────────────────────────
        if autores_unicos:
            sec_todos = tk.Frame(self._frame_autores_atarde, bg=CORES["painel2"],
                                 pady=8, padx=12)
            sec_todos.pack(fill="x", padx=8, pady=(0, 6))
            tk.Label(sec_todos, text=f"Lista completa  ({len(autores_unicos)})",
                     font=("Segoe UI", 9, "bold"),
                     fg=CORES["texto2"], bg=CORES["painel2"]).pack(anchor="w", pady=(0, 4))
            # Exibe em grade de 2 colunas
            grade = tk.Frame(sec_todos, bg=CORES["painel2"])
            grade.pack(fill="x")
            for i, nome in enumerate(autores_unicos):
                col = i % 2
                pags = paginas_por_autor.get(nome, [])
                sufixo = "  (p. " + ", ".join(str(p) for p in pags) + ")" if pags else ""
                celula = tk.Frame(grade, bg=CORES["painel2"])
                celula.grid(row=i // 2, column=col, sticky="w", padx=(0, 20))
                tk.Label(celula, text=f"  {nome}",
                         font=FONTE_PEQUENA, fg=CORES["texto2"],
                         bg=CORES["painel2"], anchor="w"
                         ).pack(side="left")
                if sufixo:
                    tk.Label(celula, text=sufixo,
                             font=FONTE_PEQUENA, fg=CORES["texto3"],
                             bg=CORES["painel2"], anchor="w"
                             ).pack(side="left")

    def _montar_aba_buscar_autor(self, parent):
        """Monta a aba de busca de autores no PDF (3 modos)."""
        # ── Seletor de modo ───────────────────────────────────────────────
        topo = tk.Frame(parent, bg=CORES["fundo"])
        topo.pack(fill="x", padx=16, pady=(14, 4))

        tk.Label(topo, text="Modo de busca:",
                 font=FONTE_LABEL, fg=CORES["texto2"],
                 bg=CORES["fundo"]).pack(anchor="w")

        self._modo_busca_autor = tk.StringVar(value="unico")
        barra_modos = tk.Frame(topo, bg=CORES["fundo"])
        barra_modos.pack(fill="x", pady=(4, 8))
        for label, valor in (("Nome único", "unico"),
                             ("Lista de nomes", "lista"),
                             ("Arquivo .txt", "arquivo")):
            tk.Radiobutton(barra_modos, text=label, variable=self._modo_busca_autor,
                           value=valor, font=FONTE_LABEL,
                           fg=CORES["texto"], bg=CORES["fundo"],
                           selectcolor=CORES["painel2"], activebackground=CORES["fundo"],
                           command=self._alternar_modo_busca_autor
                           ).pack(side="left", padx=(0, 16))

        # ── Painéis de entrada (um por modo, só um visível por vez) ───────
        self._paineis_busca_autor = {}

        # Modo 1 — nome único
        p1 = tk.Frame(topo, bg=CORES["fundo"])
        linha1 = tk.Frame(p1, bg=CORES["fundo"])
        linha1.pack(fill="x")
        self._var_busca_autor = tk.StringVar()
        entry = tk.Entry(linha1, textvariable=self._var_busca_autor,
                         font=FONTE_LABEL, bg=CORES["painel2"],
                         fg=CORES["texto"], insertbackground=CORES["texto"],
                         relief="flat", bd=0)
        entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        entry.bind("<Return>", lambda e: self._executar_busca_autor())
        tk.Button(linha1, text="Buscar", font=FONTE_LABEL, bg=CORES["azul"],
                  fg="white", relief="flat", bd=0, cursor="hand2",
                  padx=14, pady=4, command=self._executar_busca_autor
                  ).pack(side="left")
        self._paineis_busca_autor["unico"] = p1

        # Modo 2 — lista de nomes
        p2 = tk.Frame(topo, bg=CORES["fundo"])
        tk.Label(p2, text="Um nome por linha:", font=FONTE_PEQUENA,
                 fg=CORES["texto3"], bg=CORES["fundo"]).pack(anchor="w")
        self._txt_lista_nomes = tk.Text(p2, height=5, font=FONTE_LABEL,
                                        bg=CORES["painel2"], fg=CORES["texto"],
                                        insertbackground=CORES["texto"],
                                        relief="flat", bd=0)
        self._txt_lista_nomes.pack(fill="x", pady=(2, 6))
        tk.Button(p2, text="Buscar lista", font=FONTE_LABEL, bg=CORES["azul"],
                  fg="white", relief="flat", bd=0, cursor="hand2",
                  padx=14, pady=4, command=self._executar_busca_lista
                  ).pack(anchor="w")
        self._paineis_busca_autor["lista"] = p2

        # Modo 3 — arquivo .txt
        p3 = tk.Frame(topo, bg=CORES["fundo"])
        linha3 = tk.Frame(p3, bg=CORES["fundo"])
        linha3.pack(fill="x")
        self._lbl_arquivo_nomes = tk.Label(linha3, text="Nenhum arquivo selecionado",
                                           font=FONTE_PEQUENA, fg=CORES["texto3"],
                                           bg=CORES["painel2"], anchor="w",
                                           padx=8, pady=6)
        self._lbl_arquivo_nomes.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._caminho_arquivo_nomes = tk.StringVar()
        tk.Button(linha3, text="Selecionar .txt", font=FONTE_LABEL,
                  bg=CORES["painel2"], fg=CORES["texto"], relief="flat", bd=0,
                  cursor="hand2", padx=10, pady=4,
                  command=self._selecionar_arquivo_nomes).pack(side="left", padx=(0, 8))
        tk.Button(linha3, text="Buscar arquivo", font=FONTE_LABEL, bg=CORES["azul"],
                  fg="white", relief="flat", bd=0, cursor="hand2",
                  padx=14, pady=4, command=self._executar_busca_arquivo
                  ).pack(side="left")
        self._paineis_busca_autor["arquivo"] = p3

        # Exibe painel inicial
        p1.pack(fill="x", pady=(0, 4))

        # ── Status ────────────────────────────────────────────────────────
        self._lbl_status_busca_autor = tk.Label(
            topo, text="", font=FONTE_PEQUENA,
            fg=CORES["texto3"], bg=CORES["fundo"])
        self._lbl_status_busca_autor.pack(anchor="w", pady=(4, 0))

        # ── Área de resultados (rolável) ──────────────────────────────────
        self._canvas_busca_autor, self._frame_busca_autor = \
            self._criar_area_rolavel(parent)

    def _alternar_modo_busca_autor(self):
        """Troca o painel visível conforme o modo selecionado."""
        for modo, painel in self._paineis_busca_autor.items():
            if modo == self._modo_busca_autor.get():
                painel.pack(fill="x", pady=(0, 4))
            else:
                painel.pack_forget()

    def _pdf_carregado(self) -> str:
        """Retorna o caminho do PDF ou '' se nenhum carregado."""
        return self.caminho_pdf.get() if hasattr(self, "caminho_pdf") else ""

    def _executar_busca_autor(self):
        """Modo 1 — busca um único nome."""
        import threading
        pdf = self._pdf_carregado()
        if not pdf:
            self._lbl_status_busca_autor.config(
                text="Nenhum PDF carregado.", fg=CORES["amarelo"]); return
        nome = self._var_busca_autor.get().strip()
        if not nome:
            self._lbl_status_busca_autor.config(
                text="Digite um nome para buscar.", fg=CORES["amarelo"]); return
        self._lbl_status_busca_autor.config(text="Buscando...", fg=CORES["texto3"])
        self._limpar_frame(self._frame_busca_autor)
        def _run():
            res = buscar_autor_no_pdf(pdf, nome)
            self.after(0, lambda: self._mostrar_resultados_busca_autor(nome, res))
        threading.Thread(target=_run, daemon=True).start()

    def _executar_busca_lista(self):
        """Modo 2 — busca lista de nomes digitada no Text widget."""
        import threading
        pdf = self._pdf_carregado()
        if not pdf:
            self._lbl_status_busca_autor.config(
                text="Nenhum PDF carregado.", fg=CORES["amarelo"]); return
        raw = self._txt_lista_nomes.get("1.0", "end")
        nomes = [l.strip() for l in raw.splitlines() if l.strip() and not l.strip().startswith("#")]
        if not nomes:
            self._lbl_status_busca_autor.config(
                text="Nenhum nome na lista.", fg=CORES["amarelo"]); return
        self._lbl_status_busca_autor.config(
            text=f"Buscando {len(nomes)} nome(s)...", fg=CORES["texto3"])
        self._limpar_frame(self._frame_busca_autor)
        def _run():
            res = search_multiple_names(nomes, pdf)
            self.after(0, lambda: self._mostrar_resultados_multiplos(res))
        threading.Thread(target=_run, daemon=True).start()

    def _selecionar_arquivo_nomes(self):
        """Abre diálogo para escolher arquivo .txt com nomes."""
        from tkinter import filedialog
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo de nomes",
            filetypes=[("Arquivo de texto", "*.txt"), ("Todos", "*.*")])
        if caminho:
            self._caminho_arquivo_nomes.set(caminho)
            self._lbl_arquivo_nomes.config(
                text=caminho, fg=CORES["texto"])

    def _executar_busca_arquivo(self):
        """Modo 3 — busca nomes lidos de arquivo .txt."""
        import threading
        pdf = self._pdf_carregado()
        if not pdf:
            self._lbl_status_busca_autor.config(
                text="Nenhum PDF carregado.", fg=CORES["amarelo"]); return
        arq = self._caminho_arquivo_nomes.get()
        if not arq:
            self._lbl_status_busca_autor.config(
                text="Selecione um arquivo .txt primeiro.", fg=CORES["amarelo"]); return
        try:
            nomes = load_names_from_txt(arq)
        except IOError as exc:
            self._lbl_status_busca_autor.config(text=str(exc), fg=CORES["vermelho"]); return
        if not nomes:
            self._lbl_status_busca_autor.config(
                text="Arquivo sem nomes válidos.", fg=CORES["amarelo"]); return
        self._lbl_status_busca_autor.config(
            text=f"Buscando {len(nomes)} nome(s)...", fg=CORES["texto3"])
        self._limpar_frame(self._frame_busca_autor)
        def _run():
            res = search_multiple_names(nomes, pdf)
            self.after(0, lambda: self._mostrar_resultados_multiplos(res))
        threading.Thread(target=_run, daemon=True).start()

    def _mostrar_resultados_busca_autor(self, nome, resultados):
        """Renderiza cards de resultado para busca de nome único."""
        self._limpar_frame(self._frame_busca_autor)
        if not resultados:
            self._lbl_status_busca_autor.config(
                text=f'Autor "{nome}" não encontrado no PDF.',
                fg=CORES["amarelo"])
            tk.Label(self._frame_busca_autor,
                     text=f'Nenhuma ocorrência de "{nome}" no PDF.',
                     font=FONTE_LABEL, fg=CORES["texto3"],
                     bg=CORES["painel"]).pack(pady=40)
            return
        paginas_str = ", ".join(str(r["pagina"]) for r in resultados)
        total_trechos = sum(len(r["trechos"]) for r in resultados)
        self._lbl_status_busca_autor.config(
            text=(f'"{nome}" — {len(resultados)} página(s): '
                  f'{paginas_str}  —  {total_trechos} ocorrência(s)'),
            fg=CORES["verde"])
        for item in resultados:
            card = tk.Frame(self._frame_busca_autor,
                            bg=CORES["painel2"], pady=10, padx=12)
            card.pack(fill="x", padx=8, pady=(0, 6))
            tk.Label(card, text=f"Página {item['pagina']}",
                     font=FONTE_LABEL, fg=CORES["azul"],
                     bg=CORES["painel2"]).pack(anchor="w")
            for trecho in item["trechos"]:
                tk.Label(card, text=f"  …{trecho}…",
                         font=FONTE_PEQUENA, fg=CORES["texto2"],
                         bg=CORES["painel2"], wraplength=700,
                         justify="left", anchor="w").pack(fill="x", pady=(2, 0))

    def _mostrar_resultados_multiplos(self, resultados: dict):
        """Renderiza tabela resumo + cards para busca de múltiplos nomes."""
        self._limpar_frame(self._frame_busca_autor)
        encontrados = {n: p for n, p in resultados.items() if p}
        nao_encontrados = [n for n, p in resultados.items() if not p]
        total = len(resultados)
        self._lbl_status_busca_autor.config(
            text=(f"{len(encontrados)}/{total} nome(s) encontrado(s)"),
            fg=CORES["verde"] if encontrados else CORES["amarelo"])

        # ── Tabela resumo ────────────────────────────────────────────────
        tabela = tk.Frame(self._frame_busca_autor, bg=CORES["painel2"],
                          pady=8, padx=12)
        tabela.pack(fill="x", padx=8, pady=(0, 8))
        tk.Label(tabela, text="RESUMO", font=("Segoe UI", 9, "bold"),
                 fg=CORES["texto2"], bg=CORES["painel2"]).pack(anchor="w", pady=(0, 4))
        for nome, paginas in resultados.items():
            linha = tk.Frame(tabela, bg=CORES["painel2"])
            linha.pack(fill="x", pady=1)
            if paginas:
                pags_str = ", ".join(str(p) for p in paginas)
                info = f"Encontrado  —  página(s): {pags_str}"
                cor = CORES["verde"]
            else:
                info = "Não encontrado"
                cor = CORES["texto3"]
            tk.Label(linha, text=f"  {nome}", font=FONTE_LABEL,
                     fg=CORES["texto"], bg=CORES["painel2"],
                     width=28, anchor="w").pack(side="left")
            tk.Label(linha, text=info, font=FONTE_PEQUENA,
                     fg=cor, bg=CORES["painel2"],
                     anchor="w").pack(side="left", fill="x", expand=True)

        # ── Cards detalhados apenas para os encontrados ──────────────────
        for nome, paginas in encontrados.items():
            ocorrencias = buscar_autor_no_pdf(self._pdf_carregado(), nome)
            card = tk.Frame(self._frame_busca_autor,
                            bg=CORES["painel2"], pady=10, padx=12)
            card.pack(fill="x", padx=8, pady=(0, 6))
            tk.Label(card, text=nome, font=FONTE_LABEL,
                     fg=CORES["azul"], bg=CORES["painel2"]).pack(anchor="w")
            for item in ocorrencias:
                tk.Label(card, text=f"  Página {item['pagina']}",
                         font=("Segoe UI", 8, "bold"), fg=CORES["texto2"],
                         bg=CORES["painel2"]).pack(anchor="w", pady=(4, 0))
                for trecho in item["trechos"]:
                    tk.Label(card, text=f"    …{trecho}…",
                             font=FONTE_PEQUENA, fg=CORES["texto2"],
                             bg=CORES["painel2"], wraplength=680,
                             justify="left", anchor="w").pack(fill="x", pady=(1, 0))

    def _mostrar_autores(self, materias):
        self._limpar_frame(self._frame_autores)

        com_autor = [m for m in materias if m["autor_info"].get("autor")]
        sem_autor = [m for m in materias if not m["autor_info"].get("autor")]

        # Resumo no topo
        resumo = tk.Frame(self._frame_autores, bg=CORES["painel2"])
        resumo.pack(fill="x", pady=(0, 4))
        tk.Label(resumo,
                 text=f"  {len(com_autor)} autor(es) identificado(s)  |  "
                      f"{len(sem_autor)} sem autor",
                 font=("Segoe UI", 9, "bold"),
                 fg=CORES["verde"], bg=CORES["painel2"], pady=8).pack(side="left")

        # Agrupamento por autor
        autores_agrupados = {}
        for m in com_autor:
            nome = m["autor_info"]["autor"]
            autores_agrupados.setdefault(nome, []).append(m)

        # Tabela-resumo de autores
        if autores_agrupados:
            frame_tabela = tk.Frame(self._frame_autores, bg=CORES["painel3"],
                                    highlightbackground=CORES["borda"],
                                    highlightthickness=1)
            frame_tabela.pack(fill="x", padx=8, pady=(4, 8))

            tk.Label(frame_tabela, text="  RESUMO POR AUTOR",
                     font=("Segoe UI", 9, "bold"),
                     fg=CORES["lilas"], bg=CORES["painel3"],
                     pady=6).pack(anchor="w")

            for nome, lista in sorted(autores_agrupados.items(),
                                       key=lambda x: len(x[1]), reverse=True):
                tipo = lista[0]["autor_info"]["tipo"]
                tipo_map = {"reporter": "Reporter", "redacao": "Redacao",
                            "articulista": "Articulista", "desconhecido": ""}
                tipo_label = tipo_map.get(tipo, tipo)
                paginas = sorted(set(m["pagina"] for m in lista))
                pags_str = ", ".join(str(p) for p in paginas)

                lr = tk.Frame(frame_tabela, bg=CORES["painel3"])
                lr.pack(fill="x", padx=8, pady=2)
                tk.Label(lr, text=f"{nome}",
                         font=("Consolas", 9, "bold"),
                         fg=CORES["texto"], bg=CORES["painel3"],
                         anchor="w").pack(side="left")
                tk.Label(lr, text=f"  ({tipo_label})",
                         font=FONTE_BADGE, fg=CORES["ciano"],
                         bg=CORES["painel3"]).pack(side="left")
                tk.Label(lr, text=f"  {len(lista)}x  |  Pag: {pags_str}",
                         font=FONTE_MONO_P, fg=CORES["texto2"],
                         bg=CORES["painel3"]).pack(side="left", padx=(6, 0))

                # Colaboradores
                for m in lista:
                    colabs = m["autor_info"].get("colaboradores", [])
                    if colabs:
                        tk.Label(lr,
                                 text=f"  + {', '.join(colabs)}",
                                 font=FONTE_MONO_P, fg=CORES["amarelo"],
                                 bg=CORES["painel3"]).pack(side="left")
                        break

            tk.Frame(frame_tabela, bg=CORES["borda"], height=1).pack(
                fill="x", padx=8, pady=(4, 6))

        # Cards detalhados de cada matéria
        for m in materias:
            self._card_autor(m)

    def _card_autor(self, dados):
        ai = dados["autor_info"]
        tem_autor = ai.get("autor") is not None

        card = tk.Frame(self._frame_autores, bg=CORES["painel2"],
                        highlightbackground=CORES["borda"], highlightthickness=1)
        card.pack(fill="x", padx=8, pady=3)

        # Linha topo: página + autor
        topo = tk.Frame(card, bg=CORES["painel2"])
        topo.pack(fill="x", padx=10, pady=(8, 4))

        tk.Label(topo, text=f"Pag. {dados['pagina']}",
                 font=("Segoe UI", 10, "bold"),
                 fg=CORES["texto"], bg=CORES["painel2"]).pack(side="left")

        if tem_autor:
            tipo_map = {"reporter": "Reporter", "redacao": "Redacao",
                        "articulista": "Articulista", "desconhecido": ""}
            tipo_label = tipo_map.get(ai["tipo"], "")
            confianca = ai.get("confianca", 0)
            cor_conf = (CORES["verde"] if confianca >= 0.85
                        else CORES["amarelo"] if confianca >= 0.65
                        else CORES["texto2"])

            tk.Label(topo, text=f"  {ai['autor']}",
                     font=("Consolas", 10, "bold"),
                     fg=cor_conf, bg=CORES["painel2"]).pack(side="left")
            tk.Label(topo, text=f"  {tipo_label}  {confianca:.0%}",
                     font=FONTE_BADGE, fg=CORES["texto2"],
                     bg=CORES["painel3"], padx=4).pack(side="left", padx=(4, 0))

            if ai.get("colaboradores"):
                tk.Label(topo, text=f"  + {', '.join(ai['colaboradores'])}",
                         font=FONTE_BADGE, fg=CORES["amarelo"],
                         bg=CORES["painel2"]).pack(side="left", padx=(6, 0))
        else:
            tk.Label(topo, text="  Autor nao identificado",
                     font=FONTE_BADGE, fg=CORES["texto3"],
                     bg=CORES["painel2"]).pack(side="left")

        # Preview da região
        preview = gerar_preview_anuncio(
            self.caminho_pdf.get(), dados["pagina"],
            dados["x_cm"], dados["y_cm"],
            dados["w_cm"], dados["h_cm"],
            largura_thumb=300)
        if preview:
            try:
                foto = ImageTk.PhotoImage(preview)
                self._thumbnails.append(foto)
                tk.Label(card, image=foto,
                         bg=CORES["painel2"]).pack(padx=10, pady=(2, 4))
            except Exception:
                pass

        # Trecho de texto
        if dados.get("trecho"):
            txt_frame = tk.Frame(card, bg=CORES["painel3"],
                                 highlightbackground=CORES["borda"],
                                 highlightthickness=1)
            txt_frame.pack(fill="x", padx=10, pady=(0, 4))
            tk.Label(txt_frame, text=dados["trecho"],
                     font=FONTE_MONO_P, fg=CORES["texto2"],
                     bg=CORES["painel3"], anchor="w",
                     justify="left", wraplength=500,
                     padx=6, pady=4).pack(fill="x")

        # Métricas de posição
        metricas = tk.Frame(card, bg=CORES["painel2"])
        metricas.pack(fill="x", padx=10, pady=(0, 2))

        def ml(tag, cor, val):
            l = tk.Frame(metricas, bg=CORES["painel2"])
            l.pack(fill="x", pady=1)
            tk.Label(l, text=tag, font=FONTE_BADGE,
                     fg=cor, bg=CORES["painel3"], padx=2).pack(side="left")
            tk.Label(l, text=val, font=FONTE_MONO_P,
                     fg=CORES["texto2"], bg=CORES["painel2"],
                     padx=6).pack(side="left")

        ml(" POS  ", CORES["amarelo"],
           f"X {dados['x_cm']:.3f} cm  Y {dados['y_cm']:.3f} cm")
        ml(" TAM  ", CORES["ciano"],
           f"L {dados['w_cm']:.3f} cm  x  A {dados['h_cm']:.3f} cm")

        # Colunagem
        nome_jornal = self.jornal_ativo.get()
        jornal = JORNAIS_CADASTRADOS.get(nome_jornal)
        if jornal and dados["w_cm"] and dados["h_cm"]:
            info = calcular_info_colunas(
                dados["w_cm"], dados.get("x_cm") or 0.0, jornal)
            sinal = "+" if info["desvio"] >= 0 else ""
            ml(" COL  ", CORES["lilas"],
               f"Col {info['col_ini']} a {info['col_fim']}  "
               f"|  {info['num_col']} col ({info['cols_exato']:.3f})  "
               f"|  alt {dados['h_cm']:.3f} cm")
            ml(" DIM  ", CORES["ciano"],
               f"Padrao {info['larg_padrao']:.3f} cm  "
               f"|  Medido {dados['w_cm']:.3f} cm  "
               f"|  Δ {sinal}{info['desvio']:.3f} cm")

        if ai.get("raw_autor_linha"):
            ml("LINHA", CORES["texto3"],
               f"\"{ai['raw_autor_linha']}\"")

        # Botão de marcar no PDF
        tk.Button(card,
            text="Baixar PDF com marcacao",
            font=("Segoe UI", 8, "bold"),
            fg="#ffffff", bg=CORES["verde"],
            activebackground="#22a870",
            activeforeground="#ffffff",
            relief="flat", cursor="hand2",
            padx=8, pady=5,
            command=lambda d=dados: self._baixar_pdf_marcado([d])
        ).pack(fill="x", padx=10, pady=(4, 8))


if __name__ == "__main__":
    App().mainloop()
