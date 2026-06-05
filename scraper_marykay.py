#!/usr/bin/env python3
"""
Scraper de Produtos Mary Kay Brasil
Fonte  : https://loja.marykay.com.br
Estrat.: Sitemap XML → API pública VTEX → fallback scraping HTML
"""

import argparse
import json
import logging
import re
import time
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# Configurações globais
# ─────────────────────────────────────────────────────────────────────────────

BASE_URL    = "https://loja.marykay.com.br"
SITEMAP_URL = f"{BASE_URL}/sitemap.xml"

# Headers que simulam um navegador real para evitar bloqueios
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Cache-Control": "max-age=0",
}

HEADERS_JSON = {**HEADERS, "Accept": "application/json, */*;q=0.8"}

REQUEST_TIMEOUT = 25
REQUEST_DELAY   = 0.6   # segundos entre requisições
MAX_RETRIES     = 3

# URLs de teste/internas para excluir do scraping
URL_BLOCKLIST = [
    "produto-teste", "b8one", "-teste", "teste-", "/teste",
    "homolog", "staging", "sandbox", "dummy",
]

# Rotas bloqueadas (robots.txt ou não-produto)
ROUTE_BLOCKLIST = [
    "/busca", "/buscapagina", "/quick-view", "/checkout",
    "/cart", "/account", "/login", "/wishlist", "/orderPlaced",
    "/profile", "/sitemap", "/Em-Sintonia", "/em-sintonia",
    "/institucional", "/trabalhe-conosco", "/sustentabilidade",
    "/termos-de-uso", "/politica-de-privacidade",
    "/central-de-atendimento", "/fale-conosco",
    "/mapa-do-site",
]

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────

logger = logging.getLogger("scraper_marykay")


def setup_logging(log_file: str) -> None:
    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(level=logging.INFO, format=fmt)
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(logging.Formatter(fmt))
    logger.addHandler(fh)


# ─────────────────────────────────────────────────────────────────────────────
# HTTP helpers
# ─────────────────────────────────────────────────────────────────────────────

_session = requests.Session()
_session.headers.update(HEADERS)


def safe_get(
    url: str,
    params: Optional[dict] = None,
    json_mode: bool = False,
    extra_headers: Optional[dict] = None,
) -> Optional[requests.Response]:
    """GET com retries e backoff exponencial."""
    headers = HEADERS_JSON if json_mode else HEADERS
    if extra_headers:
        headers = {**headers, **extra_headers}

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(
                url,
                params=params,
                headers=headers,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )
            if resp.status_code == 200:
                return resp
            if resp.status_code in (429, 503, 502):
                wait = 2 ** attempt
                logger.warning("Rate-limit/error %d %s → aguardando %ds",
                               resp.status_code, url, wait)
                time.sleep(wait)
                continue
            logger.debug("HTTP %d para %s", resp.status_code, url)
            return None
        except requests.RequestException as exc:
            logger.warning("Tentativa %d/%d falhou para %s: %s",
                           attempt, MAX_RETRIES, url, exc)
            if attempt < MAX_RETRIES:
                time.sleep(2 ** attempt)
    return None


# ─────────────────────────────────────────────────────────────────────────────
# 1. Sitemap
# ─────────────────────────────────────────────────────────────────────────────

_SM_NS = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}


def read_sitemap_recursive(
    url: str,
    visited: Optional[set] = None,
    sitemap_rows: Optional[list] = None,
) -> list[str]:
    """Lê sitemap ou sitemap-index recursivamente. Retorna URLs de produto."""
    if visited is None:
        visited = set()
    if sitemap_rows is None:
        sitemap_rows = []

    if url in visited:
        return []
    visited.add(url)

    logger.info("Lendo sitemap: %s", url)
    resp = safe_get(url)
    if not resp:
        sitemap_rows.append({
            "sitemap_origem": url, "url_encontrada": "",
            "tipo_detectado": "erro", "incluida_sim_nao": "não",
            "motivo": "Falha ao baixar sitemap",
        })
        return []

    try:
        root = ET.fromstring(resp.content)
    except ET.ParseError as exc:
        logger.error("ParseError ao ler %s: %s", url, exc)
        sitemap_rows.append({
            "sitemap_origem": url, "url_encontrada": "",
            "tipo_detectado": "erro", "incluida_sim_nao": "não",
            "motivo": f"ParseError: {exc}",
        })
        return []

    tag = root.tag.lower()
    product_urls: list[str] = []

    if "sitemapindex" in tag:
        # Sitemap index: processar filhos
        children = root.findall("sm:sitemap", _SM_NS) or root.findall("sitemap")
        for child in children:
            loc_el = child.find("sm:loc", _SM_NS) or child.find("loc")
            if loc_el is not None and loc_el.text:
                child_url = loc_el.text.strip()
                product_urls.extend(
                    read_sitemap_recursive(child_url, visited, sitemap_rows)
                )
        return product_urls

    # Sitemap de URLs
    url_els = root.findall("sm:url", _SM_NS) or root.findall("url")
    for url_el in url_els:
        loc_el = url_el.find("sm:loc", _SM_NS) or url_el.find("loc")
        if loc_el is None or not loc_el.text:
            continue
        loc = loc_el.text.strip()
        tipo, incluida, motivo = _classify_sitemap_url(loc)
        sitemap_rows.append({
            "sitemap_origem": url, "url_encontrada": loc,
            "tipo_detectado": tipo, "incluida_sim_nao": incluida,
            "motivo": motivo,
        })
        if incluida == "sim":
            product_urls.append(loc)

    return product_urls


def _classify_sitemap_url(url: str) -> tuple[str, str, str]:
    lower = url.lower()
    for blocked in ROUTE_BLOCKLIST:
        if blocked.lower() in lower:
            return "rota-bloqueada", "não", f"Rota bloqueada: {blocked}"
    for test_word in URL_BLOCKLIST:
        if test_word in lower:
            return "teste", "não", f"URL de teste: {test_word}"
    if lower.endswith("/p") or "/p?" in lower:
        return "produto", "sim", "URL de produto (/p)"
    return "outro", "não", "Não é URL de produto (/p)"


# ─────────────────────────────────────────────────────────────────────────────
# 2. Normalização e slug
# ─────────────────────────────────────────────────────────────────────────────

def normalize_product_url(url: str) -> str:
    return url.split("?")[0].split("#")[0].rstrip("/")


def extract_slug_from_product_url(url: str) -> str:
    """
    https://loja.marykay.com.br/batom-cremoso-mary-kay/p
    → batom-cremoso-mary-kay
    """
    url = normalize_product_url(url)
    path = url.replace(BASE_URL, "").strip("/")
    if path.endswith("/p"):
        path = path[:-2]
    parts = [p for p in path.strip("/").split("/") if p and p != "p"]
    return parts[-1] if parts else path


# ─────────────────────────────────────────────────────────────────────────────
# 3. API VTEX — múltiplos endpoints
# ─────────────────────────────────────────────────────────────────────────────

def fetch_vtex_by_slug(slug: str) -> Optional[list[dict]]:
    """
    Tenta a API VTEX catalog_system com o slug do produto.
    Endpoint: /api/catalog_system/pub/products/search/{slug}
    """
    url = f"{BASE_URL}/api/catalog_system/pub/products/search/{slug}"
    resp = safe_get(url, params={"_from": "0", "_to": "9"}, json_mode=True)
    if resp:
        try:
            data = resp.json()
            if isinstance(data, list) and data:
                return data
        except ValueError:
            pass
    return None


def fetch_vtex_by_linktext(slug: str) -> Optional[list[dict]]:
    """
    Alternativa: filtro por linkText via query parameter.
    """
    url = f"{BASE_URL}/api/catalog_system/pub/products/search"
    resp = safe_get(
        url,
        params={"fq": f"productLinkText:{slug}", "_from": "0", "_to": "9"},
        json_mode=True,
    )
    if resp:
        try:
            data = resp.json()
            if isinstance(data, list) and data:
                return data
        except ValueError:
            pass
    return None


def fetch_vtex_by_sku_search(slug: str) -> Optional[list[dict]]:
    """
    Busca via fullText (último recurso na API).
    """
    url = f"{BASE_URL}/api/catalog_system/pub/products/search"
    # Converter slug em palavras para busca
    query = slug.replace("-", " ")
    resp = safe_get(
        url,
        params={"ft": query, "_from": "0", "_to": "4"},
        json_mode=True,
    )
    if resp:
        try:
            data = resp.json()
            if isinstance(data, list) and data:
                return data
        except ValueError:
            pass
    return None


def fetch_vtex_product_page_json(product_url: str) -> Optional[list[dict]]:
    """
    Tenta extrair __STATE__ (JSON hidratado) da página HTML do produto.
    O VTEX IO injeta os dados do produto no HTML como __STATE__.
    """
    resp = safe_get(product_url)
    if not resp:
        return None

    soup = BeautifulSoup(resp.text, "lxml")

    # Estratégia 1: script com __STATE__
    for script in soup.find_all("script"):
        src = script.string or ""
        if "__STATE__" in src:
            m = re.search(r"window\.__STATE__\s*=\s*(\{.+?\});?\s*</script>",
                          src, re.DOTALL)
            if not m:
                m = re.search(r"window\.__STATE__\s*=\s*(\{.+)", src, re.DOTALL)
            if m:
                try:
                    state = json.loads(m.group(1).rstrip("; \n"))
                    products = _extract_products_from_vtex_state(state)
                    if products:
                        return products
                except (json.JSONDecodeError, KeyError):
                    pass

    # Estratégia 2: script application/json com dados de produto
    for script in soup.find_all("script", {"type": "application/ld+json"}):
        try:
            data = json.loads(script.string or "{}")
            if data.get("@type") == "Product":
                return [_ld_json_to_vtex_format(data)]
        except (json.JSONDecodeError, AttributeError):
            pass

    return None


def _extract_products_from_vtex_state(state: dict) -> list[dict]:
    """Extrai produtos do __STATE__ do VTEX IO."""
    products = []
    for key, value in state.items():
        if not isinstance(value, dict):
            continue
        # Procurar nós com productName (formato típico do VTEX IO)
        if "productName" in value and "items" in value:
            products.append(value)
    return products


def _ld_json_to_vtex_format(ld: dict) -> dict:
    """Converte JSON-LD de produto para o formato mínimo VTEX esperado."""
    product_name = ld.get("name", "")
    offers = ld.get("offers", {})
    items_raw = ld.get("hasVariant") or [ld]
    items = []
    for v in items_raw:
        items.append({
            "itemId": v.get("sku", ""),
            "name": v.get("name", "").replace(product_name, "").strip(" -"),
            "nameComplete": v.get("name", ""),
            "referenceId": [{"Value": v.get("sku", "")}],
            "images": [],
        })
    return {
        "productName": product_name,
        "linkText": ld.get("url", "").split("/")[-1].replace("/p", ""),
        "categories": [ld.get("category", "")],
        "items": items,
    }


def fetch_vtex_product(product_url: str, slug: str) -> tuple[Optional[list[dict]], str]:
    """
    Tenta todos os métodos disponíveis para obter dados do produto.
    Retorna (lista_de_produtos, método_utilizado).
    """
    # 1. API direta pelo slug
    products = fetch_vtex_by_slug(slug)
    if products:
        return products, f"VTEX /search/{slug}"
    time.sleep(0.3)

    # 2. filtro por linkText
    products = fetch_vtex_by_linktext(slug)
    if products:
        return products, f"VTEX ?fq=productLinkText:{slug}"
    time.sleep(0.3)

    # 3. Busca por texto livre
    products = fetch_vtex_by_sku_search(slug)
    if products:
        return products, f"VTEX fullText:{slug}"
    time.sleep(0.3)

    # 4. Scraping da página HTML (fallback)
    products = fetch_vtex_product_page_json(product_url)
    if products:
        return products, "HTML page scraping"

    return None, "nenhum método funcionou"


# ─────────────────────────────────────────────────────────────────────────────
# 4. Seleção do melhor produto
# ─────────────────────────────────────────────────────────────────────────────

def choose_best_product_match(
    products: list[dict], slug: str
) -> tuple[Optional[dict], bool]:
    if not products:
        return None, False
    if len(products) == 1:
        return products[0], True

    slug_lower = slug.lower()
    # Match exato
    for p in products:
        link = (p.get("linkText") or "").lower()
        if link == slug_lower:
            return p, True
    # Match por prefixo
    for p in products:
        link = (p.get("linkText") or "").lower()
        if slug_lower.startswith(link) or link.startswith(slug_lower):
            return p, False
    # Melhor similaridade
    best = max(products, key=lambda p: _slug_sim(p.get("linkText", ""), slug))
    return best, False


def _slug_sim(a: str, b: str) -> float:
    a, b = a.lower(), b.lower()
    if not a or not b:
        return 0.0
    # Contar tokens em comum
    ta = set(a.split("-"))
    tb = set(b.split("-"))
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))


# ─────────────────────────────────────────────────────────────────────────────
# 5. Extração de SKUs
# ─────────────────────────────────────────────────────────────────────────────

def extract_skus_from_vtex_product(product: dict) -> list[dict]:
    product_name = (
        product.get("productName")
        or product.get("productTitle")
        or product.get("name")
        or ""
    ).strip()

    items = product.get("items") or []
    skus: list[dict] = []
    for item in items:
        sku_code  = _extract_sku_code(item)
        variation = _extract_variation(item, product_name)
        skus.append({
            "sku_code":     sku_code,
            "product_name": product_name,
            "variation":    variation,
            "raw_item":     item,
            "product":      product,
        })
    return skus


def _extract_sku_code(item: dict) -> str:
    # referenceId pode ser lista ou dict
    ref = item.get("referenceId")
    if isinstance(ref, list) and ref:
        val = ref[0].get("Value") or ref[0].get("value") or ""
        if val:
            return str(val).strip()
    if isinstance(ref, dict):
        val = ref.get("Value") or ref.get("value") or ""
        if val:
            return str(val).strip()
    for field in ("RefId", "refId", "EAN", "ean", "Ean"):
        val = item.get(field)
        if val:
            return str(val).strip()
    return str(item.get("itemId", "")).strip()


def _extract_variation(item: dict, product_name: str) -> str:
    # Campos diretos de variação (em ordem de preferência)
    variation_fields = [
        "Cor", "Tom", "Tamanho", "Volume",
        "sua cor", "Selecione a sua cor",
        "seu tom", "Selecione o seu tom",
        "Color", "Shade", "Size", "Fragrância",
    ]
    for field in variation_fields:
        val = item.get(field)
        if val:
            if isinstance(val, list):
                val = val[0] if val else ""
            cleaned = clean_variation(str(val).strip())
            if cleaned:
                return cleaned

    # Derivar de nameComplete subtraindo productName
    for key in ("nameComplete", "name"):
        full = (item.get(key) or "").strip()
        if full and full.lower() != product_name.lower():
            diff = _subtract_product_name(full, product_name)
            if diff:
                cleaned = clean_variation(diff)
                if cleaned:
                    return cleaned

    # Specifications / attributes
    for spec in (item.get("attributes") or []):
        if isinstance(spec, dict):
            v = spec.get("value") or spec.get("Value") or ""
            if v:
                cleaned = clean_variation(str(v).strip())
                if cleaned:
                    return cleaned

    return ""


def _subtract_product_name(full: str, base: str) -> str:
    """Remove o nome base da string completa e retorna o sufixo."""
    def normalize(s: str) -> str:
        return re.sub(r"\s+", " ", s.replace("®", "").strip())

    fn = normalize(full)
    bn = normalize(base)
    if fn.lower().startswith(bn.lower()):
        remainder = fn[len(bn):].strip(" -–—/|")
        return remainder
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# 6. Limpeza de variação
# ─────────────────────────────────────────────────────────────────────────────

_PREFIX_NUM_RE   = re.compile(r"^\d+\s*/\s*", re.UNICODE)
_ONLY_DIGITS_RE  = re.compile(r"^\d+$")
_INTERFACE_WORDS = {
    "adicionar", "sacola", "comprar", "selecione", "escolha",
    "ver mais", "veja mais", "disponível", "esgotado",
    "trabalhe conosco", "sustentabilidade", "termos de uso",
    "configuração de cookies", "buscar produtos",
    "adicionar à sacola", "ver detalhes",
}


def clean_variation(raw: str) -> str:
    if not raw:
        return ""
    val = raw.strip()
    # Remove prefixo "24 / Shell" → "Shell"
    val = _PREFIX_NUM_RE.sub("", val).strip()
    # Somente dígito → não é variação válida
    if _ONLY_DIGITS_RE.match(val):
        return ""
    # Palavra de interface → não é variação
    if val.lower() in _INTERFACE_WORDS:
        return ""
    return val


# ─────────────────────────────────────────────────────────────────────────────
# 7. Categoria
# ─────────────────────────────────────────────────────────────────────────────

_CATEGORY_MAP = [
    (
        ["batom", "brilho labial", "lápis labial", "labial", "lábios",
         "lip gloss", "lip liner", " lip "],
        "Lábios",
    ),
    (
        ["base ", "base time", "pó compacto", "pó translúcido", "blush",
         "bronzer", "corretivo", "primer", "cc cream", "contorno", "rosto",
         "iluminador", "corretor", "fixador de maquiagem"],
        "Rosto",
    ),
    (
        ["máscara de cílios", "cílios", "lápis para olhos", "sombra",
         "delineador", "sobrancelha", "olhos", " eye "],
        "Olhos",
    ),
    (
        ["perfume", "fragrância", "deo colônia", "colônia", "parfum",
         "eau de", "body splash", "splash"],
        "Fragrâncias",
    ),
    (
        ["gel de limpeza", "limpeza facial", "esfoliante facial",
         "sérum", "serum", "hidratante facial", "protetor solar facial",
         "microdermoabrasão", "timewise", "tônico", "demaquilante",
         "creme facial", "cuidados faciais", "máscara facial"],
        "Cuidados Faciais",
    ),
    (
        ["corporal", "creme para o corpo", "creme para mãos",
         "creme para as mãos", "hand cream", "loção corporal",
         "esfoliante corporal", "óleo corporal", "body"],
        "Cuidados Corporais",
    ),
    (
        ["kit", "presente", "necessaire", "nécessaire", "bolsa",
         "embalagem", "conjunto", "estojo", "coleção"],
        "Presentes",
    ),
    (
        ["esmalte", "nail", "unha"],
        "Unhas",
    ),
    (
        ["protetor solar", "fps", "fpf", "fator de proteção"],
        "Proteção Solar",
    ),
]


def classify_category(product: dict) -> str:
    # 1. Árvore de categorias da API
    cats = product.get("categories") or product.get("categoryTree") or []
    cat_text = " ".join(str(c).lower() for c in cats) if isinstance(cats, list) else ""
    result = _map_category(cat_text)
    if result != "Não classificado":
        return result

    # 2. Nome do produto
    name = (product.get("productName") or "").lower()
    result = _map_category(name)
    if result != "Não classificado":
        return result

    # 3. Department / brand text
    dept = (
        (product.get("department") or "")
        + " " + (product.get("brand") or "")
    ).lower()
    return _map_category(dept)


def _map_category(text: str) -> str:
    text = text.lower()
    for keywords, category in _CATEGORY_MAP:
        for kw in keywords:
            if kw.lower() in text:
                return category
    return "Não classificado"


# ─────────────────────────────────────────────────────────────────────────────
# 8. Construção das linhas da planilha
# ─────────────────────────────────────────────────────────────────────────────

def build_rows(
    product_url: str,
    product: dict,
    skus: list[dict],
    audit_rows: list[dict],
    match_exact: bool,
    method: str,
) -> list[dict]:
    rows: list[dict] = []
    category     = classify_category(product)
    product_name = (product.get("productName") or "").strip()

    variations_found: list[str] = []
    codes_found: list[str] = []
    seen: set = set()

    for sku in skus:
        code      = sku["sku_code"]
        base      = sku["product_name"] or product_name
        variation = sku["variation"]

        item_name = f"{base} - {variation}" if variation else base

        dedup_key = (code, item_name) if code else (base, variation)
        if dedup_key in seen:
            continue
        seen.add(dedup_key)

        rows.append({
            "SKU/Código":          code,
            "Nome do Item":        item_name,
            "Produto Base":        base,
            "Categoria":           category,
            "Variação/Cor/Tamanho": variation,
        })
        if variation:
            variations_found.append(variation)
        if code:
            codes_found.append(code)

    audit_rows.append({
        "URL do Produto":        product_url,
        "Produto Base":          product_name,
        "Categoria":             category,
        "Método usado":          method,
        "Quantidade de SKUs":    len(rows),
        "Variações encontradas": ", ".join(variations_found[:10]),
        "Códigos encontrados":   ", ".join(codes_found[:10]),
        "Observação":            ("" if match_exact
                                  else "Sem match exato de slug"),
    })
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# 9. Processamento de um produto
# ─────────────────────────────────────────────────────────────────────────────

def process_product_url(
    url: str,
    error_rows: list[dict],
    audit_rows: list[dict],
    debug_dir: Optional[Path],
) -> list[dict]:
    clean_url = normalize_product_url(url)
    slug      = extract_slug_from_product_url(clean_url)
    timestamp = datetime.now().isoformat()

    logger.info("Processando: %s  (slug=%s)", clean_url, slug)

    products, method = fetch_vtex_product(clean_url, slug)
    time.sleep(REQUEST_DELAY)

    if not products:
        error_rows.append({
            "URL":       clean_url,
            "Erro":      "Nenhum produto retornado por nenhum método",
            "Etapa":     "fetch_vtex_product",
            "Timestamp": timestamp,
        })
        return []

    product, match_exact = choose_best_product_match(products, slug)

    if not product:
        error_rows.append({
            "URL":       clean_url,
            "Erro":      "choose_best_product_match retornou None",
            "Etapa":     "choose_best_product_match",
            "Timestamp": timestamp,
        })
        return []

    # Dump de debug
    if debug_dir is not None:
        slug_safe = re.sub(r"[^\w\-]", "_", slug)
        debug_file = debug_dir / f"{slug_safe}.json"
        try:
            with open(debug_file, "w", encoding="utf-8") as f:
                json.dump(products, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    skus = extract_skus_from_vtex_product(product)

    if not skus:
        error_rows.append({
            "URL":       clean_url,
            "Erro":      "Produto sem items/SKUs",
            "Etapa":     "extract_skus_from_vtex_product",
            "Timestamp": timestamp,
        })
        return []

    full_method = f"{method} ({'match exato' if match_exact else 'match aprox.'})"
    return build_rows(clean_url, product, skus, audit_rows, match_exact, full_method)


# ─────────────────────────────────────────────────────────────────────────────
# 10. Salvamento dos arquivos
# ─────────────────────────────────────────────────────────────────────────────

COL_ORDER = [
    "ID", "SKU/Código", "Nome do Item", "Produto Base",
    "Categoria", "Variação/Cor/Tamanho",
]


def save_outputs(
    product_rows: list[dict],
    audit_rows: list[dict],
    error_rows: list[dict],
    sitemap_debug_rows: list[dict],
    prefix: str,
    dump_sitemap: bool,
) -> None:
    df_p = pd.DataFrame(product_rows)
    if not df_p.empty:
        df_p.insert(0, "ID", range(1, len(df_p) + 1))
    else:
        df_p = pd.DataFrame(columns=COL_ORDER)

    df_a = (pd.DataFrame(audit_rows) if audit_rows
            else pd.DataFrame(columns=[
                "URL do Produto", "Produto Base", "Categoria",
                "Método usado", "Quantidade de SKUs",
                "Variações encontradas", "Códigos encontrados", "Observação",
            ]))
    df_e = (pd.DataFrame(error_rows) if error_rows
            else pd.DataFrame(columns=["URL", "Erro", "Etapa", "Timestamp"]))

    xlsx_path = f"{prefix}.xlsx"
    csv_path  = f"{prefix}.csv"

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_p[COL_ORDER].to_excel(writer, sheet_name="Produtos", index=False)
        df_a.to_excel(writer, sheet_name="Auditoria", index=False)
        df_e.to_excel(writer, sheet_name="Erros", index=False)

    _format_excel(xlsx_path)
    df_p[COL_ORDER].to_csv(csv_path, index=False, encoding="utf-8-sig")

    logger.info("Excel salvo: %s", xlsx_path)
    logger.info("CSV salvo:   %s", csv_path)

    if dump_sitemap and sitemap_debug_rows:
        sitemap_csv = f"{prefix}_sitemap_debug.csv"
        pd.DataFrame(sitemap_debug_rows).to_csv(
            sitemap_csv, index=False, encoding="utf-8-sig"
        )
        logger.info("Sitemap debug: %s", sitemap_csv)


def _format_excel(xlsx_path: str) -> None:
    wb = load_workbook(xlsx_path)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        if ws.max_row < 1:
            continue
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        col_widths: dict[str, int] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.row == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                col_letter = get_column_letter(cell.column)
                length = len(str(cell.value or ""))
                col_widths[col_letter] = max(col_widths.get(col_letter, 10), length + 4)

        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = min(width, 60)

    wb.save(xlsx_path)


# ─────────────────────────────────────────────────────────────────────────────
# 11. Sumário no terminal
# ─────────────────────────────────────────────────────────────────────────────

def print_summary(
    product_rows: list[dict],
    error_rows: list[dict],
    total_urls: int,
    ignored: list[tuple],
) -> None:
    sep = "=" * 62
    print(f"\n{sep}")
    print("  RESUMO FINAL")
    print(sep)
    print(f"  URLs de produtos no sitemap    : {total_urls}")
    print(f"  Produtos base processados      : {len({r['Produto Base'] for r in product_rows})}")
    print(f"  SKUs/linhas geradas            : {len(product_rows)}")
    print(f"  Erros                          : {len(error_rows)}")

    counter = Counter(r["Produto Base"] for r in product_rows)
    print("\n  Top 20 produtos com mais SKUs:")
    for name, count in counter.most_common(20):
        print(f"    {count:3d}  {name}")

    if ignored:
        print(f"\n  URLs ignoradas ({len(ignored)}):")
        for url, reason in ignored[:30]:
            print(f"    {reason[:35]:35s}  {url}")
    print(sep)


# ─────────────────────────────────────────────────────────────────────────────
# 12. main
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Scraper de produtos Mary Kay Brasil via VTEX API"
    )
    parser.add_argument(
        "--max-products", type=int, default=0,
        help="Limita quantidade de produtos a processar (0 = sem limite)",
    )
    parser.add_argument(
        "--product-url", type=str, default="",
        help="Processa apenas este URL de produto",
    )
    parser.add_argument(
        "--output-prefix", type=str, default="produtos_marykay_completo",
        help="Prefixo dos arquivos de saída",
    )
    parser.add_argument(
        "--dump-debug", action="store_true",
        help="Salva JSON bruto da API em debug_json/",
    )
    parser.add_argument(
        "--dump-sitemap", action="store_true",
        help="Salva CSV de debug do sitemap",
    )
    args = parser.parse_args()

    log_file = f"{args.output_prefix}_log.txt"
    setup_logging(log_file)

    debug_dir: Optional[Path] = None
    if args.dump_debug:
        debug_dir = Path("debug_json")
        debug_dir.mkdir(exist_ok=True)

    product_rows:      list[dict] = []
    audit_rows:        list[dict] = []
    error_rows:        list[dict] = []
    sitemap_debug_rows: list[dict] = []

    # ── Modo single-URL ───────────────────────────────────────────────────
    if args.product_url:
        logger.info("Modo single-URL: %s", args.product_url)
        rows = process_product_url(
            args.product_url, error_rows, audit_rows, debug_dir
        )
        product_rows.extend(rows)
        save_outputs(
            product_rows, audit_rows, error_rows,
            sitemap_debug_rows, args.output_prefix, args.dump_sitemap,
        )
        print_summary(product_rows, error_rows, 1, [])
        logger.info("Log salvo: %s", log_file)
        return

    # ── Modo sitemap completo ─────────────────────────────────────────────
    logger.info("Lendo sitemap: %s", SITEMAP_URL)
    product_urls = read_sitemap_recursive(
        SITEMAP_URL, sitemap_rows=sitemap_debug_rows
    )

    ignored = [
        (r["url_encontrada"], r["motivo"])
        for r in sitemap_debug_rows
        if r["incluida_sim_nao"] == "não" and r["url_encontrada"]
    ]

    # Deduplicar preservando ordem
    seen_urls: set[str] = set()
    unique_urls: list[str] = []
    for u in product_urls:
        nu = normalize_product_url(u)
        if nu not in seen_urls:
            seen_urls.add(nu)
            unique_urls.append(u)

    total_urls = len(unique_urls)
    logger.info("URLs únicas de produto: %d", total_urls)

    if args.max_products and args.max_products < total_urls:
        unique_urls = unique_urls[: args.max_products]
        logger.info("Limitado a %d produtos (--max-products)", args.max_products)

    for i, url in enumerate(unique_urls, 1):
        logger.info("[%d/%d] %s", i, len(unique_urls), url)
        rows = process_product_url(url, error_rows, audit_rows, debug_dir)
        product_rows.extend(rows)

        # Salvar checkpoints a cada 50 produtos para não perder progresso
        if i % 50 == 0:
            logger.info("Checkpoint: %d produtos processados, %d linhas geradas",
                        i, len(product_rows))

    save_outputs(
        product_rows, audit_rows, error_rows,
        sitemap_debug_rows, args.output_prefix, args.dump_sitemap,
    )
    print_summary(product_rows, error_rows, total_urls, ignored)
    logger.info("Log salvo: %s", log_file)


if __name__ == "__main__":
    main()
