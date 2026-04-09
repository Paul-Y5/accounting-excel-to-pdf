#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de pré-visualização de PDFs.
Renderiza páginas de um PDF como imagens PIL para exibição na GUI.
"""

import os


def render_page(pdf_path: str, page: int = 0, dpi: int = 96):
    """Renderiza uma página do PDF como imagem PIL.

    Args:
        pdf_path: Caminho do ficheiro PDF.
        page: Índice da página (0-based).
        dpi: Resolução de renderização.

    Returns:
        Tuplo (PIL.Image, total_pages).

    Raises:
        ImportError: Se o PyMuPDF (fitz) não estiver instalado.
        FileNotFoundError: Se o ficheiro PDF não existir.
        IndexError: Se o índice da página for inválido.
    """
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"Ficheiro não encontrado: {pdf_path}")

    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise ImportError(
            "Para pré-visualizar PDFs, instale o pacote: pip install PyMuPDF"
        )

    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    if page < 0 or page >= total_pages:
        doc.close()
        raise IndexError(f"Página {page} inválida. O PDF tem {total_pages} página(s).")

    pdf_page = doc[page]
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = pdf_page.get_pixmap(matrix=mat)

    from PIL import Image
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    doc.close()
    return img, total_pages


def get_page_count(pdf_path: str) -> int:
    """Devolve o número total de páginas de um PDF.

    Args:
        pdf_path: Caminho do ficheiro PDF.

    Returns:
        Número de páginas.
    """
    import fitz
    doc = fitz.open(pdf_path)
    count = len(doc)
    doc.close()
    return count
