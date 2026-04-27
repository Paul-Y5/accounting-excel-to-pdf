#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de sequência de números de documento.

Gera e mantém sequências numéricas únicas para documentos
(faturas, recibos, notas de entrega, etc.).

Formato de saída: "<SERIE> <ANO>/<NNNN>"  ex: "FT 2026/0001"
"""

from datetime import datetime

from src.database import _get_connection


# ============================================
# INICIALIZAÇÃO
# ============================================

def init_doc_sequences_table():
    """Cria a tabela doc_sequences se ainda não existir."""
    conn = _get_connection()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS doc_sequences (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                serie       TEXT    NOT NULL UNIQUE,
                ano         INTEGER NOT NULL,
                ultimo_numero INTEGER NOT NULL DEFAULT 0,
                reset_anual INTEGER NOT NULL DEFAULT 1
            )
        """)
        conn.commit()
    finally:
        conn.close()


# ============================================
# OPERAÇÕES PRINCIPAIS
# ============================================

def get_next_number(serie: str, reset_anual: bool = True) -> str:
    """Obtém e incrementa o próximo número de documento.

    A operação é atómica: lê o contador, incrementa e persiste numa
    única transacção para evitar duplicados em uso concorrente.

    Args:
        serie:       Código da série (ex: 'FT', 'FR', 'REC').
        reset_anual: Se True, o contador reinicia a 1 em cada ano civil.

    Returns:
        Número formatado, ex: ``'FT 2026/0001'``.
    """
    ano_atual = datetime.now().year
    conn = _get_connection()
    try:
        # Criar série se ainda não existir
        conn.execute(
            """INSERT OR IGNORE INTO doc_sequences
               (serie, ano, ultimo_numero, reset_anual)
               VALUES (?, ?, 0, ?)""",
            (serie, ano_atual, 1 if reset_anual else 0),
        )

        row = conn.execute(
            "SELECT ano, ultimo_numero, reset_anual FROM doc_sequences WHERE serie = ?",
            (serie,),
        ).fetchone()

        ano_serie = row['ano']
        numero = row['ultimo_numero']
        faz_reset = bool(row['reset_anual'])

        # Reset anual se mudou o ano
        if faz_reset and ano_serie != ano_atual:
            numero = 0
            ano_serie = ano_atual
            conn.execute(
                "UPDATE doc_sequences SET ano = ?, ultimo_numero = 0 WHERE serie = ?",
                (ano_atual, serie),
            )

        numero += 1
        conn.execute(
            "UPDATE doc_sequences SET ultimo_numero = ? WHERE serie = ?",
            (numero, serie),
        )
        conn.commit()

        return f"{serie} {ano_serie}/{numero:04d}"
    finally:
        conn.close()


def peek_next_number(serie: str) -> str:
    """Pré-visualiza o próximo número sem incrementar o contador.

    Args:
        serie: Código da série (ex: 'FT').

    Returns:
        Número que seria emitido na próxima chamada a :func:`get_next_number`.
    """
    ano_atual = datetime.now().year
    conn = _get_connection()
    try:
        row = conn.execute(
            "SELECT ano, ultimo_numero, reset_anual FROM doc_sequences WHERE serie = ?",
            (serie,),
        ).fetchone()

        if not row:
            return f"{serie} {ano_atual}/0001"

        ano_serie = row['ano']
        numero = row['ultimo_numero']
        faz_reset = bool(row['reset_anual'])

        if faz_reset and ano_serie != ano_atual:
            return f"{serie} {ano_atual}/0001"

        return f"{serie} {ano_serie}/{numero + 1:04d}"
    finally:
        conn.close()


# ============================================
# GESTÃO DE SÉRIES
# ============================================

def list_series() -> list:
    """Lista todas as séries configuradas.

    Returns:
        Lista de dicts ordenada por série, cada um com:
        ``serie``, ``ano``, ``ultimo_numero``, ``reset_anual``, ``proximo``.
    """
    conn = _get_connection()
    try:
        rows = conn.execute(
            "SELECT serie, ano, ultimo_numero, reset_anual "
            "FROM doc_sequences ORDER BY serie"
        ).fetchall()
        result = []
        for row in rows:
            result.append({
                'serie': row['serie'],
                'ano': row['ano'],
                'ultimo_numero': row['ultimo_numero'],
                'reset_anual': bool(row['reset_anual']),
                'proximo': peek_next_number(row['serie']),
            })
        return result
    finally:
        conn.close()


def upsert_serie(serie: str, ultimo_numero: int = 0,
                 ano: int = None, reset_anual: bool = True):
    """Cria ou actualiza uma série.

    Args:
        serie:          Código da série.
        ultimo_numero:  Último número já emitido (0 = série nova).
        ano:            Ano a associar (padrão: ano actual).
        reset_anual:    Se True, o contador reinicia em cada ano civil.
    """
    if ano is None:
        ano = datetime.now().year
    conn = _get_connection()
    try:
        conn.execute(
            """INSERT INTO doc_sequences (serie, ano, ultimo_numero, reset_anual)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(serie) DO UPDATE SET
                   ano           = excluded.ano,
                   ultimo_numero = excluded.ultimo_numero,
                   reset_anual   = excluded.reset_anual""",
            (serie, ano, ultimo_numero, 1 if reset_anual else 0),
        )
        conn.commit()
    finally:
        conn.close()


def reset_serie(serie: str, ano: int = None):
    """Reinicia o contador de uma série para zero.

    Args:
        serie: Código da série a reiniciar.
        ano:   Ano a definir (padrão: ano actual).
    """
    if ano is None:
        ano = datetime.now().year
    conn = _get_connection()
    try:
        conn.execute(
            "UPDATE doc_sequences SET ultimo_numero = 0, ano = ? WHERE serie = ?",
            (ano, serie),
        )
        conn.commit()
    finally:
        conn.close()


def delete_serie(serie: str):
    """Remove uma série da base de dados.

    Args:
        serie: Código da série a remover.
    """
    conn = _get_connection()
    try:
        conn.execute("DELETE FROM doc_sequences WHERE serie = ?", (serie,))
        conn.commit()
    finally:
        conn.close()
