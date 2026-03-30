#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Módulo de base de dados SQLite.
Centraliza o armazenamento de histórico, perfis e cache de clientes.
"""

import json
import os
import sqlite3
from datetime import datetime

from src.config import get_config_dir


def _get_db_path() -> str:
    """Retorna o caminho do ficheiro da base de dados."""
    return os.path.join(get_config_dir(), 'conversor.db')


def _get_connection() -> sqlite3.Connection:
    """Cria e retorna uma conexão à base de dados."""
    conn = sqlite3.connect(_get_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """Inicializa a base de dados, criando as tabelas se necessário."""
    conn = _get_connection()
    try:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                source_file TEXT NOT NULL,
                source_path TEXT NOT NULL,
                output_path TEXT NOT NULL DEFAULT '',
                mode TEXT NOT NULL,
                clients_count INTEGER NOT NULL DEFAULT 0,
                success INTEGER NOT NULL DEFAULT 1,
                error TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                config_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS client_cache (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file TEXT NOT NULL,
                client_name TEXT NOT NULL,
                client_sigla TEXT NOT NULL DEFAULT '',
                nif TEXT NOT NULL DEFAULT '',
                last_seen TEXT NOT NULL,
                conversion_count INTEGER NOT NULL DEFAULT 1
            );

            CREATE INDEX IF NOT EXISTS idx_history_timestamp ON history(timestamp);
            CREATE INDEX IF NOT EXISTS idx_client_cache_source ON client_cache(source_file);
            CREATE UNIQUE INDEX IF NOT EXISTS idx_client_cache_unique
                ON client_cache(source_file, client_name);
        """)
        conn.commit()
    finally:
        conn.close()


# ============================================
# HISTÓRICO
# ============================================

def add_history_entry(source_file: str, output_path: str, mode: str,
                      clients_count: int, success: bool, error_msg: str = ''):
    """Adiciona uma entrada ao histórico."""
    conn = _get_connection()
    try:
        conn.execute(
            """INSERT INTO history (timestamp, source_file, source_path, output_path,
               mode, clients_count, success, error)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                datetime.now().isoformat(),
                os.path.basename(source_file),
                source_file,
                output_path,
                mode,
                clients_count,
                1 if success else 0,
                error_msg,
            )
        )
        # Manter apenas as últimas 500 entradas
        conn.execute("""
            DELETE FROM history WHERE id NOT IN (
                SELECT id FROM history ORDER BY id DESC LIMIT 500
            )
        """)
        conn.commit()
    finally:
        conn.close()


def get_history(limit: int = 50) -> list:
    """Retorna as últimas entradas do histórico (mais recentes primeiro)."""
    conn = _get_connection()
    try:
        cursor = conn.execute(
            "SELECT * FROM history ORDER BY id DESC LIMIT ?", (limit,)
        )
        rows = cursor.fetchall()
        return [
            {
                'timestamp': row['timestamp'],
                'source_file': row['source_file'],
                'source_path': row['source_path'],
                'output_path': row['output_path'],
                'mode': row['mode'],
                'clients_count': row['clients_count'],
                'success': bool(row['success']),
                'error': row['error'],
            }
            for row in rows
        ]
    finally:
        conn.close()


def clear_history():
    """Limpa todo o histórico."""
    conn = _get_connection()
    try:
        conn.execute("DELETE FROM history")
        conn.commit()
    finally:
        conn.close()


# ============================================
# PERFIS DE CONFIGURAÇÃO
# ============================================

def list_profiles_db() -> list:
    """Lista os nomes dos perfis guardados."""
    conn = _get_connection()
    try:
        cursor = conn.execute("SELECT name FROM profiles ORDER BY name")
        return [row['name'] for row in cursor.fetchall()]
    finally:
        conn.close()


def save_profile_db(name: str, config: dict) -> bool:
    """Guarda ou atualiza um perfil de configuração."""
    conn = _get_connection()
    try:
        now = datetime.now().isoformat()
        config_json = json.dumps(config, ensure_ascii=False)
        conn.execute(
            """INSERT INTO profiles (name, config_json, created_at, updated_at)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(name) DO UPDATE SET config_json=?, updated_at=?""",
            (name, config_json, now, now, config_json, now)
        )
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        conn.close()


def load_profile_db(name: str) -> dict:
    """Carrega um perfil de configuração."""
    import copy
    from src.config import DEFAULT_CONFIG

    conn = _get_connection()
    try:
        cursor = conn.execute(
            "SELECT config_json FROM profiles WHERE name = ?", (name,)
        )
        row = cursor.fetchone()
        if not row:
            return None
        saved = json.loads(row['config_json'])
        # Merge com defaults
        config = copy.deepcopy(DEFAULT_CONFIG)
        for section, values in saved.items():
            if section in config:
                if isinstance(values, dict):
                    config[section].update(values)
                else:
                    config[section] = values
        return config
    except Exception:
        return None
    finally:
        conn.close()


def delete_profile_db(name: str) -> bool:
    """Apaga um perfil de configuração."""
    conn = _get_connection()
    try:
        conn.execute("DELETE FROM profiles WHERE name = ?", (name,))
        conn.commit()
        return conn.total_changes > 0
    except Exception:
        return False
    finally:
        conn.close()


# ============================================
# CACHE DE CLIENTES
# ============================================

def update_client_cache(source_file: str, clients: list):
    """Atualiza a cache de clientes a partir de uma lista de dicts.

    Cada dict deve ter pelo menos 'name'. Opcionalmente 'sigla' e 'nif'.
    """
    conn = _get_connection()
    try:
        now = datetime.now().isoformat()
        for client in clients:
            name = client.get('name', '').strip()
            if not name:
                continue
            sigla = client.get('sigla', '').strip()
            nif = client.get('nif', '').strip()
            conn.execute(
                """INSERT INTO client_cache (source_file, client_name, client_sigla, nif,
                   last_seen, conversion_count)
                   VALUES (?, ?, ?, ?, ?, 1)
                   ON CONFLICT(source_file, client_name) DO UPDATE SET
                       client_sigla = ?,
                       nif = CASE WHEN ? != '' THEN ? ELSE nif END,
                       last_seen = ?,
                       conversion_count = conversion_count + 1""",
                (source_file, name, sigla, nif, now,
                 sigla, nif, nif, now)
            )
        conn.commit()
    finally:
        conn.close()


def get_cached_clients(source_file: str = None, limit: int = 200) -> list:
    """Retorna clientes da cache.

    Args:
        source_file: Se fornecido, filtra por ficheiro de origem.
        limit: Número máximo de resultados.

    Returns:
        Lista de dicts com dados dos clientes.
    """
    conn = _get_connection()
    try:
        if source_file:
            cursor = conn.execute(
                """SELECT * FROM client_cache WHERE source_file = ?
                   ORDER BY client_name LIMIT ?""",
                (source_file, limit)
            )
        else:
            cursor = conn.execute(
                "SELECT * FROM client_cache ORDER BY last_seen DESC LIMIT ?",
                (limit,)
            )
        return [
            {
                'source_file': row['source_file'],
                'name': row['client_name'],
                'sigla': row['client_sigla'],
                'nif': row['nif'],
                'last_seen': row['last_seen'],
                'conversion_count': row['conversion_count'],
            }
            for row in cursor.fetchall()
        ]
    finally:
        conn.close()


def clear_client_cache():
    """Limpa toda a cache de clientes."""
    conn = _get_connection()
    try:
        conn.execute("DELETE FROM client_cache")
        conn.commit()
    finally:
        conn.close()


# ============================================
# MIGRAÇÃO JSON → SQLite
# ============================================

def migrate_from_json():
    """Migra dados existentes de JSON para SQLite (histórico e perfis)."""
    config_dir = get_config_dir()

    # Migrar histórico
    history_path = os.path.join(config_dir, 'history.json')
    if os.path.exists(history_path):
        try:
            with open(history_path, 'r', encoding='utf-8') as f:
                entries = json.load(f)
            conn = _get_connection()
            try:
                for entry in entries:
                    conn.execute(
                        """INSERT OR IGNORE INTO history
                           (timestamp, source_file, source_path, output_path,
                            mode, clients_count, success, error)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (
                            entry.get('timestamp', ''),
                            entry.get('source_file', ''),
                            entry.get('source_path', ''),
                            entry.get('output_path', ''),
                            entry.get('mode', ''),
                            entry.get('clients_count', 0),
                            1 if entry.get('success', False) else 0,
                            entry.get('error', ''),
                        )
                    )
                conn.commit()
            finally:
                conn.close()
            # Renomear ficheiro antigo
            os.rename(history_path, history_path + '.bak')
        except Exception:
            pass

    # Migrar perfis
    profiles_dir = os.path.join(config_dir, 'profiles')
    if os.path.isdir(profiles_dir):
        try:
            for fname in os.listdir(profiles_dir):
                if fname.endswith('.json'):
                    name = fname[:-5]
                    fpath = os.path.join(profiles_dir, fname)
                    with open(fpath, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    save_profile_db(name, config)
            # Renomear pasta
            bak_dir = profiles_dir + '.bak'
            if not os.path.exists(bak_dir):
                os.rename(profiles_dir, bak_dir)
        except Exception:
            pass
