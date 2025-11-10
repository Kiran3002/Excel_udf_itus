"""
excel_index_udfs.py

Excel UDFs (xlwings) to fetch index constituents from a local SQLite database.

Configuration is read from config.ini, allowing easy migration from SQLite to RDS.

Usage (Excel formulas):
    =get_monthly_data("nifty_500","2024-03-31")
    =get_series("nifty_500","2024-03-31","2025-09-30")
    =get_matrix("2024-03-31","nifty_500")
    =get_all_data("nifty_500")
"""

import xlwings as xw
import sqlite3
import pandas as pd
import configparser
import os
import time
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from typing import Tuple
from functools import lru_cache

# -------------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------------
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.ini')
config = configparser.ConfigParser(interpolation=None)
config.read(CONFIG_FILE)

DB_TYPE = config.get('DATABASE', 'db_type', fallback='sqlite')
DB_PATH = config.get('DATABASE', 'db_path', fallback='index_data.db')
TABLE_NAME = config.get('DATABASE', 'table_name', fallback='index_components')
DATE_FORMAT = config.get('FORMAT', 'date_format', fallback='%Y-%m-%d %H:%M:%S')

RDS_CONFIG = {
    'host': config.get('DATABASE', 'host', fallback=''),
    'port': config.getint('DATABASE', 'port', fallback=5432),
    'dbname': config.get('DATABASE', 'dbname', fallback=''),
    'user': config.get('DATABASE', 'user', fallback=''),
    'password': config.get('DATABASE', 'password', fallback='')
}

# -------------------------------------------------------------------
# LOGGING SETUP
# -------------------------------------------------------------------
LOG_FILE = os.path.join(os.path.dirname(__file__), "query_log.txt")

logger = logging.getLogger("UDFLogger")
logger.setLevel(logging.INFO)

# Avoid duplicate handlers if re-imported by Excel
if not logger.handlers:
    handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    handler.setFormatter(fmt)
    logger.addHandler(handler)


def log_event(func_name: str, params: dict, status: str, msg: str, exec_time_ms: float):
    """Helper to log each UDF call."""
    logger.info(
        f"Function={func_name} | Params={params} | Time={exec_time_ms:.2f} ms | Status={status} | Msg={msg}"
    )


# -------------------------------------------------------------------
# DATABASE UTILITIES
# -------------------------------------------------------------------
def _get_connection():
    """Return database connection based on config."""
    if DB_TYPE == 'sqlite':
        if not os.path.exists(DB_PATH):
            raise FileNotFoundError(f"SQLite DB not found at path: {DB_PATH}")
        return sqlite3.connect(DB_PATH)
    else:
        import psycopg2
        return psycopg2.connect(**RDS_CONFIG)

@lru_cache(maxsize=64)
def _cached_query(sql: str, params_key: Tuple[str]):
    """Run and cache SQL queries (LRU cache for speed)."""
    conn = _get_connection()
    try:
        df = pd.read_sql_query(sql, conn, params=params_key)
        return df
    finally:
        conn.close()

def _run_query_df(sql: str, params: Tuple = ()):
    """Execute SQL with caching."""
    params_key = tuple(str(p) for p in params)
    return _cached_query(sql, params_key)


# -------------------------------------------------------------------
# VALIDATION UTILITIES
# -------------------------------------------------------------------
def _format_date(date_value: str) -> str:
    """Normalize Excel input date to match DB format."""
    try:
        date_value = str(date_value).strip().replace('"', '')
        if len(date_value) == 10:
            dt = datetime.strptime(date_value, "%Y-%m-%d")
            return dt.strftime(DATE_FORMAT)
        datetime.strptime(date_value, DATE_FORMAT)
        return date_value
    except Exception:
        return date_value


def _validate_inputs_with_types(expected_types: dict, **kwargs):
    """Validate presence and type of each argument."""
    for name, value in kwargs.items():
        if value is None or str(value).strip() == "":
            raise ValueError(f"Missing required input: {name}")
        expected_type = expected_types.get(name)
        if expected_type and not isinstance(value, expected_type):
            raise TypeError(
                f"Type Mismatch: Expected '{name}' as {expected_type.__name__}, but got {type(value).__name__}"
            )


# -------------------------------------------------------------------
# UDF DECORATOR FOR LOGGING + TIMING
# -------------------------------------------------------------------
def udf_logger(fn):
    """Excel-safe logger decorator for UDFs."""
    def xl_safe_logger_func(*args):
        start = time.perf_counter()
        params = {f"arg{i+1}": v for i, v in enumerate(args)}
        try:
            result = fn(*args)
            exec_time = (time.perf_counter() - start) * 1000
            log_event(fn.__name__, params, "SUCCESS", "OK", exec_time)
            return result
        except Exception as e:
            exec_time = (time.perf_counter() - start) * 1000
            log_event(fn.__name__, params, "FAILURE", str(e), exec_time)
            return [[f"❌ {e}"]]
    return xl_safe_logger_func



# -------------------------------------------------------------------
# EXCEL UDFS
# -------------------------------------------------------------------
@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@udf_logger
def get_monthly_data(index_name: str, date_value: str):
    """Fetch constituents for a given index as on a specific date."""
    try:
        _validate_inputs_with_types({'index_name': str, 'date_value': str},
                                    index_name=index_name, date_value=date_value)
        formatted_date = _format_date(date_value)
        sql = f"""
            SELECT company_name, sector, mcap_category, weights
            FROM {TABLE_NAME}
            WHERE index_name = ? AND (date = ? OR date LIKE ?)
            ORDER BY weights DESC
        """
        df = _run_query_df(sql, (index_name, formatted_date, formatted_date.split(" ")[0] + "%"))
        if df.empty:
            return [[f"⚠️ No data found for index='{index_name}' on '{formatted_date}'"]]
        return [df.columns.tolist()] + df.values.tolist()
    except Exception as e:
        return [[f"❌ {e}"]]


@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@udf_logger
def get_series(index_name: str, start_date: str, end_date: str):
    """Fetch index constituents and weights between start and end dates."""
    try:
        _validate_inputs_with_types(
            {'index_name': str, 'start_date': str, 'end_date': str},
            index_name=index_name, start_date=start_date, end_date=end_date)
        start_fmt = _format_date(start_date)
        end_fmt = _format_date(end_date)
        sql = f"""
            SELECT index_name, accord_code, company_name, sector,
                   mcap_category, date, weights
            FROM {TABLE_NAME}
            WHERE index_name = ? AND date BETWEEN ? AND ?
            ORDER BY date ASC, weights DESC
        """
        df = _run_query_df(sql, (index_name, start_fmt, end_fmt))
        if df.empty:
            return [[f"⚠️ No records found for '{index_name}' between {start_fmt} and {end_fmt}."]]
        return [df.columns.tolist()] + df.values.tolist()
    except Exception as e:
        return [[f"❌ {e}"]]


@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@udf_logger
def get_matrix(date_value: str, index_name: str):
    """Fetch all constituents of a given index as on a specific date."""
    try:
        _validate_inputs_with_types({'date_value': str, 'index_name': str},
                                    date_value=date_value, index_name=index_name)
        formatted_date = _format_date(date_value)
        sql = f"""
            SELECT accord_code, company_name, sector,
                   mcap_category, date, weights
            FROM {TABLE_NAME}
            WHERE index_name = ? AND (date = ? OR date LIKE ?)
            ORDER BY weights DESC
        """
        df = _run_query_df(sql, (index_name, formatted_date, formatted_date.split(" ")[0] + "%"))
        if df.empty:
            return [[f"⚠️ No records found for '{index_name}' on {formatted_date}."]]
        return [df.columns.tolist()] + df.values.tolist()
    except Exception as e:
        return [[f"❌ {e}"]]


@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@udf_logger
def get_all_data(index_name: str):
    """Fetch all available data for a specific index across all dates."""
    try:
        _validate_inputs_with_types({'index_name': str}, index_name=index_name)
        sql = f"""
            SELECT accord_code, company_name, sector,
                   mcap_category, date, weights
            FROM {TABLE_NAME}
            WHERE index_name = ?
            ORDER BY date ASC, weights DESC
        """
        df = _run_query_df(sql, (index_name,))
        if df.empty:
            return [[f"⚠️ No data found for index='{index_name}'."]]
        return [df.columns.tolist()] + df.values.tolist()
    except Exception as e:
        return [[f"❌ {e}"]]


@xw.func(category="Finance UDFs")
@udf_logger
def clear_cache():
    """Clear cached queries (useful after DB updates)."""
    _cached_query.cache_clear()
    return "✅ Cache cleared successfully."