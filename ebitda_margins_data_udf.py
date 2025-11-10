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
from datetime import datetime
from typing import Tuple
from functools import lru_cache
import win32com.client
import logging
from logging.handlers import RotatingFileHandler
import time
import inspect
import functools
import time
# -------------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------------
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.ini')
LOG_FILE = os.path.join(os.path.dirname(__file__), 'query_log.txt')

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
try:
    logger = logging.getLogger("QueryLogger")
    logger.setLevel(logging.INFO)
    handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding='utf-8')
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
except Exception:
    # If logging setup fails, don’t break Excel UDFs
    logger = None

# -------------------------------------------------------------------
# CONNECTION & UTILITIES
# -------------------------------------------------------------------
_index_checked = False  # Global flag to ensure we only check once

def _get_connection():
    """Return database connection based on config and ensure index exists."""
    global _index_checked

    if DB_TYPE == 'sqlite':
        if not os.path.exists(DB_PATH):
            raise FileNotFoundError(f"SQLite DB not found at path: {DB_PATH}")
        conn = sqlite3.connect(DB_PATH)
        if not _index_checked:
            try:
                conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_equity_index_constituents_lookup
                ON equity_index_constituents (index_name, accord_code, date);
                """)
                conn.commit()
                _index_checked = True
                if logger:
                    logger.info("✅ Verified: index idx_index_components_lookup exists.")
            except Exception as e:
                if logger:
                    logger.warning(f"⚠️ Index check skipped due to: {e}")
        return conn
    else:
        import psycopg2
        return psycopg2.connect(**RDS_CONFIG)
    
# -------------------------------------------------------------------
# CACHING
# -------------------------------------------------------------------
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
    """Execute SQL with caching and performance logging."""
    params_key = tuple(str(p) for p in params)
    start = time.perf_counter()

    df = _cached_query(sql, params_key)

    duration = round((time.perf_counter() - start) * 1000, 3)  # in ms
    if logger:
        logger.info(f"⏱️ Query time: {duration} ms | SQL Params={params_key}")
        if duration > 50:  # > 0.05 seconds threshold
            logger.warning(f"⚠️ Slow query detected ({duration} ms): {params_key}")
    return df


# -------------------------------------------------------------------
# INPUT VALIDATION
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
    """
    Validate that required inputs are provided and match expected types.
    expected_types = {'param': str, 'param2': str}
    """
    for name, value in kwargs.items():
        if value is None or str(value).strip() == "":
            raise ValueError(f"Missing required input: {name}")
        expected_type = expected_types.get(name)
        if expected_type and not isinstance(value, expected_type):
            raise TypeError(
                f"Type Mismatch: Expected '{name}' as {expected_type.__name__}, "
                f"but got {type(value).__name__}"
            )

# -------------------------------------------------------------------
# LOG DECORATOR
# -------------------------------------------------------------------


def log_call(func):
    """Decorator for logging execution time and success/failure — Excel-safe."""
    @functools.wraps(func)
    def inner_wrapper(*args):  
        start = time.perf_counter()
        status = "SUCCESS"
        error_msg = None
        try:
            result = func(*args)
            return result
        except Exception as e:
            status = "FAILED"
            error_msg = str(e)
            raise
        finally:
            duration_ms = round((time.perf_counter() - start) * 1000, 2)
            if logger:
                params = ", ".join([repr(a) for a in args])
                msg = (
                    f"{'SUCCESS' if status == 'SUCCESS' else 'FAILURE'} "
                    f"Function: {func.__name__} | Params: ({params}) | "
                    f"Duration: {duration_ms} ms | Status: {status}"
                )
                if error_msg:
                    msg += f" | Error: {error_msg}"
                logger.info(msg)
    return inner_wrapper


# -------------------------------------------------------------------
# UDFS
# -------------------------------------------------------------------
@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_monthly_data(index_name: str, date_value: str):
    """Fetch constituents for a given index as on a specific date."""
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


@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_series(index_name: str, start_date: str, end_date: str):
    """Fetch index constituents and weights between start and end dates."""
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


@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_matrix(date_value: str, index_name: str):
    """Fetch all constituents of a given index as on a specific date."""
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


@xw.func(category="Finance UDFs")
@xw.ret(expand='table')
@log_call
def get_all_data(index_name: str):
    """Fetch all available data for a specific index across all dates."""
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


@xw.func(category="Finance UDFs")
@log_call
def clear_cache():
    """Clear cached queries (useful after DB updates)."""
    _cached_query.cache_clear()
    return "✅ Cache cleared successfully."

# -------------------------------------------------------------------
# EXCEL TOOLTIP REGISTRATION
# -------------------------------------------------------------------
@xw.func
def register_excel_udfs():
    """Register UDFs with descriptions and argument help."""
    try:
        try:
            excel = xw.apps.active.api
        except:
            excel = win32com.client.Dispatch("Excel.Application")

        functions = [
            ("get_monthly_data", "Fetch index data for a given date.", ("index_name", "date_value")),
            ("get_series", "Fetch index data between two dates.", ("index_name", "start_date", "end_date")),
            ("get_matrix", "Fetch data matrix for a given date.", ("date_value", "index_name")),
            ("get_all_data", "Fetch all records for a given index.", ("index_name",)),
            ("clear_cache", "Clear cached results from memory.", ()),
        ]

        for name, desc, args in functions:
            try:
                excel.MacroOptions(
                    Macro=name,
                    Description=desc,
                    ArgumentDescriptions=list(args),
                    Category="Finance UDFs"
                )
            except Exception:
                continue

        return "✅ UDFs registered successfully! (Save & reopen workbook for tooltips)"
    except Exception as e:
        return f"⚠️ Registration failed: {e}"
