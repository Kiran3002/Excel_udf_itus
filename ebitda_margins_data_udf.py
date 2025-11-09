"""
excel_index_udfs.py

Excel UDFs (xlwings) to fetch index constituents from a local SQLite database.

Configuration is read from config.ini, allowing easy migration from SQLite to RDS.

Usage:
- Create a config.ini file in the same directory (see example below)
- Import functions via xlwings → Import Functions
- From Excel: =get_monthly_data("nifty_500","2024-03-31")
"""

import xlwings as xw
import sqlite3
import pandas as pd
import configparser
import os
from datetime import datetime
from typing import List, Tuple
from functools import lru_cache

# ---------------- CONFIG -----------------
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.ini')

# Disable interpolation so '%' in date format works normally
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
# -----------------------------------------


def _get_connection():
    """Return DB connection based on db_type."""
    if DB_TYPE == 'sqlite':
        return sqlite3.connect(DB_PATH)
    else:
        import psycopg2
        return psycopg2.connect(**RDS_CONFIG)


def _run_query_df(sql: str, params: Tuple = ()):
    """Execute SQL query and return a DataFrame."""
    conn = _get_connection()
    try:
        df = pd.read_sql_query(sql, conn, params=params)
        return df
    finally:
        conn.close()


def _format_date(date_value: str) -> str:
    """
    Convert Excel date input (e.g., 2024-03-31) to the configured DB format.
    Handles inputs with or without time components.
    """
    try:
        date_value = str(date_value).strip().replace('"', '')  # Clean up stray spaces/quotes

        # Try parsing just date (YYYY-MM-DD)
        if len(date_value) == 10:
            date_obj = datetime.strptime(date_value, "%Y-%m-%d")
            formatted = date_obj.strftime(DATE_FORMAT)
            return formatted

        # If already full timestamp (e.g. 2024-03-31 00:00:00), return as-is
        datetime.strptime(date_value, DATE_FORMAT)
        return date_value

    except Exception:
        # If format is unknown, just return raw value
        return date_value

# ---------------- UDFs ------------------



@xw.func
@xw.ret(expand='table')
def get_monthly_data(index_name: str, date_value: str):
    """
    Excel formula:
        =get_monthly_data("nifty_500", "2024-03-31")

    Returns constituents (company_name, sector, mcap_category, weights)
    for the given index as on the specified date.
    """
    try:
        if not index_name or not date_value:
            return [["Error: index_name and date are required"]]

        formatted_date = _format_date(date_value)

        # Flexible SQL: allows exact match OR date prefix match
        sql = f"""
            SELECT company_name, sector, mcap_category, weights
            FROM {TABLE_NAME}
            WHERE index_name = ?
              AND (date = ? OR date LIKE ?)
            ORDER BY weights DESC
        """

        df = _run_query_df(sql, (index_name, formatted_date, formatted_date.split(" ")[0] + "%"))

        if df.empty:
            return [[f"No data found for index='{index_name}' on '{formatted_date}'"]]

        # Return headers + rows for Excel display
        return [df.columns.tolist()] + df.values.tolist()

    except Exception as e:
        return [[f"Error: {e}"]]

@xw.func
@xw.ret(expand='table')
def get_series(index_name: str, start_date: str, end_date: str):
    """
    Excel formula:
        =get_series("nifty_500", "2024-03-31", "2025-09-30")

    Returns (index_name, accord_code, company_name, sector,
    mcap_category, date, weights) for the given index between start and end dates.
    """
    try:
        if not index_name or not start_date or not end_date:
            return [["Error: index_name, start_date, and end_date are required"]]

        start_fmt = _format_date(start_date)
        end_fmt = _format_date(end_date)

        sql = f"""
            SELECT index_name, accord_code, company_name, sector,
                   mcap_category, date, weights
            FROM {TABLE_NAME}
            WHERE index_name = ?
              AND date BETWEEN ? AND ?
            ORDER BY date ASC, weights DESC
        """

        df = _run_query_df(sql, (index_name, start_fmt, end_fmt))

        if df.empty:
            return [[f"No data found for index='{index_name}' between '{start_fmt}' and '{end_fmt}'"]]

        return [df.columns.tolist()] + df.values.tolist()

    except Exception as e:
        return [[f"Error: {e}"]]


# -------------------------------------------------------------------
# 3️⃣ Get Data of a Particular Index on a Given Date
# -------------------------------------------------------------------
@xw.func
@xw.ret(expand='table')
def get_matrix(date_value: str, index_name: str):
    """
    Excel formula:
        =get_matrix("2024-03-31", "nifty_500")

    Returns (accord_code, company_name, sector, mcap_category,
    date, weights) for the given index on the specified date.
    """
    try:
        if not date_value or not index_name:
            return [["Error: date and index_name are required"]]

        formatted_date = _format_date(date_value)

        sql = f"""
            SELECT accord_code, company_name, sector,
                   mcap_category, date, weights
            FROM {TABLE_NAME}
            WHERE index_name = ?
              AND (date = ? OR date LIKE ?)
            ORDER BY weights DESC
        """

        df = _run_query_df(sql, (index_name, formatted_date, formatted_date.split(" ")[0] + "%"))

        if df.empty:
            return [[f"No data found for index='{index_name}' on '{formatted_date}'"]]

        return [df.columns.tolist()] + df.values.tolist()

    except Exception as e:
        return [[f"Error: {e}"]]


# -------------------------------------------------------------------
# 4️⃣ Get All Data for a Particular Index
# -------------------------------------------------------------------
@xw.func
@xw.ret(expand='table')
def get_all_data(index_name: str):
    """
    Excel formula:
        =get_all_data("nifty_500")

    Returns all records for the given index across all available dates.
    """
    try:
        if not index_name:
            return [["Error: index_name is required"]]

        sql = f"""
            SELECT accord_code, company_name, sector,
                   mcap_category, date, weights
            FROM {TABLE_NAME}
            WHERE index_name = ?
            ORDER BY date ASC, weights DESC
        """

        df = _run_query_df(sql, (index_name,))

        if df.empty:
            return [[f"No data found for index='{index_name}'"]]

        return [df.columns.tolist()] + df.values.tolist()

    except Exception as e:
        return [[f"Error: {e}"]]