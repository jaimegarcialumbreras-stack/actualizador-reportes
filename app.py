import gc
import os
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st
import xlrd
from openpyxl import load_workbook

st.set_page_config(page_title="Actualizador de reportes", page_icon="📊")

st.title("📊 Actualizador de reportes")
st.write("Sube Coupons, Promotional Models y el reporte original.")

st.info(
    "Coupons se carga en la hoja Raw data. "
    "Promotional Models se carga en la hoja S.O. "
    "No se utilizan CSV."
)

coupons_file = st.file_uploader("Coupons (.xls o .xlsx)", type=["xls", "xlsx"])
promo_file = st.file_uploader(
    "Promotional Models (.xls o .xlsx)",
    type=["xls", "xlsx"]
)
report_file = st.file_uploader("Reporte original (.xlsx)", type=["xlsx"])


def is_xls(uploaded_file):
    return uploaded_file.name.lower().endswith(".xls")


def find_column(columns, desired_name):
    target = desired_name.strip().casefold()
    for column in columns:
        if str(column).strip().casefold() == target:
            return column
    return None


def iso_week_number(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        if isinstance(value, datetime):
            date_value = value
        else:
            date_value = pd.to_datetime(value, dayfirst=True, errors="coerce")

        if pd.isna(date_value):
            return ""

        return int(date_value.isocalendar().week)
    except (TypeError, ValueError, AttributeError):
        return ""


def value_for_excel(value):
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def clear_sheet(sheet):
    if sheet.max_row:
        sheet.delete_rows(1, sheet.max_row)


def write_dataframe(sheet, dataframe):
    clear_sheet(sheet)

    for col_number, column_name in enumerate(dataframe.columns, start=1):
        sheet.cell(row=1, column=col_number, value=str(column_name))

    for row_number, row in enumerate(
        dataframe.itertuples(index=False, name=None),
        start=2
    ):
        for col_number, value in enumerate(row, start=1):
            sheet.cell(
                row=row_number,
                column=col_number,
                value=value_for_excel(value)
            )


def read_coupons(file):
    file.seek(0)

    if is_xls(file):
        book = xlrd.open_workbook(
            file_contents=file.read(),
            ignore_workbook_corruption=True
        )

        sheet_names = book.sheet_names()
        if not sheet_names:
            raise ValueError("Coupons no contiene hojas.")

        selected = next(
            (name for name in sheet_names if "coupon" in name.casefold()),
            sheet_names[0]
        )

        sheet = book.sheet_by_index(sheet_names.index(selected))
        if sheet.nrows == 0:
            raise ValueError(f'La hoja "{selected}" de Coupons está vacía.')

        headers = [str(value).strip() for value in sheet.row_values(0)]
        date_column_index = None

        for index, header in enumerate(headers):
            if header.casefold() == "fecha compra":
                date_column_index = index
                break

        if date_column_index is None:
            raise ValueError(
                'No se encontró la columna "Fecha Compra" en Coupons. '
                + "Columnas detectadas: " + ", ".join(headers)
            )

        output_headers = [h for h in headers if h != "Num de SEM"]
        output_headers.insert(min(28, len(output_headers)), "Num de
