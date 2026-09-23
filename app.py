import io
from datetime import datetime

import pandas as pd
import streamlit as st
import xlrd
from openpyxl import load_workbook

st.set_page_config(page_title="Actualizador de reportes", page_icon="📊", layout="centered")

st.title("📊 Actualizador de reportes")
st.write("Sube los cinco archivos Excel. La aplicación actualizará el reporte y generará un único archivo .xlsx para descargar.")

st.info(
    "No necesitas convertir ningún archivo a CSV. "
    "Coupons, Promotional Models, S.O. y Raw data pueden ser .xls o .xlsx."
)

coupons_file = st.file_uploader("Coupons (.xls o .xlsx)", type=["xls", "xlsx"])
promo_file = st.file_uploader("Promotional Models (.xls o .xlsx)", type=["xls", "xlsx"])
excel_original = st.file_uploader("Excel original del reporte (.xlsx)", type=["xlsx"])
so_file = st.file_uploader("S.O. (.xls o .xlsx)", type=["xls", "xlsx"])
raw_file = st.file_uploader("Raw data (.xls o .xlsx)", type=["xls", "xlsx"])


def is_xls(uploaded_file):
    return uploaded_file.name.lower().endswith(".xls")


def read_excel_safely(uploaded_file, preferred_sheet_contains=None):
    """Lee la primera hoja, o la primera cuyo nombre contenga el texto indicado."""
    uploaded_file.seek(0)

    if is_xls(uploaded_file):
        workbook = xlrd.open_workbook(
            file_contents=uploaded_file.read(),
            ignore_workbook_corruption=True,
        )
        sheet_names = workbook.sheet_names()

        if not sheet_names:
            raise ValueError(f'El archivo "{uploaded_file.name}" no contiene hojas.')

        selected_sheet = sheet_names[0]
        if preferred_sheet_contains:
            search_text = preferred_sheet_contains.lower()
            selected_sheet = next(
                (name for name in sheet_names if search_text in name.lower()),
                selected_sheet,
            )

        sheet = workbook.sheet_by_index(sheet_names.index(selected_sheet))
        data = [sheet.row_values(row_index) for row_index in range(sheet.nrows)]

        if not data:
            raise ValueError(f'La hoja "{selected_sheet}" de "{uploaded_file.name}" está vacía.')

        headers = [str(header).strip() for header in data[0]]
        dataframe = pd.DataFrame(data[1:], columns=headers)
        return dataframe, selected_sheet

    uploaded_file.seek(0)
    excel = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheet_names = excel.sheet_names

    if not sheet_names:
        raise ValueError(f'El archivo "{uploaded_file.name}" no contiene hojas.')

    selected_sheet = sheet_names[0]
    if preferred_sheet_contains:
        search_text = preferred_sheet_contains.lower()
        selected_sheet = next(
            (name for name in sheet_names if search_text in name.lower()),
            selected_sheet,
        )

    uploaded_file.seek(0)
    dataframe = pd.read_excel(
        uploaded_file,
        sheet_name=selected_sheet,
        engine="openpyxl",
    )
    return dataframe, selected_sheet


def find_column(dataframe, desired_name):
    normalized = {
        str(column).strip().casefold(): column
        for column in dataframe.columns
    }
    return normalized.get(desired_name.casefold())


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


def prepare_coupons(dataframe):
    date_column = find_column(dataframe, "Fecha Compra")
    if date_column is None:
        available = ", ".join(str(column) for column in dataframe.columns)
        raise ValueError(
            'No se encontró la columna "Fecha Compra" en Coupons. '
            f"Columnas detectadas: {available}"
        )

    dataframe = dataframe.copy()
    dataframe["Num de SEM"] = dataframe[date_column].apply(iso_week_number)

    columns = [column for column in dataframe.columns if column != "Num de SEM"]
    insert_at = min(28, len(columns))
    columns.insert(insert_at, "Num de SEM")
    return dataframe[columns]


def excel_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def replace_worksheet(workbook, sheet_name, dataframe):
    if sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        workbook.remove(worksheet)

    worksheet = workbook.create_sheet(title=sheet_name)

    for column_index, column_name in enumerate(dataframe.columns, start=1):
        worksheet.cell(row=1, column=column_index, value=str(column_name))

    for row_index, row in enumerate(dataframe.itertuples(index=False, name=None), start=2):
        for column_index, value in enumerate(row, start=1):
            worksheet.cell(row=row_index, column=column_index, value=excel_value(value))

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions


def move_sheet_to_position(workbook, sheet_name, position):
    worksheet = workbook[sheet_name]
    workbook._sheets.remove(worksheet)
    workbook._sheets.insert(min(position, len(workbook._sheets)), worksheet)


if st.button("Procesar y descargar reporte", type="primary"):
    required_files = {
        "Coupons": coupons_file,
        "Promotional Models": promo_file,
        "Excel original": excel_original,
        "S.O.": so_file,
        "Raw data": raw_file,
    }
    missing = [name for name, file in required_files.items() if file is None]

    if missing:
        st.error("Faltan estos archivos: " + ", ".join(missing))
        st.stop()

    try:
        with st.spinner("Leyendo los archivos y actualizando el reporte..."):
            coupons_df, coupons_sheet = read_excel_safely(
                coupons_file,
                preferred_sheet_contains="coupon",
            )
            coupons_df = prepare_coupons(coupons_df)

            promo_df, promo_sheet = read_excel_safely(promo_file)
            so_df, so_sheet = read_excel_safely(so_file)
            raw_df, raw_sheet = read_excel_safely(raw_file)

            excel_original.seek(0)
            workbook = load_workbook(excel_original)

            target_sheets = [
                ("Coupons", coupons_df),
                ("Promotional Models", promo_df),
                ("S.O", so_df),
                ("Raw data", raw_df),
            ]

            original_positions = {
                name: workbook.sheetnames.index(name)
                for name, _ in target_sheets
                if name in workbook.sheetnames
            }

            for sheet_name, dataframe in target_sheets:
                replace_worksheet(workbook, sheet_name, dataframe)

            for sheet_name, _ in target_sheets:
                if sheet_name in original_positions:
                    move_sheet_to_position(workbook, sheet_name, original_positions[sheet_name])

            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)

        st.success("Reporte actualizado correctamente.")
        st.caption(
            f"Hoja usada en Coupons: {coupons_sheet} · "
            f"Promotional Models: {promo_sheet} · "
            f"S.O.: {so_sheet} · Raw data: {raw_sheet}"
        )

        col1, col2 = st.columns(2)
        col1.metric("Filas Coupons", len(coupons_df))
        col2.metric("Filas Promotional Models", len(promo_df))
        col1.metric("Filas S.O.", len(so_df))
        col2.metric("Filas Raw data", len(raw_df))

        st.download_button(
            label="Descargar reporte actualizado (.xlsx)",
            data=output.getvalue(),
            file_name="reporte_actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as error:
        st.error(f"Error al procesar los archivos: {error}")
        st.exception(error)
