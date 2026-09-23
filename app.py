import gc
import io
from datetime import datetime

import pandas as pd
import streamlit as st
import xlrd
from openpyxl import load_workbook

st.set_page_config(page_title="Actualizador de reportes", page_icon="📊")

st.title("📊 Actualizador de reportes")
st.write("Sube los dos archivos de datos y el reporte original.")

st.info(
    "Coupons se carga en la hoja “Raw data” y Promotional Models "
    "se carga en la hoja “S.O.”. No se utilizan archivos CSV."
)

coupons_file = st.file_uploader(
    "Coupons (.xls o .xlsx)",
    type=["xls", "xlsx"]
)

promo_file = st.file_uploader(
    "Promotional Models (.xls o .xlsx)",
    type=["xls", "xlsx"]
)

report_file = st.file_uploader(
    "Reporte original (.xlsx)",
    type=["xlsx"]
)


def is_xls(uploaded_file):
    return uploaded_file.name.lower().endswith(".xls")


def value_for_excel(value):
    """Convierte valores de pandas a valores que Excel puede guardar."""
    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    return value


def iso_week_number(value):
    """Devuelve la semana ISO de una fecha; si no puede leerla, devuelve vacío."""
    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        if isinstance(value, datetime):
            date_value = value
        else:
            date_value = pd.to_datetime(
                value,
                dayfirst=True,
                errors="coerce"
            )

        if pd.isna(date_value):
            return ""

        return int(date_value.isocalendar().week)

    except (TypeError, ValueError, AttributeError):
        return ""


def get_matching_sheet(sheet_names, text_to_find):
    """Devuelve una hoja que contenga el texto; si no, la primera."""
    if not sheet_names:
        raise ValueError("El archivo no contiene ninguna hoja.")

    for name in sheet_names:
        if text_to_find.casefold() in name.casefold():
            return name

    return sheet_names[0]


def read_xls(uploaded_file, preferred_sheet_text=""):
    """
    Lee un .xls incluso si xlrd detecta una corrupción menor.
    Devuelve dataframe y nombre de hoja empleado.
    """
    uploaded_file.seek(0)

    book = xlrd.open_workbook(
        file_contents=uploaded_file.read(),
        ignore_workbook_corruption=True
    )

    sheet_name = get_matching_sheet(
        book.sheet_names(),
        preferred_sheet_text
    )

    sheet_index = book.sheet_names().index(sheet_name)
    sheet = book.sheet_by_index(sheet_index)

    if sheet.nrows == 0:
        raise ValueError(
            f'La hoja "{sheet_name}" del archivo está vacía.'
        )

    headers = [
        str(value).strip()
        for value in sheet.row_values(0)
    ]

    rows = [
        sheet.row_values(row_number)
        for row_number in range(1, sheet.nrows)
    ]

    dataframe = pd.DataFrame(rows, columns=headers)

    return dataframe, sheet_name


def read_xlsx(uploaded_file, preferred_sheet_text=""):
    """Lee un .xlsx y devuelve dataframe y nombre de hoja empleado."""
    uploaded_file.seek(0)

    excel_file = pd.ExcelFile(
        uploaded_file,
        engine="openpyxl"
    )

    sheet_name = get_matching_sheet(
        excel_file.sheet_names,
        preferred_sheet_text
    )

    uploaded_file.seek(0)

    dataframe = pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name,
        engine="openpyxl"
    )

    return dataframe, sheet_name


def read_excel(uploaded_file, preferred_sheet_text=""):
    """Lee un .xls o .xlsx."""
    if is_xls(uploaded_file):
        return read_xls(uploaded_file, preferred_sheet_text)

    return read_xlsx(uploaded_file, preferred_sheet_text)


def find_column(dataframe, expected_name):
    """Encuentra una columna ignorando espacios y mayúsculas."""
    target = expected_name.strip().casefold()

    for column in dataframe.columns:
        if str(column).strip().casefold() == target:
            return column

    return None


def prepare_coupons(dataframe):
    """
    Añade Num de SEM basada en Fecha Compra
    y coloca esa columna en AC (posición 29).
    """
    dataframe = dataframe.copy()

    purchase_date_column = find_column(
        dataframe,
        "Fecha Compra"
    )

    if purchase_date_column is None:
        column_names = ", ".join(
            str(column)
            for column in dataframe.columns
        )

        raise ValueError(
            'No se encontró la columna "Fecha Compra" en Coupons. '
            f"Columnas detectadas: {column_names}"
        )

    dataframe["Num de SEM"] = dataframe[
        purchase_date_column
    ].apply(iso_week_number)

    columns = [
        column
        for column in dataframe.columns
        if column != "Num de SEM"
    ]

    # AC es la columna número 29 de Excel; en Python el índice es 28.
    columns.insert(
        min(28, len(columns)),
        "Num de SEM"
    )

    dataframe = dataframe[columns]

    return dataframe


def clear_sheet(worksheet):
    """Borra todo el contenido actual de una hoja."""
    if worksheet.max_row > 0:
        worksheet.delete_rows(
            1,
            worksheet.max_row
        )


def write_dataframe(worksheet, dataframe):
    """
    Limpia la hoja y escribe encabezados + filas del DataFrame.
    """
    clear_sheet(worksheet)

    for column_number, column_name in enumerate(
        dataframe.columns,
        start=1
    ):
        worksheet.cell(
            row=1,
            column=column_number,
            value=str(column_name)
        )

    for row_number, row_values in enumerate(
        dataframe.itertuples(index=False, name=None),
        start=2
    ):
        for column_number, value in enumerate(
            row_values,
            start=1
        ):
            worksheet.cell(
                row=row_number,
                column=column_number,
                value=value_for_excel(value)
            )


if st.button("Procesar y descargar reporte", type="primary"):
    missing_files = []

    if coupons_file is None:
        missing_files.append("Coupons")

    if promo_file is None:
        missing_files.append("Promotional Models")

    if report_file is None:
        missing_files.append("Reporte original")

    if missing_files:
        st.error(
            "Faltan estos archivos: "
            + ", ".join(missing_files)
        )
        st.stop()

    try:
        with st.spinner("Procesando archivos..."):
            # 1. Coupons -> Raw data
            coupons_df, coupons_sheet = read_excel(
                coupons_file,
                preferred_sheet_text="coupon"
            )

            coupons_df = prepare_coupons(coupons_df)

            # 2. Promotional Models -> S.O.
            promo_df, promo_sheet = read_excel(promo_file)

            # 3. Abrir el reporte original
            report_file.seek(0)

            workbook = load_workbook(report_file)

            if "Raw data" not in workbook.sheetnames:
                available_sheets = ", ".join(workbook.sheetnames)

                raise ValueError(
                    'El reporte no contiene la hoja "Raw data". '
                    f"Hojas disponibles: {available_sheets}"
                )

            if "S.O." not in workbook.sheetnames:
                available_sheets = ", ".join(workbook.sheetnames)

                raise ValueError(
                    'El reporte no contiene la hoja "S.O.". '
                    f"Hojas disponibles: {available_sheets}"
                )

            # 4. Actualizar únicamente las dos hojas requeridas
            write_dataframe(
                workbook["Raw data"],
                coupons_df
            )

            # Liberamos parte de la memoria antes de seguir.
            coupons_rows = len(coupons_df)
            del coupons_df
            gc.collect()

            write_dataframe(
                workbook["S.O."],
                promo_df
            )

            promo_rows = len(promo_df)
            del promo_df
            gc.collect()

            # 5. Guardar el resultado
            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)

            del workbook
            gc.collect()

        st.success("Reporte actualizado correctamente.")

        st.caption(
            f'Coupons: hoja "{coupons_sheet}" → Raw data. '
            f'Promotional Models: hoja "{promo_sheet}" → S.O.'
        )

        col1, col2 = st.columns(2)

        col1.metric(
            "Filas cargadas en Raw data",
            coupons_rows
        )

        col2.metric(
            "Filas cargadas en S.O.",
            promo_rows
        )

        st.download_button(
            label="Descargar reporte actualizado (.xlsx)",
            data=output.getvalue(),
            file_name="reporte_actualizado.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary"
        )

    except Exception as error:
        st.error(f"Error al procesar los archivos: {error}")
        st.exception(error)
