import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import io

st.set_page_config(page_title="Actualizador de Reportes", page_icon="📊")

st.title("📊 Actualizador de Reportes")
st.markdown("Sube los archivos y descarga el reporte actualizado.")

coupons_file = st.file_uploader("Coupons (.xls o .xlsx)", type=["xls", "xlsx"])
promo_file = st.file_uploader("Promotional Models (.xls o .xlsx)", type=["xls", "xlsx"])
excel_original = st.file_uploader("Excel Original (.xlsx)", type=["xlsx"])
so_file = st.file_uploader("S.O (.xlsx)", type=["xlsx"])
raw_file = st.file_uploader("Raw data (.csv)", type=["csv"])

def iso_week_number(date_value):
    if pd.isna(date_value):
        return ""
    try:
        if isinstance(date_value, str):
            for fmt in ["%d/%m/%Y", "%d-%m-%Y"]:
                try:
                    date_value = datetime.strptime(date_value, fmt)
                    break
                except ValueError:
                    continue
            else:
                date_value = pd.to_datetime(date_value)
        return date_value.isocalendar().week
    except:
        return ""

def procesar_coupons(df):
    if 'Fecha Compra' not in df.columns:
        st.error("No se encontró la columna 'Fecha Compra' en Coupons")
        return None
    
    df['Num de SEM'] = df['Fecha Compra'].apply(iso_week_number)
    
    cols = list(df.columns)
    if 'Num de SEM' in cols:
        cols.remove('Num de SEM')
        cols.insert(27, 'Num de SEM')
        df = df[cols]
    
    return df

if st.button("🚀 Procesar archivos", type="primary"):
    if not all([coupons_file, promo_file, excel_original, so_file, raw_file]):
        st.error("⚠️ Faltan archivos. Sube todos los requeridos.")
    else:
        try:
            with st.spinner("Procesando..."):
                if coupons_file.name.endswith('.xls'):
                    xl = pd.ExcelFile(coupons_file, engine='xlrd')
                else:
                    xl = pd.ExcelFile(coupons_file, engine='openpyxl')
                
                sheet_name = xl.sheet_names[0]
                for name in xl.sheet_names:
                    if 'coupon' in name.lower():
                        sheet_name = name
                        break
                
                if coupons_file.name.endswith('.xls'):
                    coupons_df = pd.read_excel(coupons_file, sheet_name=sheet_name, engine='xlrd')
                else:
                    coupons_df = pd.read_excel(coupons_file, sheet_name=sheet_name, engine='openpyxl')
                
                coupons_df = procesar_coupons(coupons_df)
                if coupons_df is None:
                    st.stop()
                
                if promo_file.name.endswith('.xls'):
                    promo_df = pd.read_excel(promo_file, engine='xlrd')
                else:
                    promo_df = pd.read_excel(promo_file, engine='openpyxl')
                
                so_df = pd.read_excel(so_file, engine='openpyxl')
                raw_df = pd.read_csv(raw_file)
                
                wb = load_workbook(excel_original)
                
                hojas = {
                    'Coupons': coupons_df,
                    'Promotional Models': promo_df,
                    'S.O': so_df,
                    'Raw data': raw_df
                }
                
                for nombre_hoja, df in hojas.items():
                    if nombre_hoja in wb.sheetnames:
                        ws = wb[nombre_hoja]
                        ws.delete_rows(1, ws.max_row)
                        
                        for col_idx, col_name in enumerate(df.columns, 1):
                            ws.cell(row=1, column=col_idx, value=col_name)
                        
                        for row_idx, row in enumerate(df.values, 2):
                            for col_idx, value in enumerate(row, 1):
                                ws.cell(row=row_idx, column=col_idx, value=value)
                    else:
                        ws = wb.create_sheet(nombre_hoja)
                        for col_idx, col_name in enumerate(df.columns, 1):
                            ws.cell(row=1, column=col_idx, value=col_name)
                        for row_idx, row in enumerate(df.values, 2):
                            for col_idx, value in enumerate(row, 1):
                                ws.cell(row=row_idx, column=col_idx, value=value)
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("✅ Archivos procesados correctamente!")
                st.metric("Coupons", f"{len(coupons_df)} filas")
                st.metric("Promotional Models", f"{len(promo_df)} filas")
                
                st.download_button(
                    label="📥 Descargar reporte actualizado",
                    data=output.getvalue(),
                    file_name="reporte_actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)
