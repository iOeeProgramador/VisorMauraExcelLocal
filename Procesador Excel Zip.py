import streamlit as st
import pandas as pd
import zipfile
import io
from datetime import datetime
import re

st.set_page_config(layout="wide")
st.title("Procesador de archivos MIA")

modo = st.radio("Selecciona el modo de operación:", ("Actualizar con ZIP", "Revisar DatosCombinados.xlsx", "Actualizar desde Responsable"))

uploaded_file = None  # Garantiza que siempre esté definida, para evitar NameError más abajo

st.write("✅ Cargando aplicación...")

if modo == "Actualizar desde Responsable":
    datos_file = st.file_uploader("Carga el archivo DatosCombinados.xlsx actual", type="xlsx", key="datos_file")
    responsable_file = st.file_uploader("Carga el archivo Excel del Responsable actualizado", type="xlsx", key="responsable_file")

    if datos_file and responsable_file:
        df_combinado = pd.read_excel(datos_file)
        df_update = pd.read_excel(responsable_file)

        backup = io.BytesIO()
        with pd.ExcelWriter(backup, engine='xlsxwriter') as writer:
            df_combinado.to_excel(writer, index=False, sheet_name='Datos')
        backup.seek(0)
        st.download_button(
            label="Descargar respaldo antes de la actualización",
            data=backup,
            file_name="Respaldo_DatosCombinados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if "LORD_ORDENES" in df_update.columns and "LLINE_ORDENES" in df_update.columns:
            df_update["KEY"] = df_update["LORD_ORDENES"].astype(str) + df_update["LLINE_ORDENES"].astype(str)
            df_combinado["KEY"] = df_combinado["LORD_ORDENES"].astype(str) + df_combinado["LLINE_ORDENES"].astype(str)

            df_combinado.set_index("KEY", inplace=True)
            df_update.set_index("KEY", inplace=True)

            for col in ["ESTADO_ESTADO", "OBSERVACION_ESTADO"]:
                if col in df_update.columns:
                    df_combinado.loc[df_update.index, col] = df_update[col]

            df_combinado.reset_index(inplace=True)
            st.success("Datos actualizados correctamente desde el archivo del responsable.")

            output_actualizado = io.BytesIO()
            with pd.ExcelWriter(output_actualizado, engine='xlsxwriter') as writer:
                df_combinado.to_excel(writer, index=False, sheet_name='Datos')
            output_actualizado.seek(0)
            st.download_button(
                label="Descargar DatosCombinados actualizado",
                data=output_actualizado,
                file_name="DatosCombinados_actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            tab1, tab2, tab3 = st.tabs(["Vista previa", "Resumen por Responsable", "Resumen por Estado"])

            with tab1:
                st.subheader("Vista previa de DatosCombinados.xlsx")
                filtro = st.text_input("Filtrar por texto (cualquier columna):", key="filtro_actualizar")
                if filtro:
                    df_filtrado = df_combinado[df_combinado.apply(lambda row: row.astype(str).str.contains(filtro, case=False).any(), axis=1)]
                    st.dataframe(df_filtrado, use_container_width=True)
                else:
                    st.dataframe(df_combinado, use_container_width=True)

            if "RESPONSABLE_GESTION" in df_combinado.columns and not df_combinado.empty:
                resumen = df_combinado.groupby("RESPONSABLE_GESTION", dropna=False).size().reset_index(name="Total Líneas")
                resumen["RESPONSABLE_GESTION"] = resumen["RESPONSABLE_GESTION"].fillna("SIN RESPONSABLE")
                resumen = resumen.sort_values(by="Total Líneas", ascending=False)
                total = resumen["Total Líneas"].sum()
                with tab2:
                    st.subheader(f"Resumen Total de Líneas por Responsable (Total: {total})")
                    st.dataframe(resumen, use_container_width=True)
                    resumen_xlsx = io.BytesIO()
                    with pd.ExcelWriter(resumen_xlsx, engine='xlsxwriter') as writer:
                        resumen.to_excel(writer, index=False, sheet_name='Resumen')
                    resumen_xlsx.seek(0)
                    st.download_button(
                        label="Descargar resumen por Responsable (Excel)",
                        data=resumen_xlsx,
                        file_name="ResumenResponsable.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            if "RESPONSABLE_GESTION" in df_combinado.columns and "ESTADO_ESTADO" in df_combinado.columns and not df_combinado.empty:
                pivot_resp_estado = df_combinado.pivot_table(
                    index="RESPONSABLE_GESTION",
                    columns="ESTADO_ESTADO",
                    aggfunc="size",
                    fill_value=0
                ).reset_index()
                with tab3:
                    st.subheader("Resumen Total de Líneas por Responsable y Estado")
                    st.dataframe(pivot_resp_estado, use_container_width=True)

if modo == "Actualizar con ZIP":
    uploaded_file = st.file_uploader("Carga tu archivo ZIP con los libros de Excel", type="zip")
elif modo == "Revisar DatosCombinados.xlsx":
    uploaded_file = st.file_uploader("Carga tu archivo Excel DatosCombinados.xlsx", type="xlsx")

    if uploaded_file is not None:
        df_combinado = pd.read_excel(uploaded_file)

        tab1, tab2, tab3 = st.tabs(["Vista previa", "Resumen por Responsable", "Resumen por Estado"])

        with tab1:
            st.subheader("Vista previa de DatosCombinados.xlsx")
            filtro = st.text_input("Filtrar por texto (cualquier columna):", key="filtro_revisar")
            if filtro:
                df_filtrado = df_combinado[df_combinado.apply(lambda row: row.astype(str).str.contains(filtro, case=False).any(), axis=1)]
                st.dataframe(df_filtrado, use_container_width=True)
            else:
                st.dataframe(df_combinado, use_container_width=True)

        if "RESPONSABLE_GESTION" in df_combinado.columns and not df_combinado.empty:
            resumen = df_combinado.groupby("RESPONSABLE_GESTION", dropna=False).size().reset_index(name="Total Líneas")
            resumen["RESPONSABLE_GESTION"] = resumen["RESPONSABLE_GESTION"].fillna("SIN RESPONSABLE")
            resumen = resumen.sort_values(by="Total Líneas", ascending=False)
            total = resumen["Total Líneas"].sum()
            with tab2:
                st.subheader(f"Resumen Total de Líneas por Responsable (Total: {total})")
                st.dataframe(resumen, use_container_width=True)

        if "RESPONSABLE_GESTION" in df_combinado.columns and "ESTADO_ESTADO" in df_combinado.columns and not df_combinado.empty:
            pivot_resp_estado = df_combinado.pivot_table(
                index="RESPONSABLE_GESTION",
                columns="ESTADO_ESTADO",
                aggfunc="size",
                fill_value=0
            ).reset_index()
            with tab3:
                st.subheader("Resumen Total de Líneas por Responsable y Estado")
                st.dataframe(pivot_resp_estado, use_container_width=True)

if uploaded_file is not None and modo == "Actualizar con ZIP":
    if uploaded_file.size == 0:
        st.error("⚠️ El archivo ZIP está vacío.")

    tab1, tab2, tab3 = st.tabs(["Vista previa", "Resumen por Responsable", "Resumen por Estado"])
    with zipfile.ZipFile(uploaded_file) as z:
        expected_files = ["ORDENES.xlsx", "INVENTARIO.xlsx", "ESTADO.xlsx", "PRECIOS.xlsx", "GESTION.xlsx"]
        file_dict = {name: z.open(name) for name in expected_files if name in z.namelist()}

        if "ORDENES.xlsx" in file_dict:
            df_ordenes = pd.read_excel(file_dict["ORDENES.xlsx"])
            df_ordenes.columns = [f"{col}_ORDENES" for col in df_ordenes.columns]

            if "LRDTE_ORDENES" in df_ordenes.columns:
                today = datetime.today()
                df_ordenes.insert(0, "CONTROL_DIAS", df_ordenes["LRDTE_ORDENES"].apply(
                    lambda x: (datetime.strptime(str(int(x)), "%Y%m%d") - today).days))

            if "INVENTARIO.xlsx" in file_dict:
                df_inventario = pd.read_excel(file_dict["INVENTARIO.xlsx"])
                df_inventario.columns = [f"{col}_INVENTARIO" for col in df_inventario.columns]
                df_inventario_unique = df_inventario.drop_duplicates(subset=["Cod. Producto_INVENTARIO"])
                df_combinado = pd.merge(df_ordenes, df_inventario_unique, left_on="LPROD_ORDENES", right_on="Cod. Producto_INVENTARIO", how="left")
            else:
                df_combinado = df_ordenes

            if "ESTADO.xlsx" in file_dict:
                df_estado = pd.read_excel(file_dict["ESTADO.xlsx"])
                df_estado.columns = [f"{col}_ESTADO" for col in df_estado.columns]
                df_combinado["KEY_ORDENES"] = df_combinado["LORD_ORDENES"].astype(str) + df_combinado["LLINE_ORDENES"].astype(str)
                df_estado["KEY_ESTADO"] = df_estado["LORD_ESTADO"].astype(str) + df_estado["LLINE_ESTADO"].astype(str)
                df_estado_unique = df_estado.drop_duplicates(subset=["KEY_ESTADO"])
                df_combinado = pd.merge(df_combinado, df_estado_unique, left_on="KEY_ORDENES", right_on="KEY_ESTADO", how="left")

            if "PRECIOS.xlsx" in file_dict:
                df_precios = pd.read_excel(file_dict["PRECIOS.xlsx"])
                df_precios.columns = [f"{col}_PRECIOS" for col in df_precios.columns]
                df_precios_unique = df_precios.drop_duplicates(subset=["LPROD_PRECIOS"])
                for col in ["VALOR_PRECIOS", "On Hand_PRECIOS"]:
                    if col in df_precios_unique.columns:
                        df_precios_unique[col] = pd.to_numeric(df_precios_unique[col], errors='coerce').fillna(0).astype(int)
                df_combinado = pd.merge(df_combinado, df_precios_unique, left_on="LPROD_ORDENES", right_on="LPROD_PRECIOS", how="left")

            if "GESTION.xlsx" in file_dict:
                df_gestion = pd.read_excel(file_dict["GESTION.xlsx"])
                df_gestion.columns = [f"{col}_GESTION" for col in df_gestion.columns]
                df_gestion_unique = df_gestion.drop_duplicates(subset=["HNAME_GESTION"])
                df_combinado = pd.merge(df_combinado, df_gestion_unique, left_on="HNAME_ORDENES", right_on="HNAME_GESTION", how="left")

            with tab1:
                st.subheader("Vista previa de DatosCombinados.xlsx")
                filtro = st.text_input("Filtrar por texto (cualquier columna):")
                if filtro:
                    df_filtrado = df_combinado[df_combinado.apply(lambda row: row.astype(str).str.contains(filtro, case=False).any(), axis=1)]
                    st.dataframe(df_filtrado, use_container_width=True)
                else:
                    st.dataframe(df_combinado, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_combinado.to_excel(writer, index=False, sheet_name='Datos')
                output.seek(0)

                st.success("Archivo DatosCombinados.xlsx generado con éxito")
                st.download_button(
                    label="Descargar DatosCombinados.xlsx",
                    data=output,
                    file_name="DatosCombinados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                if st.button("Generar ZIP por Responsable"):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zipf:
                        responsables = df_combinado["RESPONSABLE_GESTION"].fillna("SIN RESPONSABLE").unique()
                        for responsable in responsables:
                            df_responsable = df_combinado[df_combinado["RESPONSABLE_GESTION"] == responsable]
                            output_excel = io.BytesIO()
                            with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
                                df_responsable.to_excel(writer, index=False, sheet_name="Datos")
                                worksheet = writer.sheets["Datos"]
                                for col_num, _ in enumerate(df_responsable.columns):
                                    worksheet.set_column(col_num, col_num, 20, writer.book.add_format({"align": "center", "valign": "vcenter"}))
                                worksheet.autofilter(0, 0, len(df_responsable), len(df_responsable.columns) - 1)
                            output_excel.seek(0)
                            safe_name = re.sub(r'[^a-zA-Z0-9_-]', '_', str(responsable))
                            safe_name = f"{safe_name}_{datetime.today().strftime('%Y%m%d')}"
                            zipf.writestr(f"{safe_name}.xlsx", output_excel.read())
                    zip_buffer.seek(0)
                    st.success("ZIP por Responsable generado con éxito")
                    st.download_button(
                        label="Descargar ZIP con Datos por Responsable",
                        data=zip_buffer,
                        file_name="DatosPorResponsable.zip",
                        mime="application/zip"
                    )

            if "RESPONSABLE_GESTION" in df_combinado.columns and not df_combinado.empty:
                resumen = df_combinado.groupby("RESPONSABLE_GESTION", dropna=False).size().reset_index(name="Total Líneas")
                resumen["RESPONSABLE_GESTION"] = resumen["RESPONSABLE_GESTION"].fillna("SIN RESPONSABLE")
                resumen = resumen.sort_values(by="Total Líneas", ascending=False)
                total = resumen["Total Líneas"].sum()
                with tab2:
                    st.subheader(f"Resumen Total de Líneas por Responsable (Total: {total})")
                    st.dataframe(resumen, use_container_width=True)

            if "RESPONSABLE_GESTION" in df_combinado.columns and "ESTADO_ESTADO" in df_combinado.columns and not df_combinado.empty:
                pivot_resp_estado = df_combinado.pivot_table(
                    index="RESPONSABLE_GESTION",
                    columns="ESTADO_ESTADO",
                    aggfunc="size",
                    fill_value=0
                ).reset_index()
                with tab3:
                    st.subheader("Resumen Total de Líneas por Responsable y Estado")
                    st.dataframe(pivot_resp_estado, use_container_width=True)
