# App Costo por Equipo v14

## Corrección principal

Esta versión corrige el fallo reportado: al presionar **GENERAR DASHBOARD** aparecían ceros antes de que el informe quedara completamente generado.

La reparación consiste en:
1. Procesar los archivos.
2. Generar el Excel final.
3. Guardar el resultado en `st.session_state`.
4. Forzar una recarga limpia con `st.rerun()`.
5. Mostrar el dashboard solo después de esa recarga.

Así no se renderizan valores intermedios ni etiquetas fantasma.

## Ejecutar

```powershell
python -m pip install -r requirements_costo_por_equipo_v14.txt
python -m streamlit run app_costo_por_equipo_v14_excel_dashboard.py
```

## Nota

La lógica del Excel final no se modifica.
