# Ventas de marcas Febeca — Dashboard en Streamlit

Aplicación Streamlit que reemplaza el dashboard HTML/Excel: mismas reglas de negocio
(Clientes Activados = SUMA, Rotación = PROMEDIO, semana→mes por regla ISO del jueves,
Presupuesto vs Real / Ritmo, orden de Acciones por presupuesto descendente, etc.),
ahora como una app web interactiva.

## Contenido de la carpeta

```
streamlit_app/
├── app.py                  # aplicación completa (única página)
├── requirements.txt        # dependencias
├── README.md                # este archivo
└── data/
    ├── raw_clean.csv         # ventas semanales por marca (dataset por defecto)
    ├── presupuesto_raw.csv   # presupuesto mensual por marca (dataset por defecto)
    └── acciones_raw.csv      # acciones comerciales por marca (dataset por defecto)
```

Los 3 archivos CSV dentro de `data/` son el dataset que se carga automáticamente
al abrir la app (equivalente a los datos actuales de "Ventas actualizadas.xlsx").
No hace falta subir nada para empezar a usar el dashboard: ya trae información real.

## Cómo actualizar los datos

Desde la barra lateral ("Actualizar datos") se puede subir un archivo `.xlsx` con
el mismo formato de siempre:

- Hoja `Sheet1`: columnas Marca, Semana, Grupo compra, Clientes activados,
  Contribución, Rotación, Venta Neta.
- Hoja `Presupuesto`: columna Marca + una columna por mes.
- Hoja `Acciones`: columnas Marca y Acción (con celdas combinadas si una marca
  tiene varias acciones).

**El archivo que siempre se debe cargar se llama "Ventas actualizadas.xlsx"** y su
contenido reemplaza por completo los datos anteriores (ventas, presupuesto y
acciones), tal como se acordó. Si alguna hoja no se puede leer, la app avisa con
un mensaje y conserva la información anterior de esa hoja únicamente.

A diferencia de la versión HTML, aquí el archivo se procesa en el servidor con
Python/openpyxl (no en el navegador), así que no depende de que el navegador del
usuario pueda descargar una librería externa — es más robusto y funciona igual
para todos los que abran el link de la app.

## Probar en tu computadora (opcional, antes de publicar)

1. Instalar Python 3.10 o superior.
2. Dentro de la carpeta `streamlit_app`, instalar dependencias:

   ```
   pip install -r requirements.txt
   ```

3. Ejecutar la app:

   ```
   streamlit run app.py
   ```

4. Se abre automáticamente en el navegador (normalmente `http://localhost:8501`).

## Cómo publicarlo en Streamlit Community Cloud (gratis)

Streamlit Community Cloud es la forma más simple de tener el dashboard disponible
con un link público, sin pagar hosting. Pasos:

1. **Crear una cuenta de GitHub** (si no tienes una): [github.com](https://github.com).

2. **Crear un repositorio nuevo** en GitHub, por ejemplo `dashboard-marcas-febeca`.
   Puede ser privado (Streamlit Community Cloud permite conectar repos privados).

3. **Subir el contenido de esta carpeta** (`app.py`, `requirements.txt`, `README.md`
   y la carpeta `data/`) a ese repositorio. Hay dos formas:
   - Desde la web de GitHub: botón "Add file" → "Upload files", arrastrar todos
     los archivos y carpetas, y confirmar el commit.
   - Desde la computadora con git instalado:
     ```
     cd streamlit_app
     git init
     git add .
     git commit -m "Dashboard Ventas de marcas Febeca"
     git branch -M main
     git remote add origin https://github.com/TU_USUARIO/dashboard-marcas-febeca.git
     git push -u origin main
     ```

4. **Entrar a [share.streamlit.io](https://share.streamlit.io)** e iniciar sesión
   con la cuenta de GitHub (Streamlit Community Cloud es un servicio de Streamlit/
   Snowflake; la cuenta se crea gratis con el mismo login de GitHub).

5. Hacer clic en **"New app"** (o "Create app").

6. Seleccionar:
   - **Repository**: el repositorio que se creó en el paso 2.
   - **Branch**: `main`.
   - **Main file path**: `app.py`.

7. Hacer clic en **"Deploy"**. Streamlit instala automáticamente las dependencias
   listadas en `requirements.txt` y publica la app. El primer despliegue tarda
   unos 1-3 minutos.

8. Al terminar, la plataforma entrega un link público con el formato:

   ```
   https://TU-APP.streamlit.app
   ```

   Ese es el link que se puede compartir con el equipo. Cualquier persona con el
   link puede abrir el dashboard sin instalar nada.

### Actualizaciones futuras del código

Cualquier cambio que se suba al repositorio de GitHub (por ejemplo un nuevo
`app.py`) se refleja automáticamente en la app publicada — Streamlit Community
Cloud vuelve a desplegar solo cuando detecta un nuevo commit en la rama conectada.

### Actualizar los datos por defecto sin tocar código

Si se quiere que el dataset "de fábrica" (el que ve cualquiera que entra sin subir
nada) sea otro, basta con reemplazar los 3 archivos dentro de `data/` en el
repositorio de GitHub y esperar el redeploy automático. Para actualizaciones del
día a día no hace falta tocar el repositorio: se usa el uploader de la barra
lateral con el archivo "Ventas actualizadas.xlsx", como ya se hacía en la versión
HTML.

## Notas sobre reglas de negocio (sin cambios respecto al Excel/HTML)

- Semana → mes: se usa el jueves de cada semana ISO para decidir a qué mes
  pertenece (evita que semanas a caballo entre dos meses queden mal asignadas).
- Clientes Activados siempre se **suma** (nunca se promedia).
- Rotación siempre se **promedia** (nunca se suma).
- Venta y Contribución del archivo vienen en miles y se multiplican por 1000 para
  mostrarse en dólares.
- Presupuesto vs Real y Ritmo siempre se calculan sobre el mes en curso completo,
  sin importar el filtro de semana seleccionado en pantalla.
- Acciones se muestran ordenadas por presupuesto del mes en curso de mayor a
  menor; las marcas sin presupuesto asignado aparecen al final, en orden
  alfabético. Las marcas sin ninguna acción registrada se muestran con la celda
  en blanco (no con un texto de relleno).
