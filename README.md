# 🏗️ Diseño de Zapatas ACI 318-19 - XSYM

Herramienta web interactiva para el diseño estructural de zapatas de cimentación (aisladas, excéntricas, esquinadas y combinadas) conforme al reglamento **ACI 318-19**.

## 🚀 Características
- Cálculo automático de esfuerzos, cortante, punzonamiento y flexión.
- Diagramas interactivos de presiones, momento y cortante.
- Metrado de acero detallado y verificación de longitud de desarrollo.
- Exportación de memoria de cálculo a Excel (.xlsx).
- Interfaz web moderna (no requiere instalación de escritorio).

## 🛠️ Tecnologías
- **Frontend/Backend:** Python + Streamlit
- **Cálculo:** NumPy, Math
- **Gráficos:** Matplotlib
- **Reportes:** OpenPyXL

## 🌐 Despliegue en la Web (Gratis)
1. Sube este repositorio a tu cuenta de GitHub.
2. Ve a [share.streamlit.io](https://share.streamlit.io/) e inicia sesión con GitHub.
3. Haz clic en **"New app"**, selecciona este repositorio y la rama `main`.
4. El archivo principal es `app.py`. Haz clic en **"Deploy"**.
5. ¡Listo! Tu calculadora estará disponible en una URL pública gratuita.

## 📦 Ejecución Local
Si deseas probarlo en tu computadora:
```bash
pip install -r requirements.txt
streamlit run app.py