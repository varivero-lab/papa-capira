# Simulador MDP · papa capira con actualización automática (DANE – SIPSA)

Contenido:
- `index.html`: el simulador. Si encuentra `datos.json` en la misma carpeta, lo carga solo.
- `actualizar_sipsa.py`: descarga el anexo más reciente del DANE, extrae la papa capira (Central Mayorista de Antioquia) y actualiza `datos.json`.
- `.github/workflows/actualizar.yml`: ejecuta el script cada sábado en GitHub Actions y guarda los cambios.

Primer uso (rellenar el histórico): en GitHub, pestaña Actions → "Actualizar precios SIPSA" → Run workflow (cada ejecución revisa también el histórico del año),
o en tu computador: `pip install requests openpyxl` y `python actualizar_sipsa.py --anio 2026`.
