"""
Descarga automáticamente los anexos del Boletín semanal de precios mayoristas (DANE - SIPSA),
extrae el precio de la papa capira en Medellín, Central Mayorista de Antioquia,
y lo agrega a datos.json (sin duplicar semanas).

Uso:
  python actualizar_sipsa.py                 # revisa el boletín más reciente
  python actualizar_sipsa.py --anio 2026     # además recorre el histórico de 2026 (rellena semanas faltantes)
  python actualizar_sipsa.py --archivo x.xlsx  # procesa un archivo local (para pruebas)
"""
import argparse, io, json, re, sys, time, unicodedata
from pathlib import Path
from urllib.parse import urljoin

import openpyxl
import requests

PAGINA = "https://www.dane.gov.co/index.php/estadisticas-por-tema/agropecuario/sistema-de-informacion-de-precios-sipsa/mayoristas-boletin-semanal-1"
HISTORICO = PAGINA + "/boletin-mayorista-semanal-{anio}"
PRODUCTO = "papa capira"
MERCADO = "central mayorista de antioquia"
SALIDA = Path(__file__).with_name("datos.json")
CABECERAS = {"User-Agent": "simulador-mdp-papa-capira (uso academico; 1 consulta semanal)"}

MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
MES3 = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "sep": 9, "set": 9,
        "oct": 10, "nov": 11, "dic": 12}


def norm(t):
    t = unicodedata.normalize("NFD", str(t or "").lower())
    return re.sub(r"\s+", " ", "".join(c for c in t if unicodedata.category(c) != "Mn")).strip()


def semana(titulo, nombre):
    """Devuelve (inicio, fin) en formato AAAA-MM-DD a partir del título del boletín o del nombre del archivo."""
    m = re.search(r"(\d{1,2})(?: de ([a-z]+))?(?: de (\d{4}))? al (\d{1,2}) de ([a-z]+) (?:de|del) (\d{4})", norm(titulo))
    if m and m.group(5) in MESES:
        em, ey = MESES[m.group(5)], int(m.group(6))
        sm = MESES.get(m.group(2), em)
        sy = int(m.group(3)) if m.group(3) else (ey - 1 if sm > em else ey)
        return f"{sy}-{sm:02d}-{int(m.group(1)):02d}", f"{ey}-{em:02d}-{int(m.group(4)):02d}"
    m = re.search(r"(\d{1,2})([a-z]{3})(\d{1,2})([a-z]{3})-(\d{4})", norm(nombre))
    if m and m.group(2) in MES3 and m.group(4) in MES3:
        sm, em, ey = MES3[m.group(2)], MES3[m.group(4)], int(m.group(5))
        return f"{ey - 1 if sm > em else ey}-{sm:02d}-{int(m.group(1)):02d}", f"{ey}-{em:02d}-{int(m.group(3)):02d}"
    return None


def leer_anexo(contenido, nombre):
    wb = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    titulo, fila = None, None
    for hoja in wb.worksheets:
        for vals in hoja.iter_rows(values_only=True):
            vals = list(vals)
            if titulo is None:
                titulo = next((v for v in vals if isinstance(v, str) and "boletin semanal" in norm(v)), None)
            if fila is None:
                for i, v in enumerate(vals[:-1]):
                    if norm(v) == PRODUCTO and MERCADO in norm(vals[i + 1]):
                        nums = [x for x in vals[i + 2:] if isinstance(x, (int, float))]
                        if len(nums) >= 3:
                            fila = nums[:3]
        if titulo and fila:
            break
    if not fila:
        raise ValueError("no se encontró la fila de papa capira en la Central Mayorista de Antioquia")
    fechas = semana(titulo, nombre)
    if not fechas:
        raise ValueError("no se pudo identificar la semana del boletín")
    return {"key": fechas[0], "end": fechas[1], "min": fila[0], "max": fila[1], "medio": fila[2], "archivo": nombre}


def enlaces_anexos(url):
    html = requests.get(url, headers=CABECERAS, timeout=60).text
    rutas = re.findall(r'href="([^"]*anex-SIPSASemanal[^"]*\.xlsx)"', html, flags=re.I)
    return list(dict.fromkeys(urljoin(url, r) for r in rutas))  # sin duplicados, en orden


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--anio", type=int, action="append", default=[])
    ap.add_argument("--archivo", action="append", default=[])
    a = ap.parse_args()

    datos = json.loads(SALIDA.read_text(encoding="utf-8")) if SALIDA.exists() else {"semanas": []}
    por_semana = {w["key"]: w for w in datos["semanas"]}
    vistos = {w.get("archivo") for w in datos["semanas"]}
    nuevas = 0

    fuentes = [(Path(f).name, Path(f).read_bytes()) for f in a.archivo]
    if not a.archivo:
        urls = enlaces_anexos(PAGINA)
        for anio in a.anio:
            urls += enlaces_anexos(HISTORICO.format(anio=anio))
        for url in dict.fromkeys(urls):
            nombre = url.rsplit("/", 1)[-1]
            if nombre in vistos:
                continue
            time.sleep(2)  # ser amables con el servidor del DANE
            r = requests.get(url, headers=CABECERAS, timeout=120)
            r.raise_for_status()
            fuentes.append((nombre, r.content))

    for nombre, contenido in fuentes:
        try:
            w = leer_anexo(contenido, nombre)
        except Exception as e:  # un archivo con formato distinto no detiene el proceso
            print(f"OMITIDO {nombre}: {e}")
            continue
        if w["key"] not in por_semana:
            nuevas += 1
        por_semana[w["key"]] = w
        print(f"OK {nombre}: semana {w['key']} a {w['end']}, precio medio {w['medio']}")

    datos["semanas"] = sorted(por_semana.values(), key=lambda w: w["key"])
    datos["actualizado"] = time.strftime("%Y-%m-%d %H:%M UTC", time.gmtime())
    SALIDA.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{nuevas} semana(s) nueva(s); total {len(datos['semanas'])}")


if __name__ == "__main__":
    sys.exit(main())
