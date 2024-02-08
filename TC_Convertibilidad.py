# Importación de librerías.
import requests
import pdfplumber
import json
import pandas as pd

# Descarga de datos del BCRA.
url = "https://www.bcra.gob.ar/Pdfs/PublicacionesEstadisticas/infomondiae.pdf"

response = requests.get(url)

with open("InformeBCRA.pdf", "wb") as f:
    f.write(response.content)

# Se busca el precio del dólar blue y oficial.
response = requests.get("https://api.bluelytics.com.ar/v2/evolution.json")

data = json.loads(response.text)
df = pd.DataFrame(data)
registro_blue = df[df["source"] == "Blue"].iloc[0]
dolar_blue = (registro_blue["value_sell"] + registro_blue["value_buy"])/2
registro_oficial = df[df["source"] == "Oficial"].iloc[0]
dolar_oficial = (registro_oficial["value_sell"] + registro_oficial["value_buy"])/2

# Abre el archivo PDF con pdfplumber. Se extraen tablas y datos de estas.
with pdfplumber.open("InformeBCRA.pdf") as pdf:
    pages = pdf.pages
    tables = pages[3].extract_tables()[0]

    base_monetaria = float(tables[22][0].split()[2].replace(",", ""))
    leliqs = float(tables[26][0].split()[4].replace(",", ""))
    legar = float(tables[28][0].split()[4].replace(",", ""))
    pases = float(tables[30][0].split()[2].replace(",", ""))
    reservas = float(tables[36][0].split("\n")[-2].split()[3].replace(",", ""))
    adelantos = float(tables[35][0].split()[5].replace(",", ""))
    fecha = tables[20][2]
    fecha_real = tables[0][0]

    agregado1 = base_monetaria + leliqs + legar + pases
    agregado2 = base_monetaria + leliqs
    agregado3 = base_monetaria + legar + pases

    tables = pages[4].extract_tables()[-1]
    M3 = float(tables[34][0].split()[-11].replace(",", ""))
    M2 = float(tables[33][0].split()[-9].replace(",", ""))
    M1 = float(tables[32][0].split()[-11].replace(",", ""))

    # Se imprimen por pantalla los resultados.
    print(fecha_real + f"  --  Saldos al {fecha}")
    print()
    print(f"Dólar de convertibilidad = $ {agregado1 / reservas:.2f}")
    print(
        f"Diferencia contra blue = $ {agregado1 / reservas - dolar_blue:.2f}  {((agregado1 / reservas - dolar_blue) / dolar_blue * 100):.2f} %")
    print()
    print(f"Dólar Blue = $ {dolar_blue}")
    print()
    print(f"Dólar \"BNA\" = $ {agregado3 / reservas:.2f}")
    print(f"Dólar BNA = $ {dolar_oficial}")
    print(f"Diferencia = $ {agregado3 / reservas - dolar_oficial:.2f}  {((agregado3 / reservas - dolar_oficial) / dolar_oficial) * 100:.2f} %")
    print()
    print(f"Dólar \"BANCOS\" = $ {agregado2 / reservas:.2f}")
    print()
    print(f"Multiplicador (M1): {M1 / base_monetaria:.2f}")
    print(f"Multiplicador (M2): {M2 / base_monetaria:.2f}")
    print(f"Multiplicador (M3): {M3 / base_monetaria:.2f}")
    print()
    print(f"Pasivos remunerados en términos de la BM: {(agregado1 - base_monetaria) / base_monetaria:.2f} veces")
    print()
    print(f"Base Monetaria = $ {base_monetaria:_.0f}")
    print(f"Adelantos al sector público = $ {adelantos:_.0f}")
    print(f"Leliqs = $ {leliqs:_.0f}")
    print(f"Otras letras = $ {legar:_.0f}")
    print(f"Pases pasivos = $ {pases:_.0f}")
    print(f"Reservas = U$S {reservas:_.0f}")
