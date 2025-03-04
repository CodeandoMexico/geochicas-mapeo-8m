import pandas as pd
import csv
from datetime import datetime
import requests
from urllib.parse import urlparse
import os
import validators

def download_image(image_url, save_path):
    if not validators.url(image_url):
        return
    try:
        response = requests.get(image_url)
        if response.status_code == 200:
            with open(save_path, 'wb') as file:
                file.write(response.content)
            print(f"Image saved to {save_path}")
        else:
            print(f"Failed to download image. Status code: {response.status_code}")
    except requests.ConnectionError as err:
        print(f"Connection error occurred: {err}")


file_path = 'data/encuesta_20250303.xlsx'

organizaciones = pd.read_excel(file_path, sheet_name='GeoChicas 8M 2025')
actividades = pd.read_excel(file_path, sheet_name='actividades')

cabecera = [
    'indice',
    'convocatoria',
    'colectiva_nombre',
    'colectiva_url',
    'pais',
    'ciudad',
    'actividad_fecha',
    'actividad_hora',
    'actividad_localizacion',
    'actividad_localizacion_latitude',
    'actividad_localizacion_longitude',
    'actividad_localizacion_altitude',
    'actividad_localizacion_precision',
    'actividad_direccion',
    'actividad_url_imagen',
    'actividad_url_convocatoria'
]


today = datetime.now().date().isoformat().replace('-', '')
with open('data/actividades_' + today + '.csv', 'w', newline='') as file:
    writer = csv.writer(file, delimiter=',', quoting=csv.QUOTE_ALL)
    writer.writerow(cabecera)

    for i, actividad in actividades.iterrows():

        for j, organizacion in organizaciones.iterrows():
            if organizacion._uuid == actividad._submission__uuid:
                break

        row = [
            actividad._index,
            '2025',
            organizacion.colectiva_nombre if not pd.isna(organizacion.colectiva_nombre) else '',
            organizacion.colectiva_url if not pd.isna(organizacion.colectiva_url) else '',
            organizacion.pais if not pd.isna(organizacion.pais) else '',
            organizacion.ciudad if not pd.isna(organizacion.ciudad) else '',
            actividad.actividad_fecha.date(),
            actividad.actividad_fecha.time(),
            actividad.actividad_localizacion,
            actividad._actividad_localizacion_latitude,
            actividad._actividad_localizacion_longitude,
            actividad._actividad_localizacion_altitude,
            actividad._actividad_localizacion_precision,
            actividad.actividad_direccion if not pd.isna(actividad.actividad_direccion) else '',
            "{{" + actividad.actividad_url_imagen + "}}" if not pd.isna(actividad.actividad_url_imagen) else '',
            actividad.actividad_url_convocatoria if not pd.isna(actividad.actividad_url_convocatoria) else ''
        ]

        writer.writerow(row)

        if False: # not pd.isna(actividad.actividad_url_imagen):
            print('Downloading: ' + actividad.actividad_url_imagen)
            parsed_url = urlparse(actividad.actividad_url_imagen)
            file_name = os.path.basename(parsed_url.path)
            if file_name != '':
                download_image(image_url=actividad.actividad_url_imagen, save_path='images/' + file_name)
