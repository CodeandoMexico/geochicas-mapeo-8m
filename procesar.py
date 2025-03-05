import pandas as pd
import datetime
import csv
import os
import re

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
    'actividad_url_convocatoria',
    'actividad_trans_incluyente'
]

excel_files = [f for f in os.listdir('data') if f.endswith('.xlsx') or f.endswith('.xls')]
for file in excel_files:
    file_path = os.path.join('data', file)
    print(f"Processing file: {file_path}")

    try:
        xls = pd.ExcelFile(file_path)
        organizaciones = pd.read_excel(xls, sheet_name=0)  # Primera hoja
        actividades = pd.read_excel(xls, sheet_name=1)  # Segunda hoja

    except Exception as e:
        print(f"Error al leer {file_path}: {e}")
        continue # Ignoramos y procesamos el siguiente

    file_base_name = os.path.splitext(file)[0]  # Quitamos la extensión
    csv_filename = f'data/actividades_{file_base_name}.csv'

    creation_timestamp = os.stat(file_path).st_ctime
    creation_date = datetime.datetime.fromtimestamp(creation_timestamp)
    anyo = creation_date.year

    with open(csv_filename, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file, delimiter=',', quoting=csv.QUOTE_ALL)
        writer.writerow(cabecera)

        for i, actividad in actividades.iterrows():

            for j, organizacion in organizaciones.iterrows():
                if organizacion._uuid == actividad._submission__uuid:
                    break

            row = [
                actividad._index,
                anyo,
                organizacion.colectiva_nombre if not pd.isna(organizacion.colectiva_nombre) else '',
                organizacion.colectiva_url if not pd.isna(organizacion.colectiva_url) else '',
                organizacion.pais if not pd.isna(organizacion.pais) else '',
                organizacion.ciudad if not pd.isna(organizacion.ciudad) else '',
                actividad.actividad_fecha.date() if not pd.isna(actividad.actividad_fecha) else '',
                actividad.actividad_fecha.time(),
                actividad.actividad_localizacion,
                actividad._actividad_localizacion_latitude,
                actividad._actividad_localizacion_longitude,
                actividad._actividad_localizacion_altitude,
                actividad._actividad_localizacion_precision,
                actividad.actividad_direccion if not pd.isna(actividad.actividad_direccion) else '',
                "{{" + actividad.actividad_url_imagen + "}}" if not pd.isna(actividad.actividad_url_imagen) else '',
                actividad.actividad_url_convocatoria if not pd.isna(actividad.actividad_url_convocatoria) else '',
                actividad.actividad_trans_incluyente if not pd.isna(actividad.actividad_trans_incluyente) else ''
            ]

            writer.writerow(row)
