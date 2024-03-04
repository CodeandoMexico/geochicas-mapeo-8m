import pandas as pd
import csv

file_path = 'data/encuesta.xlsx'

organizaciones = pd.read_excel(file_path, sheet_name='Geochicas 8M 2024')
actividades = pd.read_excel(file_path, sheet_name='actividades')

cabecera = [
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

with open('data/actividades.csv', 'w', newline='') as file:
    writer = csv.writer(file, delimiter=',', quoting=csv.QUOTE_ALL)
    writer.writerow(cabecera)

    for i, actividad in actividades.iterrows():

        for j, organizacion in organizaciones.iterrows():
            if organizacion._uuid == actividad._submission__uuid:
                break

        row = [
            '2024',
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
