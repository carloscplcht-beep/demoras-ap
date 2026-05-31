# Programa de analisis de las demoras en Atencion Primaria

Version estable: 1.0

## Objetivo

Herramienta web estatica para analizar la accesibilidad y las demoras de agendas de Atencion Primaria a partir de un Excel local. Permite revisar tabla, filtros, indicadores DGAP, graficos, mapa por Zona Basica de Salud e informe ejecutivo imprimible.

## Carga del Excel

1. Abra la aplicacion en GitHub Pages o mediante `index.html`.
2. Pulse `Seleccionar archivo Excel`.
3. Seleccione un archivo `.xlsx` con la estructura esperada.
4. La aplicacion leera la primera hoja disponible y activara tabla, resumen, graficos e informe.

Cabeceras esperadas:

- Area
- Zona
- Centro
- CIAS
- PROFESIONAL
- Tipo visita
- Accesibilidad
- Categoria
- Codigo de centro
- Fecha Primer Hueco Libre
- Fecha Primer Hueco ID
- Fecha Corte

Las columnas `CIAS ID` y `UID`, si existen, se ignoran.

## Metricas DGAP

Los calculos solo incluyen agendas validas: filas con `Accesibilidad` numerica y no negativa.

- `% 0-2 dias`: accesibilidad entre 0 y 2 dias.
- `% 0-3 dias`: accesibilidad entre 0 y 3 dias.
- `% 0-6 dias`: accesibilidad entre 0 y 6 dias.
- `% 7 o mas dias`: accesibilidad igual o superior a 7 dias.
- Tambien se calculan total de agendas validas, demora media, mediana y maxima.

Los indicadores 0-2, 0-3 y 0-6 son acumulativos.

## Informe

La pestana `Informe` genera una vista ejecutiva imprimible adaptada a los filtros activos. Incluye cabecera, fecha de corte mas reciente del subconjunto filtrado, KPIs DGAP, resumen automatico, detalle por categoria, detalle por categoria y tipo de visita, graficos y nota metodologica.

## Mapa por Zona Basica de Salud

La pestana `Graficos` incluye un mapa analitico no geografico por Zona Basica de Salud. Cada burbuja representa una zona:

- El tamano refleja el volumen de agendas validas.
- El color refleja la demora media.
- El panel de detalle muestra agendas validas, porcentajes DGAP, media, mediana y maxima.

El mapa respeta los filtros activos.

## Privacidad

El Excel se procesa siempre en local, dentro del navegador o WebView del dispositivo. No hay backend, no se suben archivos a servidores, no se usan APIs externas y no se almacena el Excel en la nube.
