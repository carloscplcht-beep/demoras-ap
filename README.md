# Dashboard local de demoras en Atencion Primaria

## Uso rapido

1. Abra `index.html` con doble clic.
2. Pulse `Seleccionar archivo Excel`.
3. Cargue el archivo `.xlsx` que quiera analizar.
4. Navegue por `Tabla`, `Resumen ejecutivo`, `Graficos` e `Informe`.

## Requisitos del archivo

- Formato: `.xlsx`
- Hoja leida: siempre la primera hoja disponible
- Cabeceras esperadas:
  - Área
  - Zona
  - Centro
  - CIAS
  - PROFESIONAL
  - Tipo visita
  - Accesibilidad
  - Categoría
  - Código de centro
  - Fecha Primer Hueco Libre
  - Fecha Primer Hueco ID
  - Fecha Corte

## Tratamiento de columnas opcionales y duplicadas

- Si existen `CIAS ID` o `UID`, la aplicación las ignora por completo.
- Si la primera fila contiene cabeceras duplicadas, la segunda repetición se renombra internamente con sufijo `.1` para evitar conflictos.
- Esto no modifica el Excel original.

## Notas

- No necesita servidor, Node, Python ni instalacion.
- Los indicadores ejecutivos solo usan filas con `Accesibilidad` numerica.
- La tabla filtrada puede exportarse a CSV.
- Cada grafico dispone de descarga directa a PNG.
- La vista `Informe` genera un resumen ejecutivo imprimible adaptado a los filtros activos.
