# reconocimiento

Script para cruzar la "CONSULTA OSA" (relacion de obras usadas en
establecimientos, formato Excel que envia la OSA cada trimestre) contra el
catalogo propio de obras de Edimusica, e identificar automaticamente que
obras administramos, con que porcentaje, y sugerir autores para las obras
marcadas como "No identificado".

## Por que existe

Cada trimestre la OSA envia un Excel de ~30-40 mil filas para identificar
manualmente titulo por titulo. Este script automatiza el cruce usando
coincidencia exacta y difusa (fuzzy matching) de titulo + autor contra nuestra
base de obras, dejando el trabajo manual solo para los casos realmente
ambiguos.

## Instalacion

```
pip install -r requirements.txt
```

## Uso

Coloca en esta carpeta:
- El archivo que envia la OSA (por defecto se busca `CONSULTA OSA_1Q 2026.xlsx`)
- Nuestra base de obras (`Obras_edimusica.xlsx`), con columnas
  `TITULO, NOMAUTOR, CODIGO, COD ANT, PORCENTAJE`

```
python recon_osa.py
```

Genera dos archivos con timestamp (el archivo original de la OSA nunca se
modifica):

- `CONSULTA OSA_..._EDIMUSICA_<fecha>.xlsx` — copia del archivo de la OSA
  con las columnas `EDITOR`, `%` y `COMENTARIOS` llenas para las obras
  identificadas con alta confianza.
- `log_matching_<fecha>.xlsx` — auditoria fila por fila: score de titulo,
  score de autor, obra/autor de la base con la que hizo match, y si quedo
  como `AUTO` o `REVISAR`. Util para ordenar por score y priorizar la
  revision manual.

Para probar rapido con un subconjunto de filas:

```
python recon_osa.py --limit 500
```

Rutas de archivos personalizadas:

```
python recon_osa.py --osa "ruta/consulta.xlsx" --db "ruta/obras.xlsx"
```

## Logica de decision

1. **Coincidencia de titulo**: exacta (normalizada: mayusculas, sin tildes,
   sin puntuacion) primero; si no hay exacta, aproximada (rapidfuzz,
   bloqueada por palabras del titulo para que sea rapida sobre ~20k obras).
2. **Confirmacion de autor**: un titulo igual o parecido NO es suficiente
   por si solo — hay titulos genericos o tradicionales que distintos
   compositores repiten. Por eso **siempre** se exige que el autor de la
   OSA coincida (score >= 85, y con margen claro sobre el segundo candidato)
   antes de llenar `EDITOR`/`%` automaticamente. Si el autor de la OSA viene
   "No identificado" o no coincide con confianza, la fila pasa a revision
   manual con el o los autores sugeridos en `COMENTARIOS`, en vez de
   asumirlos.
3. **Obras con varios coautores**: en la base de Edimusica una misma obra
   puede tener varias filas (una por coautor), agrupadas por `COD ANT`. El
   `%` que se llena es la suma de los porcentajes de todos los coautores
   que administramos para esa obra. Si la misma obra aparece varias veces
   en el archivo de la OSA (distintos interpretes), se llena el mismo %
   total en cada aparicion — no se modifica ni se reparte el archivo de la
   OSA.
4. Las filas que ya tenian `EDITOR` o `%` llenos no se tocan.

## Ajustar el umbral de revision

Los umbrales estan al inicio de `recon_osa.py` (`AUTHOR_CONFIRM_MIN`,
`AUTHOR_MARGIN_MIN`, `TIER1_FUZZY_TITLE_MIN`, `FUZZY_CANDIDATE_CUTOFF`). Si
despues de revisar varias corridas ven que ciertos patrones de "REVISAR"
siempre resultan correctos, se pueden ajustar para automatizarlos.
