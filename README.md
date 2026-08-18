# reconocimiento

Herramienta para cruzar la "CONSULTA OSA" (relacion de obras usadas en
establecimientos, formato Excel que envia la OSA cada trimestre) contra el
catalogo propio de obras de Edimusica: identifica automaticamente que obras
administramos, con que porcentaje, deja revisar visualmente los casos
dudosos, y genera el archivo final para devolver a la OSA.

## Por que existe

Cada trimestre la OSA envia un Excel de ~30-40 mil filas para identificar
manualmente titulo por titulo. Esta herramienta automatiza el cruce usando
coincidencia exacta y difusa (fuzzy matching) de titulo + autor contra
nuestra base de obras, y deja el trabajo manual solo para los casos
realmente ambiguos — revisados uno por uno en una interfaz visual tipo
"aceptar / rechazar", en vez de en Excel.

## Instalacion

```
pip install -r requirements.txt
```

## Uso — interfaz grafica (recomendado)

```
python recon_osa_gui.py
```

O, si no quieres usar la terminal, doble clic en **`Abrir Reconocimiento
OSA.bat`**.

Flujo dentro de la app:

1. Elegir el archivo de la OSA y el archivo de obras de Edimusica.
2. Se analizan automaticamente. Se muestra cuantas obras quedaron
   identificadas con alta confianza y cuantas quedaron dudosas.
3. Las dudosas se revisan una por una en la pantalla de "juego": se ve la
   obra tal como la reporto la OSA junto a la posible coincidencia en
   nuestra base (titulo, autor(es), % y el score de similitud), y se
   decide con los botones o con las flechas del teclado (← no coincide,
   → si coincide, espacio para revisar despues, Ctrl+Z para deshacer). El
   progreso se guarda automaticamente en `sesiones_revision/`, asi que se
   puede cerrar el programa y continuar despues sin perder lo ya decidido.
4. Al generar el archivo final, se modifica **solo** lo identificado
   (automatico + confirmado a mano): se llenan `EDITOR` y `%`. Lo
   rechazado y lo que quedo pendiente no se toca. El archivo original de
   la OSA nunca se modifica — siempre se crea una copia nueva.

## Uso — linea de comandos (sin revision visual)

```
python recon_osa.py
```

Llena solo las coincidencias de alta confianza; las dudosas quedan
anotadas como sugerencia en `COMENTARIOS` (sin llenar `EDITOR`/`%`) para
revisión manual directamente en Excel. Utilidad rapida cuando no se
necesita el modo de revision interactivo.

Para probar rapido con un subconjunto de filas: `python recon_osa.py --limit 500`

Rutas de archivos personalizadas: `python recon_osa.py --osa "ruta/consulta.xlsx" --db "ruta/obras.xlsx"`

## Archivos que genera

- `..._EDIMUSICA_<fecha>.xlsx` — copia del archivo de la OSA lista para
  devolver, con `EDITOR`, `%` y `COMENTARIOS` llenos donde corresponde.
- `log_matching_<fecha>.xlsx` — auditoria fila por fila: score de titulo,
  score de autor, obra/autor de la base con la que hizo match, y la
  decision final (`AUTO`, `CONFIRMADO_MANUAL`, `RECHAZADO_MANUAL`,
  `PENDIENTE`).
- `sesiones_revision/<nombre>.json` — progreso de la revision visual para
  cada archivo de la OSA (no se versiona en git).

## Logica de decision

Principio rector: **ante la duda, no sugerir**. Es preferible dejar una
obra sin identificar que sugerir -o peor, auto-llenar- una coincidencia
que solo comparte una palabra de titulo o un apellido comun. Con esta
logica, sobre las ~37.000 filas de la consulta 1Q 2026: 566 quedan
identificadas automaticamente y solo 583 pasan a revision manual (antes
de endurecer los filtros, la revision manual llegaba a mas de 8.000
filas, la mayoria coincidencias sin ninguna relacion real).

1. **Coincidencia de titulo**: exacta (normalizada: mayusculas, sin tildes,
   sin puntuacion) primero; si no hay exacta, aproximada (rapidfuzz,
   bloqueada por palabras del titulo para que sea rapida sobre ~20k obras).
   Un titulo aproximado **debe ademas compartir la mayoria de sus palabras**
   con el candidato (no solo un score de caracteres alto): dos titulos
   cortos como "DOS FLORES" y "LOS DOS" pueden parecer similares letra por
   letra sin tener ninguna relacion real, asi que un score de caracteres
   por si solo no basta.
2. **Confirmacion de autor**: un titulo igual o parecido NO es suficiente
   por si solo — hay titulos genericos o tradicionales que distintos
   compositores repiten (o coinciden por azar). Por eso:
   - Para **auto-llenar** `EDITOR`/`%` siempre se exige que el autor de la
     OSA coincida con confianza (score >= 85, y con margen claro sobre el
     segundo candidato).
   - Para **siquiera sugerir** una revision manual, tambien hay un piso
     minimo de coincidencia de autor — un titulo exacto con un autor sin
     ninguna relacion suele ser una obra homonima distinta, no la misma
     obra, y no vale la pena mostrarla.
   - Si el autor de la OSA viene "No identificado", no hay señal de autor
     para comparar: solo se sugiere revision cuando el titulo es exacto,
     inequivoco en nuestra base (una sola obra con ese titulo) y no es
     demasiado generico/corto.
3. **Obras con varios coautores**: en la base de Edimusica una misma obra
   puede tener varias filas (una por coautor), agrupadas por `COD ANT`. El
   `%` que se llena es la suma de los porcentajes de todos los coautores
   que administramos para esa obra. Si la misma obra aparece varias veces
   en el archivo de la OSA (distintos interpretes), se llena el mismo %
   total en cada aparicion — no se modifica ni se reparte el archivo de la
   OSA.
4. Las filas que ya tenian `EDITOR` o `%` llenos no se tocan.
5. En la GUI, las filas dudosas que comparten el mismo titulo+autor se
   revisan una sola vez y la decision se aplica a todas (p. ej. la misma
   obra reportada para varios interpretes).

## Ajustar el umbral de revision

Los umbrales estan al inicio de `matching_core.py`: `AUTHOR_CONFIRM_MIN` y
`AUTHOR_MARGIN_MIN` (para auto-llenar), `REVISAR_AUTHOR_FLOOR` (piso minimo
de coincidencia de autor para siquiera sugerir), `FUZZY_CANDIDATE_CUTOFF` y
`TITLE_WORD_OVERLAP_MIN` (que tan parecido/solapado debe ser un titulo no
exacto para considerarlo candidato), `TIER1_FUZZY_TITLE_MIN` (para
auto-llenar con titulo no exacto). Si despues de revisar varias corridas
ven que ciertos patrones de "REVISAR" siempre resultan correctos, se
pueden relajar; si aparecen falsos positivos, endurecerlos mas.

## Sobre el .exe

Se intento empaquetar `recon_osa_gui.py` como ejecutable standalone con
PyInstaller (`pip install pyinstaller`, luego
`pyinstaller --onefile --windowed --name ReconocimientoOSA recon_osa_gui.py`).
El `.exe` se genera sin problema, pero en este equipo una directiva de
Control de Aplicaciones de Windows bloquea la ejecucion de binarios nuevos
sin firmar — no es un error del programa. Mientras eso no se resuelva con
el area de TI (firmarlo o agregar una excepcion), usar la app via
`python recon_osa_gui.py` o el `.bat` incluido, que si esta permitido
porque usa el Python ya instalado y confiado por el sistema.
