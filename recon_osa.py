"""
Cruce de la consulta de la OSA contra el catalogo de obras de Edimusica
(linea de comandos).

Lee "CONSULTA OSA_1Q 2026.xlsx" (formato que envia la OSA) y
"Obras_edimusica.xlsx" (nuestra base de obras), identifica que obras de la
consulta administramos, y genera una COPIA del archivo de la OSA con las
columnas EDITOR, % y COMENTARIOS llenas para los casos de alta confianza.
Los casos dudosos NO se llenan: se anotan en COMENTARIOS como sugerencia
para revision manual. El archivo original de la OSA nunca se modifica.

Para revisar los casos dudosos de forma visual e interactiva (estilo
"aceptar / rechazar"), usar la interfaz grafica: recon_osa_gui.py

Uso:
    python recon_osa.py [--limit N] [--osa RUTA] [--db RUTA]

    --limit N   Procesa solo las primeras N filas de datos (pruebas rapidas).
"""

import argparse
import time
from pathlib import Path

import openpyxl

import matching_core as mc

BASE_DIR = Path(__file__).resolve().parent
OSA_FILE_DEFAULT = BASE_DIR / "CONSULTA OSA_1Q 2026.xlsx"
DB_FILE_DEFAULT = BASE_DIR / "Obras_edimusica.xlsx"


def main():
    parser = argparse.ArgumentParser(description="Cruce CONSULTA OSA vs catalogo Edimusica")
    parser.add_argument("--osa", default=str(OSA_FILE_DEFAULT), help="Ruta al archivo de consulta de la OSA")
    parser.add_argument("--db", default=str(DB_FILE_DEFAULT), help="Ruta al archivo de obras de Edimusica")
    parser.add_argument("--limit", type=int, default=None, help="Procesar solo las primeras N filas (pruebas)")
    args = parser.parse_args()

    osa_path = Path(args.osa)
    db_path = Path(args.db)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    output_path = osa_path.with_name(f"{osa_path.stem}_EDIMUSICA_{timestamp}.xlsx")
    log_path = osa_path.with_name(f"log_matching_{timestamp}.xlsx")

    t0 = time.time()
    print(f"Cargando base de datos Edimusica: {db_path}")
    match_index = mc.MatchIndex(db_path)
    print(f"  {len(match_index.rows)} registros cargados.")
    print(f"  indice exacto: {len(match_index.exact_index)} titulos unicos")
    print(f"  indice de bloqueo: {len(match_index.blocking_index)} palabras")

    print(f"Abriendo consulta de la OSA: {osa_path}")
    wb = openpyxl.load_workbook(osa_path, data_only=False)
    ws = wb[mc.OSA_SHEET]
    max_row = ws.max_row
    last_data_row = max_row if args.limit is None else min(max_row, mc.OSA_FIRST_DATA_ROW + args.limit - 1)
    print(f"  {last_data_row - mc.OSA_FIRST_DATA_ROW + 1} filas de datos a procesar.")

    log_wb = openpyxl.Workbook()
    log_ws = log_wb.active
    log_ws.title = "log"
    log_ws.append([
        "Fila OSA", "Titulo OSA", "Autor OSA", "Decision", "Tipo match",
        "Score titulo", "Score autor", "Titulo BD", "Autor(es) BD",
        "% total BD", "Codigo BD",
    ])

    t1 = time.time()

    scan_results = mc.scan_osa_rows(ws, match_index, mc.OSA_FIRST_DATA_ROW, last_data_row)
    # Sin interfaz grafica no hay decisiones humanas: los REVISAR quedan
    # como PENDIENTE (con el comentario sugerido, sin llenar EDITOR/%).
    counters, log_rows = mc.apply_results(ws, scan_results, decisions={})
    for row in log_rows:
        log_ws.append(list(row))

    t2 = time.time()
    print(f"Matching completado en {t2 - t1:.1f}s")
    print(f"  AUTO (llenado): {counters['auto']}")
    print(f"  REVISAR / pendiente de revision (comentario sugerido): {counters['pendiente']}")
    print(f"  Sin coincidencia: {counters['sin_match']}")
    print(f"  Ya tenian EDITOR/% llenos (no tocadas): {counters['ya_lleno']}")
    print()
    print("Sugerencia: usa 'python recon_osa_gui.py' para revisar los casos")
    print("pendientes de forma visual e interactiva antes del envio final.")

    wb.save(output_path)
    log_wb.save(log_path)
    print(f"Archivo de resultado: {output_path}")
    print(f"Log de auditoria:     {log_path}")
    print(f"Tiempo total: {time.time() - t0:.1f}s")


if __name__ == "__main__":
    main()
