"""
Interfaz grafica para el cruce CONSULTA OSA vs catalogo Edimusica.

Flujo:
  1. Elegir el archivo de la OSA y el archivo de obras de Edimusica.
  2. Se analizan automaticamente (coincidencias seguras vs dudosas).
  3. Las coincidencias dudosas se revisan una por una en modo "juego"
     (botones grandes / flechas del teclado: si coincide o no coincide).
  4. Se genera el archivo final para enviar a la OSA, modificando solo las
     obras identificadas (automaticas + confirmadas a mano). El archivo
     original de la OSA nunca se modifica.

El progreso de la revision se guarda automaticamente para poder cerrar el
programa y continuar despues sin perder lo ya decidido.
"""

import json
import sys
import threading
import time
import webbrowser
from collections import deque
from pathlib import Path
from tkinter import Tk, StringVar, filedialog, messagebox
from tkinter import ttk

import openpyxl

import matching_core as mc


def app_dir():
    """Carpeta del script, o del .exe cuando esta empaquetado con PyInstaller."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


SESSIONS_DIR = app_dir() / "sesiones_revision"
SESSIONS_DIR.mkdir(exist_ok=True)

FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_SUBTITLE = ("Segoe UI", 12, "bold")
FONT_NORMAL = ("Segoe UI", 11)
FONT_BIG = ("Segoe UI", 22, "bold")
FONT_BTN = ("Segoe UI", 13, "bold")


def session_path_for(osa_path: Path) -> Path:
    return SESSIONS_DIR / f"{osa_path.stem}.json"


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Reconocimiento OSA - Edimusica")
        self.geometry("1020x760")
        self.minsize(900, 700)

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Accept.TButton", font=FONT_BTN, foreground="#0a6b1f")
        style.configure("Reject.TButton", font=FONT_BTN, foreground="#a30f0f")
        style.configure("Skip.TButton", font=FONT_BTN)
        style.configure("Big.TButton", font=FONT_BTN)

        self.osa_path = None
        self.db_path = None
        self.match_index = None
        self.wb = None
        self.ws = None
        self.scan_results = None
        self.review_groups = []
        self.decisions = {}
        self.session_file = None
        self.queue = deque()
        self.history = []

        self.container = ttk.Frame(self, padding=16)
        self.container.pack(fill="both", expand=True)

        self.show_inicio()

    # -----------------------------------------------------------------
    # utilidades comunes
    # -----------------------------------------------------------------

    def clear(self):
        for w in self.container.winfo_children():
            w.destroy()

    def save_session(self):
        if not self.session_file:
            return
        data = {
            "osa_path": str(self.osa_path),
            "db_path": str(self.db_path),
            "updated": time.strftime("%Y-%m-%d %H:%M:%S"),
            "decisions": self.decisions,
        }
        tmp = self.session_file.with_suffix(".json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        tmp.replace(self.session_file)

    # -----------------------------------------------------------------
    # Pantalla 1: inicio / seleccion de archivos
    # -----------------------------------------------------------------

    def show_inicio(self):
        self.clear()
        ttk.Label(self.container, text="Cruce CONSULTA OSA vs catalogo Edimusica",
                  font=FONT_TITLE).pack(anchor="w", pady=(0, 4))
        ttk.Label(
            self.container,
            text="Identifica que obras de la consulta de la OSA administra Edimusica,\n"
                 "y te deja revisar visualmente los casos dudosos antes de generar el archivo final.",
            font=FONT_NORMAL, foreground="#555",
        ).pack(anchor="w", pady=(0, 24))

        self.osa_var = StringVar(value=str(self.osa_path) if self.osa_path else "")
        self.db_var = StringVar(value=str(self.db_path) if self.db_path else "")

        self._file_row("Archivo CONSULTA OSA (.xlsx):", self.osa_var, self._pick_osa)
        self._file_row("Archivo Obras Edimusica (.xlsx):", self.db_var, self._pick_db)

        self.btn_analizar = ttk.Button(self.container, text="Analizar archivos",
                                        style="Big.TButton", command=self._start_analysis)
        self.btn_analizar.pack(anchor="w", pady=(24, 8))

        self.status_var = StringVar(value="")
        ttk.Label(self.container, textvariable=self.status_var, font=FONT_NORMAL,
                  foreground="#555").pack(anchor="w")

        self.progress = ttk.Progressbar(self.container, mode="indeterminate", length=400)
        self._update_analyze_btn_state()

    def _file_row(self, label, var, command):
        row = ttk.Frame(self.container)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, font=FONT_NORMAL, width=32, anchor="w").pack(side="left")
        entry = ttk.Entry(row, textvariable=var, state="readonly", width=60)
        entry.pack(side="left", padx=8)
        ttk.Button(row, text="Elegir...", command=command).pack(side="left")

    def _pick_osa(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo de la CONSULTA OSA",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=str(app_dir()),
        )
        if path:
            self.osa_path = Path(path)
            self.osa_var.set(str(self.osa_path))
            self._update_analyze_btn_state()

    def _pick_db(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo de Obras Edimusica",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=str(app_dir()),
        )
        if path:
            self.db_path = Path(path)
            self.db_var.set(str(self.db_path))
            self._update_analyze_btn_state()

    def _update_analyze_btn_state(self):
        ok = bool(self.osa_path and self.db_path)
        self.btn_analizar.configure(state="normal" if ok else "disabled")

    def _start_analysis(self):
        self.btn_analizar.configure(state="disabled")
        self.status_var.set("Analizando... esto puede tardar unos segundos.")
        self.progress.pack(anchor="w", pady=8)
        self.progress.start(12)
        threading.Thread(target=self._run_analysis, daemon=True).start()

    def _run_analysis(self):
        try:
            match_index = mc.MatchIndex(self.db_path)
            wb = openpyxl.load_workbook(self.osa_path, data_only=False)
            ws = wb[mc.OSA_SHEET]
            scan_results = mc.scan_osa_rows(ws, match_index)
            review_groups = mc.build_review_groups(scan_results)
        except Exception as exc:  # noqa: BLE001
            self.after(0, self._analysis_failed, str(exc))
            return
        self.after(0, self._analysis_done, match_index, wb, ws, scan_results, review_groups)

    def _analysis_failed(self, message):
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_analizar.configure(state="normal")
        self.status_var.set("")
        messagebox.showerror("Error al analizar", message)

    def _analysis_done(self, match_index, wb, ws, scan_results, review_groups):
        self.progress.stop()
        self.progress.pack_forget()
        self.match_index = match_index
        self.wb = wb
        self.ws = ws
        self.scan_results = scan_results
        self.review_groups = review_groups
        self.decisions = {}

        self.session_file = session_path_for(self.osa_path)
        if self.session_file.exists():
            try:
                data = json.loads(self.session_file.read_text(encoding="utf-8"))
                n = len(data.get("decisions", {}))
                if n and messagebox.askyesno(
                    "Sesion encontrada",
                    f"Ya hay {n} decisiones guardadas para este archivo "
                    f"(ultima actualizacion: {data.get('updated', '?')}).\n\n"
                    "Continuar donde quedaste?",
                ):
                    self.decisions = data.get("decisions", {})
            except Exception:
                pass

        self.show_resumen()

    # -----------------------------------------------------------------
    # Pantalla 2: resumen del analisis
    # -----------------------------------------------------------------

    def _counts(self):
        auto = sum(1 for r in self.scan_results
                   if not r["ya_lleno"] and r["match"] and r["match"]["tier"] == "AUTO")
        sin_match = sum(1 for r in self.scan_results
                         if not r["ya_lleno"] and r["match"] is None)
        ya_lleno = sum(1 for r in self.scan_results if r["ya_lleno"])
        total_grupos = len(self.review_groups)
        decididos = sum(1 for g in self.review_groups if g["key"] in self.decisions)
        aceptados = sum(1 for v in self.decisions.values() if v == "SI")
        rechazados = sum(1 for v in self.decisions.values() if v == "NO")
        pendientes = total_grupos - decididos
        filas_revisar = sum(len(g["rows"]) for g in self.review_groups)
        return dict(auto=auto, sin_match=sin_match, ya_lleno=ya_lleno,
                    total_grupos=total_grupos, decididos=decididos,
                    aceptados=aceptados, rechazados=rechazados,
                    pendientes=pendientes, filas_revisar=filas_revisar)

    def show_resumen(self):
        self.clear()
        c = self._counts()

        ttk.Label(self.container, text="Resultado del analisis", font=FONT_TITLE).pack(anchor="w", pady=(0, 16))

        info = ttk.Frame(self.container)
        info.pack(fill="x", pady=4)

        def stat(text, value, color="#222"):
            row = ttk.Frame(info)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=text, font=FONT_NORMAL, width=48, anchor="w").pack(side="left")
            ttk.Label(row, text=str(value), font=FONT_SUBTITLE, foreground=color).pack(side="left")

        stat("Identificadas automaticamente con alta confianza:", c["auto"], "#0a6b1f")
        stat("Coincidencias dudosas por revisar (agrupadas):", c["pendientes"], "#b8860b")
        stat("  -> filas de la OSA que cubren esos grupos:", c["filas_revisar"])
        stat("  -> ya revisadas antes (aceptadas / rechazadas):", f"{c['aceptados']} / {c['rechazados']}")
        stat("Sin ninguna coincidencia en nuestro catalogo:", c["sin_match"])
        stat("Filas que ya tenian EDITOR/% llenos (no se tocan):", c["ya_lleno"])

        btns = ttk.Frame(self.container)
        btns.pack(anchor="w", pady=28)

        if c["pendientes"] > 0:
            ttk.Button(
                btns, text=f"Revisar coincidencias dudosas ({c['pendientes']} pendientes)",
                style="Big.TButton", command=self.show_juego,
            ).pack(anchor="w", pady=4)
        else:
            ttk.Label(self.container, text="No quedan coincidencias dudosas por revisar.",
                      font=FONT_NORMAL, foreground="#0a6b1f").pack(anchor="w")

        ttk.Button(
            btns, text="Generar archivo final para la OSA ahora",
            command=self.show_exportar,
        ).pack(anchor="w", pady=4)

        ttk.Button(btns, text="<- Elegir otros archivos", command=self.show_inicio).pack(anchor="w", pady=(16, 0))

    # -----------------------------------------------------------------
    # Pantalla 3: juego de revision
    # -----------------------------------------------------------------

    def show_juego(self):
        self.clear()
        self.queue = deque(i for i, g in enumerate(self.review_groups) if g["key"] not in self.decisions)
        self.history = []

        top = ttk.Frame(self.container)
        top.pack(fill="x")
        ttk.Label(top, text="Es la misma obra?", font=FONT_TITLE).pack(side="left")
        ttk.Button(top, text="Terminar por ahora -> exportar", command=self.show_exportar).pack(side="right")

        self.progress_label = ttk.Label(self.container, font=FONT_NORMAL, foreground="#555")
        self.progress_label.pack(anchor="w", pady=(4, 2))
        self.progress_bar = ttk.Progressbar(self.container, mode="determinate")
        self.progress_bar.pack(fill="x", pady=(0, 16))

        # Nota: se empaqueta con fill="x" (sin expand) a proposito: si en vez
        # se usara expand=True aqui, esta tarjeta absorberia todo el espacio
        # extra al maximizar la ventana y empujaria los botones fuera del
        # area visible. Con fill="x" el layout queda estable a cualquier
        # tamano de ventana; el espacio sobrante simplemente queda en blanco
        # debajo de los controles.
        cards = ttk.Frame(self.container)
        cards.pack(fill="x", pady=(0, 16))
        cards.columnconfigure(0, weight=1)
        cards.columnconfigure(1, weight=1)

        osa_box = ttk.LabelFrame(cards, text="Reporta la OSA", padding=16)
        osa_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        bd_box = ttk.LabelFrame(cards, text="Posible coincidencia en catalogo Edimusica", padding=16)
        bd_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self.lbl_osa_titulo = ttk.Label(osa_box, font=FONT_BIG, wraplength=380, justify="left")
        self.lbl_osa_titulo.pack(anchor="w", pady=(0, 10))
        self.lbl_osa_autor = ttk.Label(osa_box, font=FONT_NORMAL, wraplength=380, justify="left")
        self.lbl_osa_autor.pack(anchor="w", pady=2)
        self.lbl_osa_interprete = ttk.Label(osa_box, font=FONT_NORMAL, wraplength=380, justify="left", foreground="#555")
        self.lbl_osa_interprete.pack(anchor="w", pady=2)
        self.lbl_osa_filas = ttk.Label(osa_box, font=("Segoe UI", 10, "italic"), foreground="#777")
        self.lbl_osa_filas.pack(anchor="w", pady=(10, 0))

        self.lbl_bd_titulo = ttk.Label(bd_box, font=FONT_BIG, wraplength=380, justify="left")
        self.lbl_bd_titulo.pack(anchor="w", pady=(0, 10))
        self.lbl_bd_autor = ttk.Label(bd_box, font=FONT_NORMAL, wraplength=380, justify="left")
        self.lbl_bd_autor.pack(anchor="w", pady=2)
        self.lbl_bd_pct = ttk.Label(bd_box, font=FONT_SUBTITLE, foreground="#0a6b1f")
        self.lbl_bd_pct.pack(anchor="w", pady=(8, 2))
        self.lbl_bd_score = ttk.Label(bd_box, font=("Segoe UI", 10), foreground="#777")
        self.lbl_bd_score.pack(anchor="w", pady=(10, 0))

        btns = ttk.Frame(self.container)
        btns.pack(fill="x", pady=24)
        ttk.Button(btns, text="❌  NO coincide  (←)", style="Reject.TButton",
                   command=self.decide_no).pack(side="left", expand=True, fill="x", padx=6, ipady=14)
        ttk.Button(btns, text="⏭  Revisar despues  (espacio)", style="Skip.TButton",
                   command=self.decide_skip).pack(side="left", expand=True, fill="x", padx=6, ipady=14)
        ttk.Button(btns, text="✅  SI, es la misma obra  (→)", style="Accept.TButton",
                   command=self.decide_yes).pack(side="left", expand=True, fill="x", padx=6, ipady=14)

        bottom = ttk.Frame(self.container)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="Deshacer ultima decision (Ctrl+Z)", command=self.undo).pack(side="left")

        self.bind("<Left>", lambda e: self.decide_no())
        self.bind("<Right>", lambda e: self.decide_yes())
        self.bind("<Return>", lambda e: self.decide_yes())
        self.bind("<space>", lambda e: self.decide_skip())
        self.bind("<Control-z>", lambda e: self.undo())

        self._refresh_card()

    def _refresh_card(self):
        total = len(self.review_groups)
        decididos = total - len(self.queue)
        self.progress_label.configure(
            text=f"{decididos} de {total} decisiones tomadas "
                 f"({sum(1 for v in self.decisions.values() if v=='SI')} aceptadas, "
                 f"{sum(1 for v in self.decisions.values() if v=='NO')} rechazadas)"
        )
        self.progress_bar.configure(maximum=max(total, 1), value=decididos)

        if not self.queue:
            self.show_resumen()
            return

        idx = self.queue[0]
        g = self.review_groups[idx]
        sample = g["sample"]
        match = g["match"]
        summary = match["summary"]

        self.lbl_osa_titulo.configure(text=sample["titulo"])
        self.lbl_osa_autor.configure(text=f"Autor: {sample['autor'] or '(sin dato)'}")
        self.lbl_osa_interprete.configure(text=f"Interprete: {sample['interprete'] or '(sin dato)'}")
        n_rows = len(g["rows"])
        self.lbl_osa_filas.configure(
            text=f"Aplica a {n_rows} fila(s) de la consulta OSA" if n_rows > 1
            else "Aplica a 1 fila de la consulta OSA"
        )

        self.lbl_bd_titulo.configure(text=summary["titulo"])
        self.lbl_bd_autor.configure(text="Autor(es): " + mc.format_authors(summary["authors"]))
        self.lbl_bd_pct.configure(text=f"% que administraria Edimusica: {summary['total_pct']:.2f}%")
        self.lbl_bd_score.configure(
            text=f"Tipo de match: {match['match_type']}  |  score titulo: {match['title_score']}  |  "
                 f"score autor: {match['author_score']}  |  codigo: {summary['codant']}"
        )

    def _decide(self, value):
        if not self.queue:
            return
        idx = self.queue.popleft()
        key = self.review_groups[idx]["key"]
        self.decisions[key] = value
        self.history.append(idx)
        self.save_session()
        self._refresh_card()

    def decide_yes(self):
        self._decide("SI")

    def decide_no(self):
        self._decide("NO")

    def decide_skip(self):
        if not self.queue:
            return
        idx = self.queue.popleft()
        if self.queue:
            self.queue.append(idx)
            self._refresh_card()
        else:
            # era el unico pendiente: no tiene caso re-mostrarlo en bucle
            self.queue.append(idx)
            messagebox.showinfo(
                "Solo queda este caso",
                "Es el unico caso pendiente. Decide Si/No o ve a exportar dejandolo pendiente.",
            )

    def undo(self):
        if not self.history:
            return
        idx = self.history.pop()
        key = self.review_groups[idx]["key"]
        self.decisions.pop(key, None)
        self.queue.appendleft(idx)
        self.save_session()
        self._refresh_card()

    # -----------------------------------------------------------------
    # Pantalla 4: exportar
    # -----------------------------------------------------------------

    def show_exportar(self):
        self.clear()
        for seq in ("<Left>", "<Right>", "<Return>", "<space>", "<Control-z>"):
            self.unbind(seq)

        c = self._counts()
        ttk.Label(self.container, text="Generar archivo final", font=FONT_TITLE).pack(anchor="w", pady=(0, 16))
        ttk.Label(
            self.container,
            text="Se creara una COPIA del archivo de la OSA (el original no se modifica),\n"
                 "llenando EDITOR y % solo en las obras identificadas automaticamente o\n"
                 "confirmadas por ti. Las que quedaron pendientes o rechazadas no se tocan.",
            font=FONT_NORMAL, foreground="#555", justify="left",
        ).pack(anchor="w", pady=(0, 20))

        ttk.Label(self.container, text=f"Se llenaran (automaticas + confirmadas): "
                                        f"{c['auto'] + c['aceptados']} coincidencias",
                  font=FONT_SUBTITLE, foreground="#0a6b1f").pack(anchor="w", pady=2)
        quedan_pendientes = max(c["total_grupos"] - c["aceptados"] - c["rechazados"], 0)
        ttk.Label(self.container, text=f"Quedaran pendientes de revisar (sin llenar): {quedan_pendientes}",
                  font=FONT_NORMAL, foreground="#b8860b").pack(anchor="w", pady=2)

        self.export_status = StringVar(value="")
        ttk.Label(self.container, textvariable=self.export_status, font=FONT_NORMAL).pack(anchor="w", pady=(20, 4))
        self.export_progress = ttk.Progressbar(self.container, mode="indeterminate", length=400)

        btns = ttk.Frame(self.container)
        btns.pack(anchor="w", pady=20)
        self.btn_exportar = ttk.Button(btns, text="Generar archivo para la OSA",
                                        style="Big.TButton", command=self._start_export)
        self.btn_exportar.pack(side="left")
        ttk.Button(btns, text="<- Volver a revisar", command=self.show_resumen).pack(side="left", padx=12)

    def _start_export(self):
        self.btn_exportar.configure(state="disabled")
        self.export_status.set("Generando archivo...")
        self.export_progress.pack(anchor="w", pady=6)
        self.export_progress.start(12)
        threading.Thread(target=self._run_export, daemon=True).start()

    def _run_export(self):
        try:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_path = self.osa_path.with_name(f"{self.osa_path.stem}_EDIMUSICA_{timestamp}.xlsx")
            log_path = self.osa_path.with_name(f"log_matching_{timestamp}.xlsx")

            wb2 = openpyxl.load_workbook(self.osa_path, data_only=False)
            ws2 = wb2[mc.OSA_SHEET]
            counters, log_rows = mc.apply_results(ws2, self.scan_results, self.decisions)

            log_wb = openpyxl.Workbook()
            log_ws = log_wb.active
            log_ws.title = "log"
            log_ws.append([
                "Fila OSA", "Titulo OSA", "Autor OSA", "Decision", "Tipo match",
                "Score titulo", "Score autor", "Titulo BD", "Autor(es) BD",
                "% total BD", "Codigo BD",
            ])
            for row in log_rows:
                log_ws.append(list(row))

            wb2.save(output_path)
            log_wb.save(log_path)
        except Exception as exc:  # noqa: BLE001
            self.after(0, self._export_failed, str(exc))
            return
        self.after(0, self._export_done, output_path, log_path, counters)

    def _export_failed(self, message):
        self.export_progress.stop()
        self.export_progress.pack_forget()
        self.btn_exportar.configure(state="normal")
        self.export_status.set("")
        messagebox.showerror("Error al generar el archivo", message)

    def _export_done(self, output_path, log_path, counters):
        self.export_progress.stop()
        self.export_progress.pack_forget()
        self.export_status.set("Listo.")
        self.clear()

        ttk.Label(self.container, text="Archivo generado", font=FONT_TITLE,
                  foreground="#0a6b1f").pack(anchor="w", pady=(0, 16))
        ttk.Label(self.container, text=f"Archivo para enviar a la OSA:\n{output_path}",
                  font=FONT_NORMAL, justify="left").pack(anchor="w", pady=6)
        ttk.Label(self.container, text=f"Log de auditoria:\n{log_path}",
                  font=FONT_NORMAL, justify="left").pack(anchor="w", pady=6)

        resumen = (
            f"Automaticas: {counters['auto']}   |   Confirmadas a mano: {counters['confirmado']}   |   "
            f"Rechazadas: {counters['rechazado']}   |   Pendientes: {counters['pendiente']}   |   "
            f"Sin coincidencia: {counters['sin_match']}"
        )
        ttk.Label(self.container, text=resumen, font=("Segoe UI", 10), foreground="#555").pack(anchor="w", pady=(16, 0))

        btns = ttk.Frame(self.container)
        btns.pack(anchor="w", pady=24)
        ttk.Button(btns, text="Abrir carpeta", command=lambda: webbrowser.open(str(output_path.parent))).pack(side="left")
        ttk.Button(btns, text="Volver al resumen", command=self.show_resumen).pack(side="left", padx=12)
        ttk.Button(btns, text="Empezar con otro archivo", command=self.show_inicio).pack(side="left", padx=12)


if __name__ == "__main__":
    app = App()
    app.mainloop()
