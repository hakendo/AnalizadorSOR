"""
Analizador SOR — Fibra Óptica
Extrae métricas de eventos/empalmes desde archivos OTDR (.sor) y exporta a Excel.
"""
from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from sor_parser import scan_folder, parse_sor
from excel_exporter import export_to_excel, build_output_path, OPENPYXL_OK


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Analizador SOR — Fibra Óptica")
        self.geometry("820x560")
        self.resizable(True, True)
        self.configure(bg="#F0F4F8")

        self._root_folder: str = ""
        self._cables: dict = {}
        self._parsed: dict = {}

        self._build_ui()
        self._check_deps()

    # ── Dependency check ─────────────────────────────────────────────────
    def _check_deps(self) -> None:
        if not OPENPYXL_OK:
            self._log("⚠  openpyxl no instalado. Ejecuta:  pip install openpyxl", warn=True)

    # ── UI construction ──────────────────────────────────────────────────
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        # Title bar
        title = tk.Frame(self, bg="#1F4E79", pady=8)
        title.pack(fill="x")
        tk.Label(title, text="Analizador SOR — Fibra Óptica",
                 font=("Segoe UI", 14, "bold"),
                 fg="white", bg="#1F4E79").pack()

        # Folder selection
        frm_folder = tk.LabelFrame(self, text=" Carpeta de archivos SOR ",
                                   bg="#F0F4F8", fg="#1F4E79",
                                   font=("Segoe UI", 9, "bold"))
        frm_folder.pack(fill="x", **pad)

        inner = tk.Frame(frm_folder, bg="#F0F4F8")
        inner.pack(fill="x", padx=6, pady=4)

        self._folder_var = tk.StringVar(value="(ninguna carpeta seleccionada)")
        tk.Label(inner, textvariable=self._folder_var,
                 bg="#F0F4F8", fg="#444444",
                 font=("Segoe UI", 9), anchor="w").pack(side="left", fill="x", expand=True)
        tk.Button(inner, text="Seleccionar Carpeta",
                  command=self._select_folder,
                  bg="#2E75B6", fg="white",
                  font=("Segoe UI", 9, "bold"),
                  relief="flat", padx=10, pady=4).pack(side="right")

        # Cables detected
        frm_cables = tk.LabelFrame(self, text=" Cables detectados ",
                                   bg="#F0F4F8", fg="#1F4E79",
                                   font=("Segoe UI", 9, "bold"))
        frm_cables.pack(fill="x", **pad)

        self._cable_frame = tk.Frame(frm_cables, bg="#F0F4F8")
        self._cable_frame.pack(fill="x", padx=6, pady=4)
        tk.Label(self._cable_frame, text="—",
                 bg="#F0F4F8", fg="#888888").pack(anchor="w")

        # Progress
        frm_prog = tk.LabelFrame(self, text=" Progreso ",
                                 bg="#F0F4F8", fg="#1F4E79",
                                 font=("Segoe UI", 9, "bold"))
        frm_prog.pack(fill="x", **pad)

        self._progress = ttk.Progressbar(frm_prog, mode="determinate", length=600)
        self._progress.pack(fill="x", padx=8, pady=(4, 2))

        self._status_var = tk.StringVar(value="Listo.")
        tk.Label(frm_prog, textvariable=self._status_var,
                 bg="#F0F4F8", fg="#444444",
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", padx=8, pady=(0, 4))

        # Log area
        frm_log = tk.LabelFrame(self, text=" Registro ",
                                bg="#F0F4F8", fg="#1F4E79",
                                font=("Segoe UI", 9, "bold"))
        frm_log.pack(fill="both", expand=True, **pad)

        self._log_text = tk.Text(frm_log, height=8, state="disabled",
                                 font=("Consolas", 8), bg="#FAFAFA",
                                 relief="flat", wrap="word")
        scrollbar = ttk.Scrollbar(frm_log, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._log_text.pack(fill="both", expand=True, padx=4, pady=4)

        # Action buttons
        frm_btns = tk.Frame(self, bg="#F0F4F8")
        frm_btns.pack(fill="x", padx=10, pady=(0, 10))

        self._btn_export = tk.Button(
            frm_btns, text="Exportar Excel",
            command=self._start_export,
            bg="#1F4E79", fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", padx=20, pady=6,
            state="disabled"
        )
        self._btn_export.pack(side="right")

        tk.Button(frm_btns, text="Analizar archivos",
                  command=self._start_analyze,
                  bg="#2E75B6", fg="white",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=20, pady=6).pack(side="right", padx=(0, 8))

    # ── Folder selection ─────────────────────────────────────────────────
    def _select_folder(self) -> None:
        folder = filedialog.askdirectory(title="Seleccionar carpeta con subcarpetas de cable SOR")
        if not folder:
            return
        self._root_folder = folder
        self._folder_var.set(folder)
        self._parsed = {}
        self._btn_export.config(state="disabled")

        # Detect cables
        cables = scan_folder(folder)
        self._cables = cables
        self._update_cable_list(cables)

        if cables:
            total = sum(len(v) for v in cables.values())
            self._log(f"Detectados {len(cables)} cable(s), {total} fibra(s) en total.")
        else:
            self._log("No se encontraron archivos SOR en las subcarpetas.", warn=True)

    def _update_cable_list(self, cables: dict) -> None:
        for widget in self._cable_frame.winfo_children():
            widget.destroy()
        if not cables:
            tk.Label(self._cable_frame, text="—", bg="#F0F4F8", fg="#888888").pack(anchor="w")
            return
        for cable_name, fibras in cables.items():
            txt = f"✓  {cable_name.strip():<40}  ({len(fibras)} fibra(s))"
            tk.Label(self._cable_frame, text=txt, bg="#F0F4F8", fg="#1F4E79",
                     font=("Segoe UI", 9)).pack(anchor="w")

    # ── Analysis (parse SOR files) ────────────────────────────────────────
    def _start_analyze(self) -> None:
        if not self._root_folder:
            messagebox.showwarning("Sin carpeta", "Selecciona primero una carpeta.")
            return
        if not self._cables:
            messagebox.showwarning("Sin SOR", "No se encontraron archivos SOR.")
            return
        threading.Thread(target=self._run_analyze, daemon=True).start()

    def _run_analyze(self) -> None:
        self._btn_export.config(state="disabled")
        self._parsed = {}

        all_paths: list[tuple[str, int, str]] = []
        for cable_name, fibras in self._cables.items():
            for item in fibras:
                all_paths.append((cable_name, item['fibra'], item['path']))

        total = len(all_paths)
        self._progress["maximum"] = total
        self._progress["value"] = 0

        errors = 0
        cables_parsed: dict[str, list[dict]] = {}

        for i, (cable_name, fibra_num, path) in enumerate(all_paths):
            self._status_var.set(f"Analizando fibra {fibra_num} — {cable_name.strip()} ...")
            result = parse_sor(path)
            if result is None or 'error' in result:
                err = result.get('error', 'error desconocido') if result else 'None'
                self._log(f"⚠  Fibra {fibra_num} ({cable_name.strip()}): {err}", warn=True)
                errors += 1
            else:
                result['fibra_num'] = fibra_num
                cables_parsed.setdefault(cable_name, []).append(result)

            self._progress["value"] = i + 1
            self.update_idletasks()

        self._parsed = cables_parsed
        ok = total - errors
        self._status_var.set(f"Análisis completo: {ok}/{total} archivos OK, {errors} errores.")
        self._log(f"Análisis finalizado: {ok} fibra(s) procesada(s), {errors} error(es).")

        if cables_parsed:
            self._btn_export.config(state="normal")

    # ── Export ────────────────────────────────────────────────────────────
    def _start_export(self) -> None:
        if not self._parsed:
            messagebox.showwarning("Sin datos", "Primero analiza los archivos SOR.")
            return
        threading.Thread(target=self._run_export, daemon=True).start()

    def _run_export(self) -> None:
        self._btn_export.config(state="disabled")
        self._status_var.set("Generando Excel...")

        # Determine reference date from first parsed result
        ref_date = None
        for fibers in self._parsed.values():
            if fibers and fibers[0].get('date'):
                ref_date = fibers[0]['date']
                break

        output_path = build_output_path(self._root_folder, ref_date)

        try:
            export_to_excel(self._parsed, output_path)
            self._status_var.set(f"Excel guardado: {os.path.basename(output_path)}")
            self._log(f"✓ Exportado: {output_path}")

            # Open file on Windows
            if sys.platform.startswith("win"):
                os.startfile(output_path)
            else:
                self._log("Abre el archivo manualmente (no Windows).")

        except ImportError as e:
            messagebox.showerror("Dependencia faltante", str(e))
            self._log(f"ERROR: {e}", warn=True)
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))
            self._log(f"ERROR exportando: {e}", warn=True)
        finally:
            self._btn_export.config(state="normal")

    # ── Log ───────────────────────────────────────────────────────────────
    def _log(self, msg: str, warn: bool = False) -> None:
        self._log_text.config(state="normal")
        tag = "warn" if warn else "normal"
        self._log_text.insert("end", msg + "\n", tag)
        self._log_text.tag_config("warn", foreground="#C55A11")
        self._log_text.see("end")
        self._log_text.config(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
