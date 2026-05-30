"""Analizador SOR — Fibra Óptica (GUI principal)"""
from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import config as cfg_mod
from sor_parser import scan_folder, parse_sor
from excel_exporter import export_to_excel, build_output_path, OPENPYXL_OK

# ── Drag-and-drop (opcional) ───────────────────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _BASE = TkinterDnD.Tk
    HAS_DND = True
except ImportError:
    _BASE = tk.Tk
    HAS_DND = False

# ── Colores ────────────────────────────────────────────────────────────────
C_BG    = "#F0F4F8"
C_DARK  = "#1F4E79"
C_MED   = "#2E75B6"
C_LIGHT = "#BDD7EE"
C_TEXT  = "#222222"


class App(_BASE):
    def __init__(self) -> None:
        super().__init__()
        self.title("Analizador SOR — Fibra Óptica")
        self.geometry("900x640")
        self.minsize(780, 560)
        self.configure(bg=C_BG)

        self._cfg = cfg_mod.load()
        self._cables: dict = {}
        self._parsed: dict = {}
        self._filter_vars: dict = {}   # {cable_name: {fibra_num: BooleanVar}}

        self._build_ui()
        self._restore_last_folder()

    # ── Restore ───────────────────────────────────────────────────────────
    def _restore_last_folder(self) -> None:
        folder = self._cfg.get('last_folder', '')
        if folder and os.path.isdir(folder):
            self._folder_var.set(folder)
            self._detect_cables(folder)

    # ── UI ────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        self._build_titlebar()
        self._build_folder_section()
        self._build_options_section()
        self._build_cables_section()
        self._build_progress_section()
        self._build_preview_section()
        self._build_buttons()

    def _build_titlebar(self) -> None:
        bar = tk.Frame(self, bg=C_DARK, pady=6)
        bar.pack(fill="x")
        tk.Label(bar, text="Analizador SOR — Fibra Óptica",
                 font=("Segoe UI", 13, "bold"),
                 fg="white", bg=C_DARK).pack(side="left", padx=14)
        tk.Button(bar, text="⚙  Configuración",
                  command=self._open_settings,
                  bg=C_MED, fg="white",
                  font=("Segoe UI", 9), relief="flat",
                  padx=10, pady=3).pack(side="right", padx=10)

    def _build_folder_section(self) -> None:
        frm = tk.LabelFrame(self, text=" Carpeta de archivos SOR ",
                            bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"))
        frm.pack(fill="x", padx=10, pady=(8, 4))

        inner = tk.Frame(frm, bg=C_BG)
        inner.pack(fill="x", padx=6, pady=4)

        self._folder_var = tk.StringVar(value="(ninguna carpeta seleccionada)")
        path_lbl = tk.Label(inner, textvariable=self._folder_var,
                            bg=C_BG, fg="#555", font=("Segoe UI", 9), anchor="w")
        path_lbl.pack(side="left", fill="x", expand=True)

        tk.Button(inner, text="Seleccionar Carpeta",
                  command=self._select_folder,
                  bg=C_MED, fg="white", font=("Segoe UI", 9, "bold"),
                  relief="flat", padx=10, pady=4).pack(side="right")

        if HAS_DND:
            drop_frm = tk.Frame(frm, bg=C_LIGHT, pady=6)
            drop_frm.pack(fill="x", padx=6, pady=(0, 4))
            tk.Label(drop_frm,
                     text="↓  Arrastra aquí una carpeta  ↓",
                     bg=C_LIGHT, fg=C_DARK,
                     font=("Segoe UI", 9)).pack()
            drop_frm.drop_target_register(DND_FILES)
            drop_frm.dnd_bind('<<Drop>>', self._on_drop)
        else:
            tk.Label(frm,
                     text="Instala tkinterdnd2 para habilitar drag & drop",
                     bg=C_BG, fg="#999", font=("Segoe UI", 8)).pack(anchor="w", padx=6)

    def _build_options_section(self) -> None:
        frm = tk.LabelFrame(self, text=" Opciones de procesamiento ",
                            bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"))
        frm.pack(fill="x", padx=10, pady=4)

        inner = tk.Frame(frm, bg=C_BG)
        inner.pack(fill="x", padx=6, pady=6)

        tk.Label(inner, text="Direcciones:", bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        self._dir_vars: dict[str, tk.BooleanVar] = {}
        for key, label in [('normal', 'Normal (bidireccional)'),
                            ('corta',  'Corta'),
                            ('larga',  'Larga')]:
            var = tk.BooleanVar(value=(key in self._cfg.get('directions', ['normal'])))
            self._dir_vars[key] = var
            tk.Checkbutton(inner, text=label, variable=var,
                           bg=C_BG, fg=C_TEXT,
                           font=("Segoe UI", 9),
                           command=self._on_options_change).pack(side="left", padx=8)

    def _build_cables_section(self) -> None:
        frm = tk.LabelFrame(self, text=" Cables detectados ",
                            bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"))
        frm.pack(fill="x", padx=10, pady=4)
        self._cable_frame = tk.Frame(frm, bg=C_BG)
        self._cable_frame.pack(fill="x", padx=6, pady=4)
        tk.Label(self._cable_frame, text="—", bg=C_BG, fg="#999").pack(anchor="w")

    def _build_progress_section(self) -> None:
        frm = tk.LabelFrame(self, text=" Progreso ",
                            bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"))
        frm.pack(fill="x", padx=10, pady=4)
        self._progress = ttk.Progressbar(frm, mode="determinate")
        self._progress.pack(fill="x", padx=8, pady=(4, 2))
        self._status_var = tk.StringVar(value="Listo.")
        tk.Label(frm, textvariable=self._status_var,
                 bg=C_BG, fg="#444", font=("Segoe UI", 9), anchor="w").pack(
                     fill="x", padx=8, pady=(0, 4))

    def _build_preview_section(self) -> None:
        frm = tk.LabelFrame(self, text=" Vista previa de datos ",
                            bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"))
        frm.pack(fill="both", expand=True, padx=10, pady=4)

        cols = ("fibra", "dir", "evento", "pos_km", "long_km",
                "perd_int", "prom_dbkm", "union_db")
        hdrs = ("Fibra", "Dirección", "N° Ev.", "Posición (km)",
                "Long. Int. (km)", "Pérd. Int. (dB)", "Prom. (dB/km)", "Unión (dB)")

        self._tree = ttk.Treeview(frm, columns=cols, show="headings",
                                  height=8, selectmode="browse")
        widths = (55, 100, 55, 100, 110, 110, 110, 90)
        for c, h, w in zip(cols, hdrs, widths):
            self._tree.heading(c, text=h)
            self._tree.column(c, width=w, anchor="center", minwidth=50)

        sb = ttk.Scrollbar(frm, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True, padx=4, pady=4)

        # Tag styles for threshold warnings
        self._tree.tag_configure("warn",  background="#FFD7D7")
        self._tree.tag_configure("caution", background="#FFE8CC")

    def _build_buttons(self) -> None:
        frm = tk.Frame(self, bg=C_BG)
        frm.pack(fill="x", padx=10, pady=(4, 10))

        self._btn_export = tk.Button(
            frm, text="Exportar Excel",
            command=self._start_export,
            bg=C_DARK, fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", padx=20, pady=6, state="disabled")
        self._btn_export.pack(side="right")

        tk.Button(frm, text="Analizar archivos",
                  command=self._start_analyze,
                  bg=C_MED, fg="white",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=20, pady=6).pack(side="right", padx=(0, 8))

        self._status_lbl = tk.Label(frm, text="", bg=C_BG, fg="#666",
                                    font=("Segoe UI", 8))
        self._status_lbl.pack(side="left")

    # ── Folder handling ───────────────────────────────────────────────────
    def _select_folder(self) -> None:
        folder = filedialog.askdirectory(title="Seleccionar carpeta raíz con subcarpetas por cable")
        if folder:
            self._load_folder(folder)

    def _on_drop(self, event) -> None:
        path = event.data.strip('{}').strip('"')
        if os.path.isdir(path):
            self._load_folder(path)

    def _load_folder(self, folder: str) -> None:
        self._folder_var.set(folder)
        self._cfg['last_folder'] = folder
        cfg_mod.save(self._cfg)
        self._parsed = {}
        self._btn_export.config(state="disabled")
        self._detect_cables(folder)

    def _detect_cables(self, folder: str) -> None:
        directions = self._selected_directions()
        self._cables = scan_folder(folder, directions)
        self._filter_vars = {
            cable: {item['fibra']: tk.BooleanVar(value=True)
                    for item in fibras}
            for cable, fibras in self._cables.items()
        }
        self._apply_saved_filter()
        self._update_cable_list()

    def _apply_saved_filter(self) -> None:
        saved = self._cfg.get('fibra_filter', {})
        for cable, fvars in self._filter_vars.items():
            saved_list = saved.get(cable, [])
            if saved_list:
                for num, var in fvars.items():
                    var.set(num in saved_list)

    def _selected_directions(self) -> list[str]:
        dirs = [k for k, v in self._dir_vars.items() if v.get()]
        return dirs or ['normal']

    def _on_options_change(self) -> None:
        self._cfg['directions'] = self._selected_directions()
        cfg_mod.save(self._cfg)
        if self._cables and self._cfg.get('last_folder'):
            self._detect_cables(self._cfg['last_folder'])

    # ── Cable list widget ─────────────────────────────────────────────────
    def _update_cable_list(self) -> None:
        for w in self._cable_frame.winfo_children():
            w.destroy()
        if not self._cables:
            tk.Label(self._cable_frame, text="—", bg=C_BG, fg="#999").pack(anchor="w")
            return
        total_f = sum(len(v) for v in self._cables.values())
        for cable, fibras in self._cables.items():
            row = tk.Frame(self._cable_frame, bg=C_BG)
            row.pack(fill="x", pady=1)
            selected = sum(1 for v in self._filter_vars.get(cable, {}).values() if v.get())
            tk.Label(row,
                     text=f"✓  {cable.strip():<45} ({selected}/{len(fibras)} fibras)",
                     bg=C_BG, fg=C_DARK, font=("Segoe UI", 9)).pack(side="left")
            tk.Button(row, text="Filtrar fibras",
                      command=lambda c=cable: self._open_filter_dialog(c),
                      bg=C_LIGHT, fg=C_DARK, font=("Segoe UI", 8),
                      relief="flat", padx=6).pack(side="left", padx=4)
        tk.Label(self._cable_frame,
                 text=f"Total: {len(self._cables)} cable(s), {total_f} fibra(s)",
                 bg=C_BG, fg="#666", font=("Segoe UI", 8)).pack(anchor="w", pady=(4, 0))

    # ── Settings dialog ───────────────────────────────────────────────────
    def _open_settings(self) -> None:
        dlg = tk.Toplevel(self)
        dlg.title("Configuración")
        dlg.geometry("420x340")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.configure(bg=C_BG)

        pad = {"padx": 12, "pady": 6}

        # OTDR equipment
        frm_eq = tk.LabelFrame(dlg, text=" Equipo OTDR ", bg=C_BG, fg=C_DARK,
                                font=("Segoe UI", 9, "bold"))
        frm_eq.pack(fill="x", **pad)
        equipment_labels = {k: v['label'] for k, v in cfg_mod.EQUIPMENT_PROFILES.items()}
        eq_var = tk.StringVar(value=equipment_labels.get(self._cfg.get('otdr_equipment', 'exfo')))
        eq_combo = ttk.Combobox(frm_eq, textvariable=eq_var,
                                values=list(equipment_labels.values()),
                                state="readonly", width=28)
        eq_combo.pack(padx=8, pady=6)

        # Thresholds
        frm_thr = tk.LabelFrame(dlg, text=" Umbrales de alerta (color naranja/rojo) ",
                                 bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"))
        frm_thr.pack(fill="x", **pad)
        thr = self._cfg.get('thresholds', {})
        thr_fields: list[tuple[str, str, tk.StringVar]] = []
        for label, key, unit in [
            ("Pérd. de unión máx.",   'perdida_union_db',      "dB"),
            ("Atenuación máx.",       'perdida_promedio_dbkm', "dB/km"),
            ("Pérd. de intervalo máx.", 'perdida_intervalo_db', "dB"),
        ]:
            row = tk.Frame(frm_thr, bg=C_BG)
            row.pack(fill="x", padx=8, pady=2)
            tk.Label(row, text=f"{label}:", bg=C_BG, fg=C_TEXT,
                     font=("Segoe UI", 9), width=28, anchor="w").pack(side="left")
            var = tk.StringVar(value=str(thr.get(key, '')))
            tk.Entry(row, textvariable=var, width=8).pack(side="left")
            tk.Label(row, text=unit, bg=C_BG, fg="#666",
                     font=("Segoe UI", 9)).pack(side="left", padx=4)
            thr_fields.append((key, label, var))

        # History file
        frm_hist = tk.LabelFrame(dlg, text=" Archivo de historial ",
                                  bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"))
        frm_hist.pack(fill="x", **pad)
        hist_row = tk.Frame(frm_hist, bg=C_BG)
        hist_row.pack(fill="x", padx=8, pady=4)
        hist_var = tk.StringVar(value=self._cfg.get('history_path', ''))
        tk.Entry(hist_row, textvariable=hist_var, width=36).pack(side="left")
        tk.Button(hist_row, text="...",
                  command=lambda: hist_var.set(
                      filedialog.askopenfilename(filetypes=[("JSON", "*.json")])),
                  bg=C_LIGHT, relief="flat", padx=4).pack(side="left", padx=4)

        # Buttons
        btn_row = tk.Frame(dlg, bg=C_BG)
        btn_row.pack(fill="x", padx=12, pady=8)

        def _save():
            label_to_key = {v: k for k, v in equipment_labels.items()}
            self._cfg['otdr_equipment'] = label_to_key.get(eq_var.get(), 'exfo')
            for key, _, var in thr_fields:
                try:
                    self._cfg['thresholds'][key] = float(var.get())
                except ValueError:
                    pass
            self._cfg['history_path'] = hist_var.get()
            cfg_mod.save(self._cfg)
            dlg.destroy()

        tk.Button(btn_row, text="Cancelar", command=dlg.destroy,
                  bg="#DDD", relief="flat", padx=14, pady=5).pack(side="right")
        tk.Button(btn_row, text="Guardar", command=_save,
                  bg=C_DARK, fg="white", font=("Segoe UI", 9, "bold"),
                  relief="flat", padx=14, pady=5).pack(side="right", padx=(0, 6))

    # ── Fiber filter dialog ───────────────────────────────────────────────
    def _open_filter_dialog(self, cable_name: str) -> None:
        fvars = self._filter_vars.get(cable_name, {})
        if not fvars:
            return

        dlg = tk.Toplevel(self)
        dlg.title(f"Filtrar fibras — {cable_name.strip()}")
        dlg.geometry("300x400")
        dlg.grab_set()
        dlg.configure(bg=C_BG)

        tk.Label(dlg, text=f"Selecciona fibras a procesar\n{cable_name.strip()}",
                 bg=C_BG, fg=C_DARK, font=("Segoe UI", 9, "bold"),
                 justify="center").pack(pady=8)

        # Scrollable frame
        canvas = tk.Canvas(dlg, bg=C_BG, highlightthickness=0)
        sb = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)
        scroll_frm = tk.Frame(canvas, bg=C_BG)
        scroll_frm.bind("<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frm, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=8)
        sb.pack(side="right", fill="y")

        for num in sorted(fvars.keys()):
            tk.Checkbutton(scroll_frm,
                           text=f"  Fibra {num}",
                           variable=fvars[num],
                           bg=C_BG, fg=C_TEXT,
                           font=("Segoe UI", 9)).pack(anchor="w")

        btn_row = tk.Frame(dlg, bg=C_BG)
        btn_row.pack(fill="x", padx=8, pady=6)
        tk.Button(btn_row, text="Todos",
                  command=lambda: [v.set(True) for v in fvars.values()],
                  bg=C_LIGHT, relief="flat", padx=8).pack(side="left")
        tk.Button(btn_row, text="Ninguno",
                  command=lambda: [v.set(False) for v in fvars.values()],
                  bg=C_LIGHT, relief="flat", padx=8).pack(side="left", padx=4)

        def _apply():
            saved = self._cfg.setdefault('fibra_filter', {})
            saved[cable_name] = [n for n, v in fvars.items() if v.get()]
            cfg_mod.save(self._cfg)
            self._update_cable_list()
            dlg.destroy()

        tk.Button(btn_row, text="Aplicar", command=_apply,
                  bg=C_DARK, fg="white", font=("Segoe UI", 9, "bold"),
                  relief="flat", padx=14).pack(side="right")

    # ── Preview update ────────────────────────────────────────────────────
    def _update_preview(self) -> None:
        for item in self._tree.get_children():
            self._tree.delete(item)

        thr = self._cfg.get('thresholds', {})
        thr_union = thr.get('perdida_union_db', 0.5)
        thr_prom  = thr.get('perdida_promedio_dbkm', 0.25)
        thr_int   = thr.get('perdida_intervalo_db', 2.0)

        count = 0
        for cable_fibers in self._parsed.values():
            for fiber in cable_fibers:
                fibra_num = fiber.get('fibra_num', '?')
                direction = fiber.get('direction', 'normal')
                for ev in fiber.get('events', []):
                    if count >= 500:
                        self._tree.insert("", "end",
                                          values=("…", "…", "…", "…", "…", "…", "…", "…"))
                        return
                    pu  = ev.get('perdida_union_db', 0)
                    pp  = ev.get('perdida_promedio_dbkm', 0)
                    pi  = ev.get('perdida_intervalo_db', 0)
                    tag = ""
                    if pu > thr_union or pp > thr_prom or pi > thr_int:
                        tag = "warn"
                    elif pu > thr_union * 0.8 or pp > thr_prom * 0.8:
                        tag = "caution"
                    self._tree.insert("", "end", tags=(tag,), values=(
                        fibra_num,
                        direction,
                        ev.get('n_evento', ''),
                        f"{ev.get('posicion_km', 0):.4f}",
                        f"{ev.get('longitud_intervalo_km', 0):.4f}",
                        f"{pi:.4f}",
                        f"{pp:.4f}",
                        f"{pu:.4f}",
                    ))
                    count += 1

    # ── Analyze ───────────────────────────────────────────────────────────
    def _start_analyze(self) -> None:
        if not self._cfg.get('last_folder'):
            messagebox.showwarning("Sin carpeta", "Selecciona primero una carpeta.")
            return
        if not self._cables:
            messagebox.showwarning("Sin SOR", "No se encontraron archivos SOR.")
            return
        threading.Thread(target=self._run_analyze, daemon=True).start()

    def _run_analyze(self) -> None:
        self._btn_export.config(state="disabled")
        self._parsed = {}

        equipment = self._cfg.get('otdr_equipment', 'exfo')
        ffilter   = self._cfg.get('fibra_filter', {})

        # Build flat list of (cable, fibra_num, direction, path)
        tasks: list[tuple[str, int, str, str]] = []
        for cable, fibras in self._cables.items():
            allowed = set(ffilter.get(cable, []))
            for item in fibras:
                num = item['fibra']
                if allowed and num not in allowed:
                    continue
                for direction, path in item['paths'].items():
                    tasks.append((cable, num, direction, path))

        self._progress["maximum"] = max(len(tasks), 1)
        self._progress["value"]   = 0
        errors = 0

        for i, (cable, fibra_num, direction, path) in enumerate(tasks):
            self._status_var.set(
                f"Analizando fibra {fibra_num} ({direction}) — {cable.strip()} …")
            result = parse_sor(path, equipment=equipment)
            if result and 'error' not in result:
                result['fibra_num']  = fibra_num
                result['direction']  = direction
                self._parsed.setdefault(cable, []).append(result)
            else:
                err = result.get('error', 'error') if result else 'None'
                self._status_lbl.config(
                    text=f"⚠ Fibra {fibra_num} ({direction}): {err[:60]}")
                errors += 1
            self._progress["value"] = i + 1
            self.update_idletasks()

        ok = len(tasks) - errors
        self._status_var.set(
            f"Análisis completo: {ok}/{len(tasks)} OK, {errors} error(es).")

        self._update_preview()
        if self._parsed:
            self._btn_export.config(state="normal")

    # ── Export ────────────────────────────────────────────────────────────
    def _start_export(self) -> None:
        if not self._parsed:
            messagebox.showwarning("Sin datos", "Primero analiza los archivos.")
            return
        threading.Thread(target=self._run_export, daemon=True).start()

    def _run_export(self) -> None:
        self._btn_export.config(state="disabled")
        self._status_var.set("Generando Excel …")

        ref_date = None
        for fibers in self._parsed.values():
            if fibers and fibers[0].get('date'):
                ref_date = fibers[0]['date']
                break

        output_path  = build_output_path(self._cfg.get('last_folder', ''), ref_date)
        history_path = self._cfg.get('history_path', '')
        if not history_path and self._cfg.get('last_folder'):
            from history import default_path
            history_path = default_path(self._cfg['last_folder'])

        try:
            export_to_excel(
                self._parsed,
                output_path,
                thresholds=self._cfg.get('thresholds'),
                history_path=history_path,
            )
            self._status_var.set(f"✓ Exportado: {os.path.basename(output_path)}")
            if sys.platform.startswith("win"):
                os.startfile(output_path)
        except ImportError as e:
            messagebox.showerror("Dependencia faltante", str(e))
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))
            self._status_var.set(f"Error: {e}")
        finally:
            self._btn_export.config(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
