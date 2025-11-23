# PoliFin1.0.py
# PoliFin - Interfaz blanca con guinda, logos IPN & UPIIZ, ER por secciones y Balance por secciones
# Guardado JSON, exportar PDF (reportlab) y Excel (pandas/openpyxl). Usa Pillow para cargar imágenes.

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import cm
from datetime import datetime
import os

# Optional: Pillow for robust image loading in Tkinter
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except:
    PIL_AVAILABLE = False

# ---------- Paths to logos (usaste dos imágenes subidas) ----------
LOGO_IPN_PATH = "/mnt/data/d2c12a3e-e1cf-4863-b53f-e66afe37d81d.png"
LOGO_UPIIZ_PATH = "/mnt/data/a9567d8d-74a9-41bd-9f99-03170b6a2094.png"

# ---------- Theme colors & fonts ----------
BG = "#FFFFFF"             # fondo blanco
GUINDA = "#7A003C"         # guinda institucional (solicitado)
CARD = "#F6EEF2"           # tarjeta muy clara (tono cercano al blanco)
FG = "#0A0A0A"             # texto principal oscuro
ACCENT = GUINDA
TITLE_FONT = ("Segoe UI", 20, "bold")
SUB_FONT = ("Segoe UI", 11)

# ---------- Utilidades ----------
def to_float(val):
    """Convierte texto a float, tolera coma y cadenas vacías."""
    try:
        if val is None:
            return 0.0
        s = str(val).strip()
        if s == "":
            return 0.0
        s = s.replace(",", "")
        return float(s)
    except:
        return 0.0

def money(x):
    """Formatea número a cadena monetaria con comas y dos decimales."""
    try:
        return f"${float(x):,.2f}"
    except:
        return "$0.00"

# ---------- App ----------
class PoliFinApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PoliFin 1.0")
        self.root.geometry("1000x700")
        self.root.configure(bg=BG)

        # central data
        self.data = {}
        self.current_report = None

        # state for section navigation
        self.current_frame = None
        self.current_entries = {}

        # logos (tk images)
        self.logo_ipn_tk = None
        self.logo_upiiz_tk = None
        self.load_logos()

        # build UI
        self.build_main_menu()

    def load_logos(self):
        """Carga logos para usar en la interfaz (intenta Pillow primero)."""
        for path in (LOGO_IPN_PATH, LOGO_UPIIZ_PATH):
            if not os.path.exists(path):
                # skip if not present
                continue
        try:
            if PIL_AVAILABLE:
                # abrir y redimensionar
                ipn = Image.open(LOGO_IPN_PATH).convert("RGBA")
                ipn = ipn.resize((90, 120), Image.LANCZOS)
                self.logo_ipn_tk = ImageTk.PhotoImage(ipn)

                upiiz = Image.open(LOGO_UPIIZ_PATH).convert("RGBA")
                upiiz = upiiz.resize((90, 90), Image.LANCZOS)
                self.logo_upiiz_tk = ImageTk.PhotoImage(upiiz)
            else:
                # tkinter.PhotoImage soporta PNG; intentar cargar directo
                if os.path.exists(LOGO_IPN_PATH):
                    self.logo_ipn_tk = tk.PhotoImage(file=LOGO_IPN_PATH)
                if os.path.exists(LOGO_UPIIZ_PATH):
                    self.logo_upiiz_tk = tk.PhotoImage(file=LOGO_UPIIZ_PATH)
        except Exception:
            # no crítico: si no se puede cargar, dejamos None
            self.logo_ipn_tk = None
            self.logo_upiiz_tk = None

    # ---------- helpers UI ----------
    def clear(self):
        for w in self.root.winfo_children():
            w.destroy()
        self.current_frame = None
        self.current_entries = {}

    def header_bar(self, title_text):
        """Crea cabecera con logos y título (color guinda)."""
        header = tk.Frame(self.root, bg=BG)
        header.pack(fill="x", padx=10, pady=6)

        left = tk.Frame(header, bg=BG)
        left.pack(side="left")
        # logos
        if self.logo_ipn_tk:
            lbl_ipn = tk.Label(left, image=self.logo_ipn_tk, bg=BG)
            lbl_ipn.pack(side="left", padx=(0,8))
        else:
            lbl_ipn = tk.Label(left, text="IPN", bg=BG, fg=GUINDA, font=("Segoe UI", 12, "bold"))
            lbl_ipn.pack(side="left", padx=(0,8))
        if self.logo_upiiz_tk:
            lbl_up = tk.Label(left, image=self.logo_upiiz_tk, bg=BG)
            lbl_up.pack(side="left", padx=(0,8))
        else:
            lbl_up = tk.Label(left, text="UPIIZ", bg=BG, fg=GUINDA, font=("Segoe UI", 12, "bold"))
            lbl_up.pack(side="left", padx=(0,8))

        center = tk.Frame(header, bg=BG)
        center.pack(side="left", expand=True)
        lbl_title = tk.Label(center, text=title_text, fg=GUINDA, bg=BG, font=("Segoe UI", 22, "bold"))
        lbl_title.pack()

        right = tk.Frame(header, bg=BG)
        right.pack(side="right")
        fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
        tk.Label(right, text=fecha, bg=BG, fg=FG, font=("Segoe UI", 9)).pack()

    def add_field(self, parent, texto):
        """Añade label + entry y devuelve widget entry."""
        lbl = tk.Label(parent, text=texto, anchor="w", bg=BG, fg=FG, font=SUB_FONT)
        lbl.pack(fill="x", pady=(8,0))
        ent = tk.Entry(parent, width=30, font=("Segoe UI", 10))
        ent.pack(fill="x", pady=(2,6))
        return ent

    # ---------- Main menu ----------
    def build_main_menu(self):
        self.clear()
        self.header_bar("POLIFIN — Generador Financiero")

        frame = tk.Frame(self.root, bg=BG)
        frame.pack(pady=12)

        tk.Label(frame, text="Selecciona el estado financiero a generar:", bg=BG, fg=FG, font=SUB_FONT).pack(pady=(4,10))

        menu_frame = tk.Frame(frame, bg=BG)
        menu_frame.pack()

        tk.Button(menu_frame, text="Estado de Resultados", bg=GUINDA, fg="white", width=28,
                  font=SUB_FONT, relief="flat", command=self.start_er_sections).grid(row=0, column=0, padx=10, pady=8)
        tk.Button(menu_frame, text="Balance General", bg=GUINDA, fg="white", width=28,
                  font=SUB_FONT, relief="flat", command=self.start_balance_sections).grid(row=1, column=0, padx=10, pady=8)

        tk.Button(frame, text="Cargar archivo (JSON)", bg=CARD, fg=FG, width=24, command=self.load_file, relief="flat").pack(pady=8)
        tk.Button(frame, text="Exportar último reporte (PDF/Excel)", bg=CARD, fg=FG, width=28, command=self.export_menu, relief="flat").pack(pady=4)
        tk.Button(frame, text="Guardar datos actuales (JSON)", bg=CARD, fg=FG, width=28, command=self.save_file, relief="flat").pack(pady=4)

        tk.Label(self.root, text="Interfaz blanca con guinda — IPN / UPIIZ", bg=BG, fg=FG, font=("Segoe UI", 9)).pack(side="bottom", pady=8)

    # ----------------- ESTADO DE RESULTADOS (secciones) -----------------
    def start_er_sections(self):
        self.er_values = {}
        self.er_sections = [
            self.er_section_ventas,
            self.er_section_compras,
            self.er_section_gastos_venta,
            self.er_section_gastos_admin,
            self.er_section_financieros,
            self.er_section_otros,
            self.er_section_impuestos,
            self.er_er_calc_and_finish
        ]
        self.er_index = 0
        self.show_er_section()

    def show_er_section(self):
        if not (0 <= self.er_index < len(self.er_sections)):
            self.build_main_menu()
            return
        # clear UI and show
        self.clear()
        self.er_sections[self.er_index]()

    def er_save_current_entries(self):
        for k,w in self.current_entries.items():
            try:
                self.er_values[k] = to_float(w.get())
            except:
                self.er_values[k] = 0.0

    def er_next(self):
        self.er_save_current_entries()
        self.er_index += 1
        self.show_er_section()

    def er_prev(self):
        self.er_save_current_entries()
        if self.er_index > 0:
            self.er_index -= 1
        self.show_er_section()

    # ER section builders
    def er_section_ventas(self):
        self.header_bar("ESTADO DE RESULTADOS — Ventas y Compras")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12)
        frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Ventas totales"] = self.add_field(frame, "Ventas totales")
        self.current_entries["devoluciones sobre ventas"] = self.add_field(frame, "Devoluciones sobre ventas")
        self.current_entries["descuentos sobre ventas"] = self.add_field(frame, "Descuentos sobre ventas")
        # restore
        for k in list(self.current_entries.keys()):
            if k in self.er_values:
                self.current_entries[k].insert(0, str(self.er_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.er_next, relief="flat").pack(side="right")

    def er_section_compras(self):
        self.header_bar("ESTADO DE RESULTADOS — Compras Totales o Brutas")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12)
        frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["inventario inicial"] = self.add_field(frame, "Inventario inicial")
        self.current_entries["compras"] = self.add_field(frame, "Compras")
        self.current_entries["gastos de compra"] = self.add_field(frame, "Gastos de compra")
        self.current_entries["devoluciones sobre compras"] = self.add_field(frame, "Devoluciones sobre compras")
        self.current_entries["descuentos sobre compras"] = self.add_field(frame, "Descuentos sobre compras")
        self.current_entries["inventario final"] = self.add_field(frame, "Inventario final")
        for k in list(self.current_entries.keys()):
            if k in self.er_values:
                self.current_entries[k].insert(0, str(self.er_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.er_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.er_next, relief="flat").pack(side="right")

    def er_section_gastos_venta(self):
        self.header_bar("ESTADO DE RESULTADOS — Gastos de Venta")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Renta de almacen"] = self.add_field(frame, "Renta de almacén")
        self.current_entries["Propaganda y publicidad"] = self.add_field(frame, "Propaganda y publicidad")
        self.current_entries["Sueldos de agentes y dependeientes"] = self.add_field(frame, "Sueldos de agentes y dependientes")
        self.current_entries["comisiones de agentes y dependientes"] = self.add_field(frame, "Comisiones de agentes y dependientes")
        self.current_entries["consumo de luz del almacen"] = self.add_field(frame, "Consumo de luz del almacén")
        for k in list(self.current_entries.keys()):
            if k in self.er_values:
                self.current_entries[k].insert(0, str(self.er_values[k]))

        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.er_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.er_next, relief="flat").pack(side="right")

    def er_section_gastos_admin(self):
        self.header_bar("ESTADO DE RESULTADOS — Gastos de Administración")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Renta de oficinas"] = self.add_field(frame, "Renta de oficinas")
        self.current_entries["sueldos del personal de oficinas"] = self.add_field(frame, "Sueldos del personal de oficinas")
        self.current_entries["papeleria y utiles"] = self.add_field(frame, "Papelería y útiles")
        self.current_entries["consumo de luz de oficinas"] = self.add_field(frame, "Consumo de luz de oficinas")
        for k in list(self.current_entries.keys()):
            if k in self.er_values:
                self.current_entries[k].insert(0, str(self.er_values[k]))

        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.er_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.er_next, relief="flat").pack(side="right")

    def er_section_financieros(self):
        self.header_bar("ESTADO DE RESULTADOS — Productos y Gastos Financieros")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Intereses cobrados"] = self.add_field(frame, "Intereses cobrados")
        self.current_entries["ganancia en cambios"] = self.add_field(frame, "Ganancia en cambios")
        self.current_entries["intereses pagados"] = self.add_field(frame, "Intereses pagados")
        self.current_entries["perdida en cambios"] = self.add_field(frame, "Pérdida en cambios")
        for k in list(self.current_entries.keys()):
            if k in self.er_values:
                self.current_entries[k].insert(0, str(self.er_values[k]))

        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.er_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.er_next, relief="flat").pack(side="right")

    def er_section_otros(self):
        self.header_bar("ESTADO DE RESULTADOS — Otros Gastos / Otros Productos")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Perdida en venta de mobiliario"] = self.add_field(frame, "Pérdida en venta de mobiliario")
        self.current_entries["perdida en venta de acciones"] = self.add_field(frame, "Pérdida en venta de acciones")
        self.current_entries["Comisiones cobradas"] = self.add_field(frame, "Comisiones cobradas")
        self.current_entries["dividendos cobrados"] = self.add_field(frame, "Dividendos cobrados")
        self.current_entries["perdida entre otros gastos y productos"] = self.add_field(frame, "Pérdida entre otros gastos y productos")
        for k in list(self.current_entries.keys()):
            if k in self.er_values:
                self.current_entries[k].insert(0, str(self.er_values[k]))

        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.er_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.er_next, relief="flat").pack(side="right")

    def er_section_impuestos(self):
        self.header_bar("ESTADO DE RESULTADOS — Impuestos")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Impuesto  sobre la renta ISR"] = self.add_field(frame, "Impuesto sobre la renta ISR")
        self.current_entries["Participacion de los trabajadores en las utilidades"] = self.add_field(frame, "Participación de los trabajadores en las utilidades")
        for k in list(self.current_entries.keys()):
            if k in self.er_values:
                self.current_entries[k].insert(0, str(self.er_values[k]))

        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.er_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Calcular →", bg=GUINDA, fg="white", command=self.er_next, relief="flat").pack(side="right")

    def er_er_calc_and_finish(self):
        # Save last screen inputs
        self.er_save_current_entries()
        vals = self.er_values

        # --- Cálculos ---
        ventas_tot = vals.get("Ventas totales", 0)
        devV = vals.get("devoluciones sobre ventas", 0)
        descV = vals.get("descuentos sobre ventas", 0)
        ventas_net = ventas_tot - devV - descV

        inv_ini = vals.get("inventario inicial", 0)
        compras = vals.get("compras", 0)
        gastos_compra = vals.get("gastos de compra", 0)
        compras_totales = compras + gastos_compra
        devC = vals.get("devoluciones sobre compras", 0)
        descC = vals.get("descuentos sobre compras", 0)
        compras_netas = compras_totales - devC - descC
        suma_merc = inv_ini + compras_netas
        inv_fin = vals.get("inventario final", 0)
        costo_vendido = suma_merc - inv_fin

        utilidad_bruta = ventas_net - costo_vendido

        gastos_venta_total = vals.get("Renta de almacen", 0) + vals.get("Propaganda y publicidad", 0) + vals.get("Sueldos de agentes y dependeientes", 0) + vals.get("comisiones de agentes y dependientes", 0) + vals.get("consumo de luz del almacen", 0)
        gastos_admin_total = vals.get("Renta de oficinas", 0) + vals.get("sueldos del personal de oficinas", 0) + vals.get("papeleria y utiles", 0) + vals.get("consumo de luz de oficinas", 0)
        gastos_operacion_total = gastos_venta_total + gastos_admin_total

        productos_financieros = vals.get("Intereses cobrados", 0) + vals.get("ganancia en cambios", 0)
        gastos_financieros = vals.get("intereses pagados", 0) + vals.get("perdida en cambios", 0)

        utilidad_operacion = utilidad_bruta - gastos_operacion_total + productos_financieros - gastos_financieros

        otros_gastos = vals.get("Perdida en venta de mobiliario", 0) + vals.get("perdida en venta de acciones", 0)
        otros_productos = vals.get("Comisiones cobradas", 0) + vals.get("dividendos cobrados", 0)
        perdida_entre_otros = vals.get("perdida entre otros gastos y productos", 0)

        utilidad_antes_isr_ptu = utilidad_operacion - otros_gastos + otros_productos - perdida_entre_otros

        isr = vals.get("Impuesto  sobre la renta ISR", 0)
        ptu = vals.get("Participacion de los trabajadores en las utilidades", 0)
        utilidad_neta = utilidad_antes_isr_ptu - isr - ptu

        # store
        self.data["estado_resultados"] = {
            "Ventas totales": ventas_tot,
            "devoluciones sobre ventas": devV,
            "descuentos sobre ventas": descV,
            "ventas netas": ventas_net,
            "inventario inicial": inv_ini,
            "compras": compras,
            "gastos de compra": gastos_compra,
            "compras totales": compras_totales,
            "devoluciones sobre compras": devC,
            "descuentos sobre compras": descC,
            "compras netas": compras_netas,
            "suma o total de mercancías": suma_merc,
            "inventario final": inv_fin,
            "costo de lo vendido": costo_vendido,
            "utilidad bruta": utilidad_bruta,
            "gastos de operación": gastos_operacion_total,
            "gastos de venta detalle": {
                "Renta de almacen": vals.get("Renta de almacen", 0),
                "Propaganda y publicidad": vals.get("Propaganda y publicidad", 0),
                "Sueldos de agentes y dependeientes": vals.get("Sueldos de agentes y dependeientes", 0),
                "comisiones de agentes y dependientes": vals.get("comisiones de agentes y dependientes", 0),
                "consumo de luz del almacen": vals.get("consumo de luz del almacen", 0)
            },
            "gastos de administracion detalle": {
                "Renta de oficinas": vals.get("Renta de oficinas", 0),
                "sueldos del personal de oficinas": vals.get("sueldos del personal de oficinas", 0),
                "papeleria y utiles": vals.get("papeleria y utiles", 0),
                "consumo de luz de oficinas": vals.get("consumo de luz de oficinas", 0)
            },
            "productos_financieros": productos_financieros,
            "gastos_financieros": gastos_financieros,
            "utilidad_operacion": utilidad_operacion,
            "otros_gastos_detalle": {
                "Perdida en venta de mobiliario": vals.get("Perdida en venta de mobiliario", 0),
                "perdida en venta de acciones": vals.get("perdida en venta de acciones", 0)
            },
            "otros_productos_detalle": {
                "Comisiones cobradas": vals.get("Comisiones cobradas", 0),
                "dividendos cobrados": vals.get("dividendos cobrados", 0)
            },
            "perdida_entre_otros": perdida_entre_otros,
            "utilidad_antes_isr_ptu": utilidad_antes_isr_ptu,
            "ISR": isr,
            "PTU": ptu,
            "utilidad_neta": utilidad_neta
        }

        self.current_report = ("estado", "reporte")
        messagebox.showinfo("Resultado", f"Estado calculado. Utilidad neta: {money(utilidad_neta)}")
        self.show_er_summary()

    def show_er_summary(self):
        self.clear()
        self.header_bar("ESTADO DE RESULTADOS — Resultado")
        vals = self.data.get("estado_resultados", {})
        txt = tk.Text(self.root, width=100, height=26, font=("Consolas", 11))
        txt.pack(padx=12, pady=10)
        def line(n, v): return f"{n}: {v:,.2f}\n"
        s = ""
        order = ["Ventas totales","devoluciones sobre ventas","descuentos sobre ventas","ventas netas",
                 "inventario inicial","compras","gastos de compra","compras totales","devoluciones sobre compras","descuentos sobre compras",
                 "compras netas","suma o total de mercancías","inventario final","costo de lo vendido","utilidad bruta"]
        for k in order:
            s += line(k, vals.get(k,0))
        s += "\nGastos de operación:\n"
        for k,v in vals.get("gastos de venta detalle",{}).items():
            s += "   " + f"{k}: {v:,.2f}\n"
        s += "Gastos de administración:\n"
        for k,v in vals.get("gastos de administracion detalle",{}).items():
            s += "   " + f"{k}: {v:,.2f}\n"
        s += f"\nProductos financieros: {vals.get('productos_financieros',0):,.2f}\n"
        s += f"Gastos financieros: {vals.get('gastos_financieros',0):,.2f}\n"
        s += f"Utilidad de operación: {vals.get('utilidad_operacion',0):,.2f}\n"
        s += f"\nUtilidad antes de ISR y PTU: {vals.get('utilidad_antes_isr_ptu',0):,.2f}\n"
        s += f"ISR: {vals.get('ISR',0):,.2f}\nPTU: {vals.get('PTU',0):,.2f}\n"
        s += f"\nUTILIDAD NETA DEL EJERCICIO: {vals.get('utilidad_neta',0):,.2f}\n"
        txt.insert("1.0", s)

        btns = tk.Frame(self.root, bg=BG); btns.pack(pady=8)
        tk.Button(btns, text="Ver Reporte", bg=GUINDA, fg="white", command=lambda: self.view_er_report("reporte"), relief="flat").pack(side="left", padx=6)
        tk.Button(btns, text="Ver Cuenta", bg=CARD, fg=FG, command=lambda: self.view_er_report("cuenta"), relief="flat").pack(side="left", padx=6)
        tk.Button(btns, text="Exportar PDF/Excel", bg=CARD, fg=FG, command=self.export_menu, relief="flat").pack(side="left", padx=6)
        tk.Button(btns, text="Volver al menú", bg=CARD, fg=FG, command=self.build_main_menu, relief="flat").pack(side="left", padx=6)

    def view_er_report(self, mode="reporte"):
        vals = self.data.get("estado_resultados", {})
        self.current_report = ("estado", mode)
        self.clear()
        self.header_bar("ESTADO DE RESULTADOS — " + ("Reporte" if mode=="reporte" else "Cuenta"))
        t = tk.Text(self.root, width=100, height=30, font=("Consolas", 11))
        t.pack(padx=10, pady=8)
        if mode == "reporte":
            lines = []
            lines.append(f"Ventas netas: {vals.get('ventas netas',0):,.2f}\n")
            lines.append(f"Costo de lo vendido: {vals.get('costo de lo vendido',0):,.2f}\n")
            lines.append(f"Utilidad bruta: {vals.get('utilidad bruta',0):,.2f}\n")
            lines.append("\nGastos de operación:\n")
            for k,v in vals.get("gastos de venta detalle",{}).items():
                lines.append(f"  {k}: {v:,.2f}\n")
            for k,v in vals.get("gastos de administracion detalle",{}).items():
                lines.append(f"  {k}: {v:,.2f}\n")
            lines.append(f"\nUtilidad de operación: {vals.get('utilidad_operacion',0):,.2f}\n")
            lines.append(f"Utilidad antes de ISR y PTU: {vals.get('utilidad_antes_isr_ptu',0):,.2f}\n")
            lines.append(f"Utilidad neta: {vals.get('utilidad_neta',0):,.2f}\n")
            t.insert("1.0", "".join(lines))
        else:
            left = f"ACTIVOS (ej.)\nVentas netas: {vals.get('ventas netas',0):,.2f}\n"
            right = f"PASIVOS + CAPITAL (ej.)\nUtilidad neta: {vals.get('utilidad_neta',0):,.2f}\n"
            t.insert("1.0", left + "\n" + right)

        tk.Button(self.root, text="Volver", bg=CARD, fg=FG, command=self.show_er_summary, relief="flat").pack(pady=8)

    # ----------------- BALANCE GENERAL (secciones) -----------------
    def start_balance_sections(self):
        self.bal_values = {}
        self.b_sections = [
            self.b_section_activo_circulante,
            self.b_section_activo_no_circulante,
            self.b_section_activo_dif,
            self.b_section_pasivo_corto,
            self.b_section_pasivo_largo,
            self.b_section_pasivo_dif,
            self.b_balance_finalize
        ]
        self.b_index = 0
        self.show_balance_section()

    def show_balance_section(self):
        if not (0 <= self.b_index < len(self.b_sections)):
            self.build_main_menu()
            return
        self.clear()
        self.b_sections[self.b_index]()

    def b_save_current_entries(self):
        for k,w in self.current_entries.items():
            try:
                self.bal_values[k] = to_float(w.get())
            except:
                self.bal_values[k] = 0.0

    def b_next(self):
        self.b_save_current_entries()
        self.b_index += 1
        self.show_balance_section()

    def b_prev(self):
        self.b_save_current_entries()
        if self.b_index > 0:
            self.b_index -= 1
        self.show_balance_section()

    def b_section_activo_circulante(self):
        self.header_bar("BALANCE GENERAL — Activo Circulante")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Caja"] = self.add_field(frame, "Caja")
        self.current_entries["Bancos"] = self.add_field(frame, "Bancos")
        self.current_entries["Inversiones temporales"] = self.add_field(frame, "Inversiones temporales")
        self.current_entries["Mercancías"] = self.add_field(frame, "Mercancías")
        self.current_entries["Inventario o almacén"] = self.add_field(frame, "Inventario o almacén")
        self.current_entries["Clientes"] = self.add_field(frame, "Clientes")
        self.current_entries["Documentos por cobrar"] = self.add_field(frame, "Documentos por cobrar")
        self.current_entries["Deudores diversos"] = self.add_field(frame, "Deudores diversos")
        self.current_entries["Inventarios (circulante)"] = self.add_field(frame, "Anticipo a proveedores")
        for k in list(self.current_entries.keys()):
            if k in self.bal_values:
                self.current_entries[k].insert(0, str(self.bal_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.b_next, relief="flat").pack(side="right")

    def b_section_activo_no_circulante(self):
        self.header_bar("BALANCE GENERAL — Activo Fijo o No Circulante")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Terrenos"] = self.add_field(frame, "Terrenos")
        self.current_entries["Edificios"] = self.add_field(frame, "Edificios")
        self.current_entries["Mobiliario"] = self.add_field(frame, "Mobiliario y equipo")
        self.current_entries["Equipo de computo electrónico"] = self.add_field(frame, "Equipo de cómputo electrónico")
        self.current_entries["Equipo de entrega o reparto"] = self.add_field(frame, "Equipo de entrega o reparto")
        self.current_entries["Depositos en garantía"] = self.add_field(frame, "Dépositos en garantía")
        self.current_entries["Inversiones permanentes"] = self.add_field(frame, "Inversiones permanentes")
        for k in list(self.current_entries.keys()):
            if k in self.bal_values:
                self.current_entries[k].insert(0, str(self.bal_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.b_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.b_next, relief="flat").pack(side="right")

    def b_section_activo_dif(self):
        self.header_bar("BALANCE GENERAL — Activo Diferido o Cargas Diferidas")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Gastos de Inversión y Desarrollo"] = self.add_field(frame, "Gastos de Inversión y Desarrollo")
        self.current_entries["Gastos en Etapas Prosperativas, de Organización y Administración"] = self.add_field(frame, "Gastos en Etapas Prosperativas, de Organización y Administración")
        self.current_entries["Gastos de Mercadotecnia"] = self.add_field(frame, "Gastos de Mercadotecnia")
        self.current_entries["Gastos de instalación"] = self.add_field(frame, "Gastos de instalación")
        self.current_entries["Papeleria y útiles"] = self.add_field(frame, "Papeleria y útiles")
        self.current_entries["Propaganda y Publicidad"] = self.add_field(frame, "Propaganda y Publicidad")
        self.current_entries["Primas de seguros"] = self.add_field(frame, "Primas de seguros")
        self.current_entries["Rentas Pagadas por Anticipado"] = self.add_field(frame, "Rentas Pagadas por Anticipado")
        self.current_entries["Intereses Pagados por Anticipado"] = self.add_field(frame, "Intereses Pagados por Anticipado")
        for k in list(self.current_entries.keys()):
            if k in self.bal_values:
                self.current_entries[k].insert(0, str(self.bal_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.b_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.b_next, relief="flat").pack(side="right")


    def b_section_pasivo_corto(self):
        self.header_bar("BALANCE GENERAL — Pasivo a Corto Plazo o Circulante")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Proveedores"] = self.add_field(frame, "Proveedores")
        self.current_entries["Acreedores diversos"] = self.add_field(frame, "Acreedores diversos")
        self.current_entries["Documentos por pagar"] = self.add_field(frame, "Documentos por pagar")
        self.current_entries["Anticipo de clientes"] = self.add_field(frame, "Anticipo de Clientes")
        self.current_entries["Gastos Pendientes de Pago, por Pagar o Acumulados"] = self.add_field(frame, "Gastos Pendientes de Pago, por Pagar o Acumulados")
        self.current_entries["Impuestos Pendientes de Pago, por Pagar o Acumulados"] = self.add_field(frame, "Impuestos Pendientes de Pago, por Pagar o Acumulados")
        for k in list(self.current_entries.keys()):
            if k in self.bal_values:
                self.current_entries[k].insert(0, str(self.bal_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.b_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.b_next, relief="flat").pack(side="right")

    def b_section_pasivo_largo(self):
        self.header_bar("BALANCE GENERAL — Pasivo a Largo Plazo o Fijo")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Hipotecas por pagar o Acreedores Hipotecarios"] = self.add_field(frame, "Hipotecas por pagar o Acreedores Hipotecarios")
        self.current_entries["Documentos por Pagar a Largo Plazo"] = self.add_field(frame, "Documentos por Pagar a Largo Plazo")
        self.current_entries["Cuentas por Pagar a Largo Plazo"] = self.add_field(frame, "Cuentas por Pagar a Largo Plazo")
        for k in list(self.current_entries.keys()):
            if k in self.bal_values:
                self.current_entries[k].insert(0, str(self.bal_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.b_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Siguiente →", bg=GUINDA, fg="white", command=self.b_next, relief="flat").pack(side="right")

    def b_section_pasivo_dif(self):
        self.header_bar("BALANCE GENERAL — Pasivo Diferido o Créditos Diferidos")
        frame = tk.Frame(self.root, bg=BG, padx=20, pady=12); frame.pack(fill="both", expand=True)
        self.current_entries = {}
        self.current_frame = frame
        self.current_entries["Rentas Cobradas por Anticipado"] = self.add_field(frame, "Rentas Cobradas por Anticipado")
        self.current_entries["Intereses Cobrados por Anticipado"] = self.add_field(frame, "Intereses Cobrados por Anticipado")
        for k in list(self.current_entries.keys()):
            if k in self.bal_values:
                self.current_entries[k].insert(0, str(self.bal_values[k]))
        nav = tk.Frame(frame, bg=BG); nav.pack(fill="x", pady=12)
        tk.Button(nav, text="← Anterior", bg=CARD, fg=FG, command=self.b_prev, relief="flat").pack(side="left")
        tk.Button(nav, text="Generar balance →", bg=GUINDA, fg="white", command=self.b_next, relief="flat").pack(side="right")

    def b_balance_finalize(self):
        self.b_save_current_entries()
        bal = self.bal_values
        activo_cir = sum([bal.get(k,0) for k in ["Caja","Bancos","Inversiones Temporales", "Mercancías", "Inventario o Almacén", "Clientes", "Documentos por Cobrar", "Deudores Diversos", "Anticipo a Proveedores"]])
        activo_no_cir = sum([bal.get(k,0) for k in ["Terrenos","Edificios","Mobiliario y equipo","Equipo de computo", "Equipo de Entrega o Reparto", "Dépositos en Garantía", "Inversiones Permanentes"]]) 
        activo_dif = sum([bal.get(k,0) for k in ["Gastos de Inversión y Desarrollo", "Gastos en Etapas Prosperativas, de Organización y Administración","Gastos de Mercadotecnia","Gastos de Instalación","Papelería y Útiles","Propaganda y Publicidad","Primas de Seguros","Rentas Pagadas por Anticipado","Intereses Pagados por Anticipado"]]) 
        total_activos = activo_cir + activo_no_cir + activo_dif

        pasivo_corto = sum([bal.get(k,0) for k in ["Proveedores","Acreedores diversos","Documentos por pagar","Anticipo de Clientes","Gastos Pendientes de Pago, por Pagar o Acumulados","Impuestos Pendientes de Pago, por Pagar o Acumulados"]])
        pasivo_largo = sum([bal.get(k,0) for k in ["Hipotecas por pagar o Acreedores Hipotecarios","Documentos por Pagar a Largo Plazo","Cuentas por Pagar a Largo Plazo"]])
        pasivo_dif = sum([bal.get(k,0) for k in ["Proveedores","Rentas Cobradas por Anticipado","Intereses Cobrados por Anticipado"]])
        total_pasivos = pasivo_corto + pasivo_largo + pasivo_dif

        self.data["balance"] = {
            "Activo Circulante detalle": {k: bal.get(k,0) for k in ["Caja","Bancos","Inversiones Temporales", "Mercancías", "Inventario o Almacén", "Clientes", "Documentos por Cobrar", "Deudores Diversos", "Anticipo a Proveedores"]},
            "Activo No Circulante detalle": {k: bal.get(k,0) for k in ["Terrenos","Edificios","Mobiliario y equipo","Equipo de computo", "Equipo de Entrega o Reparto", "Dépositos en Garantía", "Inversiones Permanentes"]},
            "Activo Diferido detalle": {k: bal.get(k,0) for k in ["Gastos de Inversión y Desarrollo", "Gastos en Etapas Prosperativas, de Organización y Administración","Gastos de Mercadotecnia","Gastos de Instalación","Papelería y Útiles","Propaganda y Publicidad","Primas de Seguros","Rentas Pagadas por Anticipado","Intereses Pagados por Anticipado"]},
            "Pasivo Corto detalle": {k: bal.get(k,0) for k in ["Proveedores","Acreedores diversos","Documentos por pagar","Anticipo de Clientes","Gastos Pendientes de Pago, por Pagar o Acumulados","Impuestos Pendientes de Pago, por Pagar o Acumulados"]},
            "Pasivo Largo detalle": {k: bal.get(k,0) for k in ["Hipotecas por pagar o Acreedores Hipotecarios","Documentos por Pagar a Largo Plazo","Cuentas por Pagar a Largo Plazo"]},
            "Pasivo Diferido detalle": {k: bal.get(k,0) for k in["Proveedores","Rentas Cobradas por Anticipado","Intereses Cobrados por Anticipado"]},
            "totales": {
                "Total Activos": total_activos,
                "Total Pasivos": total_pasivos,
                "Capital Contable": total_activos - total_pasivos 
            }
        }

        self.clear()
        self.header_bar("BALANCE GENERAL — Generado")
        tk.Button(self.root, text="Ver como Reporte", bg=GUINDA, fg="white", command=lambda: self.view_balance("reporte"), relief="flat").pack(pady=8)
        tk.Button(self.root, text="Ver como Cuenta", bg=CARD, fg=FG, command=lambda: self.view_balance("cuenta"), relief="flat").pack(pady=8)
        tk.Button(self.root, text="Volver al menú", bg=CARD, fg=FG, command=self.build_main_menu, relief="flat").pack(pady=10)

    def view_balance(self, mode="reporte"):
        bal = self.data.get("balance", {})
        self.current_report = ("balance", mode)
        self.clear()
        self.header_bar("BALANCE GENERAL — " + ("Reporte" if mode=="reporte" else "Cuenta"))
        t = tk.Text(self.root, width=100, height=30, font=("Consolas",11))
        t.pack(padx=10, pady=8)
        tot = bal.get("totales",{})
        if mode=="reporte":
            s = "BALANCE GENERAL — REPORTE\n\nACTIVOS:\n"
            for k,v in bal.get("Activo Circulante detalle",{}).items():
                s += f"  {k}: {v:,.2f}\n"
            for k,v in bal.get("Activo No Circulante detalle",{}).items():
                s += f"  {k}: {v:,.2f}\n"
            for k,v in bal.get("Activo Diferido detalle",{}).items():
                s +=f"   {k}: {v:,.2f}\n"
            s += f"\nTOTAL ACTIVOS: {tot.get('Total Activos',0):,.2f}\n\nPASIVOS:\n"
            for k,v in bal.get("Pasivo Corto detalle",{}).items():
                s += f"  {k}: {v:,.2f}\n"
            for k,v in bal.get("Pasivo Largo detalle",{}).items():
                s += f"  {k}: {v:,.2f}\n"
            for k,v in bal.get("Pasivo Diferido detalle",{}).items():
                s += f"  {k}: {v:,.2f}\n"
            s += f"\nTOTAL PASIVOS: {tot.get('Total Pasivos',0):,.2f}\n\nCAPITAL:\n"
            for k,v in bal.get("Capital detalle",{}).items():
                s += f"  {k}: {v:,.2f}\n"
            s += f"\nCAPITAL CONTABLE: {tot.get('Capital Contable',0):,.2f}\n"
        else:
            s = f"BALANCE GENERAL — FORMA DE CUENTA\n\nACTIVOS: {tot.get('Total Activos',0):,.2f} \t CAPITAL CONTABLE: {tot.get('Capital Contable',0):,.2f}\n"
        t.insert("1.0", s)
        tk.Button(self.root, text="Volver", bg=CARD, fg=FG, command=self.build_main_menu, relief="flat").pack(pady=8)

    # ----------------- Guardar / Cargar (JSON) -----------------
    def save_file(self):
        f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files","*.json")])
        if not f:
            return
        with open(f, "w", encoding="utf-8") as fp:
            json.dump(self.data, fp, indent=4, ensure_ascii=False)
        messagebox.showinfo("Guardado", f"Datos guardados en:\n{f}")

    def load_file(self):
        f = filedialog.askopenfilename(filetypes=[("JSON files","*.json"),("All files","*.*")])
        if not f:
            return
        with open(f, "r", encoding="utf-8") as fp:
            self.data = json.load(fp)
        messagebox.showinfo("Cargado", f"Datos cargados desde:\n{f}")
        # show logical view
        if "estado_resultados" in self.data:
            self.current_report = ("estado", "reporte")
            self.show_er_summary()
        elif "balance" in self.data:
            self.current_report = ("balance", "reporte")
            self.view_balance("reporte")
        else:
            self.build_main_menu()

    # ----------------- Export -----------------
    def export_menu(self):
        if not self.current_report:
            messagebox.showerror("Error", "No hay reporte generado o seleccionado para exportar.")
            return
        self.clear()
        self.header_bar("EXPORTAR REPORTE")
        tk.Button(self.root, text="Exportar a PDF", bg=GUINDA, fg="white", width=24, command=self.export_pdf, relief="flat").pack(pady=8)
        tk.Button(self.root, text="Exportar a Excel", bg=GUINDA, fg="white", width=24, command=self.export_excel, relief="flat").pack(pady=8)
        tk.Button(self.root, text="Volver", bg=CARD, fg=FG, width=20, command=self.build_main_menu, relief="flat").pack(pady=10)

    def export_pdf(self):
        kind, fmt = self.current_report
        f = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files","*.pdf")])
        if not f:
            return
        doc = SimpleDocTemplate(f, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Logos on PDF header (if files exist)
        if os.path.exists(LOGO_IPN_PATH):
            try:
                rl = RLImage(LOGO_IPN_PATH, width=3*cm, height=4*cm)
                rl.hAlign = "LEFT"
                story.append(rl)
            except:
                pass
        if os.path.exists(LOGO_UPIIZ_PATH):
            try:
                rl2 = RLImage(LOGO_UPIIZ_PATH, width=2.5*cm, height=2.5*cm)
                rl2.hAlign = "RIGHT"
                story.append(rl2)
            except:
                pass

        story.append(Spacer(1, 8))
        story.append(Paragraph("PoliFin — Reporte Financiero", styles["Title"]))
        story.append(Spacer(1, 12))

        if kind == "estado":
            vals = self.data.get("estado_resultados", {})
            rows = [["Cuenta", "Monto"]]
            order = ["Ventas totales","devoluciones sobre ventas","descuentos sobre ventas","ventas netas",
                     "inventario inicial","compras","gastos de compra","compras totales","devoluciones sobre compras","descuentos sobre compras",
                     "compras netas","suma o total de mercancías","inventario final","costo de lo vendido","utilidad bruta"]
            for k in order:
                rows.append([k, f"{vals.get(k,0):,.2f}"])
            rows.append(["Gastos de operación", ""])
            for k,v in vals.get("gastos de venta detalle",{}).items():
                rows.append([f"  {k}", f"{v:,.2f}"])
            for k,v in vals.get("gastos de administracion detalle",{}).items():
                rows.append([f"  {k}", f"{v:,.2f}"])
            rows.append(["Productos financieros", f"{vals.get('productos_financieros',0):,.2f}"])
            rows.append(["Gastos financieros", f"{vals.get('gastos_financieros',0):,.2f}"])
            rows.append(["Utilidad de operación", f"{vals.get('utilidad_operacion',0):,.2f}"])
            rows.append(["Utilidad antes de ISR y PTU", f"{vals.get('utilidad_antes_isr_ptu',0):,.2f}"])
            rows.append(["ISR", f"{vals.get('ISR',0):,.2f}"])
            rows.append(["PTU", f"{vals.get('PTU',0):,.2f}"])
            rows.append(["UTILIDAD NETA DEL EJERCICIO", f"{vals.get('utilidad_neta',0):,.2f}"])
            t = Table(rows, colWidths=[360, 140])
            t.setStyle(TableStyle([
                ("GRID",(0,0),(-1,-1),0.3,colors.grey),
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor(GUINDA)),
                ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("ALIGN",(1,1),(-1,-1),"RIGHT")
            ]))
            story.append(t)
        else:
            bal = self.data.get("balance", {})
            rows = [["Cuenta", "Monto"]]
            rows.append(["ACTIVOS", ""])
            for k,v in bal.get("Activo Circulante detalle",{}).items():
                rows.append([f"  {k}", f"{v:,.2f}"])
            for k,v in bal.get("Activo No Circulante detalle",{}).items():
                rows.append([f"  {k}", f"{v:,.2f}"])
            for k,v in bal.get("Activo Diferido detalle",{}).items():
                rows.append([f"  {k}", f"{v:,.2f}"])
            rows.append(["TOTALES", ""])
            for k,v in bal.get("totales",{}).items():
                rows.append([k, f"{v:,.2f}"])
            t = Table(rows, colWidths=[360,140])
            t.setStyle(TableStyle([
                ("GRID",(0,0),(-1,-1),0.3,colors.grey),
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor(GUINDA)),
                ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("ALIGN",(1,1),(-1,-1),"RIGHT")
            ]))
            story.append(t)

        doc.build(story)
        messagebox.showinfo("PDF generado", f"Se creó el PDF:\n{f}")

    def export_excel(self):
        kind, fmt = self.current_report or (None, None)
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")])
        if not f:
            return

        if kind == "estado":
            vals = self.data.get("estado_resultados", {})
            rows = []
            order = ["Ventas totales","devoluciones sobre ventas","descuentos sobre ventas","ventas netas",
                     "inventario inicial","compras","gastos de compra","compras totales","devoluciones sobre compras","descuentos sobre compras",
                     "compras netas","suma o total de mercancías","inventario final","costo de lo vendido","utilidad bruta"]
            for k in order:
                rows.append({"Cuenta":k, "Monto": vals.get(k,0)})
            rows.append({"Cuenta":"Gastos de operación", "Monto": ""})
            for k,v in vals.get("gastos de venta detalle",{}).items():
                rows.append({"Cuenta":"   "+k, "Monto": v})
            for k,v in vals.get("gastos de administracion detalle",{}).items():
                rows.append({"Cuenta":"   "+k, "Monto": v})
            rows.append({"Cuenta":"Productos financieros", "Monto": vals.get("productos_financieros",0)})
            rows.append({"Cuenta":"Gastos financieros", "Monto": vals.get("gastos_financieros",0)})
            rows.append({"Cuenta":"Utilidad de operación", "Monto": vals.get("utilidad_operacion",0)})
            rows.append({"Cuenta":"Utilidad antes de ISR y PTU", "Monto": vals.get("utilidad_antes_isr_ptu",0)})
            rows.append({"Cuenta":"ISR", "Monto": vals.get("ISR",0)})
            rows.append({"Cuenta":"PTU", "Monto": vals.get("PTU",0)})
            rows.append({"Cuenta":"UTILIDAD NETA DEL EJERCICIO", "Monto": vals.get("utilidad_neta",0)})
            df = pd.DataFrame(rows)
            df.to_excel(f, index=False, sheet_name="EstadoResultados")
        elif kind == "balance":
            bal = self.data.get("balance", {})
            rows = []
            rows.append({"Cuenta":"ACTIVOS", "Monto": ""})
            for k,v in bal.get("Activo Circulante detalle",{}).items():
                rows.append({"Cuenta":"  "+k, "Monto": v})
            for k,v in bal.get("Activo No Circulante detalle",{}).items():
                rows.append({"Cuenta":"  "+k, "Monto": v})
            rows.append({"Cuenta":"TOTALES", "Monto": ""})
            for k,v in bal.get("totales",{}).items():
                rows.append({"Cuenta": k, "Monto": v})
            df = pd.DataFrame(rows)
            df.to_excel(f, index=False, sheet_name="BalanceGeneral")
        else:
            messagebox.showerror("Error", "No hay reporte seleccionado para exportar.")
            return

        messagebox.showinfo("Excel generado", f"Se creó el archivo Excel:\n{f}")

# ---------- run ----------
def main():
    root = tk.Tk()
    app = PoliFinApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()