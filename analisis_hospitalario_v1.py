"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  MOTOR ANALÍTICO HOSPITALARIO  v1.0                                         ║
║  Contratación EPS·IPS · Epidemiología · ML · Costos · Predicción           ║
║  Módulos: Hospitalización · Cirugía · UCI · Ambulatorio · Financiero · ML  ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  Uso:                                                                        ║
║    python analisis_hospitalario_v1.py                  (diálogo)            ║
║    python analisis_hospitalario_v1.py --excel base.xlsx                     ║
║    python analisis_hospitalario_v1.py --excel base.xlsx --out resultados    ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

EXCEL_DEFAULT = None  # ← ruta hardcodeada opcional: r"C:\datos\base.xlsx"

import argparse, json, os, sys, warnings
from datetime import datetime
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

from scipy.stats import chi2_contingency, kruskal, pearsonr, spearmanr
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.tree import DecisionTreeRegressor, DecisionTreeClassifier
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.metrics import (mean_absolute_error, mean_squared_error, r2_score,
                             silhouette_score, accuracy_score, f1_score)
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.impute import SimpleImputer

try:
    import statsmodels.api as sm
    HAS_SM = True
except ImportError:
    HAS_SM = False

# ─── Paleta visual dark ───────────────────────────────────────────────────────
BG   = "#080818"; BG2 = "#0E0E1E"; BG3 = "#131328"; ACC = "#00F5D4"
TEXT = "#DDE4F0"; TXDS = "#6B7899"
PAL  = ["#00F5D4","#38C4FF","#FF5FAA","#FFE040","#A855F7",
        "#FF7A00","#00C875","#0066CC","#CC2E6A","#C4A800","#6A3AAD","#08B0A0"]
CMAP_HEAT = sns.diverging_palette(240, 10, as_cmap=True)

plt.rcParams.update({
    "figure.facecolor":BG,"axes.facecolor":BG2,"axes.edgecolor":"#1E1E3A",
    "axes.labelcolor":TEXT,"axes.titlecolor":TEXT,"xtick.color":TEXT,
    "ytick.color":TEXT,"text.color":TEXT,"grid.color":"#1A1A32","grid.alpha":0.5,
    "font.family":"DejaVu Sans","figure.dpi":110,"savefig.dpi":110,
})

# ─── CLI ─────────────────────────────────────────────────────────────────────
def parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=None)
    ap.add_argument("--out",   default="output_hospitalario")
    return ap.parse_args()

def seleccionar_excel(default=None):
    if default and os.path.exists(default): return default
    try:
        import tkinter as tk; from tkinter import filedialog
        r = tk.Tk(); r.withdraw(); r.attributes("-topmost", True)
        p = filedialog.askopenfilename(title="Seleccionar Excel",
            filetypes=[("Excel","*.xlsx *.xls"),("All","*.*")]); r.destroy()
        if p: return p
    except Exception: pass
    return input("Ruta al Excel: ").strip().strip('"').strip("'")

# ═══════════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ═══════════════════════════════════════════════════════════════════════════════
def fc(df, candidates, default=None):
    """find_col: devuelve la primera columna de candidates que exista en df."""
    for c in candidates:
        if c and c in df.columns: return c
    return default

def toN(s):
    try: return float(str(s).replace(",","."))
    except: return np.nan

def fmt_m(v):
    """Formato millones."""
    return f"${v/1e6:.1f}M" if abs(v) >= 1e6 else (f"${v/1e3:.0f}k" if abs(v) >= 1e3 else f"${v:.0f}")

def save_fig(out, name):
    p = os.path.join(out, f"{name}.png")
    plt.savefig(p, bbox_inches="tight", facecolor=BG)
    plt.close(); print(f"    ✓ {name}.png")

def bar_h(ax, labels, vals, colors=None, fmt=None):
    colors = colors or PAL[:len(labels)]
    bars = ax.barh(labels[::-1], vals[::-1], color=colors[::-1], alpha=0.88, edgecolor=BG)
    for b in bars:
        v = b.get_width()
        lbl = fmt(v) if fmt else f"{v:,.0f}"
        ax.text(v + abs(vals.max()*0.01 if hasattr(vals,'max') else 0.01),
                b.get_y()+b.get_height()/2, lbl, va="center", fontsize=8)
    ax.xaxis.grid(True); ax.set_axisbelow(True)
    return bars

def bar_v(ax, labels, vals, colors=None, fmt_fn=None, rotate=20):
    colors = colors or PAL[:len(labels)]
    bars = ax.bar(labels, vals, color=colors, alpha=0.88, edgecolor=BG, width=0.7)
    for b, v in zip(bars, vals):
        lbl = fmt_fn(v) if fmt_fn else f"{v:,.0f}"
        ax.text(b.get_x()+b.get_width()/2, b.get_height()+abs(v)*0.01, lbl,
                ha="center", fontsize=8, fontweight="bold")
    ax.yaxis.grid(True); ax.set_axisbelow(True)
    plt.setp(ax.get_xticklabels(), rotation=rotate, ha="right", fontsize=9)
    return bars

# ═══════════════════════════════════════════════════════════════════════════════
# 1. CARGA DE DATOS
# ═══════════════════════════════════════════════════════════════════════════════
SHEET_MAP = {
    "hosp":   ["HOSPITALIZACION","DX_TRIAGE","PRINCIPAL","EGRESO","ESTANCIA","MAIN","DATOS"],
    "cirugia":["CIRUGIA","CIRUGIA_HOSPITALARIA","QUIRURGICO","QX","CIRUGIA_AMB"],
    "uci":    ["UCI","CUIDADO_INTENSIVO","INTENSIVO","UTI"],
    "amb":    ["AMBULATORIO","CONSULTA_EXTERNA","CONSULTA","CONSULTAS","CE"],
    "cups":   ["CUPS","PROCEDIMIENTOS","PROCEDIMIENTO","CUPS_PROC"],
    "farma":  ["MEDICAMENTOS","FARMACIA","INSUMOS","MEDICAMENTO"],
    "factura":["FACTURACION","FACTURA","FACTURAS","FINANCIERO","BILLING"],
}

def cargar_hoja(xl, candidatos):
    sheets = xl.sheet_names
    for c in candidatos:
        for s in sheets:
            if c.upper() in s.upper():
                try:
                    df = pd.read_excel(xl, sheet_name=s)
                    print(f"      '{s}': {df.shape[0]:,} × {df.shape[1]}")
                    return df
                except Exception: pass
    return None

def cargar_datos(excel_path):
    print(f"\n[1] Cargando: {os.path.basename(excel_path)}")
    xl = pd.ExcelFile(excel_path)
    print(f"    Hojas encontradas: {xl.sheet_names}")
    data = {}
    for key, candidatos in SHEET_MAP.items():
        d = cargar_hoja(xl, candidatos)
        data[key] = d
        if d is not None: print(f"    [{key.upper()}] OK")
        else:             print(f"    [{key.upper()}] No encontrado — se omitirán sus gráficas")
    return data

# ═══════════════════════════════════════════════════════════════════════════════
# 2. INGENIERÍA DE VARIABLES
# ═══════════════════════════════════════════════════════════════════════════════
def preparar_hosp(df):
    """Normaliza columnas, calcula variables derivadas en hoja principal."""
    if df is None: return None
    df = df.copy()

    # Alias de columnas
    rename = {}
    for src, dsts in [
        ("INGRESO",         ["INGRESO","ID_INGRESO","NUM_INGRESO","ID","LLAVE"]),
        ("EDAD",            ["EDAD","EDAD_ANOS","AGE"]),
        ("SEXO",            ["SEXO","GENERO","GENDER"]),
        ("REGIMEN",         ["REGIMEN","REGIMEN_AFILIACION","TIPO_AFILIACION"]),
        ("EPS",             ["EAPB_NOMBRE","EPS","EPS_NOMBRE","ASEGURADORA","ENTIDAD"]),
        ("MUNICIPIO",       ["MUNICIPIO","MUN_RESIDENCIA","MUNICIPIO_RESIDENCIA","CIUDAD"]),
        ("ZONA",            ["ZONA","AREA_PROCEDENCIA","AREA","RURALIDAD"]),
        ("VIA_INGRESO",     ["VIA_INGRESO","TIPO_INGRESO","CAUSA_INGRESO"]),
        ("COLOR",           ["COLOR","COLOR_TRIAGE","TRIAGE_COLOR","CLASIFICACION_TRIAGE"]),
        ("COD_DX",          ["COD_CIE10","COD_CIE_10","COD_CIE10_HOSPITALIZACION","CIE10"]),
        ("NOM_DX",          ["NOM_DIAGNOSTICO","NOM_DX","DIAGNOSTICO","DX_PRINCIPAL","DESCRIPCION_DX"]),
        ("ESPECIALIDAD",    ["ESPECIALIDAD_TRATANTE","ESPECIALIDAD","NOMBRE_ESPECIALIDAD","ESPECIALIDAD_MEDICA"]),
        ("SERVICIO",        ["SERVICIO_HOSPITALARIO","SERVICIO","COD_SERVICIO","SALA"]),
        ("FECHA_ING",       ["FECHA_INGRESO","FECHA_INICIO","FECHA_ADM"]),
        ("FECHA_EGR",       ["FECHA_EGRESO","FECHA_FIN","FECHA_ALTA","FECHA_SALIDA"]),
        ("ESTANCIA",        ["DIAS_ESTANCIA_REDONDEADOS","DIAS_ESTANCIA","ESTANCIA_DIAS","LOS"]),
        ("MORTALIDAD",      ["MORTALIDAD","MUERTE","OBITO","FALLECIMIENTO"]),
        ("REINGRESO",       ["REINGRESO","REINGRESO_30_DIAS","READMISION"]),
        ("DESTINO",         ["DESTINO_EGRESO","DESTINO_FINAL","CONDICION_EGRESO","DESTINO"]),
        ("VALOR_FAC",       ["VALOR_FACTURADO","VALOR_TOTAL_FACTURA","VALOR_TOTAL","VALOR"]),
        ("COSTO",           ["COSTO_ESTIMADO","COSTO_TOTAL","COSTO","COSTOS"]),
        ("UTILIDAD",        ["UTILIDAD_NETA","UTILIDAD","MARGEN_UTILIDAD","GANANCIA"]),
        ("TAS",             ["TA_SISTOLICA","TENSION_SISTOLICA","T_SISTOLICA"]),
        ("TAD",             ["TA_DIASTOLICA","TENSION_DIASTOLICA","T_DIASTOLICA"]),
        ("FC",              ["FRECUENCIA_CARDIACA","F_CARDIACA","FC_CARDIACA"]),
        ("FR",              ["FRECUENCIA_RESPIRATORIA","F_RESPIRATORIA","FR"]),
        ("TEMP",            ["TEMPERATURA","TEMP","TEMPERATURA_C"]),
        ("SPO2",            ["SATUROMETRIA","SPO2","SATURACION_O2","OXIMETRIA"]),
        ("PESO",            ["PESO","PESO_KG"]),
        ("TALLA",           ["TALLA","TALLA_CM","ESTATURA"]),
    ]:
        c = fc(df, dsts)
        if c and c != src:
            rename[c] = src
    df.rename(columns=rename, inplace=True)

    # Fechas
    for col in ["FECHA_ING","FECHA_EGR","FECHA_TRIAGE","FECHA_INICIAL_HOSP","FECHA_SALIDA_HOSP"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Numéricos
    for col in ["EDAD","ESTANCIA","TAS","TAD","FC","FR","TEMP","SPO2","PESO","TALLA",
                "VALOR_FAC","COSTO","UTILIDAD"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Variables calculadas
    if "FECHA_ING" in df.columns:
        df["AÑO"]       = df["FECHA_ING"].dt.year
        df["MES"]       = df["FECHA_ING"].dt.month
        df["TRIMESTRE"] = df["FECHA_ING"].dt.quarter
        df["SEMANA"]    = df["FECHA_ING"].dt.isocalendar().week.astype(int)
        df["DIA"]       = df["FECHA_ING"].dt.day
        df["HORA"]      = df["FECHA_ING"].dt.hour
        df["DIA_SEMANA"]= df["FECHA_ING"].dt.day_name()

    if "TAS" in df.columns and "TAD" in df.columns:
        df["PAM"]         = ((df["TAS"] + 2*df["TAD"]) / 3).round(1)
        df["INDICE_SHOCK"]= (df["FC"] / df["TAS"].replace(0, np.nan)).round(3) if "FC" in df.columns else np.nan

    if "PESO" in df.columns and "TALLA" in df.columns:
        df["IMC"] = (df["PESO"] / (df["TALLA"]/100)**2).round(1)

    # Rentabilidades
    if "VALOR_FAC" in df.columns and "COSTO" in df.columns:
        df["UTILIDAD_CALC"]   = df["VALOR_FAC"] - df["COSTO"]
        df["MARGEN_PCT"]      = (df["UTILIDAD_CALC"] / df["VALOR_FAC"].replace(0,np.nan) * 100).round(2)
    if "UTILIDAD" not in df.columns and "UTILIDAD_CALC" in df.columns:
        df["UTILIDAD"] = df["UTILIDAD_CALC"]

    # Grupo etario
    if "EDAD" in df.columns:
        df["GRUPO_ETARIO"] = pd.cut(df["EDAD"],
            bins=[-1,0,5,12,17,29,44,59,74,200],
            labels=["Neonatal","Preescolar","Escolar","Adolescente","Adulto Joven","Adulto","Adulto Mayor","Vejez","Longevos"])

    # Score de complejidad clínica (proxy)
    cols_score = []
    if "ESTANCIA" in df.columns:   cols_score.append(("ESTANCIA",   0.4))
    if "FC" in df.columns:         cols_score.append(("FC",         0.15))
    if "SPO2" in df.columns:       cols_score.append(("SPO2",       0.1))
    if "TEMP" in df.columns:       cols_score.append(("TEMP",       0.05))
    if "INDICE_SHOCK" in df.columns: cols_score.append(("INDICE_SHOCK",0.1))
    if "VALOR_FAC" in df.columns:  cols_score.append(("VALOR_FAC",  0.2))
    if cols_score:
        sc = StandardScaler()
        arr = sc.fit_transform(df[[c for c,_ in cols_score]].fillna(0))
        df["SCORE_COMPLEJIDAD"] = np.clip(
            (arr * np.array([w for _,w in cols_score])).sum(axis=1) * 10 + 50, 0, 100).round(1)

    # Binario mortalidad / reingreso
    for col in ["MORTALIDAD","REINGRESO"]:
        if col in df.columns:
            df[col+"_BIN"] = df[col].astype(str).str.upper().isin(
                ["SI","SÍ","YES","1","TRUE","FALLECIDO","MUERTO","REINGRESO"]).astype(int)

    print(f"    Registros preparados: {len(df):,}")
    return df

# ═══════════════════════════════════════════════════════════════════════════════
# 3. CALIDAD DE DATOS
# ═══════════════════════════════════════════════════════════════════════════════
def reporte_calidad(df, out):
    if df is None: return {}
    print("\n[2] Calidad de datos...")
    miss = (df.isnull().sum() / len(df) * 100).sort_values(ascending=False)
    miss = miss[miss > 0].head(20)
    stats = {"total_registros": len(df), "total_columnas": len(df.columns),
             "pct_nulos": round(df.isnull().sum().sum() / (len(df)*len(df.columns))*100, 2),
             "columnas_nulos": miss.to_dict()}

    if len(miss) > 0:
        fig, ax = plt.subplots(figsize=(10,5))
        ax.barh(miss.index[::-1], miss.values[::-1], color=[PAL[i%len(PAL)] for i in range(len(miss))], alpha=0.85, edgecolor=BG)
        ax.axvline(20, color="#FF3B3B", linestyle="--", lw=1.5, label="20% umbral")
        for b in ax.patches:
            ax.text(b.get_width()+0.3, b.get_y()+b.get_height()/2, f"{b.get_width():.1f}%", va="center", fontsize=8)
        ax.set_title("Porcentaje de Datos Faltantes por Columna", fontsize=13, fontweight="bold")
        ax.set_xlabel("% Nulos"); ax.xaxis.grid(True); ax.set_axisbelow(True)
        ax.legend(framealpha=0.2); plt.tight_layout()
        save_fig(out, "F00_calidad_datos")
    return stats

# ═══════════════════════════════════════════════════════════════════════════════
# 4. ANÁLISIS HOSPITALIZACIÓN
# ═══════════════════════════════════════════════════════════════════════════════
def analisis_hospitalizacion(df, out):
    if df is None: print("\n[3] Hospitalizacion: sin datos"); return {}
    print("\n[3] Análisis de Hospitalización...")
    res = {}

    # F01 — Distribución estancia
    if "ESTANCIA" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(12,5))
        dp = df[df["ESTANCIA"].between(0,30)]["ESTANCIA"]
        n_, _, patches_ = axes[0].hist(dp, bins=31, range=(0,31), color=ACC, edgecolor=BG, alpha=0.85)
        for i,p in enumerate(patches_): p.set_facecolor(plt.cm.cool(i/31))
        axes[0].axvline(dp.mean(), color="#FFD600", linestyle="--", lw=2, label=f"Media: {dp.mean():.1f}d")
        axes[0].axvline(dp.median(), color="#FF7A00", linestyle=":", lw=2, label=f"Mediana: {dp.median():.0f}d")
        axes[0].set_title("Distribución de Estancia (≤30d)", fontsize=12, fontweight="bold")
        axes[0].set_xlabel("Días"); axes[0].legend(framealpha=0.2); axes[0].yaxis.grid(True)
        # Boxplot por grupo etario
        if "GRUPO_ETARIO" in df.columns:
            grupos = [df[df["GRUPO_ETARIO"]==g]["ESTANCIA"].dropna().values
                      for g in df["GRUPO_ETARIO"].cat.categories if len(df[df["GRUPO_ETARIO"]==g])>5]
            etiq = [str(g) for g in df["GRUPO_ETARIO"].cat.categories if len(df[df["GRUPO_ETARIO"]==g])>5]
            if grupos:
                bp = axes[1].boxplot(grupos, patch_artist=True, notch=False,
                    medianprops=dict(color="#FFD600",lw=2),
                    boxprops=dict(alpha=0.7), whiskerprops=dict(color=TEXT),
                    capprops=dict(color=TEXT), flierprops=dict(marker="o",markersize=2,alpha=0.3,markerfacecolor=ACC))
                for patch, col in zip(bp["boxes"], PAL): patch.set_facecolor(col)
                axes[1].set_xticklabels(etiq, rotation=30, ha="right", fontsize=8)
        axes[1].set_title("Estancia por Grupo Etario", fontsize=12, fontweight="bold")
        axes[1].yaxis.grid(True); axes[1].set_axisbelow(True)
        plt.suptitle("Análisis de Estancia Hospitalaria", fontsize=14, fontweight="bold", y=1.02)
        plt.tight_layout(); save_fig(out,"F01_estancia")
        res["estancia_media"] = round(float(df["ESTANCIA"].mean()),2)
        res["estancia_mediana"] = round(float(df["ESTANCIA"].median()),2)

    # F02 — Top diagnósticos y estancia por Dx
    if "NOM_DX" in df.columns and "COD_DX" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(14,6))
        td = df["NOM_DX"].value_counts().head(10)
        ns = [n[:45]+"…" if len(n)>45 else n for n in td.index]
        bar_h(axes[0], ns, td.values, PAL[:10])
        axes[0].set_title("Top 10 Diagnósticos por Frecuencia", fontsize=12, fontweight="bold")
        if "ESTANCIA" in df.columns:
            ea = df.groupby("NOM_DX")["ESTANCIA"].mean().sort_values(ascending=False).head(10)
            ns2 = [n[:45]+"…" if len(n)>45 else n for n in ea.index]
            bar_h(axes[1], ns2, ea.values, PAL[::-1][:10], fmt=lambda v:f"{v:.1f}d")
            axes[1].set_title("Top 10 Diagnósticos por Estancia Media", fontsize=12, fontweight="bold")
        plt.suptitle("Diagnósticos Principales", fontsize=14, fontweight="bold", y=1.01)
        plt.tight_layout(); save_fig(out,"F02_diagnosticos")

    # F03 — Especialidades
    if "ESPECIALIDAD" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(14,6))
        sp = df["ESPECIALIDAD"].value_counts().head(12)
        ns = [n[:38]+"…" if len(n)>38 else n for n in sp.index]
        bar_h(axes[0], ns, sp.values, PAL[:12])
        axes[0].set_title("Ingresos por Especialidad", fontsize=12, fontweight="bold")
        if "ESTANCIA" in df.columns:
            se = df.groupby("ESPECIALIDAD")["ESTANCIA"].mean().sort_values(ascending=False).head(12)
            ns2 = [n[:38]+"…" if len(n)>38 else n for n in se.index]
            bar_h(axes[1], ns2, se.values, PAL[::-1][:12], fmt=lambda v:f"{v:.1f}d")
            axes[1].set_title("Estancia Media por Especialidad", fontsize=12, fontweight="bold")
        plt.suptitle("Análisis por Especialidad", fontsize=14, fontweight="bold", y=1.01)
        plt.tight_layout(); save_fig(out,"F03_especialidades")

    # F04 — EPS / EAPB
    if "EPS" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(14,6))
        ep = df["EPS"].value_counts().head(10)
        ns = [n[:32]+"…" if len(n)>32 else n for n in ep.index]
        bar_h(axes[0], ns, ep.values, PAL[:10])
        axes[0].set_title("Top 10 EPS por Volumen", fontsize=12, fontweight="bold")
        if "ESTANCIA" in df.columns:
            ee = df.groupby("EPS")["ESTANCIA"].mean().sort_values(ascending=False).head(10)
            ns2 = [n[:32]+"…" if len(n)>32 else n for n in ee.index]
            bar_h(axes[1], ns2, ee.values, PAL[::-1][:10], fmt=lambda v:f"{v:.1f}d")
            axes[1].set_title("Estancia Media por EPS", fontsize=12, fontweight="bold")
        plt.suptitle("Comportamiento por EPS / EAPB", fontsize=14, fontweight="bold", y=1.01)
        plt.tight_layout(); save_fig(out,"F04_eps")

    # F05 — Demografía (sexo, zona, régimen)
    fig, axes = plt.subplots(1,3,figsize=(14,5))
    for ax, col, title in zip(axes, ["SEXO","ZONA","REGIMEN"], ["Sexo","Zona","Régimen"]):
        if col in df.columns:
            vc = df[col].value_counts().head(6)
            ax.pie(vc.values, labels=[str(x)[:20] for x in vc.index],
                   colors=PAL[:len(vc)], autopct="%1.1f%%", pctdistance=0.78,
                   wedgeprops=dict(width=0.55,edgecolor=BG,linewidth=2))
        ax.set_title(title, fontsize=11, fontweight="bold")
    plt.suptitle("Distribución Demográfica", fontsize=14, fontweight="bold", y=1.01)
    plt.tight_layout(); save_fig(out,"F05_demografia")

    # F06 — Municipios
    if "MUNICIPIO" in df.columns:
        fig, ax = plt.subplots(figsize=(10,6))
        tm = df["MUNICIPIO"].value_counts().head(15)
        ns = [n[:35]+"…" if len(n)>35 else n for n in tm.index]
        bar_h(ax, ns, tm.values, PAL[:15])
        ax.set_title("Top 15 Municipios de Procedencia", fontsize=13, fontweight="bold")
        plt.tight_layout(); save_fig(out,"F06_municipios")

    # F07 — Mortalidad y reingresos
    fig, axes = plt.subplots(1,2,figsize=(12,5))
    for ax, col, title in zip(axes, ["MORTALIDAD_BIN","REINGRESO_BIN"],
                              ["Mortalidad","Reingreso 30 días"]):
        if col in df.columns:
            vc = df[col].value_counts()
            colors = ["#FF3B3B" if x==1 else "#00C875" for x in vc.index]
            lbls = ["Sí" if x==1 else "No" for x in vc.index]
            ax.pie(vc.values, labels=lbls, colors=colors, autopct="%1.1f%%",
                   pctdistance=0.78, wedgeprops=dict(width=0.55,edgecolor=BG,linewidth=2))
        ax.set_title(title, fontsize=11, fontweight="bold")
    plt.suptitle("Indicadores de Calidad Clínica", fontsize=14, fontweight="bold", y=1.01)
    plt.tight_layout(); save_fig(out,"F07_calidad_clinica")

    # F08 — Tendencia temporal ingresos
    if "AÑO" in df.columns and "MES" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(14,5))
        iv = df.groupby("AÑO").size().reset_index(name="n")
        axes[0].plot(iv["AÑO"], iv["n"], "o-", color=ACC, lw=2.5, markersize=9)
        axes[0].fill_between(iv["AÑO"], iv["n"], alpha=0.12, color=ACC)
        for x, y in zip(iv["AÑO"], iv["n"]): axes[0].text(x, y+max(iv["n"])*0.01, str(y), ha="center", fontsize=9, fontweight="bold")
        axes[0].set_title("Ingresos por Año", fontsize=12, fontweight="bold"); axes[0].yaxis.grid(True)
        meses = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
        mv = df.groupby("MES").size().reset_index(name="n").sort_values("MES")
        axes[1].bar([meses.get(m,str(m)) for m in mv["MES"]], mv["n"],
                    color=PAL[:len(mv)], alpha=0.88, edgecolor=BG)
        axes[1].set_title("Ingresos por Mes", fontsize=12, fontweight="bold"); axes[1].yaxis.grid(True)
        plt.setp(axes[1].get_xticklabels(), rotation=30, ha="right")
        plt.suptitle("Tendencias Temporales", fontsize=14, fontweight="bold", y=1.01)
        plt.tight_layout(); save_fig(out,"F08_tendencias")

    # F09 — Destino de egreso
    if "DESTINO" in df.columns:
        fig, ax = plt.subplots(figsize=(9,5))
        td = df["DESTINO"].value_counts().head(8)
        bar_h(ax, [str(n)[:35] for n in td.index], td.values, PAL[:8])
        ax.set_title("Destino de Egreso", fontsize=13, fontweight="bold")
        plt.tight_layout(); save_fig(out,"F09_destino_egreso")

    # F10 — Servicios / salas
    if "SERVICIO" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(14,6))
        sv = df["SERVICIO"].value_counts().head(12)
        ns = [n[:35]+"…" if len(n)>35 else n for n in sv.index]
        bar_h(axes[0], ns, sv.values, PAL[:12])
        axes[0].set_title("Ingresos por Servicio", fontsize=12, fontweight="bold")
        if "ESTANCIA" in df.columns:
            se = df.groupby("SERVICIO")["ESTANCIA"].mean().sort_values(ascending=False).head(12)
            ns2 = [n[:35]+"…" if len(n)>35 else n for n in se.index]
            bar_h(axes[1], ns2, se.values, PAL[::-1][:12], fmt=lambda v:f"{v:.1f}d")
            axes[1].set_title("Estancia Media por Servicio", fontsize=12, fontweight="bold")
        plt.suptitle("Análisis por Servicio Hospitalario", fontsize=14, fontweight="bold", y=1.01)
        plt.tight_layout(); save_fig(out,"F10_servicios")

    print(f"    ✓ 10 figuras de hospitalización generadas")
    return res

# ═══════════════════════════════════════════════════════════════════════════════
# 5. ANÁLISIS CIRUGÍA
# ═══════════════════════════════════════════════════════════════════════════════
def analisis_cirugia(df_cir, out):
    if df_cir is None: print("\n[4] Cirugía: sin datos"); return {}
    print("\n[4] Análisis de Cirugía...")
    df = df_cir.copy()
    # Alias
    rn = {}
    for src, cands in [
        ("TIPO_CIRUGIA",    ["TIPO_CIRUGIA","TIPO_QX","TIPO_INTERVENCION"]),
        ("ESPECIALIDAD_QX", ["ESPECIALIDAD_QX","ESPECIALIDAD_CIRUGIA","ESPECIALIDAD"]),
        ("CUPS_QX",         ["CUPS_CIRUGIA","CUPS_QX","COD_CUPS","CUPS"]),
        ("DESC_CUPS",       ["DESCRIPCION_CUPS_CIRUGIA","DESCRIPCION_CUPS","NOMBRE_CUPS","DESC_QX"]),
        ("TIEMPO_QX",       ["TIEMPO_QUIRURGICO","TIEMPO_QX","DURACION_QX","TIEMPO_CIRUGIA"]),
        ("ESTANCIA_POST",   ["ESTANCIA_POST_QX","DIAS_POST_QX","ESTANCIA_POSQX"]),
        ("COMPLICACION",    ["COMPLICACION_QX","COMPLICACION","COMPLICACIONES"]),
        ("REINTERVENCION",  ["REINTERVENCION","REOPERACION","REOPERATION"]),
        ("VALOR_QX",        ["VALOR_FACTURADO","VALOR_TOTAL","VALOR"]),
        ("COSTO_QX",        ["COSTO","COSTO_TOTAL","COSTO_CIRUGIA"]),
        ("UTILIDAD_QX",     ["UTILIDAD","UTILIDAD_NETA","GANANCIA"]),
        ("FECHA_QX",        ["FECHA_CIRUGIA","FECHA_QX","FECHA_INTERVENCION"]),
        ("CIRUJANO",        ["CIRUJANO","MEDICO_QX","MEDICO_CIRUJANO"]),
    ]:
        c = fc(df, cands)
        if c and c != src: rn[c] = src
    df.rename(columns=rn, inplace=True)
    for col in ["VALOR_QX","COSTO_QX","UTILIDAD_QX","TIEMPO_QX","ESTANCIA_POST"]:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce")
    if "VALOR_QX" in df.columns and "COSTO_QX" in df.columns:
        df["UTILIDAD_QX_CALC"] = df["VALOR_QX"] - df["COSTO_QX"]
        df["MARGEN_QX"] = (df["UTILIDAD_QX_CALC"] / df["VALOR_QX"].replace(0,np.nan)*100).round(2)

    # F11 — Top CUPS quirúrgicos y especialidades
    fig, axes = plt.subplots(1,2,figsize=(14,6))
    if "DESC_CUPS" in df.columns:
        tc = df["DESC_CUPS"].value_counts().head(10)
        ns = [n[:42]+"…" if len(n)>42 else n for n in tc.index]
        bar_h(axes[0], ns, tc.values, PAL[:10])
        axes[0].set_title("Top 10 Procedimientos Quirúrgicos", fontsize=11, fontweight="bold")
    if "ESPECIALIDAD_QX" in df.columns:
        se = df["ESPECIALIDAD_QX"].value_counts().head(10)
        ns2 = [n[:38]+"…" if len(n)>38 else n for n in se.index]
        bar_h(axes[1], ns2, se.values, PAL[::-1][:10])
        axes[1].set_title("Cirugías por Especialidad", fontsize=11, fontweight="bold")
    plt.suptitle("Análisis de Cirugías", fontsize=14, fontweight="bold", y=1.01)
    plt.tight_layout(); save_fig(out,"F11_cirugia_top")

    # F12 — Tipo (ambulatoria vs hospitalaria) y complicaciones
    fig, axes = plt.subplots(1,2,figsize=(12,5))
    if "TIPO_CIRUGIA" in df.columns:
        tc = df["TIPO_CIRUGIA"].value_counts()
        axes[0].pie(tc.values, labels=[str(x)[:20] for x in tc.index], colors=PAL[:len(tc)],
                    autopct="%1.1f%%", pctdistance=0.78, wedgeprops=dict(width=0.55,edgecolor=BG,linewidth=2))
        axes[0].set_title("Tipo de Cirugía", fontsize=11, fontweight="bold")
    if "COMPLICACION" in df.columns:
        df["COMP_BIN"] = df["COMPLICACION"].astype(str).str.upper().isin(["SI","SÍ","YES","1","TRUE"]).astype(int)
        vc = df["COMP_BIN"].value_counts()
        axes[1].bar(["Sin complicación","Con complicación"], [vc.get(0,0),vc.get(1,0)],
                    color=["#00C875","#FF3B3B"], alpha=0.88, edgecolor=BG)
        axes[1].set_title("Complicaciones Quirúrgicas", fontsize=11, fontweight="bold")
        axes[1].yaxis.grid(True); axes[1].set_axisbelow(True)
    plt.suptitle("Calidad e Indicadores Quirúrgicos", fontsize=14, fontweight="bold", y=1.01)
    plt.tight_layout(); save_fig(out,"F12_cirugia_calidad")

    # F13 — Rentabilidad quirúrgica
    if "ESPECIALIDAD_QX" in df.columns and "UTILIDAD_QX_CALC" in df.columns:
        fig, ax = plt.subplots(figsize=(10,6))
        ru = df.groupby("ESPECIALIDAD_QX")["UTILIDAD_QX_CALC"].mean().sort_values(ascending=False).head(10)
        ns = [n[:35]+"…" if len(n)>35 else n for n in ru.index]
        colors = ["#00C875" if v>=0 else "#FF3B3B" for v in ru.values]
        bar_h(ax, ns, ru.values, colors, fmt=fmt_m)
        ax.axvline(0, color="white", lw=1.2)
        ax.set_title("Utilidad Media por Especialidad Quirúrgica", fontsize=13, fontweight="bold")
        plt.tight_layout(); save_fig(out,"F13_cirugia_rentabilidad")

    print("    ✓ 3 figuras de cirugía generadas")
    return {}

# ═══════════════════════════════════════════════════════════════════════════════
# 6. ANÁLISIS UCI
# ═══════════════════════════════════════════════════════════════════════════════
def analisis_uci(df_uci, out):
    if df_uci is None: print("\n[5] UCI: sin datos"); return {}
    print("\n[5] Análisis UCI...")
    df = df_uci.copy()
    rn = {}
    for src, cands in [
        ("TIPO_UCI",      ["TIPO_UCI","TIPO_CUIDADO","SERVICIO_UCI"]),
        ("DIAS_UCI",      ["DIAS_UCI","ESTANCIA_UCI","DIAS_CUIDADO_INTENSIVO"]),
        ("MORTALIDAD_UCI",["MORTALIDAD_UCI","MORTALIDAD","MUERTE_UCI"]),
        ("VM",            ["VENTILACION_MECANICA","VENTILACION","VM"]),
        ("VASOPRESOR",    ["SOPORTE_VASOPRESOR","VASOPRESOR","VASOPRESORES"]),
        ("VALOR_UCI",     ["VALOR_FACTURADO_UCI","VALOR_UCI","VALOR"]),
        ("COSTO_UCI",     ["COSTO_DIA_UCI","COSTO_UCI","COSTO"]),
        ("DX_UCI",        ["DX_PRINCIPAL_UCI","DIAGNOSTICO_UCI","DX_UCI"]),
        ("EPS",           ["EPS","EAPB_NOMBRE","ASEGURADORA"]),
    ]:
        c = fc(df, cands)
        if c and c != src: rn[c] = src
    df.rename(columns=rn, inplace=True)
    for col in ["DIAS_UCI","VALOR_UCI","COSTO_UCI"]:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce")

    # F14 — UCI tipo y días
    fig, axes = plt.subplots(1,2,figsize=(12,5))
    if "TIPO_UCI" in df.columns:
        tv = df["TIPO_UCI"].value_counts()
        axes[0].pie(tv.values, labels=[str(x)[:22] for x in tv.index], colors=PAL[:len(tv)],
                    autopct="%1.1f%%", pctdistance=0.78, wedgeprops=dict(width=0.55,edgecolor=BG,linewidth=2))
        axes[0].set_title("Distribución UCI por Tipo", fontsize=11, fontweight="bold")
    if "DIAS_UCI" in df.columns:
        dp = df["DIAS_UCI"].dropna()
        axes[1].hist(dp[dp<=30], bins=31, range=(0,31), color="#F15BB5", edgecolor=BG, alpha=0.85)
        axes[1].axvline(dp.mean(), color="#FFD600", linestyle="--", lw=2, label=f"Media: {dp.mean():.1f}d")
        axes[1].set_title("Distribución Estancia UCI", fontsize=11, fontweight="bold")
        axes[1].legend(framealpha=0.2); axes[1].yaxis.grid(True)
    plt.suptitle("Análisis UCI", fontsize=14, fontweight="bold", y=1.01)
    plt.tight_layout(); save_fig(out,"F14_uci_general")

    # F15 — Mortalidad UCI + VM + vasopresores
    fig, axes = plt.subplots(1,3,figsize=(14,5))
    for ax, col, title in zip(axes,
        ["MORTALIDAD_UCI","VM","VASOPRESOR"],
        ["Mortalidad UCI","Ventilación Mecánica","Soporte Vasopresor"]):
        if col in df.columns:
            bin_vals = df[col].astype(str).str.upper().isin(["SI","SÍ","YES","1","TRUE"]).astype(int)
            vc = bin_vals.value_counts()
            lbls = ["Sí" if x==1 else "No" for x in vc.index]
            colors = ["#FF3B3B","#00C875"] if vc.index[0]==1 else ["#00C875","#FF3B3B"]
            ax.pie(vc.values, labels=lbls, colors=colors, autopct="%1.1f%%",
                   pctdistance=0.78, wedgeprops=dict(width=0.55,edgecolor=BG,linewidth=2))
        ax.set_title(title, fontsize=10, fontweight="bold")
    plt.suptitle("Indicadores Críticos UCI", fontsize=14, fontweight="bold", y=1.01)
    plt.tight_layout(); save_fig(out,"F15_uci_indicadores")

    print("    ✓ 2 figuras UCI generadas")
    return {}

# ═══════════════════════════════════════════════════════════════════════════════
# 7. ANÁLISIS AMBULATORIO
# ═══════════════════════════════════════════════════════════════════════════════
def analisis_ambulatorio(df_amb, out):
    if df_amb is None: print("\n[6] Ambulatorio: sin datos"); return {}
    print("\n[6] Análisis Ambulatorio...")
    df = df_amb.copy()
    rn = {}
    for src, cands in [
        ("ESPECIALIDAD", ["ESPECIALIDAD_AMBULATORIA","ESPECIALIDAD","ESPECIALIDAD_MEDICA"]),
        ("TIPO_CONSULTA",["PRIMERA_VEZ_O_CONTROL","TIPO_CONSULTA","PRIMERA_VEZ"]),
        ("DX_AMB",       ["DX_CONSULTA","DIAGNOSTICO","COD_CIE10"]),
        ("EPS",          ["EPS","EAPB_NOMBRE","ASEGURADORA"]),
        ("SEXO",         ["SEXO","GENERO"]),
        ("EDAD",         ["EDAD","EDAD_ANOS"]),
        ("MUNICIPIO",    ["MUNICIPIO","MUNICIPIO_RESIDENCIA"]),
        ("VALOR",        ["VALOR_CONSULTA","VALOR_TOTAL","VALOR"]),
        ("COSTO",        ["COSTO_CONSULTA","COSTO_TOTAL","COSTO"]),
        ("UTILIDAD",     ["UTILIDAD_CONSULTA","UTILIDAD","MARGEN"]),
        ("FECHA",        ["FECHA_CONSULTA","FECHA_ATENCION","FECHA"]),
    ]:
        c = fc(df, cands)
        if c and c != src: rn[c] = src
    df.rename(columns=rn, inplace=True)
    for col in ["EDAD","VALOR","COSTO","UTILIDAD"]:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce")
    if "FECHA" in df.columns: df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")

    # F16 — Especialidades ambulatorias
    fig, axes = plt.subplots(1,2,figsize=(14,6))
    if "ESPECIALIDAD" in df.columns:
        sp = df["ESPECIALIDAD"].value_counts().head(12)
        ns = [n[:38]+"…" if len(n)>38 else n for n in sp.index]
        bar_h(axes[0], ns, sp.values, PAL[:12])
        axes[0].set_title("Consultas por Especialidad", fontsize=11, fontweight="bold")
    if "TIPO_CONSULTA" in df.columns:
        tc = df["TIPO_CONSULTA"].value_counts().head(4)
        axes[1].bar([str(x)[:22] for x in tc.index], tc.values, color=PAL[:len(tc)], alpha=0.88, edgecolor=BG)
        axes[1].set_title("Tipo de Consulta (1ª vez vs Control)", fontsize=11, fontweight="bold")
        axes[1].yaxis.grid(True); axes[1].set_axisbelow(True)
    plt.suptitle("Consulta Externa y Ambulatorio", fontsize=14, fontweight="bold", y=1.01)
    plt.tight_layout(); save_fig(out,"F16_ambulatorio")

    # F17 — EPS ambulatorio y rentabilidad
    if "EPS" in df.columns and "UTILIDAD" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(14,6))
        ev = df["EPS"].value_counts().head(10)
        ns = [n[:30]+"…" if len(n)>30 else n for n in ev.index]
        bar_h(axes[0], ns, ev.values, PAL[:10])
        axes[0].set_title("Consultas por EPS", fontsize=11, fontweight="bold")
        eu = df.groupby("EPS")["UTILIDAD"].mean().sort_values(ascending=False).head(10)
        ns2 = [n[:30]+"…" if len(n)>30 else n for n in eu.index]
        colors = ["#00C875" if v>=0 else "#FF3B3B" for v in eu.values]
        bar_h(axes[1], ns2, eu.values, colors, fmt=fmt_m)
        axes[1].axvline(0,color="white",lw=1)
        axes[1].set_title("Utilidad Media por EPS (Ambulatorio)", fontsize=11, fontweight="bold")
        plt.suptitle("EPS y Rentabilidad Ambulatoria", fontsize=14, fontweight="bold", y=1.01)
        plt.tight_layout(); save_fig(out,"F17_amb_eps_rentabilidad")

    print("    ✓ 2 figuras ambulatorio generadas")
    return {}

# ═══════════════════════════════════════════════════════════════════════════════
# 8. ANÁLISIS FINANCIERO / COSTOS
# ═══════════════════════════════════════════════════════════════════════════════
def analisis_financiero(df, cups_df, out):
    print("\n[7] Análisis Financiero...")
    res = {}

    if df is not None and "VALOR_FAC" in df.columns:
        # F18 — KPIs financieros
        total_val = df["VALOR_FAC"].sum()
        total_cos = df["COSTO"].sum() if "COSTO" in df.columns else 0
        total_util = df["UTILIDAD"].sum() if "UTILIDAD" in df.columns else (total_val - total_cos)
        margen = (total_util/total_val*100) if total_val else 0
        res.update({"total_facturado":round(float(total_val),0),
                    "total_costo":round(float(total_cos),0),
                    "total_utilidad":round(float(total_util),0),
                    "margen_global_pct":round(float(margen),2)})

        fig, axes = plt.subplots(1,3,figsize=(14,6))
        # Facturación por EPS
        if "EPS" in df.columns:
            ef = df.groupby("EPS")["VALOR_FAC"].sum().sort_values(ascending=False).head(10)
            ns = [n[:30]+"…" if len(n)>30 else n for n in ef.index]
            bar_h(axes[0], ns, ef.values, PAL[:10], fmt=fmt_m)
            axes[0].set_title("Facturación por EPS", fontsize=11, fontweight="bold")
        # Facturación vs costo (donut comparativo)
        if "COSTO" in df.columns:
            cats = ["Facturado","Costo","Utilidad"]
            vals = [max(total_val,0), max(total_cos,0), max(total_util,0)]
            axes[1].bar(cats, vals, color=["#00F5D4","#FF5FAA","#FFE040"], alpha=0.88, edgecolor=BG)
            for bar, v in zip(axes[1].patches, vals):
                axes[1].text(bar.get_x()+bar.get_width()/2, bar.get_height()+max(vals)*0.01,
                             fmt_m(v), ha="center", fontsize=9, fontweight="bold")
            axes[1].set_title("Facturado vs Costo vs Utilidad", fontsize=11, fontweight="bold")
            axes[1].yaxis.grid(True); axes[1].set_axisbelow(True)
        # Margen por especialidad
        if "ESPECIALIDAD" in df.columns and "UTILIDAD" in df.columns:
            eu = df.groupby("ESPECIALIDAD")["UTILIDAD"].mean().sort_values(ascending=False).head(10)
            ns3 = [n[:30]+"…" if len(n)>30 else n for n in eu.index]
            colors = ["#00C875" if v>=0 else "#FF3B3B" for v in eu.values]
            bar_h(axes[2], ns3, eu.values, colors, fmt=fmt_m)
            axes[2].axvline(0, color="white", lw=1)
            axes[2].set_title("Utilidad Media por Especialidad", fontsize=11, fontweight="bold")
        plt.suptitle("Análisis Financiero Integral", fontsize=14, fontweight="bold", y=1.01)
        plt.tight_layout(); save_fig(out,"F18_financiero_general")

        # F19 — Ranking rentabilidad por DX
        if "NOM_DX" in df.columns and "VALOR_FAC" in df.columns:
            fig, axes = plt.subplots(1,2,figsize=(14,6))
            rv = df.groupby("NOM_DX")["VALOR_FAC"].sum().sort_values(ascending=False).head(10)
            ns = [n[:42]+"…" if len(n)>42 else n for n in rv.index]
            bar_h(axes[0], ns, rv.values, PAL[:10], fmt=fmt_m)
            axes[0].set_title("Facturación Total por Diagnóstico", fontsize=11, fontweight="bold")
            if "UTILIDAD" in df.columns:
                ru = df.groupby("NOM_DX")["UTILIDAD"].mean().sort_values(ascending=False).head(10)
                ns2 = [n[:42]+"…" if len(n)>42 else n for n in ru.index]
                colors = ["#00C875" if v>=0 else "#FF3B3B" for v in ru.values]
                bar_h(axes[1], ns2, ru.values, colors, fmt=fmt_m)
                axes[1].axvline(0, color="white", lw=1)
                axes[1].set_title("Utilidad Media por Diagnóstico", fontsize=11, fontweight="bold")
            plt.suptitle("Rentabilidad por Diagnóstico", fontsize=14, fontweight="bold", y=1.01)
            plt.tight_layout(); save_fig(out,"F19_rentabilidad_dx")

        # F20 — Tendencia financiera mensual
        if "AÑO" in df.columns and "MES" in df.columns:
            fig, ax = plt.subplots(figsize=(12,5))
            df["PERIODO"] = df["AÑO"].astype(str)+"-"+df["MES"].astype(str).str.zfill(2)
            gf = df.groupby("PERIODO")["VALOR_FAC"].sum().sort_index()
            ax.plot(range(len(gf)), gf.values, "o-", color=ACC, lw=2.5, markersize=7)
            ax.fill_between(range(len(gf)), gf.values, alpha=0.12, color=ACC)
            if "UTILIDAD" in df.columns:
                gu = df.groupby("PERIODO")["UTILIDAD"].sum().sort_index()
                ax2 = ax.twinx()
                ax2.plot(range(len(gu)), gu.values, "s--", color="#FFE040", lw=2, markersize=6, alpha=0.8)
                ax2.set_ylabel("Utilidad", color="#FFE040")
            ax.set_xticks(range(len(gf)))
            ax.set_xticklabels(gf.index.tolist(), rotation=45, ha="right", fontsize=8)
            ax.set_title("Tendencia Financiera Mensual", fontsize=13, fontweight="bold")
            ax.set_ylabel("Facturación"); ax.yaxis.grid(True)
            plt.tight_layout(); save_fig(out,"F20_tendencia_financiera")

    # F21 — CUPS rentabilidad
    if cups_df is not None:
        cups = cups_df.copy()
        rn = {}
        for src, cands in [
            ("COD_CUPS",  ["CUPS","COD_CUPS","CUPS_COD"]),
            ("NOM_CUPS",  ["NOMBRE_CUPS","DESCRIPCION_CUPS","DESC_CUPS"]),
            ("VALOR",     ["VALOR_TOTAL","VALOR_UNITARIO","VALOR"]),
            ("COSTO",     ["COSTO_PROCEDIMIENTO","COSTO","COSTOS"]),
            ("UTILIDAD",  ["UTILIDAD_PROCEDIMIENTO","UTILIDAD","GANANCIA"]),
        ]:
            c = fc(cups, cands)
            if c and c != src: rn[c] = src
        cups.rename(columns=rn, inplace=True)
        for col in ["VALOR","COSTO","UTILIDAD"]:
            if col in cups.columns: cups[col] = pd.to_numeric(cups[col], errors="coerce")
        if "UTILIDAD" not in cups.columns and "VALOR" in cups.columns and "COSTO" in cups.columns:
            cups["UTILIDAD"] = cups["VALOR"] - cups["COSTO"]

        if "NOM_CUPS" in cups.columns and "UTILIDAD" in cups.columns:
            fig, axes = plt.subplots(1,2,figsize=(14,6))
            top_vol = cups["NOM_CUPS"].value_counts().head(10)
            ns = [n[:40]+"…" if len(n)>40 else n for n in top_vol.index]
            bar_h(axes[0], ns, top_vol.values, PAL[:10])
            axes[0].set_title("Top CUPS por Frecuencia", fontsize=11, fontweight="bold")
            ru = cups.groupby("NOM_CUPS")["UTILIDAD"].mean().sort_values(ascending=False).head(10)
            ns2 = [n[:40]+"…" if len(n)>40 else n for n in ru.index]
            colors = ["#00C875" if v>=0 else "#FF3B3B" for v in ru.values]
            bar_h(axes[1], ns2, ru.values, colors, fmt=fmt_m)
            axes[1].axvline(0,color="white",lw=1)
            axes[1].set_title("Utilidad Media por CUPS", fontsize=11, fontweight="bold")
            plt.suptitle("Rentabilidad por Procedimiento CUPS", fontsize=14, fontweight="bold", y=1.01)
            plt.tight_layout(); save_fig(out,"F21_cups_rentabilidad")

    print("    ✓ Figuras financieras generadas")
    return res

# ═══════════════════════════════════════════════════════════════════════════════
# 9. MACHINE LEARNING
# ═══════════════════════════════════════════════════════════════════════════════
def analisis_ml(df, out):
    if df is None: print("\n[8] ML: sin datos"); return {}
    print("\n[8] Machine Learning...")
    res = {}

    # ── Variables numéricas disponibles ────────────────────────────────────────
    NUM_CANDS = ["EDAD","TAS","TAD","FC","FR","TEMP","SPO2","PAM","INDICE_SHOCK","IMC",
                 "VALOR_FAC","COSTO","UTILIDAD","SCORE_COMPLEJIDAD"]
    CAT_CANDS = ["SEXO","ZONA","REGIMEN","EPS","ESPECIALIDAD","SERVICIO","VIA_INGRESO","COLOR"]
    TARGET    = "ESTANCIA"

    num_feats = [c for c in NUM_CANDS if c in df.columns and c != TARGET]
    cat_feats = [c for c in CAT_CANDS if c in df.columns]

    if TARGET not in df.columns or len(df) < 50:
        print("    Datos insuficientes para ML"); return res

    df_ml = df.copy()
    encoders = {}
    for c in cat_feats:
        le = LabelEncoder()
        df_ml[c+"_enc"] = le.fit_transform(df_ml[c].astype(str))
        encoders[c] = {str(cl):int(i) for i,cl in enumerate(le.classes_)}

    all_feats = num_feats + [c+"_enc" for c in cat_feats]
    df_ml_clean = df_ml[all_feats + [TARGET]].dropna()
    if len(df_ml_clean) < 50: print("    Datos insuficientes tras dropna"); return res

    X = df_ml_clean[all_feats].values
    y = df_ml_clean[TARGET].values
    imp = SimpleImputer(strategy="median")
    X = imp.fit_transform(X)

    X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.2, random_state=42)

    # ─── Modelos ───────────────────────────────────────────────────────────────
    models = {
        "Regresión Lineal":     LinearRegression(),
        "Ridge":                Ridge(alpha=1.0),
        "Random Forest":        RandomForestRegressor(n_estimators=150,max_depth=8,random_state=42,n_jobs=-1),
        "Gradient Boosting":    GradientBoostingRegressor(n_estimators=100,learning_rate=0.1,max_depth=4,random_state=42),
        "Árbol de Decisión":    DecisionTreeRegressor(max_depth=6,random_state=42),
    }
    ml_res = {}
    for name, model in models.items():
        try:
            model.fit(X_tr, y_tr)
            yp = model.predict(X_te)
            cv = cross_val_score(model, X, y, cv=min(5,len(df_ml_clean)//20+2),
                                 scoring="neg_mean_absolute_error", n_jobs=-1)
            ml_res[name] = {
                "MAE":   round(float(mean_absolute_error(y_te,yp)),3),
                "RMSE":  round(float(np.sqrt(mean_squared_error(y_te,yp))),3),
                "R2":    round(float(r2_score(y_te,yp)),4),
                "CV_MAE":round(float(-cv.mean()),3),
                "CV_STD":round(float(cv.std()),3),
            }
            print(f"    {name}: MAE={ml_res[name]['MAE']}  R²={ml_res[name]['R2']}")
        except Exception as e:
            print(f"    Error {name}: {e}")

    res["modelos"] = ml_res
    mejor = min(ml_res, key=lambda k: ml_res[k]["MAE"]) if ml_res else None
    res["mejor_modelo"] = mejor

    # F22 — Comparación de métricas
    if ml_res:
        fig, axes = plt.subplots(1,3,figsize=(14,5))
        names_ml = list(ml_res.keys())
        for i,(met,tit) in enumerate(zip(["MAE","RMSE","R2"],["MAE (días)","RMSE","R²"])):
            vals = [ml_res[m][met] for m in names_ml]
            bars = axes[i].bar(range(len(names_ml)), vals,
                               color=[PAL[j%len(PAL)] for j in range(len(names_ml))],
                               alpha=0.88, edgecolor=BG)
            for b,v in zip(bars,vals):
                axes[i].text(b.get_x()+b.get_width()/2, b.get_height()+max(vals)*0.01,
                             f"{v:.3f}", ha="center", fontsize=9, fontweight="bold")
            axes[i].set_xticks(range(len(names_ml)))
            axes[i].set_xticklabels([n.replace(" ","\n") for n in names_ml], fontsize=8)
            axes[i].set_title(tit, fontsize=11, fontweight="bold")
            axes[i].yaxis.grid(True); axes[i].set_axisbelow(True)
        plt.suptitle("Comparación de Modelos Predictivos — Estancia", fontsize=13, fontweight="bold", y=1.02)
        plt.tight_layout(); save_fig(out,"F22_modelos_comparacion")

    # F23 — Real vs Predicho (mejor modelo)
    if mejor:
        model_best = models[mejor]
        yb = model_best.predict(X_te)
        lim = max(y_te.max(), yb.max()) + 1
        fig, ax = plt.subplots(figsize=(7,6))
        sc = ax.scatter(y_te, yb, alpha=0.45, s=25, c=np.abs(y_te-yb),
                        cmap="plasma", edgecolors="none")
        ax.plot([0,lim],[0,lim],"--",color="#FFD600",lw=1.5,label="Línea perfecta")
        ax.set_title(f"Real vs Predicho — {mejor}", fontsize=13, fontweight="bold")
        ax.set_xlabel("Días reales"); ax.set_ylabel("Días predichos")
        ax.legend(framealpha=0.2); ax.grid(True,alpha=0.3)
        plt.colorbar(sc, ax=ax, label="Error absoluto")
        plt.tight_layout(); save_fig(out,"F23_real_predicho")

    # F24 — Importancia de variables (Random Forest)
    if "Random Forest" in models:
        rf = models["Random Forest"]
        try:
            imp_s = pd.Series(rf.feature_importances_, index=all_feats).sort_values(ascending=False).head(15)
            fig, ax = plt.subplots(figsize=(10,6))
            bar_h(ax, imp_s.index.tolist(), imp_s.values, PAL[:15], fmt=lambda v: f"{v:.3f}")
            ax.set_title("Importancia de Variables — Random Forest", fontsize=13, fontweight="bold")
            plt.tight_layout(); save_fig(out,"F24_importancia_variables")
            res["importancia_variables"] = imp_s.to_dict()
        except Exception as e: print(f"    Importancia: {e}")

    # ─── K-MEANS ───────────────────────────────────────────────────────────────
    km_feats_cands = ["EDAD","TAS","FC","TEMP","SPO2","ESTANCIA","VALOR_FAC","SCORE_COMPLEJIDAD"]
    km_feats = [c for c in km_feats_cands if c in df.columns]
    if len(km_feats) >= 3:
        Xk = df[km_feats].copy()
        imp2 = SimpleImputer(strategy="median")
        sc2  = StandardScaler()
        Xk_i = imp2.fit_transform(Xk)
        Xk_s = sc2.fit_transform(Xk_i)

        inertias=[]; sils=[]
        k_range = range(2,8)
        for k in k_range:
            km_ = KMeans(n_clusters=k, random_state=42, n_init=10)
            km_.fit(Xk_s)
            inertias.append(km_.inertia_)
            sils.append(silhouette_score(Xk_s, km_.labels_, sample_size=min(500,len(Xk_s))))

        K_OPT = max(3, min(list(k_range)[np.argmax(sils)], 5))
        km_f = KMeans(n_clusters=K_OPT, random_state=42, n_init=20)
        km_f.fit(Xk_s)
        df["CLUSTER_KMEANS"] = km_f.labels_
        res["K_optimo"] = K_OPT
        res["silhouette"] = round(float(max(sils)),4)
        print(f"    K-Means: K={K_OPT}  Silhouette={res['silhouette']}")

        # F25 — Elbow + Silhouette
        fig, (ax1,ax2) = plt.subplots(1,2,figsize=(11,4))
        ax1.plot(list(k_range),inertias,"o-",color=ACC,lw=2,markersize=8)
        ax1.axvline(K_OPT,color="#FFD600",linestyle="--",lw=1.5,label=f"K={K_OPT}")
        ax1.set_title("Método del Codo",fontsize=12,fontweight="bold"); ax1.legend(framealpha=0.3); ax1.yaxis.grid(True)
        ax2.plot(list(k_range),sils,"s-",color="#F15BB5",lw=2,markersize=8)
        ax2.axvline(K_OPT,color="#FFD600",linestyle="--",lw=1.5)
        ax2.set_title("Silhouette Score",fontsize=12,fontweight="bold"); ax2.yaxis.grid(True)
        plt.suptitle("Selección K Óptimo — Clustering Pacientes",fontsize=13,fontweight="bold",y=1.02)
        plt.tight_layout(); save_fig(out,"F25_kmeans_elbow")

        # F26 — PCA clusters
        pca = PCA(n_components=2, random_state=42)
        Xp = pca.fit_transform(Xk_s)
        fig, ax = plt.subplots(figsize=(9,6))
        for c in range(K_OPT):
            m = km_f.labels_==c
            ax.scatter(Xp[m,0],Xp[m,1],alpha=0.5,s=25,color=PAL[c],edgecolors="none",
                       label=f"Cluster {c} (n={m.sum()})")
        cen = pca.transform(km_f.cluster_centers_)
        ax.scatter(cen[:,0],cen[:,1],s=200,c="white",marker="X",zorder=5,edgecolors=BG,lw=1.5,label="Centroides")
        ax.set_title(f"Clusters K-Means (K={K_OPT}) — PCA 2D",fontsize=13,fontweight="bold")
        ax.legend(framealpha=0.2,fontsize=9); ax.grid(True,alpha=0.3)
        save_fig(out,"F26_kmeans_pca")

        # F27 — Perfil de clusters
        cluster_profiles = {}
        fig, axes = plt.subplots(2,3,figsize=(14,8))
        for idx, feat in enumerate(km_feats[:6]):
            ax = axes[idx//3][idx%3]
            data_c = [df[df["CLUSTER_KMEANS"]==c][feat].dropna().values for c in range(K_OPT)]
            bp = ax.boxplot(data_c, patch_artist=True, notch=False,
                medianprops=dict(color="#FFD600",lw=2),
                boxprops=dict(alpha=0.7), whiskerprops=dict(color=TEXT),
                capprops=dict(color=TEXT), flierprops=dict(marker="o",markersize=2,alpha=0.3))
            for patch, col in zip(bp["boxes"], PAL): patch.set_facecolor(col)
            ax.set_title(feat, fontsize=10, fontweight="bold")
            ax.set_xticklabels([f"C{c}" for c in range(K_OPT)])
            ax.yaxis.grid(True); ax.set_axisbelow(True)
        for c in range(K_OPT):
            sub = df[df["CLUSTER_KMEANS"]==c]
            cluster_profiles[int(c)] = {
                "n": int(len(sub)),
                "est_media": round(float(sub["ESTANCIA"].mean()),2) if "ESTANCIA" in sub.columns else 0,
                "edad_media": round(float(sub["EDAD"].mean()),1) if "EDAD" in sub.columns else 0,
                "valor_fac_media": round(float(sub["VALOR_FAC"].mean()),0) if "VALOR_FAC" in sub.columns else 0,
            }
        res["cluster_profiles"] = cluster_profiles
        plt.suptitle("Perfil de Clusters — Características Clínicas",fontsize=13,fontweight="bold",y=1.01)
        plt.tight_layout(); save_fig(out,"F27_kmeans_perfil")

    # ─── Regresión múltiple (statsmodels si disponible) ───────────────────────
    if HAS_SM and len(all_feats) >= 2:
        try:
            Xsm = sm.add_constant(X_tr)
            model_sm = sm.OLS(y_tr, Xsm).fit()
            coefs = {"intercept": round(float(model_sm.params[0]),4)}
            pvals = {}
            for i,feat in enumerate(all_feats):
                if i+1 < len(model_sm.params):
                    coefs[feat] = round(float(model_sm.params[i+1]),4)
                    pvals[feat] = round(float(model_sm.pvalues[i+1]),6)
            res["regresion_multiple"] = {
                "R2": round(float(model_sm.rsquared),4),
                "R2_adj": round(float(model_sm.rsquared_adj),4),
                "F_stat": round(float(model_sm.fvalue),3),
                "F_pval": round(float(model_sm.f_pvalue),6),
                "coeficientes": coefs,
                "p_valores": pvals,
            }
            # F28 — Coeficientes significativos
            sig = {k:v for k,v in pvals.items() if v < 0.05}
            if sig:
                fig, ax = plt.subplots(figsize=(10,5))
                skeys = sorted(sig, key=sig.get)[:15]
                svals = [coefs.get(k,0) for k in skeys]
                colors = ["#00C875" if v>0 else "#FF3B3B" for v in svals]
                ax.barh(skeys[::-1], svals[::-1], color=colors[::-1], alpha=0.88, edgecolor=BG)
                ax.axvline(0,color="white",lw=1.2)
                ax.set_title("Variables Significativas — Regresión Múltiple (p<0.05)",fontsize=12,fontweight="bold")
                ax.xaxis.grid(True); ax.set_axisbelow(True)
                plt.tight_layout(); save_fig(out,"F28_regresion_multiple")
        except Exception as e: print(f"    Statsmodels: {e}")

    # ─── Predicción de Costo ──────────────────────────────────────────────────
    if "VALOR_FAC" in df_ml.columns:
        TARGET_COSTO = "VALOR_FAC"
        feats_c = [c for c in all_feats + [TARGET_COSTO] if c in df_ml.columns and c != TARGET]
        df_c = df_ml[feats_c + [TARGET_COSTO]].dropna() if TARGET_COSTO not in feats_c else df_ml[feats_c].dropna()
        if len(df_c) >= 50:
            Xc = df_c[[c for c in all_feats if c in df_c.columns]].values
            yc = df_c[TARGET_COSTO].values
            Xc = SimpleImputer(strategy="median").fit_transform(Xc)
            Xct,Xce,yct,yce = train_test_split(Xc, yc, test_size=0.2, random_state=42)
            rf_c = RandomForestRegressor(n_estimators=100, max_depth=6, random_state=42, n_jobs=-1)
            rf_c.fit(Xct,yct); ycp = rf_c.predict(Xce)
            res["prediccion_costo"] = {
                "MAE": round(float(mean_absolute_error(yce,ycp)),0),
                "RMSE": round(float(np.sqrt(mean_squared_error(yce,ycp))),0),
                "R2": round(float(r2_score(yce,ycp)),4),
            }
            print(f"    Predicción costo: MAE={fmt_m(res['prediccion_costo']['MAE'])}  R²={res['prediccion_costo']['R2']}")
            # F29 — Real vs Predicho costo
            fig, ax = plt.subplots(figsize=(7,6))
            ax.scatter(yce, ycp, alpha=0.45, s=25, c=np.abs(yce-ycp), cmap="plasma", edgecolors="none")
            lim_c = max(yce.max(),ycp.max())+1
            ax.plot([0,lim_c],[0,lim_c],"--",color="#FFD600",lw=1.5)
            ax.set_title("Real vs Predicho — Valor Facturado", fontsize=13, fontweight="bold")
            ax.set_xlabel("Valor real ($)"); ax.set_ylabel("Valor predicho ($)")
            ax.grid(True,alpha=0.3); plt.tight_layout(); save_fig(out,"F29_prediccion_costo")

    # F30 — Score de complejidad por cluster / EPS / especialidad
    if "SCORE_COMPLEJIDAD" in df.columns:
        fig, axes = plt.subplots(1,2,figsize=(14,5))
        if "EPS" in df.columns:
            se = df.groupby("EPS")["SCORE_COMPLEJIDAD"].mean().sort_values(ascending=False).head(10)
            ns = [n[:30]+"…" if len(n)>30 else n for n in se.index]
            bar_h(axes[0], ns, se.values, PAL[:10], fmt=lambda v:f"{v:.1f}")
            axes[0].set_title("Score Complejidad por EPS", fontsize=11, fontweight="bold")
        if "ESPECIALIDAD" in df.columns:
            sp = df.groupby("ESPECIALIDAD")["SCORE_COMPLEJIDAD"].mean().sort_values(ascending=False).head(10)
            ns2 = [n[:30]+"…" if len(n)>30 else n for n in sp.index]
            bar_h(axes[1], ns2, sp.values, PAL[::-1][:10], fmt=lambda v:f"{v:.1f}")
            axes[1].set_title("Score Complejidad por Especialidad", fontsize=11, fontweight="bold")
        plt.suptitle("Score de Complejidad Clínica (0-100)",fontsize=13,fontweight="bold",y=1.01)
        plt.tight_layout(); save_fig(out,"F30_score_complejidad")

    # Guardar encoders en res para predictor HTML
    res["encoders"] = encoders
    res["all_feats"] = all_feats
    res["num_feats"] = num_feats
    return res

# ═══════════════════════════════════════════════════════════════════════════════
# 10. ESTADÍSTICA COMPLEMENTARIA
# ═══════════════════════════════════════════════════════════════════════════════
def analisis_estadistica(df, out):
    if df is None: return {}
    print("\n[9] Estadística complementaria...")

    # F31 — Correlación numérica
    cols_corr = ["EDAD","TAS","TAD","FC","FR","TEMP","SPO2","ESTANCIA","VALOR_FAC","SCORE_COMPLEJIDAD"]
    cols_ok = [c for c in cols_corr if c in df.columns]
    if len(cols_ok) >= 4:
        fig, ax = plt.subplots(figsize=(max(8,len(cols_ok)),max(7,len(cols_ok)-1)))
        corr = df[cols_ok].dropna().corr()
        mask = np.triu(np.ones_like(corr,dtype=bool))
        sns.heatmap(corr, mask=mask, cmap=CMAP_HEAT, vmax=1, vmin=-1, center=0,
                    annot=True, fmt=".2f", square=True, ax=ax, linewidths=.5, linecolor=BG,
                    cbar_kws={"shrink":.8}, annot_kws={"size":8,"color":TEXT})
        short = {"EDAD":"Edad","TAS":"TAS","TAD":"TAD","FC":"FC","FR":"FR","TEMP":"Temp",
                 "SPO2":"SpO2","ESTANCIA":"Estancia","VALOR_FAC":"Facturado","SCORE_COMPLEJIDAD":"Complejidad"}
        sl = [short.get(c,c) for c in cols_ok]
        ax.set_xticklabels(sl,rotation=30,ha="right",fontsize=9)
        ax.set_yticklabels(sl,rotation=0,fontsize=9)
        ax.set_title("Matriz de Correlación — Variables Numéricas",fontsize=13,fontweight="bold")
        plt.tight_layout(); save_fig(out,"F31_correlacion")

    # F32 — Signos vitales boxplots
    sv_cols = [("TAS","T.A.Sistólica\n(mmHg)"),("TAD","T.A.Diastólica\n(mmHg)"),
               ("FC","F.Cardíaca\n(lpm)"),("FR","F.Resp.\n(rpm)"),
               ("TEMP","Temperatura\n(°C)"),("SPO2","SpO₂\n(%)")]
    sv_ok = [(c,l) for c,l in sv_cols if c in df.columns]
    if sv_ok:
        fig, axes = plt.subplots(2,3,figsize=(12,7))
        for i,(col,lbl) in enumerate(sv_ok[:6]):
            ax = axes[i//3][i%3]
            dv = df[col].dropna()
            dv = dv[(dv>dv.quantile(.01)) & (dv<dv.quantile(.99))]
            ax.boxplot(dv, patch_artist=True, notch=True,
                medianprops=dict(color="#FFD600",lw=2),
                boxprops=dict(facecolor=PAL[i],alpha=0.7),
                whiskerprops=dict(color=TEXT), capprops=dict(color=TEXT),
                flierprops=dict(marker="o",markerfacecolor=ACC,markersize=3,alpha=0.3))
            ax.set_title(lbl,fontsize=10,fontweight="bold")
            ax.yaxis.grid(True); ax.set_axisbelow(True); ax.set_xticklabels([])
        for i in range(len(sv_ok),6):
            axes[i//3][i%3].set_visible(False)
        plt.suptitle("Distribución de Signos Vitales al Ingreso",fontsize=13,fontweight="bold",y=1.02)
        plt.tight_layout(); save_fig(out,"F32_signos_vitales")

    print("    ✓ Figuras estadísticas generadas")
    return {}

# ═══════════════════════════════════════════════════════════════════════════════
# 11. EXPORTAR JSON PARA DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
def exportar_json(df, data, calidad_res, hosp_res, fin_res, ml_res, out):
    print("\n[10] Exportando JSON para dashboard...")

    def safe(v):
        if isinstance(v,float) and np.isnan(v): return None
        if isinstance(v,(np.integer,)): return int(v)
        if isinstance(v,(np.floating,)): return float(v)
        if isinstance(v,dict): return {str(k):safe(vv) for k,vv in v.items()}
        if isinstance(v,list): return [safe(i) for i in v]
        return v

    # KPIs generales
    kpis = {}
    if df is not None:
        kpis["total_ingresos"] = int(df["INGRESO"].nunique()) if "INGRESO" in df.columns else int(len(df))
        kpis["edad_media"]     = round(float(df["EDAD"].mean()),1) if "EDAD" in df.columns else None
        kpis["estancia_media"] = round(float(df["ESTANCIA"].mean()),2) if "ESTANCIA" in df.columns else None
        kpis["estancia_max"]   = int(df["ESTANCIA"].max()) if "ESTANCIA" in df.columns else None
        kpis["pct_femenino"]   = round(float((df["SEXO"]=="FEMENINO").mean()*100),1) if "SEXO" in df.columns else None
        kpis["total_facturado"]= fin_res.get("total_facturado")
        kpis["total_utilidad"] = fin_res.get("total_utilidad")
        kpis["margen_global"]  = fin_res.get("margen_global_pct")
        kpis["mortalidad_pct"] = round(float(df["MORTALIDAD_BIN"].mean()*100),2) if "MORTALIDAD_BIN" in df.columns else None
        kpis["reingreso_pct"]  = round(float(df["REINGRESO_BIN"].mean()*100),2) if "REINGRESO_BIN" in df.columns else None

    # Listas para predictor
    dx_list, esp_list, eps_list, serv_list = [], [], [], []
    if df is not None:
        if "COD_DX" in df.columns and "NOM_DX" in df.columns and "ESTANCIA" in df.columns:
            for cod in df["COD_DX"].value_counts().head(100).index:
                sub = df[df["COD_DX"]==cod]
                if len(sub) > 0:
                    dx_list.append({"cod":str(cod),
                                    "nom":str(sub["NOM_DX"].iloc[0]),
                                    "n":int(len(sub)),
                                    "est_media":round(float(sub["ESTANCIA"].mean()),2),
                                    "val_media":round(float(sub["VALOR_FAC"].mean()),0) if "VALOR_FAC" in sub.columns else 0})
        if "ESPECIALIDAD" in df.columns:
            esp_list = sorted(df["ESPECIALIDAD"].dropna().unique().tolist())
        if "EPS" in df.columns:
            eps_list = sorted(df["EPS"].dropna().unique().tolist())
        if "SERVICIO" in df.columns:
            serv_list = sorted(df["SERVICIO"].dropna().unique().tolist())

    payload = {
        "generado": datetime.now().isoformat(),
        "calidad": safe(calidad_res),
        "kpis": safe(kpis),
        "hospitalizacion": safe(hosp_res),
        "financiero": safe(fin_res),
        "ml": safe({k:v for k,v in ml_res.items() if k not in ("encoders",)}),
        "opciones": {
            "diagnosticos": dx_list,
            "especialidades": esp_list,
            "eps": eps_list,
            "servicios": serv_list,
            "regimen": ["SUBSIDIADO","CONTRIBUTIVO","OTRO","VINCULADO","PARTICULAR"],
            "zona": ["URBANA","RURAL"],
            "sexo": ["FEMENINO","MASCULINO"],
            "via": ["URGENCIAS","REMITIDO","PROGRAMADO"],
        },
        "encoders": safe(ml_res.get("encoders",{})),
        "all_feats": ml_res.get("all_feats",[]),
        "cluster_profiles": safe(ml_res.get("cluster_profiles",{})),
        "stats_estancia": safe({
            "mean": round(float(df["ESTANCIA"].mean()),2) if df is not None and "ESTANCIA" in df.columns else None,
            "median": round(float(df["ESTANCIA"].median()),2) if df is not None and "ESTANCIA" in df.columns else None,
            "std":  round(float(df["ESTANCIA"].std()),2) if df is not None and "ESTANCIA" in df.columns else None,
            "p75":  round(float(df["ESTANCIA"].quantile(.75)),2) if df is not None and "ESTANCIA" in df.columns else None,
            "p90":  round(float(df["ESTANCIA"].quantile(.90)),2) if df is not None and "ESTANCIA" in df.columns else None,
        }),
    }

    path_json = os.path.join(out, "analytics_output.json")
    with open(path_json,"w",encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
    print(f"    ✓ analytics_output.json  ({os.path.getsize(path_json)//1024} KB)")
    return payload

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    args   = parse_args()
    OUT    = args.out
    os.makedirs(OUT, exist_ok=True)
    excel  = seleccionar_excel(args.excel or EXCEL_DEFAULT)
    if not excel or not os.path.exists(excel):
        print(f"ERROR: archivo no encontrado → '{excel}'"); sys.exit(1)

    t0 = datetime.now()

    # 1. Cargar
    data        = cargar_datos(excel)

    # 2. Preparar hoja principal
    df_hosp     = preparar_hosp(data.get("hosp"))

    # 3. Calidad
    cal_res     = reporte_calidad(df_hosp, OUT)

    # 4. Hospitalización
    hosp_res    = analisis_hospitalizacion(df_hosp, OUT)

    # 5. Cirugía
    analisis_cirugia(data.get("cirugia"), OUT)

    # 6. UCI
    analisis_uci(data.get("uci"), OUT)

    # 7. Ambulatorio
    analisis_ambulatorio(data.get("amb"), OUT)

    # 8. Financiero
    fin_res     = analisis_financiero(df_hosp, data.get("cups"), OUT)

    # 9. Estadística
    analisis_estadistica(df_hosp, OUT)

    # 10. ML
    ml_res      = analisis_ml(df_hosp, OUT)

    # 11. JSON
    payload     = exportar_json(df_hosp, data, cal_res, hosp_res, fin_res, ml_res, OUT)

    dt = (datetime.now()-t0).seconds
    print(f"\n{'═'*62}")
    print(f"  ✓ Análisis completo en {dt}s  |  Figuras → ./{OUT}/")
    if payload.get("kpis"):
        k = payload["kpis"]
        if k.get("total_ingresos"): print(f"  Ingresos: {k['total_ingresos']:,}")
        if k.get("estancia_media"): print(f"  Estancia media: {k['estancia_media']}d")
        if k.get("margen_global"):  print(f"  Margen global: {k['margen_global']}%")
    if ml_res.get("mejor_modelo"):
        mm = ml_res["mejor_modelo"]
        print(f"  Mejor modelo: {mm}  MAE={ml_res['modelos'][mm]['MAE']}d  R²={ml_res['modelos'][mm]['R2']}")
    if ml_res.get("K_optimo"):
        print(f"  K-Means: K={ml_res['K_optimo']}  Silhouette={ml_res['silhouette']}")
    print(f"{'═'*62}")

import matplotlib.lines as _ml
plt.mlines = _ml
if __name__ == "__main__":
    main()
