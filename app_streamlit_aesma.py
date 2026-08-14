# app.py
# ------------------------------------------------------------
# Application Streamlit - Rapport journalier AESMA
# ------------------------------------------------------------
# Lancement local :
#   pip install streamlit pandas openpyxl matplotlib seaborn numpy xlsxwriter
#   streamlit run app.py
#
# Fichiers attendus :
#   1) Fichier interventions Excel, ou URL Google Sheets export xlsx
#   2) Fichier group_f.xlsx
#   3) Fichier alarmes Excel exporté depuis le système alarmes
# ------------------------------------------------------------

import io
import math
from textwrap import fill
from datetime import time

import numpy as np
import pandas as pd
import streamlit as st

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.colors import ListedColormap, BoundaryNorm
import seaborn as sns


# ============================================================
# CONFIG STREAMLIT
# ============================================================
st.set_page_config(
    page_title="AESMA - Générateur de rapport journalier",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Générateur de rapport journalier AESMA")
st.caption("Importez les fichiers Excel, générez le rapport visuel, puis téléchargez le PDF et l'Excel.")


# ============================================================
# CONSTANTES COLONNES
# ============================================================
DATE_COL = "Date et Heure début d'intervention"
DATE_FIN_COL = "Date et Heure fin d'intervention"
SHIFT_COL = "SHIFT"
OPERATEUR_COL = "Prénom et nom"
TYPE_COL = "Choix"
EQUIPEMENT_COL = "Équipement"
COMMENT_COL = "Commentaire"
ZONE_GLOBALE_COL = "Zone geo glob"
ZONE_COL = "Zone geo"
TECH_DEFECT_COL = "Technique - prédit"
OPE_DEFECT_COL = "Operationnel - prédit"

# Style
BG = "#F3F4F6"
CARD_BG = "#FFFFFF"
CARD_EDGE = "#D1D5DB"
TITLE = "#111827"
SUBTITLE = "#6B7280"
PRIMARY = "#1F4E79"
ACCENT = "#4B5563"
SUCCESS = "#0F766E"
DANGER = "#B91C1C"
WARNING = "#D97706"
TEXT = "#1f1f1f"

plt.rcParams["figure.facecolor"] = BG
plt.rcParams["axes.facecolor"] = CARD_BG
plt.rcParams["font.family"] = "DejaVu Sans"
sns.set_style("whitegrid")


# ============================================================
# OUTILS DONNEES
# ============================================================
def clean_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.strip()
        .replace({"nan": np.nan, "None": np.nan, "": np.nan, " ": np.nan})
    )


def clean_equipment_col(df: pd.DataFrame, col: str) -> pd.DataFrame:
    df = df.copy()
    if col not in df.columns:
        return df

    s = df[col].astype(str)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace("-", ".", regex=False)
    s = s.str.replace("EC.", "EC", regex=False)
    s = s.str.replace("XC.", "XC", regex=False)
    s = s.str.replace("BC.", "BC", regex=False)
    s = s.str.replace("IU.", "IU", regex=False)

    for v in ["V1", "V2"]:
        s = s.str.replace(f" {v}", v, regex=False)
        s = s.str.replace(v, f".{v}", regex=False)
        s = s.str.replace(f"..{v}", f".{v}", regex=False)

    s = s.str.replace("A1.V1", ".V1", regex=False)
    s = s.str.replace("A2.V2", ".V2", regex=False)
    s = s.str.replace(r"^(WE\d+)$", r"\1.01", regex=True)
    s = s.str.replace(r"^IU(\d)(?=\D)", r"IU0\1", regex=True)

    df[col] = s
    return df


def safe_value_counts(df: pd.DataFrame, col: str, top_n=None) -> pd.Series:
    if col not in df.columns:
        return pd.Series(dtype=int)
    vc = clean_series(df[col]).dropna().value_counts()
    return vc.head(top_n) if top_n else vc


def wrap_text(text, width=28):
    return fill(str(text), width=width)


def autopct_count(values):
    total = sum(values)

    def inner(pct):
        count = int(round(pct * total / 100))
        return f"{count}\n({pct:.0f}%)" if pct >= 6 else ""

    return inner


def validate_required_columns(df: pd.DataFrame, required_cols: list[str], label: str):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes dans {label} : {missing}")


# ============================================================
# CHARGEMENT
# ============================================================
@st.cache_data(show_spinner=False)
def read_excel_cached(file_bytes: bytes, skiprows: int = 0) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), skiprows=skiprows)


@st.cache_data(show_spinner=False)
def read_excel_from_url(url: str) -> pd.DataFrame:
    return pd.read_excel(url)


# ============================================================
# TRAITEMENT INTERVENTIONS
# ============================================================
def prepare_interventions(df_raw: pd.DataFrame, group_raw: pd.DataFrame):
    df = clean_equipment_col(df_raw, EQUIPEMENT_COL)

    group = group_raw.copy()
    if "Equipement" in group.columns:
        group = group.rename(columns={"Equipement": EQUIPEMENT_COL})
    group = clean_equipment_col(group, EQUIPEMENT_COL)

    df = pd.merge(df, group, on=EQUIPEMENT_COL, how="left")
    df = df[~df[ZONE_GLOBALE_COL].isna()].copy()

    required_cols = [
        DATE_COL, SHIFT_COL, OPERATEUR_COL, TYPE_COL, EQUIPEMENT_COL,
        ZONE_GLOBALE_COL, ZONE_COL, TECH_DEFECT_COL, OPE_DEFECT_COL
    ]
    validate_required_columns(df, required_cols, "le fichier interventions après jointure group_f")

    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df = df.dropna(subset=[DATE_COL]).copy()

    report_date = df[DATE_COL].max().normalize()
    start_dt = report_date - pd.Timedelta(days=1) + pd.Timedelta(hours=17)
    end_dt = report_date + pd.Timedelta(hours=5)

    df_period = df[(df[DATE_COL] >= start_dt) & (df[DATE_COL] <= end_dt)].copy()
    if df_period.empty:
        raise ValueError(
            f"Aucune donnée trouvée entre {start_dt.strftime('%d/%m/%Y %H:%M')} "
            f"et {end_dt.strftime('%d/%m/%Y %H:%M')}."
        )

    for c in [SHIFT_COL, OPERATEUR_COL, TYPE_COL, EQUIPEMENT_COL, ZONE_GLOBALE_COL, ZONE_COL, TECH_DEFECT_COL, OPE_DEFECT_COL]:
        df_period[c] = clean_series(df_period[c])

    df_period["heure"] = df_period[DATE_COL].dt.hour

    return df, df_period, start_dt, end_dt


# ============================================================
# TRAITEMENT ALARMES
# ============================================================
def prepare_alarms(alarm_raw: pd.DataFrame, group_raw: pd.DataFrame):
    da = alarm_raw.copy()

    # Certains exports contiennent des colonnes vides nommées Unnamed.
    drop_cols = [
        "Unnamed: 5", "Unnamed: 8", "Unnamed: 10",
        "Unnamed: 12", "Unnamed: 13", "Unnamed: 14", "Unnamed: 15"
    ]
    da = da.drop(columns=[c for c in drop_cols if c in da.columns], errors="ignore")

    required_alarm = ["Duration", "Description", "Equipment", "On Time", "Off Time"]
    validate_required_columns(da, required_alarm, "le fichier alarmes")

    da = da[~da["Duration"].isna()].copy()
    da["Duration"] = da["Duration"].astype(str).str.replace("h", "", regex=False)
    da["Duration_timedelta"] = pd.to_timedelta(da["Duration"], errors="coerce")
    da["Duration_seconds"] = da["Duration_timedelta"].dt.total_seconds()

    da = da[da["Description"].astype(str).str.contains("Sensor Fault", na=False)].copy()
    da = da[~da["Equipment"].astype(str).str.startswith(("BN", "NC"), na=False)].copy()

    group = group_raw.copy()
    if "Equipement" in group.columns:
        group = group.rename(columns={"Equipement": "Equipment"})

    df_merge = pd.merge(group, da, on="Equipment", how="left")
    df_merge = df_merge[~df_merge[ZONE_GLOBALE_COL].isna()].copy()

    df_merge["On Time"] = pd.to_datetime(df_merge["On Time"], format="%d/%m/%y %H:%M:%S", errors="coerce")
    df_merge["Off Time"] = pd.to_datetime(df_merge["Off Time"], format="%d/%m/%y %H:%M:%S", errors="coerce")
    df_merge = df_merge.dropna(subset=["On Time"]).copy()

    df_merge["On Hour"] = df_merge["On Time"].dt.time

    start_time = pd.to_datetime("16:00:00").time()
    end_time = pd.to_datetime("04:30:00").time()
    mask = (df_merge["On Hour"] >= start_time) | (df_merge["On Hour"] <= end_time)

    df_mask = df_merge[mask].copy()
    df_mask = df_mask[df_mask["Duration_seconds"] < 600].copy()

    end_shift = pd.to_datetime("04:30:00").time()

    def assign_shift_date(row):
        hour = row["On Time"].time()
        date = row["On Time"].date()
        if hour <= end_shift:
            return date - pd.Timedelta(days=1)
        return date

    df_mask["date_real"] = df_mask.apply(assign_shift_date, axis=1)

    now = df_mask["On Time"].max()
    start_24h = now - pd.Timedelta(hours=24)
    mask_24h = df_mask["On Time"].between(start_24h, now)
    hours = df_mask["On Time"].dt.hour
    mask_nuit = (hours >= 17) | (hours < 5)

    df_filtre = df_mask[mask_24h & mask_nuit].copy()

    df_grouped = (
        df_filtre
        .groupby([ZONE_COL, "date_real"])["Duration_seconds"]
        .agg(["mean", "count"])
        .reset_index()
    )

    return df_mask, df_filtre, df_grouped


# ============================================================
# KPI
# ============================================================
def compute_kpis(df_period: pd.DataFrame, df_grouped: pd.DataFrame):
    total_pannes = len(df_period)

    shift_counts = safe_value_counts(df_period, SHIFT_COL)
    zone_counts = safe_value_counts(df_period, ZONE_COL, top_n=10)
    equip_counts = safe_value_counts(df_period, EQUIPEMENT_COL, top_n=10)
    hour_counts = df_period["heure"].value_counts().sort_index()

    top_shift = shift_counts.index[0] if not shift_counts.empty else "-"
    top_shift_nb = int(shift_counts.iloc[0]) if not shift_counts.empty else 0

    top_zone = zone_counts.index[0] if not zone_counts.empty else "-"
    top_zone_nb = int(zone_counts.iloc[0]) if not zone_counts.empty else 0

    top_equip = equip_counts.index[0] if not equip_counts.empty else "-"
    top_equip_nb = int(equip_counts.iloc[0]) if not equip_counts.empty else 0

    hour_peak = int(hour_counts.idxmax()) if not hour_counts.empty else 0
    hour_peak_nb = int(hour_counts.max()) if not hour_counts.empty else 0

    nb_equipements = df_period[EQUIPEMENT_COL].dropna().nunique()
    nb_operateurs = df_period[OPERATEUR_COL].dropna().nunique()
    nb_zones_globales = df_period[ZONE_GLOBALE_COL].dropna().nunique()
    nb_sous_zones = df_period[ZONE_COL].dropna().nunique()

    tech_defects = clean_series(df_period[TECH_DEFECT_COL]).dropna().astype("string")
    ope_defects = clean_series(df_period[OPE_DEFECT_COL]).dropna().astype("string")

    tech_defects = tech_defects[~tech_defects.str.lower().isin(["non", "nan", "none"])]
    ope_defects = ope_defects[~ope_defects.str.lower().isin(["non", "nan", "none"])]

    all_defects = pd.concat([tech_defects, ope_defects], ignore_index=True)
    defect_counts = all_defects.value_counts().head(8)

    heatmap_hour = (
        df_period
        .groupby(["heure", SHIFT_COL])
        .size()
        .unstack(fill_value=0)
    )
    heatmap_hour = heatmap_hour.reindex(list(range(24)), fill_value=0)
    heatmap_hour = heatmap_hour.loc[heatmap_hour.sum(axis=1) > 0]
    if heatmap_hour.empty:
        heatmap_hour = pd.DataFrame([[0]], index=["Aucune donnée"], columns=["-"])

    ops_zone_detail = df_period[[OPERATEUR_COL, ZONE_COL]].copy()
    ops_zone_detail[OPERATEUR_COL] = clean_series(ops_zone_detail[OPERATEUR_COL])
    ops_zone_detail[ZONE_COL] = (
        ops_zone_detail[ZONE_COL]
        .astype(str).str.strip()
        .replace(["nan", "None", "", " "], "N-A")
    )
    ops_zone_detail = ops_zone_detail.dropna(subset=[OPERATEUR_COL]).copy()

    operateur_zone_summary = (
        ops_zone_detail
        .groupby([OPERATEUR_COL, ZONE_COL], as_index=False)
        .size()
        .rename(columns={"size": "Nombre_interventions"})
        .sort_values(by=["Nombre_interventions", OPERATEUR_COL, ZONE_COL], ascending=[False, True, True])
        .reset_index(drop=True)
    )

    operateur_summary = df_period[[OPERATEUR_COL]].copy()
    operateur_summary[OPERATEUR_COL] = clean_series(operateur_summary[OPERATEUR_COL])
    operateur_summary = operateur_summary.dropna(subset=[OPERATEUR_COL])
    operateur_summary = (
        operateur_summary
        .groupby(OPERATEUR_COL, as_index=False)
        .size()
        .rename(columns={"size": "Nombre_interventions"})
        .sort_values(by=["Nombre_interventions", OPERATEUR_COL], ascending=[False, True])
        .reset_index(drop=True)
    )

    ops_heatmap = (
        operateur_zone_summary
        .pivot(index=OPERATEUR_COL, columns=ZONE_COL, values="Nombre_interventions")
        .fillna(0)
    )

    if not ops_heatmap.empty:
        row_order = ops_heatmap.sum(axis=1).sort_values(ascending=False).index
        col_order = ops_heatmap.sum(axis=0).sort_values(ascending=False).index
        ops_heatmap = ops_heatmap.loc[row_order, col_order]
        ops_heatmap = ops_heatmap.head(15)

    df_zone = df_grouped.copy()
    if not df_zone.empty:
        df_zone["date_real"] = pd.to_datetime(df_zone["date_real"], errors="coerce")
        df_zone = df_zone.sort_values(["mean", "count"], ascending=[False, False]).reset_index(drop=True)

    mean_reactivity_global = round(df_zone["mean"].mean(), 1) if not df_zone.empty else 0
    worst_zone_reactivity = df_zone.iloc[0][ZONE_COL] if not df_zone.empty else "-"
    worst_zone_reactivity_val = round(df_zone.iloc[0]["mean"], 1) if not df_zone.empty else 0
    nb_critical_90 = int((df_zone["mean"] > 90).sum()) if not df_zone.empty else 0
    nb_critical_60_90 = int(((df_zone["mean"] > 60) & (df_zone["mean"] <= 90)).sum()) if not df_zone.empty else 0

    return {
        "total_pannes": total_pannes,
        "shift_counts": shift_counts,
        "zone_counts": zone_counts,
        "equip_counts": equip_counts,
        "hour_counts": hour_counts,
        "defect_counts": defect_counts,
        "heatmap_hour": heatmap_hour,
        "operateur_zone_summary": operateur_zone_summary,
        "operateur_summary": operateur_summary,
        "ops_heatmap": ops_heatmap,
        "df_zone": df_zone,
        "top_shift": top_shift,
        "top_shift_nb": top_shift_nb,
        "top_zone": top_zone,
        "top_zone_nb": top_zone_nb,
        "top_equip": top_equip,
        "top_equip_nb": top_equip_nb,
        "hour_peak": hour_peak,
        "hour_peak_nb": hour_peak_nb,
        "nb_equipements": nb_equipements,
        "nb_operateurs": nb_operateurs,
        "nb_zones_globales": nb_zones_globales,
        "nb_sous_zones": nb_sous_zones,
        "mean_reactivity_global": mean_reactivity_global,
        "worst_zone_reactivity": worst_zone_reactivity,
        "worst_zone_reactivity_val": worst_zone_reactivity_val,
        "nb_critical_90": nb_critical_90,
        "nb_critical_60_90": nb_critical_60_90,
    }


# ============================================================
# FIGURES RAPPORT
# ============================================================
def page_template(title, start_dt, end_dt, logo_text="AESMA"):
    fig = plt.figure(figsize=(16, 9), facecolor=BG)

    ax_header = fig.add_axes([0.04, 0.90, 0.92, 0.08])
    ax_header.set_axis_off()

    header_box = patches.FancyBboxPatch(
        (0, 0), 1, 1,
        boxstyle="round,pad=0.008,rounding_size=0.02",
        linewidth=1,
        edgecolor=CARD_EDGE,
        facecolor="#F8FAFC",
        transform=ax_header.transAxes,
    )
    ax_header.add_patch(header_box)

    ax_header.text(0.02, 0.58, logo_text, fontsize=24, weight="bold", color=PRIMARY, va="center", ha="left", transform=ax_header.transAxes)
    ax_header.text(0.13, 0.62, title, fontsize=24, weight="bold", color=TITLE, va="center", ha="left", transform=ax_header.transAxes)
    ax_header.text(
        0.13, 0.26,
        f"Période analysée : du {start_dt.strftime('%d/%m/%Y %H:%M')} au {end_dt.strftime('%d/%m/%Y %H:%M')}",
        fontsize=11.5, color=SUBTITLE, va="center", ha="left", transform=ax_header.transAxes
    )

    ax_header.plot([0.01, 0.99], [0.04, 0.04], color=CARD_EDGE, lw=1, transform=ax_header.transAxes)
    return fig


def add_top_kpis(fig, kpis):
    n = len(kpis)
    left, top, total_width, gap, card_h = 0.05, 0.77, 0.90, 0.015, 0.12
    card_w = (total_width - gap * (n - 1)) / n

    for i, kpi in enumerate(kpis):
        x = left + i * (card_w + gap)
        ax = fig.add_axes([x, top, card_w, card_h])
        ax.set_axis_off()

        card = patches.FancyBboxPatch(
            (0, 0), 1, 1,
            boxstyle="round,pad=0.012,rounding_size=0.03",
            linewidth=1,
            edgecolor=CARD_EDGE,
            facecolor=CARD_BG,
            transform=ax.transAxes,
        )
        ax.add_patch(card)
        ax.add_patch(
            patches.FancyBboxPatch(
                (0, 0.92), 1, 0.08,
                boxstyle="round,pad=0.0,rounding_size=0.02",
                linewidth=0,
                facecolor=kpi.get("color", PRIMARY),
                transform=ax.transAxes,
            )
        )

        ax.text(0.05, 0.62, str(kpi["value"]), fontsize=18, weight="bold", color=TITLE, ha="left", va="center", transform=ax.transAxes)
        ax.text(0.05, 0.34, kpi["title"], fontsize=11, weight="bold", color=TITLE, ha="left", va="center", transform=ax.transAxes)
        ax.text(0.05, 0.14, kpi.get("subtitle", ""), fontsize=9, color=SUBTITLE, ha="left", va="center", transform=ax.transAxes)


def add_bottom_kpis(fig, kpis):
    n = len(kpis)
    left, bottom, total_width, gap, card_h = 0.05, 0.05, 0.90, 0.015, 0.10
    card_w = (total_width - gap * (n - 1)) / n

    for i, kpi in enumerate(kpis):
        x = left + i * (card_w + gap)
        ax = fig.add_axes([x, bottom, card_w, card_h])
        ax.set_axis_off()

        card = patches.FancyBboxPatch(
            (0, 0), 1, 1,
            boxstyle="round,pad=0.012,rounding_size=0.03",
            linewidth=1,
            edgecolor=CARD_EDGE,
            facecolor="#F8FAFC",
            transform=ax.transAxes,
        )
        ax.add_patch(card)

        ax.text(0.05, 0.62, str(kpi["value"]), fontsize=13, weight="bold", color=kpi.get("color", PRIMARY), ha="left", va="center", transform=ax.transAxes)
        ax.text(0.05, 0.34, kpi["title"], fontsize=10, weight="bold", color=TITLE, ha="left", va="center", transform=ax.transAxes)
        ax.text(0.05, 0.14, kpi.get("subtitle", ""), fontsize=8.5, color=SUBTITLE, ha="left", va="center", transform=ax.transAxes)


def style_table(tbl, header_color=PRIMARY):
    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor("#E5E7EB")
        cell.set_linewidth(0.6)
        if row == 0:
            cell.set_text_props(weight="bold", color="white")
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor("#FFFFFF" if row % 2 else "#F8FAFC")


def add_small_kpi_card(fig, x, y, w, h, title, value, color):
    ax_card = fig.add_axes([x, y, w, h])
    ax_card.axis("off")
    card = patches.FancyBboxPatch(
        (0, 0), 1, 1,
        boxstyle="round,pad=0.018,rounding_size=0.08",
        linewidth=1.2,
        edgecolor="#E5E7EB",
        facecolor="white",
        transform=ax_card.transAxes,
    )
    ax_card.add_patch(card)
    ax_card.add_patch(
        patches.FancyBboxPatch(
            (0, 0), 0.035, 1,
            boxstyle="round,pad=0.018,rounding_size=0.08",
            linewidth=0,
            facecolor=color,
            transform=ax_card.transAxes,
        )
    )
    ax_card.text(0.10, 0.62, value, ha="left", va="center", fontsize=13, fontweight="bold", color=color, transform=ax_card.transAxes)
    ax_card.text(0.10, 0.30, title, ha="left", va="center", fontsize=9, fontweight="bold", color="#111827", transform=ax_card.transAxes)


def get_reactivity_level(val):
    if pd.isna(val):
        return "N/A", "#9CA3AF", "white"
    if val > 90:
        return "> 90 sec", "#111111", "white"
    if val > 60:
        return "60 à 90 sec", "#d62828", "white"
    if val >= 45:
        return "45 à 60 sec", "#f5a623", "black"
    return "< 45 sec", "#2e7d32", "white"


def add_reactivity_cards_page(fig, df_zone):
    ax = fig.add_axes([0.05, 0.18, 0.90, 0.53])
    ax.set_facecolor(BG)
    ax.axis("off")

    if df_zone.empty:
        ax.text(0.5, 0.5, "Aucune donnée disponible pour les temps de réactivité", ha="center", va="center", fontsize=14, color=SUBTITLE, transform=ax.transAxes)
        return

    df_zone = df_zone.head(12).copy()
    n = len(df_zone)
    ncols = 3
    nrows = math.ceil(n / ncols)

    ax.set_xlim(0, ncols)
    ax.set_ylim(0, nrows + 0.6)

    legend_items = [
        ("< 45 sec", "#2e7d32"),
        ("45 à 60 sec", "#f5a623"),
        ("60 à 90 sec", "#d62828"),
        ("> 90 sec", "#111111"),
    ]

    x_positions = [0.10, 0.34, 0.58, 0.80]
    for (label, color), x in zip(legend_items, x_positions):
        fig.text(x - 0.012, 0.735, "●", color=color, fontsize=12, ha="right", va="center")
        fig.text(x, 0.735, label, color=SUBTITLE, fontsize=10, ha="left", va="center")

    card_w, card_h, x_pad, y_top_offset = 0.84, 0.74, 0.08, 0.02

    def shorten_text(txt, max_len=18):
        txt = str(txt)
        return txt if len(txt) <= max_len else txt[:max_len - 3] + "..."

    for idx, row in df_zone.reset_index(drop=True).iterrows():
        r = idx // ncols
        c = idx % ncols
        x = c + x_pad
        y = nrows - r + y_top_offset

        level_label, level_color, value_color = get_reactivity_level(row["mean"])
        zone_name = shorten_text(row[ZONE_COL], max_len=18)

        ax.add_patch(
            patches.FancyBboxPatch(
                (x + 0.015, y - card_h - 0.015), card_w, card_h,
                boxstyle="round,pad=0.018,rounding_size=0.04",
                linewidth=0,
                facecolor="#D9DDE3",
                alpha=0.35,
            )
        )
        ax.add_patch(
            patches.FancyBboxPatch(
                (x, y - card_h), card_w, card_h,
                boxstyle="round,pad=0.018,rounding_size=0.04",
                linewidth=1.2,
                edgecolor="#E5E7EB",
                facecolor=CARD_BG,
            )
        )
        ax.add_patch(
            patches.FancyBboxPatch(
                (x, y - card_h), 0.035, card_h,
                boxstyle="round,pad=0.008,rounding_size=0.03",
                linewidth=0,
                facecolor=level_color,
            )
        )

        ax.text(x + 0.07, y - 0.10, zone_name, ha="left", va="top", fontsize=10.2, fontweight="bold", color=TEXT)
        ax.text(x + 0.07, y - 0.40, f"{row['mean']:.1f} sec", ha="left", va="center", fontsize=18, fontweight="bold", color=level_color)
        ax.text(x + 0.07, y - 0.60, f"Volume alarmes : {int(row['count'])}", ha="left", va="center", fontsize=9.3, color=SUBTITLE)

        badge_w, badge_h = 0.24, 0.11
        ax.add_patch(
            patches.FancyBboxPatch(
                (x + card_w - badge_w - 0.04, y - 0.16),
                badge_w, badge_h,
                boxstyle="round,pad=0.015,rounding_size=0.03",
                linewidth=0,
                facecolor=level_color,
            )
        )
        ax.text(x + card_w - badge_w / 2 - 0.04, y - 0.105, level_label, ha="center", va="center", fontsize=7.4, fontweight="bold", color=value_color)


def build_result_linking(df_period: pd.DataFrame, df_mask: pd.DataFrame):
    interv = df_period.copy().rename(columns={EQUIPEMENT_COL: "Equipment"})
    alarms = df_mask.copy()

    if DATE_FIN_COL in interv.columns:
        interv["fin_intervention"] = pd.to_datetime(interv[DATE_FIN_COL], errors="coerce")
    else:
        interv["fin_intervention"] = pd.NaT

    alarms["equip_key"] = alarms["Equipment"].astype(str).str.upper().str.strip()
    interv["equip_key"] = interv["Equipment"].astype(str).str.upper().str.strip()

    alarms["On Time"] = pd.to_datetime(alarms["On Time"], errors="coerce")
    alarms["Off Time"] = pd.to_datetime(alarms["Off Time"], errors="coerce")
    interv["debut_intervention"] = pd.to_datetime(interv[DATE_COL], errors="coerce")

    by_cols = ["equip_key", ZONE_GLOBALE_COL, ZONE_COL]
    alarms_sorted = alarms.dropna(subset=["On Time"]).sort_values("On Time")
    interv_sorted = interv.dropna(subset=["debut_intervention"]).sort_values("debut_intervention")

    result = pd.merge_asof(
        interv_sorted,
        alarms_sorted,
        left_on="debut_intervention",
        right_on="On Time",
        by=by_cols,
        direction="backward",
        tolerance=pd.Timedelta("30min"),
    )

    result["temps_reactivite"] = result["debut_intervention"] - result["On Time"]
    result["temps_reactivite_secondes"] = result["temps_reactivite"].dt.total_seconds()

    return result


def make_figures(kpi, df_period, result, start_dt, end_dt):
    figs = []

    # Page 1 - Récap
    fig = page_template("Rapport journalier — Récapitulatif", start_dt, end_dt)
    add_top_kpis(fig, [
        {"title": "Interventions", "value": kpi["total_pannes"], "subtitle": "Volume total", "color": PRIMARY},
        {"title": "Équipements", "value": kpi["nb_equipements"], "subtitle": "Distincts", "color": ACCENT},
        {"title": "Opérateurs", "value": kpi["nb_operateurs"], "subtitle": "Impliqués", "color": SUCCESS},
        {"title": "Sous-zones", "value": kpi["nb_sous_zones"], "subtitle": "Impactées", "color": DANGER},
    ])

    ax_left = fig.add_axes([0.05, 0.22, 0.55, 0.48])
    ax_left.set_axis_off()
    ax_left.add_patch(patches.FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.015,rounding_size=0.03", linewidth=1, edgecolor=CARD_EDGE, facecolor=CARD_BG, transform=ax_left.transAxes))
    ax_left.text(0.05, 0.90, "Synthèse opérationnelle", fontsize=18, weight="bold", color=TITLE, transform=ax_left.transAxes)

    resume_items = [
        ("Shift dominant", kpi["top_shift"], f'{kpi["top_shift_nb"]} interventions', PRIMARY),
        ("Équipement critique", kpi["top_equip"], f'{kpi["top_equip_nb"]} interventions', DANGER),
        ("Sous-zone la plus impactée", kpi["top_zone"], f'{kpi["top_zone_nb"]} interventions', ACCENT),
        ("Heure la plus chargée", f'{kpi["hour_peak"]:02d}h', f'{kpi["hour_peak_nb"]} interventions', SUCCESS),
    ]
    for (label, value, subtitle, color), y in zip(resume_items, [0.72, 0.53, 0.34, 0.15]):
        ax_left.text(0.05, y + 0.07, label, fontsize=11, color=SUBTITLE, transform=ax_left.transAxes)
        ax_left.text(0.05, y, str(value), fontsize=18, weight="bold", color=TITLE, transform=ax_left.transAxes)
        ax_left.text(0.60, y, subtitle, fontsize=12, color=color, transform=ax_left.transAxes)

    ax_right = fig.add_axes([0.63, 0.22, 0.32, 0.48])
    ax_right.set_axis_off()
    ax_right.add_patch(patches.FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.015,rounding_size=0.03", linewidth=1, edgecolor=CARD_EDGE, facecolor=CARD_BG, transform=ax_right.transAxes))
    ax_right.text(0.07, 0.90, "Lecture rapide", fontsize=18, weight="bold", color=TITLE, transform=ax_right.transAxes)
    ax_right.text(0.07, 0.74, "Point d’attention principal", fontsize=11, color=SUBTITLE, transform=ax_right.transAxes)
    ax_right.text(0.07, 0.63, str(kpi["top_equip"]), fontsize=20, weight="bold", color=DANGER, transform=ax_right.transAxes)
    ax_right.text(0.07, 0.54, f'{kpi["top_equip_nb"]} interventions recensées', fontsize=12, color=TITLE, transform=ax_right.transAxes)
    ax_right.text(
        0.07, 0.36,
        "Le rapport met en évidence les zones et équipements\n"
        "nécessitant une vigilance renforcée. La concentration\n"
        "des événements permet de prioriser plus rapidement\n"
        "les actions correctives et le suivi terrain.",
        fontsize=11.2, color=SUBTITLE, linespacing=1.5, transform=ax_right.transAxes,
    )
    add_bottom_kpis(fig, [
        {"title": "Top shift", "value": kpi["top_shift"], "subtitle": f'{kpi["top_shift_nb"]} interventions', "color": PRIMARY},
        {"title": "Top zone", "value": kpi["top_zone"], "subtitle": f'{kpi["top_zone_nb"]} interventions', "color": ACCENT},
        {"title": "Top équipement", "value": kpi["top_equip"], "subtitle": f'{kpi["top_equip_nb"]} interventions', "color": DANGER},
    ])
    figs.append(("Récapitulatif", fig))

    # Page 2 - Défauts
    fig = page_template("Défauts majeurs", start_dt, end_dt)
    defect_counts = kpi["defect_counts"]
    defect_counts_plot = defect_counts.copy()
    if len(defect_counts_plot) > 6:
        top_defects = defect_counts_plot.head(5)
        autres = defect_counts_plot.iloc[5:].sum()
        defect_counts_plot = pd.concat([top_defects, pd.Series({"Autres": autres})])

    add_top_kpis(fig, [
        {"title": "Total interventions", "value": kpi["total_pannes"], "subtitle": "Sur la période", "color": PRIMARY},
        {"title": "Nb défauts cumulés", "value": int(defect_counts.sum()) if not defect_counts.empty else 0, "subtitle": "Technique + Opérationnel", "color": DANGER},
        {"title": "Défaut principal", "value": defect_counts.index[0] if not defect_counts.empty else "-", "subtitle": f"{int(defect_counts.iloc[0])} occurrences" if not defect_counts.empty else "", "color": ACCENT},
        {"title": "Nb équipements", "value": kpi["nb_equipements"], "subtitle": "Équipements distincts", "color": SUCCESS},
    ])

    ax = fig.add_axes([0.07, 0.20, 0.56, 0.52], facecolor=CARD_BG)
    ax_leg = fig.add_axes([0.66, 0.22, 0.28, 0.48], facecolor=CARD_BG)
    ax.set_aspect("equal")
    ax.set_xticks([])
    ax.set_yticks([])
    ax_leg.axis("off")

    pie_labels = defect_counts_plot.index.tolist() or ["Aucune donnée"]
    pie_sizes = defect_counts_plot.values.tolist() or [1]
    colors = ["#DBEAFE", "#BFDBFE", "#93C5FD", "#60A5FA", "#3B82F6", "#1D4ED8", "#1E40AF", "#172554"]

    wedges, _, autotexts = ax.pie(
        pie_sizes,
        startangle=90,
        counterclock=False,
        colors=colors[:len(pie_sizes)],
        labels=None,
        autopct=autopct_count(pie_sizes),
        wedgeprops=dict(width=0.38, edgecolor="white"),
        pctdistance=0.72,
    )
    ax.text(0, 0.08, f"{sum(pie_sizes)}", ha="center", va="center", fontsize=22, weight="bold", color=TITLE)
    ax.text(0, -0.12, "défauts", ha="center", va="center", fontsize=10, color=SUBTITLE)
    for t in autotexts:
        t.set_color(TITLE)
        t.set_fontsize(9)
        t.set_weight("bold")
    ax_leg.legend(wedges, [wrap_text(x, 26) for x in pie_labels], loc="center left", frameon=False, fontsize=10)
    figs.append(("Défauts majeurs", fig))

    # Page 3 - Sous-zones
    fig = page_template("Sous-zones les plus impactées", start_dt, end_dt)
    add_top_kpis(fig, [
        {"title": "Sous-zone principale", "value": kpi["top_zone"], "subtitle": f'{kpi["top_zone_nb"]} interventions', "color": DANGER},
        {"title": "Nb sous-zones", "value": kpi["nb_sous_zones"], "subtitle": "Sous-zones touchées", "color": PRIMARY},
        {"title": "Total interventions", "value": kpi["total_pannes"], "subtitle": "Sur la période", "color": ACCENT},
        {"title": "Shift dominant", "value": kpi["top_shift"], "subtitle": f'{kpi["top_shift_nb"]} interventions', "color": SUCCESS},
    ])
    ax = fig.add_axes([0.08, 0.20, 0.84, 0.52], facecolor=CARD_BG)
    zone_plot = kpi["zone_counts"].sort_values(ascending=True).tail(10)
    bars = ax.bar(zone_plot.index, zone_plot.values, color=PRIMARY, alpha=0.88)
    ax.set_title("Top sous-zones en défaut", loc="left", fontsize=16, color=TITLE, weight="bold")
    ax.set_xlabel("Nombre d'interventions")
    ax.set_ylabel("")
    ax.grid(axis="x", alpha=0.25)
    ax.tick_params(axis="x", rotation=35)
    ax.bar_label(bars, padding=4, fontsize=10, color=TITLE)
    figs.append(("Sous-zones", fig))

    # Page 4 - Equipements
    fig = page_template("Équipements critiques", start_dt, end_dt)
    add_top_kpis(fig, [
        {"title": "Équipement critique", "value": kpi["top_equip"], "subtitle": f'{kpi["top_equip_nb"]} interventions', "color": DANGER},
        {"title": "Nb équipements", "value": kpi["nb_equipements"], "subtitle": "Équipements distincts", "color": PRIMARY},
        {"title": "Total interventions", "value": kpi["total_pannes"], "subtitle": "Sur la période", "color": ACCENT},
        {"title": "Heure de pic", "value": f'{kpi["hour_peak"]:02d}h', "subtitle": f'{kpi["hour_peak_nb"]} interventions', "color": SUCCESS},
    ])
    ax = fig.add_axes([0.10, 0.20, 0.78, 0.52], facecolor=CARD_BG)
    eq_plot = kpi["equip_counts"].sort_values(ascending=True).tail(10)
    rate_plot = ((eq_plot / kpi["total_pannes"]) * 100).round(1) if kpi["total_pannes"] else eq_plot
    bars = ax.bar(eq_plot.index, eq_plot.values, color=DANGER, alpha=0.85)
    ax.set_title("Top équipements par volume d'interventions", loc="left", fontsize=16, color=TITLE, weight="bold")
    ax.set_xlabel("Nombre d'interventions")
    ax.set_ylabel("")
    ax.grid(axis="x", alpha=0.25)
    ax.tick_params(axis="x", rotation=35)
    labels = [f"{int(count)} | {rate:.1f}%" for count, rate in zip(eq_plot.values, rate_plot.values)]
    ax.bar_label(bars, labels=labels, padding=4, fontsize=10, color=TITLE)
    figs.append(("Équipements", fig))

    # Page 5 - Heatmap horaire
    fig = page_template("Heatmap horaire", start_dt, end_dt)
    add_top_kpis(fig, [
        {"title": "Heure de pic", "value": f'{kpi["hour_peak"]:02d}h', "subtitle": f'{kpi["hour_peak_nb"]} interventions', "color": DANGER},
        {"title": "Shift dominant", "value": kpi["top_shift"], "subtitle": f'{kpi["top_shift_nb"]} interventions', "color": PRIMARY},
        {"title": "Sous-zone principale", "value": kpi["top_zone"], "subtitle": f'{kpi["top_zone_nb"]} interventions', "color": ACCENT},
        {"title": "Équipement critique", "value": kpi["top_equip"], "subtitle": f'{kpi["top_equip_nb"]} interventions', "color": SUCCESS},
    ])
    ax = fig.add_axes([0.10, 0.20, 0.78, 0.52], facecolor=CARD_BG)
    sns.heatmap(
        kpi["heatmap_hour"],
        cmap="Reds",
        linewidths=0.3,
        linecolor="#E5E7EB",
        cbar=True,
        annot=True if kpi["heatmap_hour"].shape[0] <= 14 and kpi["heatmap_hour"].shape[1] <= 8 else False,
        fmt="g",
        ax=ax,
    )
    ax.set_title("Concentration horaire des interventions par shift", loc="left", fontsize=16, color=TITLE, weight="bold")
    ax.set_xlabel("")
    ax.set_ylabel("Heure")
    figs.append(("Heatmap horaire", fig))

    # Page 6 - Heatmap opérateurs / sous-zones
    fig = page_template("Heatmap opérateurs / sous-zones", start_dt, end_dt)
    top_operateur = kpi["operateur_summary"].iloc[0][OPERATEUR_COL] if len(kpi["operateur_summary"]) > 0 else "-"
    top_operateur_nb = int(kpi["operateur_summary"].iloc[0]["Nombre_interventions"]) if len(kpi["operateur_summary"]) > 0 else 0
    add_top_kpis(fig, [
        {"title": "Top opérateur", "value": top_operateur, "subtitle": f"{top_operateur_nb} interventions", "color": PRIMARY},
        {"title": "Nb opérateurs", "value": kpi["operateur_zone_summary"][OPERATEUR_COL].nunique(), "subtitle": "Opérateurs distincts", "color": ACCENT},
        {"title": "Nb sous-zones", "value": kpi["operateur_zone_summary"][ZONE_COL].nunique(), "subtitle": "Sous-localisations couvertes", "color": DANGER},
        {"title": "Total interventions", "value": kpi["total_pannes"], "subtitle": "Sur la période", "color": SUCCESS},
    ])
    ax = fig.add_axes([0.08, 0.20, 0.84, 0.52], facecolor=CARD_BG)
    if kpi["ops_heatmap"].empty:
        ax.text(0.5, 0.5, "Aucune donnée disponible pour la heatmap", ha="center", va="center", fontsize=14, color=SUBTITLE, transform=ax.transAxes)
        ax.set_xticks([])
        ax.set_yticks([])
    else:
        sns.heatmap(kpi["ops_heatmap"], cmap="Blues", linewidths=0.3, linecolor="#E5E7EB", cbar=True, annot=True, fmt=".0f", ax=ax)
        ax.set_title("Interventions par opérateur et sous-localisation", loc="left", fontsize=16, color=TITLE, weight="bold")
        ax.set_xlabel("Sous-localisation")
        ax.set_ylabel("Opérateur")
        ax.tick_params(axis="x", rotation=35, labelsize=9)
        ax.tick_params(axis="y", rotation=0, labelsize=9)
    figs.append(("Opérateurs / sous-zones", fig))

    # Page 7 - Temps de réactivité par zone
    fig = page_template("Temps de réactivité par zone", start_dt, end_dt)
    add_top_kpis(fig, [
        {"title": "Temps moyen global", "value": f'{kpi["mean_reactivity_global"]:.1f} sec', "subtitle": "Toutes zones confondues", "color": PRIMARY},
        {"title": "Zone la plus critique", "value": kpi["worst_zone_reactivity"], "subtitle": f'{kpi["worst_zone_reactivity_val"]:.1f} sec', "color": DANGER},
        {"title": "Zones > 90 sec", "value": kpi["nb_critical_90"], "subtitle": "Criticité maximale", "color": "#111111"},
        {"title": "Zones 60 à 90 sec", "value": kpi["nb_critical_60_90"], "subtitle": "À surveiller", "color": ACCENT},
    ])
    add_reactivity_cards_page(fig, kpi["df_zone"])
    figs.append(("Réactivité par zone", fig))

    # Page 8 - Temps de réactivité opérateurs / zones
    fig = page_template("Temps de réactivité opérateurs / zones", start_dt, end_dt)
    ax = fig.add_axes([0.08, 0.20, 0.84, 0.52], facecolor=CARD_BG)

    df_valid = result[result["On Time"].notna()].copy() if "On Time" in result.columns else pd.DataFrame()
    if not df_valid.empty and "Duration_seconds" in df_valid.columns:
        df_valid = df_valid[(df_valid["Duration_seconds"] >= 0) & (df_valid["Duration_seconds"] <= 600)].copy()

    if not df_valid.empty:
        heatmap2 = df_valid.pivot_table(index=OPERATEUR_COL, columns=ZONE_COL, values="Duration_seconds", aggfunc="mean")
        heatmap2 = heatmap2.replace(0, np.nan).dropna(how="all", axis=0).dropna(how="all", axis=1)

        if not heatmap2.empty:
            nb_critique = int((heatmap2 > 90).sum().sum())
            nb_mauvais = int(((heatmap2 > 60) & (heatmap2 <= 90)).sum().sum())
            add_small_kpi_card(fig, x=0.42, y=0.78, w=0.18, h=0.075, title="Cas critiques", value=f"{nb_critique} > 90s", color="#2E2E2E")
            add_small_kpi_card(fig, x=0.62, y=0.78, w=0.20, h=0.075, title="Cas à surveiller", value=f"{nb_mauvais} entre 60–90s", color="#C75C5C")

            heatmap2 = heatmap2.loc[heatmap2.mean(axis=1).sort_values().index]
            colors = ["#4CAF7A", "#D4B24C", "#C75C5C", "#2E2E2E"]
            bounds = [0, 45, 60, 90, 600]
            cmap = ListedColormap(colors)
            norm = BoundaryNorm(bounds, len(colors))

            hm2 = sns.heatmap(heatmap2, cmap=cmap, norm=norm, linewidths=0, annot=False, cbar=True, ax=ax, mask=heatmap2.isna())
            for i in range(heatmap2.shape[0]):
                for j in range(heatmap2.shape[1]):
                    val = heatmap2.iloc[i, j]
                    if pd.isna(val):
                        continue
                    ax.text(j + 0.5, i + 0.5, f"{val:.1f}", ha="center", va="center", fontsize=8, color="white" if val >= 60 else "#111827")

            ax.set_title("Temps moyen de réactivité par opérateur et zone", loc="left", fontsize=16, weight="bold")
            ax.set_xlabel("Zone")
            ax.set_ylabel("Opérateur")
            ax.tick_params(axis="x", rotation=35)
            ax.tick_params(axis="y", rotation=0)

            cbar = hm2.collections[0].colorbar
            cbar.set_ticks([22, 52, 75, 120])
            cbar.set_ticklabels(["Bon (0-45)", "À améliorer (45-60)", "À surveiller (60-90)", "Critique (>90)"])
        else:
            ax.text(0.5, 0.5, "Aucune donnée disponible", ha="center", va="center", fontsize=14, color=SUBTITLE, transform=ax.transAxes)
    else:
        ax.text(0.5, 0.5, "Aucune alarme liée aux interventions", ha="center", va="center", fontsize=14, color=SUBTITLE, transform=ax.transAxes)

    figs.append(("Réactivité opérateurs / zones", fig))

    return figs


# ============================================================
# EXPORTS
# ============================================================
def build_excel_export(kpi, df_period, result):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        kpi["operateur_zone_summary"].to_excel(writer, sheet_name="Operateurs_zones", index=False)
        kpi["operateur_summary"].to_excel(writer, sheet_name="Operateurs_resume", index=False)
        df_period.to_excel(writer, sheet_name="Donnees_filtrees", index=False)

        kpi["shift_counts"].rename_axis("SHIFT").reset_index(name="nb_interventions").to_excel(writer, sheet_name="KPI_shift", index=False)
        kpi["zone_counts"].rename_axis("Sous_Localisation").reset_index(name="nb_interventions").to_excel(writer, sheet_name="KPI_zones", index=False)
        kpi["equip_counts"].rename_axis("Equipement").reset_index(name="nb_interventions").to_excel(writer, sheet_name="KPI_equipements", index=False)
        kpi["defect_counts"].rename_axis("Defaut").reset_index(name="nb_occurrences").to_excel(writer, sheet_name="KPI_defauts", index=False)
        kpi["hour_counts"].rename_axis("heure").reset_index(name="nb_interventions").to_excel(writer, sheet_name="KPI_horaire", index=False)
        kpi["df_zone"].to_excel(writer, sheet_name="KPI_reactivite_zones", index=False)
        result.to_excel(writer, sheet_name="Alarmes_liees", index=False)
    output.seek(0)
    return output.getvalue()


def build_pdf_export(figs):
    output = io.BytesIO()
    with PdfPages(output) as pdf:
        for _, fig in figs:
            pdf.savefig(fig, bbox_inches="tight")
    output.seek(0)
    return output.getvalue()


# ============================================================
# INTERFACE
# ============================================================
with st.sidebar:
    st.header("1. Sources de données")

    source_mode = st.radio(
        "Source interventions",
        ["URL Google Sheets", "Upload Excel"],
        horizontal=False,
    )

    default_url = "https://docs.google.com/spreadsheets/d/1gzoKfkClUx9MDkTIgx7kFma9oOwk1i6p/export?format=xlsx"
    interventions_url = None
    interventions_file = None

    if source_mode == "URL Google Sheets":
        interventions_url = st.text_input("URL export XLSX", value=default_url)
    else:
        interventions_file = st.file_uploader("Fichier interventions Excel", type=["xlsx", "xls"])

    group_file = st.file_uploader("Fichier group_f.xlsx", type=["xlsx", "xls"])
    alarm_file = st.file_uploader("Fichier alarmes Excel", type=["xlsx", "xls"])

    st.header("2. Paramètres alarmes")
    alarm_skiprows = st.number_input(
        "Nombre de lignes à ignorer dans le fichier alarmes",
        min_value=0,
        max_value=50,
        value=11,
        step=1,
        help="Dans ton script initial : pd.read_excel(fichier_test, skiprows=11)",
    )
    alarm_remove_first_cols = st.number_input(
        "Nombre de premières colonnes à supprimer dans le fichier alarmes",
        min_value=0,
        max_value=20,
        value=4,
        step=1,
        help="Dans ton script initial : da = da.iloc[:, 4:]",
    )

    generate = st.button("🚀 Générer le rapport", type="primary", use_container_width=True)


if not generate:
    st.info("Charge les fichiers dans la barre latérale, puis clique sur **Générer le rapport**.")
    st.stop()

try:
    with st.spinner("Lecture des fichiers..."):
        if source_mode == "URL Google Sheets":
            df_raw = read_excel_from_url(interventions_url)
        else:
            if interventions_file is None:
                st.error("Merci d'uploader le fichier interventions.")
                st.stop()
            df_raw = read_excel_cached(interventions_file.getvalue())

        if group_file is None:
            st.error("Merci d'uploader le fichier group_f.xlsx.")
            st.stop()

        if alarm_file is None:
            st.error("Merci d'uploader le fichier alarmes.")
            st.stop()

        group_raw = read_excel_cached(group_file.getvalue())

        alarm_raw = read_excel_cached(alarm_file.getvalue(), skiprows=int(alarm_skiprows))
        if int(alarm_remove_first_cols) > 0:
            alarm_raw = alarm_raw.iloc[:, int(alarm_remove_first_cols):].copy()

    with st.spinner("Préparation des données..."):
        _, df_period, start_dt, end_dt = prepare_interventions(df_raw, group_raw)
        df_mask, df_filtre, df_grouped = prepare_alarms(alarm_raw, group_raw)
        kpi = compute_kpis(df_period, df_grouped)
        result = build_result_linking(df_period, df_mask)

    st.success(
        f"Rapport généré pour la période : "
        f"{start_dt.strftime('%d/%m/%Y %H:%M')} → {end_dt.strftime('%d/%m/%Y %H:%M')}"
    )

    # KPIs page Streamlit
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Interventions", kpi["total_pannes"])
    c2.metric("Équipements distincts", kpi["nb_equipements"])
    c3.metric("Opérateurs", kpi["nb_operateurs"])
    c4.metric("Sous-zones", kpi["nb_sous_zones"])

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Top shift", kpi["top_shift"], f'{kpi["top_shift_nb"]} interventions')
    c6.metric("Top zone", kpi["top_zone"], f'{kpi["top_zone_nb"]} interventions')
    c7.metric("Top équipement", kpi["top_equip"], f'{kpi["top_equip_nb"]} interventions')
    c8.metric("Heure pic", f'{kpi["hour_peak"]:02d}h', f'{kpi["hour_peak_nb"]} interventions')

    with st.spinner("Création des visuels et exports..."):
        figs = make_figures(kpi, df_period, result, start_dt, end_dt)
        pdf_bytes = build_pdf_export(figs)
        excel_bytes = build_excel_export(kpi, df_period, result)

    st.subheader("📥 Téléchargements")
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Télécharger le rapport PDF",
        data=pdf_bytes,
        file_name=f"rapport_journalier_aesma_{start_dt.strftime('%Y-%m-%d')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
    d2.download_button(
        "⬇️ Télécharger les données Excel",
        data=excel_bytes,
        file_name=f"rapport_journalier_aesma_{start_dt.strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.subheader("👁️ Aperçu visuel du rapport")
    tabs = st.tabs([name for name, _ in figs])
    for tab, (name, fig) in zip(tabs, figs):
        with tab:
            st.pyplot(fig, clear_figure=False, use_container_width=True)

    st.subheader("📄 Aperçu des données filtrées")
    st.dataframe(df_period.head(200), use_container_width=True)

    st.subheader("🔗 Alarmes liées aux interventions")
    st.dataframe(result.head(200), use_container_width=True)

except Exception as e:
    st.error("Une erreur est survenue pendant la génération du rapport.")
    st.exception(e)
