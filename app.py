import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import smtplib
import requests
import json
import re
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, date

try:
    import anthropic as _anthropic_lib
    _ANTHROPIC_AVAILABLE = True
except ImportError:
    _ANTHROPIC_AVAILABLE = False

st.set_page_config(
    page_title="Campaign Tracker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# THEME CONSTANTS
# ─────────────────────────────────────────────
COLOR_OK      = "#22c55e"
COLOR_WARN    = "#f59e0b"
COLOR_DANGER  = "#ef4444"
COLOR_INFO    = "#3b82f6"
COLOR_PRIMARY = "#6366f1"

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.6rem !important; }
.alert-box { border-radius: 8px; padding: 12px 16px; margin-bottom: 8px; font-size: 0.9rem; }
.alert-danger  { background: #fef2f2; border-left: 4px solid #ef4444; color: #7f1d1d; }
.alert-warning { background: #fffbeb; border-left: 4px solid #f59e0b; color: #78350f; }
.alert-info    { background: #eff6ff; border-left: 4px solid #3b82f6; color: #1e3a5f; }
.alert-success { background: #f0fdf4; border-left: 4px solid #22c55e; color: #14532d; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# DATA UTILITIES
# ─────────────────────────────────────────────

EXCEL_ERRORS = {"#div/0!", "#valeur!", "#value!", "#n/a!", "#ref!", "#name?", "finie", "nan", "none", ""}

def clean_currency(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip().lower()
    if s in EXCEL_ERRORS:
        return np.nan
    s = re.sub(r"[€\s\xa0 ]", "", str(val))
    s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s)
    except Exception:
        return np.nan


def clean_percentage(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip().lower()
    if s in EXCEL_ERRORS:
        return np.nan
    s = s.replace("%", "").replace(",", ".").strip()
    try:
        return float(s)
    except Exception:
        return np.nan


def clean_number(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip().lower()
    if s in EXCEL_ERRORS:
        return np.nan
    s = re.sub(r"\s", "", str(val))
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan


def parse_date(val):
    if pd.isna(val) or str(val).strip() == "":
        return pd.NaT
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
        try:
            return pd.to_datetime(str(val).strip(), format=fmt)
        except Exception:
            continue
    return pd.NaT


# Keywords used to detect a group-header row (should be skipped)
_GROUP_HEADER_KEYWORDS = {
    "cr", "ation", "création", "creation", "quipe", "équipe", "equipe",
    "informations", "delivery", "marge", "vue d", "mapping",
}

# Fuzzy mapping: fragments to look for (lowercase, no accents) → canonical name
_COL_FUZZY = {
    "campagne":        "Campagne",
    "geo":             "Géo",
    "entit":           "Entité",
    "stats":           "Stats",
    "levier":          "Levier",
    "statut":          "Statut",
    "affiliat":        "Affiliate",
    "account":         "Account",
    "budget":          "Budget",
    "but":             "Budget",          # "Budget" with encoding glitch
    "d but":           "Début",
    "debut":           "Début",
    "fin":             "Fin",
    "mod le":          "Modèle",
    "model":           "Modèle",
    "r m. ann":        "Rém_Ann",
    "rem. ann":        "Rém_Ann",
    "r m. net":        "Rém_NET",
    "rem. net":        "Rém_NET",
    "r m. dit":        "Rém_Éditeur",
    "rem. edit":       "Rém_Éditeur",
    "marge":           "Marge_pct",
    "objectif":        "Objectif",
    "budget restant":  "Budget_restant",
    "volume restant":  "Volume_restant",
    "volume r":        "Volume_réalisé",
    "ca interne":      "CA_interne",
    "ca externe":      "CA_externe",
    "% budget":        "Pct_Budget",
    "% temps":         "Pct_Temps",
    "annonceur":       "Annonceur",
    "march":           "Marché",
    "verticale":       "Verticale",
    "marge €":         "Marge_€",
    "marge e":         "Marge_€",
}

# Positional fallback (column index → canonical name) — matches the sample CSV layout
_COL_POSITIONAL = [
    "Campagne", "Géo", "Entité", "Stats", "Levier", "Statut",
    "Affiliate", "Account", "Budget", "Début", "Fin", "Modèle",
    "Rém_Ann", "Rém_NET", "Rém_Éditeur", "Marge_pct",
    "Objectif", "Objectif_date", "Budget_restant",
    "Objectif_date2", "Volume_restant", "Volume_réalisé",
    "CA_interne", "CA_externe", "Pct_Budget", "Pct_Temps",
    "Annonceur", "Marché", "Campagne_map", "Verticale", "Marge_€",
]

_STATUS_VALUES = {"active", "set-up", "setup", "budget fait", "en pause",
                  "pause", "finie", "budget non-atteint", "budget non atteint"}


def _strip_accents(s: str) -> str:
    import unicodedata
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def _normalize_col(s: str) -> str:
    """Lowercase, strip accents, collapse whitespace."""
    return re.sub(r"\s+", " ", _strip_accents(str(s)).lower().strip())


def _is_group_header_row(row: pd.Series) -> bool:
    """True if the row looks like a merged-cell group header (should be skipped)."""
    non_empty = [str(v).strip() for v in row if str(v).strip() not in ("", "nan")]
    if len(non_empty) == 0:
        return True
    # Group header rows tend to have many empty cells and keyword-like values
    empty_ratio = 1 - len(non_empty) / len(row)
    first = _normalize_col(non_empty[0]) if non_empty else ""
    has_keyword = any(k in first for k in _GROUP_HEADER_KEYWORDS)
    return empty_ratio > 0.5 and has_keyword


def _fuzzy_rename(df: pd.DataFrame) -> pd.DataFrame:
    """Try to rename columns using fuzzy keyword matching, then positional fallback."""
    rename_map = {}
    already_used = set()

    for col in df.columns:
        norm = _normalize_col(col)
        matched = None
        # Try fuzzy match
        for fragment, canonical in _COL_FUZZY.items():
            if fragment in norm and canonical not in already_used:
                matched = canonical
                break
        if matched:
            rename_map[col] = matched
            already_used.add(matched)

    df = df.rename(columns=rename_map)

    # Positional fallback for columns still not matched
    positional_used = set(df.columns)
    for i, canonical in enumerate(_COL_POSITIONAL):
        if canonical in df.columns:
            continue  # already renamed
        # Find the i-th column that hasn't been renamed to a canonical name yet
        unnamed_cols = [c for c in df.columns
                        if c not in set(_COL_POSITIONAL) and c not in rename_map.values()]
        if unnamed_cols:
            df = df.rename(columns={unnamed_cols[0]: canonical})

    return df


def _auto_skiprows(raw_bytes: bytes, encoding: str, sep: str, is_excel: bool) -> int:
    """Detect whether the file starts with a group-header row that should be skipped."""
    try:
        if is_excel:
            peek = pd.read_excel(io.BytesIO(raw_bytes), nrows=3,
                                 dtype=str, header=None)
        else:
            peek = pd.read_csv(io.BytesIO(raw_bytes), sep=sep, nrows=3,
                               encoding=encoding, dtype=str,
                               header=None, on_bad_lines="skip")
        if len(peek) >= 2 and _is_group_header_row(peek.iloc[0]):
            return 1   # skip the group-header row; row 1 has real column names
        return 0       # row 0 already has real column names (or is data)
    except Exception:
        return 1       # safe default for the known sample format


@st.cache_data(show_spinner="Chargement des données…")
def load_data(raw_bytes: bytes, filename: str) -> pd.DataFrame:
    is_excel = filename.lower().endswith((".xlsx", ".xls"))
    df = None
    encoding_used = "cp1252"

    for enc in ["cp1252", "iso-8859-1", "utf-8-sig", "utf-8"]:
        try:
            skip = _auto_skiprows(raw_bytes, enc, ";", is_excel)
            if is_excel:
                df = pd.read_excel(io.BytesIO(raw_bytes),
                                   skiprows=skip, dtype=str)
            else:
                df = pd.read_csv(
                    io.BytesIO(raw_bytes),
                    sep=";",
                    skiprows=skip,
                    encoding=enc,
                    on_bad_lines="skip",
                    dtype=str,
                )
            encoding_used = enc
            break
        except Exception:
            continue

    if df is None or df.empty:
        st.error("Impossible de lire le fichier. Vérifie qu'il est au format CSV (séparateur `;`) ou Excel.")
        return pd.DataFrame()

    # Drop fully-empty columns (sometimes artifacts of the CSV format)
    df = df.dropna(how="all", axis=1)

    # Drop leading unnamed/empty column that some exports add
    if df.columns[0] in ("", "Unnamed: 0") or str(df.columns[0]).startswith("Unnamed"):
        df = df.iloc[:, 1:]

    # Rename columns: fuzzy + positional fallback
    df = _fuzzy_rename(df)

    # Drop fully-empty rows and rows where Campagne is blank/NaN
    df = df.dropna(how="all")
    if "Campagne" not in df.columns:
        st.error("Colonne 'Campagne' introuvable. Vérifie la structure du fichier.")
        return pd.DataFrame()

    df = df[df["Campagne"].astype(str).str.strip().str.lower().isin(["", "nan"]) == False]

    # Strip text columns — always guard with `in df.columns`
    text_cols = ["Campagne", "Géo", "Entité", "Stats", "Levier", "Statut",
                 "Affiliate", "Account", "Modèle", "Annonceur", "Marché",
                 "Campagne_map", "Verticale"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace("nan", "")

    # Normalise Statut
    if "Statut" in df.columns:
        df["Statut"] = df["Statut"].str.lower().str.strip()
    else:
        st.warning("Colonne 'Statut' non détectée — les filtres de statut ne fonctionneront pas.")
        df["Statut"] = ""

    # Parse numeric/currency
    for col in ["Budget", "Rém_Ann", "Rém_NET", "Rém_Éditeur",
                "Budget_restant", "CA_interne", "CA_externe", "Marge_€"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_currency)

    for col in ["Marge_pct", "Pct_Budget", "Pct_Temps"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_percentage)

    for col in ["Objectif", "Objectif_date", "Objectif_date2",
                "Volume_restant", "Volume_réalisé"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)

    for col in ["Début", "Fin"]:
        if col in df.columns:
            df[col] = df[col].apply(parse_date)

    return df


# ─────────────────────────────────────────────
# ALERT ENGINE
# ─────────────────────────────────────────────

def compute_alerts(df: pd.DataFrame, budget_warn_pct: float = 80.0,
                   budget_danger_pct: float = 95.0,
                   publisher_gap_pct: float = 20.0) -> dict:
    alerts = {"danger": [], "warning": [], "publisher": []}
    if df.empty:
        return alerts

    active_statuts = {"active", "set-up"}
    active = df[df["Statut"].isin(active_statuts)].copy()

    for _, row in active.iterrows():
        name = row.get("Campagne", "—")
        pct_b = row.get("Pct_Budget", np.nan)
        b_rest = row.get("Budget_restant", np.nan)

        # DANGER: over budget OR very close
        if not pd.isna(pct_b):
            if pct_b >= budget_danger_pct or (not pd.isna(b_rest) and b_rest <= 0):
                alerts["danger"].append({
                    "campagne": name,
                    "type": "Budget épuisé / dépassé",
                    "detail": f"{pct_b:.0f}% du budget utilisé — Restant : {fmt_eur(b_rest)}",
                    "pct": pct_b,
                    "budget_restant": b_rest,
                    "levier": row.get("Levier", ""),
                    "affiliate": row.get("Affiliate", ""),
                })
            elif pct_b >= budget_warn_pct:
                alerts["warning"].append({
                    "campagne": name,
                    "type": "Budget proche de la limite",
                    "detail": f"{pct_b:.0f}% du budget utilisé — Restant : {fmt_eur(b_rest)}",
                    "pct": pct_b,
                    "budget_restant": b_rest,
                    "levier": row.get("Levier", ""),
                    "affiliate": row.get("Affiliate", ""),
                })

    # PUBLISHER COMPARISON
    # For each active campaign with remaining objectives,
    # compare Rém_Éditeur with other campaigns of the same Modèle
    if "Rém_Éditeur" in df.columns and "Modèle" in df.columns:
        for _, row in active.iterrows():
            rem_ed = row.get("Rém_Éditeur", np.nan)
            modele = row.get("Modèle", "")
            vol_rest = row.get("Volume_restant", np.nan)
            name = row.get("Campagne", "—")

            if pd.isna(rem_ed) or pd.isna(vol_rest) or vol_rest <= 0 or not modele:
                continue

            # Find candidates with lower publisher cost, same model
            same_model = active[
                (active["Modèle"] == modele) &
                (active["Campagne"] != name) &
                (active["Rém_Éditeur"].notna())
            ].copy()

            cheaper = same_model[same_model["Rém_Éditeur"] < rem_ed * (1 - publisher_gap_pct / 100)]

            for _, c in cheaper.iterrows():
                saving_per_unit = rem_ed - c["Rém_Éditeur"]
                total_saving = saving_per_unit * vol_rest
                alerts["publisher"].append({
                    "campagne": name,
                    "modele": modele,
                    "current_cost": rem_ed,
                    "cheaper_campagne": c["Campagne"],
                    "cheaper_cost": c["Rém_Éditeur"],
                    "saving_per_unit": saving_per_unit,
                    "volume_restant": vol_rest,
                    "total_saving": total_saving,
                    "affiliate": c.get("Affiliate", ""),
                })

    # Deduplicate publisher alerts (keep best saving per campaign)
    seen = {}
    for a in alerts["publisher"]:
        key = a["campagne"]
        if key not in seen or a["total_saving"] > seen[key]["total_saving"]:
            seen[key] = a
    alerts["publisher"] = list(seen.values())

    return alerts


# ─────────────────────────────────────────────
# AI SUGGESTIONS (Claude)
# ─────────────────────────────────────────────

def _campaign_context(row: dict) -> str:
    """Build a compact textual summary of a campaign for the prompt."""
    def v(k, default="N/A"):
        val = row.get(k, default)
        return default if (val is None or (isinstance(val, float) and np.isnan(val))) else str(val)

    return (
        f"Nom : {v('Campagne')}\n"
        f"Annonceur : {v('Annonceur')}  |  Verticale : {v('Verticale')}  |  Marché : {v('Marché')}\n"
        f"Levier : {v('Levier')}  |  Modèle tarifaire : {v('Modèle')}  |  Géo : {v('Géo')}\n"
        f"Statut : {v('Statut')}\n"
        f"Budget total : {v('Budget')} €  |  Budget restant : {v('Budget_restant')} €  |  % Budget utilisé : {v('Pct_Budget')} %\n"
        f"% Temps écoulé : {v('Pct_Temps')} %\n"
        f"Objectif total : {v('Objectif')}  |  Volume réalisé : {v('Volume_réalisé')}  |  Volume restant : {v('Volume_restant')}\n"
        f"Rém. Annonceur : {v('Rém_Ann')} €  |  Rém. NET : {v('Rém_NET')} €  |  Rém. Éditeur : {v('Rém_Éditeur')} €\n"
        f"Marge % : {v('Marge_pct')} %  |  Marge € : {v('Marge_€')} €\n"
        f"CA interne : {v('CA_interne')} €  |  CA externe : {v('CA_externe')} €\n"
        f"Affiliate : {v('Affiliate')}  |  Account : {v('Account')}\n"
        f"Début : {v('Début')}  |  Fin : {v('Fin')}"
    )


_SYSTEM_PROMPT = """Tu es un expert en marketing performance et en gestion de campagnes digitales (affiliation, emailing, social ads, COREG, display, CPC/CPL/CPM/CPA).
Tu travailles pour une agence qui gère des campagnes pour de grands annonceurs (automobile, travel, énergie, banque/assurance, retail, telecom...).
Tes réponses sont précises, actionnables et adaptées au contexte exact de la campagne.
Tu réponds toujours en français, avec une mise en forme claire (titres, bullet points).
Sois direct et concis — pas de blabla introductif."""


def get_ai_suggestions(api_key: str, mode: str, campaign_row: dict) -> str:
    """
    mode: "strategie" | "upsell" | "depassement"
    Returns the AI-generated suggestion as a markdown string.
    """
    if not _ANTHROPIC_AVAILABLE:
        return "❌ Package `anthropic` non installé. Lance `pip3 install anthropic`."
    if not api_key:
        return "❌ Clé API Claude manquante — configure-la dans la barre latérale."

    ctx = _campaign_context(campaign_row)

    if mode == "strategie":
        user_msg = (
            f"Voici les données d'une campagne qui approche de son budget maximum :\n\n{ctx}\n\n"
            "Propose 3 à 5 idées stratégiques concrètes pour optimiser les derniers euros de budget "
            "et maximiser les résultats avant la fin de la campagne. "
            "Tiens compte du levier, du modèle tarifaire, du volume restant et de la marge actuelle."
        )
    elif mode == "upsell":
        user_msg = (
            f"Voici les données d'une campagne performante proche de son budget :\n\n{ctx}\n\n"
            "Propose 3 à 5 arguments et actions concrètes pour convaincre l'annonceur d'augmenter "
            "le budget (upsell). Inclus : chiffres à mettre en avant, canaux à activer, "
            "angles de négociation, et proposition de valeur selon le secteur de l'annonceur."
        )
    elif mode == "depassement":
        user_msg = (
            f"Voici les données d'une campagne dont le budget est dépassé :\n\n{ctx}\n\n"
            "Propose 3 à 5 solutions immédiates pour gérer ce dépassement. "
            "Inclus : actions correctives (pause, réallocation, renégociation), "
            "comment communiquer avec l'annonceur, et comment éviter que ça se reproduise. "
            "Sois pragmatique et orienté résolution rapide."
        )
    else:
        return "Mode inconnu."

    try:
        client = _anthropic_lib.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=900,
            system=_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        return response.content[0].text
    except Exception as e:
        return f"❌ Erreur API : {e}"


# ─────────────────────────────────────────────
# NOTIFICATIONS
# ─────────────────────────────────────────────

def send_slack(webhook_url: str, message: str) -> bool:
    try:
        resp = requests.post(
            webhook_url,
            json={"text": message},
            timeout=10,
        )
        return resp.status_code == 200
    except Exception:
        return False


def send_email(smtp_host: str, smtp_port: int, smtp_user: str,
               smtp_pass: str, to_email: str, subject: str, body: str) -> bool:
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = smtp_user
        msg["To"] = to_email
        msg.attach(MIMEText(body, "html"))
        with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, to_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Email error: {e}")
        return False


def build_alert_message(alerts: dict) -> str:
    lines = [f"🔔 *Campaign Tracker — Alertes du {date.today().strftime('%d/%m/%Y')}*\n"]

    if alerts["danger"]:
        lines.append(f"🔴 *{len(alerts['danger'])} campagne(s) EN DANGER*")
        for a in alerts["danger"]:
            lines.append(f"  • {a['campagne']} — {a['detail']}")

    if alerts["warning"]:
        lines.append(f"\n⚠️ *{len(alerts['warning'])} campagne(s) PROCHES DU BUDGET*")
        for a in alerts["warning"]:
            lines.append(f"  • {a['campagne']} — {a['detail']}")

    if alerts["publisher"]:
        lines.append(f"\n💡 *{len(alerts['publisher'])} opportunité(s) éditeur*")
        for a in alerts["publisher"]:
            lines.append(
                f"  • {a['campagne']} ({a['modele']}) : "
                f"{fmt_eur(a['current_cost'])} → {fmt_eur(a['cheaper_cost'])} chez {a['cheaper_campagne']}"
                f" — économie potentielle : {fmt_eur(a['total_saving'])}"
            )
    return "\n".join(lines)


# ─────────────────────────────────────────────
# FORMATTING HELPERS
# ─────────────────────────────────────────────

def fmt_eur(val):
    if pd.isna(val):
        return "—"
    return f"{val:,.2f} €".replace(",", " ").replace(".", ",")


def fmt_pct(val):
    if pd.isna(val):
        return "—"
    return f"{val:.1f}%"


def status_badge(statut: str) -> str:
    colors = {
        "active": ("#dcfce7", "#166534"),
        "set-up": ("#dbeafe", "#1e40af"),
        "budget fait": ("#fef9c3", "#713f12"),
        "budget non-atteint": ("#fee2e2", "#7f1d1d"),
        "en pause": ("#f3f4f6", "#374151"),
        "finie": ("#f3f4f6", "#374151"),
        "budget non atteint": ("#fee2e2", "#7f1d1d"),
    }
    bg, fg = colors.get(statut.lower(), ("#f3f4f6", "#374151"))
    label = statut.capitalize()
    return f'<span style="background:{bg};color:{fg};padding:2px 8px;border-radius:12px;font-size:0.8rem;font-weight:600">{label}</span>'


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

with st.sidebar:
    st.title("📊 Campaign Tracker")
    st.divider()

    uploaded = st.file_uploader(
        "Importer CSV / Excel",
        type=["csv", "xlsx", "xls"],
        help="Export depuis votre outil de suivi (même format que Suivi Budget Mai 2026)",
    )

    st.divider()
    st.subheader("⚙️ Seuils d'alerte")
    budget_warn  = st.slider("⚠️ Budget proche (%)",  50, 99, 80, step=5)
    budget_danger = st.slider("🔴 Budget danger (%)", 50, 100, 95, step=5)
    publisher_gap = st.slider("💡 Écart éditeur min. (%)", 5, 50, 20, step=5)

    st.divider()
    st.subheader("🤖 IA — Suggestions Claude")
    anthropic_key = st.text_input(
        "Clé API Anthropic",
        type="password",
        placeholder="sk-ant-…",
        help="Obtenir une clé sur console.anthropic.com",
    )
    if not _ANTHROPIC_AVAILABLE:
        st.caption("⚠️ `pip3 install anthropic` requis")

    st.divider()
    st.subheader("🔔 Notifications")

    with st.expander("Slack"):
        slack_webhook = st.text_input("Webhook URL", type="password",
                                      placeholder="https://hooks.slack.com/…")

    with st.expander("Email"):
        smtp_host = st.text_input("Serveur SMTP", value="smtp.gmail.com")
        smtp_port = st.number_input("Port", value=465, step=1)
        smtp_user = st.text_input("Adresse email")
        smtp_pass = st.text_input("Mot de passe", type="password")
        notif_to  = st.text_input("Destinataire(s)", placeholder="vous@exemple.com")

    if st.button("📤 Envoyer les alertes maintenant", use_container_width=True):
        st.session_state["send_alerts"] = True

    st.divider()
    st.caption(f"Mise à jour : {datetime.now().strftime('%d/%m/%Y %H:%M')}")


# ─────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────

if uploaded is not None:
    raw = uploaded.read()
    fname = uploaded.name
    df = load_data(raw, fname)
else:
    # Try local sample file (dev only — won't exist on Streamlit Cloud)
    import os
    _sample = os.path.join(os.path.dirname(__file__), "sample_data.csv")
    if os.path.exists(_sample):
        with open(_sample, "rb") as f:
            raw = f.read()
        df = load_data(raw, "sample_data.csv")
        st.info("Fichier de démo chargé — importez votre propre fichier via la barre latérale.")
    else:
        df = pd.DataFrame()

if df.empty:
    st.warning("Aucune donnée chargée. Importez un fichier CSV ou Excel via la barre latérale.")
    st.stop()

# Compute alerts
alerts = compute_alerts(df, budget_warn, budget_danger, publisher_gap)

# Handle "send alerts" button
if st.session_state.get("send_alerts"):
    msg = build_alert_message(alerts)
    sent = []
    if slack_webhook:
        ok = send_slack(slack_webhook, msg)
        sent.append(f"Slack {'✅' if ok else '❌'}")
    if smtp_user and notif_to:
        body_html = msg.replace("\n", "<br>").replace("*", "<b>").replace("*", "</b>")
        ok = send_email(smtp_host, int(smtp_port), smtp_user, smtp_pass,
                        notif_to, "Campaign Tracker — Alertes", body_html)
        sent.append(f"Email {'✅' if ok else '❌'}")
    if sent:
        st.sidebar.success("Envoyé : " + " | ".join(sent))
    else:
        st.sidebar.warning("Configurez Slack ou Email avant d'envoyer.")
    st.session_state["send_alerts"] = False


# ─────────────────────────────────────────────
# GLOBAL FILTERS (top bar)
# ─────────────────────────────────────────────

total_alerts = len(alerts["danger"]) + len(alerts["warning"]) + len(alerts["publisher"])
alert_label = f"🔔 Alertes ({total_alerts})" if total_alerts else "🔔 Alertes"

tab_overview, tab_campaigns, tab_alerts, tab_publisher = st.tabs(
    ["🏠 Vue d'ensemble", "📋 Campagnes", alert_label, "💡 Éditeurs"]
)


# ─────────────────────────────────────────────
# TAB 1 — OVERVIEW
# ─────────────────────────────────────────────

with tab_overview:
    st.header("Vue d'ensemble — Mai 2026")

    active_df   = df[df["Statut"] == "active"]
    setup_df    = df[df["Statut"] == "set-up"]
    done_df     = df[df["Statut"].isin(["budget fait", "finie", "budget non-atteint", "budget non atteint"])]

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Campagnes totales", len(df))
    col2.metric("🟢 Actives", len(active_df))
    col3.metric("🔵 Set-up", len(setup_df))
    col4.metric("✅ Terminées", len(done_df))
    col5.metric("🔔 Alertes", total_alerts,
                delta=f"{len(alerts['danger'])} danger" if alerts["danger"] else None,
                delta_color="inverse")

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        total_budget = df["Budget"].sum()
        spent = df.apply(
            lambda r: (r["Budget"] - r["Budget_restant"])
            if not pd.isna(r.get("Budget")) and not pd.isna(r.get("Budget_restant"))
            else np.nan, axis=1
        ).sum()
        total_ca_int = df["CA_interne"].sum() if "CA_interne" in df.columns else 0
        total_marge  = df["Marge_€"].sum() if "Marge_€" in df.columns else 0

        st.metric("Budget total", fmt_eur(total_budget))
        st.metric("CA interne total", fmt_eur(total_ca_int))
        st.metric("Marge totale (€)", fmt_eur(total_marge))

    with c2:
        # Budget utilisation by levier
        if "Levier" in df.columns and "Budget" in df.columns:
            levier_budget = (
                df.groupby("Levier")["Budget"]
                .sum()
                .reset_index()
                .sort_values("Budget", ascending=False)
            )
            fig = px.bar(
                levier_budget, x="Levier", y="Budget",
                title="Budget par Levier",
                color="Budget",
                color_continuous_scale="Blues",
                labels={"Budget": "Budget (€)"},
            )
            fig.update_layout(height=280, margin=dict(t=40, b=0))
            st.plotly_chart(fig, use_container_width=True)

    st.divider()
    c3, c4 = st.columns(2)

    with c3:
        # Campaigns by Statut
        status_counts = df["Statut"].value_counts().reset_index()
        status_counts.columns = ["Statut", "Nombre"]
        fig2 = px.pie(
            status_counts, names="Statut", values="Nombre",
            title="Répartition par Statut",
            color_discrete_sequence=px.colors.qualitative.Pastel,
        )
        fig2.update_layout(height=300, margin=dict(t=40, b=0))
        st.plotly_chart(fig2, use_container_width=True)

    with c4:
        # Marge by Verticale
        if "Verticale" in df.columns and "Marge_€" in df.columns:
            vert = (
                df[df["Verticale"].str.strip() != ""]
                .groupby("Verticale")["Marge_€"]
                .sum()
                .reset_index()
                .sort_values("Marge_€", ascending=False)
                .head(10)
            )
            fig3 = px.bar(
                vert, x="Marge_€", y="Verticale", orientation="h",
                title="Marge par Verticale (Top 10)",
                color="Marge_€",
                color_continuous_scale="RdYlGn",
            )
            fig3.update_layout(height=300, margin=dict(t=40, b=0))
            st.plotly_chart(fig3, use_container_width=True)

    # Budget pace scatter
    if "Pct_Budget" in df.columns and "Pct_Temps" in df.columns:
        st.subheader("Rythme budgétaire (% Budget vs % Temps passé)")
        pace = df[df["Pct_Budget"].notna() & df["Pct_Temps"].notna()].copy()
        pace = pace[~pace["Pct_Temps"].astype(str).str.lower().isin(["finie", "nan"])]
        if not pace.empty:
            pace["Pct_Temps_num"] = pd.to_numeric(pace["Pct_Temps"], errors="coerce")
            pace = pace[pace["Pct_Temps_num"].notna()]
            pace["Retard"] = pace["Pct_Budget"] - pace["Pct_Temps_num"]

            fig4 = px.scatter(
                pace,
                x="Pct_Temps_num", y="Pct_Budget",
                text="Campagne",
                color="Retard",
                color_continuous_scale="RdYlGn",
                labels={"Pct_Temps_num": "% Temps passé", "Pct_Budget": "% Budget utilisé"},
                hover_data=["Campagne", "Levier", "Budget_restant"],
            )
            fig4.add_shape(type="line", x0=0, y0=0, x1=150, y1=150,
                           line=dict(color="gray", dash="dash"))
            fig4.update_traces(textposition="top center", textfont_size=8)
            fig4.update_layout(height=420, margin=dict(t=10))
            st.plotly_chart(fig4, use_container_width=True)
            st.caption("Points au-dessus de la diagonale = budget consommé plus vite que le temps écoulé")


# ─────────────────────────────────────────────
# TAB 2 — CAMPAIGN LIST & DETAIL
# ─────────────────────────────────────────────

with tab_campaigns:
    st.header("Analyse des campagnes")

    # Filters
    fc1, fc2, fc3, fc4, fc5 = st.columns(5)
    with fc1:
        f_statut = st.multiselect("Statut", sorted(df["Statut"].unique()), default=[])
    with fc2:
        leviers = sorted(df["Levier"].dropna().unique()) if "Levier" in df.columns else []
        f_levier = st.multiselect("Levier", leviers, default=[])
    with fc3:
        geos = sorted(df["Géo"].dropna().unique()) if "Géo" in df.columns else []
        f_geo = st.multiselect("Géo", geos, default=[])
    with fc4:
        models = sorted(df["Modèle"].dropna().unique()) if "Modèle" in df.columns else []
        f_model = st.multiselect("Modèle", models, default=[])
    with fc5:
        accounts = sorted(df["Account"].dropna().replace("", np.nan).dropna().unique()) if "Account" in df.columns else []
        f_account = st.multiselect("Account", accounts, default=[])

    filtered = df.copy()
    if f_statut:
        filtered = filtered[filtered["Statut"].isin(f_statut)]
    if f_levier:
        filtered = filtered[filtered["Levier"].isin(f_levier)]
    if f_geo:
        filtered = filtered[filtered["Géo"].isin(f_geo)]
    if f_model:
        filtered = filtered[filtered["Modèle"].isin(f_model)]
    if f_account:
        filtered = filtered[filtered["Account"].isin(f_account)]

    search = st.text_input("🔍 Recherche campagne", placeholder="Nom, annonceur…")
    if search:
        mask = filtered["Campagne"].str.contains(search, case=False, na=False)
        if "Annonceur" in filtered.columns:
            mask |= filtered["Annonceur"].str.contains(search, case=False, na=False)
        filtered = filtered[mask]

    st.caption(f"{len(filtered)} campagne(s) affichée(s)")

    # Display table
    display_cols = ["Campagne", "Géo", "Levier", "Modèle", "Statut",
                    "Budget", "Budget_restant", "Pct_Budget", "Pct_Temps",
                    "Volume_réalisé", "Volume_restant", "Marge_pct", "Marge_€",
                    "Rém_Ann", "Rém_NET", "Rém_Éditeur", "Affiliate", "Account"]
    disp = filtered[[c for c in display_cols if c in filtered.columns]].copy()

    # Format for display
    for col in ["Budget", "Budget_restant", "Marge_€", "Rém_Ann", "Rém_NET", "Rém_Éditeur"]:
        if col in disp.columns:
            disp[col] = disp[col].apply(lambda x: fmt_eur(x) if not pd.isna(x) else "—")
    for col in ["Pct_Budget", "Pct_Temps", "Marge_pct"]:
        if col in disp.columns:
            disp[col] = disp[col].apply(lambda x: fmt_pct(x) if not pd.isna(x) else "—")
    for col in ["Volume_réalisé", "Volume_restant"]:
        if col in disp.columns:
            disp[col] = disp[col].apply(
                lambda x: f"{int(x):,}".replace(",", " ") if not pd.isna(x) else "—"
            )

    st.dataframe(disp, use_container_width=True, height=420, hide_index=True)

    # Campaign detail
    st.divider()
    st.subheader("Détail d'une campagne")
    camp_names = sorted(filtered["Campagne"].unique())
    selected = st.selectbox("Sélectionner une campagne", camp_names)

    if selected:
        row = filtered[filtered["Campagne"] == selected].iloc[0]

        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Budget total", fmt_eur(row.get("Budget")))
        d2.metric("Budget restant", fmt_eur(row.get("Budget_restant")))
        d3.metric("% Budget utilisé", fmt_pct(row.get("Pct_Budget")))
        d4.metric("% Temps passé", fmt_pct(row.get("Pct_Temps")) if str(row.get("Pct_Temps","")).lower() not in ["finie","nan",""] else str(row.get("Pct_Temps", "—")))

        d5, d6, d7, d8 = st.columns(4)
        d5.metric("Volume réalisé", f"{int(row['Volume_réalisé']):,}".replace(",", " ") if not pd.isna(row.get("Volume_réalisé")) else "—")
        d6.metric("Volume restant", f"{int(row['Volume_restant']):,}".replace(",", " ") if not pd.isna(row.get("Volume_restant")) else "—")
        d7.metric("Marge %", fmt_pct(row.get("Marge_pct")))
        d8.metric("Marge €", fmt_eur(row.get("Marge_€")))

        d9, d10, d11, _ = st.columns(4)
        d9.metric("Rém. Annonceur", fmt_eur(row.get("Rém_Ann")))
        d10.metric("Rém. NET", fmt_eur(row.get("Rém_NET")))
        d11.metric("Rém. Éditeur", fmt_eur(row.get("Rém_Éditeur")))

        with st.expander("Toutes les données"):
            st.json({k: str(v) for k, v in row.items()})

        # AI suggestions in campaign detail
        st.subheader("🤖 Suggestions IA")
        ai_c1, ai_c2, ai_c3 = st.columns(3)
        row_dict_detail = row.to_dict()
        detail_key = selected.replace(" ", "_")[:40]

        with ai_c1:
            if st.button("💡 Idées stratégiques", key=f"detail_strat_{detail_key}"):
                with st.spinner("Génération en cours…"):
                    st.session_state[f"detail_ai_{detail_key}_strat"] = get_ai_suggestions(
                        anthropic_key, "strategie", row_dict_detail
                    )
        with ai_c2:
            if st.button("📈 Propositions upsell", key=f"detail_up_{detail_key}"):
                with st.spinner("Génération en cours…"):
                    st.session_state[f"detail_ai_{detail_key}_up"] = get_ai_suggestions(
                        anthropic_key, "upsell", row_dict_detail
                    )
        with ai_c3:
            if st.button("🚨 Solutions dépassement", key=f"detail_dep_{detail_key}"):
                with st.spinner("Génération en cours…"):
                    st.session_state[f"detail_ai_{detail_key}_dep"] = get_ai_suggestions(
                        anthropic_key, "depassement", row_dict_detail
                    )

        for suffix, label in [
            ("strat", "💡 Idées stratégiques"),
            ("up",    "📈 Propositions upsell"),
            ("dep",   "🚨 Solutions dépassement"),
        ]:
            key = f"detail_ai_{detail_key}_{suffix}"
            if key in st.session_state:
                with st.expander(label, expanded=True):
                    st.markdown(st.session_state[key])

        # Mini budget gauge
        pct = row.get("Pct_Budget", np.nan)
        if not pd.isna(pct):
            color = COLOR_DANGER if pct >= budget_danger else (COLOR_WARN if pct >= budget_warn else COLOR_OK)
            fig_g = go.Figure(go.Indicator(
                mode="gauge+number",
                value=min(pct, 150),
                title={"text": "% Budget utilisé"},
                gauge={
                    "axis": {"range": [0, 120]},
                    "bar": {"color": color},
                    "steps": [
                        {"range": [0, budget_warn], "color": "#dcfce7"},
                        {"range": [budget_warn, budget_danger], "color": "#fef9c3"},
                        {"range": [budget_danger, 120], "color": "#fee2e2"},
                    ],
                    "threshold": {"line": {"color": "red", "width": 4}, "value": 100},
                },
                number={"suffix": "%"},
            ))
            fig_g.update_layout(height=250, margin=dict(t=30, b=0))
            st.plotly_chart(fig_g, use_container_width=True)


# ─────────────────────────────────────────────
# TAB 3 — ALERTS
# ─────────────────────────────────────────────

with tab_alerts:
    st.header("Centre d'alertes")

    n_d = len(alerts["danger"])
    n_w = len(alerts["warning"])
    n_p = len(alerts["publisher"])

    ca1, ca2, ca3 = st.columns(3)
    ca1.metric("🔴 En danger", n_d)
    ca2.metric("⚠️ Proches du budget", n_w)
    ca3.metric("💡 Opportunités éditeur", n_p)

    st.divider()

    if n_d == 0 and n_w == 0 and n_p == 0:
        st.markdown('<div class="alert-box alert-success">✅ Aucune alerte active — toutes vos campagnes sont dans les seuils.</div>', unsafe_allow_html=True)

    def _ai_buttons(campagne_name: str, situation: str):
        """Render the three AI suggestion buttons for a given campaign."""
        row_data = df[df["Campagne"] == campagne_name]
        if row_data.empty:
            return
        row_dict = row_data.iloc[0].to_dict()

        btn_col1, btn_col2, btn_col3 = st.columns(3)
        key_base = campagne_name.replace(" ", "_")[:40]

        with btn_col1:
            if st.button("💡 Idées stratégiques", key=f"strat_{key_base}"):
                with st.spinner("Génération en cours…"):
                    st.session_state[f"ai_{key_base}_strat"] = get_ai_suggestions(
                        anthropic_key, "strategie", row_dict
                    )
        with btn_col2:
            if st.button("📈 Propositions upsell", key=f"up_{key_base}"):
                with st.spinner("Génération en cours…"):
                    st.session_state[f"ai_{key_base}_up"] = get_ai_suggestions(
                        anthropic_key, "upsell", row_dict
                    )
        with btn_col3:
            if situation == "danger":
                if st.button("🚨 Solutions dépassement", key=f"dep_{key_base}"):
                    with st.spinner("Génération en cours…"):
                        st.session_state[f"ai_{key_base}_dep"] = get_ai_suggestions(
                            anthropic_key, "depassement", row_dict
                        )

        for suffix, label in [("strat", "💡 Idées stratégiques"), ("up", "📈 Upsell"), ("dep", "🚨 Solutions dépassement")]:
            key = f"ai_{key_base}_{suffix}"
            if key in st.session_state:
                with st.expander(f"{label} — {campagne_name}", expanded=True):
                    st.markdown(st.session_state[key])

    # Danger
    if alerts["danger"]:
        st.subheader("🔴 Campagnes en danger")
        for a in sorted(alerts["danger"], key=lambda x: x["pct"] or 0, reverse=True):
            pct_str = fmt_pct(a["pct"])
            b_str   = fmt_eur(a["budget_restant"])
            st.markdown(
                f'<div class="alert-box alert-danger">'
                f'<b>{a["campagne"]}</b> &nbsp;|&nbsp; {a["levier"]} &nbsp;|&nbsp; Account: {a["affiliate"]}<br>'
                f'📊 {pct_str} du budget utilisé &nbsp;·&nbsp; Restant : {b_str}'
                f'</div>',
                unsafe_allow_html=True,
            )
            _ai_buttons(a["campagne"], "danger")
            st.divider()

    # Warning
    if alerts["warning"]:
        st.subheader("⚠️ Campagnes proches du budget")
        for a in sorted(alerts["warning"], key=lambda x: x["pct"] or 0, reverse=True):
            pct_str = fmt_pct(a["pct"])
            b_str   = fmt_eur(a["budget_restant"])
            st.markdown(
                f'<div class="alert-box alert-warning">'
                f'<b>{a["campagne"]}</b> &nbsp;|&nbsp; {a["levier"]} &nbsp;|&nbsp; Account: {a["affiliate"]}<br>'
                f'📊 {pct_str} du budget utilisé &nbsp;·&nbsp; Restant : {b_str}'
                f'</div>',
                unsafe_allow_html=True,
            )
            _ai_buttons(a["campagne"], "warning")
            st.divider()

    # Publisher
    if alerts["publisher"]:
        st.subheader("💡 Opportunités éditeur")
        for a in sorted(alerts["publisher"], key=lambda x: x["total_saving"], reverse=True):
            st.markdown(
                f'<div class="alert-box alert-info">'
                f'<b>{a["campagne"]}</b> ({a["modele"]}) — '
                f'Votre coût éditeur : <b>{fmt_eur(a["current_cost"])}</b> &nbsp;→&nbsp; '
                f'Plus attractif chez <b>{a["cheaper_campagne"]}</b> : <b>{fmt_eur(a["cheaper_cost"])}</b><br>'
                f'Volume restant : {int(a["volume_restant"]):,} unités &nbsp;·&nbsp; '
                f'Économie potentielle : <b>{fmt_eur(a["total_saving"])}</b>'
                f'</div>',
                unsafe_allow_html=True,
            )

    # Export alerts as CSV
    st.divider()
    all_alert_rows = []
    for a in alerts["danger"]:
        all_alert_rows.append({"Type": "Danger", "Campagne": a["campagne"], "Détail": a["detail"]})
    for a in alerts["warning"]:
        all_alert_rows.append({"Type": "Avertissement", "Campagne": a["campagne"], "Détail": a["detail"]})
    for a in alerts["publisher"]:
        all_alert_rows.append({
            "Type": "Opportunité éditeur",
            "Campagne": a["campagne"],
            "Détail": f"Coût actuel {fmt_eur(a['current_cost'])} → {fmt_eur(a['cheaper_cost'])} — économie {fmt_eur(a['total_saving'])}",
        })
    if all_alert_rows:
        alert_df = pd.DataFrame(all_alert_rows)
        csv_buf = io.StringIO()
        alert_df.to_csv(csv_buf, index=False, sep=";")
        st.download_button(
            "⬇️ Exporter les alertes (CSV)",
            data=csv_buf.getvalue().encode("utf-8-sig"),
            file_name=f"alertes_{date.today().isoformat()}.csv",
            mime="text/csv",
        )


# ─────────────────────────────────────────────
# TAB 4 — PUBLISHER COMPARISON
# ─────────────────────────────────────────────

with tab_publisher:
    st.header("Comparaison éditeurs / Coûts")

    if "Rém_Éditeur" not in df.columns or "Modèle" not in df.columns:
        st.warning("Données de rémunération éditeur non disponibles.")
    else:
        # Publisher cost by model
        pub_df = df[df["Rém_Éditeur"].notna() & df["Modèle"].notna()].copy()

        pm1, pm2 = st.columns(2)
        with pm1:
            sel_model = st.selectbox(
                "Filtrer par Modèle",
                ["Tous"] + sorted(pub_df["Modèle"].unique()),
            )
        with pm2:
            sel_statut_pub = st.multiselect(
                "Statut", sorted(pub_df["Statut"].unique()), default=["active"]
            )

        view = pub_df.copy()
        if sel_model != "Tous":
            view = view[view["Modèle"] == sel_model]
        if sel_statut_pub:
            view = view[view["Statut"].isin(sel_statut_pub)]

        if view.empty:
            st.info("Aucune campagne pour ces filtres.")
        else:
            # Scatter: Rém_NET vs Rém_Éditeur (margin visualisation)
            fig_pub = px.scatter(
                view,
                x="Rém_Éditeur", y="Rém_NET",
                color="Modèle",
                size="Budget",
                hover_name="Campagne",
                hover_data=["Levier", "Marge_pct", "Volume_restant", "Statut"],
                title="Rémunération NET vs Coût Éditeur",
                labels={"Rém_Éditeur": "Coût Éditeur (€)", "Rém_NET": "Rém. NET (€)"},
            )
            # Diagonal = break-even
            max_val = max(view["Rém_Éditeur"].max(), view["Rém_NET"].max())
            fig_pub.add_shape(type="line", x0=0, y0=0, x1=max_val, y1=max_val,
                              line=dict(color="red", dash="dash"))
            fig_pub.update_layout(height=420)
            st.plotly_chart(fig_pub, use_container_width=True)
            st.caption("Points au-dessus de la diagonale rouge = marge positive. En dessous = perte.")

            st.divider()
            # Ranking table
            st.subheader("Classement par attractivité éditeur (coût le plus bas)")
            rank = view[["Campagne", "Modèle", "Levier", "Rém_Éditeur", "Rém_NET",
                         "Marge_pct", "Volume_restant", "Statut"]].copy()
            rank = rank.sort_values("Rém_Éditeur")
            rank["Marge_pct"] = rank["Marge_pct"].apply(fmt_pct)
            rank["Rém_Éditeur"] = rank["Rém_Éditeur"].apply(fmt_eur)
            rank["Rém_NET"] = rank["Rém_NET"].apply(fmt_eur)
            rank["Volume_restant"] = rank["Volume_restant"].apply(
                lambda x: f"{int(x):,}".replace(",", " ") if not pd.isna(x) else "—"
            )
            st.dataframe(rank, use_container_width=True, hide_index=True)

            st.divider()
            # Best opportunities: active with remaining volume & best margin
            st.subheader("Top opportunités — Volume restant + Marge élevée")
            opps = pub_df[
                (pub_df["Statut"].isin(["active", "set-up"])) &
                pub_df["Volume_restant"].notna() &
                (pub_df["Volume_restant"] > 0) &
                pub_df["Marge_pct"].notna()
            ].copy()
            opps = opps.sort_values("Marge_pct", ascending=False).head(15)
            if not opps.empty:
                fig_opp = px.bar(
                    opps, x="Campagne", y="Marge_pct",
                    color="Modèle",
                    hover_data=["Volume_restant", "Rém_Éditeur", "Rém_NET"],
                    title="Marge % — Campagnes avec volume restant (Top 15)",
                    labels={"Marge_pct": "Marge (%)"},
                )
                fig_opp.update_layout(height=350, xaxis_tickangle=-45)
                st.plotly_chart(fig_opp, use_container_width=True)
