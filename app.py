import streamlit as st
import pandas as pd
import json
import os
import qrcode
import io
import hashlib
import uuid
from PIL import Image
from datetime import datetime

try:
    import openpyxl
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

try:
    from streamlit_quagga import quagga
    QUAGGA_AVAILABLE = True
except ImportError:
    quagga = None
    QUAGGA_AVAILABLE = False

# Dateipfade
USERS_FILE      = "benutzer.json"
WORKSPACES_FILE = "workspaces.json"
IMAGES_DIR      = "artikel_bilder"
os.makedirs(IMAGES_DIR, exist_ok=True)

DEFAULT_KATEGORIEN = [
    "Elektronik", "Möbel", "Bürobedarf", "Werkzeug",
    "Film", "Musik", "Buch", "Spiel", "Sonstiges"
]

# =============================================
# KATEGORIE-SPEZIFISCHE ZUSATZFELDER
# =============================================
KATEGORIE_FELDER = {
    "Film": [
        {"key": "genre", "label": "Genre", "typ": "select",
         "optionen": ["Action", "Komödie", "Drama", "Horror", "Science-Fiction",
                      "Fantasy", "Thriller", "Dokumentation", "Animation",
                      "Romance", "Krimi", "Abenteuer", "Sonstiges"]},
        {"key": "altersfreigabe", "label": "Altersfreigabe", "typ": "select",
         "optionen": ["Ohne Altersbeschränkung", "Ab 6", "Ab 12", "Ab 16", "Ab 18"]},
        {"key": "regisseur",        "label": "Regisseur",        "typ": "text"},
        {"key": "erscheinungsjahr", "label": "Erscheinungsjahr", "typ": "text"},
        {"key": "format",           "label": "Format",           "typ": "select",
         "optionen": ["DVD", "Blu-ray", "4K UHD", "VHS", "Digital", "Sonstiges"]},
    ],
    "Musik": [
        {"key": "genre", "label": "Genre", "typ": "select",
         "optionen": ["Pop", "Rock", "Hip-Hop", "Jazz", "Klassik",
                      "Electronic", "Metal", "Country", "R&B", "Sonstiges"]},
        {"key": "kuenstler",        "label": "Künstler / Band",  "typ": "text"},
        {"key": "erscheinungsjahr", "label": "Erscheinungsjahr", "typ": "text"},
        {"key": "format",           "label": "Format",           "typ": "select",
         "optionen": ["CD", "Vinyl", "Kassette", "Digital", "Sonstiges"]},
    ],
    "Buch": [
        {"key": "genre", "label": "Genre", "typ": "select",
         "optionen": ["Roman", "Sachbuch", "Krimi", "Fantasy", "Science-Fiction",
                      "Biografie", "Kinderbuch", "Manga", "Comic", "Ratgeber", "Sonstiges"]},
        {"key": "autor",            "label": "Autor",            "typ": "text"},
        {"key": "erscheinungsjahr", "label": "Erscheinungsjahr", "typ": "text"},
        {"key": "isbn",             "label": "ISBN",             "typ": "text"},
        {"key": "altersempfehlung", "label": "Altersempfehlung", "typ": "select",
         "optionen": ["Alle Altersgruppen", "Ab 6", "Ab 10", "Ab 12", "Ab 16", "Ab 18", "Erwachsene"]},
    ],
    "Spiel": [
        {"key": "genre", "label": "Genre", "typ": "select",
         "optionen": ["Brettspiel", "Kartenspiel", "Videospiel", "Rollenspiel",
                      "Puzzle", "Würfelspiel", "Partyspiel", "Sonstiges"]},
        {"key": "altersfreigabe",  "label": "Altersfreigabe",  "typ": "select",
         "optionen": ["Ohne Altersbeschränkung", "Ab 6", "Ab 12", "Ab 16", "Ab 18"]},
        {"key": "spieler_anzahl",  "label": "Spieleranzahl",   "typ": "text",
         "placeholder": "z.B. 2–4"},
        {"key": "hersteller",      "label": "Hersteller",      "typ": "text"},
    ],
    "Elektronik": [
        {"key": "hersteller", "label": "Hersteller / Marke", "typ": "text"},
        {"key": "modell",     "label": "Modell",             "typ": "text"},
        {"key": "zustand",    "label": "Zustand",            "typ": "select",
         "optionen": ["Neu", "Wie neu", "Gut", "Akzeptabel", "Defekt"]},
    ],
    "Werkzeug": [
        {"key": "hersteller", "label": "Hersteller / Marke", "typ": "text"},
        {"key": "zustand",    "label": "Zustand",            "typ": "select",
         "optionen": ["Neu", "Wie neu", "Gut", "Akzeptabel", "Defekt"]},
    ],
    "Möbel": [
        {"key": "material", "label": "Material", "typ": "text"},
        {"key": "farbe",    "label": "Farbe",    "typ": "text"},
        {"key": "zustand",  "label": "Zustand",  "typ": "select",
         "optionen": ["Neu", "Wie neu", "Gut", "Akzeptabel", "Defekt"]},
    ],
    "Bürobedarf": [
        {"key": "hersteller", "label": "Hersteller / Marke", "typ": "text"},
        {"key": "zustand",    "label": "Zustand",            "typ": "select",
         "optionen": ["Neu", "Wie neu", "Gut", "Akzeptabel", "Defekt"]},
    ],
}

# =============================================
# HILFSFUNKTIONEN
# =============================================

def load_json(file, default):
    if os.path.exists(file):
        with open(file, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(file, data):
    with open(file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def hash_passwort(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    default = [{"benutzername": "admin", "passwort": hash_passwort("admin123"),
                "rolle": "admin", "name": "Administrator"}]
    save_json(USERS_FILE, default)
    return default

def load_workspaces():
    return load_json(WORKSPACES_FILE, [])

def save_workspaces(ws_list):
    save_json(WORKSPACES_FILE, ws_list)

def get_user_workspaces(benutzername):
    return [w for w in load_workspaces()
            if w["besitzer"] == benutzername or benutzername in w.get("mitglieder", [])]

def get_ws_data_files(ws_id):
    ws_dir     = os.path.join("workspaces", ws_id)
    bilder_dir = os.path.join(ws_dir, "bilder")
    os.makedirs(bilder_dir, exist_ok=True)
    return {
        "inventar":   os.path.join(ws_dir, "inventar.json"),
        "raeume":     os.path.join(ws_dir, "raeume.json"),
        "kategorien": os.path.join(ws_dir, "kategorien.json"),
        "ausleihen":  os.path.join(ws_dir, "ausleihen.json"),
        "bilder":     bilder_dir,
    }

def init_workspace_data(ws_id):
    files = get_ws_data_files(ws_id)
    st.session_state.inventar   = load_json(files["inventar"],   [])
    st.session_state.raeume     = load_json(files["raeume"],     ["Raum 1"])
    st.session_state.kategorien = load_json(files["kategorien"], DEFAULT_KATEGORIEN)
    st.session_state.ausleihen  = load_json(files["ausleihen"],  [])
    st.session_state.ws_files   = files
    erlaubte    = st.session_state.get("erlaubte_raeume", [])
    verfuegbare = [r for r in st.session_state.raeume if not erlaubte or r in erlaubte]
    st.session_state.aktiver_raum = verfuegbare[0] if verfuegbare else None

def get_ws_image_path(artikel_id):
    bilder_dir = st.session_state.get("ws_files", {}).get("bilder", IMAGES_DIR)
    for ext in ["jpg", "jpeg", "png", "webp"]:
        path = os.path.join(bilder_dir, f"{artikel_id}.{ext}")
        if os.path.exists(path):
            return path
    return None

def save_ws_image(artikel_id, uploaded_file):
    bilder_dir = st.session_state.get("ws_files", {}).get("bilder", IMAGES_DIR)
    os.makedirs(bilder_dir, exist_ok=True)
    ext  = uploaded_file.name.split(".")[-1].lower()
    path = os.path.join(bilder_dir, f"{artikel_id}.{ext}")
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path

def generate_qr_code(artikel):
    info = (f"ID: {artikel['id']}\nName: {artikel['name']}\n"
            f"Raum: {artikel['raum']}\nMenge: {artikel['menge']}")
    qr = qrcode.QRCode(version=1, box_size=6, border=2)
    qr.add_data(info)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

def ist_ws_besitzer(ws):
    return ws["besitzer"] == st.session_state.benutzername

def get_ws_rolle(ws):
    if ws["besitzer"] == st.session_state.benutzername:
        return "admin"
    return ws.get("mitglieder_rollen", {}).get(st.session_state.benutzername, "user")

ROLLEN_RECHTE = {
    "admin":   ["inventar", "artikel_hinzufuegen", "checkout", "barcode",
                "statistiken", "export", "benutzerverwaltung"],
    "manager": ["inventar", "artikel_hinzufuegen", "checkout", "barcode",
                "statistiken", "export"],
    "user":    ["inventar", "artikel_hinzufuegen", "checkout", "barcode"],
    "viewer":  ["inventar"],
}
ROLLEN_NAMEN = {
    "admin":   "👑 Administrator",
    "manager": "🔧 Manager",
    "user":    "👤 Benutzer",
    "viewer":  "👁️ Betrachter",
}

def hat_recht(recht):
    return recht in ROLLEN_RECHTE.get(st.session_state.get("rolle", ""), [])

def hat_raum_zugriff(raum):
    if st.session_state.get("rolle") == "admin":
        return True
    erlaubte = st.session_state.get("erlaubte_raeume", [])
    return True if not erlaubte else raum in erlaubte

def get_erlaubte_raeume():
    if st.session_state.get("rolle") == "admin":
        return st.session_state.raeume
    erlaubte = st.session_state.get("erlaubte_raeume", [])
    if not erlaubte:
        return st.session_state.raeume
    return [r for r in st.session_state.raeume if r in erlaubte]

# =============================================
# SESSION STATE INITIALISIERUNG
# =============================================
for _k, _v in [
    ("eingeloggt",          False),
    ("benutzername",        ""),
    ("rolle",               ""),
    ("benutzer_name",       ""),
    ("aktiver_workspace",   None),
    ("ws_files",            {}),
    ("edit_artikel_id",     None),
    ("detail_artikel_id",   None),
    ("neu_kat_auswahl",     None),
]:
    if _k not in st.session_state:
        st.session_state[_k] = _v

# =============================================
# LOGIN
# =============================================
if not st.session_state.eingeloggt:
    st.set_page_config(page_title="Login – EVA", page_icon="🔐", layout="centered")
    st.title("🔐 EVA Inventarisierungs-App")
    st.markdown("Bitte melde dich an oder erstelle einen neuen Account.")
    st.divider()

    login_tab, register_tab = st.tabs(["🔐 Anmelden", "📝 Registrieren"])

    with login_tab:
        with st.form("login_form"):
            bn_input  = st.text_input("👤 Benutzername")
            pw_input  = st.text_input("🔑 Passwort", type="password")
            login_btn = st.form_submit_button("Anmelden", type="primary", use_container_width=True)
        if login_btn:
            users = load_users()
            user  = next((u for u in users
                          if u["benutzername"] == bn_input
                          and u["passwort"] == hash_passwort(pw_input)), None)
            if user:
                st.session_state.eingeloggt      = True
                st.session_state.benutzername    = user["benutzername"]
                st.session_state.rolle           = user["rolle"]
                st.session_state.benutzer_name   = user["name"]
                st.session_state.erlaubte_raeume = user.get("erlaubte_raeume", [])
                st.rerun()
            else:
                st.error("❌ Benutzername oder Passwort falsch!")
        st.caption("💡 Noch kein Account? Wechsle zum Tab 'Registrieren'.")

    with register_tab:
        with st.form("register_form"):
            reg_name = st.text_input("👤 Vollständiger Name", placeholder="z.B. Max Mustermann")
            reg_bn   = st.text_input("👤 Benutzername",       placeholder="z.B. max.mustermann")
            reg_pw1  = st.text_input("🔑 Passwort wählen",    type="password")
            reg_pw2  = st.text_input("🔑 Passwort wiederholen", type="password")
            reg_btn  = st.form_submit_button("✅ Account erstellen", type="primary",
                                              use_container_width=True)
        if reg_btn:
            users = load_users()
            if not reg_name.strip() or not reg_bn.strip() or not reg_pw1.strip():
                st.error("❌ Bitte alle Felder ausfüllen!")
            elif reg_pw1 != reg_pw2:
                st.error("❌ Die Passwörter stimmen nicht überein!")
            elif len(reg_pw1) < 6:
                st.error("❌ Das Passwort muss mindestens 6 Zeichen lang sein!")
            elif any(u["benutzername"] == reg_bn.strip() for u in users):
                st.warning("⚠️ Dieser Benutzername ist bereits vergeben!")
            else:
                users.append({
                    "benutzername":    reg_bn.strip(),
                    "passwort":        hash_passwort(reg_pw1),
                    "rolle":           "user",
                    "name":            reg_name.strip(),
                    "erlaubte_raeume": [],
                    "registriert_am":  datetime.now().strftime("%d.%m.%Y %H:%M"),
                })
                save_json(USERS_FILE, users)
                st.success(f"✅ Account '{reg_bn}' erstellt! Du kannst dich jetzt anmelden.")
                st.balloons()
        st.caption("🔒 Neue Accounts haben standardmäßig keine Adminrechte.")

    st.stop()

# =============================================
# WORKSPACE AUSWAHL
# =============================================
if st.session_state.aktiver_workspace is None:
    st.set_page_config(page_title="Workspace – EVA", page_icon="📂", layout="centered")
    st.title("📂 Workspace auswählen")
    st.markdown(f"Willkommen, **{st.session_state.benutzer_name}**!")
    st.divider()

    user_ws = get_user_workspaces(st.session_state.benutzername)

    if user_ws:
        st.markdown("### 📦 Deine Workspaces")
        for ws in user_ws:
            with st.container(border=True):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    label = ("👑 Eigener Workspace"
                             if ws["besitzer"] == st.session_state.benutzername
                             else f"👤 Geteilt von {ws['besitzer']}")
                    raum_count = len(load_json(get_ws_data_files(ws["id"])["raeume"], []))
                    st.markdown(f"**{ws['name']}**  \n{label} | 🏠 {raum_count} Räume")
                    if ws.get("beschreibung"):
                        st.caption(ws["beschreibung"])
                with c2:
                    pw_ein = (st.text_input("Passwort", type="password", key=f"pw_{ws['id']}")
                              if ws.get("passwort") else None)
                with c3:
                    if st.button("▶️ Öffnen", key=f"open_{ws['id']}", use_container_width=True):
                        if ws.get("passwort") and hash_passwort(pw_ein or "") != ws["passwort"]:
                            st.error("❌ Falsches Workspace-Passwort!")
                        else:
                            st.session_state.aktiver_workspace = ws
                            st.session_state.rolle = get_ws_rolle(ws)
                            mi = ws.get("mitglieder_info", {}).get(
                                st.session_state.benutzername, {})
                            st.session_state.erlaubte_raeume = mi.get("erlaubte_raeume", [])
                            init_workspace_data(ws["id"])
                            st.rerun()
                    if ist_ws_besitzer(ws):
                        if st.button("🗑️ Löschen", key=f"del_{ws['id']}", use_container_width=True):
                            all_ws = [w for w in load_workspaces() if w["id"] != ws["id"]]
                            save_workspaces(all_ws)
                            st.success(f"Workspace '{ws['name']}' gelöscht!")
                            st.rerun()
    else:
        st.info("💡 Du hast noch keinen Workspace. Erstelle deinen ersten!")

    st.divider()
    st.markdown("### ➕ Neuen Workspace erstellen")
    with st.form("ws_erstellen_form"):
        ws_name   = st.text_input("📝 Name", placeholder="z.B. Zuhause, Arbeit, Lager...")
        ws_beschr = st.text_input("💬 Beschreibung (optional)")
        ws_pw     = st.text_input("🔒 Passwort (optional)", type="password")
        ws_btn    = st.form_submit_button("✅ Workspace erstellen", type="primary",
                                           use_container_width=True)
    if ws_btn:
        if not ws_name.strip():
            st.error("Bitte einen Namen eingeben!")
        else:
            neuer_ws = {
                "id":                str(uuid.uuid4()),
                "name":              ws_name.strip(),
                "beschreibung":      ws_beschr.strip(),
                "besitzer":          st.session_state.benutzername,
                "passwort":          hash_passwort(ws_pw) if ws_pw.strip() else None,
                "mitglieder":        [],
                "mitglieder_rollen": {},
                "mitglieder_info":   {},
                "erstellt_am":       datetime.now().strftime("%d.%m.%Y %H:%M"),
            }
            all_ws = load_workspaces()
            all_ws.append(neuer_ws)
            save_workspaces(all_ws)
            st.success(f"✅ Workspace '{ws_name}' erstellt!")
            st.rerun()

    st.divider()
    st.markdown("### 🔗 Workspace beitreten")
    with st.form("ws_beitreten_form"):
        einladung = st.text_input("📎 Workspace-ID eingeben")
        join_btn  = st.form_submit_button("↗️ Beitreten", use_container_width=True)
    if join_btn:
        all_ws   = load_workspaces()
        gefunden = next((w for w in all_ws if w["id"] == einladung.strip()), None)
        if not gefunden:
            st.error("❌ Workspace nicht gefunden!")
        elif st.session_state.benutzername in gefunden.get("mitglieder", []):
            st.warning("Du bist bereits Mitglied.")
        elif gefunden["besitzer"] == st.session_state.benutzername:
            st.warning("Das ist dein eigener Workspace.")
        else:
            gefunden["mitglieder"].append(st.session_state.benutzername)
            gefunden.setdefault("mitglieder_rollen", {})[st.session_state.benutzername] = "user"
            save_workspaces(all_ws)
            st.success(f"✅ Du bist dem Workspace '{gefunden['name']}' beigetreten!")
            st.rerun()

    st.divider()
    if st.button("🚪 Abmelden", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    st.stop()

# Workspace-Daten laden
if "inventar" not in st.session_state:
    init_workspace_data(st.session_state.aktiver_workspace["id"])
if "erlaubte_raeume" not in st.session_state:
    st.session_state.erlaubte_raeume = []
if "aktiver_raum" not in st.session_state:
    e = get_erlaubte_raeume()
    st.session_state.aktiver_raum = e[0] if e else None

# =============================================
# SEITENLAYOUT
# =============================================
st.set_page_config(page_title="EVA – Inventarisierungs-App", page_icon="📦", layout="wide")
ws = st.session_state.aktiver_workspace
st.title(f"📦 {ws['name']}")

col_title, col_user = st.columns([4, 1])
with col_title:
    st.markdown(f"Workspace von **{ws['besitzer']}** | {ROLLEN_NAMEN.get(st.session_state.rolle, '')}")
with col_user:
    st.markdown(f"👤 **{st.session_state.benutzer_name}**")
    ca, cb = st.columns(2)
    with ca:
        if st.button("📂 Workspaces", use_container_width=True):
            st.session_state.aktiver_workspace = None
            for k in ["inventar", "raeume", "kategorien", "ausleihen",
                      "ws_files", "aktiver_raum", "detail_artikel_id", "neu_kat_auswahl"]:
                st.session_state.pop(k, None)
            st.rerun()
    with cb:
        if st.button("🚪 Abmelden", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

# =============================================
# SIDEBAR – Raum- & Kategorieverwaltung
# =============================================
with st.sidebar:
    st.header("🏠 Raumverwaltung")
    erlaubte_sidebar = get_erlaubte_raeume()
    if erlaubte_sidebar:
        aktiver_raum = st.selectbox(
            "Aktiver Raum", erlaubte_sidebar,
            index=erlaubte_sidebar.index(st.session_state.aktiver_raum)
            if st.session_state.aktiver_raum in erlaubte_sidebar else 0
        )
        st.session_state.aktiver_raum = aktiver_raum
    else:
        st.warning("Kein Raumzugriff vorhanden.")

    with st.expander("➕ Neuen Raum hinzufügen"):
        neuer_raum = st.text_input("Raumname", key="neuer_raum_input")
        if st.button("Raum hinzufügen"):
            if not neuer_raum.strip():
                st.error("Bitte einen Raumnamen eingeben!")
            elif neuer_raum.strip() in st.session_state.raeume:
                st.warning("Dieser Raum existiert bereits!")
            else:
                st.session_state.raeume.append(neuer_raum.strip())
                save_json(st.session_state.ws_files["raeume"], st.session_state.raeume)
                st.session_state.aktiver_raum = neuer_raum.strip()
                st.success(f"Raum '{neuer_raum}' hinzugefügt!")
                st.rerun()

    with st.expander("✏️ Raum umbenennen"):
        if st.session_state.raeume:
            raum_umb     = st.selectbox("Raum auswählen", st.session_state.raeume, key="raum_umb")
            neuer_raum_n = st.text_input("Neuer Name", value=raum_umb, key="raum_umb_name")
            if st.button("✏️ Umbenennen", key="btn_raum_umb"):
                if neuer_raum_n.strip() and neuer_raum_n.strip() != raum_umb:
                    idx = st.session_state.raeume.index(raum_umb)
                    st.session_state.raeume[idx] = neuer_raum_n.strip()
                    for a in st.session_state.inventar:
                        if a["raum"] == raum_umb:
                            a["raum"] = neuer_raum_n.strip()
                    if st.session_state.aktiver_raum == raum_umb:
                        st.session_state.aktiver_raum = neuer_raum_n.strip()
                    save_json(st.session_state.ws_files["raeume"], st.session_state.raeume)
                    save_json(st.session_state.ws_files["inventar"], st.session_state.inventar)
                    st.success(f"'{raum_umb}' wurde in '{neuer_raum_n.strip()}' umbenannt!")
                    st.rerun()

    with st.expander("🗑️ Raum löschen"):
        if st.session_state.raeume:
            raum_del = st.selectbox("Raum auswählen", st.session_state.raeume, key="raum_loeschen")
            if st.button("Raum löschen", key="btn_raum_loeschen"):
                im_raum = [a for a in st.session_state.inventar if a["raum"] == raum_del]
                if im_raum:
                    st.error(f"Raum enthält noch {len(im_raum)} Artikel.")
                else:
                    st.session_state.raeume.remove(raum_del)
                    save_json(st.session_state.ws_files["raeume"], st.session_state.raeume)
                    st.session_state.aktiver_raum = (st.session_state.raeume[0]
                                                     if st.session_state.raeume else None)
                    st.success(f"Raum '{raum_del}' gelöscht!")
                    st.rerun()

    st.divider()
    st.header("🗂️ Kategorien")

    with st.expander("➕ Neue Kategorie"):
        neue_kat = st.text_input("Kategoriename", key="neue_kat_input")
        if st.button("Kategorie hinzufügen"):
            if not neue_kat.strip():
                st.error("Bitte einen Namen eingeben!")
            elif neue_kat.strip() in st.session_state.kategorien:
                st.warning("Kategorie existiert bereits!")
            else:
                st.session_state.kategorien.append(neue_kat.strip())
                save_json(st.session_state.ws_files["kategorien"], st.session_state.kategorien)
                st.success(f"Kategorie '{neue_kat}' hinzugefügt!")
                st.rerun()

    with st.expander("✏️ Kategorie umbenennen"):
        if st.session_state.kategorien:
            kat_umb     = st.selectbox("Kategorie", st.session_state.kategorien, key="kat_umb")
            neuer_kat_n = st.text_input("Neuer Name", value=kat_umb, key="kat_umb_name")
            if st.button("✏️ Umbenennen", key="btn_kat_umb"):
                if neuer_kat_n.strip() and neuer_kat_n.strip() != kat_umb:
                    idx = st.session_state.kategorien.index(kat_umb)
                    st.session_state.kategorien[idx] = neuer_kat_n.strip()
                    for a in st.session_state.inventar:
                        if a["kategorie"] == kat_umb:
                            a["kategorie"] = neuer_kat_n.strip()
                    save_json(st.session_state.ws_files["kategorien"], st.session_state.kategorien)
                    save_json(st.session_state.ws_files["inventar"], st.session_state.inventar)
                    st.success(f"'{kat_umb}' wurde in '{neuer_kat_n.strip()}' umbenannt!")
                    st.rerun()

    with st.expander("🗑️ Kategorie löschen"):
        if st.session_state.kategorien:
            kat_del = st.selectbox("Kategorie", st.session_state.kategorien, key="kat_loeschen")
            if st.button("Kategorie löschen", key="btn_kat_loeschen"):
                in_use = [a for a in st.session_state.inventar if a["kategorie"] == kat_del]
                if in_use:
                    st.error(f"Kategorie wird noch von {len(in_use)} Artikel(n) verwendet.")
                else:
                    st.session_state.kategorien.remove(kat_del)
                    save_json(st.session_state.ws_files["kategorien"], st.session_state.kategorien)
                    st.success(f"Kategorie '{kat_del}' gelöscht!")
                    st.rerun()

# =============================================
# DETAIL-ANSICHT
# =============================================
if st.session_state.detail_artikel_id is not None:
    artikel_detail = next(
        (a for a in st.session_state.inventar
         if a["id"] == st.session_state.detail_artikel_id), None
    )
    if artikel_detail:
        with st.container(border=True):
            st.markdown(f"## 📄 Artikeldetails – {artikel_detail['name']}")
            col_img, col_info = st.columns([1, 2])

            with col_img:
                img_path = get_ws_image_path(artikel_detail["id"])
                if img_path:
                    st.image(img_path, use_container_width=True)
                else:
                    st.markdown("🖼️ *Kein Bild vorhanden*")
                uploaded = st.file_uploader(
                    "📸 Bild hochladen / ändern",
                    type=["jpg", "jpeg", "png", "webp"],
                    key=f"upload_{artikel_detail['id']}"
                )
                if uploaded:
                    save_ws_image(artikel_detail["id"], uploaded)
                    st.success("Bild gespeichert!")
                    st.rerun()

            with col_info:
                if st.session_state.get("edit_artikel_id") == artikel_detail["id"]:
                    st.markdown("### ✏️ Artikel bearbeiten")
                    with st.form(key=f"edit_form_{artikel_detail['id']}"):
                        e_name = st.text_input("Artikelname", value=artikel_detail["name"])
                        e_kat  = st.selectbox(
                            "Kategorie", st.session_state.kategorien,
                            index=st.session_state.kategorien.index(artikel_detail["kategorie"])
                            if artikel_detail["kategorie"] in st.session_state.kategorien else 0
                        )
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            e_menge = st.number_input("Gesamtmenge", min_value=1, step=1,
                                                       value=int(artikel_detail["menge"]))
                        with ec2:
                            e_preis = st.number_input("Preis (€)", min_value=0.0, step=0.01,
                                                       format="%.2f",
                                                       value=float(artikel_detail["preis"]))
                        e_raum    = st.selectbox(
                            "Raum", st.session_state.raeume,
                            index=st.session_state.raeume.index(artikel_detail["raum"])
                            if artikel_detail["raum"] in st.session_state.raeume else 0
                        )
                        e_barcode = st.text_input("Barcode",
                                                   value=artikel_detail.get("barcode", "") or "")
                        e_notiz   = st.text_area("Notiz",
                                                  value=artikel_detail.get("notiz", "") or "")

                        e_zusatz_werte = {}
                        edit_felder = KATEGORIE_FELDER.get(e_kat, [])
                        if edit_felder:
                            st.markdown(f"**📋 Zusatzfelder – {e_kat}:**")
                            ef_c1, ef_c2 = st.columns(2)
                            for ei, efd in enumerate(edit_felder):
                                with (ef_c1 if ei % 2 == 0 else ef_c2):
                                    ekey    = f"edit_{artikel_detail['id']}_{efd['key']}"
                                    aktuell = artikel_detail.get("zusatz", {}).get(efd["key"], "")
                                    if efd["typ"] == "select":
                                        opt = efd["optionen"]
                                        idx = opt.index(aktuell) if aktuell in opt else 0
                                        e_zusatz_werte[efd["key"]] = st.selectbox(
                                            efd["label"], opt, index=idx, key=ekey)
                                    else:
                                        e_zusatz_werte[efd["key"]] = st.text_input(
                                            efd["label"], value=aktuell,
                                            placeholder=efd.get("placeholder", ""), key=ekey)

                        sc1, sc2 = st.columns(2)
                        with sc1:
                            speichern = st.form_submit_button("💾 Speichern", type="primary",
                                                               use_container_width=True)
                        with sc2:
                            abbrechen = st.form_submit_button("✖️ Abbrechen",
                                                               use_container_width=True)

                    if speichern:
                        if not e_name.strip():
                            st.error("Artikelname darf nicht leer sein!")
                        else:
                            for a in st.session_state.inventar:
                                if a["id"] == artikel_detail["id"]:
                                    diff            = int(e_menge) - a["menge"]
                                    a["name"]       = e_name.strip()
                                    a["kategorie"]  = e_kat
                                    a["menge"]      = int(e_menge)
                                    a["verfuegbar"] = max(0, a.get("verfuegbar", a["menge"]) + diff)
                                    a["preis"]      = round(e_preis, 2)
                                    a["raum"]       = e_raum
                                    a["barcode"]    = e_barcode.strip()
                                    a["notiz"]      = e_notiz.strip()
                                    a["zusatz"]     = e_zusatz_werte
                                    break
                            save_json(st.session_state.ws_files["inventar"],
                                      st.session_state.inventar)
                            st.session_state.edit_artikel_id = None
                            st.success("✅ Artikel gespeichert!")
                            st.rerun()
                    if abbrechen:
                        st.session_state.edit_artikel_id = None
                        st.rerun()

                else:
                    if st.button("✏️ Artikel bearbeiten",
                                 key=f"btn_edit_{artikel_detail['id']}"):
                        st.session_state.edit_artikel_id = artikel_detail["id"]
                        st.rerun()

                    verfuegbar = artikel_detail.get("verfuegbar", artikel_detail["menge"])
                    if verfuegbar == artikel_detail["menge"]:
                        status = "✅ Verfügbar"
                    elif verfuegbar == 0:
                        status = "🔴 Ausgeliehen"
                    else:
                        status = f"⚠️ Teils verfügbar ({verfuegbar}/{artikel_detail['menge']})"

                    zusatz     = artikel_detail.get("zusatz", {})
                    felder_def = KATEGORIE_FELDER.get(artikel_detail["kategorie"], [])
                    if zusatz and felder_def:
                        st.markdown("**📋 Kategorie-Details:**")
                        z_cols = st.columns(min(len(felder_def), 3))
                        for zi, fd in enumerate(felder_def):
                            wert = zusatz.get(fd["key"], "–") or "–"
                            z_cols[zi % 3].metric(fd["label"], wert)
                        st.divider()

                    st.markdown(f"""
| Feld | Wert |
|---|---|
| **ID** | {artikel_detail['id']} |
| **Name** | {artikel_detail['name']} |
| **Kategorie** | {artikel_detail['kategorie']} |
| **Raum** | {artikel_detail['raum']} |
| **Gesamtmenge** | {artikel_detail['menge']} |
| **Verfügbar** | {verfuegbar} |
| **Status** | {status} |
| **Preis** | {artikel_detail['preis']:.2f} € |
| **Gesamtwert** | {artikel_detail['preis'] * artikel_detail['menge']:.2f} € |
| **Barcode** | {artikel_detail.get('barcode', '–') or '–'} |
| **Hinzugefügt am** | {artikel_detail['datum']} |
""")
                    st.markdown("**📝 Notiz:**")
                    neue_notiz = st.text_area(
                        "Notiz bearbeiten",
                        value=artikel_detail.get("notiz", ""),
                        key=f"notiz_{artikel_detail['id']}"
                    )
                    if st.button("💾 Notiz speichern",
                                 key=f"save_notiz_{artikel_detail['id']}"):
                        for a in st.session_state.inventar:
                            if a["id"] == artikel_detail["id"]:
                                a["notiz"] = neue_notiz.strip()
                                break
                        save_json(st.session_state.ws_files["inventar"],
                                  st.session_state.inventar)
                        st.success("Notiz gespeichert!")
                        st.rerun()

            artikel_ausleihen = [a for a in st.session_state.ausleihen
                                  if a["artikel_id"] == artikel_detail["id"]]
            if artikel_ausleihen:
                st.markdown("**📜 Ausleihhistorie dieses Artikels:**")
                df_al = pd.DataFrame(artikel_ausleihen).rename(columns={
                    "person": "Person", "checkout_datum": "Ausgabe",
                    "rueckgabe_erwartet": "Erw. Rückgabe",
                    "rueckgabe_datum": "Rückgabe", "status": "Status"
                })
                st.dataframe(
                    df_al[["Person", "Ausgabe", "Erw. Rückgabe", "Rückgabe", "Status"]],
                    use_container_width=True, hide_index=True
                )

            with st.expander("🖨️ QR-Code anzeigen"):
                qr_buf = generate_qr_code(artikel_detail)
                st.image(qr_buf, width=180)
                st.download_button(
                    "📥 QR-Code herunterladen", qr_buf,
                    f"qr_{artikel_detail['name']}.png", "image/png",
                    key=f"qr_detail_{artikel_detail['id']}"
                )

            if st.button("✖️ Detailansicht schließen", type="secondary"):
                st.session_state.detail_artikel_id = None
                st.session_state.edit_artikel_id   = None
                st.rerun()

    st.divider()

# =============================================
# TABS
# =============================================
tab_namen = ["📋 Inventar", "➕ Artikel hinzufügen"]
if hat_recht("checkout"):           tab_namen.append("🔄 Checkout / Check-in")
if hat_recht("barcode"):            tab_namen.append("📷 Barcode / QR-Code")
if hat_recht("statistiken"):        tab_namen.append("📊 Statistiken")
if hat_recht("export"):             tab_namen.append("📤 Export")
if hat_recht("benutzerverwaltung"): tab_namen.append("👥 Benutzerverwaltung")

tabs      = st.tabs(tab_namen)
tab_index = {name: i for i, name in enumerate(tab_namen)}

# -----------------------------------------------
# TAB: Inventar
# -----------------------------------------------
with tabs[tab_index["📋 Inventar"]]:
    st.subheader(f"📋 Inventar – {st.session_state.aktiver_raum}")

    ansicht_col, fc1, fc2, fc3, fc4 = st.columns([1, 2, 2, 2, 2])
    with ansicht_col:
        ansicht = st.radio("Ansicht", ["🔲 Kacheln", "📄 Tabelle"], horizontal=True,
                           label_visibility="collapsed", key="inventar_ansicht")
    with fc1:
        suche = st.text_input("🔍 Artikel suchen", placeholder="z.B. Laptop...")
    with fc2:
        filter_kat = st.selectbox("Kategorie filtern", ["Alle"] + st.session_state.kategorien)
    with fc3:
        filter_status = st.selectbox("Status filtern",
                                      ["Alle", "✅ Verfügbar", "🔴 Ausgeliehen"])
    with fc4:
        sort_option = st.selectbox("🔃 Sortierung",
                                   ["Standard", "Name ↑", "Name ↓",
                                    "Preis ↑", "Preis ↓",
                                    "Menge ↑", "Menge ↓",
                                    "Datum ↑", "Datum ↓"])

    filter_zusatz = {}
    if filter_kat != "Alle" and filter_kat in KATEGORIE_FELDER:
        select_felder = [f for f in KATEGORIE_FELDER[filter_kat] if f["typ"] == "select"]
        if select_felder:
            st.markdown(f"**🔍 Zusatzfilter für '{filter_kat}':**")
            zf_cols = st.columns(min(len(select_felder), 4))
            for zi, fd in enumerate(select_felder):
                with zf_cols[zi % 4]:
                    auswahl = st.selectbox(
                        fd["label"], ["Alle"] + fd["optionen"],
                        key=f"filter_zusatz_{fd['key']}"
                    )
                    if auswahl != "Alle":
                        filter_zusatz[fd["key"]] = auswahl

    raum_inventar = [
        a for a in st.session_state.inventar
        if a["raum"] == st.session_state.aktiver_raum and hat_raum_zugriff(a["raum"])
    ]
    if suche:
        raum_inventar = [a for a in raum_inventar if suche.lower() in a["name"].lower()]
    if filter_kat != "Alle":
        raum_inventar = [a for a in raum_inventar if a["kategorie"] == filter_kat]
    for fkey, fwert in filter_zusatz.items():
        raum_inventar = [a for a in raum_inventar
                         if a.get("zusatz", {}).get(fkey, "") == fwert]
    if filter_status == "✅ Verfügbar":
        raum_inventar = [a for a in raum_inventar if a.get("verfuegbar", a["menge"]) > 0]
    elif filter_status == "🔴 Ausgeliehen":
        raum_inventar = [a for a in raum_inventar
                         if a.get("verfuegbar", a["menge"]) < a["menge"]]

    # Sortierung
    sort_key_map = {
        "Name ↑":   (lambda a: a["name"].lower(),  False),
        "Name ↓":   (lambda a: a["name"].lower(),  True),
        "Preis ↑":  (lambda a: a["preis"],          False),
        "Preis ↓":  (lambda a: a["preis"],          True),
        "Menge ↑":  (lambda a: a["menge"],          False),
        "Menge ↓":  (lambda a: a["menge"],          True),
        "Datum ↑":  (lambda a: a.get("datum", ""), False),
        "Datum ↓":  (lambda a: a.get("datum", ""), True),
    }
    if sort_option in sort_key_map:
        sk, sr = sort_key_map[sort_option]
        raum_inventar = sorted(raum_inventar, key=sk, reverse=sr)

    if raum_inventar:
        st.markdown(f"**{len(raum_inventar)} Artikel gefunden**")

        if ansicht == "📄 Tabelle":
            # ── TABELLENANSICHT ──────────────────────────────────────────
            tbl_data = []
            for a in raum_inventar:
                verfuegbar = a.get("verfuegbar", a["menge"])
                if verfuegbar == a["menge"]:
                    status_txt = "✅ Verfügbar"
                elif verfuegbar == 0:
                    status_txt = "🔴 Ausgeliehen"
                else:
                    status_txt = f"⚠️ Teils ({verfuegbar}/{a['menge']})"
                tbl_data.append({
                    "#":          a.get("laufnummer", "–"),
                    "Name":       a["name"],
                    "Kategorie":  a["kategorie"],
                    "Raum":       a["raum"],
                    "Menge":      a["menge"],
                    "Verfügbar":  verfuegbar,
                    "Status":     status_txt,
                    "Preis (€)":  f"{a['preis']:.2f}",
                    "Gesamtwert": f"{a['preis'] * a['menge']:.2f} €",
                    "Datum":      a.get("datum", "–"),
                })
            df_tbl = pd.DataFrame(tbl_data)
            sel = st.dataframe(
                df_tbl, use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row"
            )
            # Klick auf Zeile öffnet Detailansicht
            rows = sel.selection.get("rows", []) if hasattr(sel, "selection") else []
            if rows:
                geklickt_id = raum_inventar[rows[0]]["id"]
                st.session_state.detail_artikel_id = geklickt_id
                st.session_state.edit_artikel_id   = None
                st.rerun()

        else:
            # ── KACHELANSICHT ────────────────────────────────────────────
            cols_per_row = 3
            for i in range(0, len(raum_inventar), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, a in enumerate(raum_inventar[i:i + cols_per_row]):
                    with cols[j]:
                        with st.container(border=True):
                            img_path = get_ws_image_path(a["id"])
                            if img_path:
                                st.image(img_path, use_container_width=True)
                            else:
                                st.markdown("🖼️ *Kein Bild*")

                            verfuegbar = a.get("verfuegbar", a["menge"])
                            s_icon = ("✅" if verfuegbar == a["menge"]
                                      else ("🔴" if verfuegbar == 0 else "⚠️"))
                            st.markdown(f"**{a['name']}** {s_icon}")
                            st.caption(f"📂 {a['kategorie']} | 🏠 {a['raum']}")
                            st.caption(
                                f"Menge: {a['menge']} | Verfügbar: {verfuegbar} | {a['preis']:.2f} €")

                            z_karte = a.get("zusatz", {})
                            f_karte = KATEGORIE_FELDER.get(a["kategorie"], [])
                            if z_karte and f_karte:
                                kurzinfo = " | ".join(
                                    f"{fd['label']}: {z_karte[fd['key']]}"
                                    for fd in f_karte[:2]
                                    if z_karte.get(fd["key"])
                                )
                                if kurzinfo:
                                    st.caption(f"📋 {kurzinfo}")

                            cd1, cd2 = st.columns(2)
                            with cd1:
                                if st.button("🔍 Details", key=f"detail_{a['id']}",
                                             use_container_width=True):
                                    st.session_state.detail_artikel_id = a["id"]
                                    st.session_state.edit_artikel_id   = None
                                    st.rerun()
                            with cd2:
                                if st.button("✏️ Bearbeiten", key=f"quickedit_{a['id']}",
                                             use_container_width=True):
                                    st.session_state.detail_artikel_id = a["id"]
                                    st.session_state.edit_artikel_id   = a["id"]
                                    st.rerun()
    else:
        st.info(f"Keine Artikel in Raum '{st.session_state.aktiver_raum}' gefunden.")

    st.divider()
    loeschen_tab, verschieben_tab = st.tabs(["🗑️ Artikel löschen", "📦 Artikel verschieben"])

    with loeschen_tab:
        st.subheader("🗑️ Artikel löschen")
        alle_raum = [a for a in st.session_state.inventar
                     if a["raum"] == st.session_state.aktiver_raum]
        if alle_raum:
            st.markdown("**Mehrfachauswahl:** Wähle einen oder mehrere Artikel zum Löschen aus.")
            art_optionen = {f"#{a.get('laufnummer', '?')} – {a['name']}": a["id"]
                            for a in alle_raum}
            multi_auswahl = st.multiselect(
                "Artikel auswählen", list(art_optionen.keys()),
                key="multi_loeschen"
            )
            if multi_auswahl:
                st.warning(f"⚠️ {len(multi_auswahl)} Artikel werden gelöscht!")
                if st.button(f"❌ {len(multi_auswahl)} Artikel löschen", type="primary"):
                    ids_loeschen = {art_optionen[n] for n in multi_auswahl}
                    st.session_state.inventar = [
                        a for a in st.session_state.inventar
                        if a["id"] not in ids_loeschen
                    ]
                    save_json(st.session_state.ws_files["inventar"], st.session_state.inventar)
                    st.success(f"✅ {len(ids_loeschen)} Artikel gelöscht!")
                    st.rerun()
        else:
            st.info("Keine Artikel zum Löschen vorhanden.")

    with verschieben_tab:
        st.subheader("📦 Artikel zwischen Räumen verschieben")
        alle_raum_v = [a for a in st.session_state.inventar
                       if a["raum"] == st.session_state.aktiver_raum]
        if alle_raum_v and len(st.session_state.raeume) > 1:
            v1, v2, v3 = st.columns([2, 1, 1])
            with v1:
                v_optionen = {f"#{a.get('laufnummer', '?')} – {a['name']}": a["id"]
                              for a in alle_raum_v}
                v_auswahl = st.multiselect(
                    "Artikel auswählen", list(v_optionen.keys()),
                    key="multi_verschieben"
                )
            with v2:
                ziel_raeume = [r for r in st.session_state.raeume
                               if r != st.session_state.aktiver_raum]
                ziel_raum = st.selectbox("🏠 Ziel-Raum", ziel_raeume, key="ziel_raum")
            with v3:
                st.markdown("&nbsp;", unsafe_allow_html=True)
                if st.button("➡️ Verschieben", type="primary",
                             use_container_width=True, key="btn_verschieben"):
                    if v_auswahl:
                        ids_verschieben = {v_optionen[n] for n in v_auswahl}
                        for a in st.session_state.inventar:
                            if a["id"] in ids_verschieben:
                                a["raum"] = ziel_raum
                        save_json(st.session_state.ws_files["inventar"],
                                  st.session_state.inventar)
                        st.success(f"✅ {len(ids_verschieben)} Artikel nach "
                                   f"'{ziel_raum}' verschoben!")
                        st.rerun()
                    else:
                        st.warning("Bitte mindestens einen Artikel auswählen.")
        elif len(st.session_state.raeume) <= 1:
            st.info("Es muss mindestens einen weiteren Raum geben, um Artikel zu verschieben.")
        else:
            st.info("Keine Artikel im aktuellen Raum vorhanden.")

# -----------------------------------------------
# TAB: Artikel hinzufügen
# -----------------------------------------------
with tabs[tab_index["➕ Artikel hinzufügen"]]:
    st.subheader(f"➕ Neuen Artikel hinzufügen – Raum: {st.session_state.aktiver_raum}")

    with st.container(border=True):
        neu_kat = st.selectbox(
            "🗂️ Kategorie auswählen",
            st.session_state.kategorien,
            key="neu_kat_select",
        )

        st.divider()

        h1, h2 = st.columns(2)
        with h1:
            neu_name = st.text_input("📦 Artikelname",
                                      placeholder="z.B. Inception", key="neu_name")
            neu_raum = st.selectbox(
                "🏠 Raum", get_erlaubte_raeume(),
                index=get_erlaubte_raeume().index(st.session_state.aktiver_raum)
                if st.session_state.aktiver_raum in get_erlaubte_raeume() else 0,
                key="neu_raum"
            )
        with h2:
            neu_menge   = st.number_input("🔢 Menge", min_value=1, step=1, key="neu_menge")
            neu_preis   = st.number_input("💶 Preis (€)", min_value=0.0, step=0.01,
                                           format="%.2f", key="neu_preis")
            neu_barcode = st.text_input("🔖 Barcode-Nr. (optional)",
                                         placeholder="z.B. 4012345678901", key="neu_barcode")

        neu_notiz = st.text_area("📝 Notiz (optional)", key="neu_notiz")

        neu_zusatz = {}
        kat_felder = KATEGORIE_FELDER.get(neu_kat, [])
        if kat_felder:
            st.divider()
            st.markdown(f"**📋 Zusatzfelder für '{neu_kat}':**")
            zf1, zf2 = st.columns(2)
            for fi, feld in enumerate(kat_felder):
                with (zf1 if fi % 2 == 0 else zf2):
                    fkey = f"neu_zusatz_{feld['key']}"
                    if feld["typ"] == "select":
                        neu_zusatz[feld["key"]] = st.selectbox(
                            feld["label"], feld["optionen"], key=fkey)
                    else:
                        neu_zusatz[feld["key"]] = st.text_input(
                            feld["label"],
                            placeholder=feld.get("placeholder", ""),
                            key=fkey)

        st.divider()

        if st.button("✅ Artikel hinzufügen", type="primary",
                     use_container_width=True, key="btn_neu_hinzu"):
            if not neu_name.strip():
                st.error("❌ Bitte einen Artikelnamen eingeben!")
            elif not neu_raum:
                st.error("❌ Bitte zuerst einen Raum anlegen!")
            else:
                neuer_artikel = {
                    "id":         str(uuid.uuid4()),
                    "laufnummer": max((a.get("laufnummer", 0)
                                      for a in st.session_state.inventar), default=0) + 1,
                    "name":       neu_name.strip(),
                    "kategorie":  neu_kat,
                    "menge":      int(neu_menge),
                    "verfuegbar": int(neu_menge),
                    "raum":       neu_raum,
                    "preis":      round(neu_preis, 2),
                    "barcode":    neu_barcode.strip(),
                    "notiz":      neu_notiz.strip(),
                    "zusatz":     neu_zusatz,
                    "datum":      datetime.now().strftime("%d.%m.%Y %H:%M"),
                }
                st.session_state.inventar.append(neuer_artikel)
                save_json(st.session_state.ws_files["inventar"], st.session_state.inventar)
                st.success(f"✅ '{neu_name}' wurde in '{neu_raum}' hinzugefügt!")
                st.balloons()

# -----------------------------------------------
# TAB: Checkout / Check-in
# -----------------------------------------------
if hat_recht("checkout"):
    with tabs[tab_index["🔄 Checkout / Check-in"]]:
        st.subheader("🔄 Checkout & Check-in System")
        co1, co2, co3 = st.tabs(["📤 Checkout", "📥 Check-in", "📜 Historie"])

        with co1:
            st.markdown("### 📤 Artikel ausleihen")
            verf = [a for a in st.session_state.inventar
                    if a.get("verfuegbar", a["menge"]) > 0 and hat_raum_zugriff(a["raum"])]
            if verf:
                cc1, cc2 = st.columns(2)
                with cc1:
                    labels  = [
                        f"#{a.get('laufnummer', '?')} – {a['name']} ({a['raum']}) | "
                        f"Verfügbar: {a.get('verfuegbar', a['menge'])}"
                        for a in verf
                    ]
                    co_aus  = st.selectbox("Artikel", labels, key="co_artikel")
                    art_obj = verf[labels.index(co_aus)]
                    co_menge = st.number_input(
                        "Menge", min_value=1,
                        max_value=art_obj.get("verfuegbar", art_obj["menge"]),
                        step=1, key="co_menge"
                    )
                with cc2:
                    co_person = st.text_input("Name der Person",
                                               placeholder="z.B. Max Mustermann")
                    co_datum  = st.date_input("Ausleihdatum",       value=datetime.today())
                    co_rueck  = st.date_input("Erwartete Rückgabe", value=datetime.today())
                    co_notiz  = st.text_area("Notiz (optional)", key="co_notiz")

                if st.button("📤 Auschecken", type="primary"):
                    if not co_person.strip():
                        st.error("Bitte den Namen der Person eingeben!")
                    else:
                        st.session_state.ausleihen.append({
                            "ausleihe_id":        str(uuid.uuid4()),
                            "artikel_id":         art_obj["id"],
                            "artikel_name":       art_obj["name"],
                            "raum":               art_obj["raum"],
                            "menge":              int(co_menge),
                            "person":             co_person.strip(),
                            "checkout_datum":     co_datum.strftime("%d.%m.%Y"),
                            "rueckgabe_erwartet": co_rueck.strftime("%d.%m.%Y"),
                            "rueckgabe_datum":    None,
                            "status":             "ausgeliehen",
                            "notiz":              co_notiz.strip(),
                        })
                        save_json(st.session_state.ws_files["ausleihen"],
                                  st.session_state.ausleihen)
                        for a in st.session_state.inventar:
                            if a["id"] == art_obj["id"]:
                                a["verfuegbar"] = (a.get("verfuegbar", a["menge"])
                                                   - int(co_menge))
                                break
                        save_json(st.session_state.ws_files["inventar"],
                                  st.session_state.inventar)
                        st.success(f"✅ '{art_obj['name']}' ({co_menge}x) an "
                                   f"**{co_person}** ausgeliehen.")
                        st.rerun()

                st.divider()
                st.markdown("### 📋 Aktuell ausgeliehene Artikel")
                aktiv = [a for a in st.session_state.ausleihen if a["status"] == "ausgeliehen"]
                if aktiv:
                    df_aktiv = pd.DataFrame(aktiv).rename(columns={
                        "ausleihe_id": "ID", "artikel_name": "Artikel", "raum": "Raum",
                        "menge": "Menge", "person": "Ausgeliehen an",
                        "checkout_datum": "Ausgabe",
                        "rueckgabe_erwartet": "Rückgabe erwartet", "notiz": "Notiz"
                    })
                    st.dataframe(
                        df_aktiv[["ID", "Artikel", "Raum", "Menge", "Ausgeliehen an",
                                  "Ausgabe", "Rückgabe erwartet", "Notiz"]],
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.info("Keine aktiven Ausleihen.")
            else:
                st.info("Keine verfügbaren Artikel zum Ausleihen vorhanden.")

        with co2:
            st.markdown("### 📥 Artikel zurückgeben")
            aktive = [a for a in st.session_state.ausleihen if a["status"] == "ausgeliehen"]
            if aktive:
                ci_labels = [
                    f"#{a['ausleihe_id'][:8]} – {a['artikel_name']} ({a['menge']}x) "
                    f"→ {a['person']} | Ausgabe: {a['checkout_datum']}"
                    for a in aktive
                ]
                ci_aus    = st.selectbox("Ausleihe auswählen", ci_labels, key="ci_auswahl")
                ci_obj    = aktive[ci_labels.index(ci_aus)]

                ci_c1, ci_c2 = st.columns(2)
                with ci_c1:
                    st.info(
                        f"**Artikel:** {ci_obj['artikel_name']}  \n"
                        f"**Ausgeliehen an:** {ci_obj['person']}  \n"
                        f"**Ausgabe:** {ci_obj['checkout_datum']}  \n"
                        f"**Erwartete Rückgabe:** {ci_obj['rueckgabe_erwartet']}"
                    )
                with ci_c2:
                    ci_datum = st.date_input("Rückgabedatum", value=datetime.today(),
                                              key="ci_datum")
                    ci_notiz = st.text_area("Notiz (optional)", key="ci_notiz")

                if st.button("📥 Einchecken", type="primary"):
                    for a in st.session_state.ausleihen:
                        if a["ausleihe_id"] == ci_obj["ausleihe_id"]:
                            a["status"]          = "zurückgegeben"
                            a["rueckgabe_datum"] = ci_datum.strftime("%d.%m.%Y")
                            if ci_notiz.strip():
                                a["notiz"] += f" | Rückgabe: {ci_notiz.strip()}"
                            break
                    save_json(st.session_state.ws_files["ausleihen"],
                              st.session_state.ausleihen)
                    for a in st.session_state.inventar:
                        if a["id"] == ci_obj["artikel_id"]:
                            a["verfuegbar"] = a.get("verfuegbar", 0) + ci_obj["menge"]
                            break
                    save_json(st.session_state.ws_files["inventar"],
                              st.session_state.inventar)
                    st.success(f"✅ '{ci_obj['artikel_name']}' erfolgreich zurückgebucht!")
                    st.rerun()
            else:
                st.info("Keine aktiven Ausleihen zur Rückgabe vorhanden.")

        with co3:
            st.markdown("### 📜 Vollständige Ausleihhistorie")
            hc1, hc2 = st.columns(2)
            with hc1:
                h_person = st.text_input("🔍 Nach Person suchen")
            with hc2:
                h_status = st.selectbox("Status", ["Alle", "ausgeliehen", "zurückgegeben"])

            hist = st.session_state.ausleihen
            if h_person:
                hist = [a for a in hist if h_person.lower() in a["person"].lower()]
            if h_status != "Alle":
                hist = [a for a in hist if a["status"] == h_status]

            if hist:
                df_hist = pd.DataFrame(hist).rename(columns={
                    "ausleihe_id": "ID", "artikel_name": "Artikel", "raum": "Raum",
                    "menge": "Menge", "person": "Person",
                    "checkout_datum": "Ausgabe",
                    "rueckgabe_erwartet": "Erw. Rückgabe",
                    "rueckgabe_datum": "Rückgabe",
                    "status": "Status", "notiz": "Notiz"
                })
                st.dataframe(df_hist, use_container_width=True, hide_index=True)
                csv_hist = df_hist.to_csv(index=False, sep=";", encoding="utf-8-sig")
                               st.download_button("📥 Historie als CSV", csv_hist,
                                   "ausleihhistorie.csv", "text/csv")

                st.divider()
                st.markdown("#### 🗑️ Einträge löschen")

                with st.expander("🗑️ Einzelne Einträge löschen"):
                    losch_labels = {

                          f"{a['checkout_datum']} | {a['artikel_name']} → {a['person']} [{a['status']}]": a["ausleihe_id"]
                          for a in hist
                      }
                      multi_hist = st.multiselect(
                          "Einträge auswählen",
                          list(losch_labels.keys()),
                          key="hist_multi_loeschen"
                      )
                      if multi_hist:
                          st.warning(f"⚠️ {len(multi_hist)} Eintrag/Einträge werden gelöscht!")
                          if st.button(f"❌ {len(multi_hist)} Eintrag/Einträge löschen",
                                       type="primary", key="btn_hist_einzel_del"):
                              ids_del = {losch_labels[l] for l in multi_hist}
                              st.session_state.ausleihen = [
                                  a for a in st.session_state.ausleihen
                                  if a["ausleihe_id"] not in ids_del
                              ]
                              save_json(st.session_state.ws_files["ausleihen"],
                                        st.session_state.ausleihen)
                              st.success(f"✅ {len(ids_del)} Eintrag/Einträge gelöscht!")
                              st.rerun()

                  with st.expander("🔥 Gesamte Historie löschen"):
                      st.error(
                          "⚠️ **Achtung:** Dies löscht **alle** Einträge der Ausleihhistorie "
                          "unwiderruflich."
                      )
                      bestaetigung = st.text_input(
                          "Tippe LÖSCHEN zur Bestätigung", key="hist_gesamt_confirm"
                      )
                      if st.button("🔥 Gesamte Historie löschen", type="primary",
                                   key="btn_hist_gesamt_del"):
                          if bestaetigung.strip() == "LÖSCHEN":
                              st.session_state.ausleihen = []
                              save_json(st.session_state.ws_files["ausleihen"], [])
                              st.success("✅ Gesamte Historie gelöscht!")
                              st.rerun()
                          else:
                              st.error("❌ Bitte tippe 'LÖSCHEN' zur Bestätigung.")
            else:
                st.info("Keine Einträge gefunden.")

# -----------------------------------------------
# TAB: Barcode / QR-Code
# -----------------------------------------------
if hat_recht("barcode"):
    with tabs[tab_index["📷 Barcode / QR-Code"]]:
        st.subheader("📷 Barcode & QR-Code")
        bc1, bc2 = st.tabs(["🔍 Barcode scannen", "🖨️ QR-Code generieren"])

        with bc1:
            st.markdown("Scanne einen Barcode oder gib ihn manuell ein.")
            try:
                scan_result = quagga()
                if scan_result and scan_result.get("codeResult"):
                    scanned = scan_result["codeResult"]["code"]
                    st.success(f"✅ Erkannter Barcode: `{scanned}`")
                    gefunden = [a for a in st.session_state.inventar
                                if a.get("barcode") == scanned]
                    if gefunden:
                        st.dataframe(pd.DataFrame(gefunden), use_container_width=True,
                                     hide_index=True)
                    else:
                        st.warning("Kein Artikel mit diesem Barcode gefunden.")
            except Exception:
                st.info("💡 Barcode-Scanner nicht verfügbar. "
                        "Bitte installiere: `pip install streamlit-quagga`")
                manuell = st.text_input("Barcode manuell eingeben",
                                         placeholder="z.B. 4012345678901")
                if manuell:
                    gefunden = [a for a in st.session_state.inventar
                                if a.get("barcode") == manuell]
                    if gefunden:
                        st.success("✅ Artikel gefunden:")
                        st.dataframe(pd.DataFrame(gefunden), use_container_width=True,
                                     hide_index=True)
                    else:
                        st.warning("Kein Artikel gefunden.")

        with bc2:
            st.markdown("Generiere einen QR-Code für jeden Artikel zum Ausdrucken.")
            if st.session_state.inventar:
                qr_labels = [
                    f"#{a.get('laufnummer', '?')} – {a['name']} ({a['raum']})"
                    for a in st.session_state.inventar
                ]
                qr_aus = st.selectbox("Artikel für QR-Code", qr_labels)
                qr_obj = st.session_state.inventar[qr_labels.index(qr_aus)]
                qr_buf = generate_qr_code(qr_obj)
                st.image(qr_buf, caption=f"QR-Code: {qr_obj['name']}", width=200)
                st.download_button(
                    "📥 QR-Code herunterladen", qr_buf,
                    f"qr_{qr_obj['name']}.png", "image/png",
                    key=f"qr_tab_{qr_obj['id']}"
                )
            else:
                st.info("Noch keine Artikel vorhanden.")

# -----------------------------------------------
# TAB: Statistiken
# -----------------------------------------------
if hat_recht("statistiken"):
    with tabs[tab_index["📊 Statistiken"]]:
        st.subheader("📊 Statistiken")
        gs_col, rm_col = st.columns(2)

        with gs_col:
            st.markdown("### 🌐 Gesamtübersicht")
            if st.session_state.inventar:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("📦 Artikel",      len(st.session_state.inventar))
                m2.metric("🔢 Gesamtmenge",  sum(a["menge"] for a in st.session_state.inventar))
                m3.metric("💶 Gesamtwert",
                           f"{sum(a['preis'] * a['menge'] for a in st.session_state.inventar):.2f} €")
                aktiv_ausl = len([a for a in st.session_state.ausleihen
                                  if a["status"] == "ausgeliehen"])
                m4.metric("🔴 Aktive Ausleihen", aktiv_ausl)

                st.markdown("#### Artikel pro Raum")
                raum_stats = {}
                for a in st.session_state.inventar:
                    raum_stats[a["raum"]] = raum_stats.get(a["raum"], 0) + 1
                st.bar_chart(raum_stats)

                st.markdown("#### Artikel pro Kategorie")
                kat_stats = {}
                for a in st.session_state.inventar:
                    kat_stats[a["kategorie"]] = kat_stats.get(a["kategorie"], 0) + 1
                st.bar_chart(kat_stats)
            else:
                st.info("Noch keine Artikel vorhanden.")

        with rm_col:
            st.markdown(f"### 🏠 Raum: {st.session_state.aktiver_raum}")
            raum_art = [a for a in st.session_state.inventar
                        if a["raum"] == st.session_state.aktiver_raum]
            if raum_art:
                r1, r2, r3 = st.columns(3)
                r1.metric("📦 Artikel", len(raum_art))
                r2.metric("🔢 Menge",   sum(a["menge"] for a in raum_art))
                r3.metric("💶 Wert",
                           f"{sum(a['preis'] * a['menge'] for a in raum_art):.2f} €")
            else:
                st.info(f"Keine Artikel in '{st.session_state.aktiver_raum}'.")

# -----------------------------------------------
# TAB: Export / Import
# -----------------------------------------------
if hat_recht("export"):
    with tabs[tab_index["📤 Export"]]:
        st.subheader("📤 Import & Export")
        imp_tab, exp_tab = st.tabs(["📥 Import", "📤 Export"])

        with imp_tab:
            st.markdown("### 📥 Artikel aus Excel oder CSV importieren")

            st.markdown("#### 1️⃣ Vorlage herunterladen")
            vorlage_df = pd.DataFrame(columns=[
                "name", "kategorie", "menge", "preis",
                "raum", "barcode", "notiz", "mindestmenge"
            ])
            vl1, vl2 = st.columns(2)
            with vl1:
                if XLSX_AVAILABLE:
                    vl_buf = io.BytesIO()
                    with pd.ExcelWriter(vl_buf, engine="openpyxl") as writer:
                        vorlage_df.to_excel(writer, index=False, sheet_name="Inventar")
                    vl_buf.seek(0)
                    st.download_button(
                        "📊 Excel-Vorlage", data=vl_buf,
                        file_name="inventar_vorlage.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            with vl2:
                vl_csv = vorlage_df.to_csv(index=False, sep=";", encoding="utf-8-sig")
                st.download_button(
                    "📄 CSV-Vorlage", data=vl_csv,
                    file_name="inventar_vorlage.csv", mime="text/csv",
                    use_container_width=True
                )
            st.caption("Pflichtfelder: **name**, **menge** | "
                       "Optional: kategorie, preis, raum, barcode, notiz, mindestmenge")

            st.divider()
            st.markdown("#### 2️⃣ Datei hochladen")
            upload_file = st.file_uploader(
                "Excel (.xlsx) oder CSV auswählen",
                type=["xlsx", "csv"], key="import_file"
            )
            import_raum_override = st.selectbox(
                "📍 Ziel-Raum (überschreibt 'raum'-Spalte wenn gesetzt)",
                ["— Spalte aus Datei verwenden —"] + st.session_state.raeume,
                key="import_raum"
            )
            duplikat_verhalten = st.radio(
                "🔁 Bei doppeltem Artikelname",
                ["Überspringen", "Trotzdem hinzufügen"],
                horizontal=True, key="duplikat_verhalten"
            )

            if upload_file is not None:
                try:
                    if upload_file.name.endswith(".xlsx"):
                        df_import = pd.read_excel(upload_file, dtype=str)
                    else:
                        raw_bytes = upload_file.read()
                        df_import = None
                        for enc in ["utf-8-sig", "utf-8", "latin-1", "cp1252", "iso-8859-1"]:
                            try:
                                import io as _io
                                df_import = pd.read_csv(
                                    _io.BytesIO(raw_bytes),
                                    sep=None, engine="python",
                                    dtype=str, encoding=enc
                                )
                                break
                            except Exception:
                                continue
                        if df_import is None:
                            st.error("❌ Die Datei konnte mit keinem bekannten Encoding "
                                     "gelesen werden.")
                            st.stop()

                    df_import.columns = [c.strip().lower() for c in df_import.columns]
                    df_import = df_import.fillna("")

                    st.markdown("#### 3️⃣ Vorschau")
                    st.dataframe(df_import, use_container_width=True, hide_index=True)
                    st.markdown(f"**{len(df_import)} Zeilen** bereit zum Import.")

                    fehlende = [f for f in ["name", "menge"] if f not in df_import.columns]
                    if fehlende:
                        st.error(f"❌ Pflichtfelder fehlen: {', '.join(fehlende)}")
                    else:
                        if st.button("✅ Import starten", type="primary", key="btn_import"):
                            bestehende = [a["name"].lower() for a in st.session_state.inventar]
                            importiert = uebersprungen = 0

                            for _, row in df_import.iterrows():
                                art_name = str(row.get("name", "")).strip()
                                if not art_name:
                                    uebersprungen += 1
                                    continue
                                if (duplikat_verhalten == "Überspringen"
                                        and art_name.lower() in bestehende):
                                    uebersprungen += 1
                                    continue

                                if import_raum_override != "— Spalte aus Datei verwenden —":
                                    raum_val = import_raum_override
                                else:
                                    raum_val = str(row.get("raum", "")).strip() or \
                                               st.session_state.aktiver_raum
                                    if raum_val not in st.session_state.raeume:
                                        st.session_state.raeume.append(raum_val)
                                        save_json(st.session_state.ws_files["raeume"],
                                                  st.session_state.raeume)

                                kat_val = str(row.get("kategorie", "")).strip() or "Sonstiges"
                                if kat_val not in st.session_state.kategorien:
                                    st.session_state.kategorien.append(kat_val)
                                    save_json(st.session_state.ws_files["kategorien"],
                                              st.session_state.kategorien)

                                try:
                                    menge_val = max(1, int(float(
                                        str(row.get("menge", 1)).replace(",", "."))))
                                except Exception:
                                    menge_val = 1
                                try:
                                    preis_val = round(float(
                                        str(row.get("preis", 0)).replace(",", ".")), 2)
                                except Exception:
                                    preis_val = 0.0

                                neuer = {
                                    "id":         str(uuid.uuid4()),
                                    "laufnummer": max((a.get("laufnummer", 0)
                                                       for a in st.session_state.inventar),
                                                      default=0) + importiert + 1,
                                    "name":       art_name,
                                    "kategorie":  kat_val,
                                    "menge":      menge_val,
                                    "verfuegbar": menge_val,
                                    "raum":       raum_val,
                                    "preis":      preis_val,
                                    "barcode":    str(row.get("barcode", "")).strip(),
                                    "notiz":      str(row.get("notiz", "")).strip(),
                                    "zusatz":     {},
                                    "datum":      datetime.now().strftime("%d.%m.%Y %H:%M"),
                                }
                                st.session_state.inventar.append(neuer)
                                bestehende.append(art_name.lower())
                                importiert += 1

                            save_json(st.session_state.ws_files["inventar"],
                                      st.session_state.inventar)
                            if importiert > 0:
                                st.success(f"✅ {importiert} Artikel erfolgreich importiert!")
                            if uebersprungen > 0:
                                st.warning(f"⚠️ {uebersprungen} Zeilen übersprungen.")
                            st.rerun()

                except Exception as e:
                    st.error(f"❌ Fehler beim Lesen der Datei: {e}")

        with exp_tab:
            st.markdown("### 📤 Daten exportieren")
            ex1, ex2 = st.columns(2)
            with ex1:
                exp_option = st.radio(
                    "Was möchtest du exportieren?",
                    ["Aktiver Raum", "Alle Räume", "Ausleihhistorie"],
                    key="export_option"
                )
            with ex2:
                exp_format = st.radio(
                    "Format",
                    ["📄 CSV", "📊 Excel (.xlsx)"] if XLSX_AVAILABLE else ["📄 CSV"],
                    key="export_format"
                )

            if exp_option == "Aktiver Raum":
                exp_data      = [a for a in st.session_state.inventar
                                 if a["raum"] == st.session_state.aktiver_raum]
                filename_base = f"inventar_{st.session_state.aktiver_raum}"
            elif exp_option == "Alle Räume":
                exp_data      = st.session_state.inventar
                filename_base = "inventar_gesamt"
            else:
                exp_data      = st.session_state.ausleihen
                filename_base = "ausleihhistorie"

            if exp_data:
                df_exp = pd.DataFrame(exp_data)
                if exp_format == "📄 CSV" or not XLSX_AVAILABLE:
                    csv_exp = df_exp.to_csv(index=False, sep=";", encoding="utf-8-sig")
                    st.download_button(
                        f"📥 CSV herunterladen ({len(exp_data)} Einträge)",
                        data=csv_exp, file_name=f"{filename_base}.csv", mime="text/csv"
                    )
                else:
                    xlsx_buf = io.BytesIO()
                    with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as writer:
                        df_exp.to_excel(writer, index=False, sheet_name="Export")
                    xlsx_buf.seek(0)
                    st.download_button(
                        f"📥 Excel herunterladen ({len(exp_data)} Einträge)",
                        data=xlsx_buf, file_name=f"{filename_base}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("Keine Daten zum Exportieren.")

# -----------------------------------------------
# TAB: Benutzerverwaltung
# -----------------------------------------------
if hat_recht("benutzerverwaltung"):
    with tabs[tab_index["👥 Benutzerverwaltung"]]:
        st.subheader("👥 Benutzerverwaltung")

        ws_aktuell = st.session_state.aktiver_workspace
        all_ws     = load_workspaces()
        ws_idx     = next((i for i, w in enumerate(all_ws)
                           if w["id"] == ws_aktuell["id"]), None)

        bv1, bv2, bv3, bv4 = st.tabs(["👥 Mitglieder", "🔗 Workspace teilen",
                                "⚙️ Workspace-Einstellungen", "🔑 Passwort ändern"])

        with bv1:
            st.markdown("### 👥 Mitglieder dieses Workspaces")
            users = load_users()

            # FIX: Duplikate automatisch bereinigen
            mitglieder_raw = ws_aktuell.get("mitglieder", [])
            mitglieder = list(dict.fromkeys(mitglieder_raw))
            if len(mitglieder) != len(mitglieder_raw):
                _fix = load_workspaces()
                _i = next((i for i, w in enumerate(_fix) if w["id"] == ws_aktuell["id"]), None)
                if _i is not None:
                    _fix[_i]["mitglieder"] = mitglieder
                    save_workspaces(_fix)

            mitglieder_rollen = ws_aktuell.get("mitglieder_rollen", {})
            mitglieder_info   = ws_aktuell.get("mitglieder_info", {})

            besitzer_obj = next((u for u in users
                                 if u["benutzername"] == ws_aktuell["besitzer"]), {})
            alle_mitgl = [{
                "Benutzername":   ws_aktuell["besitzer"],
                "Name":           besitzer_obj.get("name", ws_aktuell["besitzer"]),
                "Rolle":          "👑 Besitzer",
                "Erlaubte Räume": "🌐 Alle Räume",
            }]
            for m in mitglieder:
                m_obj  = next((u for u in users if u["benutzername"] == m), {})
                m_info = mitglieder_info.get(m, {})
                alle_mitgl.append({
                    "Benutzername":   m,
                    "Name":           m_obj.get("name", m),
                    "Rolle":          ROLLEN_NAMEN.get(mitglieder_rollen.get(m, "user"), "user"),
                    "Erlaubte Räume": ", ".join(m_info.get("erlaubte_raeume", [])) or "🌐 Alle Räume",
                })
            st.dataframe(pd.DataFrame(alle_mitgl), use_container_width=True, hide_index=True)

            if mitglieder:
                st.divider()
                st.markdown("#### ❌ Mitglied entfernen")
                entf_aus = st.selectbox("Mitglied", mitglieder, key="entfernen_auswahl")
                if st.button("Mitglied entfernen", key="btn_entfernen"):
                    # FIX: Sicheres Entfernen ohne ValueError
                    all_ws[ws_idx]["mitglieder"] = [
                        m for m in all_ws[ws_idx]["mitglieder"] if m != entf_aus
                    ]
                    all_ws[ws_idx].get("mitglieder_rollen", {}).pop(entf_aus, None)
                    all_ws[ws_idx].get("mitglieder_info",   {}).pop(entf_aus, None)
                    save_workspaces(all_ws)
                    st.success(f"'{entf_aus}' wurde entfernt!")
                    st.rerun()

                st.divider()
                st.markdown("#### 🏠 Raumzugriff & Rolle bearbeiten")
                em1, em2 = st.columns(2)
                with em1:
                    edit_m       = st.selectbox("Mitglied", mitglieder, key="edit_mitglied")
                    aktuelle_r   = mitglieder_rollen.get(edit_m, "user")
                    rolle_optionen = ["manager", "user", "viewer"]
                    neue_r       = st.selectbox(
                        "Rolle", rolle_optionen,
                        format_func=lambda r: ROLLEN_NAMEN[r],
                        index=rolle_optionen.index(aktuelle_r)
                        if aktuelle_r in rolle_optionen else 1,
                        key="edit_rolle"
                    )
                with em2:
                    akt_raeume  = [r for r in mitglieder_info.get(edit_m, {}).get("erlaubte_raeume", []) if r in st.session_state.raeume]
                    neue_raeume = st.multiselect(
                        "Erlaubte Räume (leer = alle)",
                          st.session_state.raeume,
                          default=[r for r in akt_raeume if r in st.session_state.raeume],
                        key="edit_raeume"
                    )
                if st.button("💾 Änderungen speichern", key="btn_edit_mitglied"):
                    all_ws[ws_idx].setdefault("mitglieder_rollen", {})[edit_m] = neue_r
                    all_ws[ws_idx].setdefault("mitglieder_info",   {}).setdefault(
                        edit_m, {})["erlaubte_raeume"] = neue_raeume
                    save_workspaces(all_ws)
                    st.success("Änderungen gespeichert!")
                    st.rerun()

        with bv2:
            st.markdown("### 🔗 Workspace teilen")
            st.markdown("Teile die folgende Workspace-ID mit Personen, die beitreten sollen:")
            st.code(ws_aktuell["id"], language=None)
            st.caption("Die Person kann die ID auf der Workspace-Auswahlseite unter "
                       "'Workspace beitreten' eingeben.")

            st.divider()
            st.markdown("#### ➕ Benutzer direkt hinzufügen")
            users = load_users()
            nicht_mitgl = [u for u in users
                           if u["benutzername"] != ws_aktuell["besitzer"]
                           and u["benutzername"] not in ws_aktuell.get("mitglieder", [])]
            if nicht_mitgl:
                dh_labels = [f"{u['name']} ({u['benutzername']})" for u in nicht_mitgl]
                dh_aus    = st.selectbox("Benutzer auswählen", dh_labels, key="direkt_hinzu")
                dh_user   = nicht_mitgl[dh_labels.index(dh_aus)]
                dh_rolle  = st.selectbox(
                    "Rolle zuweisen", ["manager", "user", "viewer"],
                    format_func=lambda r: ROLLEN_NAMEN[r], key="direkt_rolle"
                )
                dh_raeume = st.multiselect(
                    "Erlaubte Räume (leer = alle)",
                    st.session_state.raeume, key="direkt_raeume"
                )
                if st.button("✅ Hinzufügen", key="btn_direkt_hinzu"):
                    # FIX: Doppeltes Hinzufügen verhindern
                    if dh_user["benutzername"] not in all_ws[ws_idx].get("mitglieder", []):
                        all_ws[ws_idx].setdefault("mitglieder", []).append(
                            dh_user["benutzername"])
                        all_ws[ws_idx].setdefault("mitglieder_rollen", {})[
                            dh_user["benutzername"]] = dh_rolle
                        all_ws[ws_idx].setdefault("mitglieder_info", {})[
                            dh_user["benutzername"]] = {"erlaubte_raeume": dh_raeume}
                        save_workspaces(all_ws)
                        st.success(f"'{dh_user['name']}' wurde hinzugefügt!")
                        st.rerun()
                    else:
                        st.warning("⚠️ Benutzer ist bereits Mitglied!")
            else:
                st.info("Alle registrierten Benutzer sind bereits Mitglied.")

        with bv3:
            st.markdown("### ⚙️ Workspace-Einstellungen")
            neuer_ws_name  = st.text_input("Workspace-Name", value=ws_aktuell["name"])
            neue_ws_beschr = st.text_input("Beschreibung",
                                            value=ws_aktuell.get("beschreibung", ""))
            neues_ws_pw    = st.text_input(
                "Neues Passwort (leer lassen = unverändert)", type="password")
            pw_entfernen   = st.checkbox("Passwortschutz entfernen")

            if st.button("💾 Einstellungen speichern"):
                all_ws[ws_idx]["name"]         = neuer_ws_name.strip()
                all_ws[ws_idx]["beschreibung"]  = neue_ws_beschr.strip()
                if pw_entfernen:
                    all_ws[ws_idx]["passwort"] = None
                elif neues_ws_pw.strip():
                    all_ws[ws_idx]["passwort"] = hash_passwort(neues_ws_pw.strip())
                save_workspaces(all_ws)
                st.session_state.aktiver_workspace = all_ws[ws_idx]
                st.success("Einstellungen gespeichert!")
                st.rerun()
        with bv4:
            st.markdown("### 🔑 Eigenes Passwort ändern")
            with st.form("pw_aendern_form"):
                pw_alt  = st.text_input("🔐 Aktuelles Passwort",        type="password")
                pw_neu1 = st.text_input("🔑 Neues Passwort",             type="password")
                pw_neu2 = st.text_input("🔑 Neues Passwort wiederholen", type="password")
                pw_btn  = st.form_submit_button("💾 Passwort ändern", type="primary",
                                                use_container_width=True)
            if pw_btn:
                users_pw = load_users()
                user_pw_obj = next((u for u in users_pw
                                    if u["benutzername"] == st.session_state.benutzername), None)
                if not user_pw_obj:
                    st.error("❌ Benutzer nicht gefunden.")
                elif user_pw_obj["passwort"] != hash_passwort(pw_alt):
                    st.error("❌ Das aktuelle Passwort ist falsch!")
                elif len(pw_neu1) < 6:
                    st.error("❌ Das neue Passwort muss mindestens 6 Zeichen haben!")
                elif pw_neu1 != pw_neu2:
                    st.error("❌ Die neuen Passwörter stimmen nicht überein!")
                else:
                    user_pw_obj["passwort"] = hash_passwort(pw_neu1)
                    save_json(USERS_FILE, users_pw)
                    st.success("✅ Passwort erfolgreich geändert!")

            st.divider()
            st.markdown("#### 🗑️ Workspace löschen")
            st.warning("⚠️ Diese Aktion kann nicht rückgängig gemacht werden!")
            if st.button("🗑️ Workspace endgültig löschen", type="primary"):
                all_ws = [w for w in all_ws if w["id"] != ws_aktuell["id"]]
                save_workspaces(all_ws)
                st.session_state.aktiver_workspace = None
                for k in ["inventar", "raeume", "kategorien", "ausleihen",
                          "ws_files", "aktiver_raum"]:
                    st.session_state.pop(k, None)
                st.rerun()

st.markdown("---")
st.markdown("🤖 **EVA** – Inventarisierungs-App | Erstellt mit Python & Streamlit")
