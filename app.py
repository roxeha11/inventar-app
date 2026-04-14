import streamlit as st
import pandas as pd
import json
import os
import qrcode
import io
import base64
from PIL import Image
from datetime import datetime
from streamlit_quagga import quagga

# Dateien
DATA_FILE = "inventar.json"
ROOMS_FILE = "raeume.json"
CATEGORIES_FILE = "kategorien.json"
LENDING_FILE = "ausleihen.json"
IMAGES_DIR = "artikel_bilder"
os.makedirs(IMAGES_DIR, exist_ok=True)

DEFAULT_KATEGORIEN = ["Elektronik", "Möbel", "Bürobedarf", "Werkzeug", "Sonstiges"]

# --- Datenverwaltung ---
def load_json(file, default):
    if os.path.exists(file):
        with open(file, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(file, data):
    with open(file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def save_image(artikel_id, uploaded_file):
    """Speichert ein Bild für einen Artikel."""
    ext = uploaded_file.name.split(".")[-1].lower()
    path = os.path.join(IMAGES_DIR, f"{artikel_id}.{ext}")
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path

def get_image_path(artikel_id):
    """Gibt den Bildpfad eines Artikels zurück (falls vorhanden)."""
    for ext in ["jpg", "jpeg", "png", "webp"]:
        path = os.path.join(IMAGES_DIR, f"{artikel_id}.{ext}")
        if os.path.exists(path):
            return path
    return None

def generate_qr_code(artikel):
    info = f"ID: {artikel['id']}\nName: {artikel['name']}\nRaum: {artikel['raum']}\nMenge: {artikel['menge']}"
    qr = qrcode.QRCode(version=1, box_size=6, border=2)
    qr.add_data(info)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

# --- Session State Initialisierung ---
if "inventar" not in st.session_state:
    st.session_state.inventar = load_json(DATA_FILE, [])
if "raeume" not in st.session_state:
    st.session_state.raeume = load_json(ROOMS_FILE, ["Lager", "Büro", "Keller"])
if "kategorien" not in st.session_state:
    st.session_state.kategorien = load_json(CATEGORIES_FILE, DEFAULT_KATEGORIEN)
if "ausleihen" not in st.session_state:
    st.session_state.ausleihen = load_json(LENDING_FILE, [])
if "aktiver_raum" not in st.session_state:
    st.session_state.aktiver_raum = st.session_state.raeume[0] if st.session_state.raeume else None
if "detail_artikel_id" not in st.session_state:
    st.session_state.detail_artikel_id = None

# --- Seitenlayout ---
st.set_page_config(page_title="Inventarisierungs-App", page_icon="📦", layout="wide")
st.title("📦 Inventarisierungs-App")
st.markdown("Raumweise Inventarverwaltung – mit Kategorien, Barcode/QR-Code & Checkout-System.")

# =============================================
# SIDEBAR
# =============================================
with st.sidebar:

    # --- Raumverwaltung ---
    st.header("🏠 Raumverwaltung")
    if st.session_state.raeume:
        aktiver_raum = st.selectbox(
            "Aktiver Raum",
            st.session_state.raeume,
            index=st.session_state.raeume.index(st.session_state.aktiver_raum)
            if st.session_state.aktiver_raum in st.session_state.raeume else 0
        )
        st.session_state.aktiver_raum = aktiver_raum
    else:
        st.warning("Noch keine Räume vorhanden.")

    with st.expander("➕ Neuen Raum hinzufügen"):
        neuer_raum = st.text_input("Raumname", key="neuer_raum_input")
        if st.button("Raum hinzufügen"):
            if neuer_raum.strip() == "":
                st.error("Bitte einen Raumnamen eingeben!")
            elif neuer_raum.strip() in st.session_state.raeume:
                st.warning("Dieser Raum existiert bereits!")
            else:
                st.session_state.raeume.append(neuer_raum.strip())
                save_json(ROOMS_FILE, st.session_state.raeume)
                st.session_state.aktiver_raum = neuer_raum.strip()
                st.success(f"Raum '{neuer_raum}' hinzugefügt!")
                st.rerun()

    with st.expander("🗑️ Raum löschen"):
        if st.session_state.raeume:
            raum_loeschen = st.selectbox("Raum auswählen", st.session_state.raeume, key="raum_loeschen")
            if st.button("Raum löschen", key="btn_raum_loeschen"):
                artikel_im_raum = [a for a in st.session_state.inventar if a["raum"] == raum_loeschen]
                if artikel_im_raum:
                    st.error(f"Raum '{raum_loeschen}' enthält noch {len(artikel_im_raum)} Artikel.")
                else:
                    st.session_state.raeume.remove(raum_loeschen)
                    save_json(ROOMS_FILE, st.session_state.raeume)
                    st.session_state.aktiver_raum = st.session_state.raeume[0] if st.session_state.raeume else None
                    st.success(f"Raum '{raum_loeschen}' gelöscht!")
                    st.rerun()

    st.divider()

    # --- Kategorieverwaltung ---
    st.header("🗂️ Kategorien")
    with st.expander("➕ Neue Kategorie"):
        neue_kat = st.text_input("Kategoriename", key="neue_kat_input")
        if st.button("Kategorie hinzufügen"):
            if neue_kat.strip() == "":
                st.error("Bitte einen Namen eingeben!")
            elif neue_kat.strip() in st.session_state.kategorien:
                st.warning("Kategorie existiert bereits!")
            else:
                st.session_state.kategorien.append(neue_kat.strip())
                save_json(CATEGORIES_FILE, st.session_state.kategorien)
                st.success(f"Kategorie '{neue_kat}' hinzugefügt!")
                st.rerun()

    with st.expander("🗑️ Kategorie löschen"):
        if st.session_state.kategorien:
            kat_loeschen = st.selectbox("Kategorie", st.session_state.kategorien, key="kat_loeschen")
            if st.button("Kategorie löschen", key="btn_kat_loeschen"):
                artikel_mit_kat = [a for a in st.session_state.inventar if a["kategorie"] == kat_loeschen]
                if artikel_mit_kat:
                    st.error(f"Kategorie wird noch von {len(artikel_mit_kat)} Artikel(n) verwendet.")
                else:
                    st.session_state.kategorien.remove(kat_loeschen)
                    save_json(CATEGORIES_FILE, st.session_state.kategorien)
                    st.success(f"Kategorie '{kat_loeschen}' gelöscht!")
                    st.rerun()

    st.divider()

    # --- Artikel hinzufügen ---
    st.header("➕ Artikel hinzufügen")
    st.markdown(f"📍 Raum: **{st.session_state.aktiver_raum}**")
    name = st.text_input("Artikelname")
    kategorie = st.selectbox("Kategorie", st.session_state.kategorien)
    menge = st.number_input("Menge", min_value=1, step=1)
    preis = st.number_input("Preis (€)", min_value=0.0, step=0.01, format="%.2f")
    barcode_nr = st.text_input("Barcode-Nr. (optional)", placeholder="z.B. 4012345678901")
    notiz = st.text_area("Notiz (optional)")

    if st.button("✅ Artikel hinzufügen"):
        if name.strip() == "":
            st.error("Bitte einen Artikelnamen eingeben!")
        elif not st.session_state.aktiver_raum:
            st.error("Bitte zuerst einen Raum anlegen!")
        else:
            neuer_artikel = {
                "id": len(st.session_state.inventar) + 1,
                "name": name.strip(),
                "kategorie": kategorie,
                "menge": int(menge),
                "verfuegbar": int(menge),
                "raum": st.session_state.aktiver_raum,
                "preis": round(preis, 2),
                "barcode": barcode_nr.strip(),
                "notiz": notiz.strip(),
                "datum": datetime.now().strftime("%d.%m.%Y %H:%M")
            }
            st.session_state.inventar.append(neuer_artikel)
            save_json(DATA_FILE, st.session_state.inventar)
            st.success(f"Artikel '{name}' in '{st.session_state.aktiver_raum}' hinzugefügt!")

# =============================================
# HAUPTBEREICH: TABS
# =============================================
# Detail-Ansicht Modal
if st.session_state.detail_artikel_id is not None:
    artikel_detail = next((a for a in st.session_state.inventar if a["id"] == st.session_state.detail_artikel_id), None)
    if artikel_detail:
        with st.container(border=True):
            st.markdown(f"## 📄 Artikeldetails – {artikel_detail['name']}")
            col_img, col_info = st.columns([1, 2])

            with col_img:
                img_path = get_image_path(artikel_detail["id"])
                if img_path:
                    st.image(img_path, use_container_width=True)
                else:
                    st.markdown("🖼️ *Kein Bild vorhanden*")

                # Bild hochladen
                uploaded = st.file_uploader("📸 Bild hochladen / ändern", type=["jpg", "jpeg", "png", "webp"],
                                            key=f"upload_{artikel_detail['id']}")
                if uploaded:
                    save_image(artikel_detail["id"], uploaded)
                    st.success("Bild gespeichert!")
                    st.rerun()

            with col_info:
                verfuegbar = artikel_detail.get("verfuegbar", artikel_detail["menge"])
                status = "✅ Verfügbar" if verfuegbar == artikel_detail["menge"] else ("🔴 Ausgeliehen" if verfuegbar == 0 else f"⚠️ Teils verfügbar ({verfuegbar}/{artikel_detail['menge']})")

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

                # Notiz bearbeiten
                st.markdown("**📝 Notiz:**")
                neue_notiz = st.text_area("Notiz bearbeiten", value=artikel_detail.get("notiz", ""),
                                          key=f"notiz_{artikel_detail['id']}")
                if st.button("💾 Notiz speichern", key=f"save_notiz_{artikel_detail['id']}"):
                    for a in st.session_state.inventar:
                        if a["id"] == artikel_detail["id"]:
                            a["notiz"] = neue_notiz.strip()
                            break
                    save_json(DATA_FILE, st.session_state.inventar)
                    st.success("Notiz gespeichert!")
                    st.rerun()

            # Ausleihhistorie für diesen Artikel
            artikel_ausleihen = [a for a in st.session_state.ausleihen if a["artikel_id"] == artikel_detail["id"]]
            if artikel_ausleihen:
                st.markdown("**📜 Ausleihhistorie dieses Artikels:**")
                df_al = pd.DataFrame(artikel_ausleihen).rename(columns={
                    "person": "Person", "checkout_datum": "Ausgabe",
                    "rueckgabe_erwartet": "Erw. Rückgabe", "rueckgabe_datum": "Rückgabe", "status": "Status"
                })
                st.dataframe(df_al[["Person", "Ausgabe", "Erw. Rückgabe", "Rückgabe", "Status"]],
                             use_container_width=True, hide_index=True)

            # QR-Code für diesen Artikel
            with st.expander("🖨️ QR-Code anzeigen"):
                qr_buf = generate_qr_code(artikel_detail)
                st.image(qr_buf, width=180)
                st.download_button("📥 QR-Code herunterladen", qr_buf,
                                   f"qr_{artikel_detail['name']}.png", "image/png")

            if st.button("✖️ Detailansicht schließen", type="secondary"):
                st.session_state.detail_artikel_id = None
                st.rerun()

    st.divider()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Inventar",
    "🔄 Checkout / Check-in",
    "📷 Barcode / QR-Code",
    "📊 Statistiken",
    "📤 Export"
])

# -----------------------------------------------
# TAB 1: Inventar
# -----------------------------------------------
with tab1:
    st.subheader(f"📋 Inventar – {st.session_state.aktiver_raum}")

    col1, col2, col3 = st.columns(3)
    with col1:
        suche = st.text_input("🔍 Artikel suchen", placeholder="z.B. Laptop...")
    with col2:
        filter_kategorie = st.selectbox("Kategorie filtern", ["Alle"] + st.session_state.kategorien)
    with col3:
        filter_status = st.selectbox("Status filtern", ["Alle", "✅ Verfügbar", "🔴 Ausgeliehen"])

    raum_inventar = [a for a in st.session_state.inventar if a["raum"] == st.session_state.aktiver_raum]

    if suche:
        raum_inventar = [a for a in raum_inventar if suche.lower() in a["name"].lower()]
    if filter_kategorie != "Alle":
        raum_inventar = [a for a in raum_inventar if a["kategorie"] == filter_kategorie]
    if filter_status == "✅ Verfügbar":
        raum_inventar = [a for a in raum_inventar if a.get("verfuegbar", a["menge"]) > 0]
    elif filter_status == "🔴 Ausgeliehen":
        raum_inventar = [a for a in raum_inventar if a.get("verfuegbar", a["menge"]) < a["menge"]]

    if raum_inventar:
        st.markdown("### Artikel")
        # Karten-Ansicht
        cols_per_row = 3
        for i in range(0, len(raum_inventar), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, a in enumerate(raum_inventar[i:i+cols_per_row]):
                with cols[j]:
                    with st.container(border=True):
                        # Bild
                        img_path = get_image_path(a["id"])
                        if img_path:
                            st.image(img_path, use_container_width=True)
                        else:
                            st.markdown("🖼️ *Kein Bild*")

                        verfuegbar = a.get("verfuegbar", a["menge"])
                        status = "✅" if verfuegbar == a["menge"] else ("🔴" if verfuegbar == 0 else "⚠️")
                        st.markdown(f"**{a['name']}** {status}")
                        st.caption(f"📂 {a['kategorie']} | 🏠 {a['raum']}")
                        st.caption(f"Menge: {a['menge']} | Verfügbar: {verfuegbar} | {a['preis']:.2f} €")

                        if st.button("🔍 Details", key=f"detail_{a['id']}"):
                            st.session_state.detail_artikel_id = a["id"]
                            st.rerun()
    else:
        st.info(f"Keine Artikel in Raum '{st.session_state.aktiver_raum}' gefunden.")

    st.divider()
    st.subheader("🗑️ Artikel löschen")
    alle_artikel_raum = [a for a in st.session_state.inventar if a["raum"] == st.session_state.aktiver_raum]
    if alle_artikel_raum:
        artikel_namen = [f"ID {a['id']} - {a['name']}" for a in alle_artikel_raum]
        ausgewaehlter_artikel = st.selectbox("Artikel auswählen", artikel_namen)
        if st.button("❌ Artikel löschen"):
            artikel_id = int(ausgewaehlter_artikel.split(" ")[1])
            st.session_state.inventar = [a for a in st.session_state.inventar if a["id"] != artikel_id]
            save_json(DATA_FILE, st.session_state.inventar)
            st.success("Artikel wurde gelöscht!")
            st.rerun()
    else:
        st.info("Keine Artikel zum Löschen vorhanden.")

# -----------------------------------------------
# TAB 2: Checkout / Check-in
# -----------------------------------------------
with tab2:
    st.subheader("🔄 Checkout & Check-in System")

    co_tab1, co_tab2, co_tab3 = st.tabs(["📤 Checkout (Ausleihe)", "📥 Check-in (Rückgabe)", "📜 Ausleihhistorie"])

    # --- CHECKOUT ---
    with co_tab1:
        st.markdown("### 📤 Artikel ausleihen")
        verfuegbare_artikel = [a for a in st.session_state.inventar if a.get("verfuegbar", a["menge"]) > 0]

        if verfuegbare_artikel:
            col1, col2 = st.columns(2)
            with col1:
                artikel_auswahl = st.selectbox(
                    "Artikel auswählen",
                    [f"ID {a['id']} – {a['name']} ({a['raum']}) | Verfügbar: {a.get('verfuegbar', a['menge'])}"
                     for a in verfuegbare_artikel],
                    key="checkout_artikel"
                )
                ausgeliehen_id = int(artikel_auswahl.split(" ")[1])
                artikel_obj = next(a for a in st.session_state.inventar if a["id"] == ausgeliehen_id)
                max_menge = artikel_obj.get("verfuegbar", artikel_obj["menge"])
                checkout_menge = st.number_input("Menge", min_value=1, max_value=max_menge, step=1, key="checkout_menge")

            with col2:
                person_name = st.text_input("Name der Person", placeholder="z.B. Max Mustermann")
                checkout_datum = st.date_input("Ausleihdatum", value=datetime.today())
                rueckgabe_erwartet = st.date_input("Erwartete Rückgabe", value=datetime.today())
                checkout_notiz = st.text_area("Notiz (optional)", key="checkout_notiz")

            if st.button("📤 Auschecken", type="primary"):
                if not person_name.strip():
                    st.error("Bitte den Namen der Person eingeben!")
                else:
                    # Ausleihe speichern
                    ausleihe = {
                        "ausleihe_id": len(st.session_state.ausleihen) + 1,
                        "artikel_id": ausgeliehen_id,
                        "artikel_name": artikel_obj["name"],
                        "raum": artikel_obj["raum"],
                        "menge": int(checkout_menge),
                        "person": person_name.strip(),
                        "checkout_datum": checkout_datum.strftime("%d.%m.%Y"),
                        "rueckgabe_erwartet": rueckgabe_erwartet.strftime("%d.%m.%Y"),
                        "rueckgabe_datum": None,
                        "status": "ausgeliehen",
                        "notiz": checkout_notiz.strip()
                    }
                    st.session_state.ausleihen.append(ausleihe)
                    save_json(LENDING_FILE, st.session_state.ausleihen)

                    # Verfügbare Menge anpassen
                    for a in st.session_state.inventar:
                        if a["id"] == ausgeliehen_id:
                            a["verfuegbar"] = a.get("verfuegbar", a["menge"]) - int(checkout_menge)
                            break
                    save_json(DATA_FILE, st.session_state.inventar)
                    st.success(f"✅ '{artikel_obj['name']}' ({checkout_menge}x) wurde an **{person_name}** ausgeliehen.")
                    st.rerun()
        else:
            st.info("Keine verfügbaren Artikel zum Ausleihen vorhanden.")

        # Aktuelle Ausleihen anzeigen
        st.divider()
        st.markdown("### 📋 Aktuell ausgeliehene Artikel")
        aktive_ausleihen = [a for a in st.session_state.ausleihen if a["status"] == "ausgeliehen"]
        if aktive_ausleihen:
            df_aktiv = pd.DataFrame(aktive_ausleihen)
            df_aktiv = df_aktiv.rename(columns={
                "ausleihe_id": "ID", "artikel_name": "Artikel", "raum": "Raum",
                "menge": "Menge", "person": "Ausgeliehen an",
                "checkout_datum": "Ausgabe", "rueckgabe_erwartet": "Rückgabe erwartet", "notiz": "Notiz"
            })
            st.dataframe(df_aktiv[["ID", "Artikel", "Raum", "Menge", "Ausgeliehen an", "Ausgabe", "Rückgabe erwartet", "Notiz"]],
                         use_container_width=True, hide_index=True)
        else:
            st.info("Keine aktiven Ausleihen.")

    # --- CHECK-IN ---
    with co_tab2:
        st.markdown("### 📥 Artikel zurückgeben")
        aktive_ausleihen = [a for a in st.session_state.ausleihen if a["status"] == "ausgeliehen"]

        if aktive_ausleihen:
            ausleihe_auswahl = st.selectbox(
                "Ausleihe auswählen",
                [f"ID {a['ausleihe_id']} – {a['artikel_name']} ({a['menge']}x) → {a['person']} | Ausgabe: {a['checkout_datum']}"
                 for a in aktive_ausleihen],
                key="checkin_auswahl"
            )
            ausleihe_id = int(ausleihe_auswahl.split(" ")[1])
            ausleihe_obj = next(a for a in st.session_state.ausleihen if a["ausleihe_id"] == ausleihe_id)

            col1, col2 = st.columns(2)
            with col1:
                st.info(f"**Artikel:** {ausleihe_obj['artikel_name']}  \n"
                        f"**Ausgeliehen an:** {ausleihe_obj['person']}  \n"
                        f"**Ausgabe:** {ausleihe_obj['checkout_datum']}  \n"
                        f"**Erwartete Rückgabe:** {ausleihe_obj['rueckgabe_erwartet']}")
            with col2:
                rueckgabe_datum = st.date_input("Rückgabedatum", value=datetime.today(), key="checkin_datum")
                checkin_notiz = st.text_area("Notiz (optional)", key="checkin_notiz")

            if st.button("📥 Einchecken", type="primary"):
                for a in st.session_state.ausleihen:
                    if a["ausleihe_id"] == ausleihe_id:
                        a["status"] = "zurückgegeben"
                        a["rueckgabe_datum"] = rueckgabe_datum.strftime("%d.%m.%Y")
                        if checkin_notiz.strip():
                            a["notiz"] += f" | Rückgabe: {checkin_notiz.strip()}"
                        break
                save_json(LENDING_FILE, st.session_state.ausleihen)

                for a in st.session_state.inventar:
                    if a["id"] == ausleihe_obj["artikel_id"]:
                        a["verfuegbar"] = a.get("verfuegbar", 0) + ausleihe_obj["menge"]
                        break
                save_json(DATA_FILE, st.session_state.inventar)
                st.success(f"✅ '{ausleihe_obj['artikel_name']}' wurde erfolgreich zurückgebucht!")
                st.rerun()
        else:
            st.info("Keine aktiven Ausleihen zur Rückgabe vorhanden.")

    # --- HISTORIE ---
    with co_tab3:
        st.markdown("### 📜 Vollständige Ausleihhistorie")

        col1, col2 = st.columns(2)
        with col1:
            filter_person = st.text_input("🔍 Nach Person suchen", placeholder="z.B. Max Mustermann")
        with col2:
            filter_historie_status = st.selectbox("Status", ["Alle", "ausgeliehen", "zurückgegeben"])

        historie = st.session_state.ausleihen
        if filter_person:
            historie = [a for a in historie if filter_person.lower() in a["person"].lower()]
        if filter_historie_status != "Alle":
            historie = [a for a in historie if a["status"] == filter_historie_status]

        if historie:
            df_hist = pd.DataFrame(historie)
            df_hist = df_hist.rename(columns={
                "ausleihe_id": "ID", "artikel_name": "Artikel", "raum": "Raum",
                "menge": "Menge", "person": "Person", "checkout_datum": "Ausgabe",
                "rueckgabe_erwartet": "Erw. Rückgabe", "rueckgabe_datum": "Rückgabe",
                "status": "Status", "notiz": "Notiz"
            })
            st.dataframe(df_hist, use_container_width=True, hide_index=True)

            csv_hist = df_hist.to_csv(index=False, sep=";", encoding="utf-8-sig")
            st.download_button("📥 Historie als CSV", csv_hist, "ausleihhistorie.csv", "text/csv")
        else:
            st.info("Keine Einträge gefunden.")

# -----------------------------------------------
# TAB 3: Barcode / QR-Code
# -----------------------------------------------
with tab3:
    st.subheader("📷 Barcode & QR-Code")
    bc_tab1, bc_tab2 = st.tabs(["🔍 Barcode scannen", "🖨️ QR-Code generieren"])

    with bc_tab1:
        st.markdown("Scanne einen Barcode mit der Kamera, um einen Artikel zu suchen.")
        try:
            scan_result = quagga()
            if scan_result and scan_result.get("codeResult"):
                scanned_code = scan_result["codeResult"]["code"]
                st.success(f"✅ Erkannter Barcode: `{scanned_code}`")
                gefundene_artikel = [a for a in st.session_state.inventar if a.get("barcode") == scanned_code]
                if gefundene_artikel:
                    st.dataframe(pd.DataFrame(gefundene_artikel), use_container_width=True, hide_index=True)
                else:
                    st.warning("Kein Artikel mit diesem Barcode gefunden.")
        except Exception:
            st.info("💡 Barcode-Scanner nicht verfügbar. Bitte installiere: `pip install streamlit-quagga`")
            manueller_code = st.text_input("Barcode manuell eingeben", placeholder="z.B. 4012345678901")
            if manueller_code:
                gefundene_artikel = [a for a in st.session_state.inventar if a.get("barcode") == manueller_code]
                if gefundene_artikel:
                    st.success("✅ Artikel gefunden:")
                    st.dataframe(pd.DataFrame(gefundene_artikel), use_container_width=True, hide_index=True)
                else:
                    st.warning("Kein Artikel gefunden.")

    with bc_tab2:
        st.markdown("Generiere einen QR-Code für jeden Artikel zum Ausdrucken.")
        if st.session_state.inventar:
            qr_auswahl = st.selectbox(
                "Artikel für QR-Code auswählen",
                [f"ID {a['id']} – {a['name']} ({a['raum']})" for a in st.session_state.inventar]
            )
            ausgewaehlte_id = int(qr_auswahl.split(" ")[1])
            artikel_fuer_qr = next((a for a in st.session_state.inventar if a["id"] == ausgewaehlte_id), None)
            if artikel_fuer_qr:
                qr_buf = generate_qr_code(artikel_fuer_qr)
                st.image(qr_buf, caption=f"QR-Code: {artikel_fuer_qr['name']}", width=200)
                st.download_button("📥 QR-Code herunterladen", qr_buf,
                                   f"qr_{artikel_fuer_qr['name']}.png", "image/png")
        else:
            st.info("Noch keine Artikel vorhanden.")

# -----------------------------------------------
# TAB 4: Statistiken
# -----------------------------------------------
with tab4:
    st.subheader("📊 Statistiken")

    gesamt_col, raum_col = st.columns(2)
    with gesamt_col:
        st.markdown("### 🌐 Gesamtübersicht")
        if st.session_state.inventar:
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("📦 Artikel", len(st.session_state.inventar))
            col2.metric("🔢 Gesamtmenge", sum(a["menge"] for a in st.session_state.inventar))
            col3.metric("💶 Gesamtwert", f"{sum(a['preis'] * a['menge'] for a in st.session_state.inventar):.2f} €")
            aktive = len([a for a in st.session_state.ausleihen if a["status"] == "ausgeliehen"])
            col4.metric("🔴 Aktive Ausleihen", aktive)

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

    with raum_col:
        st.markdown(f"### 🏠 Raum: {st.session_state.aktiver_raum}")
        raum_artikel = [a for a in st.session_state.inventar if a["raum"] == st.session_state.aktiver_raum]
        if raum_artikel:
            col1, col2, col3 = st.columns(3)
            col1.metric("📦 Artikel", len(raum_artikel))
            col2.metric("🔢 Menge", sum(a["menge"] for a in raum_artikel))
            col3.metric("💶 Wert", f"{sum(a['preis'] * a['menge'] for a in raum_artikel):.2f} €")
        else:
            st.info(f"Keine Artikel in '{st.session_state.aktiver_raum}'.")

# -----------------------------------------------
# TAB 5: Export
# -----------------------------------------------
with tab5:
    st.subheader("📤 Export")
    export_option = st.radio("Was möchtest du exportieren?", ["Aktiver Raum", "Alle Räume", "Ausleihhistorie"])

    if export_option == "Aktiver Raum":
        export_data = [a for a in st.session_state.inventar if a["raum"] == st.session_state.aktiver_raum]
        filename = f"inventar_{st.session_state.aktiver_raum}.csv"
    elif export_option == "Alle Räume":
        export_data = st.session_state.inventar
        filename = "inventar_gesamt.csv"
    else:
        export_data = st.session_state.ausleihen
        filename = "ausleihhistorie.csv"

    if export_data:
        df_export = pd.DataFrame(export_data)
        csv = df_export.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            label=f"📥 CSV herunterladen ({len(export_data)} Einträge)",
            data=csv, file_name=filename, mime="text/csv"
        )
    else:
        st.info("Keine Daten zum Exportieren.")

# --- Footer ---
st.markdown("---")
st.markdown("🤖 **EVA** - Inventarisierungs-App | Erstellt mit Python & Streamlit")
