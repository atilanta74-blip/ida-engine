import io
import json
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import google.generativeai as genai
import streamlit as st

# --- OLDALBEALLITASOK ---
st.set_page_config(
    page_title="BSHM / B-ME // IDA",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- CSICSA & HIGH-TECH STILUS (ELES, JOL LATHATO FEHER CIMKEKKEL ES CHECKBOXOKKAL) ---
st.markdown(
    """
<style>
    .stApp {
        background: linear-gradient(135deg, #0b0f19 0%, #111827 50%, #070a12 100%);
        color: #e2e8f0;
        font-family: 'Segoe UI', Roboto, sans-serif;
    }
    .header-box {
        background: rgba(17, 24, 39, 0.85);
        border: 1px solid #1e293b;
        border-left: 5px solid #0284c7;
        padding: 20px 25px;
        border-radius: 12px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5), 0 0 15px rgba(2, 132, 199, 0.2);
        margin-bottom: 25px;
    }
    .header-title {
        font-size: 26px;
        font-weight: 800;
        letter-spacing: 1.5px;
        color: #f8fafc;
        margin: 0;
        text-transform: uppercase;
    }
    .header-sub {
        font-size: 13px;
        color: #38bdf8;
        letter-spacing: 3px;
        font-weight: 600;
        margin-top: 4px;
        text-transform: uppercase;
    }
    .kpi-card {
        background: rgba(30, 41, 59, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.08);
        padding: 14px 18px;
        border-radius: 10px;
        backdrop-filter: blur(10px);
    }
    .kpi-title { font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: #94a3b8; }
    .kpi-value { font-size: 17px; font-weight: 700; color: #38bdf8; margin-top: 3px; }

    /* ABLAKOK FELETTI CIMKEK: ELES HOFEHER, BOLD, 15PX */
    div[data-testid="stWidgetLabel"] p, 
    div[data-testid="stWidgetLabel"] label, 
    label p, 
    label {
        font-size: 15px !important;
        font-weight: 700 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        letter-spacing: 0.5px !important;
        opacity: 1 !important;
        margin-bottom: 6px !important;
    }

    /* CHECKBOXOK STILUSA: ELES FEHER SZÖVEG, NAGYOBB BETUK */
    div[data-testid="stCheckbox"] label p,
    div[data-testid="stCheckbox"] label span {
        font-size: 15px !important;
        font-weight: 700 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        letter-spacing: 0.5px !important;
    }

    /* BEMENETI MEZOK BELSO TARTALMA */
    .stTextArea textarea, .stTextInput input {
        background-color: #0f172a !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-size: 14px !important;
        line-height: 1.5 !important;
        border: 1px solid #0284c7 !important;
        border-radius: 8px !important;
    }
    .stTextArea textarea:focus, .stTextInput input:focus {
        border-color: #38bdf8 !important;
        box-shadow: 0 0 12px rgba(56, 189, 248, 0.4) !important;
    }

    /* GOMBOK ALTALANOS FORMAZASA ES FEHER SZOVEG BIZTOSITASA */
    .stButton button, .stDownloadButton button {
        width: 100% !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 20px !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        letter-spacing: 0.5px !important;
        text-transform: uppercase !important;
        transition: all 0.3s ease !important;
    }
    .stButton button *, .stDownloadButton button * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-weight: 700 !important;
    }

    /* 1. SZAKI GOMB: ELENK NARANCSSARGA HATTER */
    .st-key-btn_szaki button {
        background: linear-gradient(90deg, #ea580c 0%, #c2410c 100%) !important;
        border: 1px solid #f97316 !important;
        box-shadow: 0 4px 15px rgba(234, 88, 12, 0.5) !important;
    }
    .st-key-btn_szaki button:hover {
        background: linear-gradient(90deg, #f97316 0%, #ea580c 100%) !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(234, 88, 12, 0.7) !important;
    }

    /* 2. MERNOKI GOMB: KOBALTKEK HATTER */
    .st-key-btn_std button {
        background: linear-gradient(90deg, #0284c7 0%, #2563eb 100%) !important;
        border: 1px solid #38bdf8 !important;
        box-shadow: 0 4px 15px rgba(2, 132, 199, 0.5) !important;
    }
    .st-key-btn_std button:hover {
        background: linear-gradient(90deg, #0ea5e9 0%, #1d4ed8 100%) !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(2, 132, 199, 0.7) !important;
    }

    /* LETOLTES GOMB */
    .stDownloadButton button {
        background: linear-gradient(90deg, #059669 0%, #10b981 100%) !important;
        border: 1px solid #34d399 !important;
        box-shadow: 0 4px 15px rgba(16, 185, 129, 0.5) !important;
    }

    .stCaption p, .stCaption {
        color: #94a3b8 !important;
        -webkit-text-fill-color: #94a3b8 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# --- OLDALSAV: API KULCS BEALLITASA ---
with st.sidebar:
  st.markdown("### ⚙️ AI Rendszer Beallitas")
  api_key_input = st.text_input(
      "Google Gemini API Kulcs:",
      type="password",
      value=st.secrets.get("GEMINI_API_KEY", ""),
      help="Illeszd be az AI Studio kulcsodat",
  )
  st.markdown("---")
  st.markdown(
      "💡 **Tipp:** Ha a Streamlit Secrets-be beirtad a `GEMINI_API_KEY`-t,"
      " uresen maradhat."
  )

# --- FEJLEC ES KPI STATUS BAR ---
st.markdown(
    """
<div class="header-box">
    <div class="header-title">⚡ BSHM / B-ME // IDA</div>
    <div class="header-sub">POWERED BY NAGY ATTILA</div>
</div>
""",
    unsafe_allow_html=True,
)

col_k1, col_k2, col_k3, col_k4 = st.columns(4)
with col_k1:
  st.markdown(
      '<div class="kpi-card"><div class="kpi-title">Rendszer'
      ' Statusz</div><div class="kpi-value" style="color: #4ade80;">● ONLINE /'
      " AI READY</div></div>",
      unsafe_allow_html=True,
  )
with col_k2:
  st.markdown(
      '<div class="kpi-card"><div class="kpi-title">Sablon Revizio</div><div'
      ' class="kpi-value">Rev. 3 (2019.04.03)</div></div>',
      unsafe_allow_html=True,
  )
with col_k3:
  st.markdown(
      '<div class="kpi-card"><div class="kpi-title">Regisztracio</div><div'
      ' class="kpi-value">BME001-F-BDE-3</div></div>',
      unsafe_allow_html=True,
  )
with col_k4:
  st.markdown(
      '<div class="kpi-card"><div class="kpi-title">Ellenorzesi'
      ' Felelos</div><div class="kpi-value">Janoczki M</div></div>',
      unsafe_allow_html=True,
  )

st.markdown("<br>", unsafe_allow_html=True)

# --- BEVITELI MEZOK ---
c_left, c_right = st.columns([1, 1.4])
with c_left:
  st.markdown("### 📋 Alapadatok")
  gep_nev = st.text_input("Berendezes / Gep megnevezese", value="12k pit")
  hiba_rovid = st.text_input("Hiba megnevezese", value="12k pit hiba")
  datum_val = st.text_input("Datum", value="2026.10.14.")
  felelos_val = st.text_input("Karbantartasi Felelos", value="Nagy Attila")

with c_right:
  st.markdown("### 📝 Esemeny & Celzott Fejlesztesi Irany")
  jegyzet_szoveg = st.text_area(
      "1. Mi volt a hiba? (Muszakos jegyzet / tenyleges hiba):",
      height=110,
      value=(
          "(PIT1 becsatlakozo cv-n heveder csere megtortent es a cv kozelebb"
          " lett allitva a lehordo cv-hez.)"
      ),
  )
  fejlesztesi_cel = st.text_area(
      (
          "2. Mit hozzunk ki fejlesztesnek? (Elvart vegleges megoldas /"
          " intezkedes):"
      ),
      height=90,
      value=(
          "a regi Segment gateway controllok lecserelese multicontroll-ra a"
          " gyorsabb hibakereses miatt es hogy van hozza meg csere alkatresz a"
          " regiekkel ellentetben"
      ),
  )

# --- KRITERIUMOK LISTAJA (A GYARI TABLAZAT ALAPJAN) ---
kriterium_lista = [
    (
        "Tisztaság",
        "(A részegység (berendezés) elég tiszta, hogy ellenőrizhető legyen és az"
        " állapota ne romoljon rohamosan?)",
    ),
    ("Meghúzottság", "(A kötőelemek megfelelően meg vannak húzva?)"),
    (
        "Kenés",
        "(A részegység (berendezés) a kenési tervnek megfelelő típusú,"
        " mennyiségű kenőanyaggal, megfelelő gyakorisággal, és megfelelő"
        " módszerrel volt kenve?)",
    ),
    (
        "Karbantartás",
        "(A részegység (berendezés) szükséges elemei az előírásoknak"
        " megfelelően vannak karbantartva? A karbantartási gyakoriság"
        " megfelelő? Megfelelőek a karbantartási standardok és eljárások?)",
    ),
    (
        "Környezet",
        "(A részegység (berendezés) az előírásoknak megfelelő környezeti"
        " körülmények között üzemel?)",
    ),
    (
        "Működtetés",
        "(A megfelelő működtetést meghatározó munkafolyamatokat következetese"
        " betartják? Végez a részegység (berendezés) nem rendeltetésszerű"
        " műveleteket?)",
    ),
    (
        "Specifikáció",
        "(Működtetik a részegységet (berendezést) határértékek (sebesség,"
        " nyomás, terhelés stb.) fölött?)",
    ),
    (
        "Működés",
        "(Végez nem megfelelő, vagy fölösleges, műveletet a részegység"
        " (berendezés)?)",
    ),
    (
        "Telepítés",
        "(A részegység (berendezés) telepítése során szigorúan követték az"
        " összeszerelési útmutatót és standardot? Az elvárt telepítési"
        " körülményeket elérték?)",
    ),
    (
        "Összeszerelés",
        "(A részegység (berendezés) alkatrészei megfelelően illeszkednek"
        " egymáshoz?)",
    ),
    (
        "Gyártás",
        "(A részegység (berendezés) alkatrészei a rajzoknak megfelelően lettek"
        " legyártva és összeszerelve?)",
    ),
    (
        "Tervezés",
        "(Az egyes alkatrészek vagy eszközök megbízhatóan teszik a dolgukat?)",
    ),
    (
        "Javítás",
        "(Az üzemzavar elhárítása az elvárt idő alatt és minőségben"
        " megtörtént?)",
    ),
]

# --- BEPIPÁLHATÓ KRITÉRIUMOK KIVÁLASZTÁSA ---
st.markdown("### ☑️ Kifejtendő Kritériumok kiválasztása (az 1. oldalhoz)")
st.markdown(
    "<span style='color: #94a3b8; font-size: 13px;'>Pipáld be azokat a"
    " kritériumokat, amiket ki akarsz fejteni. <b>Ahova nem kerül pipa, az a"
    " sor teljesen üresen marad a kész Word dokumentumban!</b></span>",
    unsafe_allow_html=True,
)

cb_cols = st.columns(4)
col_items = [
    kriterium_lista[0:4],  # Tisztaság, Meghúzottság, Kenés, Karbantartás
    kriterium_lista[4:7],  # Környezet, Működtetés, Specifikáció
    kriterium_lista,  # Működés, Telepítés, Összeszerelés
    kriterium_lista,  # Gyártás, Tervezés, Javítás
]

kivalasztott_kriteriumok = []
for c_idx, items in enumerate(col_items):
  with cb_cols[c_idx]:
    for nev, kerdes in items:
      alap_ertek = True if nev in ["Karbantartás", "Javítás"] else False
      if st.checkbox(nev, value=alap_ertek, key=f"cb_{nev}"):
        kivalasztott_kriteriumok.append(nev)


# --- INTELLIGENS FLASH MOTOR ---
def elemez_geminivel(
    api_key,
    gep,
    hiba,
    datum,
    felelos,
    jegyzet,
    fejlesztes,
    kivalasztott,
    zero_bullshit=True,
):
  genai.configure(api_key=api_key)

  elerheto = []
  try:
    for m in genai.list_models():
      if "generateContent" in m.supported_generation_methods:
        elerheto.append(m.name)
  except Exception:
    pass

  flash_list = [
      m for m in elerheto if "flash" in m.lower() and "pro" not in m.lower()
  ]

  tartalek_flash = [
      "models/gemini-2.0-flash",
      "models/gemini-1.5-flash-latest",
      "models/gemini-1.5-flash-002",
      "models/gemini-1.5-flash",
      "models/gemini-1.5-flash-8b",
      "gemini-2.0-flash",
      "gemini-1.5-flash",
  ]

  kiprobalando = []
  for m in flash_list + tartalek_flash:
    if m not in kiprobalando:
      kiprobalando.append(m)

  if zero_bullshit:
    stilus_utasitas = """
        SZIGORU STILUSKOVETELMENY - ZERO-BULLSHIT / GYARI SZAKI MOD:
        - TILOS BARMIFELE 'AI-IZU' RIZSA, corporate bullshit, kormondatok vagy felesleges szovegeles!
        - TILOSAK az ilyen szavak: 'holisztikus', 'szinergia', 'proaktiv paradigma', 'optimalizacio', 'folyamatos fejlodesi kultura', 'human tenyezo'.
        - Ugy fogalmazz, mintha egy szukszavu, kokemeny gyari gepesztechnikus es karbantartas-vezeto irta volna a jegyzokonyvet:
          * Rovid, lenyegretoro, pontos gepeszeti tomondatok.
          * Pontos muszaki adatok a jegyzetbol (heveder kopas/csere, CV tavolsagallitas, elavult Segment gateway, alkatreszhiany, multicontroll modernizalas).
          * A 2. oldalon az 5 Miert kerdesek es valaszok legyenek egyenesek, lenyegretorok.
          * A vegeredmeny 100%-ban ugy nezzen ki, mintha kezzel gepelt gyari jegyzokonyv lenne, nulla AI jelleggel!
        """
  else:
    stilus_utasitas = (
        "Hivatalos mernoki stilusban, reszletesen fejtsd ki a pontokat."
    )

  kivalasztottak_szoveg = (
      ", ".join(kivalasztott)
      if kivalasztott
      else "Nincs bejelolve (minden sor maradjon ures)"
  )

  prompt = f"""
    Te egy tapasztalt mechatronikai karbantarto szakember vagy a Bridgestone gyarban.
    Elemezd az alabbi bemeneteket a gyari BSHM B-ME IDA sablon szerint:

    Berendezes: {gep}
    Hiba megnevezese: {hiba}
    Datum: {datum}
    Felelos: {felelos}

    1. TENYLEGESEN MEGTORTENT HIBA / ESEMENY:
    {jegyzet}

    2. CELZOTT FEJLESZTES / ELVART VEGLEGES MEGOLDAS:
    {fejlesztes}

    3. KIFEJTENDO KRITERIUMOK (AZ 1. OLDALHOZ):
    {kivalasztottak_szoveg}

    SZIGORU SZABALYOK:
    1. KIZAROLAG a fenti 3. pontban megadott, kivalasztott kriteriumokhoz ({kivalasztottak_szoveg}) keszits ertekelest! A nem kivalasztott kriteriumokhoz semmit ne irj.
    2. A logikai levezetést ugy epitsd fel, hogy logikusan ebbol a hibabol pontosan a megadott fejlesztesi cel ({fejlesztes}) jojjon ki vegeredmenykent az 5 Miertben es az akciotervben!
    3. FELELOS ES HATARIDO: Az Ellenintezkedes oszlopban az intezkedes leirasa utan kulon sorban mindig szerepeljen a 'Felelos: ..........' es 'Hatarido: ..........' szoveg, de ne toltsd ki nevvel/datummal!

    {stilus_utasitas}

    KIZAROLAG ervenyes JSON formatumban valaszolj, az alabbi sema szerint:
    {{
      "kriteriumok": {{
        "Kriterium Neve": {{
          "idealis": "1. ...",
          "aktualis": "1. ...",
          "hianyossag": "1. ...",
          "gyokerok": "1. ...",
          "ellenintezkedes": "1. ...\\nFelelos: ..........\\nHatarido: .........."
        }}
      }},
      "ot_miert_agak": [
        {{
          "ag_nev": "1. AG: MUSZAKI / FIZIKAI AG",
          "problema": "Problema leiras a mechanikai / fizikai szempontbol",
          "miert1": "1. Miert kerdes es valasz",
          "miert2": "2. Miert kerdes es valasz",
          "miert3": "3. Miert kerdes es valasz",
          "miert4": "4. Miert kerdes es valasz",
          "miert5_gyokerok": "GYOKEROK: A fizikai/mechanikai ok kozvetlen megnevezese"
        }},
        {{
          "ag_nev": "2. AG: KARBANTARTASI / VEZERLESI / DIAGNOSZTIKAI AG",
          "problema": "Problema leiras a hibakereses es vezerles szempontjabol",
          "miert1": "1. Miert kerdes es valasz",
          "miert2": "2. Miert kerdes es valasz",
          "miert3": "3. Miert kerdes es valasz",
          "miert4": "4. Miert kerdes es valasz",
          "miert5_gyokerok": "GYOKEROK: Vezerlesi vagy karbantartasi hianyossag"
        }},
        {{
          "ag_nev": "3. AG: ALKATRESZ-ELLATASI / FEJLESZTESI AG",
          "problema": "Problema leiras az alkatresz-elerhetoseg es modernizalas szempontjabol",
          "miert1": "1. Miert kerdes es valasz",
          "miert2": "2. Miert kerdes es valasz",
          "miert3": "3. Miert kerdes es valasz",
          "miert4": "4. Miert kerdes es valasz",
          "miert5_gyokerok": "GYOKEROK: Alkatresz-tamogatas megszunese, celzott fejlesztes szuksegessege"
        }}
      ],
      "vegleges_akcioterv": "Rovid, katonas osszefoglalo a hiba javitasarol es a celzott fejlesztesrol, a vegen:\\nFelelos: ..........\\nHatarido: .........."
    }}
    """

  utolso_hiba = None
  for mod_nev in kiprobalando:
    try:
      model = genai.GenerativeModel(mod_nev)
      try:
        res = model.generate_content(
            prompt, generation_config={"response_mime_type": "application/json"}
        )
        raw_text = res.text.strip()
      except Exception:
        res = model.generate_content(prompt)
        raw_text = res.text.strip()

      if raw_text.startswith("```"):
        lines = raw_text.splitlines()
        if lines[0].startswith("```"):
          lines = lines[1:]
        if lines and lines[-1].startswith("```"):
          lines = lines[:-1]
        raw_text = "\n".join(lines).strip()
        if raw_text.startswith("json"):
          raw_text = raw_text[4:].strip()

      return json.loads(raw_text)
    except Exception as err:
      utolso_hiba = err
      continue

  raise utolso_hiba


# --- DOKUMENTUM GENERALO MOTOR (100% GYARI FORMATUM, CSAK JANOCZKI M!) ---
def general_hivatalos_docx(
    gep, hiba, datum, felelos, ai_data, jegyzet, fejlesztes, kivalasztott
):
  doc = docx.Document()
  for section in doc.sections:
    section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

  def set_cell_bg(cell, fill_hex):
    cell._tc.get_or_add_tcPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    )

  def set_cell_pad(cell, top=35, bottom=35, left=50, right=50):
    cell._tc.get_or_add_tcPr().append(
        parse_xml(
            f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}"'
            f' w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left'
            f' w:w="{left}" w:type="dxa"/><w:right w:w="{right}"'
            ' w:type="dxa"/></w:tcMar>'
        )
    )

  def set_cell_border(cell, sz="4"):
    b_xml = (
        f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="{sz}"'
        ' w:space="0" w:color="000000"/><w:left w:val="single"'
        f' w:sz="{sz}" w:space="0" w:color="000000"/><w:bottom w:val="single"'
        f' w:sz="{sz}" w:space="0" w:color="000000"/><w:right w:val="single"'
        f' w:sz="{sz}" w:space="0" w:color="000000"/></w:tcBorders>'
    )
    cell._tc.get_or_add_tcPr().append(parse_xml(b_xml))

  p_title = doc.add_paragraph()
  r = p_title.add_run(
      "Üzemzavar elhárítási napi irányítási rendszer – Berendezés hiányosság"
      " értékelés, mélyanalízis, ellenintézkedési lap"
  )
  r.font.name = "Calibri"
  r.font.size = Pt(11)
  r.font.bold = True

  p_info = doc.add_paragraph()
  p_info.add_run("Üzemzavar leírása: ").bold = True
  p_info.add_run(f"{gep} – {hiba}          ").underline = True
  p_info.add_run("Dátum: ").bold = True
  p_info.add_run(f"{datum}          ").underline = True
  p_info.add_run("Karb. felelős: ").bold = True
  p_info.add_run(f"{felelos}").underline = True

  p1_table = doc.add_table(rows=15, cols=6)
  p1_table.alignment = WD_TABLE_ALIGNMENT.CENTER
  p1_widths = [
      Inches(1.8),
      Inches(1.75),
      Inches(1.75),
      Inches(1.75),
      Inches(1.85),
      Inches(1.8),
  ]

  c_plan = p1_table.cell(0, 0)
  c_plan.merge(p1_table.cell(0, 1)).merge(p1_table.cell(0, 2)).merge(
      p1_table.cell(0, 3)
  )
  c_plan.text = "TERVEZD MEG (PLAN)"
  p1_table.cell(0, 4).text = "VALÓSÍTSD MEG (DO)"
  p1_table.cell(0, 5).text = "ELLENŐRIZD (CHECK), STANDARDIZÁLD (ACT)"

  sub_heads = [
      "KRITÉRIUM",
      "IDEÁLIS ÁLLAPOT",
      "AKTUÁLIS ÁLLAPOT",
      "HIÁNYOSSÁG",
      "Hiányosság gyökéroka\n(ismert, vagy “5 miért” elemzés alapján)",
      "Ellenintézkedés\n(határidő, felelős, standardizálás)",
  ]
  for c_i, h in enumerate(sub_heads):
    p1_table.cell(1, c_i).text = h

  for r_i in range(2):
    for cell in p1_table.rows[r_i].cells:
      set_cell_bg(cell, "F2F2F2")
      set_cell_pad(cell, 50, 50, 50, 50)
      set_cell_border(cell, sz="6")
      p = cell.paragraphs[0]
      p.alignment = WD_ALIGN_PARAGRAPH.CENTER
      if p.runs:
        p.runs[0].font.name = "Calibri"
        p.runs[0].font.size = Pt(7.5)
        p.runs[0].font.bold = True

  # Kritériumok feldolgozása: szótár vagy lista kezelése
  raw_kriteriumok = ai_data.get("kriteriumok", {})
  crit_dict = {}
  if isinstance(raw_kriteriumok, dict):
    crit_dict = raw_kriteriumok
  elif isinstance(raw_kriteriumok, list):
    for item in raw_kriteriumok:
      if isinstance(item, dict) and "nev" in item:
        crit_dict[item["nev"]] = item

  # 13 sor pontos feltöltése a gyári sorrendben
  for r_idx, (krit_nev, krit_kerdes) in enumerate(kriterium_lista, start=2):
    is_selected = krit_nev in kivalasztott
    data = crit_dict.get(krit_nev, {})

    # Ha be van pipálva, kifejtjük; ha nincs, ÜRESEN HAGYJUK a cellákat!
    if is_selected:
      ell = data.get("ellenintezkedes", "")
      if ell and "Felelős:" not in ell:
        ell = f"{ell}\nFelelős: ..........\nHatáridő: .........."
      col_vals = [
          f"{krit_nev}\n{krit_kerdes}",
          data.get("idealis", "1. Standard előírások betartása."),
          data.get("aktualis", "1. Megfelelő üzemállapot."),
          data.get("hianyossag", "1. Eltérés tapasztalható."),
          data.get("gyokerok", "1. Kiváltó ok meghatározva."),
          (
              ell
              if ell
              else (
                  "1. Intézkedés rögzítve.\nFelelős: ..........\nHatáridő:"
                  " .........."
              )
          ),
      ]
    else:
      # TELJESEN ÜRES CELLÁK, AHOL NINCS PIPA!
      col_vals = [f"{krit_nev}\n{krit_kerdes}", "", "", "", "", ""]

    for c_i, val in enumerate(col_vals):
      cell = p1_table.cell(r_idx, c_i)
      cell.text = val
      set_cell_bg(cell, "FFFFFF")
      set_cell_pad(cell, 35, 35, 50, 50)
      set_cell_border(cell, sz="4")
      p = cell.paragraphs[0]
      p.paragraph_format.line_spacing = Pt(9)
      if p.runs:
        p.runs[0].font.name = "Calibri"
        p.runs[0].font.size = Pt(7)
        p.runs[0].font.color.rgb = RGBColor(0, 0, 0)
        if c_i == 0:
          p.runs[0].font.bold = True

  for row in p1_table.rows:
    for c_i, w in enumerate(p1_widths):
      row.cells[c_i].width = w

  doc.add_page_break()
  p2_title = doc.add_paragraph()
  r = p2_title.add_run(
      "Üzemzavar elhárítási napi irányítási rendszer – Berendezés hiányosság"
      " értékelés, mélyanalízis, ellenintézkedési lap"
  )
  r.font.name = "Calibri"
  r.font.size = Pt(11)
  r.font.bold = True

  p2_info = doc.add_paragraph()
  p2_info.add_run(
      f"Mélyanalízis (5 Miért elemzés)          Gép: {gep}          Dátum:"
      f" {datum}          Karb. felelős: {felelos}"
  ).bold = True

  p2_table = doc.add_table(rows=12, cols=5)
  p2_table.alignment = WD_TABLE_ALIGNMENT.CENTER
  col_w = Inches(10.7 / 5)

  agak = ai_data.get("ot_miert_agak", [])
  p_defs = [
      agak[i].get("problema", "") if i < len(agak) else "" for i in range(5)
  ]
  whys = [
      [agak[i].get("miert1", "") if i < len(agak) else "" for i in range(5)],
      [agak[i].get("miert2", "") if i < len(agak) else "" for i in range(5)],
      [agak[i].get("miert3", "") if i < len(agak) else "" for i in range(5)],
      [agak[i].get("miert4", "") if i < len(agak) else "" for i in range(5)],
      [
          agak[i].get("miert5_gyokerok", "") if i < len(agak) else ""
          for i in range(5)
      ],
  ]

  for c_i in range(5):
    cell = p2_table.cell(0, c_i)
    cell.text = "PROBLÉMA MEGHATÁROZÁSA"
    set_cell_bg(cell, "F2F2F2")
    set_cell_pad(cell, 35, 35, 40, 40)
    set_cell_border(cell, "6")
    cell.paragraphs[0].runs[0].font.name = "Calibri"
    cell.paragraphs[0].runs[0].font.size = Pt(8)
    cell.paragraphs[0].runs[0].font.bold = True

    cell_p = p2_table.cell(1, c_i)
    cell_p.text = p_defs[c_i]
    set_cell_bg(cell_p, "FFFFFF")
    set_cell_pad(cell_p, 35, 35, 40, 40)
    set_cell_border(cell_p, "4")
    if cell_p.paragraphs[0].runs:
      cell_p.paragraphs[0].runs[0].font.name = "Calibri"
      cell_p.paragraphs[0].runs[0].font.size = Pt(7.5)

  for step in range(5):
    h_row = 2 + step * 2
    d_row = 3 + step * 2
    for c_i in range(5):
      c_h = p2_table.cell(h_row, c_i)
      c_h.text = "( MIÉRT?"
      set_cell_bg(c_h, "F2F2F2")
      set_cell_pad(c_h, 25, 25, 40, 40)
      set_cell_border(c_h, "6")
      c_h.paragraphs[0].runs[0].font.name = "Calibri"
      c_h.paragraphs[0].runs[0].font.size = Pt(7.5)
      c_h.paragraphs[0].runs[0].font.bold = True

      c_d = p2_table.cell(d_row, c_i)
      c_d.text = whys[step][c_i]
      set_cell_bg(c_d, "FFFFFF")
      set_cell_pad(c_d, 35, 35, 40, 40)
      set_cell_border(c_d, "4")
      if c_d.paragraphs[0].runs:
        c_d.paragraphs[0].runs[0].font.name = "Calibri"
        c_d.paragraphs[0].runs[0].font.size = Pt(7.5)
        if step == 4 and c_i < len(agak):
          c_d.paragraphs[0].runs[0].font.bold = True

  for row in p2_table.rows:
    for c_i in range(5):
      row.cells[c_i].width = col_w

  doc.add_paragraph().paragraph_format.space_before = Pt(4)
  p_n = doc.add_paragraph()
  p_n.add_run("JEGYZETEK / AKCIÓTERV:").bold = True
  n_box = doc.add_table(rows=1, cols=1)
  n_box.alignment = WD_TABLE_ALIGNMENT.CENTER
  n_cell = n_box.cell(0, 0)
  n_cell.width = Inches(10.7)
  set_cell_bg(n_cell, "FFFFFF")
  set_cell_pad(n_cell, 50, 50, 60, 60)
  set_cell_border(n_cell, "6")

  alap_jegyzet = f"{jegyzet}\n\nFejlesztési cél:\n{fejlesztes}"
  n_cell.text = ai_data.get("vegleges_akcioterv") or alap_jegyzet
  n_cell.paragraphs[0].runs[0].font.name = "Calibri"
  n_cell.paragraphs[0].runs[0].font.size = Pt(8.5)

  for section in doc.sections:
    footer = section.footer
    foot_table = footer.add_table(rows=2, cols=6, width=Inches(10.7))
    foot_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers_f = [
        "Szervezet / Részleg",
        "Revízió",
        "1. kiadás dátuma",
        "Felülvizsgálat dátuma:",
        "Készítő / Felülvizsgáló:",
        "Regisztrációs szám",
    ]
    vals_f = [
        "BSHM / B-ME",
        "3.",
        "2018.07.16",
        "2019.04.03.",
        "Janóczki M",
        "BME001-F-BDE-3",
    ]
    widths_f = [
        Inches(1.8),
        Inches(1.2),
        Inches(1.6),
        Inches(1.8),
        Inches(2.5),
        Inches(1.8),
    ]
    for c_i in range(6):
      c1 = foot_table.cell(0, c_i)
      c1.text = headers_f[c_i]
      set_cell_bg(c1, "F2F2F2")
      set_cell_pad(c1, 20, 20, 30, 30)
      set_cell_border(c1, "4")
      c1.paragraphs[0].runs[0].font.name = "Calibri"
      c1.paragraphs[0].runs[0].font.size = Pt(7)
      c1.paragraphs[0].runs[0].font.bold = True

      c2 = foot_table.cell(1, c_i)
      c2.text = vals_f[c_i]
      set_cell_bg(c2, "FFFFFF")
      set_cell_pad(c2, 20, 20, 30, 30)
      set_cell_border(c2, "4")
      c2.paragraphs[0].runs[0].font.name = "Calibri"
      c2.paragraphs[0].runs[0].font.size = Pt(7.5)
    for r_f in foot_table.rows:
      for c_i, w in enumerate(widths_f):
        r_f.cells[c_i].width = w

  output = io.BytesIO()
  doc.save(output)
  output.seek(0)
  return output


# --- KÉT JÓL LÁTHATÓ GOMB ---
st.markdown("---")
c_btn1, c_btn2 = st.columns(2)

indit_szaki = False
indit_std = False

with c_btn1:
  if st.button(
      "🛠️ GYÁRI SZAKI ELEMZÉS (ZERO-BULLSHIT / NEM AI ÍZŰ)",
      key="btn_szaki",
      use_container_width=True,
  ):
    indit_szaki = True
  st.caption(
      "⚡ Tömör, katonás, életszerű gépész megfogalmazás: csak a tények, nulla"
      " tankönyvi rizsa."
  )

with c_btn2:
  if st.button(
      "⚡ RÉSZLETES MÉRNÖKI ELEMZÉS", key="btn_std", use_container_width=True
  ):
    indit_std = True
  st.caption("📘 Részletesebb szakmai magyarázatok.")

if indit_szaki or indit_std:
  api_key = api_key_input.strip()
  if not api_key:
    st.error(
        "⚠️ Kérlek, add meg a Google Gemini API kulcsot a bal oldali sávban!"
    )
  else:
    is_szaki = indit_szaki
    uzenet = (
        "🛠️ Gyári szaki elemzés fut (Zero-Bullshit mód)..."
        if is_szaki
        else "🧠 Mérnöki elemzés fut..."
    )
    with st.spinner(uzenet):
      try:
        ai_eredmeny = elemez_geminivel(
            api_key,
            gep_nev,
            hiba_rovid,
            datum_val,
            felelos_val,
            jegyzet_szoveg,
            fejlesztesi_cel,
            kivalasztott_kriteriumok,
            zero_bullshit=is_szaki,
        )
        docx_file = general_hivatalos_docx(
            gep_nev,
            hiba_rovid,
            datum_val,
            felelos_val,
            ai_eredmeny,
            jegyzet_szoveg,
            fejlesztesi_cel,
            kivalasztott_kriteriumok,
        )

        siker_szoveg = (
            "✅ Kész! Az elemzés lényegretörő, valódi gyári stílusban készült"
            " el (nulla AI-rizsa)!"
            if is_szaki
            else "✅ Az elemzés sikeresen elkészült!"
        )
        st.success(siker_szoveg)
        st.download_button(
            label="📥 HIVATALOS WORD DOKUMENTUM LETÖLTÉSE (.DOCX)",
            data=docx_file,
            file_name=(
                f"{gep_nev.replace(' ', '_')}_{hiba_rovid.replace(' ', '_')}_IDA.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )
      except Exception as e:
        st.error(f"Hiba történt az elemzés során: {e}")
