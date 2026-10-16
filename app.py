import io
import json
import unicodedata
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import google.generativeai as genai
import streamlit as st

# --- OLDALBEÁLLÍTÁSOK ---
st.set_page_config(
    page_title="BSHM / B-ME // IDA",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- CSICSA & HIGH-TECH STÍLUS (ÉLES, KONTRASZTOS FEHÉR FELIRATOKKAL ÉS CHECKBOXOKKAL) ---
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

    /* ABLAKOK FELETTI CÍMKÉK: ÉLES HÓFEHÉR, BOLD, 15PX */
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

    /* CHECKBOXOK STÍLUSA */
    div[data-testid="stCheckbox"] label p,
    div[data-testid="stCheckbox"] label span {
        font-size: 15px !important;
        font-weight: 700 !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        letter-spacing: 0.5px !important;
    }

    /* BEMENETI MEZŐK BELSŐ TARTALMA */
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

    /* GOMBOK ÁLTALÁNOS FORMÁZÁSA */
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

    /* 1. SZAKI GOMB: ÉLÉNK NARANCSSÁRGA HÁTTÉR */
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

    /* 2. MÉRNÖKI GOMB: KOBALTKÉK HÁTTÉR */
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

    /* LETÖLTÉS GOMB */
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

# --- OLDALSÁV: API KULCS BEÁLLÍTÁSA ---
with st.sidebar:
  st.markdown("### ⚙️ AI Rendszer Beállítás")
  api_key_input = st.text_input(
      "Google Gemini API Kulcs:",
      type="password",
      value=st.secrets.get("GEMINI_API_KEY", ""),
      help="Illeszd be az AI Studio kulcsodat",
  )
  st.markdown("---")
  st.markdown(
      "💡 **Tipp:** Ha a Streamlit Secrets-be beírtad a `GEMINI_API_KEY`-t,"
      " üresen maradhat."
  )

# --- FEJLÉC ÉS KPI STATUS BAR ---
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
      ' Státusz</div><div class="kpi-value" style="color: #4ade80;">● ONLINE /'
      " AI READY</div></div>",
      unsafe_allow_html=True,
  )
with col_k2:
  st.markdown(
      '<div class="kpi-card"><div class="kpi-title">Sablon Revízió</div><div'
      ' class="kpi-value">Rev. 3 (2019.04.03)</div></div>',
      unsafe_allow_html=True,
  )
with col_k3:
  st.markdown(
      '<div class="kpi-card"><div class="kpi-title">Regisztráció</div><div'
      ' class="kpi-value">BME001-F-BDE-3</div></div>',
      unsafe_allow_html=True,
  )
with col_k4:
  st.markdown(
      '<div class="kpi-card"><div class="kpi-title">Ellenőrzési'
      ' Felelős</div><div class="kpi-value">Janóczki M</div></div>',
      unsafe_allow_html=True,
  )

st.markdown("<br>", unsafe_allow_html=True)

# --- BEVITELI MEZŐK ---
c_left, c_right = st.columns([1, 1.4])
with c_left:
  st.markdown("### 📋 Alapadatok")
  gep_nev = st.text_input("Berendezés / Gép megnevezése", value="12k pit")
  hiba_rovid = st.text_input("Hiba megnevezése", value="12k pit hiba")
  datum_val = st.text_input("Dátum", value="2026.10.14.")
  felelos_val = st.text_input("Karbantartási Felelős", value="Nagy Attila")

with c_right:
  st.markdown("### 📝 Esemény & Célzott Fejlesztési Irány")
  jegyzet_szoveg = st.text_area(
      "1. Mi volt a hiba? (Műszakos jegyzet / tényleges hiba):",
      height=110,
      value=(
          "(PIT1 becsatlakozó cv-n heveder csere megtörtént és a cv közelebb"
          " lett állítva a lehordó cv-hez.)"
      ),
  )
  fejlesztesi_cel = st.text_area(
      (
          "2. Mit hozzunk ki fejlesztésnek? (Elvárt végleges megoldás /"
          " intézkedés):"
      ),
      height=90,
      value=(
          "a régi Segment gateway controllok lecserélése multicontroll-ra a"
          " gyorsabb hibakeresés miatt és hogy van hozzá még csere alkatrész a"
          " régiekkel ellentétben"
      ),
  )

# --- KRITÉRIUMOK LISTÁJA (A GYÁRI SORSZÁM ÉS SORREND SZERINT) ---
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

# --- BEPIPÁLHATÓ KRITÉRIUMOK KIVÁLASZTÁSA (MIND A 13 ALAPÉRTELMEZETTEN ÜRESEN INDUL) ---
st.markdown("### ☑️ Kifejtendő Kritériumok kiválasztása (az 1. oldalhoz)")
st.markdown(
    "<span style='color: #94a3b8; font-size: 13px;'>Pipáld be azokat a"
    " kritériumokat, amiket ki akarsz fejteni. <b>Ahova nem kerül pipa, oda"
    " automatikusan a 'Nem Releváns' bejegyzés kerül!</b></span>",
    unsafe_allow_html=True,
)

cb_col1, cb_col2, cb_col3, cb_col4 = st.columns(4)

with cb_col1:
  cb_tisztasag = st.checkbox("Tisztaság", value=False, key="chk_tisztasag")
  cb_meghuzottsag = st.checkbox(
      "Meghúzottság", value=False, key="chk_meghuzottsag"
  )
  cb_kenes = st.checkbox("Kenés", value=False, key="chk_kenes")
  cb_karbantartas = st.checkbox(
      "Karbantartás", value=False, key="chk_karbantartas"
  )

with cb_col2:
  cb_kornyezet = st.checkbox("Környezet", value=False, key="chk_kornyezet")
  cb_mukodtetis = st.checkbox("Működtetés", value=False, key="chk_mukodtetis")
  cb_specifikacio = st.checkbox(
      "Specifikáció", value=False, key="chk_specifikacio"
  )

with cb_col3:
  cb_mukodes = st.checkbox("Működés", value=False, key="chk_mukodes")
  cb_telepites = st.checkbox("Telepítés", value=False, key="chk_telepites")
  cb_osszeszereles = st.checkbox(
      "Összeszerelés", value=False, key="chk_osszeszereles"
  )

with cb_col4:
  cb_gyartas = st.checkbox("Gyártás", value=False, key="chk_gyartas")
  cb_tervezes = st.checkbox("Tervezés", value=False, key="chk_tervezes")
  cb_javitas = st.checkbox("Javítás", value=False, key="chk_javitas")

kivalasztott_kriteriumok = []
if cb_tisztasag:
  kivalasztott_kriteriumok.append("Tisztaság")
if cb_meghuzottsag:
  kivalasztott_kriteriumok.append("Meghúzottság")
if cb_kenes:
  kivalasztott_kriteriumok.append("Kenés")
if cb_karbantartas:
  kivalasztott_kriteriumok.append("Karbantartás")
if cb_kornyezet:
  kivalasztott_kriteriumok.append("Környezet")
if cb_mukodtetis:
  kivalasztott_kriteriumok.append("Működtetés")
if cb_specifikacio:
  kivalasztott_kriteriumok.append("Specifikáció")
if cb_mukodes:
  kivalasztott_kriteriumok.append("Működés")
if cb_telepites:
  kivalasztott_kriteriumok.append("Telepítés")
if cb_osszeszereles:
  kivalasztott_kriteriumok.append("Összeszerelés")
if cb_gyartas:
  kivalasztott_kriteriumok.append("Gyártás")
if cb_tervezes:
  kivalasztott_kriteriumok.append("Tervezés")
if cb_javitas:
  kivalasztott_kriteriumok.append("Javítás")


# --- ROBUSZTUS SEGÉDFÜGGVÉNYEK AZ ADATKINYERÉSHEZ ---
def norm(s):
  if not s:
    return ""
  return "".join(
      c
      for c in unicodedata.normalize("NFD", str(s).lower())
      if unicodedata.category(c) != "Mn"
  ).strip()


def extract_criterion(ai_data, krit_nev):
  raw_kriteriumok = ai_data.get("kriteriumok", {})
  t_norm = norm(krit_nev)
  found_dict = None

  if isinstance(raw_kriteriumok, dict):
    for k, v in raw_kriteriumok.items():
      k_norm = norm(k)
      if t_norm in k_norm or k_norm in t_norm:
        found_dict = v
        break
  elif isinstance(raw_kriteriumok, list):
    for item in raw_kriteriumok:
      if isinstance(item, dict):
        item_name = norm(
            item.get("nev")
            or item.get("név")
            or item.get("kriterium")
            or item.get("kritérium")
            or item.get("name")
            or ""
        )
        if t_norm in item_name or item_name in t_norm:
          found_dict = item
          break

  if not isinstance(found_dict, dict):
    return {}

  def get_val(d, *keys):
    for k in keys:
      norm_k = norm(k)
      for dk, dv in d.items():
        if norm(dk) == norm_k and dv:
          return str(dv).strip()
    return ""

  return {
      "idealis": get_val(
          found_dict,
          "idealis",
          "ideális",
          "ideal",
          "idealis_allapot",
          "ideális_állapot",
      ),
      "aktualis": get_val(
          found_dict,
          "aktualis",
          "aktuális",
          "actual",
          "aktualis_allapot",
          "aktuális_állapot",
      ),
      "hianyossag": get_val(
          found_dict,
          "hianyossag",
          "hiányosság",
          "hiany",
          "hiány",
          "defect",
      ),
      "gyokerok": get_val(
          found_dict,
          "gyokerok",
          "gyökérok",
          "gyoker_ok",
          "gyökér_ok",
          "root_cause",
      ),
      "ellenintezkedes": get_val(
          found_dict,
          "ellenintezkedes",
          "ellenintézkedés",
          "intezkedes",
          "intézkedés",
          "countermeasure",
      ),
  }


# --- INTELLIGENS FLASH MOTOR (BRIDGESTONE ÜZEMI GYAKORLAT SZERINTI MÉLYSÉGGEL) ---
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

  kivalasztottak_szoveg = (
      ", ".join(kivalasztott) if kivalasztott else "Nincs bejelölve"
  )

  prompt = f"""
    Te a Bridgestone Tatabánya tapasztalt mechatronikai karbantartási specialistája és műszaki vezetője vagy.
    Készíts hivatalos BSHM B-ME IDA elemzést a megadott adatok alapján, szigorúan a gyári valóság és az üzemeltetési gyakorlat szerint!

    Berendezés: {gep}
    Hiba: {hiba}
    Dátum: {datum}
    Felelős: {felelos}

    1. VALÓS ESEMÉNY / TÉNYLEGESEN MEGTÖRTÉNT HIBA:
    {jegyzet}

    2. CÉLZOTT FEJLESZTÉS / ELVÁRT VÉGLEGES MEGOLDÁS:
    {fejlesztes}

    3. KIFEJTENDŐ KRITÉRIUM(OK) AZ 1. OLDALON:
    {kivalasztottak_szoveg}

    SZIGORÚ TARTALMI ÉS FORMÁTUMKÖVETELMÉNYEK:
    1. TILOS ÁLTALÁNOS SABLONSZÖVEGEKET ÍRNI!
       Tilosak az ilyen mondatok: "1. Standard előírások betartása.", "1. Megfelelő üzemállapot.", "1. Eltérés tapasztalható.", "1. Kiváltó ok meghatározva.", "1. Intézkedés rögzítve."!
       Minden mező legyen 100%-ban konkrét, szakmai, a megadott jegyzetből és célzott fejlesztésből vett valós kifejezésekkel (pl. cv távolság, abroncs becsípődés, Interroll rendszer, Segment gateway kontroller, 78 db segment vezérlő, multicontroll modernizálás, hibakeresési diagnosztika, alkatrész elérhetőség, capex/opex)!

    2. AZ 1. OLDAL KRITÉRIUM CELLÁINAK FELÉPÍTÉSE (MINDEN CELLÁBAN 2 KONKRÉT SZÁMOZOTT PONT):
       Példa a 'Telepítés' kritériumra a gyári jegyzet alapján:
       - Ideális állapot:
         1. cv megfelelően van letelepítve, így nem csípődik be az abroncs
         2. Korszerű vezérlőrendszerrel (pl. multicontroll) szerelt cv, amihez biztosított a cserealkatrész
       - Aktuális állapot:
         1. Túl nagy távolságra lett telepítve a lehordó cv-hez képest, ezért az abroncs becsípődik
         2. Régi elavult Segment gateway vezérlés (1 gateway 78 segment kontrollert kezel), nehézkes hibakeresés, nincs hozzá alkatrész
       - Hiányosság:
         1. túl nagy cv távolság miatt cv becsípődött
         2. olyan cv vezérlő rendszer ahol 1 gateway controll 78 db segment controllt vezérel ezért nehéz a hibakeresés
       - Hiányosság gyökéroka:
         1. telepítési / geometriai hiányosság
         2. gazdaságosabb vezérlőrendszert választottak telepítéskor
       - Ellenintézkedés:
         1. cv beállítása megtörtént közelebb a lehordó cv-hez és heveder cserélve
         2. szükséges áttérni multicontroll vezérlésre a gyors hibakeresés és alkatrész-ellátás érdekében (capex/opex projekt)
         Felelős: ..........
         Határidő: ..........

    3. KÖZVETLEN KAPCSOLAT AZ 1. ÉS 2. OLDAL KÖZÖTT:
       - Az 1. oldalon a 'Hiányosság' oszlopba beírt számozott megállapítások (1., 2.) SZÓ SZERINT kerüljenek be a 2. oldal 'PROBLÉMA MEGHATÁROZÁSA' soraiba (1. ág, 2. ág, 3. ág)!
       - Az 5 Miértnél nem kell erőltetni az 5 szintet! 2-3 logikus lépésből érjük el a valós gyökérokot (pl. miért 1: túl nagy távolság -> miért 2: így telepítették le a gépet; vagy miért 1: nagy méretek gyártása -> miért 2: változott a termékpaletta; vagy miért 1: nem tölthető program a modulra -> miért 2: elavult megszűnt rendszer -> miért 3: beruházáskor gazdaságosabb volt ezt megvenni). A további miért mezők legyenek üres stringek ("")!
       - Csak 3 ágat tölts ki, a 4. és 5. ág maradjon üresen!

    KIZÁRÓLAG érvényes JSON formátumban válaszolj az alábbi kulcsokkal:
    {{
      "kriteriumok": {{
        "Telepítés": {{
          "idealis": "1. ...\\n2. ...",
          "aktualis": "1. ...\\n2. ...",
          "hianyossag": "1. ...\\n2. ...",
          "gyokerok": "1. ...\\n2. ...",
          "ellenintezkedes": "1. ...\\n2. ...\\nFelelős: ..........\\nHatáridő: .........."
        }}
      }},
      "ot_miert_agak": [
        {{
          "problema": "1. túl nagy cv távolság miatt cv becsípődött",
          "miert1": "cv-k között túl nagy a távolság",
          "miert2": "így telepítették le a gépet",
          "miert3": "",
          "miert4": "",
          "miert5": ""
        }},
        {{
          "problema": "1. túl nagy cv távolság miatt cv becsípődött",
          "miert1": "eddig nem gyártott nagy méretek gyártása",
          "miert2": "változott a termékpaletta",
          "miert3": "",
          "miert4": "",
          "miert5": ""
        }},
        {{
          "problema": "2. olyan cv vezérlő rendszer ahol 1 gateway controll 78 db segment controllt vezérel ezért nehéz a hibakeresés",
          "miert1": "a vezérlő látja a modulokat de nem tud rá programot feltölteni ezért nehezen azonosítható a hiba",
          "miert2": "obszolét kifutó rendszerrel lett szerelve a cv",
          "miert3": "gazdaságosabb volt ezt megvenni telepítéskor",
          "miert4": "",
          "miert5": ""
        }}
      ],
      "vegleges_akcioterv": "1. CV közelebb állítása és hevedercsere elvégezve.\\n2. Segment gateway cseréje multicontroll-ra betervezve.\\nFelelős: ..........\\nHatáridő: .........."
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


# --- DOKUMENTUM GENERÁLÓ MOTOR (100% GYÁRI FORMÁTUM, CSAK JANÓCZKI M!) ---
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

  raw_kriteriumok = ai_data.get("kriteriumok", {})
  crit_dict = {}
  if isinstance(raw_kriteriumok, dict):
    crit_dict = raw_kriteriumok
  elif isinstance(raw_kriteriumok, list):
    for item in raw_kriteriumok:
      if isinstance(item, dict) and "nev" in item:
        crit_dict[item["nev"]] = item

  for r_idx, (krit_nev, krit_kerdes) in enumerate(kriterium_lista, start=2):
    is_selected = krit_nev in kivalasztott
    data = extract_criterion(ai_data, krit_nev) if is_selected else {}

    if is_selected and data:
      ell = data.get("ellenintezkedes", "")
      if ell and "Felelős:" not in ell:
        ell = f"{ell}\nFelelős: ..........\nHatáridő: .........."
      col_vals = [
          f"{krit_nev}\n{krit_kerdes}",
          data.get("idealis", ""),
          data.get("aktualis", ""),
          data.get("hianyossag", ""),
          data.get("gyokerok", ""),
          ell,
      ]
    else:
      col_vals = [
          f"{krit_nev}\n{krit_kerdes}",
          "Nem Releváns",
          "Nem Releváns",
          "Nem Releváns",
          "Nem Releváns",
          "Nem Releváns",
      ]

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
      [agak[i].get("miert5", "") if i < len(agak) else "" for i in range(5)],
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
      c_h.text = "↓ MIÉRT?"
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
            "✅ Kész! Az elemzés a fotózott mintád szerint, gyári logikával"
            " készült el!"
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
