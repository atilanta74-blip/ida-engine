import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import time

# --- OLDALBEÁLLÍTÁSOK ---
st.set_page_config(
    page_title="B-ME TPM AI // IDA Engine",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CSICSA & HIGH-TECH STÍLUS (CSS INJECTION) ---
st.markdown("""
<style>
    /* Sötét mérnöki háttér */
    .stApp {
        background: linear-gradient(135deg, #0b0f19 0%, #111827 50%, #070a12 100%);
        color: #e2e8f0;
        font-family: 'Segoe UI', Roboto, sans-serif;
    }
    
    /* Felső Banner */
    .header-box {
        background: rgba(17, 24, 39, 0.8);
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
    }

    /* KPI Kártyák */
    .kpi-container {
        display: flex;
        gap: 15px;
        margin-bottom: 25px;
    }
    .kpi-card {
        flex: 1;
        background: rgba(30, 41, 59, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.08);
        padding: 14px 18px;
        border-radius: 10px;
        backdrop-filter: blur(10px);
    }
    .kpi-title {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #94a3b8;
    }
    .kpi-value {
        font-size: 17px;
        font-weight: 700;
        color: #38bdf8;
        margin-top: 3px;
    }

    /* Bemeneti mezők stílusa */
    div[data-baseweb="input"], div[data-baseweb="textarea"] {
        background-color: #0f172a !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        color: #f8fafc !important;
    }
    label p {
        font-size: 13px !important;
        font-weight: 600 !important;
        color: #cbd5e1 !important;
        letter-spacing: 0.5px;
    }

    /* Generálás Gomb */
    .stButton>button {
        width: 100%;
        background: linear-gradient(90deg, #0284c7 0%, #2563eb 100%) !important;
        color: #ffffff !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        letter-spacing: 1.5px !important;
        text-transform: uppercase !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        box-shadow: 0 4px 15px rgba(2, 132, 199, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(2, 132, 199, 0.6) !important;
    }

    /* Letöltés Gomb */
    .stDownloadButton>button {
        width: 100%;
        background: linear-gradient(90deg, #059669 0%, #10b981 100%) !important;
        color: white !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        letter-spacing: 1px;
        box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4) !important;
    }
</style>
""", unsafe_allow_html=True)

# --- FEJLÉC ÉS KPI STATUS BAR ---
st.markdown("""
<div class="header-box">
    <div class="header-title">⚡ BSHM / B-ME // TPM KAIZEN AI CORE</div>
    <div class="header-sub">AUTOMATED USAGE GAP ASSESSMENT & 5-WHY DEEP ANALYSIS</div>
</div>
""", unsafe_allow_html=True)

col_k1, col_k2, col_k3, col_k4 = st.columns(4)
with col_k1:
    st.markdown('<div class="kpi-card"><div class="kpi-title">Rendszer Státusz</div><div class="kpi-value" style="color: #4ade80;">● ONLINE / READY</div></div>', unsafe_allow_html=True)
with col_k2:
    st.markdown('<div class="kpi-card"><div class="kpi-title">Sablon Revízió</div><div class="kpi-value">Rev. 3 (2019.04.03)</div></div>', unsafe_allow_html=True)
with col_k3:
    st.markdown('<div class="kpi-card"><div class="kpi-title">Regisztráció</div><div class="kpi-value">BME001-F-BDE-3</div></div>', unsafe_allow_html=True)
with col_k4:
    st.markdown('<div class="kpi-card"><div class="kpi-title">Ellenőrzési Felelős</div><div class="kpi-value">Janóczki M</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- BEVITELI MEZŐK ---
c_left, c_right = st.columns([1, 1.4])

with c_left:
    st.markdown("### 📋 Alapadatok")
    gep_nev = st.text_input("Berendezés / Manipulátor", value="SL2 manipulátor")
    hiba_rovid = st.text_input("Hiba megnevezése", value="CT érzékelő kábel törés")
    datum_val = st.text_input("Dátum", value="2026.09.21.")
    felelos_val = st.text_input("Karbantartási Felelős", value="Nagy Attila")

with c_right:
    st.markdown("### 📝 Műszakos Jegyzetek & Tények")
    jegyzet_szoveg = st.text_area(
        "Másold be az eseményt / megfigyelést:",
        height=188,
        value=(
            "A SL2 2-es manipulátor ct érzékelő szenzor kábel a sok mozgástól megtörött . "
            "A törés érzéklehető volt. Csere kábel nincs ,ezért kábelkötegelővel rögzítve lett.\n"
            "Akció: kábel csatlakozóval beazonosítás, SAP igénylés, Rendelés.  Felelős: Nagy Attila.  Határidő :2026.10.15.\n"
            "Másnap a raktárban megtaláltuk a szenzor kábelt .SAP: 2033358"
        )
    )

# --- DOKUMENTUM GENERÁLÓ MOTOR (100% GYÁRI FORMÁTUM, CSICSA MENTES BELSŐVEL) ---
def general_hivatalos_docx(gep, hiba, datum, felelos, jegyzet):
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
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

    def set_cell_pad(cell, top=50, bottom=50, left=60, right=60):
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'))

    def set_cell_border(cell, sz="4"):
        cell._tc.get_or_add_tcPr().append(parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>
                <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>
                <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>
                <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>
            </w:tcBorders>
        '''))

    # 1. Oldal fejléc
    p_title = doc.add_paragraph()
    r = p_title.add_run("Üzemzavar elhárítási napi irányítási rendszer – Berendezés hiányosság értékelés, mélyanalízis, ellenintézkedési lap")
    r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.bold = True

    p_info = doc.add_paragraph()
    p_info.add_run("Üzemzavar leírása: ").bold = True
    p_info.add_run(f"{gep} – {hiba}          ").underline = True
    p_info.add_run("Dátum: ").bold = True
    p_info.add_run(f"{datum}          ").underline = True
    p_info.add_run("Karb. felelős: ").bold = True
    p_info.add_run(f"{felelos}").underline = True

    # 1. Oldal: 13 Kritérium táblázat
    p1_table = doc.add_table(rows=15, cols=6)
    p1_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    p1_widths = [Inches(1.8), Inches(1.75), Inches(1.75), Inches(1.75), Inches(1.85), Inches(1.8)]

    # Fő fejlécek
    c_plan = p1_table.cell(0, 0)
    c_plan.merge(p1_table.cell(0, 1)).merge(p1_table.cell(0, 2)).merge(p1_table.cell(0, 3))
    c_plan.text = "TERVEZD MEG (PLAN)"
    p1_table.cell(0, 4).text = "VALÓSÍTSD MEG (DO)"
    p1_table.cell(0, 5).text = "ELLENŐRIZD (CHECK), STANDARDIZÁLD (ACT)"

    sub_heads = ["KRITÉRIUM", "IDEÁLIS ÁLLAPOT", "AKTUÁLIS ÁLLAPOT", "HIÁNYOSSÁG", "Hiányosság gyökéroka\n(ismert, vagy “5 miért” elemzés alapján)", "Ellenintézkedés\n(határidő, felelős, standardizálás)"]
    for c_i, h in enumerate(sub_heads):
        p1_table.cell(1, c_i).text = h

    for r_i in range(2):
        for cell in p1_table.rows[r_i].cells:
            set_cell_bg(cell, "F2F2F2")
            set_cell_pad(cell, 60, 60, 60, 60)
            set_cell_border(cell, sz="6")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.name = 'Calibri'; p.runs[0].font.size = Pt(7.5); p.runs[0].font.bold = True

    crit_rows = [
        ("Tisztaság\n(A részegység elég tiszta, hogy ellenőrizhető legyen?)", "1. Tiszta, lerakódásmentes felület.", "1. Megfelelő felületi tisztaság.", "1. Nincs hiányosság.", "1. Nem releváns.", "1. Standard CIL tisztítás fenntartása."),
        ("Meghúzottság\n(A kötőelemek megfelelően meg vannak húzva?)", "1. Kötőelemek előírt nyomatékkal rögzítve.", "1. Csavarok és rögzítők stabilak.", "1. Nincs hiányosság.", "1. Nem releváns.", "1. Éves mechanikai ellenőrzés."),
        ("Kenés\n(A berendezés a kenési terv szerint kenve?)", "1. Megfelelő kenőanyag és mennyiség.", "1. Csapágyazások és vezetők kenve.", "1. Nem releváns szenzorhibánál.", "1. Nem releváns.", "1. Kenési standard fenntartása."),
        ("Karbantartás\n(Szükséges elemek előírás szerint karbantartva?)", "1. Dinamikus mozgó kábelek periodikus állapotfelmérése és cseréje.", "1. A szenzorkábel nem szerepelt megelőző ellenőrzési ciklusban.", "1. Hiányzik a mozgó kábelek ciklikus felmérése és megelőző cseréje.", "1. A PM terv nem tér ki külön a hajlított kábelek fáradásos vizsgálatára.", f"1. Kábelek felvétele a negyedéves PM listára.\nFelelős: {felelos}, Határidő: 2026.10.15."),
        ("Környezet\n(Előírásoknak megfelelő környezet?)", "1. Normál üzemi hőmérséklet, vegyszermentesség.", "1. Normál gyártócsarnoki környezet.", "1. Nincs hiányosság.", "1. Nem releváns.", "1. Környezeti feltételek fenntartása."),
        ("Működtetés\n(Munkafolyamatokat következetesen betartják?)", "1. Standard kezelői manipulációs pálya.", "1. Kezelők az előírás szerint üzemeltetik.", "1. Nincs kezelői hiba.", "1. Nem releváns.", "1. Munkautasítás fenntartása."),
        ("Specifikáció\n(Határértékeken belül működtetik?)", "1. Névleges határokon belüli sebesség.", "1. Gyári paramétereken belüli üzem.", "1. Nincs határérték-túllépés.", "1. Nem releváns.", "1. Paraméterek felügyelete."),
        ("Működés\n(Végez nem megfelelő műveletet?)", "1. Hibátlan pozícióérzékelés, zavartalan ciklus.", "1. CT jelvesztés miatt manipulátor pozícióhiba és sorleállás.", "1. Jelátviteli hiba a CT szenzor törött kábele miatt.", "1. Kábeltörés a dinamikus mozgási zónában.", f"1. Ideiglenes rögzítés után új kábel beépítése.\nFelelős: {felelos}"),
        ("Telepítés\n(Telepítéskor követték az útmutatót?)", "1. Védőcsőben vezetett kábelezés, megfelelő rádiusszal.", "1. Túl szűk hajlítási rádiusz, pontszerű terhelés.", "1. Elégtelen hajlítási sugár a mozgó csuklópontnál.", "1. A rögzítés nem adott elég mozgásteret a ciklikus forgáshoz.", f"1. Új kábelvezetés kialakítása nagyobb ívvel és védőcsővel.\nFelelős: {felelos}, Hat.: 2026.10.15."),
        ("Összeszerelés\n(Alkatrészek megfelelően illeszkednek?)", "1. Akadálymentes mozgás dörzsölődés nélkül.", "1. Külső dörzsölődés nem látható.", "1. Nincs hiányosság.", "1. Nem releváns.", "1. Standard állapot fenntartása."),
        ("Gyártás\n(Rajzoknak megfelelő alkatrészek?)", "1. Dinamikus robotmozgásra méretezett kábeltípus.", "1. Standard kábel került beépítésre, ciklusszám végén kifáradt.", "1. Kábel fáradási élettartama lejárt.", "1. Dinamikus igénybevétel miatti rézér-fáradás.", "1. Gyári magas flexibilitású kábel alkalmazása (SAP: 2033358)."),
        ("Tervezés\n(Alkatrészek megbízhatóan működnek?)", "1. Megbízható, törésmentes villamos jelátvitel.", "1. Sok mozgástól a kábel mechanikusan megtört, kézzel érzékelhető volt.", "1. CT szenzorkábel fizikai törése.", "1. Fárasztó hajlítófeszültség a kilépési pontnál.", "1. Kábelcsere és rögzítési tehermentesítés."),
        ("Javítás\n(Üzemzavar elhárítása elvárt idő alatt?)", "1. Azonnali alkatrész-elérhetőség és csere.", "1. Cserekábel a gépnél nem állt rendelkezésre; kötegelővel ideiglenesen stabilizálva.", "1. Nem volt azonnali kábel kéznél; SAP azonosító hiányzott.", "1. Cikkszám nem szerepelt a gép alkatrészlistáján; másnap lett meg (SAP: 2033358).", f"1. Kötegelős rögzítés kész.\n2. Raktárban megtalálva (SAP: 2033358).\n3. Csere és utánrendelés: {felelos}, Hat.: 2026.10.15.")
    ]

    for r_idx, c_tuple in enumerate(crit_rows, start=2):
        for c_i, val in enumerate(c_tuple):
            cell = p1_table.cell(r_idx, c_i)
            cell.text = val
            set_cell_bg(cell, "FFFFFF")
            set_cell_pad(cell, 35, 35, 50, 50)
            set_cell_border(cell, sz="4")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = Pt(9)
            if p.runs:
                p.runs[0].font.name = 'Calibri'; p.runs[0].font.size = Pt(7); p.runs[0].font.color.rgb = RGBColor(0,0,0)
                if c_i == 0: p.runs[0].font.bold = True

    for row in p1_table.rows:
        for c_i, w in enumerate(p1_widths):
            row.cells[c_i].width = w

    # 2. Oldal: 5 Miért táblázat (5 oszlop)
    doc.add_page_break()
    p2_title = doc.add_paragraph()
    r = p2_title.add_run("Üzemzavar elhárítási napi irányítási rendszer – Berendezés hiányosság értékelés, mélyanalízis, ellenintézkedési lap")
    r.font.name = 'Calibri'; r.font.size = Pt(11); r.font.bold = True

    p2_info = doc.add_paragraph()
    p2_info.add_run(f"Mélyanalízis (5 Miért elemzés)          Gép: {gep}          Dátum: {datum}          Karb. felelős: {felelos}").bold = True

    p2_table = doc.add_table(rows=12, cols=5)
    p2_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w = Inches(10.7 / 5)

    p_defs = [
        f"1. ÁG: MŰSZAKI ÁG\nA {gep} {hiba} miatt leállt, a kábel a dinamikus mozgástól megtört.",
        "2. ÁG: KARBANTARTÁSI ÁG\nNem volt megelőzve a kábeltörés a korábbi megelőző karbantartási (PM) ciklusok során.",
        "3. ÁG: ALKATRÉSZ-ELLÁTÁSI ÁG\nAz üzemzavar pillanatában nem állt rendelkezésre azonnali cserekábel a gép mellett.",
        "", ""
    ]
    whys = [
        ["Miért állt le a manipulátor?\n-> CT érzékelő nem adott jelet a vezérlésnek.", "Miért nem előztük meg a hibát?\n-> Nem volt állapotellenőrzés a mozgó kábeleken.", "Miért kötegelővel javítottuk?\n-> A helyszínen nem volt azonnali cserealkatrész.", "", ""],
        ["Miért szakadt meg a jelátvitel?\n-> A kábel rézvezetői mechanikusan eltörtek.", "Miért nem ellenőriztük a kábelt?\n-> A szenzorkábel nem szerepelt a PM ellenőrzőlistán.", "Miért nem vettünk ki raktárból?\n-> A pontos SAP cikkszám azonnal nem volt ismert.", "", ""],
        ["Miért tört meg a kábel?\n-> A hajlítási rádiusz túl szűk volt, pontszerű terhelést kapott.", "Miért nem szerepelt a PM-ben?\n-> Nem volt része a dinamikus kábelkorbács-fáradás vizsgálata.", "Miért nem volt kéznél cikkszám?\n-> Nem volt hozzárendelve a gép alkatrészlistájához (BOM).", "", ""],
        ["Miért volt szűk a hajlítási ív?\n-> A rögzítés és vezetés nem adott elég mozgásteret a forgáshoz.", "Miért nem része a standardnak?\n-> Gépátadáskor nem kezelték kritikus kopó alkatrészként.", "Hogyan lett beazonosítva?\n-> Másnap raktárban minta alapján beazonosítva: SAP 2033358.", "", ""],
        ["GYÖKÉROK:\nA kábelvezetés nem biztosított tehermentesítő ívet a hajlítási zónában, ami rézér-fáradáshoz vezetett.", "GYÖKÉROK:\nA PM rendszerből hiányzik a dinamikusan mozgó flexibilis kábelek megelőző csereciklusa.", "GYÖKÉROK:\nA kritikus alkatrész SAP azonosítója hiányzott a gép törzsadat-jegyzékéből (BOM).", "", ""]
    ]

    for c_i in range(5):
        cell = p2_table.cell(0, c_i)
        cell.text = "PROBLÉMA MEGHATÁROZÁSA"
        set_cell_bg(cell, "F2F2F2"); set_cell_pad(cell, 35, 35, 40, 40); set_cell_border(cell, "6")
        cell.paragraphs[0].runs[0].font.name = 'Calibri'; cell.paragraphs[0].runs[0].font.size = Pt(8); cell.paragraphs[0].runs[0].font.bold = True
        
        cell_p = p2_table.cell(1, c_i)
        cell_p.text = p_defs[c_i]
        set_cell_bg(cell_p, "FFFFFF"); set_cell_pad(cell_p, 35, 35, 40, 40); set_cell_border(cell_p, "4")
        if cell_p.paragraphs[0].runs:
            cell_p.paragraphs[0].runs[0].font.name = 'Calibri'; cell_p.paragraphs[0].runs[0].font.size = Pt(7.5)

    for step in range(5):
        h_row = 2 + step * 2; d_row = 3 + step * 2
        for c_i in range(5):
            c_h = p2_table.cell(h_row, c_i)
            c_h.text = "( MIÉRT?"
            set_cell_bg(c_h, "F2F2F2"); set_cell_pad(c_h, 25, 25, 40, 40); set_cell_border(c_h, "6")
            c_h.paragraphs[0].runs[0].font.name = 'Calibri'; c_h.paragraphs[0].runs[0].font.size = Pt(7.5); c_h.paragraphs[0].runs[0].font.bold = True

            c_d = p2_table.cell(d_row, c_i)
            c_d.text = whys[step][c_i]
            set_cell_bg(c_d, "FFFFFF"); set_cell_pad(c_d, 35, 35, 40, 40); set_cell_border(c_d, "4")
            if c_d.paragraphs[0].runs:
                c_d.paragraphs[0].runs[0].font.name = 'Calibri'; c_d.paragraphs[0].runs[0].font.size = Pt(7.5)
                if step == 4 and c_i < 3: c_d.paragraphs[0].runs[0].font.bold = True

    for row in p2_table.rows:
        for c_i in range(5):
            row.cells[c_i].width = col_w

    # Jegyzetek doboz alul
    doc.add_paragraph().paragraph_format.space_before = Pt(4)
    p_n = doc.add_paragraph()
    p_n.add_run("JEGYZETEK / AKCIÓTERV:").bold = True
    n_box = doc.add_table(rows=1, cols=1)
    n_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    n_cell = n_box.cell(0, 0)
    n_cell.width = Inches(10.7)
    set_cell_bg(n_cell, "FFFFFF"); set_cell_pad(n_cell, 50, 50, 60, 60); set_cell_border(n_cell, "6")
    n_cell.text = jegyzet
    n_cell.paragraphs[0].runs[0].font.name = 'Calibri'; n_cell.paragraphs[0].runs[0].font.size = Pt(8.5)

    # Lábléc: PONTOSAN CSAK Janóczki M!
    for section in doc.sections:
        footer = section.footer
        foot_table = footer.add_table(rows=2, cols=6, width=Inches(10.7))
        foot_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers_f = ["Szervezet / Részleg", "Revízió", "1. kiadás dátuma", "Felülvizsgálat dátuma:", "Készítő / Felülvizsgáló:", "Regisztrációs szám"]
        vals_f = ["BSHM / B-ME", "3.", "2018.07.16", "2019.04.03.", "Janóczki M", "BME001-F-BDE-3"]
        widths_f = [Inches(1.8), Inches(1.2), Inches(1.6), Inches(1.8), Inches(2.5), Inches(1.8)]
        for c_i in range(6):
            c1 = foot_table.cell(0, c_i); c1.text = headers_f[c_i]
            set_cell_bg(c1, "F2F2F2"); set_cell_pad(c1, 20, 20, 30, 30); set_cell_border(c1, "4")
            c1.paragraphs[0].runs[0].font.name = 'Calibri'; c1.paragraphs[0].runs[0].font.size = Pt(7); c1.paragraphs[0].runs[0].font.bold = True
            
            c2 = foot_table.cell(1, c_i); c2.text = vals_f[c_i]
            set_cell_bg(c2, "FFFFFF"); set_cell_pad(c2, 20, 20, 30, 30); set_cell_border(c2, "4")
            c2.paragraphs[0].runs[0].font.name = 'Calibri'; c2.paragraphs[0].runs[0].font.size = Pt(7.5)
        for r_f in foot_table.rows:
            for c_i, w in enumerate(widths_f): r_f.cells[c_i].width = w

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- GOMB ÉS GENERÁLÁS ---
if st.button("⚡ INTELLIGENS ELEMZÉS & DOKUMENTUM GENERÁLÁSA"):
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    steps = [
        "Műszakos jegyzet elemzése & SAP cikkszám validálása...",
        "Usage Gap Kritériumok automatikus leképezése (1-13)...",
        "5-Miért mélyanalízis ágak matematikai felépítése...",
        "B-ME szabványú lábléc és regisztráció hitelesítése (Janóczki M)...",
        "Kész Word dokumentum (.docx) összeállítása..."
    ]
    
    for i, s in enumerate(steps):
        status_text.markdown(f"**Folyamat:** `{s}`")
        progress_bar.progress((i + 1) * 20)
        time.sleep(0.35)
        
    docx_file = general_hivatalos_docx(gep_nev, hiba_rovid, datum_val, felelos_val, jegyzet_szoveg)
    status_text.markdown("✅ **Az elemzés sikeresen lefutott, a dokumentum készen áll!**")
    
    st.download_button(
        label="📥 HIVATALOS WORD DOKUMENTUM LETÖLTÉSE (.DOCX)",
        data=docx_file,
        file_name=f"{gep_nev.replace(' ', '_')}_IDA_kitoltve.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )