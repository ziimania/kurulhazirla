import streamlit as st
import docx
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy
import re
import io
import datetime

st.set_page_config(page_title="İSG Kurul Tutanağı Otomasyonu", page_icon="📋", layout="wide")

AYLAR_TR = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan",
    5: "Mayıs", 6: "Haziran", 7: "Temmuz", 8: "Ağustos",
    9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
}

def get_tc_text(tc, parent):
    lines = []
    for p in tc.p_lst:
        txt = Paragraph(p, parent).text.strip()
        if txt:
            lines.append(txt)
    return "\n".join(lines)

def set_tc_text_preserve_style(tc, text, bold=None):
    p_lst = tc.p_lst
    if not p_lst:
        p = OxmlElement('w:p')
        tc.append(p)
        p_lst = tc.p_lst
    first_p = p_lst[0]
    for extra_p in p_lst[1:]:
        tc.remove(extra_p)

    existing_rPr = None
    for r in first_p.findall(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is not None:
            existing_rPr = copy.deepcopy(rPr)
            break
    if existing_rPr is None:
        pPr = first_p.find(qn('w:pPr'))
        if pPr is not None:
            rPr = pPr.find(qn('w:rPr'))
            if rPr is not None:
                existing_rPr = copy.deepcopy(rPr)

    for child in list(first_p):
        if child.tag != qn('w:pPr'):
            first_p.remove(child)

    new_r = OxmlElement('w:r')
    if existing_rPr is not None:
        if bold is True and existing_rPr.find(qn('w:b')) is None:
            existing_rPr.append(OxmlElement('w:b'))
        elif bold is False:
            b_el = existing_rPr.find(qn('w:b'))
            if b_el is not None:
                existing_rPr.remove(b_el)
        new_r.append(existing_rPr)
    t_el = OxmlElement('w:t')
    t_el.set(qn('xml:space'), 'preserve')
    t_el.text = str(text)
    new_r.append(t_el)
    first_p.append(new_r)

def set_tc_multiline_decision(tc, lines, template_p_intro, template_p_bullet):
    for p in list(tc.p_lst):
        tc.remove(p)

    clean_lines = [l.strip() for l in lines if l.strip()]
    if not clean_lines:
        clean_lines = [""]

    for idx, line in enumerate(clean_lines):
        if idx == 0 or template_p_bullet is None:
            new_p = copy.deepcopy(template_p_intro)
        else:
            new_p = copy.deepcopy(template_p_bullet)

        existing_rPr = None
        for r in new_p.findall(qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            if rPr is not None:
                existing_rPr = copy.deepcopy(rPr)
                break
        for child in list(new_p):
            if child.tag != qn('w:pPr'):
                new_p.remove(child)

        new_r = OxmlElement('w:r')
        if existing_rPr is not None:
            new_r.append(existing_rPr)
        t_el = OxmlElement('w:t')
        t_el.set(qn('xml:space'), 'preserve')
        t_el.text = str(line)
        new_r.append(t_el)
        new_p.append(new_r)
        tc.append(new_p)

def parse_kurul_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    t0 = doc.tables[0]
    eski_no = get_tc_text(t0._tbl.tr_lst[0].tc_lst[1], t0).strip()
    eski_tarih_str = get_tc_text(t0._tbl.tr_lst[0].tc_lst[3], t0).strip()

    default_next_date = datetime.date.today()
    try:
        parts = eski_tarih_str.split(".")
        if len(parts) == 3:
            d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
            if m == 12:
                next_m, next_y = 1, y + 1
            else:
                next_m, next_y = m + 1, y
            default_next_date = datetime.date(next_y, next_m, min(d, 28))
    except Exception:
        pass

    t3 = doc.tables[3]
    trs = t3._tbl.tr_lst
    num_blocks = len(trs) // 7
    blocks = []
    for b in range(num_blocks):
        base = b * 7
        sira = get_tc_text(trs[base + 0].tc_lst[0], t3)
        karar_no = get_tc_text(trs[base + 0].tc_lst[2], t3)
        konu_raw = get_tc_text(trs[base + 2].tc_lst[1], t3)
        konu_clean = re.sub(r'^Gündem\s*\d+\.\s*', '', konu_raw, flags=re.IGNORECASE).strip()
        gerekce = get_tc_text(trs[base + 2].tc_lst[2], t3)
        karar_text = get_tc_text(trs[base + 4].tc_lst[1], t3)
        durum = get_tc_text(trs[base + 6].tc_lst[1], t3)
        termin = get_tc_text(trs[base + 6].tc_lst[2], t3)
        sorumlu = get_tc_text(trs[base + 6].tc_lst[3], t3)

        has_bullets = False
        p_lst = trs[base + 4].tc_lst[1].p_lst
        if len(p_lst) > 1:
            for p_el in p_lst[1:]:
                if p_el.find('.//' + qn('w:numPr')) is not None:
                    has_bullets = True
                    break

        blocks.append({
            "sira": sira,
            "karar_no": karar_no,
            "konu": konu_clean,
            "gerekce": gerekce,
            "karar": karar_text,
            "durum": durum,
            "termin": termin,
            "sorumlu": sorumlu,
            "use_bullets": has_bullets
        })

    # Karar 2 dışındaki, sondaki 2 kapanış maddesine kadar olan tüm kararlar devredenlerdir
    devredenler = blocks[1:-2] if len(blocks) > 3 else []
    if not devredenler and len(blocks) >= 1:
        devredenler = [blocks[0]]

    return {
        "eski_no": eski_no,
        "eski_tarih_str": eski_tarih_str,
        "default_next_date": default_next_date,
        "devredenler": devredenler
    }

def build_updated_kurul_docx(file_bytes, yeni_donem, yeni_tarih_str, devreden_list, yeni_kararlar_list):
    doc = docx.Document(io.BytesIO(file_bytes))

    # 1. TABLO 0: Toplantı No ve Tarih
    t0 = doc.tables[0]
    set_tc_text_preserve_style(t0._tbl.tr_lst[0].tc_lst[1], f" {yeni_donem}", bold=True)
    set_tc_text_preserve_style(t0._tbl.tr_lst[0].tc_lst[3], f" {yeni_tarih_str}", bold=True)

    # 2. TABLO 1: B. GÜNDEM MADDELERİ
    t1 = doc.tables[1]
    t1_trs = list(t1._tbl.tr_lst)
    agenda_row_template = copy.deepcopy(t1_trs[3])
    closing_agenda_1 = copy.deepcopy(t1_trs[-2])
    closing_agenda_2 = copy.deepcopy(t1_trs[-1])

    for tr in t1_trs[4:]:
        t1._tbl.remove(tr)

    # Devreden kararlar B. Gündem tablosunda tekrar listelenmez, sadece yeni eklenen kararlar girer
    for yk in yeni_kararlar_list:
        new_tr = copy.deepcopy(agenda_row_template)
        set_tc_text_preserve_style(new_tr.tc_lst[1], yk["konu"].strip(), bold=False)
        t1._tbl.append(new_tr)

    t1._tbl.append(closing_agenda_1)
    t1._tbl.append(closing_agenda_2)

    # 3. TABLO 2: Karar 1 (Açılış ve Yoklama)
    t2 = doc.tables[2]
    set_tc_text_preserve_style(t2._tbl.tr_lst[0].tc_lst[2], f"{yeni_donem}/001", bold=True)

    # 4. TABLO 3: Kararlar
    t3 = doc.tables[3]
    t3_trs = list(t3._tbl.tr_lst)

    template_block_rows = [copy.deepcopy(r) for r in t3_trs[7:14]] if len(t3_trs) >= 14 else [copy.deepcopy(r) for r in t3_trs[0:7]]
    decision_cell_paras = template_block_rows[4].tc_lst[1].p_lst
    p_intro_template = copy.deepcopy(decision_cell_paras[0])
    p_bullet_template = copy.deepcopy(decision_cell_paras[1]) if len(decision_cell_paras) > 1 else None

    closing_block_1_rows = [copy.deepcopy(r) for r in t3_trs[-14:-7]]
    closing_block_2_rows = [copy.deepcopy(r) for r in t3_trs[-7:]]

    for tr in t3_trs:
        t3._tbl.remove(tr)

    def append_decision_block(sira_val, karar_no_str, gundem_konu_str, gerekce_str, karar_lines, durum_str, termin_str, sorumlu_str, use_bullets=False):
        rows = [copy.deepcopy(r) for r in template_block_rows]
        set_tc_text_preserve_style(rows[0].tc_lst[0], str(sira_val), bold=True)
        set_tc_text_preserve_style(rows[0].tc_lst[2], karar_no_str, bold=True)
        set_tc_text_preserve_style(rows[2].tc_lst[1], gundem_konu_str, bold=True)
        set_tc_text_preserve_style(rows[2].tc_lst[2], gerekce_str, bold=False)
        set_tc_multiline_decision(
            rows[4].tc_lst[1],
            karar_lines,
            p_intro_template,
            p_bullet_template if use_bullets else None
        )
        set_tc_text_preserve_style(rows[6].tc_lst[1], durum_str, bold=False)
        set_tc_text_preserve_style(rows[6].tc_lst[2], termin_str, bold=False)
        set_tc_text_preserve_style(rows[6].tc_lst[3], sorumlu_str, bold=False)
        for r in rows:
            t3._tbl.append(r)

    # GEÇEN AYDAN DEVREDEN KARARLAR (HER BİRİ AYRI 7 SATIRLIK BLOK, SIRA=2, NO=YYYY-AA/002, KONU=Gündem 2. [Konu])
    for dk in devreden_list:
        konu_clean = re.sub(r'^Gündem\s*\d+\.\s*', '', dk["konu"].strip(), flags=re.IGNORECASE)
        gundem_konu_str = f"Gündem 2. {konu_clean}"
        lines = dk["karar"].split("\n")
        append_decision_block(
            sira_val=2,
            karar_no_str=f"{yeni_donem}/002",
            gundem_konu_str=gundem_konu_str,
            gerekce_str=dk["gerekce"],
            karar_lines=lines,
            durum_str=dk["durum"],
            termin_str=dk["termin"],
            sorumlu_str=dk["sorumlu"],
            use_bullets=dk.get("use_bullets", len(lines) > 1)
        )

    current_sira = 3

    # YENİ EKLENEN KARARLAR (SIRA 3'TEN BAŞLAR: Gündem 3., Gündem 4...)
    for yk in yeni_kararlar_list:
        konu_clean = re.sub(r'^Gündem\s*\d+\.\s*', '', yk["konu"].strip(), flags=re.IGNORECASE)
        gundem_konu_str = f"Gündem {current_sira}. {konu_clean}"
        lines = yk["karar"].split("\n")
        append_decision_block(
            sira_val=current_sira,
            karar_no_str=f"{yeni_donem}/{current_sira:03d}",
            gundem_konu_str=gundem_konu_str,
            gerekce_str=yk["gerekce"],
            karar_lines=lines,
            durum_str=yk["durum"],
            termin_str=yk["termin"],
            sorumlu_str=yk["sorumlu"],
            use_bullets=yk.get("use_bullets", len(lines) > 1)
        )
        current_sira += 1

    # Kapanış 1: Gündeme alınması talep edilen konuların belirlenmesi,
    set_tc_text_preserve_style(closing_block_1_rows[0].tc_lst[0], str(current_sira), bold=True)
    set_tc_text_preserve_style(closing_block_1_rows[0].tc_lst[2], f"{yeni_donem}/{current_sira:03d}", bold=True)
    set_tc_text_preserve_style(closing_block_1_rows[2].tc_lst[1], f"Gündem {current_sira}. Gündeme alınması talep edilen konuların belirlenmesi,", bold=True)
    for r in closing_block_1_rows:
        t3._tbl.append(r)
    current_sira += 1

    # Kapanış 2: Dilek ve Temenniler
    set_tc_text_preserve_style(closing_block_2_rows[0].tc_lst[0], f"   {current_sira}", bold=True)
    set_tc_text_preserve_style(closing_block_2_rows[0].tc_lst[2], f"{yeni_donem}/{current_sira:03d}", bold=True)
    set_tc_text_preserve_style(closing_block_2_rows[2].tc_lst[1], f"Gündem {current_sira}. Dilek ve Temenniler", bold=True)
    for r in closing_block_2_rows:
        t3._tbl.append(r)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ================= STREAMLIT ARAYÜZÜ =================
st.title("📋 İSG Kurul Tutanağı Otomasyonu")
st.markdown("Önceki ayın `.docx` tutanağını yükleyin, **Yeni Kurul Tarihini** seçin ve yeni kararları ekleyin.")

uploaded_file = st.file_uploader("1️⃣ Önceki Ayın İSG Kurul Tutanağını Yükleyin (.docx)", type=["docx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    if st.session_state.get("last_uploaded_name") != uploaded_file.name:
        parsed_data = parse_kurul_docx(file_bytes)
        st.session_state.last_uploaded_name = uploaded_file.name
        st.session_state.parsed_data = parsed_data
        st.session_state.devreden_kararlar = copy.deepcopy(parsed_data["devredenler"])
        st.session_state.yeni_kararlar = []

    parsed = st.session_state.parsed_data
    st.success(f"✅ Yüklendi: **{uploaded_file.name}** | Önceki Toplantı No: `{parsed['eski_no']}`")

    st.markdown("---")
    st.subheader("2️⃣ Yeni Kurul Tarihi ve Dönem Bilgisi")
    col_d1, col_d2, col_d3 = st.columns(3)

    secilen_tarih = col_d1.date_input(
        "Yeni Kurul Toplantı Tarihi Seçin",
        value=parsed["default_next_date"],
        format="DD.MM.YYYY"
    )
    otomatik_donem = f"{secilen_tarih.year}-{secilen_tarih.month:02d}"
    yeni_donem = col_d2.text_input("Yeni Toplantı No (Otomatik)", value=otomatik_donem)
    yeni_tarih_str = col_d3.text_input("Belgede Görünecek Tarih", value=secilen_tarih.strftime("%d.%m.%Y"))
    ay_ismi = AYLAR_TR.get(secilen_tarih.month, "")

    st.markdown("---")
    st.subheader("3️⃣ Önceki Kuruldan Devreden Kararlar (Gündem 2 Altında Ayrı Ayrı Gösterilecek)")
    st.caption("Aşağıdaki kararların her biri tutanakta ayrı ayrı Karar 2 tablosu olarak yer alacaktır. Lütfen mevcut durumlarını güncelleyin.")

    degismeyen_var = False
    updated_devreden = []

    for idx, dk in enumerate(st.session_state.devreden_kararlar):
        with st.expander(f"📌 Devreden Karar: {dk['konu']}", expanded=True):
            d_konu = st.text_input("Konu", value=dk["konu"], key=f"dk_konu_{idx}")
            d_gerekce = st.text_area("Gerekçe", value=dk["gerekce"], height=70, key=f"dk_ger_{idx}")
            d_karar = st.text_area("Alınan Karar / Uygulama", value=dk["karar"], height=120, key=f"dk_kar_{idx}")
            
            # KIRMIZI ÇERÇEVELİ VE DİKKAT ÇEKİCİ MEVCUT DURUM ALANI
            st.markdown("""
            <div style="border: 2px solid #dc2626; background-color: #fef2f2; border-radius: 8px; padding: 10px; margin-top: 10px; margin-bottom: 5px;">
                <b style="color: #b91c1c;">⚠️ DİKKAT: Mevcut Durumu Güncelleyin!</b>
            </div>
            """, unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns(3)
            default_durum_val = "Kurul Kararı Daireye İletildi" if dk["durum"] == "Uygulama sonucu bekleniyor." else dk["durum"]
            d_durum = c1.text_input("Mevcut Durum", value=default_durum_val, key=f"dk_dur_{idx}")
            
            if d_durum.strip() == "Uygulama sonucu bekleniyor.":
                degismeyen_var = True

            d_termin = c2.text_input("Termin", value=dk["termin"], key=f"dk_ter_{idx}")
            d_sorumlu = c3.text_input("Sorumlu Kişi", value=dk["sorumlu"], key=f"dk_sor_{idx}")

            updated_devreden.append({
                "konu": d_konu,
                "gerekce": d_gerekce,
                "karar": d_karar,
                "durum": d_durum,
                "termin": d_termin,
                "sorumlu": d_sorumlu
            })

    if degismeyen_var:
        st.error("🚨 **Mevcut durumu değiştirmeyecek misin?** Geçen ayki kararın durumu hâlâ 'Uygulama sonucu bekleniyor.' olarak kalmış görünüyor! Lütfen kontrol edin.")

    st.markdown("---")
    st.subheader("4️⃣ Yeni Karar Ekle (Gündem 3 ve Sonrasından Başlayacak)")

    with st.form("yeni_karar_ekle_form", clear_on_submit=True):
        yk_konu = st.text_input("Yeni Karar Konusu")
        yk_gerekce = st.text_area("Gerekçe", height=70)
        yk_karar = st.text_area("Alınan Karar / Uygulama", height=120)
        fc1, fc2, fc3 = st.columns(3)
        yk_durum = fc1.text_input("Mevcut Durum", value="Uygulama sonucu bekleniyor.")
        yk_termin = fc2.text_input("Termin", value="İvedilikle")
        yk_sorumlu = fc3.text_input("Sorumlu Kişi", value="Üst İşveren")

        submitted = st.form_submit_button("➕ Yeni Kararı Listeye Ekle")
        if submitted:
            if yk_konu.strip():
                st.session_state.yeni_kararlar.append({
                    "konu": yk_konu.strip(),
                    "gerekce": yk_gerekce.strip(),
                    "karar": yk_karar.strip(),
                    "durum": yk_durum.strip(),
                    "termin": yk_termin.strip(),
                    "sorumlu": yk_sorumlu.strip()
                })
                st.rerun()

    updated_yeni_kararlar = []
    if st.session_state.yeni_kararlar:
        st.markdown("#### Eklenen Yeni Kararlar:")
        for y_idx, yk in enumerate(st.session_state.yeni_kararlar):
            sira_no = 3 + y_idx
            with st.expander(f"🆕 Gündem {sira_no} ({yeni_donem}/{sira_no:03d}): {yk['konu']}", expanded=True):
                e_konu = st.text_input("Konu", value=yk["konu"], key=f"yk_konu_{y_idx}")
                e_gerekce = st.text_area("Gerekçe", value=yk["gerekce"], height=70, key=f"yk_ger_{y_idx}")
                e_karar = st.text_area("Alınan Karar / Uygulama", value=yk["karar"], height=120, key=f"yk_kar_{y_idx}")
                ec1, ec2, ec3, ec4 = st.columns([2, 2, 2, 1])
                e_durum = ec1.text_input("Mevcut Durum", value=yk["durum"], key=f"yk_dur_{y_idx}")
                e_termin = ec2.text_input("Termin", value=yk["termin"], key=f"yk_ter_{y_idx}")
                e_sorumlu = ec3.text_input("Sorumlu Kişi", value=yk["sorumlu"], key=f"yk_sor_{y_idx}")
                if ec4.button("🗑️ Sil", key=f"yk_del_{y_idx}"):
                    st.session_state.yeni_kararlar.pop(y_idx)
                    st.rerun()

                updated_yeni_kararlar.append({
                    "konu": e_konu,
                    "gerekce": e_gerekce,
                    "karar": e_karar,
                    "durum": e_durum,
                    "termin": e_termin,
                    "sorumlu": e_sorumlu
                })

    st.markdown("---")
    output_docx = build_updated_kurul_docx(
        file_bytes,
        yeni_donem.strip(),
        yeni_tarih_str.strip(),
        updated_devreden,
        updated_yeni_kararlar
    )
    out_filename = f"Veterinerlik {yeni_donem.strip()} {ay_ismi} İSG Kurul Tutanağı.docx".replace("  ", " ")

    st.download_button(
        label=f"📥 Yeni Tutanağı Oluştur ve İndir ({out_filename})",
        data=output_docx,
        file_name=out_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )
