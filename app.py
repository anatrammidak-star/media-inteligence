import io
import re
import urllib.parse
import feedparser
import pandas as pd
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import streamlit as st

# ==========================================
# 1. KONFIGURASI APLIKASI
# ==========================================
st.set_page_config(
    page_title="Media Intelligence & Sentiment Monitor",
    page_icon="📰",
    layout="wide"
)

st.title("📰 Media Intelligence & Sentiment Monitoring System")
st.caption("Platform Pemantauan Isu, Analisis Sentimen, dan Rekomendasi Komunikasi Strategis")

# ==========================================
# 2. ENGINE PEMINDAI BERITA (DENGAN FILTER WAKTU)
# ==========================================
def fetch_news(entity_name, issue_keywords, time_window="2d", max_results=10):
    """
    Mengambil berita via Google News RSS dengan batasan waktu spesifik.
    time_window: '1d' (24 jam), '2d' (2 hari), '7d' (1 minggu), '30d' (1 bulan), 'all' (tanpa batas).
    """
    base_query = f"{entity_name} {issue_keywords}".strip()
    
    # Menambahkan operator batasan waktu bawaan Google
    if time_window != "all":
        query = f"{base_query} when:{time_window}"
    else:
        query = base_query

    encoded_query = urllib.parse.quote(query)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=id&gl=ID&ceid=ID:id"
    
    feed = feedparser.parse(rss_url)
    articles = []

    # Korpus kata kunci sentimen reputasi
    negative_lexicon = [
        "korupsi", "rugi", "bobol", "fraud", "kasus", "gagal", "demo", "kecewa",
        "macet", "krisis", "tuntutan", "phk", "ancam", "turun", "pidana", "gugat", 
        "disorot", "hukum", "sita", "penipuan", "keluhan", "masalah"
    ]
    positive_lexicon = [
        "laba", "untung", "tumbuh", "prestasi", "penghargaan", "inovasi", "sukses",
        "dividen", "ekspansi", "naik", "resmikan", "unggul", "mitra", "positif", 
        "rekor", "terbaik", "kinerja positif", "apresiasi"
    ]

    for entry in feed.entries[:max_results]:
        title = entry.title
        link = entry.link
        published = entry.get("published", "Waktu publikasi tidak tercantum")
        
        # Pisahkan judul dan nama publisher media
        source = "Media Nasional"
        clean_title = title
        if " - " in title:
            parts = title.rsplit(" - ", 1)
            clean_title = parts[0]
            source = parts[1]

        # Bersihkan format ringkasan HTML
        raw_summary = entry.get("summary", "")
        soup = BeautifulSoup(raw_summary, "html.parser")
        snippet = soup.get_text()

        # Skoring Sentimen Sederhana
        title_lower = clean_title.lower()
        neg_count = sum(1 for word in negative_lexicon if word in title_lower)
        pos_count = sum(1 for word in positive_lexicon if word in title_lower)

        if neg_count > pos_count:
            sentiment = "Negatif"
            pr_tone = "High Risk / Potential Crisis"
            pr_rec = (
                "Rekomendasi Konsultan Media:\n"
                "1. Segera rilis 'Holding Statement' dan aktifkan crisis room monitoring.\n"
                "2. Berikan fakta berimbang kepada jurnalis peliput sebelum isu tereskalasi.\n"
                "3. Jangan bersikap reaktif di media sosial; fokus pada kanal komunikasi resmi."
            )
        elif pos_count > neg_count:
            sentiment = "Positif"
            pr_tone = "Strategic Advantage / Good Traction"
            pr_rec = (
                "Rekomendasi Konsultan Media:\n"
                "1. Amplifikasi narasi melalui media berbayar/terbina (owned & earned media).\n"
                "2. Manfaatkan momentum positif untuk 'Executive Thought Leadership' pimpinan.\n"
                "3. Dokumentasikan pencapaian sebagai bahan annual report / investor updates."
            )
        else:
            sentiment = "Netral"
            pr_tone = "Informative / Passive Coverage"
            pr_rec = (
                "Rekomendasi Konsultan Media:\n"
                "1. Pantau potensi pergeseran sentimen pada pemberitaan lanjutan.\n"
                "2. Siapkan data lembar fakta (factsheet) jika media memerlukan konfirmasi."
            )

        articles.append({
            "title": clean_title,
            "source": source,
            "published": published,
            "link": link,
            "snippet": snippet,
            "sentiment": sentiment,
            "pr_tone": pr_tone,
            "pr_rec": pr_rec
        })

    return articles

# ==========================================
# 3. GENERATOR PRESENTASI POWERPOINT (.PPTX)
# ==========================================
def generate_pptx_deck(entity_name, org_type, time_label, articles):
    prs = Presentation()
    prs.slide_width = Inches(13.33)  # Format 16:9 Widescreen
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: COVER ---
    cover_slide = prs.slides.add_slide(blank_layout)
    title_box = cover_slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.33), Inches(3.5))
    tf = title_box.text_frame
    tf.word_wrap = True
    
    p0 = tf.paragraphs[0]
    p0.text = "MEDIA INTELLIGENCE & SENTIMENT AUDIT"
    p0.font.size = Pt(36)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(20, 35, 60)

    p1 = tf.add_paragraph()
    p1.text = f"Target Entitas: {entity_name} | Sektor: {org_type}"
    p1.font.size = Pt(22)
    p1.font.color.rgb = RGBColor(70, 80, 95)

    p2 = tf.add_paragraph()
    p2.text = f"Rentang Pemantauan: {time_label} | Sumber Data: Google News RSS"
    p2.font.size = Pt(14)
    p2.font.color.rgb = RGBColor(100, 110, 125)

    p3 = tf.add_paragraph()
    p3.text = "Disusun oleh: Senior Strategic Media & Reputation Advisory"
    p3.font.size = Pt(13)
    p3.font.italic = True
    p3.font.color.rgb = RGBColor(140, 150, 160)

    # --- SLIDE 2+: SPLIT 2-KOLOM (KIRI: KLIPING, KANAN: REKOMENDASI) ---
    for idx, art in enumerate(articles, 1):
        slide = prs.slides.add_slide(blank_layout)

        # Header Slide
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.7))
        htf = header_box.text_frame
        hp = htf.paragraphs[0]
        hp.text = f"ANALISIS ISU MEDIA #{idx}: {art['source'].upper()}"
        hp.font.size = Pt(18)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(30, 41, 59)

        # SISI KIRI: CAPTURE / KLIPING PEMBERITAAN
        left_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(5.6), Inches(5.6))
        ltf = left_box.text_frame
        ltf.word_wrap = True

        lp1 = ltf.paragraphs[0]
        lp1.text = "DOKUMENTASI PEMBERITAAN"
        lp1.font.size = Pt(12)
        lp1.font.bold = True
        lp1.font.color.rgb = RGBColor(100, 116, 139)

        lp2 = ltf.add_paragraph()
        lp2.text = f"\n\"{art['title']}\""
        lp2.font.size = Pt(17)
        lp2.font.bold = True
        lp2.font.color.rgb = RGBColor(15, 23, 42)

        lp3 = ltf.add_paragraph()
        lp3.text = f"\nKanal/Penerbit: {art['source']}\nWaktu Terbit: {art['published']}\n"
        lp3.font.size = Pt(11)
        lp3.font.color.rgb = RGBColor(71, 85, 105)

        lp4 = ltf.add_paragraph()
        clean_snippet = re.sub('<[^<]+?>', '', art['snippet'])
        lp4.text = f"Cuplikan Teks:\n{clean_snippet[:260]}..."
        lp4.font.size = Pt(11)
        lp4.font.color.rgb = RGBColor(51, 65, 85)

        # SISI KANAN: REVIEW KONSULTAN & STRATEGIC RECOMMENDATIONS
        right_box = slide.shapes.add_textbox(Inches(6.8), Inches(1.3), Inches(5.7), Inches(5.6))
        rtf = right_box.text_frame
        rtf.word_wrap = True

        rp1 = rtf.paragraphs[0]
        rp1.text = "STRATEGIC MEDIA REVIEW & ADVISORY"
        rp1.font.size = Pt(12)
        rp1.font.bold = True
        rp1.font.color.rgb = RGBColor(14, 116, 144)

        rp2 = rtf.add_paragraph()
        rp2.text = f"\nKlasifikasi Sentimen: {art['sentiment'].upper()}"
        rp2.font.size = Pt(14)
        rp2.font.bold = True
        if art['sentiment'] == "Negatif":
            rp2.font.color.rgb = RGBColor(220, 38, 38)
        elif art['sentiment'] == "Positif":
            rp2.font.color.rgb = RGBColor(22, 163, 74)
        else:
            rp2.font.color.rgb = RGBColor(202, 138, 4)

        rp3 = rtf.add_paragraph()
        rp3.text = f"Tone Dampak: {art['pr_tone']}\n"
        rp3.font.size = Pt(11)
        rp3.font.italic = True
        rp3.font.color.rgb = RGBColor(100, 116, 139)

        rp4 = rtf.add_paragraph()
        rp4.text = f"{art['pr_rec']}"
        rp4.font.size = Pt(12)
        rp4.font.color.rgb = RGBColor(15, 23, 42)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# 4. TAMPILAN SIDEBAR & KONTROL INPUT
# ==========================================
with st.sidebar:
    st.header("⚙️ Pengaturan Parameter")
    entity_input = st.text_input("Nama Entitas / Organisasi", value="Bank Mandiri")
    org_type_input = st.selectbox("Jenis Organisasi", [
        "Perbankan / Finansial", 
        "BUMN / Pemerintahan", 
        "Korporasi Swasta", 
        "Lembaga Nirlaba", 
        "Personal Branding"
    ])
    keywords_input = st.text_input("Kata Kunci / Topik Isu", value="layanan kinerja laba", help="Pisahkan dengan spasi")

    # Dropdown Pemilihan Batasan Waktu
    time_options = {
        "24 Jam Terakhir (Hari Ini)": "1d",
        "2 Hari Terakhir": "2d",
        "3 Hari Terakhir": "3d",
        "7 Hari Terakhir (1 Minggu)": "7d",
        "30 Hari Terakhir (1 Bulan)": "30d",
        "Semua Waktu (Tanpa Batasan)": "all"
    }
    
    selected_time_label = st.selectbox(
        "Rentang Waktu Pemberitaan",
        options=list(time_options.keys()),
        index=1  # Default ke '2 Hari Terakhir'
    )
    time_window_val = time_options[selected_time_label]

    max_news = st.slider("Jumlah Berita Dipantau", min_value=3, max_value=20, value=6)
    run_btn = st.button("Mulai Monitoring & Analisis", type="primary")

# ==========================================
# 5. EKSEKUSI & TAMPILAN HASIL
# ==========================================
if run_btn:
    with st.spinner("Sedang memindai pemberitaan media nasional dan menyusun telaah PR..."):
        st.session_state['news_data'] = fetch_news(
            entity_name=entity_input,
            issue_keywords=keywords_input,
            time_window=time_window_val,
            max_results=max_news
        )
        st.session_state['entity'] = entity_input
        st.session_state['org_type'] = org_type_input
        st.session_state['time_label'] = selected_time_label

if 'news_data' in st.session_state:
    articles = st.session_state.get('news_data', [])

    if not articles:
        st.warning("Tidak ditemukan pemberitaan pada rentang waktu tersebut. Coba perpanjang rentang waktu atau buat kata kunci lebih umum.")
    else:
        # Metrik Ringkasan Sentimen
        pos_count = sum(1 for a in articles if a['sentiment'] == "Positif")
        neg_count = sum(1 for a in articles if a['sentiment'] == "Negatif")
        neu_count = sum(1 for a in articles if a['sentiment'] == "Netral")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Berita Dipindai", len(articles))
        col2.metric("Sentimen Positif", pos_count)
        col3.metric("Sentimen Negatif (Mitigasi)", neg_count)
        col4.metric("Sentimen Netral", neu_count)

        st.divider()

        # Tombol Download File PPTX
        pptx_file = generate_pptx_deck(
            st.session_state['entity'], 
            st.session_state['org_type'], 
            st.session_state['time_label'],
            articles
        )
        st.download_button(
            label="📥 Unduh Laporan Slide Deck (.PPTX) Siap Paparan",
            data=pptx_file,
            file_name=f"PR_Audit_{st.session_state['entity'].replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        # Pratinjau 2 Kolom di Web
        st.subheader("📋 Dashboard Telaah Media & Rekomendasi Konsultan")
        for idx, art in enumerate(articles, 1):
            with st.container():
                c_left, c_right = st.columns([1, 1], gap="medium")
                
                with c_left:
                    st.markdown(f"### [#{idx}] {art['title']}")
                    st.caption(f"**Sumber:** {art['source']} | **Waktu:** {art['published']}")
                    clean_text = re.sub('<[^<]+?>', '', art['snippet'])
                    st.markdown(f"> *\"{clean_text[:250]}...\"*")
                    st.markdown(f"[Buka Berita Asli]({art['link']})")

                with c_right:
                    badge_color = "🟢" if art['sentiment'] == "Positif" else "🔴" if art['sentiment'] == "Negatif" else "🟡"
                    st.markdown(f"#### {badge_color} Sentimen: **{art['sentiment']}**")
                    st.write(f"**PR Tone:** {art['pr_tone']}")
                    st.info(art['pr_rec'])

                st.markdown("---")
