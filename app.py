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
from pptx.enum.text import PP_ALIGN
import streamlit as st

# ==========================================
# KONFIGURASI TAMPILAN & BRANDING
# ==========================================
st.set_page_config(
    page_title="Media Intelligence & Reputation Monitor",
    page_icon="📰",
    layout="wide"
)

st.title("📰 Media Intelligence & Sentiment Monitoring System")
st.caption("Platform Pemantauan Isu, Analisis Sentimen, dan Rekomendasi Komunikasi Strategis")

# ==========================================
# FUNGSI SCRAPING & ANALISIS BERITA
# ==========================================
def fetch_news(entity_name, issue_keywords, max_results=10):
    """
    Mengambil berita via Google News RSS secara dinamis tanpa biaya API
    """
    query = f"{entity_name} {issue_keywords}".strip()
    encoded_query = urllib.parse.quote(query)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=id&gl=ID&ceid=ID:id"
    
    feed = feedparser.parse(rss_url)
    articles = []

    negative_lexicon = [
        "korupsi", "rugi", "bobol", "fraud", "kasus", "gagal", "demo", "kecewa",
        "macet", "krisis", "tuntutan", "phk", "ancam", "turun", "pidana", "gugat", "disorot", "hukum"
    ]
    positive_lexicon = [
        "laba", "untung", "tumbuh", "prestasi", "penghargaan", "inovasi", "sukses",
        "dividen", "ekspansi", "naik", "resmikan", "unggul", "mitra", "positif", "rekor"
    ]

    for entry in feed.entries[:max_results]:
        title = entry.title
        link = entry.link
        published = entry.get("published", "Waktu tidak diketahui")
        
        # Ekstrak nama media dari title jika formatnya "Judul Berita - Nama Media"
        source = "Media Nasional"
        clean_title = title
        if " - " in title:
            parts = title.rsplit(" - ", 1)
            clean_title = parts[0]
            source = parts[1]

        # Ringkasan ringkas dari deskripsi HTML
        raw_summary = entry.get("summary", "")
        soup = BeautifulSoup(raw_summary, "html.parser")
        snippet = soup.get_text()

        # Deteksi Sentimen berbasis Rule-Based / Lexicon Engine
        title_lower = clean_title.lower()
        neg_count = sum(1 for word in negative_lexicon if word in title_lower)
        pos_count = sum(1 for word in positive_lexicon if word in title_lower)

        if neg_count > pos_count:
            sentiment = "Negatif"
            pr_tone = "High Risk / Potential Crisis"
            pr_rec = (
                "Rekomendasi Konsultan:\n"
                "1. Segera siapkan 'Holding Statement' dan FAQ internal sebelum isu meluas.\n"
                "2. Monitor kanal media sosial untuk mitigasi viralisasi narasi negatif.\n"
                "3. Jangan bersikap defensif secara terbuka; lakukan penelusuran fakta objektif."
            )
        elif pos_count > neg_count:
            sentiment = "Positif"
            pr_tone = "Strategic Advantage / Good Traction"
            pr_rec = (
                "Rekomendasi Konsultan:\n"
                "1. Amplifikasi narasi melalui owned media (LinkedIn, Press Release resmi, Instagram).\n"
                "2. Jadikan momentum untuk membangun 'Executive Thought Leadership' direksi.\n"
                "3. Kirimkan apresiasi media relations kepada jurnalis yang meliput."
            )
        else:
            sentiment = "Netral"
            pr_tone = "Informative / Passive Coverage"
            pr_rec = (
                "Rekomendasi Konsultan:\n"
                "1. Pantau arah perkembangan isu lanjutan (follow-up coverage).\n"
                "2. Siapkan sudut pandang data penunjang jika ada jurnalis yang mengonfirmasi."
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
# GENERATOR PRESENTASI POWERPOINT (PPTX)
# ==========================================
def generate_pptx_deck(entity_name, org_type, articles):
    """
    Membuat file PowerPoint dengan format:
    - Sisi Kiri: Capture kartu berita (Metadata, Judul, Sumber, Link)
    - Sisi Kanan: Executive Summary & Rekomendasi Konsultan Media
    """
    prs = Presentation()
    # Atur dimensi slide 16:9 widescreen (13.33 x 7.5 inches)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: COVER ---
    cover_slide = prs.slides.add_slide(blank_layout)
    
    # Title Box
    title_box = cover_slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.33), Inches(3.0))
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
    p2.text = "Disusun oleh: Senior Strategic Media & Reputation Advisory"
    p2.font.size = Pt(14)
    p2.font.italic = True
    p2.font.color.rgb = RGBColor(120, 130, 140)

    # --- SLIDE PER PEMBERITAAN (2-COLUMN LAYOUT) ---
    for idx, art in enumerate(articles, 1):
        slide = prs.slides.add_slide(blank_layout)

        # Header Slide
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.8))
        htf = header_box.text_frame
        hp = htf.paragraphs[0]
        hp.text = f"ANALISIS ISU MEDIA #{idx}: {art['source'].upper()}"
        hp.font.size = Pt(18)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(30, 41, 59)

        # SISI KIRI: CAPTURE / RINGKASAN PEMBERITAAN
        # Background container simulasi clipping
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
        lp2.font.size = Pt(18)
        lp2.font.bold = True
        lp2.font.color.rgb = RGBColor(15, 23, 42)

        lp3 = ltf.add_paragraph()
        lp3.text = f"\nKanal/Sumber: {art['source']}\nWaktu Publikasi: {art['published']}\n"
        lp3.font.size = Pt(11)
        lp3.font.color.rgb = RGBColor(71, 85, 105)

        lp4 = ltf.add_paragraph()
        clean_snippet = re.sub('<[^<]+?>', '', art['snippet'])
        lp4.text = f"Cuplikan Isi:\n{clean_snippet[:280]}..."
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

        # Warna status sentimen
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

    # Simpan ke byte buffer
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# INTERFACE UTAMA APLIKASI
# ==========================================
with st.sidebar:
    st.header("⚙️ Pengaturan Parameter")
    entity_input = st.text_input("Nama Entitas / Organisasi", value="Bank Mandiri", help="Contoh: Bank Marindi, Telkom, Kementerian Keuangan")
    org_type_input = st.selectbox("Jenis Organisasi", ["Perbankan / Finansial", "BUMN / Pemerintahan", "Korporasi Swasta", "Lembaga Nirlaba", "Personal Branding"])
    keywords_input = st.text_input("Kata Kunci / Topik Isu", value="layanan kinerja laba", help="Pisahkan dengan spasi, misalnya: fraud kredit digital")
    max_news = st.slider("Jumlah Berita Dipantau", min_value=3, max_value=20, value=6)
    
    run_btn = st.button("Mulai Monitoring & Analisis", type="primary")

if run_btn or 'news_data' in st.session_state:
    if run_btn:
        with st.spinner("Sedang memindai pemberitaan media nasional dan menyusun telaah PR..."):
            st.session_state['news_data'] = fetch_news(entity_input, keywords_input, max_news)
            st.session_state['entity'] = entity_input
            st.session_state['org_type'] = org_type_input

    articles = st.session_state.get('news_data', [])

    if not articles:
        st.warning("Tidak ditemukan pemberitaan dengan kata kunci tersebut. Coba gunakan kata kunci yang lebih umum.")
    else:
        # Ringkasan Metrik
        pos_count = sum(1 for a in articles if a['sentiment'] == "Positif")
        neg_count = sum(1 for a in articles if a['sentiment'] == "Negatif")
        neu_count = sum(1 for a in articles if a['sentiment'] == "Netral")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Berita Dianalisis", len(articles))
        col2.metric("Sentimen Positif", pos_count)
        col3.metric("Sentimen Negatif (Perlu Mitigasi)", neg_count)
        col4.metric("Sentimen Netral", neu_count)

        st.divider()

        # Tombol Download PowerPoint
        pptx_file = generate_pptx_deck(st.session_state['entity'], st.session_state['org_type'], articles)
        st.download_button(
            label="📥 Unduh Laporan Slide Deck (.PPTX) Siap Paparan",
            data=pptx_file,
            file_name=f"PR_Media_Audit_{st.session_state['entity'].replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        # Tampilan Daftar Berita (Preview 2 Kolom di Web)
        st.subheader("📋 Dashboard Telaah Media & Rekomendasi Konsultan")
        for idx, art in enumerate(articles, 1):
            with st.container():
                c_left, c_right = st.columns([1, 1], gap="medium")
                
                with c_left:
                    st.markdown(f"### [#{idx}] {art['title']}")
                    st.caption(f"**Sumber:** {art['source']} | **Waktu:** {art['published']}")
                    st.markdown(f"> *\"{re.sub('<[^<]+?>', '', art['snippet'])[:250]}...\"*")
                    st.markdown(f"[Buka Berita Asli]({art['link']})")

                with c_right:
                    badge_color = "🟢" if art['sentiment'] == "Positif" else "🔴" if art['sentiment'] == "Negatif" else "🟡"
                    st.markdown(f"#### {badge_color} Sentimen: **{art['sentiment']}**")
                    st.write(f"**PR Tone:** {art['pr_tone']}")
                    st.info(art['pr_rec'])

                st.markdown("---")
