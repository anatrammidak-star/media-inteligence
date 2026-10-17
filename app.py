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
# 2. HELPER: EKSTRAKSI SUMMARY & SCREENSHOT
# ==========================================
def extract_bullet_summary(text, max_bullets=3):
    """
    Menyaring teks artikel menjadi 2-3 butir kalimat inti ringkas.
    """
    cleaned = re.sub(r'<[^<]+?>', '', text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', cleaned)
    bullets = [s.strip() for s in sentences if len(s.strip()) > 25]
    if not bullets:
        bullets = ["Pemberitaan memuat perkembangan operasional dan kebijakan terkini terkait entitas."]
    return bullets[:max_bullets]

def fetch_article_image_bytes(link):
    """
    Mengambil gambar tangkapan layar/thumbnail berita.
    Menggunakan WordPress mShots API (bebas biaya & tanpa browser berat di RAM).
    """
    try:
        # Layanan screenshot web gratis berbasis URL
        screenshot_url = f"https://s0.wp.com/mshots/v1/{urllib.parse.quote(link)}?w=700"
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(screenshot_url, headers=headers, timeout=5)
        if resp.status_code == 200 and len(resp.content) > 1000:
            return resp.content
    except Exception:
        pass
    return None

# ==========================================
# 3. ENGINE PEMINDAI BERITA
# ==========================================
def fetch_news(entity_name, issue_keywords, time_window="2d", max_results=6):
    base_query = f"{entity_name} {issue_keywords}".strip()
    if time_window != "all":
        query = f"{base_query} when:{time_window}"
    else:
        query = base_query

    encoded_query = urllib.parse.quote(query)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=id&gl=ID&ceid=ID:id"
    
    feed = feedparser.parse(rss_url)
    articles = []

    negative_lexicon = [
        "korupsi", "rugi", "bobol", "fraud", "kasus", "gagal", "demo", "kecewa",
        "macet", "krisis", "tuntutan", "phk", "ancam", "turun", "pidana", "gugat", 
        "disorot", "hukum", "sita", "penipuan", "keluhan", "masalah", "denda"
    ]
    positive_lexicon = [
        "laba", "untung", "tumbuh", "prestasi", "penghargaan", "inovasi", "sukses",
        "dividen", "ekspansi", "naik", "resmikan", "unggul", "mitra", "positif", 
        "rekor", "terbaik", "kinerja positif", "apresiasi", "capai"
    ]

    for entry in feed.entries[:max_results]:
        title = entry.title
        link = entry.link
        published = entry.get("published", "Waktu publikasi tidak tercantum")
        
        source = "Media Nasional"
        clean_title = title
        if " - " in title:
            parts = title.rsplit(" - ", 1)
            clean_title = parts[0]
            source = parts[1]

        raw_summary = entry.get("summary", "")
        soup = BeautifulSoup(raw_summary, "html.parser")
        snippet = soup.get_text()

        # Ekstraksi 2-3 butir isi pemberitaan
        bullet_points = extract_bullet_summary(snippet)

        # Skoring Sentimen
        title_lower = clean_title.lower()
        neg_count = sum(1 for word in negative_lexicon if word in title_lower)
        pos_count = sum(1 for word in positive_lexicon if word in title_lower)

        if neg_count > pos_count:
            sentiment = "Negatif"
            pr_tone = "High Risk / Potential Crisis"
            pr_rec = (
                "Rekomendasi Konsultan Media:\n"
                "1. Segera siapkan 'Holding Statement' dan aktifkan crisis room monitoring.\n"
                "2. Berikan fakta berimbang kepada jurnalis peliput sebelum isu tereskalasi.\n"
                "3. Jangan bersikap defensif secara terbuka; perkuat narasi penanganan konkret."
            )
        elif pos_count > neg_count:
            sentiment = "Positif"
            pr_tone = "Strategic Advantage / Good Traction"
            pr_rec = (
                "Rekomendasi Konsultan Media:\n"
                "1. Amplifikasi narasi melalui media berbayar/terbina (owned & earned channels).\n"
                "2. Manfaatkan momentum positif untuk 'Executive Thought Leadership' direksi.\n"
                "3. Dokumentasikan materi sebagai aset reputasi korporasi."
            )
        else:
            sentiment = "Netral"
            pr_tone = "Informative / Passive Coverage"
            pr_rec = (
                "Rekomendasi Konsultan Media:\n"
                "1. Pantau arah eskalasi pada siklus berita berikutnya (follow-up coverage).\n"
                "2. Siapkan data penunjang jika ada jurnalis yang mengajukan konfirmasi."
            )

        # Ambil screenshot tangkapan layar berita
        img_bytes = fetch_article_image_bytes(link)

        articles.append({
            "title": clean_title,
            "source": source,
            "published": published,
            "link": link,
            "snippet": snippet,
            "bullets": bullet_points,
            "sentiment": sentiment,
            "pr_tone": pr_tone,
            "pr_rec": pr_rec,
            "image_bytes": img_bytes
        })

    return articles

# ==========================================
# 4. GENERATOR PRESENTASI POWERPOINT (.PPTX)
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

    # --- SLIDE 2+: TATA LETAK 2 KOLOM ---
    for idx, art in enumerate(articles, 1):
        slide = prs.slides.add_slide(blank_layout)

        # Header Bar Slide
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.6))
        htf = header_box.text_frame
        hp = htf.paragraphs[0]
        hp.text = f"ANALISIS ISU MEDIA #{idx}: {art['source'].upper()}"
        hp.font.size = Pt(18)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(30, 41, 59)

        # ==========================================
        # SISI KIRI: DOKUMEN PEMBERITAAN & SCREENSHOT
        # ==========================================
        # 1. Dokumen Header Berita
        left_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.1), Inches(5.6), Inches(2.2))
        ltf = left_box.text_frame
        ltf.word_wrap = True

        lp1 = ltf.paragraphs[0]
        lp1.text = "DOKUMENTASI PEMBERITAAN"
        lp1.font.size = Pt(11)
        lp1.font.bold = True
        lp1.font.color.rgb = RGBColor(100, 116, 139)

        lp2 = ltf.add_paragraph()
        lp2.text = f"\"{art['title']}\""
        lp2.font.size = Pt(14)
        lp2.font.bold = True
        lp2.font.color.rgb = RGBColor(15, 23, 42)

        lp3 = ltf.add_paragraph()
        lp3.text = f"Kanal: {art['source']} | Waktu: {art['published']}"
        lp3.font.size = Pt(10)
        lp3.font.color.rgb = RGBColor(71, 85, 105)

        # 2. Screenshot Berita (di bawah dokumen pemberitaan)
        if art.get("image_bytes"):
            try:
                img_stream = io.BytesIO(art["image_bytes"])
                slide.shapes.add_picture(
                    img_stream, 
                    Inches(0.8), Inches(3.4), 
                    width=Inches(5.6), height=Inches(3.5)
                )
            except Exception:
                pass
        else:
            # Placeholder jika screenshot gagal dimuat
            ph_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.5), Inches(5.6), Inches(3.0))
            ph_tf = ph_box.text_frame
            ph_p = ph_tf.paragraphs[0]
            ph_p.text = "[Screenshot tangkapan layar web tidak tersedia secara langsung]"
            ph_p.font.size = Pt(11)
            ph_p.font.italic = True
            ph_p.font.color.rgb = RGBColor(148, 163, 184)

        # ==========================================
        # SISI KANAN: SUMMARY (2-3 BUTIR) + STRATEGIC ADVISORY
        # ==========================================
        right_box = slide.shapes.add_textbox(Inches(6.8), Inches(1.1), Inches(5.7), Inches(5.9))
        rtf = right_box.text_frame
        rtf.word_wrap = True

        # Bagian A: Summary Pemberitaan (2-3 Butir) - Posisi DI ATAS Advisory
        rp1 = rtf.paragraphs[0]
        rp1.text = "RINGKASAN PEMBERITAAN (KEY TAKEAWAYS)"
        rp1.font.size = Pt(11)
        rp1.font.bold = True
        rp1.font.color.rgb = RGBColor(14, 116, 144)

        for b in art['bullets']:
            bp = rtf.add_paragraph()
            bp.text = f"• {b}"
            bp.font.size = Pt(11)
            bp.font.color.rgb = RGBColor(30, 41, 59)

        # Pembatas Antar Bagian
        div_p = rtf.add_paragraph()
        div_p.text = "―" * 28
        div_p.font.size = Pt(9)
        div_p.font.color.rgb = RGBColor(203, 213, 225)

        # Bagian B: Strategic Media Review & Advisory
        adv_title = rtf.add_paragraph()
        adv_title.text = "STRATEGIC MEDIA REVIEW & ADVISORY"
        adv_title.font.size = Pt(11)
        adv_title.font.bold = True
        adv_title.font.color.rgb = RGBColor(14, 116, 144)

        sent_p = rtf.add_paragraph()
        sent_p.text = f"Sentimen: {art['sentiment'].upper()} ({art['pr_tone']})"
        sent_p.font.size = Pt(12)
        sent_p.font.bold = True
        if art['sentiment'] == "Negatif":
            sent_p.font.color.rgb = RGBColor(220, 38, 38)
        elif art['sentiment'] == "Positif":
            sent_p.font.color.rgb = RGBColor(22, 163, 74)
        else:
            sent_p.font.color.rgb = RGBColor(202, 138, 4)

        rec_p = rtf.add_paragraph()
        rec_p.text = f"{art['pr_rec']}"
        rec_p.font.size = Pt(11)
        rec_p.font.color.rgb = RGBColor(15, 23, 42)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# 5. TAMPILAN SIDEBAR & KONTROL INPUT
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

    max_news = st.slider("Jumlah Berita Dipantau", min_value=3, max_value=12, value=6)
    run_btn = st.button("Mulai Monitoring & Analisis", type="primary")

# ==========================================
# 6. EKSEKUSI & TAMPILAN HASIL
# ==========================================
if run_btn:
    with st.spinner("Sedang memindai pemberitaan, mengambil capture layar, dan menyusun telaah konsultan..."):
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
        st.warning("Tidak ditemukan pemberitaan pada rentang waktu tersebut. Coba perpanjang rentang waktu atau gunakan kata kunci yang lebih umum.")
    else:
        # Metrik Ringkasan
        pos_count = sum(1 for a in articles if a['sentiment'] == "Positif")
        neg_count = sum(1 for a in articles if a['sentiment'] == "Negatif")
        neu_count = sum(1 for a in articles if a['sentiment'] == "Netral")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Berita Dipindai", len(articles))
        col2.metric("Sentimen Positif", pos_count)
        col3.metric("Sentimen Negatif", neg_count)
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
            label="📥 Unduh Laporan Slide Deck (.PPTX) Lengkap Screenshot & Summary",
            data=pptx_file,
            file_name=f"PR_Audit_{st.session_state['entity'].replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        # Pratinjau di Layar Web
        st.subheader("📋 Dashboard Telaah Media & Rekomendasi Konsultan")
        for idx, art in enumerate(articles, 1):
            with st.container():
                c_left, c_right = st.columns([1, 1], gap="medium")
                
                with c_left:
                    st.markdown(f"### [#{idx}] {art['title']}")
                    st.caption(f"**Sumber:** {art['source']} | **Waktu:** {art['published']}")
                    
                    # Tampilan Screenshot di Web
                    if art.get("image_bytes"):
                        st.image(art["image_bytes"], caption=f"Capture Layar Pemberitaan #{idx}", use_container_width=True)
                    st.markdown(f"[🔗 Baca Berita Asli]({art['link']})")

                with c_right:
                    # 1. Summary Berita (2-3 Butir) di ATAS Strategic Advisory
                    st.markdown("#### 📝 Ringkasan Isi Pemberitaan:")
                    for b in art['bullets']:
                        st.markdown(f"- {b}")
                    
                    st.markdown("---")
                    
                    # 2. Strategic Advisory
                    badge_color = "🟢" if art['sentiment'] == "Positif" else "🔴" if art['sentiment'] == "Negatif" else "🟡"
                    st.markdown(f"#### {badge_color} Sentimen: **{art['sentiment']}** ({art['pr_tone']})")
                    st.info(art['pr_rec'])

                st.markdown("---")
