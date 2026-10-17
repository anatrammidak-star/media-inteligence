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
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="Media & Social Intelligence Monitor",
    page_icon="📡",
    layout="wide"
)

st.title("📡 Media & Social Intelligence Monitoring System")
st.caption("Pemantauan Komprehensif Berita Online & Media Sosial (Instagram, Facebook, Threads, X, TikTok) + Rekomendasi PR Eksekutif")

# ==========================================
# 2. GENERATOR KARTU CAPTURE VISUAL (ANTI GAGAL)
# ==========================================
def generate_press_capture_card(channel_type, source, title, published, snippet, sentiment):
    """
    Membuat kartu dokumentasi visual profesional (PNG) secara programatik.
    Menghilangkan risiko screenshot mutar-mutar/gagal di PowerPoint.
    """
    width, height = 700, 480
    bg_color = (255, 255, 255)
    
    # Warna sentimen
    if sentiment == "Negatif":
        accent_color = (220, 38, 38)     # Merah
        badge_bg = (254, 226, 226)
    elif sentiment == "Positif":
        accent_color = (22, 163, 74)     # Hijau
        badge_bg = (220, 252, 231)
    else:
        accent_color = (202, 138, 4)     # Kuning/Amber
        badge_bg = (254, 243, 199)

    img = Image.new("RGB", (width, height), bg_color)
    draw = ImageDraw.Draw(img)

    # Frame border luar
    draw.rectangle([(0, 0), (width - 1, height - 1)], outline=(203, 213, 225), width=2)
    
    # Header bar
    draw.rectangle([(0, 0), (width, 50)], fill=(241, 245, 249))
    draw.rectangle([(0, 0), (8, height)], fill=accent_color)  # Garis aksen kiri

    # Simbol channel
    type_tag = "[MEDIA SOSIAL]" if channel_type == "medsos" else "[MEDIA ONLINE]"
    draw.text((25, 16), f"{type_tag}  •  {source.upper()}", fill=(30, 41, 59))

    # Badge Sentimen di kanan atas
    draw.rectangle([(width - 140, 12), (width - 20, 38)], fill=badge_bg)
    draw.text((width - 125, 17), sentiment.upper(), fill=accent_color)

    # Waktu
    draw.text((25, 65), f"Waktu Rilis / Deteksi: {published}", fill=(100, 116, 139))

    # Helper text wrapper sederhana
    def wrap_text(text, limit=48):
        lines = []
        words = text.split()
        cur = []
        for w in words:
            if sum(len(x) + 1 for x in cur) + len(w) <= limit:
                cur.append(w)
            else:
                lines.append(" ".join(cur))
                cur = [w]
        if cur:
            lines.append(" ".join(cur))
        return lines

    # Judul Pemberitaan / Konten Medsos
    y_text = 95
    draw.text((25, y_text), "HEADLINE / POST TITLE:", fill=(71, 85, 105))
    y_text += 22

    for line in wrap_text(title, limit=40)[:3]:
        draw.text((25, y_text), line, fill=(15, 23, 42))
        y_text += 25

    # Garis pembatas isi
    y_text += 10
    draw.line([(25, y_text), (width - 25, y_text)], fill=(226, 232, 240), width=1)
    y_text += 15

    # Cuplikan Konten
    draw.text((25, y_text), "KUTIPAN / CAPTION / ISI KONTEN:", fill=(71, 85, 105))
    y_text += 22
    for line in wrap_text(snippet, limit=48)[:5]:
        draw.text((25, y_text), line, fill=(51, 65, 85))
        y_text += 22

    # Footer Watermark
    draw.rectangle([(0, height - 35), (width, height)], fill=(248, 250, 252))
    draw.text((25, height - 25), "DIVERIFIKASI MEDIA INTELLIGENCE ADVISORY SYSTEM", fill=(148, 163, 184))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()

# ==========================================
# 3. ENGINE PEMINDAI BERITA ONLINE & MEDSOS
# ==========================================
def extract_bullet_summary(text, max_bullets=3):
    cleaned = re.sub(r'<[^<]+?>', '', text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', cleaned)
    bullets = [s.strip() for s in sentences if len(s.strip()) > 20]
    if not bullets:
        bullets = ["Pemberitaan memuat dinamika perkembangan isu operasional entitas."]
    return bullets[:max_bullets]

def fetch_media_and_social(entity_name, issue_keywords, time_window, max_results, include_news, include_social):
    items = []
    
    # 1. Konstruksi Query Media Online
    queries = []
    if include_news:
        q_news = f"{entity_name} {issue_keywords}".strip()
        queries.append((q_news, "online_news"))

    if include_social:
        # Menargetkan postingan Instagram, Facebook, Threads, X/Twitter, TikTok
        social_targets = "(site:instagram.com OR site:facebook.com OR site:threads.net OR site:twitter.com OR site:x.com OR site:tiktok.com)"
        q_social = f"{entity_name} {issue_keywords} {social_targets}".strip()
        queries.append((q_social, "medsos"))

    negative_lexicon = [
        "korupsi", "rugi", "bobol", "fraud", "kasus", "gagal", "demo", "kecewa",
        "macet", "krisis", "tuntutan", "phk", "ancam", "turun", "pidana", "gugat", 
        "disorot", "hukum", "sita", "penipuan", "keluhan", "masalah", "denda", "skandal", "viral", "hujat"
    ]
    positive_lexicon = [
        "laba", "untung", "tumbuh", "prestasi", "penghargaan", "inovasi", "sukses",
        "dividen", "ekspansi", "naik", "resmikan", "unggul", "mitra", "positif", 
        "rekor", "terbaik", "kinerja positif", "apresiasi", "capai", "solutif", "bangga"
    ]

    for q, channel_type in queries:
        if time_window != "all":
            final_query = f"{q} when:{time_window}"
        else:
            final_query = q

        encoded_query = urllib.parse.quote(final_query)
        rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=id&gl=ID&ceid=ID:id"
        feed = feedparser.parse(rss_url)

        for entry in feed.entries:
            title = entry.title
            link = entry.link
            published = entry.get("published", "Terkini")

            # Identifikasi Media / Akun Medsos
            source = "Media Nasional"
            clean_title = title
            if " - " in title:
                parts = title.rsplit(" - ", 1)
                clean_title = parts[0]
                source = parts[1]

            # Deteksi platform medsos dari link/title
            if channel_type == "medsos" or any(s in link.lower() for s in ["instagram.com", "facebook.com", "threads.net", "twitter.com", "x.com", "tiktok.com"]):
                channel_type = "medsos"
                if "instagram.com" in link.lower() or "instagram" in source.lower():
                    source = "Instagram"
                elif "facebook.com" in link.lower() or "facebook" in source.lower():
                    source = "Facebook"
                elif "threads.net" in link.lower():
                    source = "Threads"
                elif "twitter.com" in link.lower() or "x.com" in link.lower():
                    source = "X (Twitter)"
                elif "tiktok.com" in link.lower():
                    source = "TikTok"
                else:
                    source = "Social Network"

            raw_summary = entry.get("summary", "")
            soup = BeautifulSoup(raw_summary, "html.parser")
            snippet = soup.get_text()

            bullet_points = extract_bullet_summary(snippet)

            # Analisis Sentimen
            title_lower = clean_title.lower()
            snippet_lower = snippet.lower()
            neg_count = sum(1 for w in negative_lexicon if w in title_lower or w in snippet_lower)
            pos_count = sum(1 for w in positive_lexicon if w in title_lower or w in snippet_lower)

            if neg_count > pos_count:
                sentiment = "Negatif"
                pr_tone = "High Risk / Potential Crisis"
                if channel_type == "medsos":
                    pr_rec = (
                        "Rekomendasi Respons Medsos:\n"
                        "1. Aktifkan tim Social Media Response (Customer Care / PR) untuk tanggapan santun di kolom komentar.\n"
                        "2. Hindari menghapus komentar secara sepihak agar tidak memicu backlash netizen.\n"
                        "3. Arahkan pelapor/komplain ke saluran resmi (Direct Message / Call Center resmi)."
                    )
                else:
                    pr_rec = (
                        "Rekomendasi Konsultan Media:\n"
                        "1. Siapkan 'Holding Statement' resmi dan sebarkan ke media tier-1.\n"
                        "2. Jalin koordinasi proaktif dengan wartawan terkait untuk hak jawab berimbang.\n"
                        "3. Evaluasi potensi eskalasi krisis ke level otoritas pengawas."
                    )
            elif pos_count > neg_count:
                sentiment = "Positif"
                pr_tone = "Strategic Advantage / Good Traction"
                if channel_type == "medsos":
                    pr_rec = (
                        "Rekomendasi Respons Medsos:\n"
                        "1. Repost / quote konten positif ke kanal Instagram Story atau Threads resmi entitas.\n"
                        "2. Bangun engagement positif dengan memberikan 'like' atau komentar apresiasi.\n"
                        "3. Gunakan materi ini sebagai Social Proof untuk audiens muda."
                    )
                else:
                    pr_rec = (
                        "Rekomendasi Konsultan Media:\n"
                        "1. Amplifikasi melalui siaran pers resmi dan jaringan asosiasi industri.\n"
                        "2. Jadikan rujukan dalam agenda public expose / media briefing pimpinan."
                    )
            else:
                sentiment = "Netral"
                pr_tone = "Informative / Passive Coverage"
                pr_rec = (
                    "Rekomendasi Konsultan Media:\n"
                    "1. Pantau interaksi lanjutan netizen & jurnalis untuk mendeteksi sentimen turunan.\n"
                    "2. Sediakan factsheet ringkas jika ada permohonan klarifikasi."
                )

            # Buat capture visual card anti-gagal
            card_bytes = generate_press_capture_card(
                channel_type, source, clean_title, published, snippet, sentiment
            )

            items.append({
                "channel_type": channel_type,
                "title": clean_title,
                "source": source,
                "published": published,
                "link": link,
                "snippet": snippet,
                "bullets": bullet_points,
                "sentiment": sentiment,
                "pr_tone": pr_tone,
                "pr_rec": pr_rec,
                "image_bytes": card_bytes
            })

            # Batasan jika jumlah bebas/spesifik tercapai
            if max_results != "unlimited" and len(items) >= int(max_results):
                break
        
        if max_results != "unlimited" and len(items) >= int(max_results):
            break

    return items

# ==========================================
# 4. GENERATOR POWERPOINT MULTI-CHANNEL
# ==========================================
def generate_pptx_deck(entity_name, org_type, time_label, articles):
    prs = Presentation()
    prs.slide_width = Inches(13.33)  # Format 16:9
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # COVER SLIDE
    cover_slide = prs.slides.add_slide(blank_layout)
    title_box = cover_slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.33), Inches(3.5))
    tf = title_box.text_frame
    tf.word_wrap = True
    
    p0 = tf.paragraphs[0]
    p0.text = "MEDIA & SOCIAL INTELLIGENCE AUDIT"
    p0.font.size = Pt(36)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(20, 35, 60)

    p1 = tf.add_paragraph()
    p1.text = f"Target Entitas: {entity_name} | Kategori: {org_type}"
    p1.font.size = Pt(22)
    p1.font.color.rgb = RGBColor(70, 80, 95)

    p2 = tf.add_paragraph()
    p2.text = f"Cakupan: Media Online & Jejaring Sosial | Rentang Waktu: {time_label}"
    p2.font.size = Pt(14)
    p2.font.color.rgb = RGBColor(100, 110, 125)

    # CONTENT SLIDES (SPLIT LAYOUT)
    for idx, art in enumerate(articles, 1):
        slide = prs.slides.add_slide(blank_layout)

        # Header Title
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.6))
        htf = header_box.text_frame
        hp = htf.paragraphs[0]
        channel_label = "MEDIA SOSIAL" if art['channel_type'] == "medsos" else "MEDIA ONLINE"
        hp.text = f"AUDIT #{idx} [{channel_label}] : {art['source'].upper()}"
        hp.font.size = Pt(18)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(30, 41, 59)

        # SISI KIRI: CAPTURE CARD VISUAL (EMBEDDED LOKAL - TIDAK MUTAR)
        if art.get("image_bytes"):
            img_stream = io.BytesIO(art["image_bytes"])
            slide.shapes.add_picture(
                img_stream, 
                Inches(0.8), Inches(1.3), 
                width=Inches(5.7), height=Inches(5.6)
            )

        # SISI KANAN: 2-3 BUTIR RINGKASAN + REKOMENDASI ADVISORY
        right_box = slide.shapes.add_textbox(Inches(6.8), Inches(1.3), Inches(5.7), Inches(5.6))
        rtf = right_box.text_frame
        rtf.word_wrap = True

        # Summary
        rp1 = rtf.paragraphs[0]
        rp1.text = "RINGKASAN PEMBERITAAN / POSTINGAN (KEY TAKEAWAYS)"
        rp1.font.size = Pt(11)
        rp1.font.bold = True
        rp1.font.color.rgb = RGBColor(14, 116, 144)

        for b in art['bullets']:
            bp = rtf.add_paragraph()
            bp.text = f"• {b}"
            bp.font.size = Pt(11)
            bp.font.color.rgb = RGBColor(30, 41, 59)

        div_p = rtf.add_paragraph()
        div_p.text = "―" * 28
        div_p.font.size = Pt(9)
        div_p.font.color.rgb = RGBColor(203, 213, 225)

        # Strategic Advisory
        adv_title = rtf.add_paragraph()
        adv_title.text = "STRATEGIC MEDIA & SOCIAL ADVISORY"
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
# 5. SIDEBAR & KONTROL INPUT FLEKSIBEL
# ==========================================
with st.sidebar:
    st.header("⚙️ Konfigurasi Monitoring")
    entity_input = st.text_input("Nama Entitas / Target", value="Bank Mandiri")
    org_type_input = st.selectbox("Kategori Entitas", [
        "Perbankan / Finansial", 
        "BUMN / Pemerintahan", 
        "Korporasi Swasta", 
        "Lembaga Publik / NGO", 
        "Personal / Tokoh Publik"
    ])
    keywords_input = st.text_input("Kata Kunci Isu Spesifik", value="layanan laba kartu kredit fraud")

    st.subheader("🌐 Saluran Sumber Data")
    col_ch1, col_ch2 = st.columns(2)
    with col_ch1:
        inc_news = st.checkbox("Media Berita Online", value=True)
    with col_ch2:
        inc_social = st.checkbox("Media Sosial", value=True, help="Instagram, Facebook, Threads, X/Twitter, TikTok")

    st.subheader("⏳ Rentang Waktu")
    time_options = {
        "24 Jam Terakhir (Hari Ini)": "1d",
        "2 Hari Terakhir": "2d",
        "3 Hari Terakhir": "3d",
        "7 Hari Terakhir (1 Minggu)": "7d",
        "30 Hari Terakhir (1 Bulan)": "30d",
        "Semua Waktu (Tanpa Batas)": "all"
    }
    selected_time_label = st.selectbox("Pilih Rentang", list(time_options.keys()), index=1)
    time_val = time_options[selected_time_label]

    st.subheader("📊 Batasan Jumlah Temuan")
    limit_mode = st.radio("Pilihan Kuota:", ["Tentukan Jumlah", "Bebas (Ambil Semua Berita)"])
    if limit_mode == "Tentukan Jumlah":
        max_news_count = st.number_input("Jumlah Temuan Diinginkan", min_value=1, max_value=100, value=15, step=5)
    else:
        max_news_count = "unlimited"
        st.info("Sistem akan menarik seluruh berita & postingan yang tersedia di feed pencarian.")

    run_btn = st.button("🚀 Mulai Pemindaian Lintas Kanal", type="primary")

# ==========================================
# 6. TAMPILAN DASHBOARD & EKSPOR
# ==========================================
if run_btn:
    if not inc_news and not inc_social:
        st.error("Pilih minimal satu saluran (Media Berita Online atau Media Sosial)!")
    else:
        with st.spinner("Memindai portal berita dan jejaring sosial, menghasilkan visual capture, serta menyusun telaah konsultan..."):
            st.session_state['news_data'] = fetch_media_and_social(
                entity_name=entity_input,
                issue_keywords=keywords_input,
                time_window=time_val,
                max_results=max_news_count,
                include_news=inc_news,
                include_social=inc_social
            )
            st.session_state['entity'] = entity_input
            st.session_state['org_type'] = org_type_input
            st.session_state['time_label'] = selected_time_label

if 'news_data' in st.session_state:
    articles = st.session_state.get('news_data', [])

    if not articles:
        st.warning("Tidak ditemukan pemberitaan atau konten medsos pada rentang waktu ini. Coba perpanjang rentang waktu atau variasikan kata kunci.")
    else:
        # Ringkasan KPI
        total_items = len(articles)
        pos_count = sum(1 for a in articles if a['sentiment'] == "Positif")
        neg_count = sum(1 for a in articles if a['sentiment'] == "Negatif")
        neu_count = sum(1 for a in articles if a['sentiment'] == "Netral")
        medsos_count = sum(1 for a in articles if a['channel_type'] == "medsos")
        online_count = sum(1 for a in articles if a['channel_type'] == "online_news")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Temuan", total_items)
        c2.metric("Media Online", online_count)
        c3.metric("Medsos (IG/FB/X/Threads)", medsos_count)
        c4.metric("Sentimen Positif", pos_count)
        c5.metric("Sentimen Negatif (Krisis)", neg_count)

        st.divider()

        # Tombol Download PowerPoint
        pptx_file = generate_pptx_deck(
            st.session_state['entity'], 
            st.session_state['org_type'], 
            st.session_state['time_label'],
            articles
        )
        st.download_button(
            label=f"📥 Unduh Laporan Slide Deck ({total_items} Slide Analisis .PPTX)",
            data=pptx_file,
            file_name=f"PR_Audit_{st.session_state['entity'].replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        # Pratinjau di Web
        st.subheader("📋 Kliping Digital & Rekomendasi Media Konsultan")
        for idx, art in enumerate(articles, 1):
            with st.container():
                col_left, col_right = st.columns([1, 1], gap="medium")
                
                with col_left:
                    st.image(art["image_bytes"], caption=f"Capture Dokumen #{idx} - {art['source']}", use_container_width=True)
                    st.markdown(f"[🔗 Buka Tautan Asli ({art['source']})]({art['link']})")

                with col_right:
                    badge_channel = "📱 [MEDIA SOSIAL]" if art['channel_type'] == "medsos" else "📰 [MEDIA ONLINE]"
                    st.markdown(f"### {badge_channel} #{idx}: {art['title']}")
                    st.caption(f"**Sumber/Kanal:** {art['source']} | **Waktu:** {art['published']}")
                    
                    st.markdown("#### 📝 Ringkasan Isi:")
                    for b in art['bullets']:
                        st.markdown(f"- {b}")
                    
                    st.markdown("---")
                    badge_color = "🟢" if art['sentiment'] == "Positif" else "🔴" if art['sentiment'] == "Negatif" else "🟡"
                    st.markdown(f"#### {badge_color} Sentimen: **{art['sentiment']}** ({art['pr_tone']})")
                    st.info(art['pr_rec'])

                st.markdown("---")
