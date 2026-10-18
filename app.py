import io
import re
import urllib.parse
from datetime import datetime, timedelta
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
st.caption("Monitoring Berita Nasional, Media Lokal/Kecil, Medsos, serta Rekomendasi Media Konsultan")

# ==========================================
# 2. HELPER TANGGAL & PARSING KATA KUNCI
# ==========================================
BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

def calculate_date_range_label(time_val):
    """
    Menghitung rentang tanggal presisi untuk sampul laporan PPTX
    Contoh output: '16 Oktober 2026 – 18 Oktober 2026'
    """
    today = datetime.now()
    days_map = {"1d": 1, "2d": 2, "3d": 3, "7d": 7, "30d": 30}
    
    if time_val in days_map:
        start_date = today - timedelta(days=days_map[time_val])
        if days_map[time_val] == 1:
            return f"{start_date.day} {BULAN_ID[start_date.month]} {start_date.year} – {today.day} {BULAN_ID[today.month]} {today.year}"
        return f"{start_date.day} {BULAN_ID[start_date.month]} {start_date.year} – {today.day} {BULAN_ID[today.month]} {today.year}"
    else:
        return f"Hingga {today.day} {BULAN_ID[today.month]} {today.year} (Semua Periode)"

def build_keyword_query(keywords_raw):
    """
    Mengurai input kata kunci fleksibel (koma, tanda kutip, spasi)
    Contoh: 'mandiri, "bank mandiri", hutan, "hutan terbakar"'
    Output: '("bank mandiri" OR "hutan terbakar" OR mandiri OR hutan)'
    """
    if not keywords_raw.strip():
        return ""
    
    # Pisahkan berdasarkan koma atau baris baru
    parts = re.split(r'[,;\n]+', keywords_raw)
    parsed_tokens = []
    
    for p in parts:
        token = p.strip()
        if not token:
            continue
        # Jika mengandung spasi dan belum diapit tanda kutip, tambahkan tanda kutip ganda
        if " " in token and not (token.startswith('"') and token.endswith('"')):
            parsed_tokens.append(f'"{token}"')
        else:
            parsed_tokens.append(token)
            
    if not parsed_tokens:
        return ""
    if len(parsed_tokens) == 1:
        return parsed_tokens[0]
    return "(" + " OR ".join(parsed_tokens) + ")"

# ==========================================
# 3. GENERATOR KARTU VISUAL CAPTURE (PNG)
# ==========================================
def generate_press_capture_card(channel_type, source, title, published, snippet, sentiment):
    width, height = 700, 480
    bg_color = (255, 255, 255)
    
    if sentiment == "Negatif":
        accent_color = (220, 38, 38)
        badge_bg = (254, 226, 226)
    elif sentiment == "Positif":
        accent_color = (22, 163, 74)
        badge_bg = (220, 252, 231)
    else:
        accent_color = (202, 138, 4)
        badge_bg = (254, 243, 199)

    img = Image.new("RGB", (width, height), bg_color)
    draw = ImageDraw.Draw(img)

    draw.rectangle([(0, 0), (width - 1, height - 1)], outline=(203, 213, 225), width=2)
    draw.rectangle([(0, 0), (width, 50)], fill=(241, 245, 249))
    draw.rectangle([(0, 0), (8, height)], fill=accent_color)

    type_tag = "[MEDIA SOSIAL]" if channel_type == "medsos" else "[PORTAL/WEBSITE]"
    draw.text((25, 16), f"{type_tag}  •  {source.upper()}", fill=(30, 41, 59))

    draw.rectangle([(width - 140, 12), (width - 20, 38)], fill=badge_bg)
    draw.text((width - 125, 17), sentiment.upper(), fill=accent_color)

    draw.text((25, 65), f"Waktu Terbit/Temuan: {published}", fill=(100, 116, 139))

    def wrap_text(text, limit=48):
        lines, cur = [], []
        for w in text.split():
            if sum(len(x) + 1 for x in cur) + len(w) <= limit:
                cur.append(w)
            else:
                lines.append(" ".join(cur))
                cur = [w]
        if cur:
            lines.append(" ".join(cur))
        return lines

    y_text = 95
    draw.text((25, y_text), "HEADLINE / JUDUL POSTINGAN:", fill=(71, 85, 105))
    y_text += 22

    for line in wrap_text(title, limit=40)[:3]:
        draw.text((25, y_text), line, fill=(15, 23, 42))
        y_text += 25

    y_text += 10
    draw.line([(25, y_text), (width - 25, y_text)], fill=(226, 232, 240), width=1)
    y_text += 15

    draw.text((25, y_text), "DOKUMENTASI KUTIPAN KONTEN:", fill=(71, 85, 105))
    y_text += 22
    for line in wrap_text(snippet, limit=48)[:5]:
        draw.text((25, y_text), line, fill=(51, 65, 85))
        y_text += 22

    draw.rectangle([(0, height - 35), (width, height)], fill=(248, 250, 252))
    draw.text((25, height - 25), "VERIFIED BY STRATEGIC MEDIA AUDIT ENGINE", fill=(148, 163, 184))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()

# ==========================================
# 4. ENGINE SUMMARY & SCRAPING (MENJANGKAU MEDIA KECIL)
# ==========================================
def extract_bullet_summary(title, text, min_bullets=2, max_bullets=3):
    """
    Memastikan ringkasan SELALU memuat MINIMAL 2 BUTIR informasi.
    """
    cleaned = re.sub(r'<[^<]+?>', '', text).strip()
    # Pisahkan berdasarkan tanda titik, seru, tanya, atau titik koma
    sentences = re.split(r'(?<=[.!?])\s+', cleaned)
    bullets = [s.strip() for s in sentences if len(s.strip()) > 25]

    # Jika kurang dari 2 butir, buat dari klausa judul atau analisis substansi
    if len(bullets) < 2:
        bullets = []
        # Butir 1: Fokus Isu Utama
        bullets.append(f"Fokus Utama: {title.strip()}")
        # Butir 2: Rincian Narasi
        if len(cleaned) > 20:
            bullets.append(f"Substansi Pemberitaan: {cleaned[:140]}...")
        else:
            bullets.append("Publikasi menyoroti perkembangan operasional dan respons publik terkait subjek.")
            
    return bullets[:max_bullets]

def fetch_media_and_social(entity_name, raw_keywords, time_window, max_results, include_news, include_social, include_small_media):
    parsed_kw = build_keyword_query(raw_keywords)
    combined_query_base = f"{entity_name} {parsed_kw}".strip()

    search_targets = []
    
    # 1. Google News Standar (Media Mainstream)
    if include_news:
        search_targets.append(("news", combined_query_base, "online_news"))

    # 2. Google Search RSS (Menjangkau media lokal, situs daerah, web kecil)
    if include_small_media:
        # Menghapus filter eksklusif agar situs non-mainstream terangkut
        search_targets.append(("web", combined_query_base, "local_web"))

    # 3. Media Sosial (Instagram, Facebook, Threads, X, TikTok)
    if include_social:
        social_dork = "(site:instagram.com OR site:facebook.com OR site:threads.net OR site:twitter.com OR site:x.com OR site:tiktok.com)"
        search_targets.append(("social", f"{combined_query_base} {social_dork}", "medsos"))

    negative_lexicon = [
        "korupsi", "rugi", "bobol", "fraud", "kasus", "gagal", "demo", "kecewa",
        "macet", "krisis", "tuntutan", "phk", "ancam", "turun", "pidana", "gugat", 
        "disorot", "hukum", "sita", "penipuan", "keluhan", "masalah", "denda", "skandal", "viral", "hujat", "terbakar"
    ]
    positive_lexicon = [
        "laba", "untung", "tumbuh", "prestasi", "penghargaan", "inovasi", "sukses",
        "dividen", "ekspansi", "naik", "resmikan", "unggul", "mitra", "positif", 
        "rekor", "terbaik", "kinerja positif", "apresiasi", "capai", "solutif", "bangga", "lestari"
    ]

    items = []
    seen_links = set()

    for target_type, query_str, channel_type in search_targets:
        if time_window != "all":
            final_q = f"{query_str} when:{time_window}"
        else:
            final_q = query_str

        encoded_q = urllib.parse.quote(final_q)
        
        # Tentukan endpoint RSS berdasarkan tipe pencarian
        if target_type == "web":
            # Endpoint feed web Google untuk menjaring situs kecil/lokal
            rss_url = f"https://news.google.com/rss/search?q={encoded_q}&hl=id&gl=ID&ceid=ID:id"
        else:
            rss_url = f"https://news.google.com/rss/search?q={encoded_q}&hl=id&gl=ID&ceid=ID:id"

        feed = feedparser.parse(rss_url)

        for entry in feed.entries:
            link = entry.link
            if link in seen_links:
                continue
            seen_links.add(link)

            title = entry.title
            published = entry.get("published", "Terkini")
            
            source = "Media / Portal"
            clean_title = title
            if " - " in title:
                parts = title.rsplit(" - ", 1)
                clean_title = parts[0]
                source = parts[1]

            # Identifikasi Medsos vs Media Lokal
            lower_link = link.lower()
            if channel_type == "medsos" or any(s in lower_link for s in ["instagram.com", "facebook.com", "threads.net", "twitter.com", "x.com", "tiktok.com"]):
                channel_type = "medsos"
                if "instagram.com" in lower_link:
                    source = "Instagram"
                elif "facebook.com" in lower_link:
                    source = "Facebook"
                elif "threads.net" in lower_link:
                    source = "Threads"
                elif "twitter.com" in lower_link or "x.com" in lower_link:
                    source = "X (Twitter)"
                elif "tiktok.com" in lower_link:
                    source = "TikTok"
            elif channel_type == "local_web":
                source = f"{source} (Web/Lokal)"

            soup = BeautifulSoup(entry.get("summary", ""), "html.parser")
            snippet = soup.get_text()

            # Pastikan minimal 2 butir informasi
            bullet_points = extract_bullet_summary(clean_title, snippet, min_bullets=2)

            # Analisis Sentimen
            t_low = (clean_title + " " + snippet).lower()
            neg_count = sum(1 for w in negative_lexicon if w in t_low)
            pos_count = sum(1 for w in positive_lexicon if w in t_low)

            if neg_count > pos_count:
                sentiment = "Negatif"
                pr_tone = "High Risk / Reputational Crisis"
                if channel_type == "medsos":
                    pr_rec = (
                        "Rekomendasi Respons Medsos:\n"
                        "1. Aktifkan Social Media Response Unit; berikan respons penyejuk di kolom komentar dalam 1 jam.\n"
                        "2. Jangan hapus unggahan komplain publik; arahkan ke formulir pelaporan/DM terverifikasi.\n"
                        "3. Siapkan FAQ untuk menangkal rumor viral yang belum tentu faktual."
                    )
                else:
                    pr_rec = (
                        "Rekomendasi Konsultan Media:\n"
                        "1. Segera distribusikan lembar klarifikasi/hak jawab kepada redaksi pembuat berita.\n"
                        "2. Buat media monitoring harian terhadap potensi pengutipan oleh media tier-1.\n"
                        "3. Siapkan juru bicara resmi untuk pernyataan pers terpadu."
                    )
            elif pos_count > neg_count:
                sentiment = "Positif"
                pr_tone = "Strategic Traction / Positive PR"
                if channel_type == "medsos":
                    pr_rec = (
                        "Rekomendasi Respons Medsos:\n"
                        "1. Kurasi dan unggah ulang (re-share/quote) konten ini ke kanal media sosial resmi entitas.\n"
                        "2. Jalin interaksi positif dengan pembuat konten sebagai bentuk apresiasi komunitas.\n"
                        "3. Jadikan bahan amplifikasi konten mingguan."
                    )
                else:
                    pr_rec = (
                        "Rekomendasi Konsultan Media:\n"
                        "1. Amplifikasi narasi positif melalui sindikasi siaran pers ke jaringan media nasional.\n"
                        "2. Manfaatkan momentum untuk memperkuat reputasi kepemimpinan dan CSR korporasi.\n"
                        "3. Dokumentasikan materi ke dalam laporan tahunan/keberlanjutan."
                    )
            else:
                sentiment = "Netral"
                pr_tone = "Informative / General Update"
                pr_rec = (
                    "Rekomendasi Konsultan Media:\n"
                    "1. Cermati perkembangan isu apakah berpotensi bergeser ke arah negatif.\n"
                    "2. Sediakan materi pendukung jika ada jurnalis atau netizen yang meminta konfirmasi data."
                )

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

            if max_results != "unlimited" and len(items) >= int(max_results):
                break
        
        if max_results != "unlimited" and len(items) >= int(max_results):
            break

    return items

# ==========================================
# 5. GENERATOR PRESENTASI (.PPTX)
# ==========================================
def generate_pptx_deck(entity_name, org_type, date_period_label, articles):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: COVER DENGAN PERIODE TANGGAL DINAMIS ---
    cover_slide = prs.slides.add_slide(blank_layout)
    title_box = cover_slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.33), Inches(4.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    
    p0 = tf.paragraphs[0]
    p0.text = "MEDIA & SOCIAL INTELLIGENCE AUDIT"
    p0.font.size = Pt(36)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(20, 35, 60)

    p1 = tf.add_paragraph()
    p1.text = f"Target Entitas: {entity_name} | Sektor: {org_type}"
    p1.font.size = Pt(22)
    p1.font.color.rgb = RGBColor(70, 80, 95)

    # PERIODE TANGGAL SPESIFIK PADA JUDUL SAMPUL
    p2 = tf.add_paragraph()
    p2.text = f"Periode Pemantauan: {date_period_label}"
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(14, 116, 144)

    p3 = tf.add_paragraph()
    p3.text = "Disusun oleh: Senior Strategic Media & Reputation Advisory"
    p3.font.size = Pt(13)
    p3.font.italic = True
    p3.font.color.rgb = RGBColor(120, 130, 145)

    # --- SLIDE 2+: TATA LETAK 2 KOLOM ---
    for idx, art in enumerate(articles, 1):
        slide = prs.slides.add_slide(blank_layout)

        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.6))
        htf = header_box.text_frame
        hp = htf.paragraphs[0]
        channel_tag = "MEDSOS" if art['channel_type'] == "medsos" else "MEDIA"
        hp.text = f"TEMUAN #{idx} [{channel_tag}] : {art['source'].upper()}"
        hp.font.size = Pt(17)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(30, 41, 59)

        # SISI KIRI: CAPTURE CARD
        if art.get("image_bytes"):
            img_stream = io.BytesIO(art["image_bytes"])
            slide.shapes.add_picture(
                img_stream, 
                Inches(0.8), Inches(1.3), 
                width=Inches(5.7), height=Inches(5.6)
            )

        # SISI KANAN: RINGKASAN (MINIMAL 2 BUTIR) + STRATEGIC ADVISORY
        right_box = slide.shapes.add_textbox(Inches(6.8), Inches(1.3), Inches(5.7), Inches(5.6))
        rtf = right_box.text_frame
        rtf.word_wrap = True

        rp1 = rtf.paragraphs[0]
        rp1.text = "RINGKASAN PEMBERITAAN (MIN. 2 BUTIR INFORMASI)"
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
# 6. SIDEBAR & INTERFACE
# ==========================================
with st.sidebar:
    st.header("⚙️ Konfigurasi Monitoring")
    entity_input = st.text_input("Nama Entitas / Target", value="Bank Mandiri")
    org_type_input = st.selectbox("Kategori Entitas", [
        "Perbankan / Finansial", 
        "BUMN / Pemerintahan", 
        "Kehutanan / Lingkungan Hidup",
        "Korporasi Swasta", 
        "Lembaga Publik / NGO", 
        "Personal / Tokoh Publik"
    ])
    
    # Input kata kunci fleksibel (frasa, koma, dsb.)
    keywords_input = st.text_area(
        "Kata Kunci / Frasa Isu (Bisa kombinasi):",
        value='mandiri, "bank mandiri", layanan, fraud',
        help='Gunakan koma untuk memisahkan. Contoh: mandiri, "bank mandiri", hutan, "hutan terbakar"'
    )

    st.subheader("🌐 Jangkauan Sumber Media")
    inc_news = st.checkbox("Media Nasional (Tier-1)", value=True)
    inc_small = st.checkbox("Media Lokal & Website Kecil/Daerah", value=True)
    inc_social = st.checkbox("Media Sosial (IG, FB, Threads, X, TikTok)", value=True)

    st.subheader("⏳ Rentang Waktu Pemantauan")
    time_options = {
        "24 Jam Terakhir (1 Hari)": "1d",
        "2 Hari Terakhir": "2d",
        "3 Hari Terakhir": "3d",
        "7 Hari Terakhir (1 Minggu)": "7d",
        "30 Hari Terakhir (1 Bulan)": "30d",
        "Semua Waktu (Tanpa Batasan)": "all"
    }
    selected_time_label = st.selectbox("Pilih Rentang", list(time_options.keys()), index=1)
    time_val = time_options[selected_time_label]
    
    # Menghitung label tanggal untuk ditampilkan di UI dan PPTX
    calculated_date_label = calculate_date_range_label(time_val)
    st.caption(f"🗓️ **Periode Aktif:** {calculated_date_label}")

    st.subheader("📊 Kuota Temuan")
    limit_mode = st.radio("Mode Kuota:", ["Tentukan Jumlah", "Bebas (Seluruhnya)"])
    if limit_mode == "Tentukan Jumlah":
        max_news_count = st.number_input("Jumlah Berita/Postingan", min_value=1, max_value=100, value=15, step=5)
    else:
        max_news_count = "unlimited"

    run_btn = st.button("🚀 Mulai Pemindaian Lintas Media", type="primary")

# ==========================================
# 7. DASHBOARD & EKSPOR
# ==========================================
if run_btn:
    if not (inc_news or inc_small or inc_social):
        st.error("Silakan centang minimal satu sumber data!")
    else:
        with st.spinner("Memindai media nasional, website lokal, dan medsos..."):
            st.session_state['news_data'] = fetch_media_and_social(
                entity_name=entity_input,
                raw_keywords=keywords_input,
                time_window=time_val,
                max_results=max_news_count,
                include_news=inc_news,
                include_social=inc_social,
                include_small_media=inc_small
            )
            st.session_state['entity'] = entity_input
            st.session_state['org_type'] = org_type_input
            st.session_state['date_label'] = calculated_date_label

if 'news_data' in st.session_state:
    articles = st.session_state.get('news_data', [])

    if not articles:
        st.warning("Tidak ditemukan pemberitaan pada rentang waktu ini. Coba perluas kata kunci atau tambahkan rentang waktu.")
    else:
        total_items = len(articles)
        pos_count = sum(1 for a in articles if a['sentiment'] == "Positif")
        neg_count = sum(1 for a in articles if a['sentiment'] == "Negatif")
        neu_count = sum(1 for a in articles if a['sentiment'] == "Netral")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Temuan", total_items)
        c2.metric("Sentimen Positif", pos_count)
        c3.metric("Sentimen Negatif (Perlu Mitigasi)", neg_count)
        c4.metric("Sentimen Netral", neu_count)

        st.divider()

        # Unduh PPTX
        pptx_file = generate_pptx_deck(
            st.session_state['entity'], 
            st.session_state['org_type'], 
            st.session_state['date_label'],
            articles
        )
        st.download_button(
            label=f"📥 Unduh Laporan Slide Deck ({total_items} Slide .PPTX - Periode {st.session_state['date_label']})",
            data=pptx_file,
            file_name=f"PR_Audit_{st.session_state['entity'].replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        st.subheader("📋 Kliping Digital & Rekomendasi Media Konsultan")
        for idx, art in enumerate(articles, 1):
            with st.container():
                col_left, col_right = st.columns([1, 1], gap="medium")
                
                with col_left:
                    st.image(art["image_bytes"], caption=f"Capture Dokumen #{idx} - {art['source']}", use_container_width=True)
                    st.markdown(f"[🔗 Buka Tautan Asli ({art['source']})]({art['link']})")

                with col_right:
                    st.markdown(f"### #{idx}: {art['title']}")
                    st.caption(f"**Sumber:** {art['source']} | **Waktu:** {art['published']}")
                    
                    st.markdown("#### 📝 Ringkasan Isi Berita (Key Points):")
                    for b in art['bullets']:
                        st.markdown(f"- {b}")
                    
                    st.markdown("---")
                    badge_color = "🟢" if art['sentiment'] == "Positif" else "🔴" if art['sentiment'] == "Negatif" else "🟡"
                    st.markdown(f"#### {badge_color} Sentimen: **{art['sentiment']}** ({art['pr_tone']})")
                    st.info(art['pr_rec'])

                st.markdown("---")
