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
st.caption("Pemantauan Berita Nasional, Media Lokal, Medsos + Rekomendasi PR Eksekutif & Slide PPTX Lengkap")

# ==========================================
# 2. HELPER TANGGAL & PARSING KATA KUNCI
# ==========================================
BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

def calculate_date_range_label(time_val):
    today = datetime.now()
    days_map = {"1d": 1, "2d": 2, "3d": 3, "7d": 7, "30d": 30}
    
    if time_val in days_map:
        start_date = today - timedelta(days=days_map[time_val])
        return f"{start_date.day} {BULAN_ID[start_date.month]} {start_date.year} – {today.day} {BULAN_ID[today.month]} {today.year}"
    else:
        return f"Hingga {today.day} {BULAN_ID[today.month]} {today.year} (Semua Periode)"

def build_keyword_query(keywords_raw):
    if not keywords_raw.strip():
        return ""
    parts = re.split(r'[,;\n]+', keywords_raw)
    parsed_tokens = []
    for p in parts:
        token = p.strip()
        if not token:
            continue
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
# 3. GENERATOR KARTU VISUAL CAPTURE (PNG NATIVE)
# ==========================================
def generate_press_capture_card(channel_type, source, title, published, snippet, sentiment):
    """
    Membuat file gambar PNG berkualitas tinggi dari clipping berita.
    Dijamin selalu valid dan langsung muncul di slide PowerPoint.
    """
    width, height = 800, 520
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

    # Frame Border Luar
    draw.rectangle([(0, 0), (width - 1, height - 1)], outline=(203, 213, 225), width=3)
    
    # Top Bar Header
    draw.rectangle([(0, 0), (width, 55)], fill=(241, 245, 249))
    draw.rectangle([(0, 0), (10, height)], fill=accent_color)

    # Label Media / Channel
    tag = "[MEDIA SOSIAL]" if channel_type == "medsos" else "[PORTAL BERITA]"
    draw.text((25, 18), f"{tag}  •  {source.upper()}", fill=(30, 41, 59))

    # Badge Sentimen
    draw.rectangle([(width - 150, 12), (width - 25, 42)], fill=badge_bg)
    draw.text((width - 135, 18), sentiment.upper(), fill=accent_color)

    # Waktu
    draw.text((25, 70), f"Waktu Rilis / Temuan: {published}", fill=(100, 116, 139))

    def wrap_text(text, limit=46):
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

    # Headline
    y = 105
    draw.text((25, y), "HEADLINE / JUDUL:", fill=(71, 85, 105))
    y += 24
    for line in wrap_text(title, limit=42)[:3]:
        draw.text((25, y), line, fill=(15, 23, 42))
        y += 26

    # Garis Pembatas
    y += 12
    draw.line([(25, y), (width - 25, y)], fill=(226, 232, 240), width=1)
    y += 18

    # Cuplikan Isi
    draw.text((25, y), "KUTIPAN KONTEN / TEKS ARTIKEL:", fill=(71, 85, 105))
    y += 24
    for line in wrap_text(snippet, limit=48)[:5]:
        draw.text((25, y), line, fill=(51, 65, 85))
        y += 24

    # Footer Watermark
    draw.rectangle([(0, height - 38), (width, height)], fill=(248, 250, 252))
    draw.text((25, height - 26), "DIVERIFIKASI MEDIA INTELLIGENCE ADVISORY SYSTEM", fill=(148, 163, 184))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()

# ==========================================
# 4. SUMMARY MINIMAL 2 BUTIR
# ==========================================
def extract_bullet_summary(title, text, min_bullets=2, max_bullets=3):
    cleaned = re.sub(r'<[^<]+?>', '', text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', cleaned)
    bullets = [s.strip() for s in sentences if len(s.strip()) > 25]

    if len(bullets) < 2:
        bullets = []
        bullets.append(f"Fokus Utama: {title.strip()}")
        if len(cleaned) > 20:
            bullets.append(f"Substansi Pemberitaan: {cleaned[:140]}...")
        else:
            bullets.append("Publikasi menyoroti perkembangan operasional dan respons publik terkait subjek.")
            
    return bullets[:max_bullets]

# ==========================================
# 5. ENGINE PEMINDAI BERITA & MEDSOS
# ==========================================
def fetch_media_and_social(entity_name, raw_keywords, time_window, max_results, include_news, include_social, include_small_media):
    parsed_kw = build_keyword_query(raw_keywords)
    combined_query_base = f"{entity_name} {parsed_kw}".strip()

    search_targets = []
    if include_news:
        search_targets.append(("news", combined_query_base, "online_news"))
    if include_small_media:
        search_targets.append(("web", combined_query_base, "local_web"))
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
        rss_url = f"https://news.google.com/rss/search?q={encoded_q}&hl=id&gl=ID&ceid=ID:id"
        feed = feedparser.parse(rss_url)

        for entry in feed.entries:
            link = entry.link
            if link in seen_links:
                continue
            seen_links.add(link)

            title = entry.title
            published = entry.get("published", "Terkini")
            
            source = "Media Nasional"
            clean_title = title
            if " - " in title:
                parts = title.rsplit(" - ", 1)
                clean_title = parts[0]
                source = parts[1]

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

            bullet_points = extract_bullet_summary(clean_title, snippet, min_bullets=2)

            t_low = (clean_title + " " + snippet).lower()
            neg_count = sum(1 for w in negative_lexicon if w in t_low)
            pos_count = sum(1 for w in positive_lexicon if w in t_low)

            if neg_count > pos_count:
                sentiment = "Negatif"
                pr_tone = "High Risk / Potential Crisis"
                if channel_type == "medsos":
                    pr_rec = (
                        "Rekomendasi Respons Medsos:\n"
                        "1. Aktifkan Social Media Response Unit untuk menjawab komplain dengan santun.\n"
                        "2. Jangan menghapus kritik warganet secara sepihak; arahkan ke jalur pelaporan resmi.\n"
                        "3. Pantau laju interaksi untuk mencegah eskalasi menjadi trending topic."
                    )
                else:
                    pr_rec = (
                        "Rekomendasi Konsultan Media:\n"
                        "1. Segera siapkan 'Holding Statement' resmi dan sebarkan hak jawab proporsional.\n"
                        "2. Adakan media briefing proaktif kepada jurnalis peliput untuk meluruskan fakta.\n"
                        "3. Siapkan FAQ mitigasi isu bagi juru bicara institusi."
                    )
            elif pos_count > neg_count:
                sentiment = "Positif"
                pr_tone = "Strategic Traction / Good PR"
                if channel_type == "medsos":
                    pr_rec = (
                        "Rekomendasi Respons Medsos:\n"
                        "1. Repost dan amplifikasi postingan ke kanal Instagram/Threads resmi institusi.\n"
                        "2. Jalin engagement dengan memberikan likes/komentar apresiasi resmi.\n"
                        "3. Gunakan materi sebagai 'Social Proof' kredibilitas layanan."
                    )
                else:
                    pr_rec = (
                        "Rekomendasi Konsultan Media:\n"
                        "1. Amplifikasi narasi positif melalui siaran pers lanjutan ke media nasional.\n"
                        "2. Jadikan rujukan pada public expose pimpinan dan bahan laporan tahunan.\n"
                        "3. Jaga hubungan relasi media dengan jurnalis yang meliput prestasi ini."
                    )
            else:
                sentiment = "Netral"
                pr_tone = "Informative / General Update"
                pr_rec = (
                    "Rekomendasi Konsultan Media:\n"
                    "1. Monitor dinamika pemberitaan lanjutan (follow-up coverage).\n"
                    "2. Sediakan factsheet ringkas jika terdapat awak media yang meminta rincian data."
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
# 6. GENERATOR PPTX DENGAN LINK & SLIDE RANGKUMAN AKHIR
# ==========================================
def generate_pptx_deck(entity_name, org_type, date_period_label, articles):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: COVER ---
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
    p1.text = f"Target Entitas: {entity_name} | Kategori: {org_type}"
    p1.font.size = Pt(22)
    p1.font.color.rgb = RGBColor(70, 80, 95)

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

    # --- SLIDE 2+: PER ITEM DOKUMENTASI (2 KOLOM) ---
    for idx, art in enumerate(articles, 1):
        slide = prs.slides.add_slide(blank_layout)

        # Header Title
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.6))
        htf = header_box.text_frame
        hp = htf.paragraphs[0]
        channel_tag = "MEDSOS" if art['channel_type'] == "medsos" else "MEDIA"
        hp.text = f"DOKUMENTASI #{idx} [{channel_tag}] : {art['source'].upper()}"
        hp.font.size = Pt(17)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(30, 41, 59)

        # SISI KIRI: GAMBAR CAPTURE PNG (EMBEDDED LOKAL)
        if art.get("image_bytes"):
            img_stream = io.BytesIO(art["image_bytes"])
            slide.shapes.add_picture(
                img_stream, 
                Inches(0.8), Inches(1.2), 
                width=Inches(5.6), height=Inches(4.8)
            )

        # LINK TAUTAN AKTIF DI BAWAH SCREENSHOT (BISA DIKLIK)
        link_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.1), Inches(5.6), Inches(0.8))
        ltf = link_box.text_frame
        ltf.word_wrap = True
        lp = ltf.paragraphs[0]
        lp.text = f"🔗 Tautan Sumber ({art['source']}):\n"
        lp.font.size = Pt(10)
        lp.font.bold = True
        lp.font.color.rgb = RGBColor(71, 85, 105)

        # Tambahkan Hyperlink aktif
        link_p = ltf.add_paragraph()
        link_p.text = art['link'][:65] + ("..." if len(art['link']) > 65 else "")
        link_p.font.size = Pt(10)
        link_p.font.underline = True
        link_p.font.color.rgb = RGBColor(37, 99, 235)
        link_p.hyperlink.address = art['link']

        # SISI KANAN: RINGKASAN MINIMAL 2 BUTIR + STRATEGIC ADVISORY
        right_box = slide.shapes.add_textbox(Inches(6.7), Inches(1.2), Inches(5.8), Inches(5.7))
        rtf = right_box.text_frame
        rtf.word_wrap = True

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

    # ==========================================
    # SLIDE TERAKHIR: EXECUTIVE STRATEGIC ADVISORY SUMMARY
    # ==========================================
    summary_slide = prs.slides.add_slide(blank_layout)

    sum_header = summary_slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.8))
    stf = sum_header.text_frame
    sp = stf.paragraphs[0]
    sp.text = "EXECUTIVE SUMMARY: STRATEGIC ACTIONABLE RECOMMENDATIONS"
    sp.font.size = Pt(20)
    sp.font.bold = True
    sp.font.color.rgb = RGBColor(20, 35, 60)

    # 3 Kolom Matriks Rekomendasi (Negatif, Netral, Positif)
    col_w = Inches(3.7)
    gap = Inches(0.3)
    start_x = Inches(0.8)
    top_y = Inches(1.3)
    card_h = Inches(5.6)

    # Kategori 1: Krisis & Isu Negatif
    neg_box = summary_slide.shapes.add_textbox(start_x, top_y, col_w, card_h)
    ntf = neg_box.text_frame
    ntf.word_wrap = True
    np1 = ntf.paragraphs[0]
    np1.text = "ISU NEGATIF / KRISIS"
    np1.font.size = Pt(14)
    np1.font.bold = True
    np1.font.color.rgb = RGBColor(220, 38, 38)

    np2 = ntf.add_paragraph()
    np2.text = (
        "Langkah Mitigasi Eksekutif:\n"
        "• Segera rilis Hak Jawab / Holding Statement resmi dalam 24 jam.\n"
        "• Lakukan penelusuran fakta objektif internal sebelum juru bicara berbicara di publik.\n"
        "• Berikan penanganan santun & terarah pada kanal keluhan publik/medsos.\n"
        "• Pantau agar isu lokal tidak tereskalasi menjadi isu nasional di media tier-1."
    )
    np2.font.size = Pt(11)
    np2.font.color.rgb = RGBColor(51, 65, 85)

    # Kategori 2: Isu Netral
    neu_box = summary_slide.shapes.add_textbox(start_x + col_w + gap, top_y, col_w, card_h)
    netf = neu_box.text_frame
    netf.word_wrap = True
    nep1 = netf.paragraphs[0]
    nep1.text = "ISU NETRAL / INFORMATIF"
    nep1.font.size = Pt(14)
    nep1.font.bold = True
    nep1.font.color.rgb = RGBColor(202, 138, 4)

    nep2 = netf.add_paragraph()
    nep2.text = (
        "Langkah Pengawasan & Edukasi:\n"
        "• Cermati perkembangan isu apakah terdapat potensi pergeseran persepsi publik.\n"
        "• Sediakan lembar fakta (factsheet) pendukung bagi wartawan yang meminta data komparasi.\n"
        "• Edukasi publik melalui artikel opini atau infografis sederhana di situs web resmi."
    )
    nep2.font.size = Pt(11)
    nep2.font.color.rgb = RGBColor(51, 65, 85)

    # Kategori 3: Isu Positif & Prestasi
    pos_box = summary_slide.shapes.add_textbox(start_x + (col_w + gap) * 2, top_y, col_w, card_h)
    ptf = pos_box.text_frame
    ptf.word_wrap = True
    pp1 = ptf.paragraphs[0]
    pp1.text = "ISU POSITIF / PRESTASI"
    pp1.font.size = Pt(14)
    pp1.font.bold = True
    pp1.font.color.rgb = RGBColor(22, 163, 74)

    pp2 = ptf.add_paragraph()
    pp2.text = (
        "Langkah Amplifikasi Reputasi:\n"
        "• Perluas jangkauan narasi prestasi ke owned media (LinkedIn, Press Release, IG Reels/Threads).\n"
        "• Manfaatkan momentum keberhasilan untuk Thought Leadership Direksi/Pimpinan.\n"
        "• Jalin silaturahmi apresiasi dengan jurnalis yang memberitakan secara berimbang."
    )
    pp2.font.size = Pt(11)
    pp2.font.color.rgb = RGBColor(51, 65, 85)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# 7. SIDEBAR KONTROL INPUT
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
    
    keywords_input = st.text_area(
        "Kata Kunci / Frasa Isu (Bisa kombinasi):",
        value='mandiri, "bank mandiri", layanan, fraud',
        help='Pisahkan dengan koma atau baris baru. Contoh: mandiri, "bank mandiri", hutan, "hutan terbakar"'
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
# 8. DASHBOARD & EKSPOR
# ==========================================
if run_btn:
    if not (inc_news or inc_small or inc_social):
        st.error("Silakan centang minimal satu sumber data!")
    else:
        with st.spinner("Memindai media nasional, web lokal, medsos, dan menyiapkan slide eksekutif..."):
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
        c2_neg = c3.metric("Sentimen Negatif (Mitigasi)", neg_count)
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
            label=f"📥 Unduh Laporan Slide Deck ({total_items + 2} Slide .PPTX Termasuk Rangkuman Eksekutif)",
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
                    st.markdown(f"🔗 **Tautan Langsung:** [{art['link'][:70]}...]({art['link']})")

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
