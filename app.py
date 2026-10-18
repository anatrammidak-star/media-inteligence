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
# 1. KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Media & Social Intelligence Monitor",
    page_icon="📡",
    layout="wide"
)

st.title("📡 Media & Social Intelligence Monitoring System")
st.caption("Pemantauan Berita Nasional, Media Lokal/Kecil, Medsos, serta Rekomendasi Konsultan Media & Slide PPTX Lengkap")

# ==========================================
# 2. HELPER TANGGAL & PARSER KATA KUNCI
# ==========================================
BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

def calculate_date_range_label(time_val):
    """
    Menghitung rentang tanggal presisi untuk sampul laporan PPTX
    Contoh: '16 Oktober 2026 – 18 Oktober 2026'
    """
    today = datetime.now()
    days_map = {"1d": 1, "2d": 2, "3d": 3, "7d": 7, "30d": 30}
    
    if time_val in days_map:
        start_date = today - timedelta(days=days_map[time_val])
        return f"{start_date.day} {BULAN_ID[start_date.month]} {start_date.year} – {today.day} {BULAN_ID[today.month]} {today.year}"
    else:
        return f"Hingga {today.day} {BULAN_ID[today.month]} {today.year} (Semua Periode)"

def build_keyword_query(keywords_raw):
    """
    Mengurai input kata kunci fleksibel (koma, tanda kutip ganda, spasi).
    Contoh: 'mandiri, "bank mandiri", hutan, "hutan terbakar"'
    Output: '("bank mandiri" OR "hutan terbakar" OR mandiri OR hutan)'
    """
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
# 3. GENERATOR KARTU CAPTURE VISUAL (PNG)
# ==========================================
def generate_press_capture_card(channel_type, source, title, published, snippet, sentiment):
    """
    Membuat file gambar PNG kartu kliping berkualitas tinggi secara programatik.
    Disematkan langsung ke PowerPoint (embedded), ringan, dan anti-gagal/tidak loading mutar.
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

    # Frame border luar
    draw.rectangle([(0, 0), (width - 1, height - 1)], outline=(203, 213, 225), width=3)
    
    # Header Bar
    draw.rectangle([(0, 0), (width, 55)], fill=(241, 245, 249))
    draw.rectangle([(0, 0), (10, height)], fill=accent_color)

    # Label Channel
    tag = "[MEDIA SOSIAL]" if channel_type == "medsos" else "[PORTAL BERITA]"
    draw.text((25, 18), f"{tag}  •  {source.upper()}", fill=(30, 41, 59))

    # Badge Sentimen di kanan atas
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
    draw.text((25, y), "HEADLINE / JUDUL POSTINGAN:", fill=(71, 85, 105))
    y += 24
    for line in wrap_text(title, limit=42)[:3]:
        draw.text((25, y), line, fill=(15, 23, 42))
        y += 26

    # Garis Pembatas
    y += 12
    draw.line([(25, y), (width - 25, y)], fill=(226, 232, 240), width=1)
    y += 18

    # Cuplikan Konten
    draw.text((25, y), "DOKUMENTASI KUTIPAN KONTEN:", fill=(71, 85, 105))
    y += 24
    for line in wrap_text(snippet, limit=48)[:5]:
        draw.text((25, y), line, fill=(51, 65, 85))
        y += 24

    # Footer Watermark
    draw.rectangle([(0, height - 38), (width, height)], fill=(248, 250, 252))
    draw.text((25, height - 26), "VERIFIED BY STRATEGIC MEDIA AUDIT ENGINE", fill=(148, 163, 184))

    buf = io.BytesIO()
    img.save(buf, format
