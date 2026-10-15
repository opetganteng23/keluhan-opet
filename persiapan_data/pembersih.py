"""Pra-pemrosesan teks keluhan sebelum dianalisis (dipakai kedua metode).

Menghapus emoji/simbol, link, mention, awalan retweet, dan kode HTML;
hashtag dipertahankan katanya (#macet -> macet); spasi dirapikan.
"""

import html
import re
import unicodedata

POLA_RETWEET = re.compile(r"^\s*RT\s+@\w+:?\s*", re.IGNORECASE)
POLA_LINK = re.compile(r"(https?://\S+|www\.\S+)", re.IGNORECASE)
POLA_MENTION = re.compile(r"(?<!\w)@\w+")
POLA_HASHTAG = re.compile(r"#(\w+)")
POLA_SPASI = re.compile(r"\s+")

# Karakter penyusun emoji yang bukan kategori "So" (penyambung & pemilih varian)
KARAKTER_EMOJI_TAMBAHAN = {"‍", "︎", "️", "⃣"}


def _bukan_emoji(karakter: str) -> bool:
    """True jika karakter bukan emoji/simbol gambar."""
    if karakter in KARAKTER_EMOJI_TAMBAHAN:
        return False
    kategori = unicodedata.category(karakter)
    # So = simbol lain (emoji, ikon), Sk = pengubah (warna kulit), Cs/Co = karakter tak valid/pribadi
    if kategori in ("So", "Cs", "Co"):
        return False
    if kategori == "Sk" and ord(karakter) > 0x2000:
        return False
    # Huruf regional (bendera) dan tag emoji
    return not (0x1F1E6 <= ord(karakter) <= 0x1F1FF or 0xE0020 <= ord(karakter) <= 0xE007F)


def hapus_emoji(teks: str) -> str:
    """Menghapus semua emoji dan simbol gambar dari teks."""
    return "".join(k for k in teks if _bukan_emoji(k))


def bersihkan_teks(teks: str) -> str:
    """Menetralkan teks: buang emoji, link, mention, RT, kode HTML; rapikan spasi."""
    if not teks:
        return ""
    teks = html.unescape(teks)
    teks = POLA_RETWEET.sub("", teks)
    teks = POLA_LINK.sub(" ", teks)
    teks = POLA_MENTION.sub(" ", teks)
    teks = POLA_HASHTAG.sub(r"\1", teks)
    teks = hapus_emoji(teks)
    return POLA_SPASI.sub(" ", teks).strip()
