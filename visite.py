"""
Generator "List Pasien per DPJP" (format template.docx).

Alur:
  1. Teks laporan bangsal (berantakan) -> dibersihkan dulu (buang URL PACS, spasi) supaya hemat token.
  2. AI (Claude) cuma mengelompokkan data ke JSON ringkas per pasien.
  3. Kode ini yang menyusun docx: tabel, warna, flag lab abnormal (pakai data-normal.xlsx),
     rekap jumlah pasien, dan catatan bawah. Jadi AI tidak perlu menulis format apa pun.

Pemakaian:
  python visite.py RB6.txt RB2.txt ... -o list.docx            (pakai Claude Code + langganan Pro;
                                                                 otomatis pakai API kalau ada ANTHROPIC_API_KEY)
  python visite.py RB6.txt --prompt-only > prompt.txt           (tanpa API: paste ke claude.ai,
  python visite.py --from-json RB6.json -o list.docx              simpan jawabannya jadi .json, lalu render)

Hasil AI di-cache per pasien di folder cache/: pasien yang teksnya sama dengan hari sebelumnya
tidak dikirim ulang ke AI, cuma pasien baru/berubah. Tiap run menyimpan snapshot ke riwayat/
dan membandingkan dengan list terakhir sebelumnya (pasien baru, keluar, pindah ruangan).
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

BASE = Path(__file__).parent
TEMPLATE = BASE / "template.docx"
NORMAL_XLSX = BASE / "data-normal.xlsx"
CACHE = BASE / "cache"
RIWAYAT = BASE / "riwayat"  # snapshot list tiap hari, untuk dibandingkan dengan hari berikutnya
DEFAULT_MODEL = "claude-haiku-4-5"

# keyword (huruf kecil, dicari di teks DPJP) -> (nama lengkap judul, nama pendek di header tabel)
DOKTER = [
    ("safruddin", "dr. SAFRUDDIN, Sp.B-KBD, MM", "dr. Safruddin"),
    ("adi", "Dr. dr. ADI MURADI MUHAR, Sp.B, Subsp.BD(K)", "dr. Adi"),
    ("siti", "dr. SITI FATHIYA, Sp.B, Subsp.BD(K)", "dr. Siti"),
    ("edwin", "dr. EDWIN SALEH SIREGAR, Sp.B, Subsp.BD(K)", "dr. Edwin"),
    ("budi", "dr. BUDI IRWAN, Sp.B-KBD", "dr. Budi"),
]

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]

KOLOM = [("Kamar", 630), ("Identitas", 1340), ("Diagnosis", 1580), ("Objectives", 2110),
         ("Masalah", 1010), ("Diet", 530), ("Tatalaksana", 1920), ("Planning", 2380)]

MERAH = RGBColor(0xCC, 0x00, 0x00)
PUTIH = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x1F, 0x38, 0x64)
ABU = RGBColor(0x80, 0x80, 0x80)

# ---------------------------------------------------------------- nilai normal lab
# Nilai cadangan untuk pemeriksaan yang tidak ada di data-normal.xlsx. (min, max); None = tanpa batas.
NORMAL_DEFAULT = {
    "na": (135, 145), "k": (3.5, 5.0), "cl": (98, 106),
    "ur": (10, 50), "cr": (0.6, 1.2), "egfr": (60, None), "gds": (70, 200),
    "biltotal": (None, 1.2), "bildirek": (None, 0.3), "sgot": (None, 40), "sgpt": (None, 41),
    "inr": (0.8, 1.2), "amilase": (None, 100), "lipase": (None, 60),
}
# alias nama di laporan -> kunci
ALIAS = {
    "hb": "hb", "hgb": "hb", "hemoglobin": "hb",
    "ht": "ht", "hct": "ht", "hematokrit": "ht",
    "leu": "leu", "leuko": "leu", "wbc": "leu", "leukosit": "leu",
    "plt": "plt", "tromb": "plt", "trombosit": "plt",
    "alb": "alb", "albumin": "alb",
    "cea": "cea", "mg": "mg", "magnesium": "mg",
    "na": "na", "k": "k", "cl": "cl", "ur": "ur", "ureum": "ur", "cr": "cr", "kreatinin": "cr",
    "egfr": "egfr", "gds": "gds", "bil total": "biltotal", "biltotal": "biltotal", "bil.total": "biltotal",
    "bil direk": "bildirek", "bildirek": "bildirek", "bil.direk": "bildirek",
    "sgot": "sgot", "sgpt": "sgpt", "inr": "inr", "amilase": "amilase", "lipase": "lipase",
}
XLSX_KEY = {"hemoglobin": "hb", "leukosit": "leu", "hematokrit": "ht", "trombosit": "plt",
            "albumin": "alb", "cea": "cea", "magnesium": "mg"}
RIBUAN = {"leu", "plt"}  # di laporan ditulis 8.460 / 119.000 (titik = ribuan)


def _num(s):
    return float(str(s).replace(",", "").strip())


def load_normal():
    ranges = dict(NORMAL_DEFAULT)
    if not NORMAL_XLSX.exists():
        return ranges
    import openpyxl
    ws = openpyxl.load_workbook(NORMAL_XLSX, data_only=True).active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[1] or row[3] is None:
            continue
        nama = str(row[1]).lower()
        key = next((v for k, v in XLSX_KEY.items() if nama.startswith(k)), None)
        if not key:
            continue
        rj = row[3]
        if isinstance(rj, (dt.datetime, dt.date)):  # Excel mengubah "12 - 16" jadi tanggal
            rj = f"{rj.month} - {rj.day}"
        rj = str(rj)
        try:
            if rj.startswith("<"):
                ranges[key] = (None, _num(rj[1:]))
            elif rj.startswith(">"):
                ranges[key] = (_num(rj[1:]), None)
            else:
                lo, hi = rj.split(" - ")
                ranges[key] = (_num(lo), _num(hi))
        except ValueError:
            pass
    return ranges


NORMAL = load_normal()


def abnormal(nama, nilai):
    """True kalau nilai lab di luar rentang normal."""
    key = ALIAS.get(nama.lower().strip().rstrip(":"))
    if not key or key not in NORMAL:
        return False
    v = nilai.strip().lstrip("<>")
    try:
        if key in RIBUAN:
            x = float(v.replace(".", "").replace(",", ".")) if re.fullmatch(r"\d{1,3}(\.\d{3})+", v) else float(v.replace(",", "."))
            if x < 100:  # ditulis dalam ribuan, mis. 8.5
                x *= 1000
        else:
            x = float(v.replace(",", "."))
    except ValueError:
        return False
    lo, hi = NORMAL[key]
    return (lo is not None and x < lo) or (hi is not None and x > hi)


# ---------------------------------------------------------------- AI
PROMPT = """Ubah laporan visite bangsal bedah di bawah jadi JSON. Tanggal laporan (hari ini): {tgl}.
Ruangan bagian ini: {ruang}.
Balas HANYA JSON valid, tanpa penjelasan, format:
{{"pasien":[{{
 "i":"nomor blok '### PASIEN n' asal pasien ini (angka)",
 "r":"ruangan pasien (mis RB1, RA3, Ruang 206, Pav 5). Biasanya = ruangan bagian ini, kecuali ada judul ruangan lain di teks (mis baris '206' sebelum pasien)",
 "k":"kamar/bed, mis 3.10.2",
 "n":"nama (Title Case)","jk":"Lk/Pr","u":"umur angka","rm":"no RM",
 "dpjp":"nama DPJP apa adanya","tm":"tgl masuk dd/mm/yyyy","to":"tgl operasi atau kosong",
 "st":"jadwal PBJ/operasi dari instruksi, format 'PBJ dd/mm' / 'Op dd/mm: tindakan' / ''. Tanggal HANYA bila tertulis. 'Op hari ini: tindakan' HANYA bila ada jadwal kamar operasi/jam (mis 'MOT-2 08.00') atau tertulis 'hari ini'; rencana/pro operasi tanpa jadwal -> ''. 'PBJ hari ini' -> 'PBJ hari ini'",
 "dx":"diagnosis lengkap apa adanya (1 string)",
 "tl":"tgl lab terbaru yang ADA nilainya dd/mm (judul 'Lab dd/mm' tanpa nilai diabaikan)",
 "o":["keluhan & tanda vital bila tertulis, mis 'S: nyeri luka operasi','TD 100/60, HR 69, RR 20, T 36.5, SpO2 98%'"],
 "lab":["HANYA lab terbaru, format 'Nama nilai' dipisah ', ', mis 'Hb 10.0, Ht 30.7, Leu 8.460, PLT 119.000','Na 136, K 4.1, Cl 106','Alb 2.56','Ur 56, Cr 0.94, eGFR 88.95','Bil total 3.38, Bil direk 2.53'"],
 "pj":["penunjang terbaru singkat, 'Nama: hasil sangat ringkas / susul bacaan', mis 'Foto thorax: cardiomegaly','CT abdomen (03/02): mets hepar'"],
 "m":["daftar masalah ringkas (bhs Indonesia, <=6 kata/item). Awali '!' jika masalah lab abnormal/kritis, mis '!Anemia (Hb 8.5)'"],
 "d":["diet, kosong jika tidak disebut"],
 "tx":["terapi singkat, mis 'RL 20tpm','Ranitidine/12jam'; termasuk terapi TS lain"],
 "p":["rencana/instruksi singkat (dari I:, rencana operasi, konsul, cek lab) HANYA yang tertulis"]
}}]}}
Aturan: jangan mengarang data; field kosong = "" atau []. Buang URL. Singkat.
Semua isi HARUS berasal dari blok pasien itu sendiri: jangan menambah terapi/rencana umum
(mis 'monitor kondisi', 'rawat luka') yang tidak tertulis, jangan memakai data pasien lain.
Angka disalin persis (koma desimal boleh jadi titik).
Baris rekap 'dr. X : n' / 'Total Pasien' bukan pasien. Lab bisa berupa tabel per baris
(no, nama, satuan, nilai, rujukan): ambil nilainya saja. Tidak ada pasien -> {{"pasien":[]}}.
Laporan dibagi blok '### PASIEN n': tiap blok = 1 pasien, urutan jawaban sama dengan urutan blok.

LAPORAN:
{teks}"""


def bersihkan(teks):
    """Buang hal yang tidak perlu dikirim ke AI (hemat token)."""
    out = []
    for line in teks.splitlines():
        if re.match(r"\s*PACS\s*:", line, re.I) or line.strip().lower() == "document:":
            continue
        line = re.sub(r"https?://\S+", "", line)
        line = re.sub(r"[ \t]+", " ", line).strip()
        if line or (out and out[-1]):
            out.append(line)
    return "\n".join(out).strip()


def buat_prompt(blok, tgl, ruang="-"):
    """blok: list (nomor, teks pasien yang sudah dibersihkan)."""
    teks = "\n\n".join(f"### PASIEN {i}\n{t}" for i, t in blok)
    return PROMPT.format(tgl=tgl.strftime("%d/%m/%Y"), ruang=ruang or "-", teks=teks)


# baris identitas: "Lucia/Pr/43th/01017416", "TUKMAIDA ... / 47 / PR / 01018825", "ZUNAIDAH/ 63thn / pr / ..."
def baris_identitas(line, berikut=""):
    """Baris 'Nama/Lk/umur/RM'. Juga dikenali kalau tanpa Lk/Pr tapi diakhiri no RM
    ('RINA NOVALIA/01018848') atau baris berikutnya 'DPJP ...'."""
    f = [x.strip(" *").lower() for x in line.split("/")]
    if len(re.sub(r"[^a-z]", "", f[0])) < 3 or line.lstrip(" *").lower().startswith("dpjp"):
        return False
    if len(f) >= 3 and any(x in ("lk", "pr", "l", "p") for x in f[1:4]):
        return True
    if len(f) >= 2 and re.fullmatch(r"\d{6,8}", f[-1]):
        return True
    return bool(re.match(r"\s*dpjp\b", berikut, re.I)) and not re.match(r"\s*(dx|r/|i:|s:|o:)", line, re.I)


BARIS_BED = re.compile(r"^\*?\s*\d+(?:[.,]\d+)*\s*\*?$")  # "3.4.1", "6,05", "206"


def pecah_pasien(teks):
    """Teks satu ruangan -> list teks per pasien (sudah dibersihkan). Teks sebelum pasien pertama
    (rekap 'dr. X : n') dibuang. Kosong kalau tidak ada baris identitas yang dikenali."""
    lines = bersihkan(teks).splitlines()
    mulai = []
    for i, line in enumerate(lines):
        if baris_identitas(line, lines[i + 1] if i + 1 < len(lines) else ""):
            # nomor bed/ruangan tepat di atas identitas ikut blok pasien ini
            mulai.append(i - 1 if i and BARIS_BED.match(lines[i - 1]) else i)
    return ["\n".join(lines[a:b]).strip() for a, b in zip(mulai, mulai[1:] + [len(lines)])]


# judul ruangan: "RB 1 :", "RA 3", "PAVILIUN 5", "Pav 6", "Ruang 206"
HEADER_RUANG = re.compile(r"^\s*(RB|RA|PAV(?:ILIUN)?\.?|RUANG(?:AN)?)\s*(\d+)\s*:?\s*$", re.I)


def pecah_ruangan(teks):
    """Satu file bisa berisi beberapa ruangan -> [(ruang, teks)], supaya tiap panggilan AI kecil."""
    bagian, ruang, buf = [], "", []
    for line in teks.splitlines():
        m = HEADER_RUANG.match(line)
        if m:
            if "\n".join(buf).strip():
                bagian.append((ruang, "\n".join(buf)))
            ruang, buf = norm_ruang(f"{m.group(1)} {m.group(2)}"), []
        else:
            buf.append(line)
    if "\n".join(buf).strip():
        bagian.append((ruang, "\n".join(buf)))
    # buang bagian yang cuma rekap (tidak ada baris identitas pasien 'Nama/Lk/...')
    return [(r, t) for r, t in bagian if re.search(r"[A-Za-z]\s*/\s*(Lk|Pr|\d)", t, re.I)]


def ambil_json(s):
    s = s.strip()
    s = re.sub(r"^```(?:json)?|```$", "", s, flags=re.M).strip()
    return json.loads(s[s.find("{"): s.rfind("}") + 1])


def panggil_ai(prompt, model):
    try:
        import anthropic
    except ImportError:
        sys.exit("Install dulu: pip install anthropic")
    client = anthropic.Anthropic()
    msg = client.messages.create(model=model, max_tokens=16000,
                                 messages=[{"role": "user", "content": prompt}])
    u = msg.usage
    print(f"  token: input {u.input_tokens}, output {u.output_tokens}", file=sys.stderr)
    return "".join(b.text for b in msg.content if b.type == "text")


def panggil_cli(prompt, model):
    """Pakai Claude Code (`claude -p`) -> memakai kuota langganan Pro/Max, tanpa API key.
    Tools, plugin, hook, MCP dimatikan supaya token tambahan minimal."""
    exe = shutil.which("claude")
    if not exe:
        sys.exit("Perintah 'claude' tidak ditemukan. Install Claude Code lalu login: npm i -g @anthropic-ai/claude-code")
    cmd = [exe, "-p", "--model", model, "--tools", "", "--setting-sources", "",
           "--strict-mcp-config", "--no-session-persistence", "--output-format", "json",
           "--system-prompt", "Kamu konverter laporan medis ke JSON. Balas hanya JSON valid."]
    with tempfile.TemporaryDirectory() as tmp:  # folder kosong: tidak ada CLAUDE.md proyek yang ikut terbaca
        r = subprocess.run(cmd, input=prompt, capture_output=True, text=True, encoding="utf-8", cwd=tmp,
                           env={**os.environ, "MAX_THINKING_TOKENS": "0"})  # tanpa thinking = hemat kuota
    try:
        res = json.loads(r.stdout)
    except json.JSONDecodeError:
        sys.exit(f"claude gagal (exit {r.returncode}):\n{r.stdout[-2000:]}\n{r.stderr[-2000:]}")
    if res.get("is_error"):
        sys.exit(f"claude error: {res.get('result')}")
    u = res.get("usage", {})
    print(f"  token: input {u.get('input_tokens')}, output {u.get('output_tokens')}", file=sys.stderr)
    return res.get("result", "")


RELATIF = re.compile(r"\b(hari ini|besok|lusa|kemarin)\b", re.I)


def kunci_cache(blok, ruang, tgl, model):
    """Cache per pasien: teks pasien sama dengan hari sebelumnya -> tidak dikirim ulang ke AI.
    Tanggal hanya ikut kunci kalau ada kata relatif (besok/hari ini) tanpa tanggal di baris yang
    sama, karena artinya berubah tiap hari. 'PBJ hari ini (30/09)' tidak perlu."""
    versi = hashlib.sha1(PROMPT.encode()).hexdigest()[:8]  # prompt diubah -> cache lama tidak dipakai
    relatif = any(RELATIF.search(l) and not re.search(r"\d{1,2}/\d{1,2}", l) for l in blok.splitlines())
    tg = tgl.isoformat() if relatif else ""
    return hashlib.sha1("|".join([model, versi, ruang or "", tg, blok]).encode()).hexdigest()[:16]


def proses_bagian(nama, ruang, teks, tgl, model, engine):
    """Satu ruangan: pasien yang sudah ada di cache dipakai, sisanya dikirim ke AI sekaligus."""
    CACHE.mkdir(exist_ok=True)
    label = f"{nama}-{ruang or 'x'}".replace(" ", "")
    blok = pecah_pasien(teks) or [bersihkan(teks)]  # format tak dikenal -> kirim utuh
    hasil, kirim = {}, []
    for i, b in enumerate(blok, 1):
        cf = CACHE / f"p-{kunci_cache(b, ruang, tgl, model)}.json"
        if cf.exists():
            hasil[i] = json.loads(cf.read_text(encoding="utf-8"))
        else:
            kirim.append((i, b, cf))
    print(f"{label}: {len(blok)} blok pasien, {len(hasil)} dari cache, {len(kirim)} dikirim ke AI",
          file=sys.stderr)
    if kirim:
        prompt = buat_prompt([(i, b) for i, b, _ in kirim], tgl, ruang)
        jawaban = panggil_cli(prompt, model) if engine == "cli" else panggil_ai(prompt, model)
        pasien = ambil_json(jawaban).get("pasien", [])
        per_blok = {}
        if len(kirim) == 1:  # satu blok: semua jawaban milik blok itu
            per_blok[kirim[0][0]] = pasien
        else:
            for p in pasien:
                try:
                    per_blok.setdefault(int(p.get("i")), []).append(p)
                except (TypeError, ValueError):
                    print(f"{label}: pasien tanpa nomor blok diabaikan: {p.get('n')}", file=sys.stderr)
        for i, b, cf in kirim:
            if i in per_blok:
                hasil[i] = per_blok[i]
                cf.write_text(json.dumps(per_blok[i], ensure_ascii=False, indent=1), encoding="utf-8")
            else:  # tidak di-cache supaya dicoba lagi di run berikutnya
                print(f"{label}: AI tidak mengembalikan blok {i}: {b.splitlines()[0][:60]}", file=sys.stderr)
    return {"ruang": ruang, "pasien": [p for i in sorted(hasil) for p in hasil[i]]}


def kumpul_bagian(files):
    """Pecah semua file per ruangan -> [(nama_file, ruang, teks)].
    File belakang = koreksi: ruangan yang sama menggantikan ruangan di file depan
    (posisi tetap), mis `visite.py laporan.txt koreksi_rb6.txt`.
    Item files boleh path, atau (nama, teks) langsung (dari web)."""
    bagian = {}
    for path in files:
        if isinstance(path, tuple):
            nama, teks = path
        else:
            nama, teks = Path(path).stem, Path(path).read_text(encoding="utf-8", errors="replace")
        for i, (r, t) in enumerate(pecah_ruangan(teks) or [("", teks)]):
            key = r or f"{nama}#{i}"
            if key in bagian:
                print(f"{r}: diganti dengan versi {nama}", file=sys.stderr)
            bagian[key] = (nama, r, t)
    return list(bagian.values())


def proses_file(files, tgl, model, engine, paralel=4):
    """Kirim tiap ruangan ke AI paralel. Return list data per ruangan."""
    with ThreadPoolExecutor(max_workers=paralel) as ex:
        futs = [ex.submit(proses_bagian, n, r, t, tgl, model, engine) for n, r, t in kumpul_bagian(files)]
        return [f.result() for f in futs]


# ---------------------------------------------------------------- docx helpers
def shade(el_pr, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    el_pr.append(shd)


TOTAL_LEBAR = sum(w for _, w in KOLOM)


def cell_setup(cell, width=None, fill=None, span=None, mar=(20, 30), valign="top"):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("w:tcW", "w:gridSpan", "w:shd", "w:tcMar", "w:vAlign"):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
    if width:
        w = OxmlElement("w:tcW")
        # lebar relatif (pct, 1/50 persen) supaya tabel autofit ke lebar halaman
        w.set(qn("w:w"), str(round(width * 5000 / TOTAL_LEBAR)))
        w.set(qn("w:type"), "pct")
        tcPr.append(w)
    if span:
        g = OxmlElement("w:gridSpan")
        g.set(qn("w:val"), str(span))
        tcPr.append(g)
    if fill:
        shade(tcPr, fill)
    m = OxmlElement("w:tcMar")
    for side, val in (("top", mar[0]), ("left", mar[1]), ("bottom", mar[0]), ("right", mar[1])):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(val))
        e.set(qn("w:type"), "dxa")
        m.append(e)
    tcPr.append(m)
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), valign)
    tcPr.append(va)


def run(p, text, bold=False, color=None, size=9):
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    if color is not None:
        r.font.color.rgb = color
    return r


def para(container, after=15, before=15, align=None):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(after / 20)
    pf.space_before = Pt(before / 20)
    if align:
        p.alignment = align
    return p


def isi_cell(cell, lines):
    """lines: list of list-of-(text, bold, color). Paragraf pertama pakai yang sudah ada."""
    if not lines:
        lines = [[("-", False, None)]]
    first = True
    for segs in lines:
        if first:
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0.75)
            p.paragraph_format.space_before = Pt(0.75)
            first = False
        else:
            p = para(cell)
        for text, bold, color in segs:
            run(p, text, bold, color)


def plain(items):
    return [[(str(x), False, None)] for x in items if str(x).strip()]


def lab_line(line):
    segs = []
    parts = [x.strip() for x in line.lstrip("! ").split(",") if x.strip()]
    # gabung kembali angka yang terpecah koma desimal, mis 'Cr 0,94'
    for i, part in enumerate(parts):
        mt = re.match(r"^(.*?)\s+([<>]?\d[\d.,]*)\s*$", part)
        bad = bool(mt) and abnormal(mt.group(1), mt.group(2))
        if i:
            segs.append((", ", False, None))
        segs.append((part, False, MERAH if bad else None))
    return segs


def split_dx(dx):
    dx = re.sub(r"\s+", " ", dx or "").strip()
    if not dx:
        return []
    parts = re.split(r"\s+(?=\+\s)|\s+(?=d/t\s)", dx)
    return [x.strip() for x in parts if x.strip()]


def kol_objectives(p):
    lines = []
    lab = [x for x in p.get("lab", []) if str(x).strip()]
    if lab:
        # AI kadang sudah menulis judul 'Lab dd/mm: ...' sendiri -> jangan dobel judul
        if not re.match(r"\s*lab\b", lab[0], re.I):
            lines.append([(f"Lab {p.get('tl', '')}:".replace(" :", ":"), True, None)])
        lines += [lab_line(x) for x in lab]
    for x in p.get("pj", []):
        if ":" in x:
            a, b = x.split(":", 1)
            lines.append([(a + ": ", True, None), (b.strip(), False, None)])
        elif x.strip():
            lines.append([(x, False, None)])
    lines += [[(str(x).strip(), False, None)] for x in p.get("o", []) if str(x).strip()]
    return lines


def kol_lanjut(p, f):
    """Isi kolom; kalau diteruskan dari list sebelumnya (laporan hari ini kosong), beri keterangan tanggal."""
    lines = plain(p.get(f, []))
    tg = p.get("_lanjut", {}).get(f)
    return lines


def kol_masalah(p):
    out = []
    for x in p.get("m", []):
        x = str(x).strip()
        if x:
            red = x.startswith("!")
            out.append([(x.lstrip("! "), False, MERAH if red else None)])
    return out


def kol_identitas(p):
    lines = [[(p.get("n", "-"), True, None)]]
    info = " / ".join(x for x in [p.get("jk", ""), f"{p['u']} th" if p.get("u") else "",
                                  f"RM {p['rm']}" if p.get("rm") else ""] if x)
    if info:
        lines.append([(info, False, None)])
    if p.get("to"):
        lines.append([(f"Op: {p['to']}", False, None)])
    tag = " — ".join(x for x in ["NEW" if p.get("baru") else "", p.get("st", "")] if x)
    if tag:
        lines.append([(tag, True, MERAH)])
    return lines


def tabel(doc, judul, pasien):
    t = doc.add_table(rows=2 + len(pasien), cols=len(KOLOM))
    tblPr = t._tbl.tblPr
    # AutoFit to Window: lebar tabel 100% halaman, layout autofit
    e = OxmlElement("w:tblW")
    e.set(qn("w:w"), "5000")
    e.set(qn("w:type"), "pct")
    tblPr.append(e)
    e = OxmlElement("w:tblLayout")
    e.set(qn("w:type"), "autofit")
    tblPr.append(e)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "auto")
        borders.append(b)
    tblPr.append(borders)
    for gc, (_, w) in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), KOLOM):
        gc.set(qn("w:w"), str(w))

    # baris judul ruangan
    c = t.rows[0].cells[0].merge(t.rows[0].cells[-1])
    cell_setup(c, sum(w for _, w in KOLOM), "2E74B5", span=len(KOLOM), mar=(40, 60), valign="center")
    run(c.paragraphs[0], judul, True, PUTIH, 10)
    # header kolom
    for cell, (nama, w) in zip(t.rows[1].cells, KOLOM):
        cell_setup(cell, w, "4472C4", valign="center")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(cell.paragraphs[0], nama, True, PUTIH, 9)
    # repeat header di halaman berikut
    for r in t.rows[:2]:
        trPr = r._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:tblHeader"))

    for row, p in zip(t.rows[2:], pasien):
        isi = [
            [[(p.get("k", "-"), False, None)]],
            kol_identitas(p),
            plain(split_dx(p.get("dx", ""))),
            kol_objectives(p),
            kol_masalah(p),
            kol_lanjut(p, "d"),
            kol_lanjut(p, "tx"),
            plain(p.get("p", [])),
        ]
        for cell, (_, w), lines in zip(row.cells, KOLOM, isi):
            cell_setup(cell, w)
            isi_cell(cell, lines)
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
    return t


# ---------------------------------------------------------------- susun dokumen
def cari_dokter(dpjp):
    s = (dpjp or "").lower()
    for key, full, short in DOKTER:
        if re.search(rf"\b{key}\b", s):
            return key, full, short
    nama = dpjp.strip() or "DPJP tidak diketahui"
    return nama.lower(), nama, nama


def norm_ruang(r):
    r = re.sub(r"\s+", " ", str(r or "").strip().rstrip(":")).strip()
    m = re.fullmatch(r"(RB|RA)\s*(\d+)", r, re.I)
    if m:
        return f"{m.group(1).upper()}{m.group(2)}"
    m = re.fullmatch(r"PAV(?:ILIUN|ILION)?\.?\s*(\d+)", r, re.I)
    if m:
        return f"Pav {m.group(1)}"
    m = re.fullmatch(r"(?:RUANG(?:AN)?\s*)?(\d+)", r, re.I)
    if m:
        return f"Ruang {m.group(1)}"
    return r


def ruang_pasien(p, default):
    """Pakai 'r' dari AI hanya kalau memang nama ruangan (AI kadang mengisi nomor bed)."""
    r = str(p.get("r") or "").strip()
    if not r or r == str(p.get("k", "")).strip():
        return default
    if re.fullmatch(r"(RB|RA|PAV(ILIUN)?\.?|RUANG(AN)?)\s*\d+|\d{3}", r, re.I):
        return norm_ruang(r)
    if not p.get("k"):
        p["k"] = r  # ternyata nomor bed
    return default


def parse_tgl(s):
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime((s or "").strip(), fmt).date()
        except ValueError:
            pass
    return None


def tgl_pbj_lewat(p, tgl):
    """Tanggal PBJ yang sudah lewat (pasien sudah pulang tapi masih tersalin di laporan), atau None."""
    for s in [p.get("st", "")] + list(p.get("p", [])):
        m = re.search(r"pbj\D{0,20}?(\d{1,2})/(\d{1,2})", str(s), re.I)
        if not m:
            continue
        try:
            d = dt.date(tgl.year, int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        if d.month == 12 and tgl.month == 1:
            d = d.replace(year=tgl.year - 1)
        if d < tgl:
            return d
    return None


URUTAN_RUANG = ["RB1", "Ruang 206", "RA6", "Pav 5", "RB2", "RB3", "RB6", "Pav 6"]  # urutan list asli

# batas nilai kritis untuk catatan akhir: nama lab -> (bawah, atas)
KRITIS = {"hb": (7, None), "leu": (3000, None), "plt": (100000, None), "na": (125, 155),
          "k": (3.0, 6.0), "alb": (2.2, None), "gdp": (None, 300), "gds": (None, 300),
          "gd": (None, 300), "bil total": (None, 3)}


def nilai_lab(nama, s):
    s = s.strip()
    if nama in ("leu", "plt"):
        v = float(re.sub(r"[.,]", "", s)) if re.fullmatch(r"\d{1,3}([.,]\d{3})+", s) else float(s.replace(",", "."))
        return v * 1000 if v < 1000 else v
    return float(s.replace(",", "."))


def hal_penting(p, tgl):
    """Hal yang perlu disorot di catatan: POD 0-2, lab kritis, rencana emergency/transfusi."""
    out, sudah = [], set()
    to = [d for d in (parse_tgl(x) for x in re.findall(r"\d{1,2}/\d{1,2}/\d{2,4}", str(p.get("to") or ""))) if d]
    pro_op = re.search(r"\bop\b", p.get("st", ""), re.I)  # pro op hari ini sudah disebut sendiri
    if to and 0 <= (tgl - max(to)).days <= 2 and not pro_op:
        op = re.sub(r"\s*\(?POD-?\s*\d+\)?", "", re.split(r"\s+d/t\s+", str(p.get("dx", "")), maxsplit=1)[0]).strip()
        out.append(f"{op} POD-{(tgl - max(to)).days}" if op else f"POD-{(tgl - max(to)).days}")
    for baris in p.get("lab", []):
        for nama, s in re.findall(r"([A-Za-z][A-Za-z ]*?)\s+(\d+(?:[.,]\d+)*)", str(baris)):
            key = nama.strip().lower()
            if key not in KRITIS or key in sudah:  # lab terbaru saja (baris pertama)
                continue
            sudah.add(key)
            try:
                v = nilai_lab(key, s)
            except ValueError:
                continue
            lo, hi = KRITIS[key]
            if (lo is not None and v < lo) or (hi is not None and v > hi):
                out.append(f"{nama.strip()} {s}")
    out += [str(x) for x in p.get("p", []) if re.search(r"emergency|cito|transfusi", str(x), re.I)
            and not re.search(r"selesai|post transfusi|cek", str(x), re.I)]
    return out


def norm_status(st, tgl):
    """AI menyalin status + tanggal; kode yang menentukan apakah itu hari ini."""
    st = (st or "").strip()
    if re.search(r"besok pulang|pulang hari ini|rencana pulang", st, re.I):
        return "PBJ hari ini"  # catatan 'Post HD besok pulang' di laporan = pulang hari ini
    m = re.search(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?", st)
    if not m:
        if "hari ini" not in st.lower():
            return ""
        m = re.match(r"(\d{1,2})/(\d{1,2})", f"{tgl.day}/{tgl.month}")
    try:
        d = dt.date(tgl.year, int(m.group(2)), int(m.group(1)))
    except ValueError:
        return ""
    if d != tgl:
        return ""  # jadwal lama / yang akan datang tetap terlihat di kolom Planning
    tindakan = st.split(":", 1)[1].strip() if ":" in st else ""
    if "pbj" in st.lower():
        return "PBJ hari ini"
    if re.search(r"\b(CT|MRI|USG|foto|rontgen|x-?ray|lab)\b", tindakan, re.I):
        return ""  # jadwal pemeriksaan penunjang, bukan operasi
    return "Pro Op hari ini" + (f": {tindakan}" if tindakan else "")


def id_pasien(p):
    """Kunci pasien antar hari: no RM (tanpa nol depan), kalau kosong pakai nama."""
    rm = re.sub(r"\D", "", str(p.get("rm") or "")).lstrip("0")
    return f"rm:{rm}" if rm else "n:" + re.sub(r"[^a-z]", "", str(p.get("n") or "").lower())


LANJUT = {"tx": "Tatalaksana", "d": "Diet"}  # kolom yang diteruskan dari list sebelumnya bila kosong
DETAIL = ["jk", "u", "tm", "to", "st", "dx", "tl", "o", "lab", "pj", "m", "p"]  # ikut disimpan di riwayat


def simpan_teks(tgl, teks):
    """Simpan teks laporan mentah hari itu: [(nama, teks)] -> riwayat/teks/YYYY-MM-DD.json."""
    d = RIWAYAT / "teks"
    d.mkdir(parents=True, exist_ok=True)
    (d / f"{tgl.isoformat()}.json").write_text(
        json.dumps([{"nama": n, "teks": t} for n, t in teks], ensure_ascii=False, indent=1), encoding="utf-8")


def isi(xs):
    return [str(x).strip() for x in xs or [] if str(x).strip() not in ("", "-")]


def snapshot(semua):
    kini = {}
    for r, p in semua:
        v = {"n": p.get("n", ""), "rm": p.get("rm", ""), "r": r, "k": p.get("k", ""), "dpjp": p.get("dpjp", "")}
        for f in DETAIL:  # untuk halaman riwayat pasien di web
            if p.get(f):
                v[f] = p[f]
        for f in LANJUT:
            if isi(p.get(f)):
                v[f], v[f + "_tgl"] = isi(p[f]), p.get("_lanjut", {}).get(f, "")
        kini.setdefault(id_pasien(p), v)
    return kini


def simpan_snapshot(semua, tgl):
    RIWAYAT.mkdir(exist_ok=True)
    (RIWAYAT / f"{tgl.isoformat()}.json").write_text(
        json.dumps(snapshot(semua), ensure_ascii=False, indent=1), encoding="utf-8")


def banding_kemarin(semua, tgl):
    """Bandingkan dengan snapshot terakhir sebelum tgl di riwayat/.
    Kolom LANJUT yang kosong hari ini diisi dari list sebelumnya (ditandai p['_lanjut']).
    Return (tanggal_banding, baru, keluar, pindah) atau None kalau belum ada riwayat."""
    kini = snapshot(semua)
    lama = sorted(f for f in RIWAYAT.glob("*.json") if f.stem < tgl.isoformat()) if RIWAYAT.exists() else []
    if not lama:
        return None
    dulu = json.loads(lama[-1].read_text(encoding="utf-8"))
    tgl_lama = dt.date.fromisoformat(lama[-1].stem)
    pasangan = cocokkan(kini, dulu)
    for _, p in semua:
        prev = dulu.get(pasangan.get(id_pasien(p)), {})
        if p.get("_dpjp_warisan") and prev.get("dpjp"):
            p["dpjp"] = prev["dpjp"]
        for f in LANJUT:
            if not isi(p.get(f)) and prev.get(f):
                p[f] = prev[f]
                p.setdefault("_lanjut", {})[f] = prev.get(f + "_tgl") or f"{tgl_lama:%d/%m}"
    baru = [k for k in kini if k not in pasangan]
    lama_cocok = set(pasangan.values())
    keluar = [v for k, v in dulu.items() if k not in lama_cocok]
    # pindah = ruangan DAN bed berubah (bed sama = cuma beda penamaan ruangan, mis 'RA3' vs 'Ruang 206')
    pindah = [(v, dulu[pasangan[k]]["r"]) for k, v in kini.items()
              if k in pasangan and dulu[pasangan[k]]["r"] != v["r"]
              and beda_bed(dulu[pasangan[k]].get("k"), v.get("k"))]
    return tgl_lama, set(baru), keluar, pindah


def beda_bed(a, b):
    """Bed kosong = tidak diketahui -> jangan dianggap pindah."""
    a, b = str(a or "").strip(" ."), str(b or "").strip(" .")
    return bool(a and b and a != b)


def kata_nama(n):
    return re.findall(r"[a-z]{2,}", str(n or "").lower())


def cocokkan(kini, dulu):
    """{id_kini: id_dulu}. Utama: no RM. Cadangan: nama (RM bisa salah ketik / dokumen lama
    dibuat manual): kata pertama nama sama; kalau lebih dari satu kandidat, dua kata pertama."""
    pasangan = {k: k for k in kini if k in dulu}
    sisa = {k: v for k, v in dulu.items() if k not in pasangan}
    for k, v in kini.items():
        if k in pasangan:
            continue
        kn = kata_nama(v["n"])
        for panjang in (1, 2):
            cand = [d for d, w in sisa.items() if kn and kata_nama(w["n"])[:panjang] == kn[:panjang]]
            if len(cand) == 1:
                pasangan[k] = cand[0]
                del sisa[cand[0]]
                break
            if not cand:
                break
    return pasangan


def impor_docx(path, tgl):
    """Jadikan dokumen list pasien lama (docx, format tabel aplikasi ini) sebagai riwayat tgl,
    supaya run hari berikutnya bisa dibandingkan dengannya."""
    snap = {}
    for t in Document(path).tables:
        judul = t.rows[0].cells[0].text
        if "|" not in judul:
            continue
        ruang, dok = [x.strip() for x in judul.split("|", 1)]
        dpjp = dok.rsplit(":", 1)[0].strip()
        for row in t.rows[2:]:
            c = row.cells
            ident = [p.text.strip() for p in c[1].paragraphs if p.text.strip()]
            if not ident:
                continue
            rm = re.search(r"RM\s*(\d+)", " ".join(ident))
            p = {"n": ident[0], "rm": rm.group(1) if rm else "", "k": c[0].text.strip()}
            v = {**p, "r": norm_ruang(ruang), "dpjp": dpjp}
            for f, kol in (("d", 5), ("tx", 6)):
                xs = isi(x.text for x in c[kol].paragraphs)
                if xs:
                    v[f] = xs
            snap[id_pasien(p)] = v
    RIWAYAT.mkdir(exist_ok=True)
    out = RIWAYAT / f"{tgl.isoformat()}.json"
    out.write_text(json.dumps(snap, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Riwayat {tgl:%d/%m/%Y}: {len(snap)} pasien dari {path} -> {out}", file=sys.stderr)


def buat_docx(data_list, tgl, out, hari_baru=2):
    doc = Document(TEMPLATE)
    body = doc.element.body
    for el in list(body):
        if el.tag != qn("w:sectPr"):
            body.remove(el)

    semua = []  # (ruang, pasien)
    sudah_pulang = []  # PBJ tanggal lalu -> tidak dimasukkan list
    for data in data_list:
        ruang = norm_ruang(data.get("ruang", "?"))
        dpjp_atas = ""
        for p in data.get("pasien", []):
            # DPJP kosong -> ikut DPJP pasien terdekat di atasnya (urutan sesuai laporan)
            if str(p.get("dpjp") or "").strip():
                dpjp_atas = p["dpjp"]
            else:
                p["dpjp"] = dpjp_atas
                p["_dpjp_warisan"] = True  # diganti DPJP list sebelumnya bila ada
            d_pbj = tgl_pbj_lewat(p, tgl)
            if d_pbj:
                sudah_pulang.append(f"{p.get('n')} ({ruang_pasien(p, ruang)}, PBJ {d_pbj:%d/%m})")
                continue
            # cadangan: AI kadang lupa isi st tapi PBJ tertulis di planning
            kandidat = [p.get("st", "")] + [x for x in p.get("p", []) if re.search(r"pbj|pulang", str(x), re.I)]
            p["st"] = next((s for s in (norm_status(k, tgl) for k in kandidat) if s), "")
            semua.append((ruang_pasien(p, ruang), p))
    ruang_order = list(dict.fromkeys(r for r, _ in semua))
    ruang_order.sort(key=lambda r: URUTAN_RUANG.index(r) if r in URUTAN_RUANG else len(URUTAN_RUANG))
    banding = banding_kemarin(semua, tgl)
    simpan_snapshot(semua, tgl)
    if banding:
        for _, p in semua:
            p["baru"] = id_pasien(p) in banding[1]

    # header
    def judul(text, size, color=NAVY, bold=True, after=10):
        p = para(doc, after=after, before=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        run(p, text, bold, color, size)
    judul("LIST PASIEN PER DPJP", 14)
    judul("DIVISI BEDAH DIGESTIF — DEPARTEMEN ILMU BEDAH", 11)
    judul("RSUP H. ADAM MALIK MEDAN", 11)
    judul(f"{tgl.day} {BULAN[tgl.month - 1]} {tgl.year}", 10, None)
    rekap = " | ".join(f"{r}: {sum(1 for x, _ in semua if x == r)}" for r in ruang_order)
    judul(f"Total Pasien: {len(semua)}  |  {rekap}", 9, None, False)

    pbj = [f"{p.get('n')} ({r})" for r, p in semua if "pbj" in p.get("st", "").lower()]
    proop = [(r, p) for r, p in semua if re.search(r"\bop\b", p.get("st", ""), re.I)]
    info = []
    if pbj:
        info.append("PBJ hari ini: " + ", ".join(pbj))
    if proop:
        info.append("Pro Op: " + ", ".join(
            f"{p.get('n')} ({p['st'].split(':', 1)[1].strip() if ':' in p['st'] else r})" for r, p in proop))
    if info:
        judul("  |  ".join(info), 8, None, False)
    para(doc, 0, 0)

    # kelompok per DPJP
    grup = {}
    for r, p in semua:
        key, full, short = cari_dokter(p.get("dpjp", ""))
        grup.setdefault(key, {"full": full, "short": short, "items": []})["items"].append((r, p))
    order = [k for k, *_ in DOKTER] + [k for k in grup if k not in {d[0] for d in DOKTER}]
    no = 0
    for key in order:
        if key not in grup:
            continue
        g = grup[key]
        no += 1
        p = para(doc, after=40, before=80)
        pPr = p._p.get_or_add_pPr()
        shade(pPr, "1F3864")
        p.paragraph_format.left_indent = Pt(6)
        p.paragraph_format.right_indent = Pt(6)
        run(p, f"{no}. {g['full']} — {len(g['items'])} Pasien", True, PUTIH, 12)
        for ruang in ruang_order:
            ps = [x for r, x in g["items"] if r == ruang]
            if ps:
                tabel(doc, f"{ruang}  |  {g['short']}: {len(ps)}", ps)
                para(doc, 0, 0)

    # catatan (format list asli: pasien baru, PBJ, pro op, lalu hal penting/kritis saja)
    def tindakan(x):
        return x["st"].split(":", 1)[1].strip() if ":" in x.get("st", "") else ""

    baru = []
    for r, p in semua:
        tm = parse_tgl(p.get("tm"))
        # ada riwayat: baru = tidak ada di list sebelumnya; belum ada: pakai tgl masuk
        if p.get("baru") if banding else (tm and 0 <= (tgl - tm).days <= hari_baru):
            baru.append(f"{p.get('n')} ({r}, pro {tindakan(p)})" if tindakan(p) else f"{p.get('n')} ({r})")
    p = para(doc, after=40, before=40)
    run(p, "Catatan: ", True)
    if baru:
        run(p, "Pasien baru: " + ", ".join(baru) + ". ")
    if pbj:
        run(p, "PBJ hari ini: " + ", ".join(pbj) + ". ")
    if proop:
        run(p, "Pro Op hari ini: " + ", ".join(f"{p_.get('n')} ({r}, {tindakan(p_)})" if tindakan(p_)
                                               else f"{p_.get('n')} ({r})" for r, p_ in proop) + ". ")
    for r, x in semua:
        hal = hal_penting(x, tgl)
        if hal:
            run(p, f"{x.get('n')}: {', '.join(hal)}. ", True, MERAH)

    # info pembanding cukup di terminal, tidak masuk dokumen
    keluar = [v for v in (banding[2] if banding else [])
              if not any(s.startswith(f"{v['n']} (") for s in sudah_pulang)]
    if keluar:
        print("Keluar dari list (pulang/pindah divisi): "
              + ", ".join(f"{v['n']} ({v['r']})" for v in keluar), file=sys.stderr)
    if banding and banding[3]:
        print("Pindah ruangan: " + ", ".join(f"{v['n']} ({r0} → {v['r']})" for v, r0 in banding[3]),
              file=sys.stderr)
    if sudah_pulang:
        print("Tidak dimasukkan (PBJ sudah lewat, cek bila ternyata belum pulang): "
              + ", ".join(sudah_pulang), file=sys.stderr)

    doc.save(out)
    print(f"Tersimpan: {out}  ({len(semua)} pasien)", file=sys.stderr)
    info = {"total": len(semua), "rekap": rekap, "catatan": p.text, "file": Path(out).name,
            "dibuat": dt.datetime.now().isoformat(timespec="seconds"),
            "keluar": [f"{v['n']} ({v['r']})" for v in keluar],
            "pindah": [f"{v['n']} ({r0} → {v['r']})" for v, r0 in (banding[3] if banding else [])],
            "sudah_pulang": sudah_pulang}
    (RIWAYAT / "meta").mkdir(parents=True, exist_ok=True)
    (RIWAYAT / "meta" / f"{tgl.isoformat()}.json").write_text(
        json.dumps(info, ensure_ascii=False, indent=1), encoding="utf-8")
    return info


# ---------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(description="Buat list pasien visite (docx) dari laporan bangsal.")
    ap.add_argument("files", nargs="*", help="file .txt laporan; file berikutnya = koreksi (ruangan sama diganti)")
    ap.add_argument("-o", "--out", default=None, help="nama file docx output")
    ap.add_argument("--tanggal", help="tanggal laporan dd/mm/yyyy (default hari ini)")
    ap.add_argument("--engine", choices=["cli", "api"],
                    default="api" if os.environ.get("ANTHROPIC_API_KEY") else "cli",
                    help="cli = Claude Code (langganan Pro), api = Anthropic API (API key)")
    ap.add_argument("--model", default=os.environ.get("VISITE_MODEL"),
                    help="default: haiku (cli) / claude-haiku-4-5 (api); bisa 'sonnet' kalau hasil kurang rapi")
    ap.add_argument("--prompt-only", action="store_true", help="cetak prompt saja (untuk dipaste ke claude.ai)")
    ap.add_argument("--from-json", nargs="+", help="render dari file JSON (hasil AI) tanpa memanggil API")
    ap.add_argument("--impor-docx", metavar="DOCX",
                    help="jadikan list docx lama sebagai riwayat --tanggal (pembanding hari berikutnya)")
    a = ap.parse_args()

    tgl = parse_tgl(a.tanggal) if a.tanggal else dt.date.today()
    if a.impor_docx:
        if not a.tanggal:
            ap.error("--impor-docx butuh --tanggal (tanggal dokumen itu)")
        impor_docx(a.impor_docx, tgl)
        return
    if a.prompt_only:
        sys.stdout.reconfigure(encoding="utf-8")
        for _, r, t in kumpul_bagian(a.files):
            blok = pecah_pasien(t) or [bersihkan(t)]
            print(buat_prompt(list(enumerate(blok, 1)), tgl, r), end="\n\n=====\n\n")
        return
    if a.from_json:
        data = [ambil_json(Path(f).read_text(encoding="utf-8")) for f in a.from_json]
    elif a.files:
        model = a.model or ("haiku" if a.engine == "cli" else DEFAULT_MODEL)
        data = proses_file(a.files, tgl, model, a.engine)
        simpan_teks(tgl, [(Path(f).stem, Path(f).read_text(encoding="utf-8", errors="replace")) for f in a.files])
    else:
        ap.error("butuh file .txt atau --from-json")
    out = a.out or f"/output/List Pasien {tgl.strftime('%d-%m-%Y')}.docx"
    buat_docx(data, tgl, out)


if __name__ == "__main__":
    main()
