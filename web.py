"""
Dashboard web lokal untuk visite.py.

  python web.py            -> buka http://127.0.0.1:8000

Halaman:
  /              input laporan (tempel teks), cek data baru, lalu proses jadi docx
  /riwayat       daftar laporan per hari  -> /riwayat/YYYY-MM-DD
  /pasien        daftar semua pasien      -> /pasien/<id> (riwayat per hari pasien itu)
  /ruangan       ruangan yang dikenal sistem (tambah/hapus/urutkan)
  /edit/YYYY-MM-DD   preview & edit data hasil AI, lalu buat ulang docx

Hanya mendengar di 127.0.0.1 (data pasien tidak keluar dari komputer ini).
Tanpa library tambahan: cukup Python standar + dependensi visite.py.
"""
import datetime as dt
import html
import io
import json
import os
import re
import sys
import threading
import traceback
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

import visite

HOST, PORT = "127.0.0.1", int(os.environ.get("VISITE_PORT", 8000))
ENGINE = "api" if os.environ.get("ANTHROPIC_API_KEY") else "cli"
MODEL = os.environ.get("VISITE_MODEL") or ("haiku" if ENGINE == "cli" else visite.DEFAULT_MODEL)
E = html.escape


# ---------------------------------------------------------------- data riwayat
def tgl_hari():
    """Tanggal yang punya riwayat, terbaru dulu."""
    if not visite.RIWAYAT.exists():
        return []
    return sorted((f.stem for f in visite.RIWAYAT.glob("*.json")), reverse=True)


def baca(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default


def snapshot(tgl):
    return baca(visite.RIWAYAT / f"{tgl}.json", {})


def meta(tgl):
    return baca(visite.RIWAYAT / "meta" / f"{tgl}.json", {})


def teks_hari(tgl):
    return baca(visite.RIWAYAT / "teks" / f"{tgl}.json", [])


def data_hari(tgl):
    """Hasil AI mentah hari itu (None kalau list dibuat sebelum fitur edit)."""
    return baca(visite.RIWAYAT / "data" / f"{tgl}.json")


def statistik_ruang():
    """{ruang: {hari, akhir, n}} dari riwayat: berapa hari muncul, tanggal terakhir, jumlah pasien saat itu."""
    st = {}
    for tgl in sorted(tgl_hari()):
        hit = {}
        for v in snapshot(tgl).values():
            hit[v.get("r") or "?"] = hit.get(v.get("r") or "?", 0) + 1
        for r, n in hit.items():
            s = st.setdefault(r, {"hari": 0})
            s.update(hari=s["hari"] + 1, akhir=tgl, n=n)
    return st


def ubah_ruangan(aksi, nama):
    xs = visite.daftar_ruangan()
    nama = re.sub(r"\s+", " ", str(nama or "")).strip(" :*")
    kunci = visite.kunci_ruang(nama)
    idx = next((i for i, x in enumerate(xs) if visite.kunci_ruang(x) == kunci), None)
    if aksi == "tambah":
        if not kunci or len(nama) > 30 or "/" in nama:
            raise ValueError("Nama ruangan tidak valid.")
        if kunci.isdigit():
            raise ValueError("Nama ruangan tidak boleh angka saja (tertukar dengan nomor bed). Tulis mis. 'Ruang 206'.")
        if idx is not None:
            raise ValueError(f"'{xs[idx]}' sudah ada di daftar.")
        xs.append(visite.norm_ruang(nama))
    elif idx is None:
        raise ValueError(f"Ruangan '{nama}' tidak ada di daftar.")
    elif aksi == "hapus":
        xs.pop(idx)
    else:
        raise ValueError("Aksi tidak dikenal.")
    visite.simpan_ruangan(xs)
    return xs


TEKS_F = ("r", "k", "n", "jk", "u", "rm", "dpjp", "tm", "to", "st", "dx", "tl")
LIST_F = ("o", "lab", "pj", "m", "d", "tx", "p")


def bersih_data(data):
    """Data hasil edit dari browser -> struktur proses_file yang rapi (string / list string)."""
    if not isinstance(data, list):
        raise ValueError("Data tidak valid.")
    out = []
    for d in data:
        ps = []
        for p in (d.get("pasien") if isinstance(d, dict) else None) or []:
            if not isinstance(p, dict):
                continue
            x = {f: str(p[f]).strip() for f in TEKS_F if str(p.get(f) or "").strip()}
            for f in LIST_F:
                v = p.get(f) or []
                x[f] = [str(i).strip() for i in (v if isinstance(v, list) else [v]) if str(i).strip()]
            if p.get("_c"):
                x["_c"] = str(p["_c"])
            ps.append(x)
        out.append({"ruang": str(d.get("ruang") or ""), "pasien": ps})
    return out


def registri():
    """Satukan pasien lintas hari (RM bisa beda tulis antar hari -> pakai visite.cocokkan).
    Return {kunci: [(tgl, data), ...]} urut tanggal, dan {(tgl, id): kunci}."""
    orang, peta, prev, prev_peta = {}, {}, {}, {}
    for tgl in sorted(tgl_hari()):
        snap = snapshot(tgl)
        pas = visite.cocokkan(snap, prev) if prev else {}
        cur = {}
        for i, v in snap.items():
            kunci = prev_peta.get(pas.get(i)) or i
            cur[i] = kunci
            peta[(tgl, i)] = kunci
            orang.setdefault(kunci, []).append((tgl, v))
        prev, prev_peta = snap, cur
    return orang, peta


# ---------------------------------------------------------------- cek data sebelum proses
def identitas(blok):
    """(nama, rm, bed) dari teks mentah satu pasien, tanpa AI."""
    lines = blok.splitlines()
    for j, l in enumerate(lines):
        if visite.baris_identitas(l, lines[j + 1] if j + 1 < len(lines) else ""):
            nama = re.split(r"[,\[(]", re.sub(r"^[\d.,\s*]+", "", l.split("/")[0]))[0].strip(" *")
            rm = re.findall(r"\b\d{6,8}\b", l)
            bed = lines[j - 1].strip(" *") if j and visite.BARIS_BED.match(lines[j - 1]) else ""
            return nama.title(), (rm[-1] if rm else ""), bed
    return (lines[0][:40] if lines else "?"), "", ""


def daftar_blok(teks_list):
    """[(nama, teks)] -> list pasien mentah {ruang, n, rm, k, blok, sumber}."""
    out = []
    for sumber, r, t in visite.kumpul_bagian(teks_list):
        for b in visite.pecah_pasien(t) or [visite.bersihkan(t)]:
            n, rm, k = identitas(b)
            out.append({"ruang": visite.norm_ruang(r) if r else "?", "n": n, "rm": rm, "k": k,
                        "blok": b, "sumber": sumber, "r_asli": r})
    return out


def norm_baris(s):
    return re.sub(r"\s+", " ", s).strip().lower()


def analisa(teks_list, tgl):
    """Penjelasan sebelum diproses: ruangan, pasien baru/keluar, baris data baru per pasien,
    pasien yang dikirim ke AI vs dari cache."""
    kini = daftar_blok(teks_list)
    iso = tgl.isoformat()
    sebelum = [d for d in tgl_hari() if d < iso]
    snap_lama = snapshot(sebelum[0]) if sebelum else {}
    tgl_lama = sebelum[0] if sebelum else None

    # teks mentah hari sebelumnya (kalau tersimpan) untuk diff per pasien
    teks_lama = []
    for d in sorted(sebelum, reverse=True):
        teks_lama = teks_hari(d)
        if teks_lama:
            tgl_teks = d
            break
    lama_blok = daftar_blok([(x["nama"], x["teks"]) for x in teks_lama]) if teks_lama else []

    def kunci(p, i):
        return visite.id_pasien({"rm": p["rm"], "n": p["n"]}) + ("" if i == 0 else f"#{i}")

    def peta(lst):
        d = {}
        for p in lst:
            i, k = 0, kunci(p, 0)
            while k in d:
                i += 1
                k = kunci(p, i)
            d[k] = p
        return d

    d_kini = peta(kini)
    pas_snap = visite.cocokkan({k: {"n": v["n"]} for k, v in d_kini.items()}, snap_lama) if snap_lama else {}
    d_lama = peta(lama_blok)
    pas_teks = visite.cocokkan({k: {"n": v["n"]} for k, v in d_kini.items()},
                               {k: {"n": v["n"]} for k, v in d_lama.items()}) if d_lama else {}

    pasien = []
    for k, p in d_kini.items():
        cache = (visite.CACHE / f"p-{visite.kunci_cache(p['blok'], p['r_asli'], tgl, MODEL)}.json").exists()
        lama = d_lama.get(pas_teks.get(k))
        tambah, hapus = [], 0
        if lama:
            set_lama = {norm_baris(x) for x in lama["blok"].splitlines()}
            set_kini = {norm_baris(x) for x in p["blok"].splitlines()}
            tambah = [x.strip() for x in p["blok"].splitlines()
                      if x.strip() and norm_baris(x) not in set_lama]
            hapus = sum(1 for x in set_lama if x and x not in set_kini)
        pbj = [str(x) for x in p["blok"].splitlines()
               if "pbj" in x.lower() and visite.tgl_pbj_lewat({"st": x}, tgl)]
        baru = bool(snap_lama) and k not in pas_snap
        status = "baru" if baru else ("berubah" if tambah or hapus else ("sama" if lama else "?"))
        pasien.append({"ruang": p["ruang"], "k": p["k"], "n": p["n"], "rm": p["rm"], "sumber": p["sumber"],
                       "cache": cache, "status": status, "tambah": tambah, "hapus": hapus, "pbj_lewat": pbj,
                       "dpjp_lama": snap_lama.get(pas_snap.get(k), {}).get("dpjp", "")})
    cocok = set(pas_snap.values())
    keluar = [f"{v['n']} ({v.get('r', '')})" for i, v in snap_lama.items() if i not in cocok]
    ruang, kenal = {}, {visite.kunci_ruang(x) for x in visite.daftar_ruangan()}
    for p in sorted(pasien, key=lambda p: visite.urutan_ruang(p["ruang"])):
        ruang.setdefault(p["ruang"], {"n": 0, "sumber": p["sumber"],
                                      "baru": p["ruang"] != "?" and visite.kunci_ruang(p["ruang"]) not in kenal})["n"] += 1
    return {"tgl": iso, "tgl_lama": tgl_lama, "tgl_teks": tgl_teks if teks_lama else None,
            "total": len(pasien), "ke_ai": sum(1 for p in pasien if not p["cache"]),
            "ruang": ruang, "pasien": pasien, "keluar": keluar,
            "sudah_ada": (visite.RIWAYAT / f"{iso}.json").exists(), "model": f"{ENGINE}/{MODEL}"}


# ---------------------------------------------------------------- proses (thread latar)
JOBS = {}
KUNCI_JOB = threading.Lock()


class Tee(io.TextIOBase):
    """Salin output stderr visite.py ke log job (untuk ditampilkan di web)."""

    def __init__(self, job, asli):
        self.job, self.asli, self.buf = job, asli, ""

    def write(self, s):
        self.asli.write(s)
        self.buf += s
        *baris, self.buf = self.buf.split("\n")
        self.job["log"] += [b for b in baris if b.strip() and "[rtk]" not in b]
        return len(s)

    def flush(self):
        self.asli.flush()


def buat_docx(data, tgl):
    out = visite.BASE / f"output/List Pasien {tgl:%d-%m-%Y}.docx"
    try:
        return visite.buat_docx(data, tgl, out)
    except PermissionError:
        raise RuntimeError(f"'{out.name}' sedang dibuka (Word?). Tutup dulu lalu coba lagi.")


def jalankan(job, teks_list, tgl):
    if not KUNCI_JOB.acquire(blocking=False):
        job.update(status="gagal", error="Masih ada proses lain berjalan.")
        return
    asli = sys.stderr
    sys.stderr = Tee(job, asli)
    try:
        data = visite.proses_file(teks_list, tgl, MODEL, ENGINE)
        visite.simpan_teks(tgl, teks_list)
        visite.simpan_data(tgl, data)  # sebelum buat_docx (yang mengubah isi data)
        job["hasil"] = buat_docx(data, tgl)
        job["status"] = "selesai"
    except BaseException as e:  # sys.exit dari panggil_cli juga ditangkap
        job.update(status="gagal", error=str(e) or e.__class__.__name__)
        traceback.print_exc(file=asli)
    finally:
        sys.stderr = asli
        KUNCI_JOB.release()


# ---------------------------------------------------------------- HTML
CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1d2433;--mute:#6b7385;--line:#e3e6ec;--navy:#1f3864;
--acc:#2457c5;--red:#c62828;--redbg:#fdecec;--grn:#2e7d32;--grnbg:#e8f5e9;--amb:#a15c00;--ambbg:#fff4e0;--chip:#eef1f6}
@media (prefers-color-scheme:dark){:root{--bg:#12151b;--card:#1b2029;--ink:#e6e9ef;--mute:#98a1b3;--line:#2c3340;
--navy:#9db7ea;--acc:#7ea6ff;--red:#ff8a80;--redbg:#3a1f22;--grn:#81c784;--grnbg:#1e3324;--amb:#ffcc80;--ambbg:#3a2e1a;--chip:#262c37}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
header{background:var(--navy);color:#fff;padding:0 16px}@media (prefers-color-scheme:dark){header{background:#1f3864}}
header .in{max-width:1200px;margin:auto;display:flex;gap:20px;align-items:center;height:52px;overflow-x:auto}
header b{font-size:15px;white-space:nowrap}header a{color:#cfdcff;text-decoration:none;white-space:nowrap;padding:4px 0}
header a.on{color:#fff;border-bottom:2px solid #fff}
main{max-width:1200px;margin:auto;padding:20px 16px 60px}
h1{font-size:20px;margin:0 0 4px}h2{font-size:16px;margin:24px 0 8px}.mute{color:var(--mute)}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin:12px 0}
label{font-weight:600;display:block;margin:10px 0 4px}
textarea{width:100%;min-height:220px;font:12px/1.4 ui-monospace,Consolas,monospace;padding:10px;border:1px solid var(--line);
border-radius:8px;background:var(--bg);color:var(--ink);resize:vertical}
input[type=date],input[type=search]{padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit}
button,.btn{background:var(--acc);color:#fff;border:0;border-radius:8px;padding:9px 16px;font:inherit;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}
button.sec,.btn.sec{background:var(--chip);color:var(--ink)}button:disabled{opacity:.5;cursor:default}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px}
.stat b{display:block;font-size:22px}.stat span{color:var(--mute);font-size:12px}
table{width:100%;border-collapse:collapse;background:var(--card)}th,td{text-align:left;padding:7px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-size:12px;color:var(--mute);font-weight:600;background:var(--chip)}
.wrap{overflow-x:auto;border:1px solid var(--line);border-radius:10px}
.tag{display:inline-block;padding:1px 8px;border-radius:99px;font-size:11px;font-weight:700;white-space:nowrap}
.baru{background:var(--redbg);color:var(--red)}.berubah{background:var(--ambbg);color:var(--amb)}
.sama{background:var(--grnbg);color:var(--grn)}.q{background:var(--chip);color:var(--mute)}
.add{font:12px/1.45 ui-monospace,Consolas,monospace;color:var(--grn);margin:4px 0 0;padding-left:10px;border-left:3px solid var(--grn);white-space:pre-wrap}
.warn{background:var(--ambbg);color:var(--amb);border-radius:8px;padding:10px 12px;margin:8px 0}
.err{background:var(--redbg);color:var(--red);border-radius:8px;padding:10px 12px;margin:8px 0}
.log{font:12px/1.4 ui-monospace,Consolas,monospace;background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:10px;max-height:260px;overflow:auto;white-space:pre-wrap}
a{color:var(--acc)}ul{margin:4px 0;padding-left:20px}details summary{cursor:pointer;color:var(--acc)}
.red{color:var(--red);font-weight:600}.small{font-size:12px}
.tl{border-left:3px solid var(--line);padding-left:14px;margin-left:6px}
.tl .card{position:relative}.tl .card:before{content:"";position:absolute;left:-22px;top:20px;width:11px;height:11px;border-radius:50%;background:var(--acc)}
.grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:4px 20px}
input[type=text]{padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit}
body.lebar main{max-width:none;padding-left:8px;padding-right:8px}
.edt{min-width:1100px;table-layout:fixed}.edt td{padding:3px}
.edt tr>:nth-child(1){width:6%}.edt tr>:nth-child(2){width:15%}.edt tr>:nth-child(3){width:12%}
.edt tr>:nth-child(4){width:19%}.edt tr>:nth-child(5){width:11%}.edt tr>:nth-child(6){width:7%}
.edt tr>:nth-child(7),.edt tr>:nth-child(8){width:13%}.edt tr>:nth-child(9){width:4%;text-align:center;padding:3px 0}
.ed{display:block;width:100%;margin:1px 0;padding:3px 6px;border:1px solid var(--line);border-radius:5px;
background:var(--bg);color:var(--ink);font:12px/1.35 system-ui,-apple-system,"Segoe UI",sans-serif}
textarea.ed{min-height:0;resize:vertical;overflow:hidden}.ed:focus{outline:2px solid var(--acc);outline-offset:-1px}
.ed.ubah{border-color:var(--amb)}tr.dihapus td{opacity:.35}tr.pbaru td{background:var(--grnbg)}
button.mini{padding:3px 9px;font-size:12px;font-weight:600}
"""

# dipakai halaman Input (setelah proses) dan /edit/<tgl>
EDITOR_JS = r"""
function esc(s){return String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function post(u,b){const r=await fetch(u,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});
const j=await r.json();if(!r.ok)throw new Error(j.error||r.statusText);return j}
// beberapa field dalam satu textbox: teks() untuk ditampilkan, baca() mengisi field dari teks
const IDENT=[['n','Nama'],['jk','JK'],['u','Umur'],['rm','RM'],['dpjp','DPJP'],['tm','Masuk'],['to','Operasi'],['st','Status']];
const GABUNG={
 ident:{teks:p=>IDENT.map(([f,l])=>`${l}: ${p[f]??''}`).join('\n'),
  baca:(p,t)=>{const lab=Object.fromEntries(IDENT.map(([f,l])=>[l.toLowerCase(),f]));IDENT.forEach(([f])=>p[f]='');
   for(const x of t.split('\n')){const m=x.match(/^\s*([^:]+?)\s*:\s*(.*)$/),f=m&&lab[m[1].toLowerCase()];if(f)p[f]=m[2].trim()}}},
 // judul bagian = baris persis 'Lab dd/mm:' / 'Penunjang:' / 'Klinis:'; baris di bawahnya = isi bagian itu
 obj:{teks:p=>[`Lab ${p.tl||''}:`.replace(' :',':'),...(p.lab||[]),'Penunjang:',...(p.pj||[]),'Klinis:',...(p.o||[])].join('\n'),
  baca:(p,t)=>{p.tl='';p.lab=[];p.pj=[];p.o=[];let cur='lab';
   for(const x of t.split('\n')){const s=x.trim();if(!s)continue;const m=s.match(/^lab\s*([\d\/.-]*)\s*:$/i);
    if(m){cur='lab';p.tl=m[1];continue}
    if(/^penunjang\s*:$/i.test(s)){cur='pj';continue}
    if(/^klinis\s*:$/i.test(s)){cur='o';continue}
    p[cur].push(s)}}}};
// kolom sama dengan docx; tipe: '' input, area = teks panjang, list = 1 baris 1 item, gabung = GABUNG
const KOL_ED=[
 ['Kamar',[['k','bed'],['r','ruang','ruang']]],
 ['Identitas',[['ident','Status: PBJ dd/mm / Op dd/mm: tindakan','gabung']]],
 ['Diagnosis',[['dx','diagnosis','area']]],
 ['Objectives',[['obj','Lab: 1 baris = 1 item (mis Hb 10.0, Leu 8.460). Penunjang: mis Foto thorax: normal. Klinis: keluhan / tanda vital','gabung']]],
 ['Masalah',[['m','awali ! = merah','list']]],
 ['Diet',[['d','diet','list']]],
 ['Tatalaksana',[['tx','terapi (kosong = ikut list sebelumnya)','list']]],
 ['Planning',[['p','rencana','list']]]];
function edField(si,pi,p,[f,ph,tipe],ruang){
 const v=tipe==='gabung'?GABUNG[f].teks(p):tipe==='list'?(p[f]||[]).join('\n'):String(p[f]??'');
 if(f==='r')ph='ruang (kosong = '+ruang+')';
 const a=`class="ed" data-s="${si}" data-p="${pi}" data-f="${f}" data-t="${tipe||''}" placeholder="${esc(ph)}" title="${esc(ph)}"`;
 if(tipe==='list'||tipe==='area'||tipe==='gabung')return `<textarea ${a} rows="${Math.max(tipe==='area'?3:2,v.split('\n').length)}">${esc(v)}</textarea>`;
 return `<input ${a} value="${esc(v)}"${tipe==='ruang'?' list="dl-ruang"':''}>`}
function barisEd(si,pi,p,ruang){
 return `<tr${p._baru?' class="pbaru"':''}>${KOL_ED.map(k=>`<td>${k[1].map(x=>edField(si,pi,p,x,ruang)).join('')}</td>`).join('')}
 <td><input type="checkbox" class="hapus" data-s="${si}" data-p="${pi}"></td></tr>`}
function editor(el,tgl,data,ruangan){
 let h=`<datalist id="dl-ruang">${ruangan.map(r=>`<option value="${esc(r)}">`).join('')}</datalist>
 <p class="mute small">Klik sel untuk mengubah. Satu baris = satu item. Identitas: ubah teks setelah "Nama:", "RM:", dst.
 Objectives: isi di bawah judul <b>Lab dd/mm:</b>, <b>Penunjang:</b>, <b>Klinis:</b> (judulnya jangan dihapus).
 Centang <b>Hapus</b> untuk membuang pasien dari list. Pasien yang tidak terbaca AI bisa ditambah lewat tombol
 <b>+ Tambah pasien</b> di bawah tiap ruangan.</p>`;
 data.forEach((d,si)=>{
  h+=`<h2>${esc(d.ruang||'Tanpa ruangan')} <span class="mute small"><span class="jml">${d.pasien.length}</span> pasien</span></h2>
  <div class="wrap"><table class="edt"><tr>${KOL_ED.map(k=>`<th>${k[0]}</th>`).join('')}<th>Hapus</th></tr>`;
  d.pasien.forEach((p,pi)=>{h+=barisEd(si,pi,p,d.ruang||'-')});
  h+=`</table></div><p><button class="sec mini tambah" data-s="${si}">+ Tambah pasien di ${esc(d.ruang||'sini')}</button></p>`});
 h+=`<div class="card"><div class="row"><button class="bsimpan">Simpan &amp; buat ulang docx</button>
 <span class="mute small">Riwayat tanggal ${esc(tgl)} ditimpa. Hasil edit juga masuk cache: teks pasien yang sama besok memakai data ini.</span></div>
 <div class="hsimpan"></div></div>`;
 el.innerHTML=h;
 document.body.classList.add('lebar');  // editor selebar layar
 const tinggi=t=>{t.style.height='auto';t.style.height=(t.scrollHeight+2)+'px'};  // textarea memanjang sesuai isi
 el.querySelectorAll('textarea.ed').forEach(tinggi);
 addEventListener('resize',()=>el.querySelectorAll('textarea.ed').forEach(tinggi));
 el.addEventListener('click',e=>{const b=e.target.closest('.tambah');if(!b)return;
  const si=+b.dataset.s,d=data[si],pi=d.pasien.push({_baru:true})-1;
  const tb=el.querySelectorAll('table.edt')[si].tBodies[0];
  tb.insertAdjacentHTML('beforeend',barisEd(si,pi,d.pasien[pi],d.ruang||'-'));
  tb.lastElementChild.querySelectorAll('textarea.ed').forEach(tinggi);
  el.querySelectorAll('.jml')[si].textContent=d.pasien.length;
  tb.lastElementChild.querySelector('[data-f="ident"]').focus()});
 el.addEventListener('input',e=>{const t=e.target;if(!t.classList.contains('ed'))return;
  if(t.tagName==='TEXTAREA')tinggi(t);
  const p=data[t.dataset.s].pasien[t.dataset.p];
  if(t.dataset.t==='gabung')GABUNG[t.dataset.f].baca(p,t.value);
  else p[t.dataset.f]=t.dataset.t==='list'?t.value.split('\n').map(x=>x.trim()).filter(Boolean):t.value;
  t.classList.add('ubah')});
 el.addEventListener('change',e=>{if(e.target.classList.contains('hapus'))e.target.closest('tr').classList.toggle('dihapus',e.target.checked)});
 const bt=el.querySelector('.bsimpan'),out=el.querySelector('.hsimpan');
 bt.onclick=async()=>{
  const hapus=new Set([...el.querySelectorAll('.hapus:checked')].map(c=>c.dataset.s+'/'+c.dataset.p));
  const kirim=data.map((d,si)=>({...d,pasien:d.pasien.filter((_,pi)=>!hapus.has(si+'/'+pi))}));
  if(kirim.some(d=>d.pasien.some(p=>!String(p.n??'').trim()))){
   out.innerHTML='<div class="err">Ada pasien tanpa nama. Isi "Nama:" di kolom Identitas, atau centang Hapus.</div>';return}
  bt.disabled=true;out.innerHTML='<p class="mute">Menyimpan…</p>';
  try{const x=await post('/api/simpan',{tanggal:tgl,data:kirim});
   el.querySelectorAll('.ed.ubah').forEach(t=>t.classList.remove('ubah'));
   out.innerHTML=`<div class="row" style="margin-top:10px"><b>Tersimpan — ${x.total} pasien.</b>
   <a class="btn" href="/unduh/${encodeURIComponent(x.file)}">Unduh ${esc(x.file)}</a>
   <a class="btn sec" href="/riwayat/${esc(tgl)}">Lihat riwayat</a></div>`}
  catch(e){out.innerHTML=`<div class="err">${esc(e.message)}</div>`}
  bt.disabled=false}}
async function muatEditor(el,tgl){el.innerHTML='<p class="mute">Memuat data…</p>';
 const r=await fetch('/api/data/'+tgl),j=await r.json();
 if(!r.ok){el.innerHTML=`<div class="warn">${esc(j.error)}</div>`;return}
 editor(el,tgl,j.data,j.ruangan)}
"""


def halaman(judul, isi, aktif=""):
    nav = "".join(f'<a href="{u}" class="{"on" if aktif == k else ""}">{t}</a>'
                  for k, u, t in [("input", "/", "Input Laporan"), ("riwayat", "/riwayat", "Riwayat Harian"),
                                  ("pasien", "/pasien", "Riwayat Pasien"), ("ruangan", "/ruangan", "Ruangan")])
    return f"""<!doctype html><html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{E(judul)} · List Pasien</title>
<style>{CSS}</style></head><body><header><div class="in"><b>List Pasien Visite</b>{nav}</div></header>
<main>{isi}</main></body></html>"""


def tgl_id(iso):
    try:
        d = dt.date.fromisoformat(iso)
    except (TypeError, ValueError):
        return iso or "-"
    hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"][d.weekday()]
    return f"{hari}, {d.day} {visite.BULAN[d.month - 1]} {d.year}"


def daftar(xs):
    xs = [x for x in (xs or []) if str(x).strip()]
    return "<ul>" + "".join(f"<li>{E(str(x))}</li>" for x in xs) + "</ul>" if xs else '<span class="mute">-</span>'


def hal_input():
    hari_ini = dt.date.today().isoformat()
    isi = f"""
<h1>Input Laporan</h1>
<p class="mute">Tempel teks laporan bangsal. Klik <b>Cek Data</b> dulu untuk melihat data baru, lalu <b>Proses</b>.
Model: {E(ENGINE)}/{E(MODEL)}</p>
<div class="card">
 <div class="row"><label style="margin:0">Tanggal laporan</label><input type="date" id="tgl" value="{hari_ini}"></div>
 <label for="t0">Laporan</label>
 <textarea id="t0" placeholder="Tempel seluruh laporan di sini (bisa berisi banyak ruangan: RB1, RB2, Pav 5, ...)"></textarea>
 <div id="koreksi"></div>
 <div class="row" style="margin-top:10px">
  <button class="sec" type="button" onclick="tambahKoreksi()">+ Tambah koreksi</button>
  <span class="mute small">Koreksi = teks ruangan yang diperbaiki; ruangan yang sama menggantikan versi di laporan.</span>
 </div>
 <div class="row" style="margin-top:16px"><button id="bcek" onclick="cek()">Cek Data</button></div>
</div>
<div id="hasil"></div>
<script>{EDITOR_JS}</script>
<script>
let nK=0, terakhir=null;
function tambahKoreksi(){{nK++;const d=document.createElement('div');
d.innerHTML=`<label for="k${{nK}}">Koreksi ${{nK}}</label><textarea id="k${{nK}}" style="min-height:140px"></textarea>`;
document.getElementById('koreksi').appendChild(d)}}
function kumpul(){{const t=[{{nama:'laporan',teks:document.getElementById('t0').value}}];
for(let i=1;i<=nK;i++){{const v=document.getElementById('k'+i).value;if(v.trim())t.push({{nama:'koreksi '+i,teks:v}})}}
return {{tanggal:document.getElementById('tgl').value,teks:t}}}}
async function cek(){{const b=kumpul();if(!b.teks[0].teks.trim()){{alert('Laporan masih kosong');return}}
const h=document.getElementById('hasil');h.innerHTML='<div class="card mute">Mengecek…</div>';
try{{const a=await post('/api/cek',b);terakhir=b;tampil(a)}}catch(e){{h.innerHTML=`<div class="err">${{esc(e.message)}}</div>`}}}}
function tampil(a){{
 const urut={{baru:0,berubah:1,'?':2,sama:3}};
 const ps=[...a.pasien].sort((x,y)=>urut[x.status]-urut[y.status]);
 const n=s=>a.pasien.filter(p=>p.status===s).length;
 let h=`<h2>Penjelasan data — ${{esc(a.tgl)}}</h2><div class="stats">
 <div class="stat"><b>${{a.total}}</b><span>pasien di laporan</span></div>
 <div class="stat"><b class="red">${{n('baru')}}</b><span>pasien baru</span></div>
 <div class="stat"><b>${{n('berubah')}}</b><span>pasien ada data baru</span></div>
 <div class="stat"><b>${{a.keluar.length}}</b><span>tidak ada lagi vs list sebelumnya</span></div>
 <div class="stat"><b>${{a.ke_ai}}</b><span>dikirim ke AI (sisanya cache)</span></div></div>`;
 h+=`<div class="card small">Ruangan: ${{Object.entries(a.ruang).map(([r,v])=>`<b>${{esc(r)}}</b> ${{v.n}}${{v.baru?' <span class="tag baru">ruangan baru</span>':''}}${{v.sumber!=='laporan'?' <span class="tag berubah">'+esc(v.sumber)+'</span>':''}}`).join(' · ')}}<br>
 Pembanding: ${{a.tgl_lama?'list '+esc(a.tgl_lama):'<i>belum ada riwayat</i>'}}${{a.tgl_teks?' · teks laporan '+esc(a.tgl_teks):' · <i>teks laporan sebelumnya tidak tersimpan, data baru per baris tidak bisa dibandingkan</i>'}}</div>`;
 if(a.ruang['?'])h+=`<div class="warn">${{a.ruang['?'].n}} pasien tanpa judul ruangan yang dikenali. Kalau ini ruangan baru,
 tambahkan namanya di <a href="/ruangan" target="_blank">Ruangan</a> lalu klik Cek Data lagi.</div>`;
 if(a.sudah_ada)h+=`<div class="warn">Riwayat tanggal ${{esc(a.tgl)}} sudah ada dan akan ditimpa kalau diproses.</div>`;
 const pbj=a.pasien.filter(p=>p.pbj_lewat.length);
 if(pbj.length)h+=`<div class="warn"><b>PBJ sudah lewat (tidak akan dimasukkan list):</b> ${{pbj.map(p=>esc(p.n)+' — '+esc(p.pbj_lewat[0])).join('; ')}}</div>`;
 if(a.keluar.length)h+=`<div class="card"><b>Tidak ada lagi di laporan (pulang/pindah?):</b> ${{a.keluar.map(esc).join(', ')}}</div>`;
 h+=`<div class="wrap"><table><tr><th>Status</th><th>Ruang/Bed</th><th>Pasien</th><th>Data baru dibanding laporan sebelumnya</th><th>AI</th></tr>`;
 for(const p of ps){{
  let d='';
  if(p.status==='baru')d='<span class="mute">Pasien baru</span>';
  else if(p.tambah.length)d=`<details${{p.tambah.length<=6?' open':''}}><summary>${{p.tambah.length}} baris baru${{p.hapus?', '+p.hapus+' baris hilang':''}}</summary><div class="add">${{p.tambah.map(esc).join('\\n')}}</div></details>`;
  else if(p.hapus)d=`<span class="mute">${{p.hapus}} baris dihapus, tidak ada baris baru</span>`;
  else if(p.status==='sama')d='<span class="mute">Tidak berubah</span>';
  else d='<span class="mute">Tidak ada pembanding</span>';
  if(p.status==='baru'&&p.tambah.length)d+=`<div class="add">${{p.tambah.map(esc).join('\\n')}}</div>`;
  h+=`<tr><td><span class="tag ${{p.status==='?'?'q':p.status}}">${{p.status==='?'?'—':p.status}}</span></td>
  <td>${{esc(p.ruang)}}<br><span class="mute small">${{esc(p.k)}}</span></td>
  <td><b>${{esc(p.n)}}</b><br><span class="mute small">${{p.rm?'RM '+esc(p.rm):''}}</span></td><td>${{d}}</td>
  <td class="small">${{p.cache?'<span class="mute">cache</span>':'<b>kirim</b>'}}</td></tr>`}}
 h+=`</table></div><div class="card"><div class="row"><button id="bproses" onclick="proses()">Proses &amp; buat docx</button>
 <span class="mute small">${{a.ke_ai}} pasien dikirim ke AI (${{esc(a.model)}}).</span></div><div id="job"></div></div>`;
 document.getElementById('hasil').innerHTML=h}}
async function proses(){{const bt=document.getElementById('bproses');bt.disabled=true;
const j=document.getElementById('job');j.innerHTML='<p class="mute">Memproses…</p><div class="log" id="log"></div>';
let r;try{{r=await post('/api/proses',terakhir)}}catch(e){{j.innerHTML=`<div class="err">${{esc(e.message)}}</div>`;bt.disabled=false;return}}
const t=setInterval(async()=>{{const s=await (await fetch('/api/job/'+r.id)).json();
 const lg=document.getElementById('log');if(lg){{lg.textContent=s.log.join('\\n');lg.scrollTop=lg.scrollHeight}}
 if(s.status==='jalan')return;clearInterval(t);bt.disabled=false;
 if(s.status==='gagal'){{j.innerHTML+=`<div class="err">Gagal: ${{esc(s.error)}}</div>`;return}}
 const x=s.hasil;j.innerHTML=`<div class="card" style="border-color:var(--grn)"><b>Selesai — ${{x.total}} pasien.</b>
 <div class="row" style="margin:10px 0"><a class="btn" href="/unduh/${{encodeURIComponent(x.file)}}">Unduh ${{esc(x.file)}}</a>
 <a class="btn sec" href="/riwayat/${{esc(terakhir.tanggal)}}">Lihat riwayat hari ini</a></div>
 <p class="small">${{esc(x.catatan)}}</p></div><details><summary>Log</summary><div class="log">${{esc(s.log.join('\\n'))}}</div></details>
 <h2>Preview &amp; edit data</h2><div id="edit"></div>`;
 muatEditor(document.getElementById('edit'),terakhir.tanggal)}},1500)}}
</script>"""
    return halaman("Input", isi, "input")


def hal_riwayat():
    rows = []
    for tgl in tgl_hari():
        m, snap = meta(tgl), snapshot(tgl)
        f = m.get("file")
        unduh = (f'<a href="/unduh/{quote(f)}">docx</a>' if f and (visite.BASE / f).exists() else "")
        rows.append(f"""<tr><td><a href="/riwayat/{tgl}"><b>{E(tgl_id(tgl))}</b></a></td><td>{len(snap)}</td>
<td class="small">{E(m.get("rekap", ""))}</td><td class="small">{"ya" if teks_hari(tgl) else "-"}</td><td>{unduh}</td></tr>""")
    isi = ("<h1>Riwayat Harian</h1><p class='mute'>Satu baris = satu list pasien yang pernah dibuat.</p>"
           + ('<div class="wrap"><table><tr><th>Tanggal</th><th>Pasien</th><th>Per ruangan</th><th>Teks laporan</th><th>File</th></tr>'
              + "".join(rows) + "</table></div>" if rows else "<div class='card mute'>Belum ada riwayat.</div>"))
    return halaman("Riwayat", isi, "riwayat")


def urut_pasien(snap):
    return sorted(snap.items(), key=lambda kv: (visite.urutan_ruang(kv[1].get("r", "")), kv[1].get("k", "")))


def hal_hari(tgl):
    snap = snapshot(tgl)
    if not snap:
        return None
    m = meta(tgl)
    _, peta = registri()
    hari = sorted(tgl_hari())
    i = hari.index(tgl)
    nav = " · ".join(x for x in [
        f'<a href="/riwayat/{hari[i - 1]}">← {E(hari[i - 1])}</a>' if i else "",
        f'<a href="/riwayat/{hari[i + 1]}">{E(hari[i + 1])} →</a>' if i + 1 < len(hari) else ""] if x)
    f = m.get("file")
    isi = f"<h1>{E(tgl_id(tgl))}</h1><p class='mute'>{len(snap)} pasien · {E(m.get('rekap', ''))} {('· ' + nav) if nav else ''}</p>"
    tombol = []
    if f and (visite.BASE / f).exists():
        tombol.append(f'<a class="btn" href="/unduh/{quote(f)}">Unduh {E(f)}</a>')
    if data_hari(tgl) is not None:
        tombol.append(f'<a class="btn sec" href="/edit/{tgl}">Edit data</a>')
    if tombol:
        isi += f'<p class="row">{"".join(tombol)}</p>'
    if m.get("catatan"):
        isi += f'<div class="card"><b>Catatan</b><p class="small">{E(m["catatan"].removeprefix("Catatan: "))}</p>'
        for k, t in (("keluar", "Keluar dari list"), ("pindah", "Pindah ruangan"),
                     ("sudah_pulang", "Tidak dimasukkan (PBJ lewat)")):
            if m.get(k):
                isi += f'<p class="small"><b>{t}:</b> {E(", ".join(m[k]))}</p>'
        isi += "</div>"
    rows = []
    for pid, v in urut_pasien(snap):
        kunci = peta.get((tgl, pid), pid)
        st = v.get("st", "")
        rows.append(f"""<tr><td>{E(v.get("r", ""))}<br><span class="mute small">{E(v.get("k", ""))}</span></td>
<td><a href="/pasien/{quote(kunci, safe='')}"><b>{E(v.get("n", ""))}</b></a><br><span class="mute small">{E(" / ".join(x for x in [v.get("jk", ""), (v.get("u") or "") and f"{v['u']} th", v.get("rm", "") and "RM " + v["rm"]] if x))}</span>
{f'<br><span class="tag baru">{E(st)}</span>' if st else ""}</td>
<td class="small">{E(v.get("dpjp", ""))}</td><td class="small">{E(v.get("dx", "")) or '<span class="mute">-</span>'}</td>
<td class="small">{daftar(v.get("m"))}</td></tr>""")
    isi += ('<h2>Pasien</h2><div class="wrap"><table><tr><th>Ruang</th><th>Pasien</th><th>DPJP</th><th>Diagnosis</th><th>Masalah</th></tr>'
            + "".join(rows) + "</table></div>")
    teks = teks_hari(tgl)
    if teks:
        isi += "<h2>Teks laporan asli</h2>" + "".join(
            f'<details class="card"><summary>{E(x["nama"])} ({len(x["teks"].splitlines())} baris)</summary>'
            f'<div class="log" style="max-height:600px">{E(x["teks"])}</div></details>' for x in teks)
    return halaman(tgl, isi, "riwayat")


def hal_edit(tgl):
    if not snapshot(tgl) and data_hari(tgl) is None:
        return None
    isi = f"""<h1>Edit data — {E(tgl_id(tgl))}</h1>
<p class="mute">Hasil AI untuk tanggal ini. Simpan = docx &amp; riwayat tanggal ini dibuat ulang dari data di bawah.
<a href="/riwayat/{tgl}">← kembali ke riwayat</a></p><div id="ed"></div>
<script>{EDITOR_JS}</script><script>muatEditor(document.getElementById('ed'),{json.dumps(tgl)})</script>"""
    return halaman(f"Edit {tgl}", isi, "riwayat")


def hal_ruangan():
    xs, st = visite.daftar_ruangan(), statistik_ruang()
    kenal = {visite.kunci_ruang(x) for x in xs}

    def info(r):
        s = st.get(r)
        if not s:
            return "<td class='mute small' colspan='3'>belum pernah ada di list</td>"
        return f"<td class='small'>{E(tgl_id(s['akhir']))}</td><td>{s['n']}</td><td>{s['hari']}</td>"

    def tombol(aksi, r, teks, kelas="sec"):
        return f'<button class="mini {kelas}" data-n="{E(r)}" onclick="aksi(\'{aksi}\',this.dataset.n)">{teks}</button>'

    rows = "".join(f"""<tr><td>{i}</td><td><b>{E(r)}</b></td>
<td class="small">{"otomatis" if visite.HEADER_RUANG.match(r) else "baris judul persis <code>" + E(r) + "</code>"}</td>
{info(r)}<td>{tombol("hapus", r, "Hapus")}</td></tr>"""
                   for i, r in enumerate(xs, 1))
    lain = [r for r in st if r != "?" and visite.kunci_ruang(r) not in kenal]
    rows_lain = "".join(f"<tr><td><b>{E(r)}</b></td>{info(r)}<td>{tombol('tambah', r, '+ Tambah', '')}</td></tr>"
                        for r in sorted(lain))
    isi = f"""<h1>Ruangan</h1>
<p class="mute">Ruangan yang dikenali sistem saat memecah laporan. Judul <b>RB n</b>, <b>RA n</b>, <b>Pav n</b>, <b>Ruang n</b>
dikenali otomatis. Ruangan dengan nama lain (mis. ICU, HCU, Melati 2) harus ditambahkan dulu: di laporan, tulis namanya
sebagai satu baris judul sendiri (boleh diakhiri ":") sebelum pasien-pasien ruangan itu.</p>
<div class="card"><form class="row" onsubmit="event.preventDefault();aksi('tambah',this.nama.value)">
<input type="text" name="nama" placeholder="Nama ruangan, mis. ICU" maxlength="30" required>
<button>Tambah ruangan</button></form></div>
<h2>Daftar ruangan <span class="mute small">(urut otomatis: RA, RB, Pav, Ruang, lalu nama lain; per kelompok urut angka)</span></h2>
<div class="wrap"><table><tr><th>#</th><th>Ruangan</th><th>Dikenali dari</th><th>Terakhir di list</th><th>Pasien saat itu</th><th>Hari</th><th></th></tr>
{rows or "<tr><td colspan='7' class='mute'>Daftar kosong.</td></tr>"}</table></div>
{f'''<h2>Terdata di riwayat, belum di daftar</h2><div class="wrap"><table><tr><th>Ruangan</th><th>Terakhir di list</th>
<th>Pasien saat itu</th><th>Hari</th><th></th></tr>{rows_lain}</table></div>''' if lain else ""}
<script>{EDITOR_JS}
async function aksi(a,n){{if(a==='hapus'&&!confirm('Hapus '+n+' dari daftar?'))return;
try{{await post('/api/ruangan',{{aksi:a,nama:n}});location.reload()}}catch(e){{alert(e.message)}}}}</script>"""
    return halaman("Ruangan", isi, "ruangan")


def hal_pasien(q=""):
    orang, _ = registri()
    rows = []
    for kunci, hist in sorted(orang.items(), key=lambda kv: (kv[1][-1][0], kv[1][-1][1].get("n", "")), reverse=True):
        t_akhir, v = hist[-1]
        nama = " ".join({h[1].get("n", ""): 0 for h in hist})
        if q and q.lower() not in (nama + " " + v.get("rm", "")).lower():
            continue
        aktif = t_akhir == (tgl_hari() or [""])[0]
        rows.append(f"""<tr><td><a href="/pasien/{quote(kunci, safe='')}"><b>{E(v.get("n", ""))}</b></a>
{'' if aktif else ' <span class="tag q">tidak di list terakhir</span>'}<br><span class="mute small">{E("RM " + v["rm"] if v.get("rm") else "")}</span></td>
<td>{E(v.get("r", ""))} <span class="mute small">{E(v.get("k", ""))}</span></td><td class="small">{E(v.get("dpjp", ""))}</td>
<td>{len(hist)}</td><td class="small">{E(hist[0][0])} – {E(t_akhir)}</td></tr>""")
    isi = f"""<h1>Riwayat Pasien</h1>
<form class="row" style="margin:10px 0"><input type="search" name="q" value="{E(q)}" placeholder="Cari nama / RM"><button>Cari</button></form>
<div class="wrap"><table><tr><th>Pasien</th><th>Ruang terakhir</th><th>DPJP</th><th>Hari</th><th>Periode</th></tr>{"".join(rows)}</table></div>"""
    return halaman("Pasien", isi, "pasien")


def beda(a, b):
    """Item di b yang tidak ada di a (untuk menandai yang baru)."""
    sa = {norm_baris(str(x)) for x in a or []}
    return {norm_baris(str(x)) for x in b or []} - sa


def hal_satu_pasien(kunci):
    orang, _ = registri()
    hist = orang.get(kunci)
    if not hist:
        return None
    v = hist[-1][1]
    isi = f"""<h1>{E(v.get("n", ""))}</h1><p class="mute">{E(" / ".join(x for x in [v.get("jk", ""), (v.get("u") or "") and f"{v['u']} th", v.get("rm", "") and "RM " + v["rm"]] if x))}
· {len(hist)} hari di list ({E(hist[0][0])} – {E(hist[-1][0])})</p><div class="tl">"""
    prev, kartu = {}, []
    for tgl, x in hist:
        def kol(f, judul, ket=""):
            baru = beda(prev.get(f), x.get(f)) if prev else set()
            items = x.get(f) or []
            if not items:
                return f"<div><b class='small'>{judul}</b><div class='mute small'>-</div></div>"
            li = "".join(f"<li class='{'red' if norm_baris(str(i)) in baru else ''}'>{E(str(i))}</li>" for i in items)
            return f"<div><b class='small'>{judul}</b>{ket}<ul class='small'>{li}</ul></div>"
        ubah = []
        if prev:
            for f, t in (("r", "ruang"), ("k", "bed"), ("dpjp", "DPJP"), ("dx", "diagnosis")):
                if prev.get(f) and x.get(f) and prev.get(f) != x.get(f):
                    ubah.append(t)
        tx_lanjut = f" <span class='mute small'>(sesuai list {E(x['tx_tgl'])})</span>" if x.get("tx_tgl") else ""
        kartu.append(f"""<div class="card"><div class="row"><b><a href="/riwayat/{tgl}">{E(tgl_id(tgl))}</a></b>
<span class="mute">{E(x.get("r", ""))} {E(x.get("k", ""))} · {E(x.get("dpjp", ""))}</span>
{f'<span class="tag baru">{E(x["st"])}</span>' if x.get("st") else ""}
{f'<span class="tag berubah">berubah: {E(", ".join(ubah))}</span>' if ubah else ""}</div>
<p class="small"><b>Dx:</b> {E(x.get("dx", "")) or "-"}{f" · <b>Op:</b> {E(x['to'])}" if x.get("to") else ""}</p>
<div class="grid2">{kol("lab", "Lab " + E(x.get("tl", "")))}{kol("m", "Masalah")}{kol("d", "Diet")}
{kol("tx", "Tatalaksana", tx_lanjut)}{kol("p", "Planning")}{kol("pj", "Penunjang")}</div></div>""")
        prev = x
    isi += "".join(reversed(kartu)) + "</div><p class='mute small'>Merah = item baru dibanding hari sebelumnya.</p>"
    return halaman(v.get("n", "Pasien"), isi, "pasien")


# ---------------------------------------------------------------- server
class Handler(BaseHTTPRequestHandler):
    def kirim(self, kode, body, tipe="text/html; charset=utf-8", header=None):
        b = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(kode)
        self.send_header("Content-Type", tipe)
        self.send_header("Content-Length", str(len(b)))
        for k, v in (header or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(b)

    def json(self, obj, kode=200):
        self.kirim(kode, json.dumps(obj, ensure_ascii=False), "application/json; charset=utf-8")

    def log_message(self, *a):
        pass

    def do_GET(self):
        u = urlparse(self.path)
        p = unquote(u.path)
        try:
            if p == "/":
                return self.kirim(200, hal_input())
            if p == "/riwayat":
                return self.kirim(200, hal_riwayat())
            m = re.fullmatch(r"/riwayat/(\d{4}-\d{2}-\d{2})", p)
            if m:
                h = hal_hari(m.group(1))
                return self.kirim(200, h) if h else self.kirim(404, halaman("404", "<h1>Tidak ada</h1>"))
            m = re.fullmatch(r"/edit/(\d{4}-\d{2}-\d{2})", p)
            if m:
                h = hal_edit(m.group(1))
                return self.kirim(200, h) if h else self.kirim(404, halaman("404", "<h1>Tidak ada</h1>"))
            m = re.fullmatch(r"/api/data/(\d{4}-\d{2}-\d{2})", p)
            if m:
                data = data_hari(m.group(1))
                if data is None:
                    return self.json({"error": "Data hasil AI tanggal ini tidak tersimpan (list dibuat sebelum fitur edit). "
                                               "Proses ulang laporannya di Input Laporan; pasien yang teksnya sama "
                                               "diambil dari cache, tidak dikirim ulang ke AI."}, 404)
                return self.json({"data": data, "ruangan": visite.daftar_ruangan()})
            if p == "/ruangan":
                return self.kirim(200, hal_ruangan())
            if p == "/pasien":
                return self.kirim(200, hal_pasien(parse_qs(u.query).get("q", [""])[0]))
            if p.startswith("/pasien/"):
                h = hal_satu_pasien(p[len("/pasien/"):])
                return self.kirim(200, h) if h else self.kirim(404, halaman("404", "<h1>Pasien tidak ada</h1>"))
            if p.startswith("/api/job/"):
                job = JOBS.get(p.rsplit("/", 1)[1])
                return self.json(job) if job else self.json({"error": "job tidak ada"}, 404)
            if p.startswith("/unduh/"):
                f = visite.BASE / Path(p[len("/unduh/"):]).name  # nama file saja, tidak bisa keluar folder
                if f.suffix != ".docx" or not f.exists():
                    return self.kirim(404, "tidak ada", "text/plain")
                return self.kirim(200, f.read_bytes(),
                                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                  {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(f.name)}"})
            self.kirim(404, halaman("404", "<h1>Halaman tidak ada</h1>"))
        except Exception as e:
            traceback.print_exc()
            self.kirim(500, halaman("Error", f"<div class='err'>{E(repr(e))}</div>"))

    def do_POST(self):
        try:
            n = int(self.headers.get("Content-Length") or 0)
            body = json.loads(self.rfile.read(n) or b"{}")
            if self.path == "/api/ruangan":
                return self.json({"ruangan": ubah_ruangan(body.get("aksi"), body.get("nama"))})
            tgl = dt.date.fromisoformat(body.get("tanggal") or dt.date.today().isoformat())
            if self.path == "/api/simpan":
                data = bersih_data(body.get("data"))
                if not KUNCI_JOB.acquire(blocking=False):
                    return self.json({"error": "Masih ada proses lain berjalan, tunggu selesai."}, 409)
                try:
                    visite.perbarui_cache(data_hari(tgl.isoformat()) or [], data)
                    visite.simpan_data(tgl, data)
                    return self.json(buat_docx(data, tgl))
                finally:
                    KUNCI_JOB.release()
            teks =[(x.get("nama") or f"teks {i}", x.get("teks", "")) for i, x in enumerate(body.get("teks", []))
                    if str(x.get("teks", "")).strip()]
            if not teks:
                return self.json({"error": "Teks laporan kosong"}, 400)
            if self.path == "/api/cek":
                return self.json(analisa(teks, tgl))
            if self.path == "/api/proses":
                if KUNCI_JOB.locked():
                    return self.json({"error": "Masih ada proses lain berjalan, tunggu selesai."}, 409)
                jid = uuid.uuid4().hex[:8]
                JOBS[jid] = {"status": "jalan", "log": [], "hasil": None, "error": ""}
                threading.Thread(target=jalankan, args=(JOBS[jid], teks, tgl), daemon=True).start()
                return self.json({"id": jid})
            self.json({"error": "tidak ada"}, 404)
        except ValueError as e:  # input tidak valid (nama ruangan, tanggal, data edit)
            self.json({"error": str(e)}, 400)
        except Exception as e:
            traceback.print_exc()
            self.json({"error": str(e)}, 500)


def main():
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    print(f"Dashboard: {url}  (Ctrl+C untuk berhenti)")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
