# ===== PARSER REAL FIX - SESUAI JUDUL KOLOM PDF =====
def parse_pdf_real_all(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc: full_text += page.get_text("text") + "\n"
    m = re.search(r"(\d{2}-\d{2}-\d{4})\s*s/d\s*(\d{2}-\d{2}-\d{4})", full_text)
    if m:
        tgl_akhir = m.group(2)
        bln = int(tgl_akhir.split("-")[1]); thn = tgl_akhir.split("-")[2]
        bulan_nama = ["","JANUARI","FEBRUARI","MARET","APRIL","MEI","JUNI","JULI","AGUSTUS","SEPTEMBER","OKTOBER","NOVEMBER","DESEMBER"][bln]
        periode = f"{bulan_nama}-{thn}"
    else:
        periode = normalize_periode(os.path.basename(pdf_path))

    data = []
    for page in doc:
        text = page.get_text("text")
        lines = text.split("\n")
        pic = "UNKNOWN"
        for l in lines:
            if l.strip().isupper() and 3 < len(l.strip()) < 30 and "PENGAMBILAN" not in l and "TOTAL" not in l and "Tanggal" not in l:
                if len(l.strip().split())<=4:
                    pic = l.strip()
        for i, line in enumerate(lines):
            if "274" not in line: continue
            kode_match = re.search(r"(274\d{7})", line)
            if not kode_match: continue
            kode = kode_match.group(1)
            combined = line + " " + (lines[i+1] if i+1 < len(lines) else "")
            # Ambil semua angka: 2.756.000 atau 9 atau 2
            all_nums = re.findall(r"\d{1,3}(?:\.\d{3})+|\b\d+\b", combined)
            clean=[]
            for n in all_nums:
                if n==kode: continue
                if "." in n: clean.append(int(n.replace(".","")))
                else:
                    try: clean.append(int(n))
                    except: pass
            # Minimal 13 angka: Bln, CD_Total, CD_Reward, CD_Round, CD_Psn, SA_Total, SA_Reward, SA_Round, SA_Psn, TOTAL_Total, TOTAL_Reward, TOTAL_Pasien, TOTAL_Round
            if len(clean) < 13: continue
            try:
                Bln = clean[0] # <-- INI BULAN, JANGAN DIPAKAI JADI PASIEN!
                CD_Total = clean[1]
                CD_Psn_cabang = clean[4] # <-- Psn per cabang CIK DI TIRO - INI YANG BENAR!
                SA_Total = clean[5]
                SA_Psn_cabang = clean[8] # <-- Psn per cabang SULTAN AGUNG - INI YANG BENAR!
                TOTAL_Total = clean[9]
                TOTAL_Pasien_total = clean[11] # <-- pasien di TOTAL = total gabungan - INI YANG BENAR!

                # Nama dokter
                nama = line.split(kode)[-1]
                nama = re.sub(r"\s+\d+\s+[\d.]+\s*$", "", nama).strip()[:80]
                if len(nama)<3: nama = f"dr. {kode}"

                data.append({
                    "Kode_Dokter": kode,
                    "Nama_Dokter": nama,
                    "PIC": pic,
                    "Periode": periode,
                    "Bln": Bln,
                    "CD_Omzet": CD_Total, # Omzet CD
                    "CD_Pasien": CD_Psn_cabang, # FIX: Psn cabang, bukan Bln!
                    "SA_Omzet": SA_Total,
                    "SA_Pasien": SA_Psn_cabang, # FIX: Psn cabang
                    "Total_Omzet": TOTAL_Total,
                    "Total_Pasien": TOTAL_Pasien_total # FIX: pasien total
                })
            except: continue
    doc.close()
    df = pd.DataFrame(data)
    if not df.empty:
        df = df.groupby(["Kode_Dokter","Periode"], as_index=False).agg({
            "Nama_Dokter":"first","PIC":"first",
            "CD_Omzet":"max","CD_Pasien":"max",
            "SA_Omzet":"max","SA_Pasien":"max",
            "Total_Omzet":"max","Total_Pasien":"max"
        })
    return df, periode
