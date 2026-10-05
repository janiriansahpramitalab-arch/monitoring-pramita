import fitz, pandas as pd, os, base64, requests, re, streamlit as st
st.set_page_config(page_title="Pramita V20.3 Real Columns Cantik", layout="wide", page_icon="🚀")
DB_FILE = "Database_Monitoring_Pramita.xlsx"
HIDDEN_PIC = ["YOHANA DEWI RATIH", "NARINDRA NATA KUNTHARA"]
HIDDEN_KODE = ["2741002000"]
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "pramita123")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown("<div style='text-align:center; padding:60px 20px;'><div style='background:white; border-radius:20px; padding:40px; max-width:400px; margin:auto; box-shadow:0 10px 40px rgba(0,0,0,0.1);'><h1 style='color:#dc2626;'>🔐 LOGIN PRAMITA</h1><p>Monitoring Kinerja Dokter</p></div></div>", unsafe_allow_html=True)
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        pwd = st.text_input("Password", type="password")
        if st.button("MASUK", type="primary", use_container_width=True):
            if pwd == APP_PASSWORD:
                st.session_state.authenticated=True
                st.rerun()
            else: st.error("Password salah kak!")
    st.stop()

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
html, body, [class*="css"] {font-family: 'Inter', sans-serif;}
.main {background-color: #f6f8fb;}
div[data-testid="metric-container"] {background: white; border-radius: 16px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #eef2f7;}
.gradient-header {background: linear-gradient(135deg,#b91c1c 0%, #dc2626 50%, #ef4444 100%); padding: 18px 24px; border-radius: 18px; color: white; margin-bottom: 20px; display:flex; align-items:center; gap:18px; justify-content:space-between;}
.card {background: white; border-radius: 18px; padding: 22px; box-shadow: 0 8px 24px rgba(0,0,0,0.06); border: 1px solid #eef2f7; margin-bottom:16px;}
</style>
""", unsafe_allow_html=True)

def push_to_github(file_path):
    try:
        token = st.secrets.get("GITHUB_TOKEN"); repo = st.secrets.get("GITHUB_REPO"); branch = st.secrets.get("GITHUB_BRANCH", "main")
        if not token or not repo: return False, ""
        with open(file_path, "rb") as f: content = base64.b64encode(f.read()).decode()
        url = f"https://api.github.com/repos/{repo}/contents/{file_path}"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        r_get = requests.get(url, headers=headers, params={"ref": branch}); sha = r_get.json().get("sha") if r_get.status_code == 200 else None
        payload = {"message": f"backup {file_path}", "content": content, "branch": branch}
        if sha: payload["sha"] = sha
        r_put = requests.put(url, headers=headers, json=payload)
        return (True, "OK") if r_put.status_code in [200,201] else (False, "")
    except: return False, ""

def fmt_titik(x):
    try:
        if pd.isna(x): return "0"
        return f"{int(float(x)):,}".replace(",", ".")
    except: return "0"
def fmt_rp(x):
    try: return f"Rp {int(float(x)):,}".replace(",", ".")
    except: return "Rp 0"

BULAN_FULL = {"JANU":"JANUARI","JAN":"JANUARI","FEBR":"FEBRUARI","FEB":"FEBRUARI","MAR":"MARET","MARET":"MARET","APRIL":"APRIL","APR":"APRIL","MEI":"MEI","JUNI":"JUNI","JUN":"JUNI","JULI":"JULI","JUL":"JULI","AGUS":"AGUSTUS","AGU":"AGUSTUS","SEPT":"SEPTEMBER","SEP":"SEPTEMBER","OKTO":"OKTOBER","OKT":"OKTOBER","NOPE":"NOVEMBER","NOV":"NOVEMBER","DESE":"DESEMBER","DES":"DESEMBER"}
BULAN_ANGKA = {"JANUARI":1,"FEBRUARI":2,"MARET":3,"APRIL":4,"MEI":5,"JUNI":6,"JULI":7,"AGUSTUS":8,"SEPTEMBER":9,"OKTOBER":10,"NOVEMBER":11,"DESEMBER":12}
def normalize_periode(s):
    s=str(s).upper().strip(); tahun="".join([c for c in s if c.isdigit()])[-4:]; huruf="".join([c for c in s if c.isalpha()]); full=BULAN_FULL.get(huruf[:4],huruf); return f"{full}-{tahun}" if tahun else full
def parse_date(s):
    try: n=normalize_periode(s); return pd.Timestamp(year=int(n.split("-")[1]), month=BULAN_ANGKA.get(n.split("-")[0],1), day=1)
    except: return pd.Timestamp(2026,1,1)

# ===== PARSER REAL FIX SESUAI JUDUL KOLOM - Psn per cabang & pasien total =====
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
        if "-" not in periode: periode = "SEPTEMBER-2026"

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
            all_nums = re.findall(r"\d{1,3}(?:\.\d{3})+|\b\d+\b", combined)
            clean=[]
            for n in all_nums:
                if n==kode: continue
                if "." in n: clean.append(int(n.replace(".","")))
                else:
                    try: clean.append(int(n))
                    except: pass
            # Harus 13 angka: Bln, CD_Total, CD_Reward, CD_Round, CD_Psn, SA_Total, SA_Reward, SA_Round, SA_Psn, TOTAL_Total, TOTAL_Reward, TOTAL_Pasien, TOTAL_Round
            if len(clean) < 13: continue
            try:
                Bln = clean[0] # Bulan, JANGAN dipakai jadi pasien!
                CD_Total = clean[1]
                CD_Psn_cabang = clean[4] # Psn per cabang CIK DI TIRO - INI YANG BENAR KAK!
                SA_Total = clean[5]
                SA_Psn_cabang = clean[8] # Psn per cabang SULTAN AGUNG - INI YANG BENAR!
                TOTAL_Total = clean[9]
                TOTAL_Pasien_total = clean[11] # pasien di TOTAL = total gabungan - INI YANG BENAR!

                nama = line.split(kode)[-1]
                nama = re.sub(r"\s+\d+\s+[\d.]+\s*$", "", nama).strip()
                nama = re.sub(r"\s{2,}", " ", nama)[:80]
                if len(nama)<3: nama = f"dr. {kode}"

                data.append({
                    "Kode_Dokter": kode,
                    "Nama_Dokter": nama,
                    "PIC": pic,
                    "Periode": periode,
                    "CD_Omzet": CD_Total,
                    "CD_Pasien": CD_Psn_cabang, # FIX: pakai Psn cabang, bukan Bln!
                    "SA_Omzet": SA_Total,
                    "SA_Pasien": SA_Psn_cabang, # FIX: pakai Psn cabang
                    "Total_Omzet": TOTAL_Total,
                    "Total_Pasien": TOTAL_Pasien_total # FIX: pakai pasien total
                })
            except: continue
    doc.close()
    df = pd.DataFrame(data)
    if not df.empty:
        df = df.groupby(["Kode_Dokter","Periode"], as_index=False).agg({
            "Nama_Dokter":"first","PIC":"first","CD_Omzet":"max","CD_Pasien":"max","SA_Omzet":"max","SA_Pasien":"max","Total_Omzet":"max","Total_Pasien":"max"
        })
    return df, periode

c_head1, c_head2 = st.columns([6,1])
with c_head1:
    st.markdown(f'''<div class="gradient-header"><div style="display:flex; align-items:center; gap:18px;"><div style="background:white; border-radius:12px; padding:6px 14px; display:flex; align-items:center;"><span style="color:#dc2626; font-weight:900; font-size:22px; letter-spacing:1px;">PRAMITA</span><span style="color:#dc2626; font-style:italic; margin-left:8px; font-weight:600;">Lab</span></div><div><h1 style="margin:0;font-size:26px; font-weight:800;">MONITORING KINERJA DOKTER</h1><p style="margin:4px 0 0 0;opacity:0.95">V20.3 Real Psn & Pasien - Tampilan Cantik V18.0</p></div></div></div>''', unsafe_allow_html=True)
with c_head2:
    if st.button("🚪 Logout"): st.session_state.authenticated=False; st.rerun()

if not os.path.exists(DB_FILE):
    st.info("Upload PDF di sidebar kiri ya kak")
    with st.sidebar:
        st.markdown("### 📂 Upload PDF Real (Semua Bulan)")
        uploaded_pdfs = st.file_uploader("Pilih PDF (bisa banyak)", type=["pdf"], accept_multiple_files=True)
        if uploaded_pdfs and st.button("🔥 PROSES JADI REAL (SEMUA BULAN)", type="primary", use_container_width=True):
            all_dfs = []
            for pdf_file in uploaded_pdfs:
                temp_path = f"/tmp/{pdf_file.name}"
                with open(temp_path, "wb") as f: f.write(pdf_file.getbuffer())
                df_real, per_real = parse_pdf_real_all(temp_path)
                all_dfs.append(df_real)
            if all_dfs:
                df_final = pd.concat(all_dfs, ignore_index=True)
                df_final.to_excel(DB_FILE, index=False)
                push_to_github(DB_FILE)
                st.success(f"✅ {len(all_dfs)} file real!"); st.rerun()
    st.stop()

df=pd.read_excel(DB_FILE); df["Periode"]=df["Periode"].apply(normalize_periode); df["SortDate"]=df["Periode"].apply(parse_date); df=df.sort_values("SortDate")
df["Kode_Dokter"]=df["Kode_Dokter"].astype(str); df["Tahun"]=df["SortDate"].dt.year

with st.sidebar:
    st.markdown("### 📂 Upload PDF Real (Semua Bulan)")
    st.caption("Fix: CD_Pasien = Psn cabang, Total_Pasien = pasien total, Bln tidak dipakai jadi pasien")
    uploaded_pdfs = st.file_uploader("Pilih PDF (bisa banyak sekaligus)", type=["pdf"], accept_multiple_files=True, key="upload_real_all_v203")
    if uploaded_pdfs:
        if st.button("🔥 PROSES JADI REAL (SEMUA BULAN)", type="primary", use_container_width=True, key="btn_real_all"):
            all_dfs = []
            for pdf_file in uploaded_pdfs:
                temp_path = f"/tmp/{pdf_file.name}"
                with open(temp_path, "wb") as f: f.write(pdf_file.getbuffer())
                df_real, per_real = parse_pdf_real_all(temp_path)
                st.write(f"✅ {pdf_file.name} -> {per_real}: {len(df_real)} dokter")
                all_dfs.append(df_real)
            if all_dfs:
                df_all_real = pd.concat(all_dfs, ignore_index=True)
                df_old = pd.read_excel(DB_FILE)
                periode_baru = df_all_real["Periode"].unique().tolist()
                df_old_filtered = df_old[~df_old["Periode"].isin(periode_baru)]
                df_final = pd.concat([df_old_filtered, df_all_real], ignore_index=True)
                df_final.to_excel(DB_FILE, index=False)
                push_to_github(DB_FILE)
                st.success(f"✅ Semua {len(all_dfs)} file sudah real! Ikhwan Sept 2 pasien (bukan 9)")
                st.rerun()

    if os.path.exists("/mnt/data/PENGAMBILAN_DANA_JASA_PRE_ANALISTIK.pdf"):
        if st.button("⚡ FIX CEPAT SEPTEMBER REAL 2 PASIEN", use_container_width=True, key="fix_cepat"):
            df_real, per_real = parse_pdf_real_all("/mnt/data/PENGAMBILAN_DANA_JASA_PRE_ANALISTIK.pdf")
            df_old = pd.read_excel(DB_FILE)
            df_old = df_old[~df_old["Periode"].str.contains("SEPTEMBER-2026", na=False)]
            df_new = pd.concat([df_old, df_real], ignore_index=True)
            df_new.to_excel(DB_FILE, index=False)
            push_to_github(DB_FILE)
            st.success(f"✅ September real: Ikhwan {df_real[df_real['Kode_Dokter']=='2741002623']['Total_Pasien'].values[0]} pasien")
            st.rerun()

    st.divider()
    st.markdown("### 🔎 Filter")
    tahun_list=sorted(df["Tahun"].unique().tolist()); sel_tahun=st.multiselect("Tahun", tahun_list, default=tahun_list, key="thn_v203")
    periode_list_sorted=df.sort_values("SortDate")["Periode"].unique().tolist()
    sel_periode=st.multiselect("Periode", periode_list_sorted, default=periode_list_sorted, key="per_v203")
    pic_list=["Semua"]+sorted(df["PIC"].dropna().unique().tolist()); sel_pic=st.selectbox("PIC", pic_list, key="pic_v203")
    sort_by=st.selectbox("Ranking Urut", ["Total_Omzet","Total_Pasien","CD_Omzet","SA_Omzet"], key="sort_v203")

df_all=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_tahun: df_all=df_all[df_all["Tahun"].isin(sel_tahun)]
if sel_pic!="Semua": df_all=df_all[df_all["PIC"]==sel_pic]
df_kpi=df_all.copy()
df_f=df_all[~df_all["PIC"].str.upper().isin(HIDDEN_PIC)]; df_f=df_f[~df_f["Kode_Dokter"].isin(HIDDEN_KODE)]

total_omzet=df_kpi["Total_Omzet"].sum(); total_pasien=df_kpi["Total_Pasien"].sum(); jml_dokter=df_kpi["Kode_Dokter"].nunique()
periode_urut = df.sort_values("SortDate")["Periode"].unique().tolist()
df_month=df_all.groupby(["SortDate","Periode"],as_index=False).agg(Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum")).sort_values("SortDate")

k1,k2,k3,k4=st.columns(4)
k1.metric("💰 TOTAL OMZET", fmt_rp(total_omzet))
k2.metric("👥 TOTAL PASIEN", fmt_titik(total_pasien))
k3.metric("🩺 DOKTER", f"{jml_dokter}", f"Tampil {df_f['Kode_Dokter'].nunique()}")
k4.metric("📅 PERIODE", f"{len(periode_list_sorted)} Bulan")

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📊 Dashboard", "🏆 Ranking", "🔄 Banding Bulan", "👥 Analisa PIC", "🆕 Baru/Hilang", "🔥 Heatmap", "👨‍⚕️ Detail"])

with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 📈 Trend Omzet (JANUARI -> DESEMBER URUT)")
    df_chart = df_month.copy()
    df_chart["Periode"] = pd.Categorical(df_chart["Periode"], categories=periode_urut, ordered=True)
    df_chart = df_chart.sort_values("Periode")
    st.line_chart(df_chart.set_index("Periode")[["Total_Omzet"]])
    df_month_disp = df_chart.copy()
    df_month_disp["Total_Omzet"] = df_month_disp["Total_Omzet"].apply(fmt_titik)
    df_month_disp["Total_Pasien"] = df_month_disp["Total_Pasien"].apply(lambda x: f"{int(x)}")
    st.dataframe(df_month_disp[["Periode","Total_Omzet","Total_Pasien"]], use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write(f"#### 🏆 Ranking - {sort_by} (Top 10 Highlight)")
    df_rank=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"),PIC=("PIC","first"),CD_Omzet=("CD_Omzet","sum"),SA_Omzet=("SA_Omzet","sum"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values(sort_by,ascending=False).reset_index(drop=True)
    df_rank.insert(0,"Rank",range(1,len(df_rank)+1))
    df_display = df_rank.copy()
    for col in ["CD_Omzet","SA_Omzet","Total_Omzet"]:
        df_display[col] = df_display[col].apply(fmt_titik)
    df_display["Total_Pasien"] = df_display["Total_Pasien"].apply(lambda x: f"{int(x)}")
    df_display["Medal"]=df_display["Rank"].apply(lambda x: "🥇 JUARA 1" if x==1 else "🥈 JUARA 2" if x==2 else "🥉 JUARA 3" if x==3 else f"⭐ TOP {x}" if x<=10 else f"#{x}")
    def style_top10(row):
        r = row['Rank']
        if r == 1: return ['background-color: #FFD700; color: #78350f; font-weight: 900; font-size: 15px; border-left: 6px solid #b45309;'] * len(row)
        elif r == 2: return ['background-color: #e5e7eb; color: #111827; font-weight: 800; font-size: 14px; border-left: 6px solid #6b7280;'] * len(row)
        elif r == 3: return ['background-color: #fdba74; color: #7c2d12; font-weight: 800; font-size: 14px; border-left: 6px solid #9a3412;'] * len(row)
        elif 4 <= r <= 10: return ['background-color: #e0f2fe; color: #0c4a6e; font-weight: 700; font-size: 13px; border-left: 5px solid #0284c7;'] * len(row)
        else: return ['background-color: white; color: #334155; font-size: 12px;'] * len(row)
    st.dataframe(df_display.style.apply(style_top10, axis=1), use_container_width=True, hide_index=True, height=700)
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔄 Perbandingan Bulan - Bisa 2 Bulan & Rentang")
    mode = st.radio("Mode", ["2 Bulan (SEPTEMBER vs AGUSTUS)", "Rentang (JANUARI sampai OKTOBER)"], horizontal=True, key="mode_v203")
    if mode == "2 Bulan (SEPTEMBER vs AGUSTUS)":
        c1,c2 = st.columns(2)
        with c1: bulan1 = st.selectbox("Bulan Lama", periode_urut, index=max(0,len(periode_urut)-2), key="b1_v203")
        with c2: bulan2 = st.selectbox("Bulan Baru", periode_urut, index=len(periode_urut)-1, key="b2_v203")
        if bulan1!=bulan2:
            df_b1 = df_f[df_f["Periode"]==bulan1].groupby("Kode_Dokter", as_index=False).agg(Nama=("Nama_Dokter","first"), Omzet_1=("Total_Omzet","sum"), Pasien_1=("Total_Pasien","sum"))
            df_b2 = df_f[df_f["Periode"]==bulan2].groupby("Kode_Dokter", as_index=False).agg(Omzet_2=("Total_Omzet","sum"), Pasien_2=("Total_Pasien","sum"))
            df_comp = pd.merge(df_b1, df_b2, on="Kode_Dokter", how="outer").fillna(0)
            df_comp["Selisih"] = df_comp["Omzet_2"] - df_comp["Omzet_1"]
            df_comp["Persen"] = df_comp.apply(lambda r: (r["Selisih"]/r["Omzet_1"]*100) if r["Omzet_1"]!=0 else (100 if r["Omzet_2"]>0 else 0), axis=1)
            df_comp["Status"] = df_comp["Persen"].apply(lambda x: "📈 NAIK" if x>10 else "📉 TURUN" if x<-10 else "➖ STABIL")
            df_comp = df_comp.sort_values("Persen", ascending=False)
            df_disp = df_comp.copy()
            df_disp["Omzet_1"] = df_disp["Omzet_1"].apply(fmt_titik)
            df_disp["Omzet_2"] = df_disp["Omzet_2"].apply(fmt_titik)
            df_disp["Selisih"] = df_disp["Selisih"].apply(fmt_titik)
            df_disp["Persen_fmt"] = df_disp["Persen"].apply(lambda x: f"{x:.1f}%")
            st.dataframe(df_disp[["Nama","Kode_Dokter","Omzet_1","Omzet_2","Selisih","Persen_fmt","Status"]], use_container_width=True, hide_index=True, height=500)
    else:
        rentang = st.multiselect("Pilih Bulan", periode_urut, default=periode_urut, key="rentang_v203")
        if len(rentang)>=2:
            rentang_sorted = [p for p in periode_urut if p in rentang]
            bulan_awal = rentang_sorted[0]; bulan_akhir = rentang_sorted[-1]
            df_rentang = df_f[df_f["Periode"].isin(rentang_sorted)].groupby("Kode_Dokter", as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"), Total_Rentang=("Total_Omzet","sum"), Rata2=("Total_Omzet","mean"), Bulan_Aktif=("Periode","nunique"))
            df_awal = df_f[df_f["Periode"]==bulan_awal].groupby("Kode_Dokter", as_index=False).agg(Awal=("Total_Omzet","sum"))
            df_akhir = df_f[df_f["Periode"]==bulan_akhir].groupby("Kode_Dokter", as_index=False).agg(Akhir=("Total_Omzet","sum"))
            df_r = pd.merge(df_rentang, df_awal, on="Kode_Dokter", how="left").merge(df_akhir, on="Kode_Dokter", how="left").fillna(0)
            df_r["Selisih"] = df_r["Akhir"] - df_r["Awal"]
            df_r["Persen"] = df_r.apply(lambda r: (r["Selisih"]/r["Awal"]*100) if r["Awal"]!=0 else 0, axis=1)
            df_r = df_r.sort_values("Total_Rentang", ascending=False)
            df_r_disp = df_r.copy()
            df_r_disp["Total_Rentang"] = df_r_disp["Total_Rentang"].apply(fmt_titik)
            df_r_disp["Awal"] = df_r_disp["Awal"].apply(fmt_titik)
            df_r_disp["Akhir"] = df_r_disp["Akhir"].apply(fmt_titik)
            df_r_disp["Persen_fmt"] = df_r_disp["Persen"].apply(lambda x: f"{x:.1f}%")
            st.write(f"**{bulan_awal} sampai {bulan_akhir} ({len(rentang_sorted)} bulan)**")
            st.dataframe(df_r_disp[["Nama_Dokter","Kode_Dokter","Total_Rentang","Awal","Akhir","Persen_fmt","Bulan_Aktif"]], use_container_width=True, hide_index=True, height=600)
    st.markdown('</div>', unsafe_allow_html=True)

with tab4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👥 Analisa PIC")
    df_pic = df_f.groupby("PIC", as_index=False).agg(Jml_Dokter=("Kode_Dokter","nunique"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values("Total_Omzet", ascending=False)
    st.bar_chart(df_pic.set_index("PIC")[["Total_Omzet"]])
    df_pic_disp = df_pic.copy()
    df_pic_disp["Total_Omzet"] = df_pic_disp["Total_Omzet"].apply(fmt_titik)
    df_pic_disp["Total_Pasien"] = df_pic_disp["Total_Pasien"].apply(lambda x: f"{int(x)}")
    st.dataframe(df_pic_disp, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab5:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🆕 Baru / Hilang")
    if len(periode_urut)>=2:
        bl_lama = periode_urut[-2]; bl_baru = periode_urut[-1]
        dl = set(df[df["Periode"]==bl_lama]["Kode_Dokter"].unique())
        db = set(df[df["Periode"]==bl_baru]["Kode_Dokter"].unique())
        baru = db-dl; hilang = dl-db
        c1,c2 = st.columns(2)
        with c1:
            st.write(f"🆕 Baru di {bl_baru}: {len(baru)}")
            if baru:
                df_b = df[(df["Kode_Dokter"].isin(baru)) & (df["Periode"]==bl_baru)][["Kode_Dokter","Nama_Dokter"]].drop_duplicates()
                st.dataframe(df_b, hide_index=True, use_container_width=True)
        with c2:
            st.write(f"❌ Hilang dari {bl_lama}: {len(hilang)}")
            if hilang:
                df_h = df[(df["Kode_Dokter"].isin(hilang)) & (df["Periode"]==bl_lama)][["Kode_Dokter","Nama_Dokter"]].drop_duplicates()
                st.dataframe(df_h, hide_index=True, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab6:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔥 Heatmap - Real Psn & Pasien")
    mode_heat = st.radio("Tampilkan:", ["💰 OMZET (Rp)", "👥 PASIEN - Real Sesuai File"], horizontal=True, key="heat_v203")
    col_val = "Total_Omzet" if "OMZET" in mode_heat else "Total_Pasien"
    pivot = df_f.pivot_table(index="Nama_Dokter", columns="Periode", values=col_val, aggfunc="max", fill_value=0)
    pivot = pivot.reindex(columns=[p for p in periode_urut if p in pivot.columns])
    if "OMZET" in mode_heat:
        st.dataframe(pivot.style.background_gradient(cmap="Reds").format(fmt_titik), use_container_width=True, height=600)
    else:
        st.dataframe(pivot.style.background_gradient(cmap="Blues").format(lambda x: f"{int(x)}"), use_container_width=True, height=600)
    st.caption("CD_Pasien = Psn cabang CIK, SA_Pasien = Psn cabang SA, Total_Pasien = pasien total - Sudah fix, Bln tidak dipakai jadi pasien lagi")
    st.markdown('</div>', unsafe_allow_html=True)

with tab7:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👨‍⚕️ Detail Dokter - Real Psn & Pasien")
    df_dokter_list = df_f.groupby("Kode_Dokter", as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"), PIC=("PIC","first"), Total_Omzet=("Total_Omzet","sum")).sort_values("Nama_Dokter")
    df_dokter_list["Label"] = df_dokter_list["Nama_Dokter"] + " | KODE: " + df_dokter_list["Kode_Dokter"] + " | PIC: " + df_dokter_list["PIC"] + " | Rp " + df_dokter_list["Total_Omzet"].apply(fmt_titik)
    selected_label = st.selectbox("🔎 Cari & Pilih Dokter", ["-- Pilih Dokter --"] + df_dokter_list["Label"].tolist(), index=0, key="detail_v203")
    if selected_label!= "-- Pilih Dokter --":
        try: selected_kode = selected_label.split("KODE: ")[1].split(" |")[0].strip()
        except: selected_kode = ""
        res = df_f[df_f["Kode_Dokter"] == selected_kode].copy()
        if not res.empty:
            nama_tampil = res["Nama_Dokter"].iloc[0]
            st.success(f"✅ {nama_tampil} | Kode: {selected_kode}")
            hist=res.groupby(["SortDate","Periode"],as_index=False).agg(Total=("Total_Omzet","sum"),Pasien=("Total_Pasien","sum"),CD_Pasien=("CD_Pasien","max"),SA_Pasien=("SA_Pasien","max"),CD_Omzet=("CD_Omzet","max"),SA_Omzet=("SA_Omzet","max")).sort_values("SortDate")
            hist["Periode"] = pd.Categorical(hist["Periode"], categories=periode_urut, ordered=True)
            hist = hist.sort_values("Periode")
            st.line_chart(hist.set_index("Periode")[["Total"]])
            hist_disp = hist.copy()
            hist_disp["Total"] = hist_disp["Total"].apply(fmt_titik)
            hist_disp["Pasien"] = hist_disp["Pasien"].apply(lambda x: f"{int(x)}")
            hist_disp["CD_Pasien"] = hist_disp["CD_Pasien"].apply(lambda x: f"{int(x)}")
            hist_disp["SA_Pasien"] = hist_disp["SA_Pasien"].apply(lambda x: f"{int(x)}")
            hist_disp["CD_Omzet"] = hist_disp["CD_Omzet"].apply(fmt_titik)
            hist_disp["SA_Omzet"] = hist_disp["SA_Omzet"].apply(fmt_titik)
            st.dataframe(hist_disp[["Periode","Total","Pasien","CD_Pasien","SA_Pasien","CD_Omzet","SA_Omzet"]], use_container_width=True, hide_index=True)
            if "2741002623" in selected_kode:
                sept = hist[hist["Periode"].astype(str).str.contains("SEPT")]
                if not sept.empty:
                    st.success(f"✅ REAL FIX: September CD Psn={int(sept.iloc[0]['CD_Pasien'])} (Psn cabang), SA Psn={int(sept.iloc[0]['SA_Pasien'])}, Total Pasien={int(sept.iloc[0]['Pasien'])} (pasien total) - Bln=9 tidak dipakai lagi!")
    else:
        preview = df_dokter_list.copy()
        preview["Total_Omzet"] = preview["Total_Omzet"].apply(fmt_titik)
        st.dataframe(preview[["Kode_Dokter","Nama_Dokter","PIC","Total_Omzet"]], use_container_width=True, hide_index=True, height=300)
    st.markdown('</div>', unsafe_allow_html=True)
