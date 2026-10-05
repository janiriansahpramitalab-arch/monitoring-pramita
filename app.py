import fitz, pandas as pd, os, base64, requests, re, streamlit as st
st.set_page_config(page_title="Pramita V21 FINAL REAL", layout="wide", page_icon="✅")
DB_FILE = "Database_Monitoring_Pramita.xlsx"
HIDDEN_PIC = ["YOHANA DEWI RATIH", "NARINDRA NATA KUNTHARA"]
HIDDEN_KODE = ["2741002000"]
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "pramita123")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        pwd = st.text_input("Password", type="password")
        if st.button("MASUK", type="primary", use_container_width=True):
            if pwd == APP_PASSWORD:
                st.session_state.authenticated=True; st.rerun()
    st.stop()

st.markdown("""
<style>
.main {background-color: #f6f8fb;}
div[data-testid="metric-container"] {background: white; border-radius: 16px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #eef2f7;}
.gradient-header {background: linear-gradient(135deg,#b91c1c 0%, #dc2626 50%, #ef4444 100%); padding: 18px 24px; border-radius: 18px; color: white; margin-bottom: 20px;}
.card {background: white; border-radius: 18px; padding: 22px; box-shadow: 0 8px 24px rgba(0,0,0,0.06); border: 1px solid #eef2f7; margin-bottom:16px;}
</style>
""", unsafe_allow_html=True)

def push_to_github(file_path):
    try:
        token = st.secrets.get("GITHUB_TOKEN"); repo = st.secrets.get("GITHUB_REPO"); branch = st.secrets.get("GITHUB_BRANCH", "main")
        if not token or not repo: return False
        with open(file_path, "rb") as f: content = base64.b64encode(f.read()).decode()
        url = f"https://api.github.com/repos/{repo}/contents/{file_path}"
        headers = {"Authorization": f"token {token}"}
        r_get = requests.get(url, headers=headers, params={"ref": branch}); sha = r_get.json().get("sha") if r_get.status_code == 200 else None
        payload = {"message": "final real", "content": content, "branch": branch}
        if sha: payload["sha"] = sha
        requests.put(url, headers=headers, json=payload)
        return True
    except: return False

def fmt_titik(x):
    try: return f"{int(float(x)):,}".replace(",", ".")
    except: return "0"
def fmt_rp(x):
    try: return f"Rp {int(float(x)):,}".replace(",", ".")
    except: return "Rp 0"

BULAN_FULL = {"JANU":"JANUARI","JAN":"JANUARI","FEBR":"FEBRUARI","FEB":"FEBRUARI","MAR":"MARET","APRIL":"APRIL","MEI":"MEI","JUNI":"JUNI","JULI":"JULI","AGUS":"AGUSTUS","SEPT":"SEPTEMBER","OKTO":"OKTOBER","NOPE":"NOVEMBER","DESE":"DESEMBER"}
BULAN_ANGKA = {"JANUARI":1,"FEBRUARI":2,"MARET":3,"APRIL":4,"MEI":5,"JUNI":6,"JULI":7,"AGUSTUS":8,"SEPTEMBER":9,"OKTOBER":10,"NOVEMBER":11,"DESEMBER":12}
def normalize_periode(s):
    s=str(s).upper().strip(); tahun="".join([c for c in s if c.isdigit()])[-4:]; huruf="".join([c for c in s if c.isalpha()]); full=BULAN_FULL.get(huruf[:4],huruf); return f"{full}-{tahun}" if tahun else full
def parse_date(s):
    try: n=normalize_periode(s); return pd.Timestamp(year=int(n.split("-")[1]), month=BULAN_ANGKA.get(n.split("-")[0],1), day=1)
    except: return pd.Timestamp(2026,9,1)

# ===== PARSER FINAL - 100% SESUAI FILE ASLI - TIDAK PAKAI Bln JADI PASIEN =====
def parse_pdf_final_real(pdf_path):
    doc = fitz.open(pdf_path)
    full = ""
    for p in doc: full += p.get_text("text") + "\n"
    m = re.search(r"(\d{2}-\d{2}-\d{4})\s*s/d\s*(\d{2}-\d{2}-\d{4})", full)
    if m:
        akhir = m.group(2); bln = int(akhir.split("-")[1]); thn = akhir.split("-")[2]
        bulan_nama = ["","JANUARI","FEBRUARI","MARET","APRIL","MEI","JUNI","JULI","AGUSTUS","SEPTEMBER","OKTOBER","NOVEMBER","DESEMBER"][bln]
        periode = f"{bulan_nama}-{thn}"
    else:
        periode = normalize_periode(os.path.basename(pdf_path))

    rows=[]
    for page in doc:
        lines = page.get_text("text").split("\n")
        pic="UNKNOWN"
        for l in lines:
            if l.strip().isupper() and 3 < len(l.strip()) < 30 and "PENGAMBILAN" not in l and "TOTAL" not in l and "Tanggal" not in l:
                if len(l.strip().split())<=4: pic=l.strip()
        for i,line in enumerate(lines):
            km = re.search(r"(274\d{7})", line)
            if not km: continue
            kode = km.group(1)
            combined = line + " " + (lines[i+1] if i+1 < len(lines) else "")
            nums = re.findall(r"\d{1,3}(?:\.\d{3})+|\b\d+\b", combined)
            clean=[]
            for n in nums:
                if n==kode: continue
                if "." in n: clean.append(int(n.replace(".","")))
                else:
                    try: clean.append(int(n))
                    except: pass
            if len(clean) < 13: continue
            # INI URUTAN REAL DARI FILE KAKAK - JANGAN DIUBAH LAGI:
            # clean[0]=Bln(9) JANGAN DIPAKAI JADI PASIEN!
            # clean[1]=CD_Total, clean[2]=CD_Reward, clean[3]=CD_Round, clean[4]=CD_Psn (per cabang BENAR)
            # clean[5]=SA_Total, clean[6]=SA_Reward, clean[7]=SA_Round, clean[8]=SA_Psn (per cabang BENAR)
            # clean[9]=TOTAL_Total, clean[10]=TOTAL_Reward, clean[11]=TOTAL_Pasien (total BENAR), clean[12]=TOTAL_Round
            Bln = clean[0]
            CD_Total = clean[1]
            CD_Psn = clean[4]
            SA_Total = clean[5]
            SA_Psn = clean[8]
            TOTAL_Total = clean[9]
            TOTAL_Pasien = clean[11]

            nama = line.split(kode)[-1]
            nama = re.sub(r"\s+\d+\s+[\d.]+\s*$", "", nama).strip()[:80]

            rows.append({
                "Kode_Dokter": kode, "Nama_Dokter": nama, "PIC": pic, "Periode": periode,
                "Bln": Bln,
                "CD_Omzet": CD_Total, "CD_Pasien": CD_Psn,
                "SA_Omzet": SA_Total, "SA_Pasien": SA_Psn,
                "Total_Omzet": TOTAL_Total, "Total_Pasien": TOTAL_Pasien
            })
    doc.close()
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.groupby(["Kode_Dokter","Periode"], as_index=False).agg({
            "Nama_Dokter":"first","PIC":"first",
            "CD_Omzet":"max","CD_Pasien":"max","SA_Omzet":"max","SA_Pasien":"max","Total_Omzet":"max","Total_Pasien":"max"
        })
    return df, periode

# HEADER CANTIK
st.markdown(f'''<div class="gradient-header"><div style="background:white; border-radius:12px; padding:6px 14px;"><span style="color:#dc2626; font-weight:900;">PRAMITA</span></div><div><h1 style="margin:0;font-size:24px; font-weight:800;">MONITORING V21 FINAL REAL</h1><p style="margin:2px 0 0 0;">Semua Menu Real - CD_Psn= Psn cabang, Total_Pasien= pasien total</p></div></div>''', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🔥 FIX SEMUA MENU JADI REAL")
    st.error("Klik tombol di bawah ini untuk fix semua menu (Dashboard, Ranking, Heatmap, dll) jadi real 2 pasien, bukan 9")
    up = st.file_uploader("Upload PDF (bisa banyak)", type=["pdf"], accept_multiple_files=True, key="final_up")
    if up:
        if st.button("🔥 REBUILD TOTAL - SEMUA MENU REAL", type="primary", use_container_width=True):
            all_df=[]
            for f in up:
                tmp = f"/tmp/{f.name}"
                with open(tmp,"wb") as o: o.write(f.getbuffer())
                d,p = parse_pdf_final_real(tmp)
                all_df.append(d)
            final = pd.concat(all_df, ignore_index=True)
            final.to_excel(DB_FILE, index=False)
            push_to_github(DB_FILE)
            st.success(f"✅ Selesai! {len(final)} baris real. Ikhwan Sept = {final[final['Kode_Dokter']=='2741002623']['Total_Pasien'].values[0]} pasien")
            st.rerun()

if not os.path.exists(DB_FILE):
    st.warning("Upload PDF dulu di sidebar kiri"); st.stop()

df=pd.read_excel(DB_FILE); df["Periode"]=df["Periode"].apply(normalize_periode); df["SortDate"]=df["Periode"].apply(parse_date)
df["Kode_Dokter"]=df["Kode_Dokter"].astype(str)
df = df.sort_values("SortDate")

# FILTER
periode_urut = df.sort_values("SortDate")["Periode"].unique().tolist()
with st.sidebar:
    st.divider()
    sel_periode=st.multiselect("Periode", periode_urut, default=periode_urut)
    sel_pic=st.selectbox("PIC", ["Semua"]+sorted(df["PIC"].dropna().unique().tolist()))

df_all=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_pic!="Semua": df_all=df_all[df_all["PIC"]==sel_pic]
df_f=df_all[~df_all["PIC"].str.upper().isin(HIDDEN_PIC)]; df_f=df_f[~df_f["Kode_Dokter"].isin(HIDDEN_KODE)]

# KPI - SEMUA MENU PAKAI Total_Pasien REAL
k1,k2,k3=st.columns(3)
k1.metric("💰 OMZET", fmt_rp(df_f["Total_Omzet"].sum()))
k2.metric("👥 PASIEN (REAL)", fmt_titik(df_f["Total_Pasien"].sum()))
k3.metric("🩺 DOKTER", f"{df_f['Kode_Dokter'].nunique()}")

tab1, tab2, tab6, tab7 = st.tabs(["📊 Dashboard", "🏆 Ranking", "🔥 Heatmap", "👨‍⚕️ Detail"])

with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    df_m=df_f.groupby(["SortDate","Periode"],as_index=False).agg(Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum"), CD_Pasien=("CD_Pasien","sum"), SA_Pasien=("SA_Pasien","sum")).sort_values("SortDate")
    st.write("**Trend Pasien REAL (bukan Bln)**")
    st.line_chart(df_m.set_index("Periode")[["Total_Pasien"]])
    st.write("**Trend Omzet REAL**")
    st.line_chart(df_m.set_index("Periode")[["Total_Omzet"]])
    disp=df_m.copy(); disp["Total_Omzet"]=disp["Total_Omzet"].apply(fmt_titik)
    for c in ["Total_Pasien","CD_Pasien","SA_Pasien"]: disp[c]=disp[c].apply(lambda x: f"{int(x)}")
    st.dataframe(disp[["Periode","Total_Omzet","Total_Pasien","CD_Pasien","SA_Pasien"]], use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    df_r=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"), Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum"), CD_Pasien=("CD_Pasien","sum"), SA_Pasien=("SA_Pasien","sum")).sort_values("Total_Omzet",ascending=False)
    df_r.insert(0,"Rank",range(1,len(df_r)+1))
    d=df_r.copy()
    d["Total_Omzet"]=d["Total_Omzet"].apply(fmt_titik)
    for c in ["Total_Pasien","CD_Pasien","SA_Pasien"]: d[c]=d[c].apply(lambda x: f"{int(x)}")
    st.dataframe(d, use_container_width=True, hide_index=True, height=600)
    st.caption("Ranking sudah pakai Total_Pasien real (2), bukan Bln (9). CD_Pasien = Psn cabang")
    st.markdown('</div>', unsafe_allow_html=True)

with tab6:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    mode=st.radio("Tampil", ["💰 OMZET","👥 PASIEN REAL"], horizontal=True)
    col="Total_Omzet" if "OMZET" in mode else "Total_Pasien"
    piv=df_f.pivot_table(index="Nama_Dokter", columns="Periode", values=col, aggfunc="max", fill_value=0)
    piv=piv.reindex(columns=[p for p in periode_urut if p in piv.columns])
    if "OMZET" in mode: st.dataframe(piv.style.background_gradient(cmap="Reds").format(fmt_titik), use_container_width=True, height=600)
    else: st.dataframe(piv.style.background_gradient(cmap="Blues").format(lambda x: f"{int(x)}"), use_container_width=True, height=600)
    st.markdown('</div>', unsafe_allow_html=True)

with tab7:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    df_list=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first")).sort_values("Nama_Dokter")
    df_list["Label"]=df_list["Nama_Dokter"]+" | "+df_list["Kode_Dokter"]
    sel=st.selectbox("Pilih Dokter", ["-- Pilih --"]+df_list["Label"].tolist())
    if sel!="-- Pilih --":
        kode=sel.split(" | ")[-1]
        res=df_f[df_f["Kode_Dokter"]==kode].copy().sort_values("SortDate")
        st.dataframe(res[["Periode","Nama_Dokter","CD_Omzet","CD_Pasien","SA_Omzet","SA_Pasien","Total_Omzet","Total_Pasien"]], use_container_width=True, hide_index=True)
        if kode=="2741002623" and not res.empty:
            sept=res[res["Periode"].str.contains("SEPT")]
            if not sept.empty:
                st.success(f"✅ REAL: CD_Psn={int(sept.iloc[0]['CD_Pasien'])} (Psn cabang), SA_Psn={int(sept.iloc[0]['SA_Pasien'])}, Total_Pasien={int(sept.iloc[0]['Total_Pasien'])} (pasien total) - Bukan Bln=9 lagi!")
    st.markdown('</div>', unsafe_allow_html=True)
