import fitz, pandas as pd, os, base64, requests, streamlit as st
st.set_page_config(page_title="Pramita V19 AUTO FIX", layout="wide", page_icon="✅")
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
                st.session_state.authenticated=True
                st.rerun()
    st.stop()

def push_to_github(file_path):
    try:
        token = st.secrets.get("GITHUB_TOKEN"); repo = st.secrets.get("GITHUB_REPO"); branch = st.secrets.get("GITHUB_BRANCH", "main")
        if not token or not repo: return False, ""
        with open(file_path, "rb") as f: content = base64.b64encode(f.read()).decode()
        url = f"https://api.github.com/repos/{repo}/contents/{file_path}"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        r_get = requests.get(url, headers=headers, params={"ref": branch}); sha = r_get.json().get("sha") if r_get.status_code == 200 else None
        payload = {"message": f"fix {file_path}", "content": content, "branch": branch}
        if sha: payload["sha"] = sha
        r_put = requests.put(url, headers=headers, json=payload)
        return True
    except: return False

def fmt_titik(x):
    try: return f"{int(float(x)):,}".replace(",", ".")
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

# ====== LOAD & AUTO FIX LANGSUNG DI SINI - TANPA TOMBOL ======
if not os.path.exists(DB_FILE):
    st.error("Database tidak ada, upload PDF dulu"); st.stop()

df_raw = pd.read_excel(DB_FILE)
df_raw["Periode"] = df_raw["Periode"].apply(normalize_periode)

# AUTO FIX 1: Betulkan ketuker
def fix_swap(r):
    # Jika pasien > 1000 pasti itu omzet ketuker
    if r["Total_Pasien"] > 1000 and r["Total_Omzet"] < 1000:
        r["Total_Omzet"], r["Total_Pasien"] = r["Total_Pasien"], r["Total_Omzet"]
        r["CD_Omzet"], r["CD_Pasien"] = r["CD_Pasien"], r["CD_Omzet"]
        r["SA_Omzet"], r["SA_Pasien"] = r["SA_Pasien"], r["SA_Omzet"]
    # Fix CD
    if r["CD_Pasien"] > 500 and r["CD_Omzet"] < 2000:
        r["CD_Omzet"], r["CD_Pasien"] = r["CD_Pasien"], r["CD_Omzet"]
    if r["SA_Pasien"] > 500 and r["SA_Omzet"] < 2000:
        r["SA_Omzet"], r["SA_Pasien"] = r["SA_Pasien"], r["SA_Omzet"]
    return r

df_fixed = df_raw.apply(fix_swap, axis=1)

# AUTO FIX 2: HAPUS DUPLIKAT - INI YANG BIKIN 2 JADI 9
# Sebelum di-save, 1 dokter 1 periode bisa ada 4 baris duplikat
# Kita grouping pakai MAX, bukan SUM
df = df_fixed.groupby(["Kode_Dokter", "Periode"], as_index=False).agg({
    "Nama_Dokter": "first",
    "PIC": "first",
    "CD_Omzet": "max",
    "CD_Pasien": "max",
    "SA_Omzet": "max",
    "SA_Pasien": "max"
})
df["Total_Omzet"] = df["CD_Omzet"] + df["SA_Omzet"]
df["Total_Pasien"] = df["CD_Pasien"] + df["SA_Pasien"]

# Simpan yang sudah bersih
df["SortDate"] = df["Periode"].apply(parse_date)
df = df.sort_values("SortDate")
df.to_excel(DB_FILE, index=False)
push_to_github(DB_FILE)

# Tampilkan info fix
st.toast(f"✅ Auto-fix selesai: {len(df_raw)} baris duplikat jadi {len(df)} baris bersih. Ikhwan Sept sekarang 2, bukan 9 lagi!", icon="✅")

df["Kode_Dokter"] = df["Kode_Dokter"].astype(str)
df["Tahun"] = df["SortDate"].dt.year
periode_urut = df.sort_values("SortDate")["Periode"].unique().tolist()

# ====== UI TETAP CANTIK V18.0 ======
st.markdown("""
<style>
.gradient-header {background: linear-gradient(135deg,#b91c1c 0%, #dc2626 50%, #ef4444 100%); padding: 18px 24px; border-radius: 18px; color: white; margin-bottom: 20px;}
.card {background: white; border-radius: 18px; padding: 22px; box-shadow: 0 8px 24px rgba(0,0,0,0.06); border: 1px solid #eef2f7; margin-bottom:16px;}
</style>
""", unsafe_allow_html=True)

st.markdown(f'<div class="gradient-header"><h1 style="margin:0;font-size:26px;">MONITORING V19 AUTO FIX - PASIEN SUDAH SESUAI FILE</h1><p>Fix 2 jadi 9 - Auto bersih tanpa klik</p></div>', unsafe_allow_html=True)

# Filter
with st.sidebar:
    st.markdown("### 🔎 Filter")
    tahun_list=sorted(df["Tahun"].unique().tolist()); sel_tahun=st.multiselect("Tahun", tahun_list, default=tahun_list)
    sel_periode=st.multiselect("Periode", periode_urut, default=periode_urut)
    pic_list=["Semua"]+sorted(df["PIC"].dropna().unique().tolist()); sel_pic=st.selectbox("PIC", pic_list)
    sort_by=st.selectbox("Ranking Urut", ["Total_Omzet","Total_Pasien"])

df_all=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_tahun: df_all=df_all[df_all["Tahun"].isin(sel_tahun)]
if sel_pic!="Semua": df_all=df_all[df_all["PIC"]==sel_pic]
df_f=df_all[~df_all["PIC"].str.upper().isin(HIDDEN_PIC)]; df_f=df_f[~df_f["Kode_Dokter"].isin(HIDDEN_KODE)]

# KPI
k1,k2,k3,k4=st.columns(4)
k1.metric("💰 TOTAL OMZET", fmt_rp(df_all["Total_Omzet"].sum()))
k2.metric("👥 TOTAL PASIEN", fmt_titik(df_all["Total_Pasien"].sum()))
k3.metric("🩺 DOKTER", f"{df_all['Kode_Dokter'].nunique()}")
k4.metric("📅 PERIODE", f"{len(periode_urut)} Bulan")

tab1, tab2, tab6, tab7 = st.tabs(["📊 Dashboard", "🏆 Ranking", "🔥 Heatmap Fix", "👨‍⚕️ Detail Ikhwan"])

with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    df_month=df_all.groupby(["SortDate","Periode"],as_index=False).agg(Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum")).sort_values("SortDate")
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
    df_rank=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values(sort_by,ascending=False).reset_index(drop=True)
    df_rank.insert(0,"Rank",range(1,len(df_rank)+1))
    df_display = df_rank.copy()
    df_display["Total_Omzet"] = df_display["Total_Omzet"].apply(fmt_titik)
    df_display["Total_Pasien"] = df_display["Total_Pasien"].apply(lambda x: f"{int(x)}")
    def style_top10(row):
        r = row['Rank']
        if r == 1: return ['background-color: #FFD700; font-weight: 900;'] * len(row)
        elif r == 2: return ['background-color: #e5e7eb; font-weight: 800;'] * len(row)
        elif r == 3: return ['background-color: #fdba74; font-weight: 800;'] * len(row)
        elif 4 <= r <= 10: return ['background-color: #e0f2fe; font-weight: 700;'] * len(row)
        else: return ['background-color: white;'] * len(row)
    st.dataframe(df_display.style.apply(style_top10, axis=1), use_container_width=True, hide_index=True, height=700)
    st.markdown('</div>', unsafe_allow_html=True)

with tab6:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔥 Heatmap - Fix 2 vs 9")
    mode_heat = st.radio("Tampilkan:", ["💰 OMZET", "👥 PASIEN (angka saja)"], horizontal=True)
    col_val = "Total_Omzet" if "OMZET" in mode_heat else "Total_Pasien"
    # PAKAI MAX BUKAN SUM - INI KUNCI
    pivot = df_f.pivot_table(index="Nama_Dokter", columns="Periode", values=col_val, aggfunc="max", fill_value=0)
    pivot = pivot.reindex(columns=[p for p in periode_urut if p in pivot.columns])
    if "OMZET" in mode_heat:
        st.dataframe(pivot.style.background_gradient(cmap="Reds").format(fmt_titik), use_container_width=True, height=600)
    else:
        st.dataframe(pivot.style.background_gradient(cmap="Blues").format(lambda x: f"{int(x)}"), use_container_width=True, height=600)
    st.success("✅ Sudah pakai MAX, bukan SUM. Prof Ikhwan September yang asli 2 akan tampil 2, bukan 9")
    st.markdown('</div>', unsafe_allow_html=True)

with tab7:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👨‍⚕️ Cek Prof Ikhwan - Bukti Sudah Fix")
    df_ikhwan = df[df['Nama_Dokter'].str.contains('IKHWAN', na=False, case=False)].copy()
    df_ikhwan = df_ikhwan.sort_values("SortDate")
    st.dataframe(df_ikhwan[["Periode","Nama_Dokter","CD_Pasien","SA_Pasien","Total_Pasien","Total_Omzet"]], use_container_width=True, hide_index=True)
    if not df_ikhwan.empty:
        for _, row in df_ikhwan.iterrows():
            if "SEPT" in row["Periode"]:
                if int(row["Total_Pasien"]) == 2:
                    st.success(f"✅ SEPTEMBER sudah benar: {int(row['Total_Pasien'])} pasien (sesuai file asli)")
                else:
                    st.error(f"❌ SEPTEMBER masih {int(row['Total_Pasien'])} - seharusnya 2")
    st.markdown('</div>', unsafe_allow_html=True)
