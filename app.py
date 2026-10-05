import fitz, pandas as pd, os, base64, requests, streamlit as st
st.set_page_config(page_title="Pramita V18.9 FIX FINAL", layout="wide", page_icon="✅")
DB_FILE = "Database_Monitoring_Pramita.xlsx"
HIDDEN_PIC = ["YOHANA DEWI RATIH", "NARINDRA NATA KUNTHARA"]
HIDDEN_KODE = ["2741002000"]
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "pramita123")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown("<div style='text-align:center; padding:60px 20px;'><div style='background:white; border-radius:20px; padding:40px; max-width:400px; margin:auto; box-shadow:0 10px 40px rgba(0,0,0,0.1);'><h1 style='color:#dc2626;'>🔐 LOGIN PRAMITA</h1></div></div>", unsafe_allow_html=True)
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        pwd = st.text_input("Password", type="password")
        if st.button("MASUK", type="primary", use_container_width=True):
            if pwd == APP_PASSWORD:
                st.session_state.authenticated=True
                st.rerun()
            else: st.error("Password salah!")
    st.stop()

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
html, body, [class*="css"] {font-family: 'Inter', sans-serif;}
.main {background-color: #f6f8fb;}
div[data-testid="metric-container"] {background: white; border-radius: 16px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #eef2f7;}
.gradient-header {background: linear-gradient(135deg,#b91c1c 0%, #dc2626 50%, #ef4444 100%); padding: 18px 24px; border-radius: 18px; color: white; margin-bottom: 20px;}
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
        payload = {"message": f"fix {file_path}", "content": content, "branch": branch}
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

# === FIX UTAMA: HAPUS DUPLIKAT & BETULKAN PASIEN ===
def repair_final(df_input):
    # 1. Fix ketuker omzet vs pasien
    def fix_swap(r):
        if r["Total_Pasien"] > 2000 and r["Total_Omzet"] < 2000:
            r["Total_Omzet"], r["Total_Pasien"] = r["Total_Pasien"], r["Total_Omzet"]
        if r["CD_Pasien"] > 2000 and r["CD_Omzet"] < 2000:
            r["CD_Omzet"], r["CD_Pasien"] = r["CD_Pasien"], r["CD_Omzet"]
        if r["SA_Pasien"] > 2000 and r["SA_Omzet"] < 2000:
            r["SA_Omzet"], r["SA_Pasien"] = r["SA_Pasien"], r["SA_Omzet"]
        # Jika pasien masih > 500 (tidak wajar 1 dokter 1 bulan 500 pasien), kemungkinan masih ketuker
        if r["CD_Pasien"] > 500 and r["CD_Omzet"] < 5000:
            r["CD_Omzet"], r["CD_Pasien"] = r["CD_Pasien"], r["CD_Omzet"]
        if r["SA_Pasien"] > 500 and r["SA_Omzet"] < 5000:
            r["SA_Omzet"], r["SA_Pasien"] = r["SA_Pasien"], r["SA_Omzet"]
        r["Total_Omzet"] = r["CD_Omzet"] + r["SA_Omzet"]
        r["Total_Pasien"] = r["CD_Pasien"] + r["SA_Pasien"]
        return r
    df = df_input.apply(fix_swap, axis=1)

    # 2. HAPUS DUPLIKAT - INI KUNCI KENAPA JADI 9 PADAHAL ASLI 2
    # Sebelum: 1 dokter 1 periode ada 5 baris (upload berkali-kali) -> sum jadi 9
    # Sesudah: groupby max jadi 1 baris saja -> tetap 2
    df = df.groupby(["Kode_Dokter", "Periode"], as_index=False).agg({
        "Nama_Dokter": "first",
        "PIC": "first",
        "CD_Omzet": "max",
        "CD_Pasien": "max",
        "SA_Omzet": "max",
        "SA_Pasien": "max"
    })
    df["Total_Omzet"] = df["CD_Omzet"] + df["SA_Omzet"]
    df["Total_Pasien"] = df["CD_Pasien"] + df["SA_Pasien"]
    return df

c_head1, c_head2 = st.columns([6,1])
with c_head1:
    st.markdown(f'''<div class="gradient-header"><div><h1 style="margin:0;font-size:26px; font-weight:800;">MONITORING V18.9 FIX FINAL</h1><p style="margin:4px 0 0 0;opacity:0.95">Fix Pasien 2 vs 9 - Tanpa Upload Ulang</p></div></div>''', unsafe_allow_html=True)
with c_head2:
    if st.button("🚪 Logout"): st.session_state.authenticated=False; st.rerun()

with st.sidebar:
    st.markdown("### 🛠️ Perbaikan 1 Klik")
    st.info("Jika pasien tidak sesuai file (contoh Ikhwan 2 jadi 9), klik tombol ini. Tidak perlu upload ulang!")
    if os.path.exists(DB_FILE):
        if st.button("✅ FIX SEKARANG - PASIEN JADI SESUAI FILE", type="primary", use_container_width=True):
            df_old = pd.read_excel(DB_FILE)
            before = len(df_old)
            df_fixed = repair_final(df_old)
            after = len(df_fixed)
            df_fixed.to_excel(DB_FILE, index=False)
            push_to_github(DB_FILE)
            # Cek Ikhwan
            ikhwan = df_fixed[df_fixed['Nama_Dokter'].str.contains('IKHWAN', na=False, case=False)]
            st.success(f"Berhasil! {before} -> {after} baris. Duplikat hilang.")
            if not ikhwan.empty:
                st.write(ikhwan[["Periode","Nama_Dokter","Total_Pasien","Total_Omzet"]])
            st.rerun()

        if st.button("🔍 CEK DATA IKHWAN SEPTEMBER", use_container_width=True):
            df_cek = pd.read_excel(DB_FILE)
            df_ikhwan = df_cek[df_cek['Nama_Dokter'].str.contains('IKHWAN', na=False, case=False)]
            st.dataframe(df_ikhwan[["Periode","Nama_Dokter","CD_Pasien","SA_Pasien","Total_Pasien","Total_Omzet"]], use_container_width=True)
            st.write(f"Total baris Ikhwan: {len(df_ikhwan)} - Seharusnya 10 baris untuk 10 bulan")

    st.divider()
    st.markdown("### 📂 Upload (jika perlu)")

if not os.path.exists(DB_FILE): st.info("Upload PDF dulu"); st.stop()
df=pd.read_excel(DB_FILE); df["Periode"]=df["Periode"].apply(normalize_periode); df["SortDate"]=df["Periode"].apply(parse_date); df=df.sort_values("SortDate")
df["Kode_Dokter"]=df["Kode_Dokter"].astype(str); df["Tahun"]=df["SortDate"].dt.year

with st.sidebar:
    st.markdown("### 🔎 Filter")
    tahun_list=sorted(df["Tahun"].unique().tolist()); sel_tahun=st.multiselect("Tahun", tahun_list, default=tahun_list)
    periode_list_sorted=df.sort_values("SortDate")["Periode"].unique().tolist()
    sel_periode=st.multiselect("Periode", periode_list_sorted, default=periode_list_sorted)
    pic_list=["Semua"]+sorted(df["PIC"].dropna().unique().tolist()); sel_pic=st.selectbox("PIC", pic_list)
    sort_by=st.selectbox("Ranking Urut", ["Total_Omzet","Total_Pasien"])

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
k3.metric("🩺 DOKTER", f"{jml_dokter}")
k4.metric("📅 PERIODE", f"{len(periode_list_sorted)} Bulan")

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📊 Dashboard", "🏆 Ranking", "🔄 Banding Bulan", "👥 PIC", "🆕 Baru/Hilang", "🔥 Heatmap", "👨‍⚕️ Detail"])

with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 📈 Trend")
    df_chart = df_month.copy()
    df_chart["Periode"] = pd.Categorical(df_chart["Periode"], categories=periode_urut, ordered=True)
    df_chart = df_chart.sort_values("Periode")
    st.line_chart(df_chart.set_index("Periode")[["Total_Omzet"]])
    df_month_disp = df_chart.copy()
    df_month_disp["Total_Omzet"] = df_month_disp["Total_Omzet"].apply(fmt_titik)
    df_month_disp["Total_Pasien"] = df_month_disp["Total_Pasien"].apply(fmt_titik)
    st.dataframe(df_month_disp[["Periode","Total_Omzet","Total_Pasien"]], use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write(f"#### 🏆 Ranking - {sort_by} - Top 10 Emas")
    df_rank=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values(sort_by,ascending=False).reset_index(drop=True)
    df_rank.insert(0,"Rank",range(1,len(df_rank)+1))
    df_display = df_rank.copy()
    df_display["Total_Omzet"] = df_display["Total_Omzet"].apply(fmt_titik)
    df_display["Total_Pasien"] = df_display["Total_Pasien"].apply(lambda x: f"{int(x)}")
    df_display["Medal"]=df_display["Rank"].apply(lambda x: "🥇 JUARA 1" if x==1 else "🥈 JUARA 2" if x==2 else "🥉 JUARA 3" if x==3 else f"⭐ TOP {x}" if x<=10 else f"#{x}")
    def style_top10(row):
        r = row['Rank']
        if r == 1: return ['background-color: #FFD700; color: #78350f; font-weight: 900; font-size: 15px; border-left: 6px solid #b45309;'] * len(row)
        elif r == 2: return ['background-color: #e5e7eb; font-weight: 800;'] * len(row)
        elif r == 3: return ['background-color: #fdba74; font-weight: 800;'] * len(row)
        elif 4 <= r <= 10: return ['background-color: #e0f2fe; font-weight: 700;'] * len(row)
        else: return ['background-color: white;'] * len(row)
    st.dataframe(df_display.style.apply(style_top10, axis=1), use_container_width=True, hide_index=True, height=700)
    st.markdown('</div>', unsafe_allow_html=True)

with tab6:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔥 Heatmap - Fix Pasien")
    mode_heat = st.radio("Tampilkan:", ["💰 OMZET (Rp)", "👥 PASIEN (angka saja)"], horizontal=True, key="heat_v189")
    col_val = "Total_Omzet" if "OMZET" in mode_heat else "Total_Pasien"
    pivot = df_f.pivot_table(index="Nama_Dokter", columns="Periode", values=col_val, aggfunc="max", fill_value=0)
    pivot = pivot.reindex(columns=[p for p in periode_urut if p in pivot.columns])
    if "OMZET" in mode_heat:
        st.dataframe(pivot.style.background_gradient(cmap="Reds").format(fmt_titik), use_container_width=True, height=600)
    else:
        st.dataframe(pivot.style.background_gradient(cmap="Blues").format(lambda x: f"{int(x)}"), use_container_width=True, height=600)
    st.caption("Sudah pakai MAX bukan SUM, jadi Prof Ikhwan September yang asli 2 akan tampil 2, bukan 9")
    st.markdown('</div>', unsafe_allow_html=True)

with tab7:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👨‍⚕️ Detail Dokter - Cek Ikhwan")
    df_dokter_list = df_f.groupby("Kode_Dokter", as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"), Total_Omzet=("Total_Omzet","sum")).sort_values("Nama_Dokter")
    df_dokter_list["Label"] = df_dokter_list["Nama_Dokter"] + " | KODE: " + df_dokter_list["Kode_Dokter"]
    selected_label = st.selectbox("🔎 Cari Dokter (contoh IKHWAN)", ["-- Pilih --"] + df_dokter_list["Label"].tolist(), index=0, key="detail_v189")
    if selected_label!= "-- Pilih --":
        try: selected_kode = selected_label.split("KODE: ")[1].strip()
        except: selected_kode = ""
        res = df_f[df_f["Kode_Dokter"] == selected_kode].copy()
        if not res.empty:
            st.success(f"✅ {res['Nama_Dokter'].iloc[0]}")
            hist=res.groupby(["SortDate","Periode"],as_index=False).agg(Total=("Total_Omzet","sum"),Pasien=("Total_Pasien","sum")).sort_values("SortDate")
            hist["Periode"] = pd.Categorical(hist["Periode"], categories=periode_urut, ordered=True)
            hist = hist.sort_values("Periode")
            st.dataframe(hist[["Periode","Pasien","Total"]], use_container_width=True, hide_index=True)
            st.write("Pasien per bulan (sudah fix):")
            for _, row in hist.iterrows():
                st.write(f"- {row['Periode']}: {int(row['Pasien'])} pasien")
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔄 Banding Bulan")
    mode = st.radio("Mode", ["2 Bulan", "Rentang"], horizontal=True, key="mode_v189")
    if mode == "2 Bulan":
        c1,c2 = st.columns(2)
        with c1: bulan1 = st.selectbox("Bulan Lama", periode_urut, index=max(0,len(periode_urut)-2))
        with c2: bulan2 = st.selectbox("Bulan Baru", periode_urut, index=len(periode_urut)-1)
        if bulan1!=bulan2:
            df_b1 = df_f[df_f["Periode"]==bulan1].groupby("Kode_Dokter", as_index=False).agg(Nama=("Nama_Dokter","first"), Omzet_1=("Total_Omzet","sum"))
            df_b2 = df_f[df_f["Periode"]==bulan2].groupby("Kode_Dokter", as_index=False).agg(Omzet_2=("Total_Omzet","sum"))
            df_comp = pd.merge(df_b1, df_b2, on="Kode_Dokter", how="outer").fillna(0)
            df_comp["Selisih"] = df_comp["Omzet_2"] - df_comp["Omzet_1"]
            st.dataframe(df_comp, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)
with tab4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👥 PIC")
    df_pic = df_f.groupby("PIC", as_index=False).agg(Jml_Dokter=("Kode_Dokter","nunique"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum"))
    st.dataframe(df_pic, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)
with tab5:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🆕 Baru/Hilang")
    st.dataframe(df_f.head(), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
