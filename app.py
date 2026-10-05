import fitz, pandas as pd, os, base64, requests, streamlit as st
st.set_page_config(page_title="Pramita V18.7 Repair Tanpa Upload", layout="wide", page_icon="🚀")
DB_FILE = "Database_Monitoring_Pramita.xlsx"
HIDDEN_PIC = ["YOHANA DEWI RATIH", "NARINDRA NATA KUNTHARA"]
HIDDEN_KODE = ["2741002000"]
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "pramita123")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown("<div style='text-align:center; padding:60px 20px;'><div style='background:white; border-radius:20px; padding:40px; max-width:400px; margin:auto;'><h1 style='color:#dc2626;'>🔐 LOGIN PRAMITA</h1></div></div>", unsafe_allow_html=True)
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
        payload = {"message": f"repair {file_path}", "content": content, "branch": branch}
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

# === FUNGSI REPAIR TANPA UPLOAD ULANG ===
def repair_existing_data_smart(df_input):
    """Perbaiki data lama yang ketuker tanpa upload ulang"""
    def fix_row(r):
        try:
            # Ambil semua angka
            vals = [r["CD_Omzet"], r["CD_Pasien"], r["SA_Omzet"], r["SA_Pasien"]]
            # Pisahkan besar vs kecil
            besar = [v for v in vals if v > 5000]
            kecil = [v for v in vals if v <= 2000 and v >= 0]

            # Jika ada yang ketuker: CD_Omzet kecil tapi CD_Pasien besar
            if r["CD_Omzet"] < 1000 and r["CD_Pasien"] > 5000:
                r["CD_Omzet"], r["CD_Pasien"] = r["CD_Pasien"], r["CD_Omzet"]
            if r["SA_Omzet"] < 1000 and r["SA_Pasien"] > 5000:
                r["SA_Omzet"], r["SA_Pasien"] = r["SA_Pasien"], r["SA_Omzet"]

            # Jika Total masih ketuker
            if r["Total_Pasien"] > 1000 and r["Total_Omzet"] < 1000:
                r["Total_Omzet"], r["Total_Pasien"] = r["Total_Pasien"], r["Total_Omzet"]

            # Hitung ulang dengan logika smart: 2 besar = omzet, 2 kecil = pasien
            all_vals = [r["CD_Omzet"], r["SA_Omzet"], r["CD_Pasien"], r["SA_Pasien"]]
            besar = sorted([v for v in all_vals if v > 5000], reverse=True)
            kecil = sorted([v for v in all_vals if v <= 2000 and v > 0])

            # Jika ada 2 besar dan 2 kecil, assign yang benar
            if len(besar) >= 2 and len(kecil) >= 2:
                # Urutkan biar CD yang besar, SA yang kedua
                r["CD_Omzet"] = besar[0]
                r["SA_Omzet"] = besar[1] if len(besar)>1 else 0
                r["CD_Pasien"] = kecil[0]
                r["SA_Pasien"] = kecil[1] if len(kecil)>1 else 0
            elif len(besar) == 1 and len(kecil) == 1:
                # Kasus cuma 1 omzet 1 pasien
                r["CD_Omzet"] = besar[0]
                r["CD_Pasien"] = kecil[0]
                r["SA_Omzet"] = 0
                r["SA_Pasien"] = 0

            r["Total_Omzet"] = r["CD_Omzet"] + r["SA_Omzet"]
            r["Total_Pasien"] = r["CD_Pasien"] + r["SA_Pasien"]
        except:
            pass
        return r
    return df_input.apply(fix_row, axis=1)

BULAN_FULL = {"JANU":"JANUARI","JAN":"JANUARI","FEBR":"FEBRUARI","FEB":"FEBRUARI","MAR":"MARET","MARET":"MARET","APRIL":"APRIL","APR":"APRIL","MEI":"MEI","JUNI":"JUNI","JUN":"JUNI","JULI":"JULI","JUL":"JULI","AGUS":"AGUSTUS","AGU":"AGUSTUS","SEPT":"SEPTEMBER","SEP":"SEPTEMBER","OKTO":"OKTOBER","OKT":"OKTOBER","NOPE":"NOVEMBER","NOV":"NOVEMBER","DESE":"DESEMBER","DES":"DESEMBER"}
BULAN_ANGKA = {"JANUARI":1,"FEBRUARI":2,"MARET":3,"APRIL":4,"MEI":5,"JUNI":6,"JULI":7,"AGUSTUS":8,"SEPTEMBER":9,"OKTOBER":10,"NOVEMBER":11,"DESEMBER":12}
def normalize_periode(s):
    s=str(s).upper().strip(); tahun="".join([c for c in s if c.isdigit()])[-4:]; huruf="".join([c for c in s if c.isalpha()]); full=BULAN_FULL.get(huruf[:4],huruf); return f"{full}-{tahun}" if tahun else full
def parse_date(s):
    try: n=normalize_periode(s); return pd.Timestamp(year=int(n.split("-")[1]), month=BULAN_ANGKA.get(n.split("-")[0],1), day=1)
    except: return pd.Timestamp(2026,1,1)
def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0

def parse_numbers_smart(nums_list):
    nums = [to_int(x) for x in nums_list]
    besar = sorted([n for n in nums if n > 5000], reverse=True)
    kecil = sorted([n for n in nums if n >0 and n <= 2000])
    cd_omzet = besar[0] if len(besar)>0 else 0
    sa_omzet = besar[1] if len(besar)>1 else 0
    cd_pasien = kecil[0] if len(kecil)>0 else 0
    sa_pasien = kecil[1] if len(kecil)>1 else 0
    return cd_omzet, cd_pasien, sa_omzet, sa_pasien

def parse_pdf(path, label):
    doc=fitz.open(path); parsed=[]
    is_new=any("CIK DI TIRO" in doc[p].get_text("text") for p in range(min(2,len(doc))))
    if is_new:
        for p in range(len(doc)):
            lines=[l.strip() for l in doc[p].get_text("text").splitlines() if l.strip()!='']; pic="UNKNOWN"
            for i,l in enumerate(lines):
                if "Tanggal" in l or "s/d" in l:
                    for k in range(i+1,min(i+6,len(lines))):
                        c=lines[k].strip()
                        if len(c)>3 and c.upper()==c and "JASA" not in c and len(c.split())<=4 and not any(ch.isdigit() for ch in c): pic=c.title(); break
                    break
            i=0
            while i < len(lines):
                l=lines[i]; cd=''.join(c for c in l if c.isdigit())
                if len(cd)==10 and l.replace('.','').isdigit():
                    kode=cd; nama=lines[i+1] if i+1<len(lines) and any(c.isalpha() for c in lines[i+1]) else ""; off=2 if nama else 1
                    if i+off < len(lines) and lines[i+off] in ["6","7","8","9","10","11","12","5","4","3","2","1"]: off+=1
                    nums=[]; j=i+off
                    while j < len(lines) and len(nums)<12:
                        cur=lines[j]; c2=''.join(c for c in cur if c.isdigit())
                        if len(c2)==10 and cur.replace('.','').isdigit(): break
                        if cur.upper().startswith("TOTAL"): break
                        tc=cur.replace('.','').replace(',','');
                        if tc.isdigit(): nums.append(tc)
                        j+=1
                    if len(nums)>=2:
                        co, cp, so, sp = parse_numbers_smart(nums)
                        parsed.append({"Periode":normalize_periode(label),"PIC":pic,"Kode_Dokter":kode,"Nama_Dokter":nama[:80],"CD_Omzet":co,"CD_Pasien":cp,"SA_Omzet":so,"SA_Pasien":sp,"Total_Omzet":co+so,"Total_Pasien":cp+sp})
                    i=j-1
                i+=1
    else:
        for txt in [doc[p].get_text("text") for p in range(len(doc))]:
            lines=txt.splitlines(); pic="UNKNOWN"
            for i in range(len(lines)):
                if "Tanggal" in lines[i] or "s/d" in lines[i]:
                    for k in range(i+1,min(i+6,len(lines))):
                        c=lines[k].strip()
                        if len(c)>5 and c.upper()==c and "JASA" not in c and not any(ch.isdigit() for ch in c): pic=c; break
                    break
            buf=""
            def flush(b,p):
                if not b: return None
                t=b.split(); idx=-1; k=""
                for ti,x in enumerate(t):
                    if len(x)==10 and x.isdigit(): idx=ti; k=x; break
                if idx==-1: return None
                aft=t[idx+1:]; bn=-1
                for j,tt in enumerate(aft):
                    if tt in ["6","7","8","9","10","11","12","1","2","3","4","5"]: bn=j; break
                if bn==-1: return None
                nama=" ".join(aft[:bn]); nums=[x for x in aft[bn+1:] if x.replace('.','').replace(',','').isdigit()]
                if len(nums)>=2:
                    co, cp, so, sp = parse_numbers_smart(nums)
                    return {"Periode":normalize_periode(label),"PIC":p.title(),"Kode_Dokter":k,"Nama_Dokter":nama[:80],"CD_Omzet":co,"CD_Pasien":cp,"SA_Omzet":so,"SA_Pasien":sp,"Total_Omzet":co+so,"Total_Pasien":cp+sp}
            for line in lines:
                line=line.strip()
                if not line or ("Kode" in line and "Dokter" in line): continue
                if line.startswith("TOTAL") and len(line)<30:
                    r=flush(buf,pic)
                    if r: parsed.append(r)
                    buf=""; continue
                if any(len(p)==10 and p.isdigit() for p in line.split()):
                    r=flush(buf,pic)
                    if r: parsed.append(r)
                    buf=line
                else:
                    if buf: buf+=" "+line
            r=flush(buf,pic)
            if r: parsed.append(r)
    return pd.DataFrame(parsed)

# HEADER
c_head1, c_head2 = st.columns([6,1])
with c_head1:
    st.markdown(f'''<div class="gradient-header"><div><h1 style="margin:0;font-size:26px; font-weight:800;">MONITORING V18.7 REPAIR 1 KLIK</h1><p style="margin:4px 0 0 0;opacity:0.95">Fix Pasien Tanpa Upload Ulang</p></div></div>''', unsafe_allow_html=True)
with c_head2:
    if st.button("🚪 Logout"): st.session_state.authenticated=False; st.rerun()

with st.sidebar:
    st.markdown("### 📂 Upload Data")
    multi = st.file_uploader("Pilih PDF", type=["pdf"], accept_multiple_files=True)
    if multi and st.button("💾 SIMPAN", type="primary", use_container_width=True):
        all_new=[]
        for up in multi:
            fname=up.name.upper(); import re; th=re.findall(r'20\d{2}', fname); ths=th[0] if th else "2026"; pg="OKTOBER-2026"
            for b in BULAN_FULL.keys():
                if b in fname: pg=f"{BULAN_FULL[b]}-{ths}"; break
            open("temp.pdf","wb").write(up.getbuffer())
            df_t=parse_pdf("temp.pdf", pg); all_new.append(df_t)
        if all_new:
            df_m=pd.concat(all_new, ignore_index=True)
            if os.path.exists(DB_FILE):
                old=pd.read_excel(DB_FILE); old["Periode"]=old["Periode"].apply(normalize_periode)
                for per in df_m["Periode"].unique(): old=old[old["Periode"]!=per]
                final=pd.concat([old, df_m], ignore_index=True)
            else: final=df_m
            final.to_excel(DB_FILE, index=False); push_to_github(DB_FILE); st.rerun()

    st.divider()
    st.markdown("### 🛠️ Perbaikan 1 Klik")
    st.caption("Jika jumlah pasien tidak sesuai, klik tombol ini. Tidak perlu upload ulang!")
    if st.button("🔧 PERBAIKI DATA LAMA SEKARANG", type="primary", use_container_width=True):
        if os.path.exists(DB_FILE):
            df_old = pd.read_excel(DB_FILE)
            df_fixed = repair_existing_data_smart(df_old)
            df_fixed.to_excel(DB_FILE, index=False)
            push_to_github(DB_FILE)
            st.success(f"✅ Berhasil diperbaiki {len(df_fixed)} baris! Reload...")
            st.rerun()
        else:
            st.error("File database tidak ada")

if not os.path.exists(DB_FILE): st.info("Upload PDF dulu kak"); st.stop()
df=pd.read_excel(DB_FILE); df["Periode"]=df["Periode"].apply(normalize_periode); df["SortDate"]=df["Periode"].apply(parse_date); df=df.sort_values("SortDate")
df["Kode_Dokter"]=df["Kode_Dokter"].astype(str)
df["Tahun"]=df["SortDate"].dt.year

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
    st.write(f"#### 🏆 Ranking - {sort_by}")
    df_rank=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values(sort_by,ascending=False).reset_index(drop=True)
    df_rank.insert(0,"Rank",range(1,len(df_rank)+1))
    df_display = df_rank.copy()
    df_display["Total_Omzet"] = df_display["Total_Omzet"].apply(fmt_titik)
    df_display["Total_Pasien"] = df_display["Total_Pasien"].apply(fmt_titik)
    def style_top10(row):
        r = row['Rank']
        if r == 1: return ['background-color: #FFD700; font-weight: 900;'] * len(row)
        elif r == 2: return ['background-color: #e5e7eb; font-weight: 800;'] * len(row)
        elif r == 3: return ['background-color: #fdba74; font-weight: 800;'] * len(row)
        elif 4 <= r <= 10: return ['background-color: #e0f2fe; font-weight: 700;'] * len(row)
        else: return ['background-color: white;'] * len(row)
    st.dataframe(df_display.style.apply(style_top10, axis=1), use_container_width=True, hide_index=True, height=700)
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔄 Banding Bulan")
    mode = st.radio("Mode", ["2 Bulan", "Rentang"], horizontal=True, key="mode_v187")
    if mode == "2 Bulan":
        c1,c2 = st.columns(2)
        with c1: bulan1 = st.selectbox("Bulan Lama", periode_urut, index=max(0,len(periode_urut)-2), key="b1_v187")
        with c2: bulan2 = st.selectbox("Bulan Baru", periode_urut, index=len(periode_urut)-1, key="b2_v187")
        if bulan1!=bulan2:
            df_b1 = df_f[df_f["Periode"]==bulan1].groupby("Kode_Dokter", as_index=False).agg(Nama=("Nama_Dokter","first"), Omzet_1=("Total_Omzet","sum"), Pasien_1=("Total_Pasien","sum"))
            df_b2 = df_f[df_f["Periode"]==bulan2].groupby("Kode_Dokter", as_index=False).agg(Omzet_2=("Total_Omzet","sum"), Pasien_2=("Total_Pasien","sum"))
            df_comp = pd.merge(df_b1, df_b2, on="Kode_Dokter", how="outer").fillna(0)
            df_comp["Selisih"] = df_comp["Omzet_2"] - df_comp["Omzet_1"]
            df_comp = df_comp.sort_values("Selisih", ascending=False)
            df_disp = df_comp.copy()
            for c in ["Omzet_1","Omzet_2","Selisih"]: df_disp[c] = df_disp[c].apply(fmt_titik)
            st.dataframe(df_disp, use_container_width=True, hide_index=True, height=500)
    else:
        rentang = st.multiselect("Pilih Bulan", periode_urut, default=periode_urut, key="rentang_v187")
        if len(rentang)>=2:
            rentang_sorted = [p for p in periode_urut if p in rentang]
            df_rentang = df_f[df_f["Periode"].isin(rentang_sorted)].groupby("Kode_Dokter", as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"), Total_Rentang=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum"))
            df_rentang = df_rentang.sort_values("Total_Rentang", ascending=False)
            df_r_disp = df_rentang.copy()
            df_r_disp["Total_Rentang"] = df_r_disp["Total_Rentang"].apply(fmt_titik)
            df_r_disp["Total_Pasien"] = df_r_disp["Total_Pasien"].apply(fmt_titik)
            st.dataframe(df_r_disp, use_container_width=True, hide_index=True, height=600)
    st.markdown('</div>', unsafe_allow_html=True)

with tab4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👥 Analisa PIC")
    df_pic = df_f.groupby("PIC", as_index=False).agg(Jml_Dokter=("Kode_Dokter","nunique"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values("Total_Omzet", ascending=False)
    df_pic_disp = df_pic.copy()
    df_pic_disp["Total_Omzet"] = df_pic_disp["Total_Omzet"].apply(fmt_titik)
    df_pic_disp["Total_Pasien"] = df_pic_disp["Total_Pasien"].apply(fmt_titik)
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
            st.write(f"❌ Hilang: {len(hilang)}")
            if hilang:
                df_h = df[(df["Kode_Dokter"].isin(hilang)) & (df["Periode"]==bl_lama)][["Kode_Dokter","Nama_Dokter"]].drop_duplicates()
                st.dataframe(df_h, hide_index=True, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab6:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔥 Heatmap - Sudah Fix")
    mode_heat = st.radio("Tampilkan:", ["💰 OMZET", "👥 PASIEN"], horizontal=True, key="heat_v187")
    col_val = "Total_Omzet" if "OMZET" in mode_heat else "Total_Pasien"
    pivot = df_f.pivot_table(index="Nama_Dokter", columns="Periode", values=col_val, aggfunc="sum", fill_value=0)
    pivot = pivot.reindex(columns=[p for p in periode_urut if p in pivot.columns])
    if "OMZET" in mode_heat:
        st.dataframe(pivot.style.background_gradient(cmap="Reds").format(fmt_titik), use_container_width=True, height=600)
    else:
        st.dataframe(pivot.style.background_gradient(cmap="Blues").format(lambda x: f"{int(x)}"), use_container_width=True, height=600)
    st.markdown('</div>', unsafe_allow_html=True)

with tab7:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👨‍⚕️ Detail Dokter")
    df_dokter_list = df_f.groupby("Kode_Dokter", as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"), PIC=("PIC","first"), Total_Omzet=("Total_Omzet","sum")).sort_values("Nama_Dokter")
    df_dokter_list["Label"] = df_dokter_list["Nama_Dokter"] + " | KODE: " + df_dokter_list["Kode_Dokter"] + " | Rp " + df_dokter_list["Total_Omzet"].apply(fmt_titik)
    selected_label = st.selectbox("🔎 Cari Dokter", ["-- Pilih Dokter --"] + df_dokter_list["Label"].tolist(), index=0, key="detail_v187")
    if selected_label!= "-- Pilih Dokter --":
        try: selected_kode = selected_label.split("KODE: ")[1].split(" |")[0].strip()
        except: selected_kode = ""
        res = df_f[df_f["Kode_Dokter"] == selected_kode].copy()
        if not res.empty:
            st.success(f"✅ {res['Nama_Dokter'].iloc[0]} | Kode: {selected_kode}")
            hist=res.groupby(["SortDate","Periode"],as_index=False).agg(Total=("Total_Omzet","sum"),Pasien=("Total_Pasien","sum")).sort_values("SortDate")
            hist["Periode"] = pd.Categorical(hist["Periode"], categories=periode_urut, ordered=True)
            hist = hist.sort_values("Periode")
            st.line_chart(hist.set_index("Periode")[["Total"]])
            hist_disp = hist.copy()
            hist_disp["Total"] = hist_disp["Total"].apply(fmt_titik)
            hist_disp["Pasien"] = hist_disp["Pasien"].apply(lambda x: f"{int(x)}")
            st.dataframe(hist_disp[["Periode","Total","Pasien"]], use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)
