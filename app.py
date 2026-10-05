import fitz, pandas as pd, os, base64, requests, streamlit as st

st.set_page_config(page_title="Pramita V18.2 Fix Only", layout="wide", page_icon="🚀")
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
def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0
def parse_pdf(path, label):
    doc=fitz.open(path); parsed=[]; is_new=any("CIK DI TIRO" in doc[p].get_text("text") for p in range(min(2,len(doc))))
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
                    if len(nums)>=4:
                        co=to_int(nums[0]); cp=to_int(nums[3]) if len(nums)>3 else 0; so=to_int(nums[4]) if len(nums)>4 else 0; sp=to_int(nums[7]) if len(nums)>7 else 0
                        if cp>10000: cp=to_int(nums[1]) if len(nums)>1 else 0
                        if sp>10000: sp=to_int(nums[5]) if len(nums)>5 else 0
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
                if len(nums)>=8: return {"Periode":normalize_periode(label),"PIC":p.title(),"Kode_Dokter":k,"Nama_Dokter":nama[:80],"CD_Omzet":to_int(nums[0]),"CD_Pasien":to_int(nums[3]),"SA_Omzet":to_int(nums[4]),"SA_Pasien":to_int(nums[7]),"Total_Omzet":to_int(nums[0])+to_int(nums[4]),"Total_Pasien":to_int(nums[3])+to_int(nums[7])}
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

c_head1, c_head2 = st.columns([6,1])
with c_head1:
    st.markdown(f'''<div class="gradient-header"><div style="display:flex; align-items:center; gap:18px;"><div style="background:white; border-radius:12px; padding:6px 14px; display:flex; align-items:center;"><span style="color:#dc2626; font-weight:900; font-size:22px; letter-spacing:1px;">PRAMITA</span><span style="color:#dc2626; font-style:italic; margin-left:8px; font-weight:600;">Lab</span></div><div><h1 style="margin:0;font-size:26px; font-weight:800;">MONITORING KINERJA DOKTER V18.2</h1><p style="margin:4px 0 0 0;opacity:0.95">Tampilan V18.0 Cantik + Fix Banding & Detail Only</p></div></div></div>''', unsafe_allow_html=True)
with c_head2:
    if st.button("🚪 Logout"): st.session_state.authenticated=False; st.rerun()

with st.sidebar:
    st.markdown("### 📂 Upload Data")
    multi = st.file_uploader("Pilih PDF (bisa banyak)", type=["pdf"], accept_multiple_files=True)
    if multi and st.button("💾 SIMPAN PERMANEN", type="primary", use_container_width=True):
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
            final["Total_Omzet"]=final["CD_Omzet"]+final["SA_Omzet"]; final["Total_Pasien"]=final["CD_Pasien"]+final["SA_Pasien"]
            final.to_excel(DB_FILE, index=False); push_to_github(DB_FILE); st.rerun()

if not os.path.exists(DB_FILE): st.info("Upload PDF dulu kak"); st.stop()
df=pd.read_excel(DB_FILE); df["Periode"]=df["Periode"].apply(normalize_periode); df["SortDate"]=df["Periode"].apply(parse_date); df=df.sort_values("SortDate")
df["Kode_Dokter"]=df["Kode_Dokter"].astype(str); df["Total_Omzet"]=df["CD_Omzet"]+df["SA_Omzet"]; df["Total_Pasien"]=df["CD_Pasien"]+df["SA_Pasien"]
df["Tahun"]=df["SortDate"].dt.year

with st.sidebar:
    st.markdown("### 🔎 Filter")
    tahun_list=sorted(df["Tahun"].unique().tolist()); sel_tahun=st.multiselect("Tahun", tahun_list, default=tahun_list)
    periode_list_sorted=df.sort_values("SortDate")["Periode"].unique().tolist()
    sel_periode=st.multiselect("Periode", periode_list_sorted, default=periode_list_sorted)
    pic_list=["Semua"]+sorted(df["PIC"].dropna().unique().tolist()); sel_pic=st.selectbox("PIC", pic_list)
    sort_by=st.selectbox("Ranking Urut", ["Total_Omzet","Total_Pasien","CD_Omzet","SA_Omzet"])

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
    df_month_disp["Total_Pasien"] = df_month_disp["Total_Pasien"].apply(fmt_titik)
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
    df_display["Total_Pasien"] = df_display["Total_Pasien"].apply(fmt_titik)
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
    st.write("### 🔄 Perbandingan Bulan - FIXED")
    mode = st.radio("Mode", ["2 Bulan (SEPTEMBER vs AGUSTUS)", "Rentang (JANUARI sampai OKTOBER)"], horizontal=True, key="mode_banding")

    if mode == "2 Bulan (SEPTEMBER vs AGUSTUS)":
        c1,c2 = st.columns(2)
        with c1: bulan1 = st.selectbox("Bulan Lama", periode_urut, index=max(0,len(periode_urut)-2), key="b1_fix")
        with c2: bulan2 = st.selectbox("Bulan Baru", periode_urut, index=len(periode_urut)-1, key="b2_fix")
        if bulan1!=bulan2:
            df_b1 = df_f[df_f["Periode"]==bulan1].groupby("Kode_Dokter", as_index=False).agg(Nama=("Nama_Dokter","first"), Omzet_1=("Total_Omzet","sum"), Pasien_1=("Total_Pasien","sum"))
            df_b2 = df_f[df_f["Periode"]==bulan2].groupby("Kode_Dokter", as_index=False).agg(Omzet_2=("Total_Omzet","sum"), Pasien_2=("Total_Pasien","sum"))
            df_comp = pd.merge(df_b1, df_b2, on="Kode_Dokter", how="outer").fillna(0)
            df_comp["Selisih_Omzet"] = df_comp["Omzet_2"] - df_comp["Omzet_1"]
            df_comp["Persen"] = df_comp.apply(lambda r: (r["Selisih_Omzet"]/r["Omzet_1"]*100) if r["Omzet_1"]!=0 else (100 if r["Omzet_2"]>0 else 0), axis=1)
            df_comp["Status"] = df_comp["Persen"].apply(lambda x: "📈 NAIK" if x>10 else "📉 TURUN" if x<-10 else "➖ STABIL")
            df_comp = df_comp.sort_values("Persen", ascending=False)
            df_disp = df_comp.copy()
            df_disp["Omzet_1"] = df_disp["Omzet_1"].apply(fmt_titik)
            df_disp["Omzet_2"] = df_disp["Omzet_2"].apply(fmt_titik)
            df_disp["Selisih_Omzet"] = df_disp["Selisih_Omzet"].apply(fmt_titik)
            df_disp["Persen_fmt"] = df_disp["Persen"].apply(lambda x: f"{x:.1f}%")
            st.dataframe(df_disp[["Nama","Kode_Dokter","Omzet_1","Omzet_2","Selisih_Omzet","Persen_fmt","Status"]], use_container_width=True, hide_index=True, height=500)
    else:
        st.write("**Pilih Rentang (misal JANUARI-2026 sampai OKTOBER-2026)**")
        rentang = st.multiselect("Pilih Bulan", periode_urut, default=periode_urut, key="rentang_fix")
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
        else:
            st.warning("Pilih minimal 2 bulan")
    st.markdown('</div>', unsafe_allow_html=True)

with tab4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👥 Analisa PIC - Siapa PIC Paling Jago?")
    df_pic = df_f.groupby("PIC", as_index=False).agg(Jml_Dokter=("Kode_Dokter","nunique"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum"),CD_Omzet=("CD_Omzet","sum"),SA_Omzet=("SA_Omzet","sum")).sort_values("Total_Omzet", ascending=False)
    df_pic["Avg_per_Dokter"] = df_pic["Total_Omzet"] / df_pic["Jml_Dokter"]
    st.bar_chart(df_pic.set_index("PIC")[["Total_Omzet"]])
    df_pic_disp = df_pic.copy()
    for col in ["Total_Omzet","CD_Omzet","SA_Omzet","Avg_per_Dokter"]: df_pic_disp[col] = df_pic_disp[col].apply(fmt_titik)
    df_pic_disp["Total_Pasien"] = df_pic_disp["Total_Pasien"].apply(fmt_titik)
    st.dataframe(df_pic_disp, use_container_width=True, hide_index=True)
    sel_pic_detail = st.selectbox("Pilih PIC untuk lihat dokter nya", df_pic["PIC"].tolist())
    if sel_pic_detail:
        df_pic_dok = df_f[df_f["PIC"]==sel_pic_detail].groupby("Kode_Dokter", as_index=False).agg(Nama=("Nama_Dokter","first"), Omzet=("Total_Omzet","sum"), Pasien=("Total_Pasien","sum")).sort_values("Omzet", ascending=False)
        df_pic_dok["Omzet"] = df_pic_dok["Omzet"].apply(fmt_titik)
        st.dataframe(df_pic_dok, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab5:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🆕 Dokter Baru / Hilang (Churn Detection)")
    if len(periode_urut) >= 2:
        bulan_lama = periode_urut[-2]; bulan_baru = periode_urut[-1]
        st.info(f"Membandingkan **{bulan_lama}** vs **{bulan_baru}**")
        dokter_lama = set(df[df["Periode"]==bulan_lama]["Kode_Dokter"].unique())
        dokter_baru = set(df[df["Periode"]==bulan_baru]["Kode_Dokter"].unique())
        baru = dokter_baru - dokter_lama; hilang = dokter_lama - dokter_baru; tetap = dokter_baru & dokter_lama
        k1,k2,k3 = st.columns(3)
        k1.metric("✅ Tetap", f"{len(tetap)}"); k2.metric("🆕 Baru", f"{len(baru)}", delta=f"{len(baru)}"); k3.metric("❌ Hilang", f"{len(hilang)}", delta=f"-{len(hilang)}", delta_color="inverse")
        colA,colB = st.columns(2)
        with colA:
            st.write(f"**🆕 Baru di {bulan_baru}:**")
            if baru:
                df_baru = df[df["Kode_Dokter"].isin(baru) & (df["Periode"]==bulan_baru)][["Kode_Dokter","Nama_Dokter","PIC","Total_Omzet"]].drop_duplicates()
                df_baru["Total_Omzet"] = df_baru["Total_Omzet"].apply(fmt_titik)
                st.dataframe(df_baru, use_container_width=True, hide_index=True)
        with colB:
            st.write(f"**❌ Hilang:**")
            if hilang:
                df_hilang = df[df["Kode_Dokter"].isin(hilang) & (df["Periode"]==bulan_lama)][["Kode_Dokter","Nama_Dokter","PIC","Total_Omzet"]].drop_duplicates()
                df_hilang["Total_Omzet"] = df_hilang["Total_Omzet"].apply(fmt_titik)
                st.dataframe(df_hilang, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab6:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 🔥 Heatmap Pasien - Bulan Sepi / Rame")
    pivot = df_f.pivot_table(index="Nama_Dokter", columns="Periode", values="Total_Omzet", aggfunc="sum", fill_value=0)
    pivot = pivot.reindex(columns=[p for p in periode_urut if p in pivot.columns])
    st.dataframe(pivot.style.background_gradient(cmap="RdYlGn_r").format(lambda x: fmt_titik(x)), use_container_width=True, height=600)
    st.markdown('</div>', unsafe_allow_html=True)

with tab7:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 👨‍⚕️ Detail Dokter - FIXED Anti Error")
    try:
        df_dokter_list = df_f.groupby("Kode_Dokter", as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"), PIC=("PIC","first"), Total_Omzet=("Total_Omzet","sum")).sort_values("Nama_Dokter")
        df_dokter_list["Label"] = df_dokter_list["Nama_Dokter"] + " | KODE: " + df_dokter_list["Kode_Dokter"] + " | PIC: " + df_dokter_list["PIC"] + " | Rp " + df_dokter_list["Total_Omzet"].apply(fmt_titik)

        selected_label = st.selectbox("🔎 Cari & Pilih Dokter (ketik nama atau kode)", ["-- Pilih Dokter --"] + df_dokter_list["Label"].tolist(), index=0, key="detail_fix_v182")

        if selected_label!= "-- Pilih Dokter --":
            try:
                selected_kode = selected_label.split("KODE: ")[1].split(" |")[0].strip()
            except:
                selected_kode = selected_label.split("|")[1].strip() if "|" in selected_label else ""
            res = df_f[df_f["Kode_Dokter"] == selected_kode].copy()
            if not res.empty:
                nama_tampil = res["Nama_Dokter"].iloc[0]
                st.success(f"✅ Dokter terpilih: **{nama_tampil}** | Kode: **{selected_kode}** | PIC: {res['PIC'].iloc[0]}")
                hist=res.groupby(["SortDate","Periode"],as_index=False).agg(Total=("Total_Omzet","sum"),Pasien=("Total_Pasien","sum")).sort_values("SortDate")
                hist["Periode"] = pd.Categorical(hist["Periode"], categories=periode_urut, ordered=True)
                hist = hist.sort_values("Periode")
                st.line_chart(hist.set_index("Periode")[["Total"]])
                hist_disp = hist.copy()
                hist_disp["Total"] = hist_disp["Total"].apply(fmt_titik)
                hist_disp["Pasien"] = hist_disp["Pasien"].apply(fmt_titik)
                st.dataframe(hist_disp[["Periode","Total","Pasien"]], use_container_width=True, hide_index=True)
                detail_bulan = res.groupby(["Periode","SortDate"], as_index=False).agg(CD=("CD_Omzet","sum"),SA=("SA_Omzet","sum"),Total=("Total_Omzet","sum"),Pasien=("Total_Pasien","sum")).sort_values("SortDate")
                detail_bulan_disp = detail_bulan.copy()
                for c in ["CD","SA","Total"]: detail_bulan_disp[c] = detail_bulan_disp[c].apply(fmt_titik)
                st.dataframe(detail_bulan_disp[["Periode","CD","SA","Total","Pasien"]], use_container_width=True, hide_index=True)
        else:
            st.info("👆 Silahkan pilih dokter di atas. Ketik nama atau kode dokter, nanti muncul semua pilihan dengan kode nya biar tidak salah.")
            preview = df_dokter_list.copy()
            preview["Total_Omzet"] = preview["Total_Omzet"].apply(fmt_titik)
            st.dataframe(preview[["Kode_Dokter","Nama_Dokter","PIC","Total_Omzet"]], use_container_width=True, hide_index=True, height=300)
    except Exception as e:
        st.error(f"Detail error: {e}")

    st.markdown('</div>', unsafe_allow_html=True)
