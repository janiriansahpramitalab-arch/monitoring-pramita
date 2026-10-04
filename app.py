import fitz, pandas as pd, os, base64, requests, streamlit as st

st.set_page_config(page_title="Pramita Modern V16.1 Logo", layout="wide", page_icon="💎")
DB_FILE = "Database_Monitoring_Pramita.xlsx"
LOGO_FILE = "logo-pramita.png" # <-- upload logo dengan nama ini
HIDDEN_PIC = ["YOHANA DEWI RATIH", "NARINDRA NATA KUNTHARA"]
HIDDEN_KODE = ["2741002000"]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
html, body, [class*="css"] {font-family: 'Inter', sans-serif;}
.main {background-color: #f6f8fb;}
div[data-testid="metric-container"] {
    background: white; border-radius: 16px; padding: 20px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #eef2f7;
}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {background: white; border-radius: 10px; padding: 10px 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.04);}
.stTabs [aria-selected="true"] {background: linear-gradient(135deg,#2563eb,#1e40af); color:white!important;}
.card {background: white; border-radius: 18px; padding: 22px; box-shadow: 0 8px 24px rgba(0,0,0,0.06); border: 1px solid #eef2f7; margin-bottom:16px;}
.gradient-header {
    background: linear-gradient(135deg,#1e3a8a 0%, #2563eb 50%, #06b6d4 100%);
    padding: 24px 28px; border-radius: 20px; color: white; margin-bottom: 20px;
    display:flex; align-items:center; gap:20px;
}
.logo-img {width: 70px; height: 70px; background: white; border-radius: 14px; padding: 8px; object-fit: contain;}
</style>
""", unsafe_allow_html=True)

def get_logo_base64():
    if os.path.exists(LOGO_FILE):
        with open(LOGO_FILE, "rb") as f: return base64.b64encode(f.read()).decode()
    return None

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

BULAN_FULL = {"JANU":"JANUARI","JAN":"JANUARI","FEBR":"FEBRUARI","FEB":"FEBRUARI","MAR":"MARET","APRIL":"APRIL","APR":"APRIL","MEI":"MEI","JUNI":"JUNI","JUN":"JUNI","JULI":"JULI","JUL":"JULI","AGUS":"AGUSTUS","AGU":"AGUSTUS","SEPT":"SEPTEMBER","SEP":"SEPTEMBER","OKTO":"OKTOBER","OKT":"OKTOBER","NOPE":"NOVEMBER","NOV":"NOVEMBER","DESE":"DESEMBER","DES":"DESEMBER"}
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
                    if i+off < len(lines) and lines[i+off] in ["6","7","8","9","10","11","12"]: off+=1
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

# HEADER DENGAN LOGO
logo_b64 = get_logo_base64()
if logo_b64:
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" class="logo-img">'
else:
    logo_html = '<div class="logo-img" style="display:flex;align-items:center;justify-content:center;font-size:32px">🧬</div>'

st.markdown(f'<div class="gradient-header">{logo_html}<div><h1 style="margin:0;font-size:28px">PRAMITA LAB</h1><p style="margin:4px 0 0 0;opacity:0.9">Monitoring Kinerja Dokter - Modern Dashboard V16.1 | Auto-Permanen</p></div></div>', unsafe_allow_html=True)

with st.sidebar:
    if logo_b64: st.image(LOGO_FILE, width=180)
    st.markdown("### 📂 Upload")
    multi = st.file_uploader("PDF Bulanan", type=["pdf"], accept_multiple_files=True)
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
            final.to_excel(DB_FILE, index=False); push_to_github(DB_FILE); st.rerun()

if not os.path.exists(DB_FILE): st.info("Upload PDF dulu kak"); st.stop()
df=pd.read_excel(DB_FILE); df["Periode"]=df["Periode"].apply(normalize_periode); df["SortDate"]=df["Periode"].apply(parse_date); df=df.sort_values("SortDate")
df["Kode_Dokter"]=df["Kode_Dokter"].astype(str); df["Total_Omzet"]=df["CD_Omzet"]+df["SA_Omzet"]; df["Total_Pasien"]=df["CD_Pasien"]+df["SA_Pasien"]
df["Tahun"]=df["SortDate"].dt.year

with st.sidebar:
    st.divider(); st.markdown("### 🔎 Filter")
    tahun_list=sorted(df["Tahun"].unique().tolist()); sel_tahun=st.multiselect("Tahun", tahun_list, default=tahun_list)
    periode_list=df.sort_values("SortDate")["Periode"].unique().tolist(); sel_periode=st.multiselect("Periode", periode_list, default=periode_list)
    cabang_opsi=st.selectbox("Cabang", ["Semua","Cik Di Tiro","Sultan Agung"])
    pic_list=["Semua"]+sorted(df["PIC"].dropna().unique().tolist()); sel_pic=st.selectbox("PIC", pic_list)
    sort_by=st.selectbox("Ranking Urut", ["Total_Omzet","Total_Pasien"])

df_all=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_tahun: df_all=df_all[df_all["Tahun"].isin(sel_tahun)]
if sel_pic!="Semua": df_all=df_all[df_all["PIC"]==sel_pic]
df_kpi=df_all.copy()
df_f=df_all[~df_all["PIC"].str.upper().isin(HIDDEN_PIC)]; df_f=df_f[~df_f["Kode_Dokter"].isin(HIDDEN_KODE)]

total_omzet=df_kpi["Total_Omzet"].sum(); total_pasien=df_kpi["Total_Pasien"].sum(); jml_dokter=df_kpi["Kode_Dokter"].nunique()
df_month=df.groupby("SortDate",as_index=False).agg(Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum")).sort_values("SortDate")
df_month["Growth_O"]=df_month["Total_Omzet"].pct_change()*100; last_o=df_month["Growth_O"].iloc[-1] if len(df_month)>1 else 0
df_month["Growth_P"]=df_month["Total_Pasien"].pct_change()*100; last_p=df_month["Growth_P"].iloc[-1] if len(df_month)>1 else 0

k1,k2,k3,k4=st.columns(4)
k1.metric("💰 TOTAL OMZET", f"Rp {total_omzet/1_000_000_000:.2f} M", f"{last_o:.1f}%")
k2.metric("👥 TOTAL PASIEN", f"{total_pasien:,}", f"{last_p:.1f}%")
k3.metric("🩺 DOKTER", f"{jml_dokter}", f"Tampil {df_f['Kode_Dokter'].nunique()}")
k4.metric("📅 PERIODE", f"{len(periode_list)} Bulan")

tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "🏆 Ranking", "📈 Analytics", "👨‍⚕️ Detail"])

with tab1:
    c1,c2=st.columns([2,1])
    with c1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 📈 Trend Omzet & Pasien (Include Hidden)")
        st.line_chart(df_month.set_index("SortDate")[["Total_Omzet","Total_Pasien"]])
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 🔥 Top Performer")
        top=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama=("Nama_Dokter","first"), Omzet=("Total_Omzet","sum"), Pasien=("Total_Pasien","sum")).sort_values("Omzet",ascending=False).head(3)
        for i,row in top.iterrows():
            medal="🥇" if i==0 else "🥈" if i==1 else "🥉"
            st.write(f"{medal} **{row['Nama'][:28]}**")
            st.caption(f"Rp {row['Omzet']:,} | {row['Pasien']} pasien")
            st.divider()
        st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write(f"#### 🏆 Ranking Modern - {sort_by}")
    df_rank=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"),PIC=("PIC","first"),CD_Omzet=("CD_Omzet","sum"),SA_Omzet=("SA_Omzet","sum"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values(sort_by,ascending=False).reset_index(drop=True)
    df_rank.insert(0,"Rank",range(1,len(df_rank)+1)); df_rank["Medal"]=df_rank["Rank"].apply(lambda x: "🥇" if x==1 else "🥈" if x==2 else "🥉" if x==3 else f"#{x}")
    st.dataframe(df_rank, use_container_width=True, hide_index=True, height=650)
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 📊 Per Cabang")
    df_cabang=df_f.groupby("Periode",as_index=False).agg(CD=("CD_Omzet","sum"), SA=("SA_Omzet","sum")).sort_values("Periode")
    st.bar_chart(df_cabang.set_index("Periode")[["CD","SA"]])
    st.markdown('</div>', unsafe_allow_html=True)

with tab4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    q=st.text_input("🔎 Cari Dokter")
    if q:
        res=df_f[df_f["Nama_Dokter"].str.contains(q,case=False,na=False) | df_f["Kode_Dokter"].str.contains(q,case=False,na=False)].sort_values("SortDate")
        if not res.empty:
            hist=res.groupby("Periode",as_index=False).agg(Total=("Total_Omzet","sum"),Pasien=("Total_Pasien","sum"),Date=("SortDate","first")).sort_values("Date")
            st.line_chart(hist.set_index("Periode")[["Total"]])
            st.dataframe(hist, use_container_width=True, hide_index=True)
        else: st.warning("Tidak ditemukan / hidden")
    st.markdown('</div>', unsafe_allow_html=True)
