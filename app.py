import fitz, pandas as pd, os, base64, requests, streamlit as st
import altair as alt

st.set_page_config(page_title="Pramita V16.7 Urut Fix", layout="wide", page_icon="✅")
DB_FILE = "Database_Monitoring_Pramita.xlsx"
HIDDEN_PIC = ["YOHANA DEWI RATIH", "NARINDRA NATA KUNTHARA"]
HIDDEN_KODE = ["2741002000"]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
html, body, [class*="css"] {font-family: 'Inter', sans-serif;}
.main {background-color: #f6f8fb;}
div[data-testid="metric-container"] {background: white; border-radius: 16px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #eef2f7;}
.gradient-header {background: linear-gradient(135deg,#b91c1c 0%, #dc2626 50%, #ef4444 100%); padding: 18px 24px; border-radius: 18px; color: white; margin-bottom: 20px; display:flex; align-items:center; gap:18px;}
.card {background: white; border-radius: 18px; padding: 22px; box-shadow: 0 8px 24px rgba(0,0,0,0.06); border: 1px solid #eef2f7; margin-bottom:16px;}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {background: white; border-radius: 10px; padding: 10px 20px;}
.stTabs [aria-selected="true"] {background: linear-gradient(135deg,#dc2626,#991b1b); color:white!important;}
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

st.markdown(f'''
<div class="gradient-header">
    <div style="background:white; border-radius:12px; padding:6px 14px; display:flex; align-items:center;">
        <span style="color:#dc2626; font-weight:900; font-size:22px; letter-spacing:1px;">PRAMITA</span>
        <span style="color:#dc2626; font-style:italic; margin-left:8px; font-weight:600;">Lab</span>
    </div>
    <div>
        <h1 style="margin:0;font-size:26px; font-weight:800;">MONITORING KINERJA DOKTER</h1>
        <p style="margin:4px 0 0 0;opacity:0.95">V16.7 Grafik Urut Januari-Desember FIXED</p>
    </div>
</div>
''', unsafe_allow_html=True)

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

# FIX URUT: Buat list urut yang benar JANUARI-2026, FEBRUARI-2026 dst
periode_urut = df.sort_values("SortDate")["Periode"].unique().tolist()
df_month=df.groupby(["SortDate","Periode"],as_index=False).agg(Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum")).sort_values("SortDate")

k1,k2,k3,k4=st.columns(4)
k1.metric("💰 TOTAL OMZET", f"Rp {total_omzet/1_000_000_000:.2f} M")
k2.metric("👥 TOTAL PASIEN", f"{total_pasien:,}")
k3.metric("🩺 DOKTER", f"{jml_dokter}", f"Tampil {df_f['Kode_Dokter'].nunique()}")
k4.metric("📅 PERIODE", f"{len(periode_list_sorted)} Bulan")

tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "🏆 Ranking", "📈 Analytics", "👨‍⚕️ Detail"])

with tab1:
    c1,c2=st.columns([2,1])
    with c1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 📈 Trend Omzet (JANUARI -> DESEMBER URUT)")
        chart1 = alt.Chart(df_month).mark_line(point=True).encode(
            x=alt.X('Periode:N', sort=periode_urut, title='Periode'),
            y=alt.Y('Total_Omzet:Q', title='Omzet'),
            tooltip=['Periode','Total_Omzet']
        ).properties(height=300)
        st.altair_chart(chart1, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.write("### 🔥 Top 3 Performer")
        top=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama=("Nama_Dokter","first"), Omzet=("Total_Omzet","sum"), Pasien=("Total_Pasien","sum")).sort_values("Omzet",ascending=False).head(3)
        for i,row in top.iterrows():
            medal="🥇" if i==0 else "🥈" if i==1 else "🥉"
            st.write(f"{medal} **{row['Nama'][:28]}**")
            st.caption(f"Rp {row['Omzet']:,} | {row['Pasien']} pasien")
            st.divider()
        st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write(f"#### 🏆 Ranking - {sort_by} (Top 10 Highlight)")
    df_rank=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"),PIC=("PIC","first"),CD_Omzet=("CD_Omzet","sum"),SA_Omzet=("SA_Omzet","sum"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum")).sort_values(sort_by,ascending=False).reset_index(drop=True)
    df_rank.insert(0,"Rank",range(1,len(df_rank)+1))
    df_rank["Medal"]=df_rank["Rank"].apply(lambda x: "🥇 JUARA 1" if x==1 else "🥈 JUARA 2" if x==2 else "🥉 JUARA 3" if x==3 else f"⭐ TOP {x}" if x<=10 else f"#{x}")
    def style_top10(row):
        r = row['Rank']
        if r == 1: return ['background-color: #FFD700; color: #78350f; font-weight: 900; font-size: 15px; border-left: 6px solid #b45309;'] * len(row)
        elif r == 2: return ['background-color: #e5e7eb; color: #111827; font-weight: 800; font-size: 14px; border-left: 6px solid #6b7280;'] * len(row)
        elif r == 3: return ['background-color: #fdba74; color: #7c2d12; font-weight: 800; font-size: 14px; border-left: 6px solid #9a3412;'] * len(row)
        elif 4 <= r <= 10: return ['background-color: #e0f2fe; color: #0c4a6e; font-weight: 700; font-size: 13px; border-left: 5px solid #0284c7;'] * len(row)
        else: return ['background-color: white; color: #334155; font-size: 12px;'] * len(row)
    st.dataframe(df_rank.style.apply(style_top10, axis=1), use_container_width=True, hide_index=True, height=700)
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.write("### 📊 Per Cabang (Urut Januari-Desember)")
    df_cabang=df_f.groupby(["SortDate","Periode"],as_index=False).agg(CD=("CD_Omzet","sum"), SA=("SA_Omzet","sum")).sort_values("SortDate")
    chart3 = alt.Chart(df_cabang).transform_fold(['CD','SA'], as_=['Cabang','Omzet']).mark_bar().encode(
        x=alt.X('Periode:N', sort=periode_urut),
        y='Omzet:Q',
        color='Cabang:N',
        tooltip=['Periode','Cabang','Omzet']
    )
    st.altair_chart(chart3, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab4:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    q=st.text_input("🔎 Cari Dokter (nama/kode)")
    if q:
        res=df_f[df_f["Nama_Dokter"].str.contains(q,case=False,na=False) | df_f["Kode_Dokter"].str.contains(q,case=False,na=False)].copy()
        if not res.empty:
            hist=res.groupby(["SortDate","Periode"],as_index=False).agg(Total=("Total_Omzet","sum"),Pasien=("Total_Pasien","sum")).sort_values("SortDate")
            periode_dokter_urut = hist["Periode"].tolist() # sudah urut SortDate
            st.write(f"#### Grafik {q} - URUT JANUARI->DESEMBER")
            chart4 = alt.Chart(hist).mark_line(point=True, color='#2563eb').encode(
                x=alt.X('Periode:N', sort=periode_dokter_urut, title='Periode (JANUARI-2026)'),
                y=alt.Y('Total:Q', title='Omzet'),
                tooltip=['Periode','Total','Pasien']
            ).properties(height=350)
            st.altair_chart(chart4, use_container_width=True)
            display_df = hist.sort_values("SortDate")[["Periode","Total","Pasien"]]
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        else: st.warning("Tidak ditemukan / hidden")
    st.markdown('</div>', unsafe_allow_html=True)
