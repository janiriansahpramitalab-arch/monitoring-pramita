import fitz, pandas as pd, os, base64, requests
import streamlit as st

st.set_page_config(page_title="Monitoring Pramita V14 AUTO", layout="wide", page_icon="📊")
DB_FILE = "Database_Monitoring_Pramita.xlsx"

# ============ AUTO BACKUP KE GITHUB ============
def push_to_github(file_path):
    try:
        token = st.secrets.get("GITHUB_TOKEN")
        repo = st.secrets.get("GITHUB_REPO")
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        if not token or not repo:
            return False, "Secrets GITHUB_TOKEN / REPO belum diisi"

        with open(file_path, "rb") as f:
            content = base64.b64encode(f.read()).decode()

        url = f"https://api.github.com/repos/{repo}/contents/{file_path}"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        r_get = requests.get(url, headers=headers, params={"ref": branch})
        sha = r_get.json().get("sha") if r_get.status_code == 200 else None

        payload = {"message": f"Auto backup {file_path}", "content": content, "branch": branch}
        if sha:
            payload["sha"] = sha

        r_put = requests.put(url, headers=headers, json=payload)
        if r_put.status_code in [200, 201]:
            return True, "Berhasil auto-backup permanen ke Github!"
        else:
            return False, f"Gagal push: {r_put.text[:300]}"
    except Exception as e:
        return False, str(e)

# ============ PARSER ============
BULAN_FULL = {"JANU": "JANUARI", "JAN": "JANUARI","FEBR": "FEBRUARI", "FEB": "FEBRUARI","MARET": "MARET", "MAR": "MARET","APRIL": "APRIL", "APR": "APRIL","MEI": "MEI","JUNI": "JUNI", "JUN": "JUNI","JULI": "JULI", "JUL": "JULI","AGUS": "AGUSTUS", "AGU": "AGUSTUS","SEPT": "SEPTEMBER", "SEP": "SEPTEMBER","OKTO": "OKTOBER", "OKT": "OKTOBER","NOPE": "NOVEMBER", "NOV": "NOVEMBER","DESE": "DESEMBER", "DES": "DESEMBER"}
BULAN_ANGKA = {"JANUARI":1,"FEBRUARI":2,"MARET":3,"APRIL":4,"MEI":5,"JUNI":6,"JULI":7,"AGUSTUS":8,"SEPTEMBER":9,"OKTOBER":10,"NOVEMBER":11,"DESEMBER":12}

def normalize_periode(s):
    s=str(s).upper().strip()
    tahun="".join([c for c in s if c.isdigit()])[-4:]
    huruf="".join([c for c in s if c.isalpha()])
    pre=huruf[:4]
    full=BULAN_FULL.get(pre,huruf)
    return f"{full}-{tahun}" if tahun else full

def parse_date(s):
    try:
        n=normalize_periode(s)
        return pd.Timestamp(year=int(n.split("-")[1]), month=BULAN_ANGKA.get(n.split("-")[0],1), day=1)
    except:
        return pd.Timestamp(2026,1,1)

def to_int(s):
    try:
        return int(str(s).replace('.','').replace(',','').strip())
    except:
        return 0

def parse_pdf(path, label):
    doc=fitz.open(path)
    parsed=[]
    is_new=any("CIK DI TIRO" in doc[p].get_text("text") for p in range(min(2,len(doc))))
    if is_new:
        for p in range(len(doc)):
            lines=[l.strip() for l in doc[p].get_text("text").splitlines() if l.strip()!='']
            pic="UNKNOWN"
            for i,l in enumerate(lines):
                if "Tanggal" in l or "s/d" in l:
                    for k in range(i+1,min(i+6,len(lines))):
                        c=lines[k].strip()
                        if len(c)>3 and c.upper()==c and "JASA" not in c and len(c.split())<=4 and not any(ch.isdigit() for ch in c):
                            pic=c.title()
                            break
                    break
            i=0
            while i < len(lines):
                l=lines[i]
                cd=''.join(c for c in l if c.isdigit())
                if len(cd)==10 and l.replace('.','').isdigit():
                    kode=cd
                    nama=lines[i+1] if i+1<len(lines) and any(c.isalpha() for c in lines[i+1]) else ""
                    off=2 if nama else 1
                    if i+off < len(lines) and lines[i+off] in ["6","7","8","9","10","11","12","5","4","3","2","1"]:
                        off+=1
                    nums=[]
                    j=i+off
                    while j < len(lines) and len(nums)<12:
                        cur=lines[j]
                        c2=''.join(c for c in cur if c.isdigit())
                        if len(c2)==10 and cur.replace('.','').isdigit():
                            break
                        if cur.upper().startswith("TOTAL"):
                            break
                        tc=cur.replace('.','').replace(',','')
                        if tc.isdigit():
                            nums.append(tc)
                        j+=1
                    if len(nums)>=4:
                        co=to_int(nums[0])
                        cp=to_int(nums[3]) if len(nums)>3 else 0
                        so=to_int(nums[4]) if len(nums)>4 else 0
                        sp=to_int(nums[7]) if len(nums)>7 else 0
                        if cp>10000:
                            cp=to_int(nums[1]) if len(nums)>1 else 0
                        if sp>10000:
                            sp=to_int(nums[5]) if len(nums)>5 else 0
                        parsed.append({"Periode":normalize_periode(label),"PIC":pic,"Kode_Dokter":kode,"Nama_Dokter":nama[:80],"CD_Omzet":co,"CD_Pasien":cp,"SA_Omzet":so,"SA_Pasien":sp,"Total_Omzet":co+so,"Total_Pasien":cp+sp})
                    i=j-1
                i+=1
    else:
        for txt in [doc[p].get_text("text") for p in range(len(doc))]:
            lines=txt.splitlines()
            pic="UNKNOWN"
            for i in range(len(lines)):
                if "Tanggal" in lines[i] or "s/d" in lines[i]:
                    for k in range(i+1,min(i+6,len(lines))):
                        c=lines[k].strip()
                        if len(c)>5 and c.upper()==c and "JASA" not in c and not any(ch.isdigit() for ch in c):
                            pic=c
                            break
                    break
            buf=""
            def flush(b,p):
                if not b:
                    return None
                t=b.split()
                k=""
                idx=-1
                for ti,x in enumerate(t):
                    if len(x)==10 and x.isdigit():
                        idx=ti
                        k=x
                        break
                if idx==-1:
                    return None
                aft=t[idx+1:]
                bn=-1
                for j,tt in enumerate(aft):
                    if tt in ["6","7","8","9","10","11","12","1","2","3","4","5"]:
                        bn=j
                        break
                if bn==-1:
                    return None
                nama=" ".join(aft[:bn])
                nums=[x for x in aft[bn+1:] if x.replace('.','').replace(',','').isdigit()]
                if len(nums)>=8:
                    return {"Periode":normalize_periode(label),"PIC":p.title(),"Kode_Dokter":k,"Nama_Dokter":nama[:80],"CD_Omzet":to_int(nums[0]),"CD_Pasien":to_int(nums[3]),"SA_Omzet":to_int(nums[4]),"SA_Pasien":to_int(nums[7]),"Total_Omzet":to_int(nums[0])+to_int(nums[4]),"Total_Pasien":to_int(nums[3])+to_int(nums[7])}
            for line in lines:
                line=line.strip()
                if not line or ("Kode" in line and "Dokter" in line):
                    continue
                if line.startswith("TOTAL") and len(line)<30:
                    r=flush(buf,pic)
                    if r:
                        parsed.append(r)
                    buf=""
                    continue
                if any(len(p)==10 and p.isdigit() for p in line.split()):
                    r=flush(buf,pic)
                    if r:
                        parsed.append(r)
                    buf=line
                else:
                    if buf:
                        buf+=" "+line
            r=flush(buf,pic)
            if r:
                parsed.append(r)
    return pd.DataFrame(parsed)

# ============ UI ============
st.title("📊 Monitoring Pramita V14 AUTO-PERMANEN")
has_token = "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets

with st.sidebar:
    st.header("📂 Upload Sekali Jadi Permanen")
    if has_token:
        st.success(f"✅ Auto-permanen AKTIF\n{st.secrets['GITHUB_REPO']}")
    else:
        st.error("⚠️ Belum setting Token - data masih sementara. Set di Streamlit Secrets.")

    multi = st.file_uploader("Pilih SEMUA PDF sekaligus", type=["pdf"], accept_multiple_files=True)
    if multi:
        all_new=[]
        for up in multi:
            fname=up.name.upper()
            import re
            th=re.findall(r'20\d{2}', fname)
            ths=th[0] if th else "2026"
            pg="OKTOBER-2026"
            for b in BULAN_FULL.keys():
                if b in fname:
                    pg=f"{BULAN_FULL[b]}-{ths}"
                    break
            open("temp.pdf","wb").write(up.getbuffer())
            df_t=parse_pdf("temp.pdf", pg)
            df_t["File"]=up.name
            all_new.append(df_t)
        if all_new:
            df_m=pd.concat(all_new, ignore_index=True)
            st.success(f"Terbaca {len(df_m)} dokter dari {len(multi)} file")
            if st.button("💾 SIMPAN PERMANEN", type="primary"):
                if os.path.exists(DB_FILE):
                    old=pd.read_excel(DB_FILE)
                    old["Periode"]=old["Periode"].apply(normalize_periode)
                    for per in df_m["Periode"].unique():
                        old=old[old["Periode"]!=per]
                    final=pd.concat([old, df_m.drop(columns=["File"])], ignore_index=True)
                else:
                    final=df_m.drop(columns=["File"])
                final["Total_Omzet"]=final["CD_Omzet"]+final["SA_Omzet"]
                final["Total_Pasien"]=final["CD_Pasien"]+final["SA_Pasien"]
                final.to_excel(DB_FILE, index=False)
                if has_token:
                    with st.spinner("Menyimpan permanen ke Github..."):
                        ok, msg = push_to_github(DB_FILE)
                        if ok:
                            st.success("✅ "+msg)
                            st.balloons()
                        else:
                            st.error(msg)
                            st.warning("Data tersimpan sementara, tapi gagal auto-backup. Cek Secrets.")
                else:
                    st.warning("Tersimpan sementara! Segera setting Token biar permanen.")
                st.rerun()

if not os.path.exists(DB_FILE):
    st.info("Database kosong. Upload PDF di sidebar kiri. Jika sudah setting Token, ini akan jadi permanen selamanya.")
    st.stop()

df=pd.read_excel(DB_FILE)
df["Periode"]=df["Periode"].apply(normalize_periode)
df["SortDate"]=df["Periode"].apply(parse_date)
df=df.sort_values("SortDate")
df["Kode_Dokter"]=df["Kode_Dokter"].astype(str)
df["Total_Omzet"]=df["CD_Omzet"]+df["SA_Omzet"]
df["Total_Pasien"]=df["CD_Pasien"]+df["SA_Pasien"]

with st.sidebar:
    st.divider()
    st.write(f"Database: {len(df)} baris")
    period_options=df.sort_values("SortDate")["Periode"].unique().tolist()
    sel_periode=st.multiselect("Periode", period_options, default=period_options)
    sel_pic=st.selectbox("PIC", ["Semua"]+sorted(df["PIC"].dropna().unique().tolist()))
    sort_by=st.selectbox("Urut", ["Total_Omzet","Total_Pasien"])
    if os.path.exists(DB_FILE) and not has_token:
        with open(DB_FILE,"rb") as f:
            st.download_button("⬇️ Download Manual (karena belum auto)", data=f.read(), file_name=DB_FILE)

df_f=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_pic!="Semua":
    df_f=df_f[df_f["PIC"]==sel_pic]

df_rank=df_f.groupby("Kode_Dokter",as_index=False).agg(Nama_Dokter=("Nama_Dokter","first"),PIC=("PIC","first"),CD_Omzet=("CD_Omzet","sum"),CD_Pasien=("CD_Pasien","sum"),SA_Omzet=("SA_Omzet","sum"),SA_Pasien=("SA_Pasien","sum"),Total_Omzet=("Total_Omzet","sum"),Total_Pasien=("Total_Pasien","sum"),Jumlah_Bulan=("Periode","nunique")).sort_values(by=sort_by,ascending=False) if len(sel_periode)>1 else df_f.drop_duplicates("Kode_Dokter").sort_values(by=sort_by,ascending=False)
df_rank=df_rank.reset_index(drop=True)
df_rank.insert(0,"No",range(1,len(df_rank)+1))

def hl(r):
    return ['font-weight: bold; background-color: #FFF176' if r.name<3 else 'font-weight: bold; background-color: #C8E6C9' if r.name<10 else '' for _ in r]

st.subheader(f"🏆 Ranking {sort_by} - {len(df_rank)} Dokter {'(AUTO-PERMANEN AKTIF)' if has_token else '(BELUM PERMANEN)'}")
st.dataframe(df_rank.style.apply(hl,axis=1),use_container_width=True,height=600,hide_index=True)
