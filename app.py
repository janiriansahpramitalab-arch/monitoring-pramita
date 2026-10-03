import fitz, pandas as pd, os, io
import streamlit as st

st.set_page_config(page_title="Monitoring Pramita V10 FINAL", layout="wide", page_icon="📊")
DB_FILE = "Database_Monitoring_Pramita.xlsx"

PIC_HIDDEN_LIST = ["NARINDRA NATA KUNTHARA", "YOHANA DEWI RATIH"]

BULAN_FULL = {
    "JANU": "JANUARI", "JAN": "JANUARI",
    "FEBR": "FEBRUARI", "FEB": "FEBRUARI",
    "MARET": "MARET", "MAR": "MARET",
    "APRIL": "APRIL", "APR": "APRIL",
    "MEI": "MEI",
    "JUNI": "JUNI", "JUN": "JUNI",
    "JULI": "JULI", "JUL": "JULI",
    "AGUS": "AGUSTUS", "AGU": "AGUSTUS",
    "SEPT": "SEPTEMBER", "SEP": "SEPTEMBER",
    "OKTO": "OKTOBER", "OKT": "OKTOBER",
    "NOPE": "NOVEMBER", "NOV": "NOVEMBER",
    "DESE": "DESEMBER", "DES": "DESEMBER"
}
BULAN_ANGKA = {"JANUARI":1,"FEBRUARI":2,"MARET":3,"APRIL":4,"MEI":5,"JUNI":6,"JULI":7,"AGUSTUS":8,"SEPTEMBER":9,"OKTOBER":10,"NOVEMBER":11,"DESEMBER":12}

def normalize_periode(per_str):
    s = str(per_str).upper().strip()
    tahun = "".join([c for c in s if c.isdigit()])[-4:]
    huruf = "".join([c for c in s if c.isalpha()])
    prefix = huruf[:4]
    full = BULAN_FULL.get(prefix, huruf)
    if full in BULAN_ANGKA:
        return f"{full}-{tahun}" if tahun else full
    return f"{BULAN_FULL.get(prefix, prefix)}-{tahun}"

def parse_date(per_str):
    try:
        norm = normalize_periode(per_str)
        nama = norm.split("-")[0]
        thn = int(norm.split("-")[1])
        bln = BULAN_ANGKA.get(nama, 1)
        return pd.Timestamp(year=thn, month=bln, day=1)
    except:
        return pd.Timestamp(year=2026, month=1, day=1)

def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0

def parse_pdf(pdf_path, periode_label):
    doc = fitz.open(pdf_path)
    parsed = []
    is_new = any("CIK DI TIRO" in doc[p].get_text("text") for p in range(min(2,len(doc))))
    if is_new:
        for p_idx in range(len(doc)):
            txt=doc[p_idx].get_text("text")
            lines=[l.strip() for l in txt.splitlines() if l.strip()!='']
            pic="UNKNOWN"
            for i,l in enumerate(lines):
                if "Tanggal" in l or "s/d" in l:
                    for k in range(i+1, min(i+6,len(lines))):
                        cand=lines[k].strip()
                        if len(cand)>3 and cand.upper()==cand and "JASA" not in cand and "PENGAMBILAN" not in cand and "PRE" not in cand:
                            if not any(ch.isdigit() for ch in cand) and len(cand.split())<=4:
                                pic=cand.title(); break
                    break
            i=0
            while i < len(lines):
                l=lines[i]
                cd=''.join(c for c in l if c.isdigit())
                if len(cd)==10 and l.replace('.','').isdigit():
                    kode=cd; nama=lines[i+1] if i+1<len(lines) and any(c.isalpha() for c in lines[i+1]) else ""; offset=2 if nama else 1
                    if i+offset < len(lines) and lines[i+offset] in ["6","7","8","9","10","11","12","5","4","3","2","1"]: offset+=1
                    nums=[]; j=i+offset
                    while j < len(lines) and len(nums)<12:
                        cur=lines[j]; c2=''.join(c for c in cur if c.isdigit())
                        if len(c2)==10 and cur.replace('.','').isdigit(): break
                        if cur.upper().startswith("TOTAL"): break
                        if cur.startswith("Yang Membuat"): break
                        tc=cur.replace('.','').replace(',','')
                        if tc.isdigit(): nums.append(tc)
                        j+=1
                    if len(nums)>=4:
                        if len(nums)>=12:
                            cik_total=to_int(nums[0]); cik_psn=to_int(nums[3]); sultan_total=to_int(nums[4]); sultan_psn=to_int(nums[7]); total_omzet=to_int(nums[8]); total_psn=to_int(nums[10])
                        else:
                            cik_total=to_int(nums[0]); cik_psn=to_int(nums[3]) if len(nums)>3 else 0; sultan_total=to_int(nums[4]) if len(nums)>4 else 0; sultan_psn=to_int(nums[7]) if len(nums)>7 else 0; total_omzet=cik_total+sultan_total; total_psn=cik_psn+sultan_psn
                            if len(nums)>=12: total_omzet=to_int(nums[8]); total_psn=to_int(nums[10])
                        parsed.append({"Periode":normalize_periode(periode_label),"PIC":pic,"Kode_Dokter":kode,"Nama_Dokter":nama[:60],"CD_Omzet":cik_total,"CD_Pasien":cik_psn,"SA_Omzet":sultan_total,"SA_Pasien":sultan_psn,"Total_Omzet":total_omzet,"Total_Pasien":total_psn})
                    i=j-1
                i+=1
    else:
        for txt in [doc[p].get_text("text") for p in range(len(doc))]:
            lines=txt.splitlines(); pic="UNKNOWN"
            for i in range(len(lines)):
                if "Tanggal" in lines[i] or "s/d" in lines[i]:
                    for k in range(i+1, min(i+6, len(lines))):
                        cand=lines[k].strip()
                        if len(cand)>5 and cand.upper()==cand and "JASA" not in cand:
                            if not any(ch.isdigit() for ch in cand): pic=cand; break
                    break
            buffer=""
            def flush(buf,pic):
                if not buf: return None
                toks=buf.split(); kode_idx=-1; kode=""
                for ti,t in enumerate(toks):
                    if len(t)==10 and t.isdigit(): kode_idx=ti; kode=t; break
                if kode_idx==-1: return None
                after=toks[kode_idx+1:]; bln=-1
                for j,tt in enumerate(after):
                    if tt in ["6","7","8","9","10","11","12","1","2","3","4","5"]: bln=j; break
                if bln==-1: return None
                nama=" ".join(after[:bln]); nums=[x for x in after[bln+1:] if x.replace('.','').replace(',','').isdigit()]
                if len(nums)>=12:
                    return {"Periode":normalize_periode(periode_label),"PIC":pic.title(),"Kode_Dokter":kode,"Nama_Dokter":nama,"CD_Omzet":to_int(nums[0]),"CD_Pasien":to_int(nums[3]),"SA_Omzet":to_int(nums[4]),"SA_Pasien":to_int(nums[7]),"Total_Omzet":to_int(nums[8]),"Total_Pasien":to_int(nums[10])}
                return None
            for line in lines:
                line=line.strip()
                if not line: continue
                if "Kode" in line and "Dokter" in line: continue
                if line.startswith("TOTAL") and len(line)<30:
                    r=flush(buffer,pic)
                    if r: parsed.append(r)
                    buffer=""; continue
                has_code=any(len(p)==10 and p.isdigit() for p in line.split())
                if has_code:
                    r=flush(buffer,pic)
                    if r: parsed.append(r)
                    buffer=line
                else:
                    if buffer: buffer+=" "+line
            r=flush(buffer,pic)
            if r: parsed.append(r)
    return pd.DataFrame(parsed)

st.title("📊 Monitoring Pramita V10 FINAL - No Urut 1,2,3")

with st.sidebar:
    st.header("📂 Upload PDF")
    periode=st.text_input("Periode (cth: SEPTEMBER-2026)", value="OKTOBER-2026")
    up=st.file_uploader("Upload PDF", type=["pdf"])
    if up:
        open("temp.pdf","wb").write(up.getbuffer())
        df_new=parse_pdf("temp.pdf", periode)
        st.success(f"Terbaca {len(df_new)} dokter -> {normalize_periode(periode)}")
        st.dataframe(df_new.head(3))
        if st.button("💾 Simpan"):
            if os.path.exists(DB_FILE):
                old=pd.read_excel(DB_FILE)
                old["Periode"] = old["Periode"].apply(normalize_periode)
                old=old[old["Periode"]!=normalize_periode(periode)]
                all_df=pd.concat([old,df_new], ignore_index=True)
            else: all_df=df_new
            all_df.to_excel(DB_FILE,index=False); st.success("Tersimpan!"); st.rerun()
    st.divider()
    if os.path.exists(DB_FILE):
        df_tmp=pd.read_excel(DB_FILE)
        if not df_tmp.empty:
            df_tmp["Periode"] = df_tmp["Periode"].apply(normalize_periode)
            df_tmp["SortDate"] = df_tmp["Periode"].apply(parse_date)
            sorted_p = df_tmp.sort_values("SortDate")["Periode"].unique()
            del_per=st.selectbox("Hapus periode", sorted_p)
            if st.button(f"Hapus {del_per}"):
                df_del=df_tmp[df_tmp["Periode"]!=del_per]; df_del.drop(columns=["SortDate"]).to_excel(DB_FILE,index=False); st.rerun()
            if st.button("⚠️ HAPUS SEMUA"):
                os.remove(DB_FILE); st.rerun()

if not os.path.exists(DB_FILE): st.info("Upload dulu"); st.stop()
df=pd.read_excel(DB_FILE)
df["Periode"] = df["Periode"].apply(normalize_periode)
df["SortDate"] = df["Periode"].apply(parse_date)
df = df.sort_values("SortDate")
df.drop(columns=["SortDate"]).to_excel(DB_FILE, index=False)
df["SortDate"] = df["Periode"].apply(parse_date)
df = df.sort_values("SortDate")

# LOGIKA HIDE BERDASARKAN FILE TERAKHIR
latest_periode = df.sort_values("SortDate")["Periode"].iloc[-1]
latest_pics_upper = df[df["Periode"]==latest_periode]["PIC"].astype(str).str.upper().unique().tolist()

hidden_active = [t for t in PIC_HIDDEN_LIST if any(t in lp or lp in t for lp in latest_pics_upper)]
pics_to_hide_real = [p for p in df["PIC"].unique() if any(h in str(p).upper() for h in hidden_active)]

with st.sidebar:
    st.divider()
    st.header("🔎 Filter Ranking")
    period_options = df.sort_values("SortDate")["Periode"].unique().tolist()
    sel_periode=st.multiselect("Pilih Periode", period_options, default=period_options)
    sel_pic=st.selectbox("PIC", ["Semua"]+sorted(df["PIC"].dropna().unique().tolist()))
    sort_by=st.selectbox("Urut Ranking", ["Total_Omzet","Total_Pasien","CD_Omzet","SA_Omzet"])
    st.divider()
    st.caption(f"File terakhir: {latest_periode}")
    if pics_to_hide_real:
        st.warning(f"🙈 Hide Aktif: {', '.join(pics_to_hide_real)}")
    else:
        st.success("✅ Tidak ada PIC disembunyikan")

df_f=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_pic!="Semua": df_f=df_f[df_f["PIC"]==sel_pic]

c1,c2,c3=st.columns(3)
c1.metric("Total Omzet (Semua PIC)", f"Rp {df_f['Total_Omzet'].sum():,}")
c2.metric("Total Pasien (Semua PIC)", f"{df_f['Total_Pasien'].sum():,}")
c3.metric("Dokter Unik", f"{df_f['Kode_Dokter'].nunique()}")

# RANKING TANPA PIC HIDDEN
df_rank_base = df_f[~df_f["PIC"].isin(pics_to_hide_real)] if pics_to_hide_real else df_f

if len(sel_periode) > 1:
    df_rank = df_rank_base.groupby(["Kode_Dokter","Nama_Dokter"], as_index=False).agg(
        PIC=("PIC","first"),
        CD_Omzet=("CD_Omzet","sum"), CD_Pasien=("CD_Pasien","sum"),
        SA_Omzet=("SA_Omzet","sum"), SA_Pasien=("SA_Pasien","sum"),
        Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum"),
        Jumlah_Bulan=("Periode","nunique")
    ).sort_values(by=sort_by, ascending=False)
else:
    df_rank = df_rank_base.sort_values(by=sort_by, ascending=False)

# --- FIX UTAMA NO URUT 1,2,3 ---
df_rank = df_rank.reset_index(drop=True)
df_rank.insert(0, "No", range(1, len(df_rank)+1))
if "SortDate" in df_rank.columns:
    df_rank = df_rank.drop(columns=["SortDate"])

st.subheader(f"🏆 Ranking {sort_by} - {len(df_rank)} Dokter")
st.dataframe(df_rank, use_container_width=True, height=550, hide_index=True)

buf=io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as w:
    df_rank.to_excel(w, index=False, sheet_name='Ranking_No_Urut')
st.download_button("📥 Export Excel Ranking (No 1,2,3)", data=buf.getvalue(), file_name=f"Ranking_{'_'.join(sel_periode)}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.subheader("📅 Trend MoM Urut Januari-Desember")
trend = df.groupby("Periode", as_index=False).agg(Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum"), Kode_Dokter=("Kode_Dokter","nunique"))
trend["SortDate"] = trend["Periode"].apply(parse_date)
trend = trend.sort_values("SortDate")
st.line_chart(trend.set_index("Periode")[["Total_Omzet"]])
st.dataframe(trend[["Periode","Total_Omzet","Total_Pasien","Kode_Dokter"]].reset_index(drop=True), use_container_width=True, hide_index=True)
