import fitz, pandas as pd, os, io
import streamlit as st

st.set_page_config(page_title="Monitoring Pramita V6", layout="wide", page_icon="📊")
DB_FILE = "Database_Monitoring_Pramita.xlsx"

# Mapping biar JANU,FEBR,Agus,Maret dll kebaca semua
BULAN_MAP = {
    "JANU":1, "JAN":1,
    "FEBR":2, "FEB":2,
    "MARET":3, "MARET-":3, "MAR":3,
    "APRIL":4, "APR":4,
    "MEI":5, "MEI-":5,
    "JUNI":6, "JUN":6,
    "JULI":7, "JUL":7,
    "AGUS":8, "AGU":8, "AUG":8,
    "SEPT":9, "SEP":9,
    "OKTO":10, "OKT":10, "OCT":10,
    "NOPE":11, "NOV":11,
    "DESE":12, "DES":12, "DEC":12
}

def parse_periode_to_date(per_str):
    try:
        per = str(per_str).upper()
        # ambil 4 huruf awal bulan
        prefix = "".join([c for c in per if c.isalpha()])[:4]
        bulan = BULAN_MAP.get(prefix, 1)
        tahun = int("".join([c for c in per if c.isdigit()][-4:]))
        return pd.Timestamp(year=tahun, month=bulan, day=1)
    except:
        return pd.Timestamp(year=2026, month=1, day=1)

def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0

def parse_pdf(pdf_path, periode_label):
    #... (pakai fungsi parse_pdf V5 yang kemarin sudah bisa baca Juni)...
    doc = fitz.open(pdf_path)
    parsed = []
    is_new_format=False
    for p in range(min(2,len(doc))):
        if "CIK DI TIRO" in doc[p].get_text("text"): is_new_format=True; break
    if is_new_format:
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
                        parsed.append({"Periode":periode_label,"PIC":pic,"Kode_Dokter":kode,"Nama_Dokter":nama[:60],"CD_Omzet":cik_total,"CD_Pasien":cik_psn,"SA_Omzet":sultan_total,"SA_Pasien":sultan_psn,"Total_Omzet":total_omzet,"Total_Pasien":total_psn})
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
                    return {"Periode":periode_label,"PIC":pic.title(),"Kode_Dokter":kode,"Nama_Dokter":nama,"CD_Omzet":to_int(nums[0]),"CD_Pasien":to_int(nums[3]),"SA_Omzet":to_int(nums[4]),"SA_Pasien":to_int(nums[7]),"Total_Omzet":to_int(nums[8]),"Total_Pasien":to_int(nums[10])}
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

st.title("📊 Monitoring Pramita V6 - Urut Bulan Otomatis")

with st.sidebar:
    st.header("📂 Upload")
    periode=st.text_input("Periode (JANU-2026, FEBR-2026, dst)", value="OKTO-2026")
    up=st.file_uploader("Upload PDF", type=["pdf"])
    if up:
        open("temp.pdf","wb").write(up.getbuffer())
        df_new=parse_pdf("temp.pdf", periode.upper())
        st.success(f"Terbaca {len(df_new)} dokter")
        if len(df_new)>0:
            st.dataframe(df_new.head(10))
            if st.button("💾 Simpan"):
                if os.path.exists(DB_FILE):
                    old=pd.read_excel(DB_FILE); old=old[old["Periode"]!=periode.upper()]
                    all_df=pd.concat([old,df_new], ignore_index=True)
                else: all_df=df_new
                all_df.to_excel(DB_FILE,index=False); st.success("Tersimpan!"); st.rerun()
    st.divider()
    st.header("🗑️ Hapus Data")
    if os.path.exists(DB_FILE):
        df_tmp=pd.read_excel(DB_FILE)
        if not df_tmp.empty:
            # Urutkan pilihan hapus juga Jan-Des
            df_tmp["SortDate"] = df_tmp["Periode"].apply(parse_periode_to_date)
            sorted_periods = df_tmp.sort_values("SortDate")["Periode"].unique()
            del_per=st.selectbox("Hapus periode", sorted_periods)
            if st.button(f"Hapus {del_per}"):
                df_del=df_tmp[df_tmp["Periode"]!=del_per]; df_del.to_excel(DB_FILE,index=False); st.rerun()
            if st.button("⚠️ HAPUS SEMUA"):
                os.remove(DB_FILE); st.rerun()

if not os.path.exists(DB_FILE): st.info("Upload dulu"); st.stop()
df=pd.read_excel(DB_FILE)

# Buat kolom tanggal untuk sorting
df["SortDate"] = df["Periode"].apply(parse_periode_to_date)
df = df.sort_values("SortDate")

with st.sidebar:
    st.divider()
    st.header("🔎 Filter Ranking")
    # Urutkan pilihan filter Jan-Des juga
    period_options = df.sort_values("SortDate")["Periode"].unique().tolist()
    sel_periode=st.multiselect("Pilih Periode", period_options, default=period_options)
    sel_pic=st.selectbox("PIC", ["Semua"]+sorted(df["PIC"].dropna().unique().tolist()))
    sort_by=st.selectbox("Urut Ranking", ["Total_Omzet","Total_Pasien","CD_Omzet","SA_Omzet"])

df_f=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_pic!="Semua": df_f=df_f[df_f["PIC"]==sel_pic]

# Ranking Akumulasi Anti Double
if len(sel_periode) > 1:
    df_rank = df_f.groupby(["Kode_Dokter","Nama_Dokter"], as_index=False).agg(
        PIC=("PIC","first"), CD_Omzet=("CD_Omzet","sum"), CD_Pasien=("CD_Pasien","sum"),
        SA_Omzet=("SA_Omzet","sum"), SA_Pasien=("SA_Pasien","sum"),
        Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum"),
        Jumlah_Bulan=("Periode","nunique")
    ).sort_values(by=sort_by, ascending=False)
    total_dokter_unik = df_f["Kode_Dokter"].nunique()
else:
    df_rank = df_f.sort_values(by=sort_by, ascending=False)
    total_dokter_unik = df_f["Kode_Dokter"].nunique()

c1,c2,c3,c4=st.columns(4)
c1.metric("Total Omzet", f"Rp {df_f['Total_Omzet'].sum():,}")
c2.metric("Total Pasien", f"{df_f['Total_Pasien'].sum():,}")
c3.metric("Dokter Unik", f"{total_dokter_unik}")
c4.metric("PIC", f"{df_f['PIC'].nunique()}")

st.subheader(f"🏆 Ranking {sort_by} - {len(df_rank)} Dokter Unik")

buf=io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as w:
    df_rank.to_excel(w, index=False, sheet_name='Ranking_Akumulasi')
st.download_button("📥 Export Excel", data=buf.getvalue(), file_name=f"Ranking_Akumulasi_{'_'.join(sel_periode)}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.dataframe(df_rank, use_container_width=True, height=450)

st.divider()
st.subheader("📅 Trend MoM - Sudah Urut Jan-Des")

trend = df.groupby("Periode", as_index=False).agg(
    Total_Omzet=("Total_Omzet","sum"),
    Total_Pasien=("Total_Pasien","sum"),
    Kode_Dokter=("Kode_Dokter","nunique")
)
trend["SortDate"] = trend["Periode"].apply(parse_periode_to_date)
trend = trend.sort_values("SortDate")

# Grafik urut
st.line_chart(trend.set_index("Periode")[["Total_Omzet"]], use_container_width=True)

# Tabel juga urut
st.dataframe(trend[["Periode","Total_Omzet","Total_Pasien","Kode_Dokter"]], use_container_width=True)
