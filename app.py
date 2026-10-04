import fitz, pandas as pd, os, io, json
import streamlit as st
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="Monitoring Pramita V13.5 ANTI KOSONG", layout="wide", page_icon="📊")
DB_FILE = "Database_Monitoring_Pramita.xlsx"
HIDDEN_FILE = "hidden_config.json"

PIC_HIDDEN_LIST = ["NARINDRA NATA KUNTHARA", "YOHANA DEWI RATIH"]
DOKTER_HIDDEN_KODE_DEFAULT = ["2741002000"]

def load_hidden():
    if os.path.exists(HIDDEN_FILE):
        try:
            with open(HIDDEN_FILE, 'r') as f:
                data = json.load(f)
                return data.get("PIC", PIC_HIDDEN_LIST), data.get("DOKTER_KODE", DOKTER_HIDDEN_KODE_DEFAULT)
        except: return PIC_HIDDEN_LIST, DOKTER_HIDDEN_KODE_DEFAULT
    return PIC_HIDDEN_LIST, DOKTER_HIDDEN_KODE_DEFAULT

def save_hidden(pic_list, kode_list):
    with open(HIDDEN_FILE, 'w') as f:
        json.dump({"PIC": pic_list, "DOKTER_KODE": kode_list}, f)

PIC_HIDDEN_LIST_LOAD, DOKTER_HIDDEN_KODE = load_hidden()

BULAN_FULL = {"JANU": "JANUARI", "JAN": "JANUARI","FEBR": "FEBRUARI", "FEB": "FEBRUARI","MARET": "MARET", "MAR": "MARET","APRIL": "APRIL", "APR": "APRIL","MEI": "MEI","JUNI": "JUNI", "JUN": "JUNI","JULI": "JULI", "JUL": "JULI","AGUS": "AGUSTUS", "AGU": "AGUSTUS","SEPT": "SEPTEMBER", "SEP": "SEPTEMBER","OKTO": "OKTOBER", "OKT": "OKTOBER","NOPE": "NOVEMBER", "NOV": "NOVEMBER","DESE": "DESEMBER", "DES": "DESEMBER"}
BULAN_ANGKA = {"JANUARI":1,"FEBRUARI":2,"MARET":3,"APRIL":4,"MEI":5,"JUNI":6,"JULI":7,"AGUSTUS":8,"SEPTEMBER":9,"OKTOBER":10,"NOVEMBER":11,"DESEMBER":12}

def normalize_periode(per_str):
    s = str(per_str).upper().strip(); tahun = "".join([c for c in s if c.isdigit()])[-4:]; huruf = "".join([c for c in s if c.isalpha()]); prefix = huruf[:4]; full = BULAN_FULL.get(prefix, huruf)
    if full in BULAN_ANGKA: return f"{full}-{tahun}" if tahun else full
    return f"{BULAN_FULL.get(prefix, prefix)}-{tahun}"
def parse_date(per_str):
    try: norm = normalize_periode(per_str); nama = norm.split("-")[0]; thn = int(norm.split("-")[1]); bln = BULAN_ANGKA.get(nama, 1); return pd.Timestamp(year=thn, month=bln, day=1)
    except: return pd.Timestamp(year=2026, month=1, day=1)
def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0

def parse_pdf(pdf_path, periode_label):
    doc = fitz.open(pdf_path); parsed = []
    is_new = any("CIK DI TIRO" in doc[p].get_text("text") for p in range(min(2,len(doc))))
    if is_new:
        for p_idx in range(len(doc)):
            txt=doc[p_idx].get_text("text"); lines=[l.strip() for l in txt.splitlines() if l.strip()!='']; pic="UNKNOWN"
            for i,l in enumerate(lines):
                if "Tanggal" in l or "s/d" in l:
                    for k in range(i+1, min(i+6,len(lines))):
                        cand=lines[k].strip()
                        if len(cand)>3 and cand.upper()==cand and "JASA" not in cand and "PENGAMBILAN" not in cand and "PRE" not in cand:
                            if not any(ch.isdigit() for ch in cand) and len(cand.split())<=4: pic=cand.title(); break
                    break
            i=0
            while i < len(lines):
                l=lines[i]; cd=''.join(c for c in l if c.isdigit())
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
                        if tc.isdigit(): nums.append(tc);
                        j+=1
                    if len(nums)>=4:
                        cik_o = to_int(nums[0]); cik_p = to_int(nums[3]) if len(nums)>3 else 0
                        sul_o = to_int(nums[4]) if len(nums)>4 else 0; sul_p = to_int(nums[7]) if len(nums)>7 else 0
                        if cik_p > 10000: cik_p = to_int(nums[1]) if len(nums)>1 else 0
                        if sul_p > 10000: sul_p = to_int(nums[5]) if len(nums)>5 else 0
                        tot_o = cik_o + sul_o; tot_p = cik_p + sul_p
                        parsed.append({"Periode":normalize_periode(periode_label),"PIC":pic,"Kode_Dokter":kode,"Nama_Dokter":nama[:80],"CD_Omzet":cik_o,"CD_Pasien":cik_p,"SA_Omzet":sul_o,"SA_Pasien":sul_p,"Total_Omzet":tot_o,"Total_Pasien":tot_p})
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
                if len(nums)>=8:
                    cik_o=to_int(nums[0]); cik_p=to_int(nums[3]); sul_o=to_int(nums[4]); sul_p=to_int(nums[7])
                    tot_o=cik_o+sul_o; tot_p=cik_p+sul_p
                    return {"Periode":normalize_periode(periode_label),"PIC":pic.title(),"Kode_Dokter":kode,"Nama_Dokter":nama[:80],"CD_Omzet":cik_o,"CD_Pasien":cik_p,"SA_Omzet":sul_o,"SA_Pasien":sul_p,"Total_Omzet":tot_o,"Total_Pasien":tot_p}
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

st.title("📊 Monitoring Pramita V13.5 ANTI KOSONG")

with st.sidebar:
    st.header("📂 Upload PDF")
    periode=st.text_input("Periode", value="OKTOBER-2026")
    up=st.file_uploader("Upload PDF", type=["pdf"])
    if up:
        open("temp.pdf","wb").write(up.getbuffer())
        df_new=parse_pdf("temp.pdf", periode)
        st.success(f"Terbaca {len(df_new)} -> {normalize_periode(periode)}")
        if st.button("💾 Simpan PDF ke Database"):
            if os.path.exists(DB_FILE):
                old=pd.read_excel(DB_FILE); old["Periode"] = old["Periode"].apply(normalize_periode)
                old=old[old["Periode"]!=normalize_periode(periode)]
                all_df=pd.concat([old,df_new], ignore_index=True)
            else: all_df=df_new
            all_df["Total_Omzet"] = all_df["CD_Omzet"] + all_df["SA_Omzet"]
            all_df["Total_Pasien"] = all_df["CD_Pasien"] + all_df["SA_Pasien"]
            all_df.to_excel(DB_FILE,index=False); st.success("Tersimpan!"); st.rerun()
    st.divider()
    st.header("♻️ RESTORE DATABASE")
    st.caption("Jika data kosong, upload file Excel backup disini")
    up_excel = st.file_uploader("Upload Database_Monitoring_Pramita.xlsx", type=["xlsx"])
    if up_excel:
        open(DB_FILE,"wb").write(up_excel.getbuffer())
        st.success("Database berhasil di-restore!"); st.rerun()

if not os.path.exists(DB_FILE):
    st.warning("⚠️ Database belum ada / hilang karena restart server.")
    st.info("Silahkan upload file `Database_Monitoring_Pramita.xlsx` di sidebar kiri > RESTORE DATABASE. Jika tidak punya, upload ulang semua PDF bulan Jan-Okt.")
    st.stop()

df=pd.read_excel(DB_FILE)
df["Periode"] = df["Periode"].apply(normalize_periode)
df["SortDate"] = df["Periode"].apply(parse_date)
df = df.sort_values("SortDate")
df["Kode_Dokter"] = df["Kode_Dokter"].astype(str)
df["Total_Omzet"] = df["CD_Omzet"] + df["SA_Omzet"]
df["Total_Pasien"] = df["CD_Pasien"] + df["SA_Pasien"]

latest_periode = df.sort_values("SortDate")["Periode"].iloc[-1]
latest_pics_upper = df[df["Periode"]==latest_periode]["PIC"].astype(str).str.upper().unique().tolist()
hidden_active = [t for t in PIC_HIDDEN_LIST_LOAD if any(t in lp or lp in t for lp in latest_pics_upper)]
pics_to_hide_real = [p for p in df["PIC"].unique() if any(h in str(p).upper() for h in hidden_active)]

with st.sidebar:
    st.divider()
    st.header("🔎 Filter Ranking")
    period_options = df.sort_values("SortDate")["Periode"].unique().tolist()
    sel_periode=st.multiselect("Pilih Periode", period_options, default=period_options)
    sel_pic=st.selectbox("PIC", ["Semua"]+sorted(df["PIC"].dropna().unique().tolist()))
    sort_by=st.selectbox("Urut Ranking", ["Total_Omzet","Total_Pasien","CD_Omzet","SA_Omzet"])
    if st.button("🔧 Perbaiki Total Pasien"):
        df["Total_Omzet"] = df["CD_Omzet"] + df["SA_Omzet"]
        df["Total_Pasien"] = df["CD_Pasien"] + df["SA_Pasien"]
        df.drop(columns=["SortDate"]).to_excel(DB_FILE,index=False)
        st.success("Diperbaiki!"); st.rerun()
    st.divider()
    if os.path.exists(DB_FILE):
        with open(DB_FILE,"rb") as f:
            st.download_button("💾 Download Backup Database", data=f.read(), file_name="Database_Monitoring_Pramita.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

df_f=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_pic!="Semua": df_f=df_f[df_f["PIC"]==sel_pic]
df_rank_base = df_f.copy()
if pics_to_hide_real: df_rank_base = df_rank_base[~df_rank_base["PIC"].isin(pics_to_hide_real)]
df_rank_base = df_rank_base[~df_rank_base["Kode_Dokter"].astype(str).isin(DOKTER_HIDDEN_KODE)]

if len(sel_periode) > 1:
    df_rank = df_rank_base.groupby("Kode_Dokter", as_index=False).agg(
        Nama_Dokter=("Nama_Dokter","first"), PIC=("PIC","first"),
        CD_Omzet=("CD_Omzet","sum"), CD_Pasien=("CD_Pasien","sum"),
        SA_Omzet=("SA_Omzet","sum"), SA_Pasien=("SA_Pasien","sum"),
        Total_Omzet=("Total_Omzet","sum"), Total_Pasien=("Total_Pasien","sum"),
        Jumlah_Bulan=("Periode","nunique")
    ).sort_values(by=sort_by, ascending=False)
else:
    df_rank_base = df_rank_base.drop_duplicates(subset=["Kode_Dokter"])
    df_rank = df_rank_base.sort_values(by=sort_by, ascending=False)

df_rank = df_rank.reset_index(drop=True)
df_rank.insert(0, "No", range(1, len(df_rank)+1))
def highlight_top10(row):
    idx = row.name
    if idx < 3: return ['font-weight: bold; background-color: #FFF176; color: black'] * len(row)
    elif idx < 10: return ['font-weight: bold; background-color: #C8E6C9; color: black'] * len(row)
    else: return [''] * len(row)

st.subheader(f"🏆 Ranking {sort_by} - {len(df_rank)} Dokter")
st.dataframe(df_rank.style.apply(highlight_top10, axis=1), use_container_width=True, height=650, hide_index=True)
buf=io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    df_rank.to_excel(writer, index=False, sheet_name='Ranking_FIX')
    ws = writer.sheets['Ranking_FIX']
    fill_gold = PatternFill(start_color="FFF176", end_color="FFF176", fill_type="solid")
    fill_green = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    font_bold = Font(bold=True)
    for r_idx in range(2, min(12, len(df_rank)+2)):
        for c in range(1, len(df_rank.columns)+1):
            cell = ws.cell(row=r_idx, column=c)
            cell.font = font_bold
            cell.fill = fill_gold if r_idx <= 4 else fill_green
st.download_button("📥 Export Excel FIX", data=buf.getvalue(), file_name=f"Ranking_FIX_{'_'.join(sel_periode)}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
