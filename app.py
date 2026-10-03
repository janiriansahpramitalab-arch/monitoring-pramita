import fitz
import pandas as pd
import os
from datetime import datetime
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="Monitoring Dokter Pramita", layout="wide", page_icon="📊")
DB_FILE = "Database_Monitoring_Pramita.xlsx"

def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0

def parse_pdf(pdf_path, periode_label="Sep-2026"):
    doc = fitz.open(pdf_path)
    parsed = []
    for txt in [doc[p].get_text("text") for p in range(len(doc))]:
        lines = txt.splitlines()
        pic = "UNKNOWN"
        for i in range(len(lines)):
            if "Tanggal" in lines[i] or "s/d" in lines[i]:
                for k in range(i+1, min(i+6, len(lines))):
                    cand = lines[k].strip()
                    if len(cand)>5 and cand.upper()==cand and "JASA" not in cand:
                        if not any(ch.isdigit() for ch in cand):
                            pic = cand; break
                break
        buffer=""
        def flush(buf, pic):
            if not buf: return None
            toks = buf.split()
            kode_idx=-1
            for ti, t in enumerate(toks):
                if len(t)==10 and t.isdigit():
                    kode_idx=ti; kode=t; break
            if kode_idx==-1: return None
            after = toks[kode_idx+1:]
            bln=-1
            for j, tt in enumerate(after):
                if tt in ["9","10","11","12"]: bln=j; break
            if bln==-1: return None
            nama=" ".join(after[:bln])
            nums=[x for x in after[bln+1:] if x.replace('.','').replace(',','').isdigit()]
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

st.title("📊 Monitoring Dokter - Pramita Lab")
with st.sidebar:
    st.header("📂 Upload PDF")
    periode = st.text_input("Periode", value=datetime.now().strftime("%b-%Y"))
    up = st.file_uploader("Upload PDF", type=["pdf"])
    if up:
        open("temp.pdf","wb").write(up.getbuffer())
        df_new = parse_pdf("temp.pdf", periode)
        st.success(f"Terbaca {len(df_new)} dokter")
        st.dataframe(df_new.head())
        if st.button("💾 Simpan"):
            if os.path.exists(DB_FILE):
                old=pd.read_excel(DB_FILE)
                old=old[old["Periode"]!=periode]
                all_df=pd.concat([old,df_new])
            else: all_df=df_new
            all_df.to_excel(DB_FILE,index=False)
            st.success("Tersimpan!"); st.rerun()

if os.path.exists(DB_FILE):
    df=pd.read_excel(DB_FILE)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Omzet",f"Rp {df['Total_Omzet'].sum():,}")
    c2.metric("Pasien",f"{df['Total_Pasien'].sum():,}")
    c3.metric("Dokter",f"{df['Kode_Dokter'].nunique()}")
    c4.metric("PIC",f"{df['PIC'].nunique()}")
    st.subheader("🏆 Ranking - Omzet Tertinggi di Atas")
    st.dataframe(df.sort_values("Total_Omzet",ascending=False),use_container_width=True)
    top10=df.sort_values("Total_Omzet",ascending=False).head(10)
    fig,ax=plt.subplots(figsize=(8,5))
    ax.barh(top10["Nama_Dokter"].str[:25][::-1],top10["Total_Omzet"][::-1])
    st.pyplot(fig)
else:
    st.info("Upload PDF September dulu untuk mulai")
