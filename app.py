import fitz
import pandas as pd
import os
from datetime import datetime
import streamlit as st
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Monitoring Dokter Pramita PRO", layout="wide", page_icon="📊")
DB_FILE = "Database_Monitoring_Pramita.xlsx"

def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0

def parse_pdf(pdf_path, periode_label):
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
            kode_idx=-1; kode=""
            for ti, t in enumerate(toks):
                if len(t)==10 and t.isdigit():
                    kode_idx=ti; kode=t; break
            if kode_idx==-1: return None
            after = toks[kode_idx+1:]
            bln=-1
            for j, tt in enumerate(after):
                if tt in ["7","8","9","10","11","12"]: bln=j; break
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

st.title("📊 Monitoring Dokter - Pramita Lab PRO")

# Sidebar Upload
with st.sidebar:
    st.header("📂 Upload Data Bulanan")
    periode = st.text_input("Periode (Contoh: Sep-2026)", value=datetime.now().strftime("%b-%Y"))
    up = st.file_uploader("Upload PDF Pramita", type=["pdf"])
    if up:
        open("temp.pdf","wb").write(up.getbuffer())
        df_new = parse_pdf("temp.pdf", periode)
        st.success(f"Terbaca {len(df_new)} dokter")
        st.dataframe(df_new.head(10))
        if st.button("💾 Simpan ke Database"):
            if os.path.exists(DB_FILE):
                old=pd.read_excel(DB_FILE)
                old=old[old["Periode"]!=periode]
                all_df=pd.concat([old,df_new], ignore_index=True)
            else: all_df=df_new
            all_df.to_excel(DB_FILE,index=False)
            st.success(f"Data {periode} tersimpan!"); st.rerun()

# Main
if not os.path.exists(DB_FILE):
    st.info("Database belum ada. Upload PDF September dulu.")
    st.stop()

df = pd.read_excel(DB_FILE)

# Filter
with st.sidebar:
    st.divider()
    st.header("🔎 Filter")
    sel_periode = st.multiselect("Pilih Periode", sorted(df["Periode"].unique()), default=sorted(df["Periode"].unique()))
    sel_pic = st.selectbox("Pilih PIC", ["Semua"]+sorted(df["PIC"].dropna().unique().tolist()))
    sort_by = st.selectbox("Urut Ranking", ["Total_Omzet","Total_Pasien","CD_Omzet","SA_Omzet"])

df_f = df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_pic!="Semua": df_f = df_f[df_f["PIC"]==sel_pic]

# KPI
c1,c2,c3,c4 = st.columns(4)
c1.metric("Total Omzet", f"Rp {df_f['Total_Omzet'].sum():,}")
c2.metric("Total Pasien", f"{df_f['Total_Pasien'].sum():,}")
c3.metric("Jml Dokter", f"{df_f['Kode_Dokter'].nunique()}")
c4.metric("Jml PIC", f"{df_f['PIC'].nunique()}")

# Ranking
st.subheader(f"🏆 Ranking Dokter - {sort_by} Tertinggi di Atas")
df_rank = df_f.sort_values(by=sort_by, ascending=False)

# Export Excel Button - FITUR BARU
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    df_rank.to_excel(writer, index=False, sheet_name='Ranking')
    # Sheet rekap per PIC
    df_f.groupby("PIC")[["Total_Omzet","Total_Pasien"]].sum().sort_values("Total_Omzet",ascending=False).to_excel(writer, sheet_name='Rekap_PIC')
writer_bytes = buffer.getvalue()

st.download_button(
    label="📥 Export Ranking ke Excel (Laporan Pimpinan)",
    data=writer_bytes,
    file_name=f"Ranking_Dokter_{'_'.join(sel_periode)}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.dataframe(df_rank, use_container_width=True, height=450)

# Charts
col1, col2 = st.columns(2)
with col1:
    st.subheader("📈 Top 10 Omzet")
    top10 = df_rank.head(10)
    fig, ax = plt.subplots(figsize=(8,5))
    ax.barh(top10["Nama_Dokter"].str[:20][::-1], top10["Total_Omzet"][::-1], color="#1f77b4")
    plt.tight_layout(); st.pyplot(fig)
with col2:
    st.subheader("👥 Top 10 Pasien")
    top10p = df_f.sort_values("Total_Pasien",ascending=False).head(10)
    fig2, ax2 = plt.subplots(figsize=(8,5))
    ax2.barh(top10p["Nama_Dokter"].str[:20][::-1], top10p["Total_Pasien"][::-1], color="#2ca02c")
    plt.tight_layout(); st.pyplot(fig2)

# FITUR BARU: Trend Bulan ke Bulan
st.divider()
st.subheader("📅 Trend Bulan ke Bulan (MoM) - PRO")
if len(df["Periode"].unique())>=1:
    # Coba urutkan periode dengan benar
    trend = df.groupby("Periode")[["Total_Omzet","Total_Pasien","Kode_Dokter"]].agg({"Total_Omzet":"sum","Total_Pasien":"sum","Kode_Dokter":"nunique"}).reset_index()
    # Urutkan berdasarkan tanggal
    try:
        trend["SortDate"] = pd.to_datetime(trend["Periode"], format="%b-%Y")
        trend = trend.sort_values("SortDate")
    except:
        trend = trend.sort_values("Periode")

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.write("**Trend Omzet**")
        st.line_chart(trend.set_index("Periode")[["Total_Omzet"]])
    with col_t2:
        st.write("**Trend Pasien & Jumlah Dokter Aktif**")
        st.line_chart(trend.set_index("Periode")[["Total_Pasien","Kode_Dokter"]])

    st.dataframe(trend.drop(columns=["SortDate"], errors='ignore'), use_container_width=True)
    st.caption("Upload bulan Okt, Nov, Des nanti, grafik ini akan otomatis membentuk garis trend untuk melihat pertumbuhan.")
else:
    st.info("Baru 1 periode, upload bulan depan untuk melihat trend.")

# Detail Dokter
st.subheader("🔍 Detail Per Dokter (MoM)")
search = st.text_input("Ketik nama / kode dokter untuk lihat trend nya")
if search:
    det = df[df["Nama_Dokter"].str.contains(search, case=False, na=False) | df["Kode_Dokter"].astype(str).str.contains(search)]
    if not det.empty:
        st.dataframe(det.sort_values("Periode"))
        st.line_chart(det.groupby("Periode")[["Total_Omzet","Total_Pasien"]].sum())
    else:
        st.warning("Tidak ketemu")

st.divider()
st.caption("Pramita Lab PRO - Export & Trend | Monitoring Omzet")
