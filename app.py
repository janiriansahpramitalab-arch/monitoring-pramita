import fitz, pandas as pd, os, io
from datetime import datetime
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="Monitoring Pramita PRO V4", layout="wide", page_icon="📊")
DB_FILE = "Database_Monitoring_Pramita.xlsx"

def to_int(s):
    try: return int(str(s).replace('.','').replace(',','').strip())
    except: return 0

def parse_pdf(pdf_path, periode_label):
    doc = fitz.open(pdf_path)
    parsed = []
    is_new_format = False
    for p in range(min(2, len(doc))):
        if "CIK DI TIRO" in doc[p].get_text("text"):
            is_new_format = True
            break

    if is_new_format: # FORMAT JUNI (CIK DI TIRO / SULTAN AGUNG)
        for p_idx in range(len(doc)):
            txt = doc[p_idx].get_text("text")
            lines = [l.strip() for l in txt.splitlines() if l.strip()!='']
            pic="UNKNOWN"
            for i,l in enumerate(lines):
                if "Tanggal" in l or "s/d" in l:
                    for k in range(i+1, min(i+6, len(lines))):
                        cand = lines[k].strip()
                        if len(cand)>3 and cand.upper()==cand and "JASA" not in cand and "PENGAMBILAN" not in cand and "PRE" not in cand:
                            if not any(ch.isdigit() for ch in cand) and len(cand.split())<=4:
                                pic=cand.title(); break
                    break
            i=0
            while i < len(lines):
                l = lines[i]
                cd = ''.join(c for c in l if c.isdigit())
                is_kode = len(cd)==10 and l.replace('.','').isdigit()
                if is_kode:
                    kode=cd
                    nama=lines[i+1] if i+1<len(lines) and any(c.isalpha() for c in lines[i+1]) else ""
                    # skip bulan
                    offset = 2 if nama else 1
                    if i+offset < len(lines) and lines[i+offset] in ["6","7","8","9","10","11","12"]:
                        offset+=1
                    nums=[]
                    j=i+offset
                    while j < len(lines) and len(nums)<12:
                        cur=lines[j]
                        c2=''.join(c for c in cur if c.isdigit())
                        if len(c2)==10 and cur.replace('.','').isdigit(): break
                        if cur.upper().startswith("TOTAL"): break
                        if cur.startswith("Yang Membuat"): break
                        tc=cur.replace('.','').replace(',','')
                        if tc.isdigit(): nums.append(tc)
                        j+=1
                    if len(nums)>=4:
                        if len(nums)>=12:
                            cik_total=to_int(nums[0]); cik_psn=to_int(nums[3])
                            sultan_total=to_int(nums[4]); sultan_psn=to_int(nums[7])
                            total_omzet=to_int(nums[8]); total_psn=to_int(nums[10])
                        else:
                            cik_total=to_int(nums[0]); cik_psn=to_int(nums[3]) if len(nums)>3 else 0
                            sultan_total=to_int(nums[4]) if len(nums)>4 else 0
                            sultan_psn=to_int(nums[7]) if len(nums)>7 else 0
                            total_omzet=cik_total+sultan_total
                            total_psn=cik_psn+sultan_psn
                            if len(nums)>=12:
                                total_omzet=to_int(nums[8]); total_psn=to_int(nums[10])
                        parsed.append({"Periode":periode_label,"PIC":pic,"Kode_Dokter":kode,"Nama_Dokter":nama[:60],"CD_Omzet":cik_total,"CD_Pasien":cik_psn,"SA_Omzet":sultan_total,"SA_Pasien":sultan_psn,"Total_Omzet":total_omzet,"Total_Pasien":total_psn})
                    i=j-1
                i+=1
    else: # FORMAT LAMA SEP
        for txt in [doc[p].get_text("text") for p in range(len(doc))]:
            lines=txt.splitlines()
            pic="UNKNOWN"
            for i in range(len(lines)):
                if "Tanggal" in lines[i] or "s/d" in lines[i]:
                    for k in range(i+1, min(i+6, len(lines))):
                        cand=lines[k].strip()
                        if len(cand)>5 and cand.upper()==cand and "JASA" not in cand:
                            if not any(ch.isdigit() for ch in cand):
                                pic=cand; break
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
                    if tt in ["7","8","9","10","11","12"]: bln=j; break
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

# UI SAMA SEPERTI V3 + HAPUS
st.title("📊 Monitoring Pramita PRO V4 - Support Juni")
with st.sidebar:
    st.header("📂 Upload")
    periode=st.text_input("Periode", value="JUN-2026")
    up=st.file_uploader("Upload PDF", type=["pdf"])
    if up:
        open("temp.pdf","wb").write(up.getbuffer())
        df_new=parse_pdf("temp.pdf", periode)
        st.success(f"Terbaca {len(df_new)} dokter")
        if len(df_new)>0:
            st.dataframe(df_new.head(10))
            if st.button("💾 Simpan"):
                if os.path.exists(DB_FILE):
                    old=pd.read_excel(DB_FILE); old=old[old["Periode"]!=periode]
                    all_df=pd.concat([old,df_new], ignore_index=True)
                else: all_df=df_new
                all_df.to_excel(DB_FILE,index=False); st.success("Tersimpan!"); st.rerun()
    st.divider()
    st.header("🗑️ Hapus Data")
    if os.path.exists(DB_FILE):
        df_tmp=pd.read_excel(DB_FILE)
        if not df_tmp.empty:
            del_per=st.selectbox("Hapus periode", df_tmp["Periode"].unique())
            if st.button(f"Hapus {del_per}"):
                df_del=df_tmp[df_tmp["Periode"]!=del_per]; df_del.to_excel(DB_FILE,index=False); st.rerun()
            if st.button("⚠️ HAPUS SEMUA"):
                os.remove(DB_FILE); st.rerun()

if not os.path.exists(DB_FILE): st.info("Upload dulu"); st.stop()
df=pd.read_excel(DB_FILE)
with st.sidebar:
    st.divider()
    sel_periode=st.multiselect("Pilih Periode", sorted(df["Periode"].unique()), default=sorted(df["Periode"].unique()))
    sel_pic=st.selectbox("PIC", ["Semua"]+sorted(df["PIC"].dropna().unique().tolist()))
    sort_by=st.selectbox("Ranking", ["Total_Omzet","Total_Pasien","CD_Omzet","SA_Omzet"])
df_f=df[df["Periode"].isin(sel_periode)] if sel_periode else df
if sel_pic!="Semua": df_f=df_f[df_f["PIC"]==sel_pic]
c1,c2,c3,c4=st.columns(4)
c1.metric("Total Omzet", f"Rp {df_f['Total_Omzet'].sum():,}"); c2.metric("Pasien", f"{df_f['Total_Pasien'].sum():,}")
c3.metric("Dokter", f"{df_f['Kode_Dokter'].nunique()}"); c4.metric("PIC", f"{df_f['PIC'].nunique()}")
st.subheader(f"🏆 Ranking {sort_by}")
df_rank=df_f.sort_values(by=sort_by, ascending=False)
buf=io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as w:
    df_rank.to_excel(w, index=False, sheet_name='Ranking')
    df_f.groupby("PIC")[["Total_Omzet","Total_Pasien"]].sum().sort_values("Total_Omzet",ascending=False).to_excel(w, sheet_name='Rekap_PIC')
st.download_button("📥 Export Excel", data=buf.getvalue(), file_name=f"Ranking_{'_'.join(sel_periode)}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.dataframe(df_rank, use_container_width=True, height=400)
st.divider(); st.subheader("📅 Trend MoM")
trend=df.groupby("Periode")[["Total_Omzet","Total_Pasien","Kode_Dokter"]].agg({"Total_Omzet":"sum","Total_Pasien":"sum","Kode_Dokter":"nunique"}).reset_index()
st.line_chart(trend.set_index("Periode")[["Total_Omzet"]]); st.dataframe(trend, use_container_width=True)
