import streamlit as st
import pandas as pd
import datetime
import io
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Set Page Config
st.set_page_config(
    page_title="Sistem Serah Terima Kasir - RS Adhyaksa Jatim",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Navigasi Menu Sidebar
st.sidebar.title("📌 Navigasi Menu")
menu_pilihan = st.sidebar.radio(
    "Pilih Jenis Form:",
    ["Serah Terima Shift", "Closing Harian / Tutup Shift"]
)
# Custom Styling
st.markdown("""
<style>
    .main-header { font-size: 24px; font-weight: bold; color: #1e4d2b; text-align: center; margin-bottom: 5px; }
    .sub-header { font-size: 15px; color: #444; text-align: center; margin-bottom: 20px; }
    .stButton>button { width: 100%; background-color: #1e4d2b; color: white; font-weight: bold; }
    .stAlert { padding: 8px 15px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏥 SISTEM INTEGRASI SERAH TERIMA CLOSING KASIR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Rumah Sakit Adhyaksa Jawa Timur</div>', unsafe_allow_html=True)

# Sidebar: File Upload & Configuration
st.sidebar.header("📂 1. Upload Tarikan File SIMRS")
uploaded_file = st.sidebar.file_uploader("Unggah File SIMRS (.xlsx / .csv)", type=["xlsx", "csv", "xls"])

df = None
file_parsed = False

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        file_parsed = True
        st.sidebar.success("✅ File SIMRS berhasil diuraikan!")
    except Exception as e:
        st.sidebar.error(f"❌ Gagal membaca file: {e}")
else:
    st.sidebar.info("📌 Unggah file laporan SIMRS untuk merekap transaksi & biaya admin EDC secara otomatis.")

# ==========================================
# MODUL UTAMA FORM SERAH TERIMA KASIR
# ==========================================
with st.container():
    col1, col2, col3 = st.columns([1.1, 1, 1])
    
    with col1:
        st.subheader("📋 Informasi Shift & Tanggal")
        tgl_shift = st.date_input("Tanggal Shift", value=datetime.date.today())
        shift_opt = st.selectbox("Shift Operasional", [
            "PAGI (07.00 - 14.00 WIB)",
            "SIANG (14.00 - 21.00 WIB)",
            "MALAM (21.00 - 07.00 WIB)"
        ])
        petugas_lama = st.text_input("Petugas Shift Lama (Menyerahkan)", "CHORI CHOIRUNNISA'")
        petugas_baru = st.text_input("Petugas Shift Baru (Menerima)", "ABDUL JALIL SANTRI AJI")
        pj_kasir = st.text_input("Penanggung Jawab Kasir", "")

    # Variabel Default Rekap
    penerimaan_tunai = 0.0
    total_non_tunai_bruto = 0.0
    total_biaya_admin = 0.0
    df_grouped_final = pd.DataFrame()
    df_display_clean = pd.DataFrame()

    # Processing SIMRS Data (Item-based / Multi-row format support)
    if file_parsed and df is not None:
        df.columns = [str(c).strip() for c in df.columns]
        
        cols_map = {str(c).strip().lower(): c for c in df.columns}
        col_tgl = cols_map.get('tanggal transaksi') or cols_map.get('tanggal') or cols_map.get('tgl_transaksi')
        col_notx = cols_map.get('no. transaksi') or cols_map.get('no.rawat/no.nota') or cols_map.get('no_rawat')
        col_pasien = cols_map.get('nama pasien') or cols_map.get('pasien')
        col_jenis = cols_map.get('jenis pembayaran') or cols_map.get('jenis/cara bayar') or cols_map.get('metode pembayaran')
        col_total_trx = cols_map.get('total transaksi') or cols_map.get('pendapatan bersih') or cols_map.get('subtotal after discount')
        col_item_name = cols_map.get('nama item') or cols_map.get('item')
        col_sub_item = cols_map.get('subtotal item') or cols_map.get('subtotal') or cols_map.get('harga satuan item')

        def clean_numeric(val):
            if pd.isna(val): return 0.0
            if isinstance(val, (int, float)): return float(val)
            val_str = str(val).replace('.', '').replace(',', '.')
            try:
                return float(val_str)
            except:
                return 0.0

        if col_sub_item and col_sub_item in df.columns:
            df['clean_sub_item'] = df[col_sub_item].apply(clean_numeric)
        else:
            df['clean_sub_item'] = 0.0

        # Identifikasi baris admin EDC / QRIS / QR secara spesifik berdasarkan item name
        if col_item_name and col_item_name in df.columns:
            admin_mask = df[col_item_name].astype(str).str.lower().str.contains('admin edc|admin qris|admin qr|edc|qris', na=False)
            df['is_admin_item'] = admin_mask
        else:
            df['is_admin_item'] = False

        if col_notx and col_notx in df.columns:
            trx_grouped_list = []
            for trx_id, group in df.groupby(col_notx):
                if pd.isna(trx_id) or str(trx_id).strip() == '':
                    continue
                
                tgl_val = group[col_tgl].iloc[0] if col_tgl and col_tgl in group.columns else "-"
                pasien_val = group[col_pasien].iloc[0] if col_pasien and col_pasien in group.columns else "-"
                jenis_val = group[col_jenis].iloc[0] if col_jenis and col_jenis in group.columns else "-"
                
                if col_total_trx and col_total_trx in group.columns:
                    bruto_val = clean_numeric(group[col_total_trx].iloc[0])
                else:
                    bruto_val = float(group['clean_sub_item'].sum())
                
                admin_val = float(group[group['is_admin_item']]['clean_sub_item'].sum())
                netto_val = bruto_val - admin_val
                
                trx_grouped_list.append({
                    'No. Transaksi': trx_id,
                    'Tanggal': tgl_val,
                    'Nama Pasien': pasien_val,
                    'Cara Bayar': jenis_val,
                    'clean_bersih': bruto_val,
                    'clean_admin': admin_val,
                    'clean_netto': netto_val
                })
            
            if trx_grouped_list:
                df_grouped_final = pd.DataFrame(trx_grouped_list)
                df_display_clean = df_grouped_final[['No. Transaksi', 'Tanggal', 'Nama Pasien', 'Cara Bayar', 'clean_bersih', 'clean_admin', 'clean_netto']].rename(columns={
                    'clean_bersih': 'Nominal (Rp)',
                    'clean_admin': 'Admin EDC/QRIS (Rp)',
                    'clean_netto': 'Netto Setelah Potongan (Rp)'
                })

                cash_mask = df_grouped_final['Cara Bayar'].astype(str).str.upper().str.contains('CASH|TUNAI')
                penerimaan_tunai = float(df_grouped_final[cash_mask]['clean_bersih'].sum())
                
                nontunai_mask = ~cash_mask
                total_non_tunai_bruto = float(df_grouped_final[nontunai_mask]['clean_bersih'].sum())
                total_biaya_admin = float(df_grouped_final['clean_admin'].sum())
                
                st.sidebar.success(f"📊 Auto-rekap berhasil: {len(df_grouped_final)} transaksi unik terbaca.")
                st.sidebar.info(f"🔍 Total Admin EDC/QRIS terdeteksi: Rp {total_biaya_admin:,.2f}")
        else:
            st.sidebar.warning("⚠️ Kolom No. Transaksi tidak ditemukan dalam file SIMRS.")

    with col2:
        st.subheader("💰 Transaksi Tunai (Rp)")
        modal_awal = st.number_input("Saldo Awal Kas Shift (Modal)", value=500000.0, step=50000.0)
        penerimaan_tunai = st.number_input("Penerimaan Tunai Pelayanan", value=penerimaan_tunai, step=10000.0)
        piutang_tunai = st.number_input("Pelunasan Piutang Tunai", value=0.0, step=10000.0)
        deposit_tunai = st.number_input("Penerimaan Deposit Tunai", value=0.0, step=10000.0)
        refund_tunai = st.number_input("Dikurangi: Refund Tunai", value=0.0, step=10000.0)
        
        total_tunai_netto = penerimaan_tunai + piutang_tunai + deposit_tunai - refund_tunai
        st.success(f" Total Netto Tunai: **Rp {total_tunai_netto:,.2f}**")

    with col3:
        st.subheader("💳 Transaksi Non-Tunai (Rp)")
        penerimaan_nontunai_pelayanan = st.number_input("Penerimaan Non-Tunai Pelayanan", value=total_non_tunai_bruto, step=50000.0)
        piutang_nontunai = st.number_input("Pelunasan Piutang Non-Tunai", value=0.0, step=10000.0)
        deposit_nontunai = st.number_input("Penerimaan Deposit Non-Tunai", value=0.0, step=10000.0)
        refund_nontunai = st.number_input("Dikurangi: Refund Non-Tunai", value=0.0, step=10000.0)
        total_biaya_admin = st.number_input("Potongan Biaya Admin EDC/QRIS", value=total_biaya_admin, step=1000.0)
        
        total_nontunai_bruto_all = penerimaan_nontunai_pelayanan + piutang_nontunai + deposit_nontunai - refund_nontunai
        total_non_tunai_netto = total_nontunai_bruto_all - total_biaya_admin
        st.info(f" Total Netto Non-Tunai: **Rp {total_non_tunai_netto:,.2f}**")

# ==========================================
# MODUL RINGKASAN DATA OLAHAN SIMRS
# ==========================================
if file_parsed and not df_display_clean.empty:
    with st.expander("🔍 Ringkasan Data Transaksi SIMRS (Fokus: Nominal, Admin, & Netto)", expanded=True):
        st.dataframe(df_display_clean, use_container_width=True)
        
        # --- TAMBAHAN KODE UNTUK TOTAL TUNAI & NON-TUNAI ---
        if 'Cara Bayar' in df_grouped_final.columns:
            cash_filter = df_grouped_final['Cara Bayar'].astype(str).str.upper().str.contains('CASH|TUNAI')
            tot_tunai_simrs = float(df_grouped_final[cash_filter]['clean_bersih'].sum())
            tot_nontunai_simrs = float(df_grouped_final[~cash_filter]['clean_bersih'].sum())
            tot_admin_simrs = float(df_grouped_final['clean_admin'].sum())
            
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                st.metric("Total SIMRS Tunai", f"Rp {tot_tunai_simrs:,.2f}")
            with sc2:
                st.metric("Total SIMRS Non-Tunai", f"Rp {tot_nontunai_simrs:,.2f}")
            with sc3:
                st.metric("Total Admin EDC/QRIS", f"Rp {tot_admin_simrs:,.2f}")
        # --------------------------------------------------

        # --- UBAH MENJADI FORMAT EXCEL (.xlsx) DENGAN TOTAL TUNAI & NON-TUNAI TERPISAH ---
        import io
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Tulis data utama ke sheet Excel
            df_display_clean.to_excel(writer, index=False, sheet_name='Ringkasan SIMRS')
            
            # Ambil worksheet untuk menambahkan baris total terpisah
            worksheet = writer.sheets['Ringkasan SIMRS']
            
            # Tentukan baris awal untuk total (berjarak 2 baris dari data terakhir)
            start_row = len(df_display_clean) + 3
            
            # Hitung total khusus Tunai & Non-Tunai dari df_grouped_final jika ada
            if 'Cara Bayar' in df_grouped_final.columns:
                cash_filter = df_grouped_final['Cara Bayar'].astype(str).str.upper().str.contains('CASH|TUNAI')
                
                tunai_nominal = float(df_grouped_final[cash_filter]['clean_bersih'].sum())
                tunai_admin = float(df_grouped_final[cash_filter]['clean_admin'].sum())
                tunai_netto = tunai_nominal - tunai_admin
                
                nontunai_nominal = float(df_grouped_final[~cash_filter]['clean_bersih'].sum())
                nontunai_admin = float(df_grouped_final[~cash_filter]['clean_admin'].sum())
                nontunai_netto = nontunai_nominal - nontunai_admin
            else:
                tunai_nominal, tunai_admin, tunai_netto = 0, 0, 0
                nontunai_nominal, nontunai_admin, nontunai_netto = 0, 0, 0

            # Baris 1: Total Tunai
            worksheet.cell(row=start_row, column=1, value="TOTAL TUNAI")
            try:
                worksheet.cell(row=start_row, column=5, value=tunai_nominal)
                worksheet.cell(row=start_row, column=6, value=tunai_admin)
                worksheet.cell(row=start_row, column=7, value=tunai_netto)
            except:
                pass
                
            # Baris 2: Total Non-Tunai
            row_nontunai = start_row + 1
            worksheet.cell(row=row_nontunai, column=1, value="TOTAL NON-TUNAI")
            try:
                worksheet.cell(row=row_nontunai, column=5, value=nontunai_nominal)
                worksheet.cell(row=row_nontunai, column=6, value=nontunai_admin)
                worksheet.cell(row=row_nontunai, column=7, value=nontunai_netto)
            except:
                pass
                
        excel_data = output.getvalue()

        # SIMPAN KE SESSION STATE UNTUK FORM CLOSING HARIAN
        st.session_state['simrs_tunai'] = float(tunai_netto)
        st.session_state['simrs_nontunai'] = float(nontunai_netto)
        st.session_state['simrs_piutang_tunai'] = 0.0  # Sesuaikan jika ada variabel piutang
        st.session_state['simrs_deposit_tunai'] = 0.0 # Sesuaikan jika ada variabel deposit
        st.session_state['simrs_refund_tunai'] = 0.0  # Sesuaikan jika ada variabel refund

        st.download_button(
            label="📥 Unduh Ringkasan SIMRS (XLSX)",
            data=excel_data,
            file_name=f"Ringkasan_SIMRS_{tgl_shift}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Cash Breakdown & Calculations
st.markdown("---")
st.subheader("💵 Input Pecahan Uang Fisik Kasir (Cash Count)")
col_pec1, col_pec2, col_pec3, col_pec4 = st.columns(4)

with col_pec1:
    l100 = st.number_input("100.000 (Lembar)", min_value=0, value=1)
    l50 = st.number_input("50.000 (Lembar)", min_value=0, value=1)
with col_pec2:
    l20 = st.number_input("20.000 (Lembar)", min_value=0, value=2)
    l10 = st.number_input("10.000 (Lembar)", min_value=0, value=13)
with col_pec3:
    l5 = st.number_input("5.000 (Lembar)", min_value=0, value=16)
    l2 = st.number_input("2.000 (Lembar)", min_value=0, value=25)
with col_pec4:
    l1 = st.number_input("1.000 (Lembar/Keping)", min_value=0, value=1)
    logam = st.number_input("Total Uang Logam (Rp)", min_value=0.0, value=49000.0, step=100.0)

total_kas_seharusnya = modal_awal + total_tunai_netto
total_uang_fisik = (l100 * 100000) + (l50 * 50000) + (l20 * 20000) + (l10 * 10000) + (l5 * 5000) + (l2 * 2000) + (l1 * 1000) + logam
selisih_kas = total_uang_fisik - total_kas_seharusnya
total_pendapatan_netto = total_tunai_netto + total_non_tunai_netto

st.markdown("---")
st.subheader("⏳ Transaksi / Tagihan Dalam Proses (Pending / Outstanding)")
initial_pending_data = pd.DataFrame([
    {
        "Nama / No RM": "Budi / RM-00129", 
        "Keterangan / Kendala": "Menunggu konfirmasi settlement EDC", 
        "Status / Tindak Lanjut": "Dalam Proses EDC"
    }
])
edited_pending_df = st.data_editor(
    initial_pending_data,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Nominal (Rp)": st.column_config.NumberColumn(format="Rp %.2f", step=10000.0)
    }
)

# Aman dari KeyError meskipun nama kolom diubah pada data_editor
if not edited_pending_df.empty:
    # Cari kolom yang berisi angka/nominal (biasanya kolom terakhir atau yang bertipe numerik)
    numeric_cols = edited_pending_df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        total_nominal_pending = float(edited_pending_df[numeric_cols[0]].sum())
    else:
        # Fallback ambil kolom terakhir jika tipe datanya belum terdeteksi numerik otomatis
        total_nominal_pending = float(pd.to_numeric(edited_pending_df.iloc[:, -1], errors='coerce').sum())
else:
    total_nominal_pending = 0.0

st.markdown("---")
st.subheader("📊 Hasil Rekonsiliasi Otomatis Shift")
m_col1, m_col2, m_col3, m_col4 = st.columns(4)
with m_col1:
    st.metric(label="Total Kas Fisik (Modal+Tunai)", value=f"Rp {total_kas_seharusnya:,.2f}")
with m_col2:
    st.metric(label="Uang Tunai Aktual (Brankas)", value=f"Rp {total_uang_fisik:,.2f}")
with m_col3:
    status_selisih = "PAS / SESUAI" if selisih_kas == 0 else ("LEBIH" if selisih_kas > 0 else "KURANG")
    st.metric(label=f"Selisih Kas ({status_selisih})", value=f"Rp {selisih_kas:,.2f}")
with m_col4:
    st.metric(label="Total Admin EDC / QRIS", value=f"Rp {total_biaya_admin:,.2f}")

st.markdown("---")
st.subheader("📝 Catatan Tambahan Kasir")
catatan_tambahan = st.text_area("Catatan Tambahan", "Uang lebih Rp 22 karena pasien tidak mau menerima kembalian")

# Generate PDF function
def create_pdf():
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, alignment=1, spaceAfter=2)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, alignment=1, spaceAfter=10)
    normal_bold = ParagraphStyle('NormalBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9)
    normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=8)
    
    elements = []
    elements.append(Paragraph("RUMAH SAKIT ADHYAKSA JAWA TIMUR", title_style))
    elements.append(Paragraph("FORMULIR SERAH TERIMA & CLOSING KASIR SHIFT", subtitle_style))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1e4d2b"), spaceAfter=10))
    
    meta_data = [
        [Paragraph("<b>Tanggal Shift:</b>", normal_style), Paragraph(str(tgl_shift), normal_style), Paragraph("<b>Petugas Menyerahkan:</b>", normal_style), Paragraph(petugas_lama, normal_style)],
        [Paragraph("<b>Shift Operasional:</b>", normal_style), Paragraph(shift_opt, normal_style), Paragraph("<b>Petugas Menerima:</b>", normal_style), Paragraph(petugas_baru, normal_style)]
    ]
    t_meta = Table(meta_data, colWidths=[100, 170, 110, 170])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F2F4F3")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_meta)
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b>1. REKAPITULASI PENERIMAAN KAS & NON-TUNAI</b>", normal_bold))
    rekap_data = [
        ["Uraian Penerimaan", "Penerimaan Tunai (Rp)", "Penerimaan Non-Tunai (Rp)", "Total Netto (Rp)"],
        ["Saldo Awal Kas / Modal Kembalian", f"{modal_awal:,.2f}", "-", f"{modal_awal:,.2f}"],
        ["1. Penerimaan Pelayanan", f"{penerimaan_tunai:,.2f}", f"{penerimaan_nontunai_pelayanan:,.2f}", f"{(penerimaan_tunai + penerimaan_nontunai_pelayanan):,.2f}"],
        ["2. Pelunasan Piutang", f"{piutang_tunai:,.2f}", f"{piutang_nontunai:,.2f}", f"{(piutang_tunai + piutang_nontunai):,.2f}"],
        ["3. Penerimaan Deposit", f"{deposit_tunai:,.2f}", f"{deposit_nontunai:,.2f}", f"{(deposit_tunai + deposit_nontunai):,.2f}"],
        ["4. Dikurangi: Batal / Refund", f"({refund_tunai:,.2f})", f"({refund_nontunai:,.2f})", f"({(refund_tunai + refund_nontunai):,.2f})"],
        ["5. Dikurangi: Potongan Admin EDC / QRIS", "-", f"({total_biaya_admin:,.2f})", f"({total_biaya_admin:,.2f})"],
        ["GRAND TOTAL PENDAPATAN SHIFT", f"{total_tunai_netto:,.2f}", f"{total_non_tunai_netto:,.2f}", f"{total_pendapatan_netto:,.2f}"]
    ]
    t_rekap = Table(rekap_data, colWidths=[210, 110, 115, 115])
    t_rekap.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e4d2b")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#E8F5E9")),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_rekap)
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph("<b>2. RINCIAN UANG FISIK BRANKAS (CASH COUNT)</b>", normal_bold))
    cash_data = [
        ["Pecahan", "Jumlah", "Total Nominal", "Pecahan", "Jumlah", "Total Nominal"],
        ["Rp 100.000", str(l100), f"Rp {l100*100000:,.2f}", "Rp 5.000", str(l5), f"Rp {l5*5000:,.2f}"],
        ["Rp 50.000", str(l50), f"Rp {l50*50000:,.2f}", "Rp 2.000", str(l2), f"Rp {l2*2000:,.2f}"],
        ["Rp 20.000", str(l20), f"Rp {l20*20000:,.2f}", "Rp 1000", str(l1), f"Rp {l1*1000:,.2f}"],
        ["Rp 10.000", str(l10), f"Rp {l10*10000:,.2f}", "Uang Logam", "-", f"Rp {logam:,.2f}"],
        ["TOTAL UANG FISIK AKTUAL", "", "", "", "", f"Rp {total_uang_fisik:,.2f}"],
        ["KAS SEHARUSNYA (MODAL + TUNAI)", "", "", "", "", f"Rp {total_kas_seharusnya:,.2f}"],
        [f"SELISIH KAS ({status_selisih})", "", "", "", "", f"Rp {selisih_kas:,.2f}"]
    ]
    t_cash = Table(cash_data, colWidths=[90, 50, 135, 90, 50, 135])
    t_cash.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#444444")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,0), (2,-1), 'RIGHT'),
        ('ALIGN', (4,0), (5,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('SPAN', (0,5), (4,5)),
        ('SPAN', (0,6), (4,6)),
        ('SPAN', (0,7), (4,7)),
        ('FONTNAME', (0,5), (-1,-1), 'Helvetica-Bold'),
        ('BACKGROUND', (0,7), (-1,7), colors.HexColor("#FFF3E0")),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_cash)
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("<b>3. TRANSAKSI / TAGIHAN DALAM PROSES (PENDING / OUTSTANDING)</b>", normal_bold))
    pending_table_data = [["No.", "Nama / No RM", "Keterangan / Kendala", "Status / Tindak Lanjut"]]
    
    if not edited_pending_df.empty:
    # Definisikan style paragraf untuk tabel agar teks bisa otomatis turun (wrap)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    styles = getSampleStyleSheet()
    table_text_style = ParagraphStyle(
        'TableTextCustom',
        parent=styles['Normal'],
        fontSize=8,
        leading=10
    )

    for idx, row in edited_pending_df.reset_index(drop=True).iterrows():
        try:
            nama_orm = str(row.iloc[0]) if len(row) > 0 else ""
            ket = str(row.iloc[1]) if len(row) > 1 else ""
            status_tindakan = str(row.iloc[2]) if len(row) > 2 else ""
        except:
            nama_orm = ""
            ket = ""
            status_tindakan = ""
        
        # BUNGKUS DENGAN PARAGRAPH AGAR TEKS OTOMATIS TURUN KE BAWAH
        pending_table_data.append([
            str(idx + 1),
            Paragraph(nama_orm, table_text_style),
            Paragraph(ket, table_text_style),
            status_tindakan
        ])
else:
    pending_table_data.append(["-", "Tidak ada transaksi pending", "-", "-"])

    t_pending = Table(pending_table_data, colWidths=[30, 130, 230, 114])
    t_pending.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e4d2b")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (3,1), (3,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#E3F2FD")),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_pending)
    elements.append(Spacer(1, 10))
    
    # Buat style khusus catatan agar bisa terbungkus rapi
    from reportlab.lib.styles import ParagraphStyle
    catatan_style = ParagraphStyle(
        'CatatanStylePdf', 
        parent=normal_style, 
        fontSize=8, 
        leading=10
    )

# Gunakan style tersebut pada Paragraph Catatan Kasir
elements.append(Paragraph(f"<b>Catatan Kasir:</b> {catatan_tambahan}", catatan_style))
    
    elements.append(Paragraph("<b>4. PERNYATAAN SERAH TERIMA ANTAR SHIFT</b>", normal_bold))
    pernyataan_text = "Kas, dokumen, dan informasi transaksi shift telah diperiksa dan diserahterimakan sesuai kondisi pada saat pergantian shift."
    elements.append(Paragraph(pernyataan_text, normal_style))
    elements.append(Spacer(1, 15))
    
    pj_name = pj_kasir if pj_kasir.strip() != "" else " ( .................................... ) "
    sig_data = [
        ["Petugas Shift Lama (Menyerahkan)", "Petugas Shift Baru (Menerima)", "Mengetahui (Penanggung Jawab Kasir)"],
        ["\n\n\n________________________", "\n\n\n________________________", "\n\n\n________________________"],
        [petugas_lama, petugas_baru, pj_name]
    ]
    t_sig = Table(sig_data, colWidths=[180, 180, 190])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),
    ]))
    elements.append(t_sig)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

st.markdown("---")
st.subheader("🖨️ Cetak & Unduh Dokumen Closing")
pdf_bytes = create_pdf()

st.download_button(
    label="📄 Unduh Form Closing Kasir (PDF)",
    data=pdf_bytes,
    file_name=f"Serah_Terima_Kasir_{tgl_shift}.pdf",
    mime="application/pdf"
)
# --- FORM CLOSING HARIAN TAMBAHAN DI BAGIAN BAWAH ---
if menu_pilihan == "Closing Harian / Tutup Shift":
    st.title("📑 Form Closing Harian & Tutup Shift Kasir")
    st.markdown("Form pendapatan otomatis dari tarikan SIMRS, dilengkapi rincian pecahan uang fisik.")

    # Ambil data otomatis dari session state hasil SIMRS
    default_tunai_pelayanan = st.session_state.get('simrs_tunai_pelayanan', 0.0)
    
    # PERHATIKAN INI: Perbaiki nama variabel berikut agar sesuai dengan data session state Anda
    default_piutang_t = st.session_state.get('simrs_piutang_tunai', 0.0)
    default_deposit_t = st.session_state.get('simrs_deposit_tunai', 0.0)
    default_refund_t = st.session_state.get('simrs_refund_tunai', 0.0)
    
    default_nontunai_brutto = st.session_state.get('simrs_nontunai_brutto', 0.0)
    default_nontunai_admin = st.session_state.get('simrs_nontunai_admin', 0.0)

    with st.form("form_closing_harian"):
        col1, col2, col3 = st.columns(3)
        with col1:
            tanggal_closing = st.date_input("Tanggal Closing")
        with col2:
            jam_closing = st.text_input("Jam Closing (WIB)", value="07.30")
        with col3:
            nama_kasir = st.text_input("Penanggung Jawab Kasir", value="")
            
        st.subheader("1. Pendapatan & Transaksi (Otomatis dari SIMRS)")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### Sisi Tunai")
            # Menggunakan variabel default yang sudah diperbaiki namanya di atas
            penerimaan_tunai = st.number_input("Penerimaan Tunai Pelayanan (Rp)", min_value=0.0, value=default_tunai_pelayanan, step=1000.0)
            piutang_tunai = st.number_input("Pembayaran Piutang Tunai (Rp)", min_value=0.0, value=default_piutang_t, step=1000.0)
            deposit_tunai = st.number_input("Penerimaan Deposit Tunai (Rp)", min_value=0.0, value=default_deposit_t, step=1000.0)
            refund_tunai = st.number_input("Pengembalian / Refund Tunai (Rp)", min_value=0.0, value=default_refund_t, step=1000.0)
            
        with c2:
            st.markdown("##### Sisi Non-Tunai")
            penerimaan_nontunai = st.number_input("Penerimaan Non-Tunai (QRIS/EDC) (Rp)", min_value=0.0, value=default_nontunai_brutto, step=1000.0)
            biaya_admin = st.number_input("Biaya Admin EDC/QRIS (Rp)", min_value=0.0, value=default_nontunai_admin, step=1000.0)
            
            # Total Non-Tunai Bersih terhitung otomatis
            total_nontunai_bersih = penerimaan_nontunai - biaya_admin
            st.info(f"**Total Non-Tunai Netto:** Rp {total_nontunai_bersih:,.2f}")

        st.subheader("2. Rincian Pecahan Uang Tunai Fisik")
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1:
            l_100k = st.number_input("Lembar 100.000", min_value=0, value=0, step=1)
            l_50k = st.number_input("Lembar 50.000", min_value=0, value=0, step=1)
        with col_f2:
            l_20k = st.number_input("Lembar 20.000", min_value=0, value=0, step=1)
            l_10k = st.number_input("Lembar 10.000", min_value=0, value=0, step=1)
        with col_f3:
            l_5k = st.number_input("Lembar 5.000", min_value=0, value=0, step=1)
            l_2k = st.number_input("Lembar 2.000", min_value=0, value=0, step=1)
        with col_f4:
            l_1k = st.number_input("Lembar 1.000", min_value=0, value=0, step=1)
            logam = st.number_input("Total Koin / Logam (Rp)", min_value=0.0, value=0.0, step=500.0)

        catatan_closing = st.text_area("Catatan Tambahan", placeholder="Tuliskan catatan atau kendala jika ada...")
        submitted_closing = st.form_submit_button("Generate Laporan Closing PDF")

    if submitted_closing:
        total_tunai_sebelum = penerimaan_tunai + piutang_tunai + deposit_tunai - refund_tunai
        total_nontunai_bersih = penerimaan_nontunai + piutang_nontunai + deposit_nontunai - biaya_admin
        total_fisik = (l_100k * 100000) + (l_50k * 50000) + (l_20k * 20000) + (l_10k * 10000) + (l_5k * 5000) + (l_2k * 2000) + (l_1k * 1000) + logam
        selisih_fisik = total_fisik - total_tunai_sebelum

        st.success("Data Closing berhasil dihitung! Silakan unduh PDF di bawah ini:")

        def generate_closing_pdf():
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            styles = getSampleStyleSheet()
            
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=12, alignment=1, fontName='Helvetica-Bold')
            normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontSize=8, fontName='Helvetica')
            bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontSize=8, fontName='Helvetica-Bold')

            elements.append(Paragraph("FORM SERAH TERIMA KAS CLOSING KASIR", title_style))
            elements.append(Spacer(1, 10))
            
            meta_data = [
                [Paragraph(f"<b>Tanggal :</b> {tanggal_closing}", normal_style), Paragraph(f"<b>Jam Closing :</b> {jam_closing} WIB", normal_style)],
                [Paragraph(f"<b>Penanggung Jawab Kasir :</b> {nama_kasir}", normal_style), Paragraph("<b>Bendahara Penerimaan :</b>", normal_style)]
            ]
            t_meta = Table(meta_data, colWidths=[270, 270])
            t_meta.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
            elements.append(t_meta)
            elements.append(Spacer(1, 10))

            # Tabel Pendapatan (Menggunakan Paragraph agar tag <b> terbaca rapi)
            tabel_pendapatan = [
                ["No", "Uraian", "Nominal (Rp)", "No", "Uraian", "Nominal (Rp)"],
                ["1", "Penerimaan Tunai Pelayanan", f"{penerimaan_tunai:,.2f}", "5", "Penerimaan Non-Tunai (QRIS/EDC)", f"{penerimaan_nontunai:,.2f}"],
                ["2", "Pembayaran Piutang Tunai", f"{piutang_tunai:,.2f}", "6", "Pembayaran Piutang Non-Tunai", f"{piutang_nontunai:,.2f}"],
                ["3", "Penerimaan Deposit Tunai", f"{deposit_tunai:,.2f}", "7", "Penerimaan Deposit Non-Tunai", f"{deposit_nontunai:,.2f}"],
                ["4", "Pengembalian / Refund Tunai", f"{refund_tunai:,.2f}", "8", "Biaya Admin EDC/QRIS", f"{biaya_admin:,.2f}"],
                [
                    Paragraph("<b>T</b>", bold_style), 
                    Paragraph("<b>TOTAL KAS SHIFT SEBELUM SERAH TERIMA</b>", bold_style), 
                    Paragraph(f"<b>{total_tunai_sebelum:,.2f}</b>", bold_style), 
                    Paragraph("<b>T</b>", bold_style), 
                    Paragraph("<b>TOTAL PENERIMAAN NON TUNAI BERSIH</b>", bold_style), 
                    Paragraph(f"<b>{total_nontunai_bersih:,.2f}</b>", bold_style)
                ]
            ]
            t_pend = Table(tabel_pendapatan, colWidths=[20, 160, 90, 20, 160, 90])
            t_pend.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#e0e0e0")),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (2,0), (2,-1), 'RIGHT'),
                ('ALIGN', (5,0), (5,-1), 'RIGHT'),
            ]))
            elements.append(t_pend)
            elements.append(Spacer(1, 10))

            # Tabel Ringkasan Serah Terima
            elements.append(Paragraph("<b>RINGKASAN SERAH TERIMA</b>", bold_style))
            tabel_serah = [
                ["No", "Uraian", "Nominal (Rp)", "Keterangan"],
                ["1", "Total Penerimaan Tunai", f"{penerimaan_tunai:,.2f}", ""],
                ["2", "Dikurangi : Pengembalian / Refund Tunai", f"{refund_tunai:,.2f}", ""],
                ["3", "Total uang tunai yang seharusnya diserahkan", f"{total_tunai_sebelum:,.2f}", ""],
                ["4", "Total uang tunai aktual yang diserahkan (Fisik)", f"{total_fisik:,.2f}", ""],
                ["5", "Selisih (+/-)", f"{selisih_fisik:,.2f}", "Diisi setelah penghitungan fisik"]
            ]
            t_s = Table(tabel_serah, colWidths=[20, 240, 100, 180])
            t_s.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#e0e0e0")),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('ALIGN', (2,1), (2,-1), 'RIGHT'),
            ]))
            elements.append(t_s)
            elements.append(Spacer(1, 10))

            # Tabel Pecahan Fisik
            elements.append(Paragraph("<b>RINCIAN UANG TUNAI YANG DISERAHKAN</b>", bold_style))
            tabel_fisik = [
                ["Pecahan", "Jumlah Lembar / Keping", "Total (Rp)", "Pecahan", "Jumlah Lembar / Keping", "Total (Rp)"],
                ["100.000", str(l_100k), f"{l_100k*100000:,.2f}", "5.000", str(l_5k), f"{l_5k*5000:,.2f}"],
                ["50.000", str(l_50k), f"{l_50k*50000:,.2f}", "2.000", str(l_2k), f"{l_2k*2000:,.2f}"],
                ["20.000", str(l_20k), f"{l_20k*20000:,.2f}", "1.000", str(l_1k), f"{l_1k*1000:,.2f}"],
                ["10.000", str(l_10k), f"{l_10k*10000:,.2f}", "Logam", "-", f"{logam:,.2f}"],
            ]
            t_f = Table(tabel_fisik, colWidths=[60, 110, 100, 60, 110, 100])
            t_f.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#e0e0e0")),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('ALIGN', (2,0), (2,-1), 'RIGHT'),
                ('ALIGN', (5,0), (5,-1), 'RIGHT'),
            ]))
            elements.append(t_f)
            elements.append(Spacer(1, 15))

            # Bagian Tanda Tangan (DIPERIKSA / DITERIMA OLEH)
            elements.append(Paragraph("<b>DIPERIKSA / DITERIMA OLEH:</b>", bold_style))
            elements.append(Spacer(1, 5))
            tabel_ttd = [
                ["Nama Petugas", "Jabatan", "Tanda Tangan", "Waktu"],
                [nama_kasir or "...........................", "Penanggung Jawab Kasir", "\n\n", ""],
                ["...........................", "Bendahara Penerimaan", "\n\n", ""]
            ]
            t_ttd_obj = Table(tabel_ttd, colWidths=[140, 150, 120, 130])
            t_ttd_obj.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#e0e0e0")),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            elements.append(t_ttd_obj)

            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

# --- TEMPATKAN KODE DI SINI (SEJAJAR KIRI / DI LUAR FUNGSI) ---
        pdf_data = generate_closing_pdf()
        st.download_button(
            label="📥 Unduh PDF Closing Harian",
            data=pdf_data,
            file_name=f"Closing_Kasir_{tanggal_closing}.pdf",
            mime="application/pdf"
            )
