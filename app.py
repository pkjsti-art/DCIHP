import io
import re
import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
import streamlit as st

# ==========================================
# KONFIGURASI HALAMAN & TAMPILAN FUTURISTIK (DARK PURPLE NEON / AI THEME)
# ==========================================
st.set_page_config(
    page_title="DATA CONTROL IHP",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS untuk mengadopsi tema UI/UX Cyberpunk / AI Dashboard (Dark & Glowing Purple)
st.markdown(
    """
    <style>
    /* Global background & text */
    .main {
        background-color: #030712;
        color: #f8fafc;
    }
    .stApp {
        background-color: #030712;
    }
    
    /* Main Header Glowing Container */
    .main-header {
        background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
        padding: 2.5rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
        border: 1px solid rgba(168, 85, 247, 0.3);
        box-shadow: 0 0 40px rgba(168, 85, 247, 0.15), inset 0 0 20px rgba(168, 85, 247, 0.1);
        text-align: center;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        background: linear-gradient(90deg, #c084fc, #e879f9, #f43f5e);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .main-header p {
        margin: 0.75rem 0 0 0;
        color: #cbd5e1;
        font-size: 1.1rem;
    }

    /* File Uploader Container (Card Style) */
    div[data-testid="stFileUploader"] {
        background: #0b0f19;
        border: 1px solid rgba(168, 85, 247, 0.2);
        padding: 1.25rem;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4);
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(168, 85, 247, 0.6);
        box-shadow: 0 0 20px rgba(168, 85, 247, 0.25);
    }

    /* Custom Gradient Button (Neon Purple Style) */
    .stButton > button {
        background: linear-gradient(135deg, #7c3aed 0%, #9333ea 50%, #c084fc 100%);
        color: white;
        font-weight: 700;
        padding: 0.85rem 2rem;
        border-radius: 10px;
        border: none;
        box-shadow: 0 0 25px rgba(147, 51, 234, 0.4);
        transition: all 0.3s ease;
        width: 100%;
        letter-spacing: 0.05em;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #6d28d9 0%, #7e22ce 50%, #a855f7 100%);
        box-shadow: 0 0 35px rgba(168, 85, 247, 0.7);
        transform: translateY(-2px);
    }

    /* Download Button Styling */
    div.stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #10b981 100%);
        color: white;
        font-weight: 700;
        border-radius: 10px;
        border: none;
        box-shadow: 0 0 20px rgba(16, 185, 129, 0.3);
    }
    div.stDownloadButton > button:hover {
        box-shadow: 0 0 30px rgba(16, 185, 129, 0.6);
    }

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #070913;
        border-right: 1px solid rgba(168, 85, 247, 0.15);
        color: #f8fafc;
    }
    section[data-testid="stSidebar"] .stMarkdown {
        color: #e2e8f0;
    }

    /* Dataframe / Table container dark mode fix */
    div[data-testid="stDataFrame"] {
        border: 1px solid rgba(168, 85, 247, 0.2);
        border-radius: 10px;
        padding: 5px;
        background-color: #0b0f19;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# Fungsi untuk membersihkan spasi gaib / karakter tersembunyi dari string
def clean_text(val):
  if isinstance(val, str):
    cleaned = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", val)
    return cleaned.strip()
  return val


def clean_dataframe(df):
  df = df.copy()
  df.columns = [clean_text(str(col)) for col in df.columns]
  for col in df.select_dtypes(include=["object", "string"]):
    df[col] = df[col].apply(clean_text)
  return df


# Fungsi pintar untuk membersihkan dan mengubah string angka
def parse_numeric(val):
  if pd.isna(val) or val == "":
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  val_str = str(val).strip()
  if "," in val_str and "." in val_str:
    val_str = val_str.replace(".", "").replace(",", ".")
  elif "," in val_str:
    val_str = val_str.replace(",", ".")
  try:
    return float(val_str)
  except ValueError:
    return 0.0


# Inisialisasi Session State
if "processed" not in st.session_state:
  st.session_state.processed = False
  st.session_state.final_report = None
  st.session_state.df_ihp = None
  st.session_state.df_memphis = None
  st.session_state.df_gu = None
  st.session_state.df_heatset = None
  st.session_state.processed_data = None


# ==========================================
# HEADER UTAMA DASHBOARD (FUTURISTIK)
# ==========================================
st.markdown(
    """
    <div class="main-header">
        <h1>✨ DATA CONTROL IHP AI</h1>
        <p>Advanced Data Processing, Validation & Automated Reconciliation System</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Sidebar Modern
st.sidebar.markdown("### ⚡ Control Panel")
menu = st.sidebar.selectbox("Pilih Menu Utama", ["Pengelolaan Data IHP"])

if menu == "Pengelolaan Data IHP":
  st.markdown("### 🔮 Unggah Dokumen Master")
  st.markdown(
      "Silakan unggah file master **IHP**, **Memphis**, dan **Benang** pada"
      " area di bawah ini:"
  )

  col1, col2, col3 = st.columns(3)

  with col1:
    uploaded_ihp = st.file_uploader(
        "1️⃣ Data Master IHP (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="ihp",
    )

  with col2:
    uploaded_memphis = st.file_uploader(
        "2️⃣ Data Master Memphis (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="memphis",
    )

  with col3:
    uploaded_benang = st.file_uploader(
        "3️⃣ Data Acuan Benang (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="benang",
    )

  st.markdown("<br>", unsafe_allow_html=True)

  if st.button("⚡ PROSES ANALISIS & REKONSILIASI", type="primary"):
    if uploaded_ihp and uploaded_memphis and uploaded_benang:
      try:
        with st.spinner(
            "🔮 AI sedang membersihkan, memetakan, dan memproses data..."
        ):
          # 1. Membaca File
          if uploaded_ihp.name.endswith(".csv"):
            df_ihp = pd.read_csv(uploaded_ihp, dtype=str)
          else:
            df_ihp = pd.read_excel(uploaded_ihp, dtype=str)

          if uploaded_memphis.name.endswith(".csv"):
            df_memphis = pd.read_csv(uploaded_memphis, dtype=str)
          else:
            df_memphis = pd.read_excel(uploaded_memphis, dtype=str)

          if uploaded_benang.name.endswith(".csv"):
            df_benang = pd.read_csv(uploaded_benang, dtype=str)
          else:
            df_benang = pd.read_excel(uploaded_benang, dtype=str)

          # Basic Cleaning
          df_ihp = clean_dataframe(df_ihp)
          df_memphis = clean_dataframe(df_memphis)
          df_benang = clean_dataframe(df_benang)

          # ==========================================
          # 2. PROCESSING DATA MEMPHIS & GU (Hapus baris total di bawah)
          # ==========================================
          if len(df_memphis) > 0:
            last_row_val = str(df_memphis.iloc[-1, 0]).strip().lower()
            if (
                "total" in last_row_val
                or last_row_val == "nan"
                or last_row_val == ""
            ):
              df_memphis = df_memphis.iloc[:-1].copy()

          if "C. Penerimaan" in df_memphis.columns:
            df_memphis["C. Penerimaan"] = df_memphis["C. Penerimaan"].apply(
                parse_numeric
            )
          if "Penerimaan" in df_memphis.columns:
            df_memphis["Penerimaan"] = df_memphis["Penerimaan"].apply(
                parse_numeric
            )

          if "C. Penerimaan" in df_memphis.columns:
            df_memphis = df_memphis[df_memphis["C. Penerimaan"] != 0]

          if "Bobin" in df_memphis.columns:
            df_memphis = df_memphis[
                df_memphis["Bobin"].astype(str).str.lower() != "tanpa bobin"
            ]

          # Pisahkan baris GU
          df_gu = pd.DataFrame(columns=df_memphis.columns)
          if "No. LOT" in df_memphis.columns:
            mask_gu = (
                df_memphis["No. LOT"]
                .astype(str)
                .str.contains("GU", case=False, na=False)
            )
            df_gu = df_memphis[mask_gu].copy()
            df_memphis = df_memphis[~mask_gu].copy()

          # Tambahkan Kolom Penerimaan (GR) di Data Memphis
          if "Penerimaan" in df_memphis.columns:
            penerimaan_gr_vals = df_memphis["Penerimaan"] * 1000
            idx_pen = df_memphis.columns.get_loc("Penerimaan")
            if "Penerimaan (GR)" in df_memphis.columns:
              df_memphis = df_memphis.drop(columns=["Penerimaan (GR)"])
            df_memphis.insert(idx_pen + 1, "Penerimaan (GR)", penerimaan_gr_vals)
          else:
            df_memphis["Penerimaan (GR)"] = 0

          # Tambahkan Kolom Penerimaan (GR) di Data GU juga
          if "Penerimaan" in df_gu.columns:
            penerimaan_gr_gu = df_gu["Penerimaan"] * 1000
            idx_pen_gu = df_gu.columns.get_loc("Penerimaan")
            if "Penerimaan (GR)" in df_gu.columns:
              df_gu = df_gu.drop(columns=["Penerimaan (GR)"])
            df_gu.insert(idx_pen_gu + 1, "Penerimaan (GR)", penerimaan_gr_gu)
          else:
            df_gu["Penerimaan (GR)"] = 0

          # ==========================================
          # 3. MAPPING BENANG
          # ==========================================
          mapping_memphis_to_ihp = {}
          if len(df_benang.columns) >= 2:
            for _, row in df_benang.iterrows():
              m_name = str(row.iloc[0]).strip()
              i_name = str(row.iloc[1]).strip()
              if m_name and m_name != "nan" and i_name and i_name != "nan":
                mapping_memphis_to_ihp[m_name] = i_name

          # ==========================================
          # 4. PROCESSING DATA MASTER IHP & HEATSET
          # ==========================================
          df_heatset = pd.DataFrame(columns=df_ihp.columns)
          if "Kode Barang Jadi" in df_ihp.columns:
            mask_heatset = (
                df_ihp["Kode Barang Jadi"]
                .astype(str)
                .str.contains("TWP.050.8", case=False, na=False)
            )
            df_heatset = df_ihp[mask_heatset].copy()
            df_ihp = df_ihp[~mask_heatset].copy()

          # Fungsi pemrosesan kalkulasi IHP / Heatset
          def process_ihp_calculations(target_df):
            target_df = target_df.copy()
            if "Target Qty" in target_df.columns:
              target_df["Target Qty"] = target_df["Target Qty"].apply(
                  parse_numeric
              )
            if "Qty" in target_df.columns:
              target_df["Qty"] = target_df["Qty"].apply(parse_numeric)

            item_c = None
            for col in ["Nama Item", "Item", "Benang", "Nama Benang"]:
              if col in target_df.columns:
                item_c = col
                break
            if not item_c:
              item_c = target_df.columns[0]

            inv_mapping = {v: k for k, v in mapping_memphis_to_ihp.items()}
            std_names = target_df[item_c].apply(
                lambda x: inv_mapping.get(x, x)
            )

            for col_drop in ["Harga", "Total"]:
              if col_drop in target_df.columns:
                target_df = target_df.drop(columns=[col_drop])

            if (
                "Target Unit" in target_df.columns
                and "Target Qty" in target_df.columns
            ):
              target_gr_values = target_df.apply(
                  lambda row: (
                      row["Target Qty"]
                      if str(row["Target Unit"]).strip().upper() == "GR"
                      else row["Target Qty"] * 1000
                  ),
                  axis=1,
              )
              idx_target_unit = target_df.columns.get_loc("Target Unit")
              if "Qty Target Standar (GR)" in target_df.columns:
                target_df = target_df.drop(columns=["Qty Target Standar (GR)"])
              target_df.insert(
                  idx_target_unit + 1,
                  "Qty Target Standar (GR)",
                  target_gr_values,
              )
            else:
              target_df["Qty Target Standar (GR)"] = 0

            if "Unit" in target_df.columns and "Qty" in target_df.columns:
              qty_gr_values = target_df.apply(
                  lambda row: (
                      row["Qty"]
                      if str(row["Unit"]).strip().upper() == "GR"
                      else row["Qty"] * 1000
                  ),
                  axis=1,
              )
              idx_qty = target_df.columns.get_loc("Qty")
              if "Qty BB Standar (GR)" in target_df.columns:
                target_df = target_df.drop(columns=["Qty BB Standar (GR)"])
              target_df.insert(
                  idx_qty + 1, "Qty BB Standar (GR)", qty_gr_values
              )
            else:
              target_df["Qty BB Standar (GR)"] = 0

            selisih_temp = (
                target_df["Qty BB Standar (GR)"]
                - target_df["Qty Target Standar (GR)"]
            )
            crosscheck_vals = selisih_temp.apply(
                lambda x: "CORRECT" if abs(x) < 1e-5 else "INCORRECT"
            )

            for col_drop in ["Crosscheck Qty", "Selisih (GR)"]:
              if col_drop in target_df.columns:
                target_df = target_df.drop(columns=[col_drop])

            idx_bb_std = target_df.columns.get_loc("Qty BB Standar (GR)")
            target_df.insert(idx_bb_std + 1, "Crosscheck Qty", crosscheck_vals)
            target_df.insert(idx_bb_std + 2, "Selisih (GR)", selisih_temp)

            if "Nama Benang Standar" in target_df.columns:
              target_df = target_df.drop(columns=["Nama Benang Standar"])
            target_df["Nama Benang Standar"] = std_names

            return target_df

          df_ihp = process_ihp_calculations(df_ihp)
          df_heatset = process_ihp_calculations(df_heatset)

          # ==========================================
          # 5. COMPARE & BUILD FINAL REPORT
          # ==========================================
          memphis_benang_col = None
          for col in ["Benang", "Nama Benang", "Item", "Nama Item"]:
            if col in df_memphis.columns:
              memphis_benang_col = col
              break
          if not memphis_benang_col:
            memphis_benang_col = df_memphis.columns[0]

          kode_col_ihp = None
          for col in ["Kode", "Kode Item", "Kode Pemakaian", "No Bukti"]:
            if col in df_ihp.columns:
              kode_col_ihp = col
              break
          if not kode_col_ihp and len(df_ihp.columns) > 0:
            kode_col_ihp = df_ihp.columns[0]

          item_col_ihp = None
          for col in ["Nama Item", "Item", "Benang", "Nama Benang"]:
            if col in df_ihp.columns:
              item_col_ihp = col
              break
          if not item_col_ihp:
            item_col_ihp = df_ihp.columns[0]

          memphis_benangs = df_memphis[memphis_benang_col].unique()
          final_rows = []

          for b_name in memphis_benangs:
            sub_memphis = df_memphis[df_memphis[memphis_benang_col] == b_name]
            total_cones = sub_memphis["C. Penerimaan"].sum()
            total_qty_m = sub_memphis["Penerimaan (GR)"].sum()

            ihp_target_name = mapping_memphis_to_ihp.get(b_name, b_name)

            sub_ihp = df_ihp[
                (df_ihp[item_col_ihp] == ihp_target_name)
                | (df_ihp["Nama Benang Standar"] == b_name)
            ]
            total_qty_ihp = sub_ihp["Qty BB Standar (GR)"].sum()

            selisih = round(total_qty_m - total_qty_ihp, 2)
            keterangan = (
                "Benang belum di IHP"
                if total_qty_ihp == 0 and len(sub_ihp) == 0
                else ""
            )

            if len(sub_ihp) > 0 and kode_col_ihp:
              kodes = sub_ihp[kode_col_ihp].dropna().unique().tolist()
            else:
              kodes = ["-"]
            if not kodes:
              kodes = ["-"]

            kode_gabung = ", ".join([str(k) for k in kodes])

            final_rows.append({
                "Nama Benang": b_name,
                "Total Cones Memphis": total_cones,
                "Total Qty Memphis": total_qty_m,
                "Total QTY IHP": total_qty_ihp,
                "Selisih IHP - Memphis": selisih,
                "Satuan": "GR",
                "Kode Pemakaian IHP": kode_gabung,
                "Keterangan": keterangan,
            })

          final_report = pd.DataFrame(final_rows)

          # ==========================================
          # 6. GENERATE EXCEL BUFFER
          # ==========================================
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_report.to_excel(
                writer, sheet_name="Hasil Olah Data", index=False
            )
            df_ihp.to_excel(
                writer, sheet_name="Data Master IHP", index=False
            )
            df_memphis.to_excel(
                writer, sheet_name="Data Memphis", index=False
            )
            df_gu.to_excel(writer, sheet_name="Data GU", index=False)
            df_heatset.to_excel(writer, sheet_name="Heatset", index=False)

            # Mengatur Wrap Text dan Lebar Kolom untuk Sheet 'Hasil Olah Data'
            workbook = writer.book
            ws = workbook["Hasil Olah Data"]

            for col in ws.iter_cols(
                min_row=1, max_row=1, min_col=1, max_col=ws.max_column
            ):
              for cell in col:
                if cell.value == "Kode Pemakaian IHP":
                  col_letter = cell.column_letter
                  ws.column_dimensions[col_letter].width = 45
                  for r in range(2, ws.max_row + 1):
                    ws[f"{col_letter}{r}"].alignment = Alignment(
                        wrap_text=True, vertical="center"
                    )

          processed_data = output.getvalue()

          # Simpan ke session state
          st.session_state.final_report = final_report
          st.session_state.df_ihp = df_ihp
          st.session_state.df_memphis = df_memphis
          st.session_state.df_gu = df_gu
          st.session_state.df_heatset = df_heatset
          st.session_state.processed_data = processed_data
          st.session_state.processed = True

      except Exception as e:
        st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")
    else:
      st.warning(
          "⚠️ Mohon unggah ketiga file master (Data Master IHP, Data Master"
          " Memphis, dan Data Benang) terlebih dahulu!"
      )

  # Tampilkan Hasil jika data sudah diproses
  if st.session_state.processed:
    st.success("🎉 Data Berhasil Diproses dan Direkonsiliasi!")
    st.markdown("### 📋 Hasil Olah Data & Rekonsiliasi")
    st.dataframe(st.session_state.final_report, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label="📥 Unduh Laporan Excel Lengkap (Multi-Sheet)",
        data=st.session_state.processed_data,
        file_name="Laporan_Data_Control_IHP.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
