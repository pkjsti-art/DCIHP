import io
import re
import pandas as pd
import streamlit as st

# ==========================================
# KONFIGURASI HALAMAN & TAMPILAN ESTETIK (UI/UX)
# ==========================================
st.set_page_config(
    page_title="DATA CONTROL IHP",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS untuk memperindah tampilan Streamlit (Modern & Trendy)
st.markdown(
    """
    <style>
    .main {
        background-color: #f8fafc;
    }
    .main-header {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-header h1 {
        margin: 0;
        font-size: 2.2rem;
        font-weight: 700;
        letter-spacing: -0.025em;
    }
    .main-header p {
        margin: 0.5rem 0 0 0;
        color: #94a3b8;
        font-size: 1rem;
    }
    .stButton > button {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        font-weight: 600;
        padding: 0.75rem 1.5rem;
        border-radius: 8px;
        border: none;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        box-shadow: 0 6px 8px -1px rgba(37, 99, 235, 0.3);
        transform: translateY(-1px);
    }
    section[data-testid="stSidebar"] {
        background-color: #0f172a;
        color: white;
    }
    section[data-testid="stSidebar"] .stMarkdown {
        color: #cbd5e1;
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


# Inisialisasi Session State agar data tidak hilang saat tombol download diklik
if "processed" not in st.session_state:
  st.session_state.processed = False
  st.session_state.final_report = None
  st.session_state.df_ihp = None
  st.session_state.df_memphis = None
  st.session_state.df_gu = None
  st.session_state.df_heatset = None
  st.session_state.processed_data = None


# ==========================================
# HEADER UTAMA DASHBOARD
# ==========================================
st.markdown(
    """
    <div class="main-header">
        <h1>📊 DATA CONTROL IHP</h1>
        <p>Sistem Pengelolaan, Validasi, dan Rekonsiliasi Data Produksi & Inventory</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Sidebar Modern
st.sidebar.markdown("### ⚙️ Navigasi")
menu = st.sidebar.selectbox("Pilih Menu Utama", ["Pengelolaan Data IHP"])

if menu == "Pengelolaan Data IHP":
  st.markdown("### 📂 Unggah Dokumen Master")
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

  if st.button("🚀 Proses Analisis & Rekonsiliasi Data", type="primary"):
    if uploaded_ihp and uploaded_memphis and uploaded_benang:
      try:
        with st.spinner(
            "✨ Sistem sedang membersihkan dan memproses data secara"
            " otomatis..."
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
          # 2. PROCESSING DATA MEMPHIS
          # ==========================================
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

          df_gu = pd.DataFrame(columns=df_memphis.columns)
          if "No. LOT" in df_memphis.columns:
            mask_gu = (
                df_memphis["No. LOT"]
                .astype(str)
                .str.contains("GU", case=False, na=False)
            )
            df_gu = df_memphis[mask_gu].copy()
            df_memphis = df_memphis[~mask_gu].copy()

          if "Penerimaan" in df_memphis.columns:
            penerimaan_gr_vals = df_memphis["Penerimaan"] * 1000
            idx_pen = df_memphis.columns.get_loc("Penerimaan")
            if "Penerimaan (GR)" in df_memphis.columns:
              df_memphis = df_memphis.drop(columns=["Penerimaan (GR)"])
            df_memphis.insert(idx_pen + 1, "Penerimaan (GR)", penerimaan_gr_vals)
          else:
            df_memphis["Penerimaan (GR)"] = 0

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
          # 4. PROCESSING DATA MASTER IHP
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

          if "Target Qty" in df_ihp.columns:
            df_ihp["Target Qty"] = df_ihp["Target Qty"].apply(parse_numeric)
          if "Qty" in df_ihp.columns:
            df_ihp["Qty"] = df_ihp["Qty"].apply(parse_numeric)

          item_col_ihp = None
          for col in ["Nama Item", "Item", "Benang", "Nama Benang"]:
            if col in df_ihp.columns:
              item_col_ihp = col
              break
          if not item_col_ihp:
            item_col_ihp = df_ihp.columns[0]

          inv_mapping = {v: k for k, v in mapping_memphis_to_ihp.items()}
          std_names = df_ihp[item_col_ihp].apply(
              lambda x: inv_mapping.get(x, x)
          )

          for col_drop in ["Harga", "Total"]:
            if col_drop in df_ihp.columns:
              df_ihp = df_ihp.drop(columns=[col_drop])

          if "Target Unit" in df_ihp.columns and "Target Qty" in df_ihp.columns:
            target_gr_values = df_ihp.apply(
                lambda row: (
                    row["Target Qty"]
                    if str(row["Target Unit"]).strip().upper() == "GR"
                    else row["Target Qty"] * 1000
                ),
                axis=1,
            )
            idx_target_unit = df_ihp.columns.get_loc("Target Unit")
            if "Qty Target Standar (GR)" in df_ihp.columns:
              df_ihp = df_ihp.drop(columns=["Qty Target Standar (GR)"])
            df_ihp.insert(
                idx_target_unit + 1,
                "Qty Target Standar (GR)",
                target_gr_values,
            )
          else:
            df_ihp["Qty Target Standar (GR)"] = 0

          if "Unit" in df_ihp.columns and "Qty" in df_ihp.columns:
            qty_gr_values = df_ihp.apply(
                lambda row: (
                    row["Qty"]
                    if str(row["Unit"]).strip().upper() == "GR"
                    else row["Qty"] * 1000
                ),
                axis=1,
            )
            idx_qty = df_ihp.columns.get_loc("Qty")
            if "Qty BB Standar (GR)" in df_ihp.columns:
              df_ihp = df_ihp.drop(columns=["Qty BB Standar (GR)"])
            df_ihp.insert(idx_qty + 1, "Qty BB Standar (GR)", qty_gr_values)
          else:
            df_ihp["Qty BB Standar (GR)"] = 0

          selisih_temp = (
              df_ihp["Qty BB Standar (GR)"] - df_ihp["Qty Target Standar (GR)"]
          )
          crosscheck_vals = selisih_temp.apply(
              lambda x: "CORRECT" if abs(x) < 1e-5 else "INCORRECT"
          )

          for col_drop in ["Crosscheck Qty", "Selisih (GR)"]:
            if col_drop in df_ihp.columns:
              df_ihp = df_ihp.drop(columns=[col_drop])

          idx_bb_std = df_ihp.columns.get_loc("Qty BB Standar (GR)")
          df_ihp.insert(idx_bb_std + 1, "Crosscheck Qty", crosscheck_vals)
          df_ihp.insert(idx_bb_std + 2, "Selisih (GR)", selisih_temp)

          if "Nama Benang Standar" in df_ihp.columns:
            df_ihp = df_ihp.drop(columns=["Nama Benang Standar"])
          df_ihp["Nama Benang Standar"] = std_names

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
          processed_data = output.getvalue()

          # Simpan ke session state agar aman dari rerun saat download
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

  # Tampilkan Hasil jika data sudah diproses (Persisten dengan session_state)
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
