import io
import re
import pandas as pd
import streamlit as st

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="DATA CONTROL IHP", page_icon="📊", layout="wide"
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


# Fungsi pintar untuk membersihkan dan mengubah string angka (menangani format koma/titik)
def parse_numeric(val):
  if pd.isna(val) or val == "":
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  # Jika berupa string, bersihkan karakter koma/titik pemisah ribuan jika ada
  val_str = str(val).strip()
  # Ganti koma jadi titik jika format desimal menggunakan koma
  # Contoh: "1.250,50" -> "1250.50"
  if "," in val_str and "." in val_str:
    val_str = val_str.replace(".", "").replace(",", ".")
  elif "," in val_str:
    val_str = val_str.replace(",", ".")
  try:
    return float(val_str)
  except ValueError:
    return 0.0


# Title Aplikasi
st.title("DATA CONTROL IHP")

# Sidebar
st.sidebar.header("Menu Pengelolaan Data IHP")
menu = st.sidebar.selectbox("Pilih Menu", ["Pengelolaan Data IHP"])

if menu == "Pengelolaan Data IHP":
  st.subheader("Upload File Master")

  uploaded_ihp = st.file_uploader(
      "Upload Data Master IHP (Excel/CSV)", type=["xlsx", "xls", "csv"]
  )
  uploaded_memphis = st.file_uploader(
      "Upload Data Master Memphis (Excel/CSV)", type=["xlsx", "xls", "csv"]
  )
  uploaded_benang = st.file_uploader(
      "Upload Data Benang (Excel/CSV)", type=["xlsx", "xls", "csv"]
  )

  if st.button("Proses Data", type="primary"):
    if uploaded_ihp and uploaded_memphis and uploaded_benang:
      try:
        with st.spinner(
            "Sedang memproses dan membersihkan data, mohon tunggu..."
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

          # Hapus baris dengan C. Penerimaan bernilai 0
          if "C. Penerimaan" in df_memphis.columns:
            df_memphis = df_memphis[df_memphis["C. Penerimaan"] != 0]

          # Hapus baris yang nilai Bobin-nya "Tanpa Bobin"
          if "Bobin" in df_memphis.columns:
            df_memphis = df_memphis[
                df_memphis["Bobin"].astype(str).str.lower() != "tanpa bobin"
            ]

          # Pisahkan baris "GU" pada kolom No. LOT
          df_gu = pd.DataFrame(columns=df_memphis.columns)
          if "No. LOT" in df_memphis.columns:
            mask_gu = (
                df_memphis["No. LOT"]
                .astype(str)
                .str.contains("GU", case=False, na=False)
            )
            df_gu = df_memphis[mask_gu].copy()
            df_memphis = df_memphis[~mask_gu].copy()

          # Buat kolom baru Penerimaan (GR)
          if "Penerimaan" in df_memphis.columns:
            df_memphis["Penerimaan (GR)"] = df_memphis["Penerimaan"] * 1000
          else:
            df_memphis["Penerimaan (GR)"] = 0

          # ==========================================
          # 3. PROCESSING DATA MASTER IHP
          # ==========================================
          # Pisahkan baris "Heatset" (TWP.050.8 pada Kode Barang Jadi)
          df_heatset = pd.DataFrame(columns=df_ihp.columns)
          if "Kode Barang Jadi" in df_ihp.columns:
            mask_heatset = (
                df_ihp["Kode Barang Jadi"]
                .astype(str)
                .str.contains("TWP.050.8", case=False, na=False)
            )
            df_heatset = df_ihp[mask_heatset].copy()
            df_ihp = df_ihp[~mask_heatset].copy()

          # Konversi numerik target & qty dengan parser pintar
          if "Target Qty" in df_ihp.columns:
            df_ihp["Target Qty"] = df_ihp["Target Qty"].apply(parse_numeric)
          if "Qty" in df_ihp.columns:
            df_ihp["Qty"] = df_ihp["Qty"].apply(parse_numeric)

          # Qty Target Standar (GR) - Dibuat TEPAT di sebelah kanan kolom 'Target Unit'
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
            df_ihp.insert(
                idx_target_unit + 1,
                "Qty Target Standar (GR)",
                target_gr_values,
            )
          else:
            df_ihp["Qty Target Standar (GR)"] = 0

          # Qty BB Standar (GR) - Dibuat TEPAT di sebelah kanan kolom 'Qty'
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
            df_ihp.insert(idx_qty + 1, "Qty BB Standar (GR)", qty_gr_values)
          else:
            df_ihp["Qty BB Standar (GR)"] = 0

          # Crosscheck Qty & Selisih (GR)
          df_ihp["Selisih (GR)"] = (
              df_ihp["Qty BB Standar (GR)"] - df_ihp["Qty Target Standar (GR)"]
          )
          df_ihp["Crosscheck Qty"] = df_ihp["Selisih (GR)"].apply(
              lambda x: "CORRECT" if x == 0 else "INCORRECT"
          )

          # ==========================================
          # 4. MAPPING BENANG & COMPARE DATA
          # ==========================================
          mapping_dict = {}
          if len(df_benang.columns) >= 2:
            for _, row in df_benang.iterrows():
              k, v = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
              mapping_dict[k] = v
              mapping_dict[v] = k

          if "Benang" in df_memphis.columns:
            df_memphis["Benang_Standard"] = df_memphis["Benang"].apply(
                lambda x: mapping_dict.get(x, x)
            )
          else:
            df_memphis["Benang_Standard"] = ""

          item_col_ihp = None
          for col in ["Nama Item", "Item", "Benang", "Nama Benang"]:
            if col in df_ihp.columns:
              item_col_ihp = col
              break

          if item_col_ihp:
            df_ihp["Item_Standard"] = df_ihp[item_col_ihp].apply(
                lambda x: mapping_dict.get(x, x)
            )
          else:
            df_ihp["Item_Standard"] = df_ihp.iloc[:, 0]

          # Agregasi Memphis per Benang
          memphis_agg = (
              df_memphis.groupby("Benang_Standard")
              .agg(
                  Total_Cones_Memphis=("C. Penerimaan", "sum"),
                  Total_Qty_Memphis=("Penerimaan (GR)", "sum"),
              )
              .reset_index()
          )

          # Agregasi IHP per Item/Benang
          ihp_agg = (
              df_ihp.groupby("Item_Standard")
              .agg(
                  Total_Qty_IHP=("Qty BB Standar (GR)", "sum"),
                  Kode_Pemakaian_IHP=(
                      "Kode" if "Kode" in df_ihp.columns else df_ihp.columns[0],
                      "first",
                  ),
              )
              .reset_index()
          )

          merged_result = pd.merge(
              memphis_agg,
              ihp_agg,
              left_on="Benang_Standard",
              right_on="Item_Standard",
              how="outer",
          ).fillna(0)

          # Susun Tabel Akhir Hasil Olah Data
          final_report = pd.DataFrame()
          final_report["Nama Benang"] = merged_result.apply(
              lambda row: (
                  row["Item_Standard"]
                  if row["Benang_Standard"] == 0
                  or pd.isna(row["Benang_Standard"])
                  else row["Benang_Standard"]
              ),
              axis=1,
          )
          final_report["Total Cones Memphis"] = merged_result[
              "Total_Cones_Memphis"
          ]
          final_report["Total Qty Memphis"] = merged_result[
              "Total_Qty_Memphis"
          ]
          final_report["Total QTY IHP"] = merged_result["Total_Qty_IHP"]
          final_report["Selisih IHP - Memphis"] = (
              final_report["Total Qty Memphis"]
              - final_report["Total QTY IHP"]
          )
          final_report["Satuan"] = "GR"
          final_report["Kode Pemakaian IHP"] = merged_result[
              "Kode_Pemakaian_IHP"
          ]

          st.success("Data Berhasil Diproses!")

          # Tampilkan Tabel
          st.subheader("Tabel Hasil Olah Data")
          st.dataframe(final_report, use_container_width=True)

          # ==========================================
          # 5. DOWNLOAD EXCEL MULTI-SHEET
          # (Hapus kolom bantu internal sebelum diexport ke Excel)
          # ==========================================
          df_ihp_export = df_ihp.drop(columns=["Item_Standard"], errors="ignore")
          df_memphis_export = df_memphis.drop(
              columns=["Benang_Standard"], errors="ignore"
          )
          df_gu_export = df_gu.drop(
              columns=["Benang_Standard"], errors="ignore"
          )
          df_heatset_export = df_heatset.drop(
              columns=["Item_Standard"], errors="ignore"
          )

          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_report.to_excel(
                writer, sheet_name="Hasil Olah Data", index=False
            )
            df_ihp_export.to_excel(
                writer, sheet_name="Data Master IHP", index=False
            )
            df_memphis_export.to_excel(
                writer, sheet_name="Data Memphis", index=False
            )
            df_gu_export.to_excel(writer, sheet_name="Data GU", index=False)
            df_heatset_export.to_excel(writer, sheet_name="Heatset", index=False)
          processed_data = output.getvalue()

          st.download_button(
              label="📥 Unduh Data Excel (Multi-Sheet)",
              data=processed_data,
              file_name="Laporan_Data_Control_IHP.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

      except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
    else:
      st.warning(
          "Mohon unggah ketiga file (Data Master IHP, Data Master Memphis, dan"
          " Data Benang) terlebih dahulu!"
      )
