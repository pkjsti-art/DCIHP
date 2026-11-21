import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="DATA CONTROL IHP",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main {
        background-color: #030712;
        color: #f8fafc;
    }
    .stApp {
        background-color: #030712;
    }
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
    section[data-testid="stSidebar"] {
        background-color: #070913;
        border-right: 1px solid rgba(168, 85, 247, 0.15);
        color: #f8fafc;
    }
    section[data-testid="stSidebar"] .stMarkdown {
        color: #e2e8f0;
    }
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


def try_parse_strict_number(val):
  if pd.isna(val):
    return True, None
  if isinstance(val, (int, float)):
    return True, float(val)
  s = str(val).strip()
  if s == "" or s.lower() == "nan":
    return True, None
  if re.match(r"^[+-]?0\d+$", s):
    return False, None
  try:
    return True, float(s)
  except ValueError:
    pass
  if re.match(r"^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$", s):
    try:
      return True, float(s.replace(".", "").replace(",", "."))
    except ValueError:
      return False, None
  if re.match(r"^[+-]?\d+,\d+$", s):
    try:
      return True, float(s.replace(",", "."))
    except ValueError:
      return False, None
  return False, None


def convert_numeric_columns_for_export(df):
  df_out = df.copy()
  skip_keywords = [
      "lot",
      "batch",
      "kode",
      "keterangan",
      "tanggal",
      "tgl",
      "satuan",
      "unit",
      "benang",
      "item",
      "produk",
      "bobin",
      "status",
      "crosscheck",
  ]
  for col in df_out.columns:
    col_lower = str(col).strip().lower()
    if col_lower in ["no", "no."]:
      continue
    if pd.api.types.is_numeric_dtype(df_out[col]):
      if not df_out[col].notna().any():
        df_out[col] = ""
      continue
    non_empty_mask = (
        df_out[col].notna()
        & (df_out[col].astype(str).str.strip() != "")
        & (df_out[col].astype(str).str.strip().str.lower() != "nan")
    )
    if not non_empty_mask.any():
      df_out[col] = ""
      continue
    if (
        col_lower.startswith("no ")
        or col_lower.startswith("no.")
        or any(k in col_lower for k in skip_keywords)
    ):
      continue
    parsed_vals = []
    all_numeric = True
    for v in df_out[col]:
      ok, num = try_parse_strict_number(v)
      if not ok:
        all_numeric = False
        break
      parsed_vals.append(num)
    if all_numeric:
      df_out[col] = pd.Series(parsed_vals, index=df_out.index, dtype="float64")
  return df_out


def make_unique_columns(cols):
  seen = {}
  result = []
  for idx, c in enumerate(cols):
    name = clean_text(str(c))
    if not name or name.lower() == "nan":
      name = f"Kolom_{idx + 1}"
    if name in seen:
      seen[name] += 1
      result.append(f"{name}_{seen[name]}")
    else:
      seen[name] = 0
      result.append(name)
  return result


def find_col_ci(df, candidates):
  lower_map = {str(c).strip().lower(): c for c in df.columns}
  for cand in candidates:
    if cand.lower() in lower_map:
      return lower_map[cand.lower()]
  return None


def style_standard_sheet(ws, df, add_total_row=True):
  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(color="FFFFFF", bold=True)
  total_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  total_font = Font(color="FFFFFF", bold=True)
  thick_side = Side(style="medium", color="000000")
  thick_border = Border(
      left=thick_side, right=thick_side, top=thick_side, bottom=thick_side
  )

  max_col = len(df.columns)
  data_end_row = len(df) + 1

  numeric_col_indices = set()
  no_col_indices = set()
  for idx, col_name in enumerate(df.columns, 1):
    col_lower = str(col_name).strip().lower()
    if col_lower in ["no", "no."]:
      no_col_indices.add(idx)
    elif pd.api.types.is_numeric_dtype(df[col_name]):
      if df[col_name].notna().any():
        numeric_col_indices.add(idx)

  for col_idx in range(1, max_col + 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border

  for r in range(2, data_end_row + 1):
    for col_idx in range(1, max_col + 1):
      cell = ws.cell(row=r, column=col_idx)
      cell.border = thick_border
      if col_idx in numeric_col_indices:
        cell.number_format = "#,##0.00"
        cell.alignment = Alignment(horizontal="right", vertical="center")
      elif col_idx in no_col_indices:
        cell.number_format = "0"
        cell.alignment = Alignment(horizontal="center", vertical="center")

  if max_col > 0:
    last_col_letter = get_column_letter(max_col)
    filter_end_row = data_end_row if len(df) > 0 else 1
    ws.auto_filter.ref = f"A1:{last_col_letter}{filter_end_row}"

  if add_total_row and len(df) > 0 and len(numeric_col_indices) > 0:
    total_row = data_end_row + 1
    for col_idx in range(1, max_col + 1):
      cell = ws.cell(row=total_row, column=col_idx)
      cell.fill = total_fill
      cell.font = total_font
      cell.border = thick_border
      if col_idx == 1 and col_idx not in numeric_col_indices:
        cell.value = "TOTAL"
        cell.alignment = Alignment(horizontal="center", vertical="center")
      elif col_idx in numeric_col_indices:
        col_letter = get_column_letter(col_idx)
        cell.value = f"=SUBTOTAL(109,{col_letter}2:{col_letter}{data_end_row})"
        cell.number_format = "#,##0.00"
        cell.alignment = Alignment(horizontal="right", vertical="center")

  for col_idx in range(1, max_col + 1):
    col_letter = get_column_letter(col_idx)
    max_len = 0
    for r in range(1, ws.max_row + 1):
      val = ws.cell(row=r, column=col_idx).value
      if val is not None:
        val_str = str(val)
        if val_str.startswith("="):
          val_len = 16
        elif col_idx in numeric_col_indices and isinstance(val, (int, float)):
          val_len = len(f"{val:,.2f}")
        else:
          val_len = len(val_str)
        if val_len > max_len:
          max_len = val_len
    ws.column_dimensions[col_letter].width = max(max_len + 6, 14)


def style_laporan_gu_sheet(ws, df_gu_memphis, df_gu_mb):
  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(color="FFFFFF", bold=True)
  total_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  total_font = Font(color="FFFFFF", bold=True)
  thick_side = Side(style="medium", color="000000")
  thick_border = Border(
      left=thick_side, right=thick_side, top=thick_side, bottom=thick_side
  )

  ws.merge_cells("A1:D1")
  ws["A1"] = "DATA REKAP GU DARI MEMPHIS"
  for c in range(1, 5):
    cell = ws.cell(row=1, column=c)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border

  left_cols = list(df_gu_memphis.columns)
  for idx, col_name in enumerate(left_cols, 1):
    cell = ws.cell(row=2, column=idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border

  left_num_cols = {
      idx
      for idx, col_name in enumerate(left_cols, 1)
      if pd.api.types.is_numeric_dtype(df_gu_memphis[col_name])
      and df_gu_memphis[col_name].notna().any()
  }

  if len(df_gu_memphis) > 0:
    for r_idx, row in enumerate(df_gu_memphis.itertuples(index=False), 3):
      for c_idx, val in enumerate(row, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.border = thick_border
        if c_idx in left_num_cols:
          cell.number_format = "#,##0.00"
          cell.alignment = Alignment(horizontal="right", vertical="center")
  else:
    for c_idx in range(1, 5):
      ws.cell(row=3, column=c_idx).border = thick_border

  left_data_end = 2 + len(df_gu_memphis) if len(df_gu_memphis) > 0 else 3
  left_total_row = left_data_end + 1
  for c_idx in range(1, 5):
    cell = ws.cell(row=left_total_row, column=c_idx)
    cell.fill = total_fill
    cell.font = total_font
    cell.border = thick_border
    if c_idx == 1:
      cell.value = "TOTAL"
      cell.alignment = Alignment(horizontal="center", vertical="center")
    elif c_idx in left_num_cols:
      col_letter = get_column_letter(c_idx)
      cell.value = f"=SUBTOTAL(109,{col_letter}3:{col_letter}{left_data_end})"
      cell.number_format = "#,##0.00"
      cell.alignment = Alignment(horizontal="right", vertical="center")

  tab_left = Table(displayName="TabelGUMemphis", ref=f"A2:D{left_data_end}")
  ws.add_table(tab_left)

  ws.merge_cells("F1:K1")
  ws["F1"] = "DATA REKAP GU DARI MB (ERP)"
  for c in range(6, 12):
    cell = ws.cell(row=1, column=c)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border

  right_cols = list(df_gu_mb.columns)
  for idx, col_name in enumerate(right_cols, 6):
    cell = ws.cell(row=2, column=idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border

  right_num_cols = {
      idx
      for idx, col_name in enumerate(right_cols, 6)
      if pd.api.types.is_numeric_dtype(df_gu_mb[col_name])
      and df_gu_mb[col_name].notna().any()
  }

  if len(df_gu_mb) > 0:
    for r_idx, row in enumerate(df_gu_mb.itertuples(index=False), 3):
      for offset, val in enumerate(row):
        c_idx = 6 + offset
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.border = thick_border
        if c_idx in right_num_cols:
          cell.number_format = "#,##0.00"
          cell.alignment = Alignment(horizontal="right", vertical="center")
  else:
    for c_idx in range(6, 12):
      ws.cell(row=3, column=c_idx).border = thick_border

  right_data_end = 2 + len(df_gu_mb) if len(df_gu_mb) > 0 else 3
  right_total_row = right_data_end + 1
  for c_idx in range(6, 12):
    cell = ws.cell(row=right_total_row, column=c_idx)
    cell.fill = total_fill
    cell.font = total_font
    cell.border = thick_border
    if c_idx == 6:
      cell.value = "TOTAL"
      cell.alignment = Alignment(horizontal="center", vertical="center")
    elif c_idx in right_num_cols:
      col_letter = get_column_letter(c_idx)
      cell.value = f"=SUBTOTAL(109,{col_letter}3:{col_letter}{right_data_end})"
      cell.number_format = "#,##0.00"
      cell.alignment = Alignment(horizontal="right", vertical="center")

  tab_right = Table(displayName="TabelGUMB", ref=f"F2:K{right_data_end}")
  ws.add_table(tab_right)

  e2 = ws.cell(row=2, column=5, value="Selisih Total Memphis - Total ERP")
  e2.fill = header_fill
  e2.font = header_font
  e2.alignment = Alignment(horizontal="center", vertical="center")
  e2.border = thick_border

  e3 = ws.cell(
      row=3, column=5, value=f"=D{left_total_row}-H{right_total_row}"
  )
  e3.number_format = "#,##0.00"
  e3.font = Font(bold=True)
  e3.alignment = Alignment(horizontal="right", vertical="center")
  e3.border = thick_border

  for col_idx in range(1, 12):
    col_letter = get_column_letter(col_idx)
    max_len = 0
    for r in range(2, ws.max_row + 1):
      val = ws.cell(row=r, column=col_idx).value
      if val is not None:
        val_str = str(val)
        if val_str.startswith("="):
          val_len = 16
        elif isinstance(val, (int, float)):
          val_len = len(f"{val:,.2f}")
        else:
          val_len = len(val_str)
        if val_len > max_len:
          max_len = val_len
    ws.column_dimensions[col_letter].width = max(max_len + 6, 15)


if "processed" not in st.session_state:
  st.session_state.processed = False
  st.session_state.final_report = None
  st.session_state.df_ihp = None
  st.session_state.df_memphis = None
  st.session_state.df_gu = None
  st.session_state.df_heatset = None
  st.session_state.processed_data = None


st.markdown(
    """
    <div class="main-header">
        <h1>✨ DATA CONTROL IHP</h1>
        <p>Data Processing, Validation & Automated Reconciliation System</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.sidebar.markdown("### ⚡ DCIHP")
menu = st.sidebar.selectbox("", ["Pengelolaan Data IHP"])

if menu == "Pengelolaan Data IHP":
  st.markdown("### 🔮 Unggah Dokumen Master")
  st.markdown(
      "Silakan unggah file master **IHP**, **Memphis**, **Benang**, dan **MB"
      " (Opsional)** pada area di bawah ini:"
  )

  col1, col2, col3, col4 = st.columns(4)

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

  with col4:
    uploaded_mb = st.file_uploader(
        "4️⃣ Data Master MB - Opsional (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="mb",
    )

  st.markdown("<br>", unsafe_allow_html=True)

  if st.button("⚡ PROSES ANALISIS & REKONSILIASI", type="primary"):
    if uploaded_ihp and uploaded_memphis and uploaded_benang:
      try:
        with st.spinner(
            "🔮 AI sedang membersihkan, memetakan, dan memproses data..."
        ):
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

          df_ihp = clean_dataframe(df_ihp)
          df_memphis = clean_dataframe(df_memphis)
          df_benang = clean_dataframe(df_benang)

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

          if "Penerimaan" in df_gu.columns:
            penerimaan_gr_gu = df_gu["Penerimaan"] * 1000
            idx_pen_gu = df_gu.columns.get_loc("Penerimaan")
            if "Penerimaan (GR)" in df_gu.columns:
              df_gu = df_gu.drop(columns=["Penerimaan (GR)"])
            df_gu.insert(idx_pen_gu + 1, "Penerimaan (GR)", penerimaan_gr_gu)
          else:
            df_gu["Penerimaan (GR)"] = 0

          mapping_memphis_to_ihp = {}
          if len(df_benang.columns) >= 2:
            for _, row in df_benang.iterrows():
              m_name = str(row.iloc[0]).strip()
              i_name = str(row.iloc[1]).strip()
              if m_name and m_name != "nan" and i_name and i_name != "nan":
                mapping_memphis_to_ihp[m_name] = i_name

          nama_to_kode_item_map = {}
          if "Nama Item" in df_ihp.columns and "Kode Item" in df_ihp.columns:
            for _, r_ihp in df_ihp[["Nama Item", "Kode Item"]].dropna().iterrows():
              ni_val = str(r_ihp["Nama Item"]).strip()
              ki_val = str(r_ihp["Kode Item"]).strip()
              if ni_val and ni_val != "nan" and ki_val and ki_val != "nan":
                if ni_val not in nama_to_kode_item_map:
                  nama_to_kode_item_map[ni_val] = ki_val

          if (
              "Nama Barang Jadi" in df_ihp.columns
              and "Kode Barang Jadi" in df_ihp.columns
          ):
            for _, r_ihp in (
                df_ihp[["Nama Barang Jadi", "Kode Barang Jadi"]]
                .dropna()
                .iterrows()
            ):
              nbj_val = str(r_ihp["Nama Barang Jadi"]).strip()
              kbj_val = str(r_ihp["Kode Barang Jadi"]).strip()
              if nbj_val and nbj_val != "nan" and kbj_val and kbj_val != "nan":
                if nbj_val not in nama_to_kode_item_map:
                  nama_to_kode_item_map[nbj_val] = kbj_val

          df_heatset = pd.DataFrame(columns=df_ihp.columns)
          if "Kode Barang Jadi" in df_ihp.columns:
            mask_heatset = (
                df_ihp["Kode Barang Jadi"]
                .astype(str)
                .str.contains("TWP.050.8", case=False, na=False)
            )
            df_heatset = df_ihp[mask_heatset].copy()
            df_ihp = df_ihp[~mask_heatset].copy()

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

          no_col_ihp = find_col_ci(df_ihp, ["No", "No."])
          kode_group_col = find_col_ci(
              df_ihp, ["Kode", "Kode Pemakaian", "No Bukti"]
          )
          if no_col_ihp and kode_group_col and not df_ihp.empty:
            df_ihp[no_col_ihp] = pd.factorize(df_ihp[kode_group_col])[0] + 1

          no_col_hs = find_col_ci(df_heatset, ["No", "No."])
          kode_group_col_hs = find_col_ci(
              df_heatset, ["Kode", "Kode Pemakaian", "No Bukti"]
          )
          if no_col_hs and kode_group_col_hs and not df_heatset.empty:
            df_heatset[no_col_hs] = (
                pd.factorize(df_heatset[kode_group_col_hs])[0] + 1
            )

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
          detail_kode_rows = []
          detail_row_counter = 2
          start_row_map = {}

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

            if len(kodes) > 1:
              kode_summary = f"{kodes[0]} ... (+{len(kodes)-1} lainnya)"
            else:
              kode_summary = str(kodes[0])

            added_any = False
            start_idx = detail_row_counter
            for k in kodes:
              if k != "-":
                detail_kode_rows.append({
                    "Nama Benang": b_name,
                    "Kode Pemakaian IHP": k,
                })
                detail_row_counter += 1
                added_any = True

            if added_any:
              start_row_map[b_name] = start_idx

            final_rows.append({
                "Nama Benang": b_name,
                "Total Cones Memphis": total_cones,
                "Total Qty Memphis": total_qty_m,
                "Total QTY IHP": total_qty_ihp,
                "Selisih IHP - Memphis": selisih,
                "Satuan": "GR",
                "Kode Pemakaian IHP": kode_summary,
                "Keterangan": keterangan,
            })

          final_report = pd.DataFrame(final_rows)
          df_detail_kode = (
              pd.DataFrame(detail_kode_rows)
              if detail_kode_rows
              else pd.DataFrame(columns=["Nama Benang", "Kode Pemakaian IHP"])
          )

          col_benang_gu = None
          for col in ["Benang", "Nama Benang", "Item", "Nama Item"]:
            if col in df_gu.columns:
              col_benang_gu = col
              break
          if not col_benang_gu and len(df_gu.columns) > 0:
            col_benang_gu = df_gu.columns[0]

          def lookup_kode_gu(benang_memphis):
            b_clean = str(benang_memphis).strip()
            ihp_name = mapping_memphis_to_ihp.get(b_clean, b_clean)
            kode_found = nama_to_kode_item_map.get(ihp_name, "")
            if not kode_found:
              kode_found = nama_to_kode_item_map.get(b_clean, "")
            return kode_found

          if not df_gu.empty and col_benang_gu:
            cols_gu_to_sum = ["Penerimaan (GR)"]
            if "C. Penerimaan" in df_gu.columns:
              cols_gu_to_sum.insert(0, "C. Penerimaan")

            df_laporan_gu = df_gu.groupby(col_benang_gu, as_index=False)[
                cols_gu_to_sum
            ].sum()

            ordered_cols = [col_benang_gu]
            if "C. Penerimaan" in df_laporan_gu.columns:
              ordered_cols.append("C. Penerimaan")
            if "Penerimaan (GR)" in df_laporan_gu.columns:
              ordered_cols.append("Penerimaan (GR)")
            df_laporan_gu = df_laporan_gu[ordered_cols]
            df_laporan_gu.insert(
                0,
                "kode",
                df_laporan_gu[col_benang_gu].map(lookup_kode_gu),
            )
            if col_benang_gu != "Benang":
              df_laporan_gu = df_laporan_gu.rename(
                  columns={col_benang_gu: "Benang"}
              )
          else:
            df_laporan_gu = pd.DataFrame(
                columns=["kode", "Benang", "C. Penerimaan", "Penerimaan (GR)"]
            )

          col_kbj = (
              "Kode Barang Jadi"
              if "Kode Barang Jadi" in df_heatset.columns
              else None
          )
          col_nbj = (
              "Nama Barang Jadi"
              if "Nama Barang Jadi" in df_heatset.columns
              else None
          )
          col_ki = (
              "Kode Item"
              if "Kode Item" in df_heatset.columns
              else ("Kode" if "Kode" in df_heatset.columns else None)
          )
          col_ni = (
              "Nama Item"
              if "Nama Item" in df_heatset.columns
              else ("Item" if "Item" in df_heatset.columns else None)
          )

          if col_kbj and col_nbj and col_ki and col_ni and not df_heatset.empty:
            group_cols = [col_kbj, col_nbj, col_ki, col_ni]
            df_heatset["Qty Target Standar (GR)"] = pd.to_numeric(
                df_heatset["Qty Target Standar (GR)"], errors="coerce"
            ).fillna(0)
            df_heatset["Qty BB Standar (GR)"] = pd.to_numeric(
                df_heatset["Qty BB Standar (GR)"], errors="coerce"
            ).fillna(0)

            df_laporan_heatset = df_heatset.groupby(
                group_cols, as_index=False
            ).agg({
                "Qty Target Standar (GR)": "sum",
                "Qty BB Standar (GR)": "sum",
            })

            df_laporan_heatset = df_laporan_heatset.rename(
                columns={
                    col_kbj: "Kode Barang Jadi",
                    col_nbj: "Nama Barang Jadi",
                    col_ki: "Kode Item",
                    col_ni: "Nama Item",
                }
            )

            df_laporan_heatset = df_laporan_heatset[[
                "Kode Barang Jadi",
                "Nama Barang Jadi",
                "Qty Target Standar (GR)",
                "Kode Item",
                "Nama Item",
                "Qty BB Standar (GR)",
            ]]
          else:
            df_laporan_heatset = pd.DataFrame(
                columns=[
                    "Kode Barang Jadi",
                    "Nama Barang Jadi",
                    "Qty Target Standar (GR)",
                    "Kode Item",
                    "Nama Item",
                    "Qty BB Standar (GR)",
                ]
            )

          if (
              "Qty BB Standar (GR)" in df_laporan_heatset.columns
              and "Qty Target Standar (GR)" in df_laporan_heatset.columns
          ):
            df_laporan_heatset["Selisih (GR)"] = (
                df_laporan_heatset["Qty BB Standar (GR)"]
                - df_laporan_heatset["Qty Target Standar (GR)"]
            )
          else:
            df_laporan_heatset["Selisih (GR)"] = 0

          ihp_kbj = find_col_ci(df_ihp, ["Kode Barang Jadi", "Kode Bahan Jadi"])
          ihp_nbj = find_col_ci(df_ihp, ["Nama Barang Jadi", "Nama Bahan Jadi"])
          ihp_ki = find_col_ci(df_ihp, ["Kode Item", "Kode"])
          ihp_ni = find_col_ci(df_ihp, ["Nama Item", "Item"])

          if ihp_kbj and ihp_nbj and ihp_ki and ihp_ni and not df_ihp.empty:
            df_ihp_tmp = df_ihp.copy()
            df_ihp_tmp["Qty Target Standar (GR)"] = pd.to_numeric(
                df_ihp_tmp["Qty Target Standar (GR)"], errors="coerce"
            ).fillna(0)
            df_laporan_ihp = df_ihp_tmp.groupby(
                [ihp_kbj, ihp_nbj, ihp_ki, ihp_ni], as_index=False, sort=False
            )["Qty Target Standar (GR)"].sum()
            df_laporan_ihp = df_laporan_ihp.rename(
                columns={
                    ihp_kbj: "KODE BAHAN JADI",
                    ihp_nbj: "NAMA BAHAN JADI",
                    ihp_ki: "KODE ITEM",
                    ihp_ni: "NAMA ITEM",
                    "Qty Target Standar (GR)": "Total QTY IHP (GR)",
                }
            )
            df_laporan_ihp["Satuan"] = "GR"
            df_laporan_ihp = df_laporan_ihp[[
                "KODE BAHAN JADI",
                "NAMA BAHAN JADI",
                "KODE ITEM",
                "NAMA ITEM",
                "Total QTY IHP (GR)",
                "Satuan",
            ]]
          else:
            df_laporan_ihp = pd.DataFrame(
                columns=[
                    "KODE BAHAN JADI",
                    "NAMA BAHAN JADI",
                    "KODE ITEM",
                    "NAMA ITEM",
                    "Total QTY IHP (GR)",
                    "Satuan",
                ]
            )

          df_mb = None
          mask_mb_filter = None
          df_laporan_gu_mb = pd.DataFrame(
              columns=[
                  "Kode Item",
                  "Nama Item",
                  "Qty BB Standar (GR)",
                  "Kode Barang Jadi",
                  "Nama Barang Jadi",
                  "Qty Target Standar (GR)",
              ]
          )

          if uploaded_mb is not None:
            if uploaded_mb.name.endswith(".csv"):
              df_mb_raw = pd.read_csv(uploaded_mb, header=None, dtype=str)
            else:
              df_mb_raw = pd.read_excel(uploaded_mb, header=None, dtype=str)

            if len(df_mb_raw) > 8:
              df_mb_raw = df_mb_raw.iloc[4:-4].copy()
            elif len(df_mb_raw) > 4:
              df_mb_raw = df_mb_raw.iloc[4:].copy()

            for col in df_mb_raw.columns:
              df_mb_raw[col] = df_mb_raw[col].apply(
                  lambda x: clean_text(x) if isinstance(x, str) else x
              )

            df_mb_raw = df_mb_raw.replace(
                to_replace=r"^\s*$", value=pd.NA, regex=True
            )
            df_mb_raw = df_mb_raw.dropna(how="all").reset_index(drop=True)

            if len(df_mb_raw) > 1:
              df_mb_raw.columns = make_unique_columns(df_mb_raw.iloc[0].tolist())
              df_mb = df_mb_raw.iloc[1:].reset_index(drop=True)
              df_mb = clean_dataframe(df_mb)
              df_mb = df_mb.replace(
                  to_replace=r"^\s*$", value=pd.NA, regex=True
              )
              df_mb = df_mb.dropna(how="all").reset_index(drop=True)
              df_mb = df_mb.fillna("")

              mb_target_qty_col = find_col_ci(df_mb, ["Target Qty"])
              mb_target_unit_col = find_col_ci(df_mb, ["Target Unit"])
              mb_qty_col = find_col_ci(df_mb, ["Qty"])
              mb_unit_col = find_col_ci(df_mb, ["Unit"])

              if mb_target_qty_col:
                df_mb[mb_target_qty_col] = df_mb[mb_target_qty_col].apply(
                    parse_numeric
                )
              if mb_qty_col:
                df_mb[mb_qty_col] = df_mb[mb_qty_col].apply(parse_numeric)

              if mb_target_unit_col and mb_target_qty_col:
                mb_target_gr = df_mb.apply(
                    lambda row: (
                        row[mb_target_qty_col]
                        if str(row[mb_target_unit_col]).strip().upper() == "GR"
                        else row[mb_target_qty_col] * 1000
                    ),
                    axis=1,
                )
                idx_tu = df_mb.columns.get_loc(mb_target_unit_col)
                if "Qty Target Standar (GR)" in df_mb.columns:
                  df_mb = df_mb.drop(columns=["Qty Target Standar (GR)"])
                df_mb.insert(idx_tu + 1, "Qty Target Standar (GR)", mb_target_gr)
              else:
                df_mb["Qty Target Standar (GR)"] = 0.0

              if mb_unit_col and mb_qty_col:
                mb_bb_gr = df_mb.apply(
                    lambda row: (
                        row[mb_qty_col]
                        if str(row[mb_unit_col]).strip().upper() == "GR"
                        else row[mb_qty_col] * 1000
                    ),
                    axis=1,
                )
                idx_q = df_mb.columns.get_loc(mb_qty_col)
                if "Qty BB Standar (GR)" in df_mb.columns:
                  df_mb = df_mb.drop(columns=["Qty BB Standar (GR)"])
                df_mb.insert(idx_q + 1, "Qty BB Standar (GR)", mb_bb_gr)
              else:
                df_mb["Qty BB Standar (GR)"] = 0.0

              mask_mb_filter = pd.Series(False, index=df_mb.index)

              col_nama_produk = find_col_ci(
                  df_mb, ["Nama Produk", "Nama Barang Jadi"]
              )
              if col_nama_produk:
                mask_mb_filter = mask_mb_filter | df_mb[col_nama_produk].astype(
                    str
                ).str.contains(r"HTC|BDP", case=False, na=False)

              ket_cols = [
                  c for c in df_mb.columns if "keterangan" in str(c).lower()
              ]
              for kc in ket_cols:
                mask_mb_filter = mask_mb_filter | df_mb[kc].astype(
                    str
                ).str.contains(r"GU|GULUNG ULANG", case=False, na=False)

              col_mb_kbj = find_col_ci(
                  df_mb, ["Kode Barang Jadi", "Kode Produk"]
              )
              if col_mb_kbj:
                mask_mb_filter = mask_mb_filter | df_mb[col_mb_kbj].astype(
                    str
                ).str.contains(r"TWP\.100|TWP\.999", case=False, na=False)

              df_mb_filtered = df_mb[mask_mb_filter].copy()

              mb_ki = find_col_ci(df_mb_filtered, ["Kode Item"])
              mb_ni = find_col_ci(df_mb_filtered, ["Nama Item", "Item"])
              mb_kbj = find_col_ci(
                  df_mb_filtered, ["Kode Barang Jadi", "Kode Produk"]
              )
              mb_nbj = find_col_ci(
                  df_mb_filtered, ["Nama Barang Jadi", "Nama Produk"]
              )

              if (
                  mb_ki
                  and mb_ni
                  and mb_kbj
                  and mb_nbj
                  and not df_mb_filtered.empty
              ):
                df_mb_filtered["Qty BB Standar (GR)"] = pd.to_numeric(
                    df_mb_filtered["Qty BB Standar (GR)"], errors="coerce"
                ).fillna(0)
                df_mb_filtered["Qty Target Standar (GR)"] = pd.to_numeric(
                    df_mb_filtered["Qty Target Standar (GR)"], errors="coerce"
                ).fillna(0)

                df_laporan_gu_mb = df_mb_filtered.groupby(
                    [mb_ki, mb_ni, mb_kbj, mb_nbj], as_index=False
                ).agg({
                    "Qty BB Standar (GR)": "sum",
                    "Qty Target Standar (GR)": "sum",
                })
                df_laporan_gu_mb = df_laporan_gu_mb.rename(
                    columns={
                        mb_ki: "Kode Item",
                        mb_ni: "Nama Item",
                        mb_kbj: "Kode Barang Jadi",
                        mb_nbj: "Nama Barang Jadi",
                    }
                )
                df_laporan_gu_mb = df_laporan_gu_mb[[
                    "Kode Item",
                    "Nama Item",
                    "Qty BB Standar (GR)",
                    "Kode Barang Jadi",
                    "Nama Barang Jadi",
                    "Qty Target Standar (GR)",
                ]]

          final_report_exp = convert_numeric_columns_for_export(final_report)
          df_detail_kode_exp = convert_numeric_columns_for_export(
              df_detail_kode
          )
          df_gu_exp = convert_numeric_columns_for_export(df_gu)
          df_laporan_gu_exp = convert_numeric_columns_for_export(df_laporan_gu)
          df_laporan_gu_mb_exp = convert_numeric_columns_for_export(
              df_laporan_gu_mb
          )
          df_heatset_exp = convert_numeric_columns_for_export(df_heatset)
          df_laporan_heatset_exp = convert_numeric_columns_for_export(
              df_laporan_heatset
          )
          df_laporan_ihp_exp = convert_numeric_columns_for_export(
              df_laporan_ihp
          )
          df_ihp_exp = convert_numeric_columns_for_export(df_ihp)
          df_memphis_exp = convert_numeric_columns_for_export(df_memphis)
          df_mb_exp = (
              convert_numeric_columns_for_export(df_mb)
              if df_mb is not None
              else None
          )

          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_report_exp.to_excel(
                writer, sheet_name="Hasil Olah Data", index=False
            )
            df_detail_kode_exp.to_excel(
                writer, sheet_name="Detail Kode IHP", index=False
            )
            df_gu_exp.to_excel(writer, sheet_name="Data GU", index=False)

            workbook = writer.book
            ws_lap_gu = workbook.create_sheet(title="Laporan GU")
            style_laporan_gu_sheet(
                ws_lap_gu, df_laporan_gu_exp, df_laporan_gu_mb_exp
            )

            df_heatset_exp.to_excel(
                writer, sheet_name="Data Heatset", index=False
            )
            df_laporan_heatset_exp.to_excel(
                writer, sheet_name="Laporan Heatset", index=False
            )
            df_laporan_ihp_exp.to_excel(
                writer, sheet_name="Laporan IHP", index=False
            )
            df_ihp_exp.to_excel(
                writer, sheet_name="Data Master IHP", index=False
            )
            df_memphis_exp.to_excel(
                writer, sheet_name="Data Memphis", index=False
            )

            if df_mb_exp is not None:
              df_mb_exp.to_excel(writer, sheet_name="Master MB", index=False)

            standard_sheets = [
                ("Hasil Olah Data", final_report_exp, True),
                ("Detail Kode IHP", df_detail_kode_exp, False),
                ("Data GU", df_gu_exp, True),
                ("Data Heatset", df_heatset_exp, True),
                ("Laporan Heatset", df_laporan_heatset_exp, True),
                ("Laporan IHP", df_laporan_ihp_exp, True),
                ("Data Master IHP", df_ihp_exp, True),
                ("Data Memphis", df_memphis_exp, True),
            ]
            if df_mb_exp is not None:
              standard_sheets.append(("Master MB", df_mb_exp, True))

            for s_name, s_df, add_tot in standard_sheets:
              style_standard_sheet(
                  workbook[s_name],
                  s_df,
                  add_total_row=add_tot,
              )

            if (
                df_mb_exp is not None
                and mask_mb_filter is not None
                and len(df_mb_exp) > 0
            ):
              ws_mb = workbook["Master MB"]
              for i, is_match in enumerate(mask_mb_filter):
                if not is_match:
                  ws_mb.row_dimensions[i + 2].hidden = True

            ws_hasil = workbook["Hasil Olah Data"]
            col_idx_nama_benang = None
            col_idx_kode = None

            for idx, cell in enumerate(ws_hasil[1], 1):
              if cell.value == "Nama Benang":
                col_idx_nama_benang = idx
              elif cell.value == "Kode Pemakaian IHP":
                col_idx_kode = idx
                if (
                    ws_hasil.column_dimensions[cell.column_letter].width
                    is None
                    or ws_hasil.column_dimensions[cell.column_letter].width < 35
                ):
                  ws_hasil.column_dimensions[cell.column_letter].width = 35

            if col_idx_nama_benang and col_idx_kode:
              hyperlink_font = Font(color="0563C1", underline="single")
              for row in range(2, len(final_report_exp) + 2):
                b_name_cell = ws_hasil.cell(
                    row=row, column=col_idx_nama_benang
                ).value
                kode_cell = ws_hasil.cell(row=row, column=col_idx_kode)

                if b_name_cell in start_row_map and "lainnya" in str(
                    kode_cell.value
                ):
                  target_row = start_row_map[b_name_cell]
                  kode_cell.hyperlink = f"#'Detail Kode IHP'!A{target_row}"
                  kode_cell.font = hyperlink_font

          processed_data = output.getvalue()

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
