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
          detail_kode_rows = [] # TAMBAHAN OPSI B: Untuk menyimpan rincian kode

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

            # --- MODIFIKASI OPSI B (Meringkas Kode di Laporan Utama) ---
            if len(kodes) > 1:
                kode_summary = f"{kodes[0]} ... (+{len(kodes)-1} lainnya)"
            else:
                kode_summary = str(kodes[0])

            # Menyimpan detail kode ke list terpisah untuk Sheet baru
            for k in kodes:
                if k != "-":
                    detail_kode_rows.append({
                        "Nama Benang": b_name,
                        "Kode Pemakaian IHP": k
                    })
            # -----------------------------------------------------------

            final_rows.append({
                "Nama Benang": b_name,
                "Total Cones Memphis": total_cones,
                "Total Qty Memphis": total_qty_m,
                "Total QTY IHP": total_qty_ihp,
                "Selisih IHP - Memphis": selisih,
                "Satuan": "GR",
                "Kode Pemakaian IHP": kode_summary, # Menampilkan versi ringkas
                "Keterangan": keterangan,
            })

          final_report = pd.DataFrame(final_rows)
          df_detail_kode = pd.DataFrame(detail_kode_rows) if detail_kode_rows else pd.DataFrame(columns=["Nama Benang", "Kode Pemakaian IHP"])


          # ==========================================
          # 5B. REKAP LAPORAN GU & HEATSET
          # ==========================================
          # 1. Pembuatan Laporan GU (Group by Benang)
          col_benang_gu = None
          for col in ["Benang", "Nama Benang", "Item", "Nama Item"]:
              if col in df_gu.columns:
                  col_benang_gu = col
                  break
          if not col_benang_gu and len(df_gu.columns) > 0:
              col_benang_gu = df_gu.columns[0]
              
          if not df_gu.empty and col_benang_gu:
              df_laporan_gu = df_gu.groupby(col_benang_gu, as_index=False)["Penerimaan (GR)"].sum()
          else:
              df_laporan_gu = pd.DataFrame(columns=["Benang", "Penerimaan (GR)"])

          # 2. Pembuatan Laporan Heatset (2 Bagian: Barang Jadi & Bahan Baku)
          col_kbj = "Kode Barang Jadi" if "Kode Barang Jadi" in df_heatset.columns else None
          col_nbj = "Nama Barang Jadi" if "Nama Barang Jadi" in df_heatset.columns else None
          
          if col_kbj and col_nbj and not df_heatset.empty:
              df_heatset_bj = df_heatset.groupby([col_kbj, col_nbj], as_index=False)["Qty Target Standar (GR)"].sum()
          else:
              df_heatset_bj = pd.DataFrame(columns=["Kode Barang Jadi", "Nama Barang Jadi", "Qty Target Standar (GR)"])

          # Penentuan dinamis untuk kolom Kode dan Nama Bahan Baku
          col_ki = "Kode" if "Kode" in df_heatset.columns else ("Kode Item" if "Kode Item" in df_heatset.columns else "Kode")
          col_ni = "Nama Item" if "Nama Item" in df_heatset.columns else ("Item" if "Item" in df_heatset.columns else "Nama Item")
          
          if col_ki in df_heatset.columns and col_ni in df_heatset.columns and not df_heatset.empty:
              df_heatset_bb = df_heatset.groupby([col_ki, col_ni], as_index=False)["Qty BB Standar (GR)"].sum()
          else:
              df_heatset_bb = pd.DataFrame(columns=[col_ki, col_ni, "Qty BB Standar (GR)"])


          # ==========================================
          # 6. GENERATE EXCEL BUFFER
          # ==========================================
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Menyimpan Sheet Utama & Detail Baru
            final_report.to_excel(writer, sheet_name="Hasil Olah Data", index=False)
            df_detail_kode.to_excel(writer, sheet_name="Detail Kode IHP", index=False)
            
            # Menyimpan Sheet Laporan Rekapitulasi
            df_laporan_gu.to_excel(writer, sheet_name="Laporan GU", index=False)
            
            # Menyimpan Laporan Heatset (2 Tabel disatukan dalam 1 sheet, Startcol memberi jarak)
            df_heatset_bj.to_excel(writer, sheet_name="Laporan Heatset", index=False, startcol=0) # Kolom A, B, C
            df_heatset_bb.to_excel(writer, sheet_name="Laporan Heatset", index=False, startcol=4) # Kolom E, F, G (Kolom D kosong)

            # Menyimpan Data Master Original
            df_ihp.to_excel(writer, sheet_name="Data Master IHP", index=False)
            df_memphis.to_excel(writer, sheet_name="Data Memphis", index=False)
            df_gu.to_excel(writer, sheet_name="Data GU", index=False)
            df_heatset.to_excel(writer, sheet_name="Heatset", index=False)

            # Mengatur Lebar Kolom untuk Sheet 'Hasil Olah Data' (Menghapus Wrap Text)
            workbook = writer.book
            ws = workbook["Hasil Olah Data"]

            for col in ws.iter_cols(min_row=1, max_row=1, min_col=1, max_col=ws.max_column):
              for cell in col:
                if cell.value == "Kode Pemakaian IHP":
                  col_letter = cell.column_letter
                  ws.column_dimensions[col_letter].width = 35 # Dilebarkan sedikit, teks sisa tersembunyi dengan rapi di satu baris

          processed_data = output.getvalue()
