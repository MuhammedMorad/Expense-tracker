import pandas as pd
import re
import os


class InvoiceParser:

    def __init__(self, file_path='data/uploaded_invoices.xlsx'):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.file_path = os.path.join(base_dir, file_path)
        self._ensure_file()

    def _ensure_file(self):
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)
        if not os.path.exists(self.file_path):
            pd.DataFrame(columns=['ID']).to_excel(self.file_path, index=False)

    def _read(self):
        try:
            return pd.read_excel(self.file_path)
        except:
            return pd.DataFrame(columns=['ID'])

    def _write(self, df):
        df.to_excel(self.file_path, index=False)

    # ── Parse uploaded Excel file ─────────────────────────────────────────

    def parse_uploaded_invoice(self, file_path):
        """
        Parse invoice Excel files that may contain multiple invoice blocks
        inside the same sheet.
        """
        try:
            # Pick the right engine
            if file_path.lower().endswith('.xls'):
                try:
                    import xlrd
                    engine = 'xlrd'
                except ImportError:
                    print("Parse Invoice Error: xlrd not installed")
                    return []
            else:
                engine = 'openpyxl'

            date_pattern = r'\d{4}-\d{2}-\d{2}'
            all_results = []

            try:
                xls = pd.ExcelFile(file_path, engine=engine)
            except OSError as e:
                # Fallback: sometimes .xls files are named .xlsx
                if engine == 'openpyxl':
                    try:
                        xls = pd.ExcelFile(file_path, engine='xlrd')
                    except:
                        print("Parse Invoice Error:", e)
                        return []
                else:
                    print("Parse Invoice Error:", e)
                    return []
            except Exception as e:
                print("Parse Invoice Error:", e)
                return []

            for sheet in xls.sheet_names:
                try:
                    df_raw = xls.parse(sheet, header=None, dtype=str).fillna("")
                except:
                    continue

                if df_raw.empty:
                    continue

                # ── locate all invoice headers (row containing الصنف) ──
                header_rows = []
                for i, row in df_raw.iterrows():
                    row_vals = [str(v).strip() for v in row.tolist()]
                    if "الصنف" in row_vals:
                        header_rows.append(i)

                if not header_rows:
                    continue

                header_rows.append(df_raw.index[-1] + 1)

                # ── process each invoice block ──
                for blk_i, hdr_i in enumerate(header_rows[:-1]):
                    next_hdr = header_rows[blk_i + 1]

                    # find invoice date above the table
                    invoice_date = "Unknown"
                    for r in range(max(0, hdr_i - 10), hdr_i):
                        row_text = " ".join(df_raw.loc[r].astype(str).tolist())
                        dates = re.findall(date_pattern, row_text)
                        if dates:
                            invoice_date = dates[-1]

                    # detect expenses / net
                    expenses = 0
                    net = 0
                    for r in range(hdr_i, next_hdr):
                        row_text = " ".join(df_raw.loc[r].astype(str).tolist())
                        if "تنزيل" in row_text or "منصرف" in row_text:
                            for cell in df_raw.loc[r]:
                                try:
                                    v = float(str(cell).replace(",", "").strip())
                                    if v > 0: expenses = v
                                except: pass
                        if "صافي" in row_text or "صافى" in row_text:
                            for cell in df_raw.loc[r]:
                                try:
                                    v = float(str(cell).replace(",", "").strip())
                                    if v > 0: net = v
                                except: pass

                    # build column names
                    header_vals = [str(v).strip() for v in df_raw.loc[hdr_i].tolist()]
                    col_names = [v if v not in ("", "nan") else f"_c{i}" for i, v in enumerate(header_vals)]

                    data_rows = df_raw.loc[hdr_i + 1: next_hdr - 1].copy()
                    if data_rows.empty:
                        continue

                    data_rows.columns = col_names[:len(data_rows.columns)]
                    if "الصنف" not in data_rows.columns:
                        continue

                    # Filter out helper rows and empty items
                    exclude_keywords = ['تنزيل', 'صافي', 'صافى', 'اجمالي', 'اجمالى', 'المنصرف', 'المجموع', 'total']
                    data_rows = data_rows[data_rows["الصنف"].astype(str).str.strip() != ""]
                    
                    def num(val):
                        try:
                            return float(str(val).replace(",", "").strip())
                        except:
                            return 0

                    def get_val_by_keys(row_data, keys):
                        for k in keys:
                            for col in row_data.index:
                                if k in str(col):
                                    return num(row_data[col])
                        return 0

                    for _, r in data_rows.iterrows():
                        item = str(r.get("الصنف", "")).strip()
                        if not item or item.lower() == "nan" or any(kw in item for kw in exclude_keywords):
                            continue

                        all_results.append({
                            "التاريخ": invoice_date,
                            "الصنف": item,
                            "العدد": get_val_by_keys(r, ["عدد", "العدد"]),
                            "السعر": get_val_by_keys(r, ["سعر", "السعر"]),
                            "الكيلو": get_val_by_keys(r, ["الكيلو", "كيلو"]),
                            "الاجمالي": get_val_by_keys(r, ["اجمالي", "اجمالى", "إجمالي"]),
                            "تنزيل المنصرف": expenses,
                            "صافي الفاتورة": net
                        })

            return all_results

        except Exception as e:
            print("Parse Invoice Error:", e)
            import traceback
            traceback.print_exc()
            return []

    # ── CRUD for saved invoices ───────────────────────────────────────────

    def save_uploaded_invoice_data(self, data):
        if not data:
            return False
        df_old = self._read()
        df_new = pd.DataFrame(data)
        start_id = 1
        if not df_old.empty and 'ID' in df_old.columns:
            start_id = int(pd.to_numeric(df_old['ID'], errors='coerce').max()) + 1
        df_new['ID'] = range(start_id, start_id + len(df_new))
        df = pd.concat([df_old, df_new], ignore_index=True)
        self._write(df)
        return True

    def get_uploaded_invoices(self):
        return self._read().fillna('').to_dict('records')

    def update_invoice(self, record_id, data):
        df = self._read()
        idx = df.index[df['ID'] == int(record_id)]
        if idx.empty:
            return False
        for k, v in data.items():
            if k in df.columns:
                df.loc[idx, k] = v
        self._write(df)
        return True

    def delete_invoice(self, record_id):
        df = self._read()
        df = df[df['ID'] != int(record_id)]
        self._write(df)
        return True

    def delete_all_invoices(self):
        df = self._read()
        df = df.iloc[0:0]
        self._write(df)
