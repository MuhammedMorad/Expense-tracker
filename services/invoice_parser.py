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
        Parse an invoice Excel file that may contain multiple invoice blocks
        in a single sheet. Each block starts with a header row containing 'الصنف'.
        """
        try:
            # Pick the right engine
            if file_path.lower().endswith('.xls'):
                try:
                    import xlrd  # noqa
                    engine = 'xlrd'
                except ImportError:
                    print("Parse Invoice Error: xlrd not installed. Run: pip install xlrd")
                    return []
            else:
                engine = 'openpyxl'

            date_pattern = r'\d{4}-\d{2}-\d{2}'
            all_results = []
            xls_file = pd.ExcelFile(file_path, engine=engine)

            for sheet_name in xls_file.sheet_names:
                df_raw = xls_file.parse(sheet_name, header=None, dtype=str)
                if df_raw.empty:
                    continue

                # Find ALL rows where 'الصنف' appears as a cell value
                header_row_indices = []
                for i, row in df_raw.iterrows():
                    if 'الصنف' in [str(v).strip() for v in row.tolist()]:
                        header_row_indices.append(i)

                if not header_row_indices:
                    continue

                header_row_indices.append(df_raw.index[-1] + 1)  # sentinel

                for blk_num, hdr_i in enumerate(header_row_indices[:-1]):
                    next_hdr_i = header_row_indices[blk_num + 1]

                    # Date: look in rows BEFORE this header
                    search_from = header_row_indices[blk_num - 1] if blk_num > 0 else df_raw.index[0]
                    invoice_date = 'Unknown'
                    for ri in range(search_from, hdr_i):
                        if ri not in df_raw.index:
                            continue
                        row_text = ' '.join(df_raw.loc[ri].astype(str).tolist())
                        dates = re.findall(date_pattern, row_text)
                        if dates:
                            invoice_date = dates[-1]

                    # تنزيل المنصرف / صافي الفاتورة: within this block
                    expenses_deduction = 0
                    net_invoice = 0
                    for ri in range(hdr_i, next_hdr_i):
                        if ri not in df_raw.index:
                            continue
                        row_text = ' '.join(df_raw.loc[ri].astype(str).tolist())
                        if 'تنزيل' in row_text or 'منصرف' in row_text:
                            for cell in df_raw.loc[ri]:
                                try:
                                    v = float(str(cell).strip().replace(',', ''))
                                    if v > 0:
                                        expenses_deduction = v
                                except: pass
                        if 'صافي' in row_text or 'صافى' in row_text:
                            for cell in df_raw.loc[ri]:
                                try:
                                    v = float(str(cell).strip().replace(',', ''))
                                    if v > 0:
                                        net_invoice = v
                                except: pass

                    # Build column map from header row
                    header_vals = [str(v).strip() for v in df_raw.loc[hdr_i].tolist()]
                    col_names = [v if v not in ('nan', '') else f'_col{ci}' for ci, v in enumerate(header_vals)]

                    if next_hdr_i - 1 < hdr_i + 1:
                        continue
                    data_rows = df_raw.loc[hdr_i + 1: next_hdr_i - 1].copy()
                    if data_rows.empty:
                        continue

                    data_rows.columns = col_names[:len(data_rows.columns)]
                    if 'الصنف' not in data_rows.columns:
                        continue

                    exclude = r'^الصنف$|^nan$|تنزيل|صافي|اجمالي|صافى|اجمالى|الاجمالي|المنصرف'
                    data_rows = data_rows[~data_rows['الصنف'].astype(str).str.strip().str.contains(exclude, na=False, regex=True)]
                    data_rows = data_rows[data_rows['الصنف'].astype(str).str.strip().replace('nan', '') != '']

                    def get_num(row_s, keys):
                        for k in keys:
                            for c in row_s.index:
                                if k in str(c):
                                    v = row_s[c]
                                    if pd.notna(v) and str(v).strip() not in ('', 'nan'):
                                        try:
                                            return float(str(v).replace(',', '').strip())
                                        except: pass
                        return 0

                    for _, r in data_rows.iterrows():
                        item = str(r.get('الصنف', '')).strip()
                        if not item or item.lower() == 'nan':
                            continue
                        all_results.append({
                            'التاريخ':        invoice_date,
                            'الصنف':          item,
                            'العدد':          get_num(r, ['عدد', 'العدد']),
                            'السعر':          get_num(r, ['سعر', 'السعر']),
                            'الكيلو':         get_num(r, ['الكيلو', 'كيلو']),
                            'الاجمالي':       get_num(r, ['اجمالى', 'الاجمالي', 'إجمالي', 'اجمالي']),
                            'تنزيل المنصرف': expenses_deduction,
                            'صافي الفاتورة': net_invoice,
                        })

            return all_results

        except Exception as e:
            print(f"Parse Invoice Error: {e}")
            import traceback; traceback.print_exc()
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