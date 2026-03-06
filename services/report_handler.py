import pandas as pd
import os


class ReportHandler:

    def __init__(self, file_path="data/uploaded_reports.xlsx"):

        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        self.file_path = os.path.join(base_dir, file_path)

        self.ensure_file()

    def ensure_file(self):

        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

        if not os.path.exists(self.file_path):
            pd.DataFrame(columns=["ID"]).to_excel(self.file_path, index=False)

    def get_all(self):

        df = pd.read_excel(self.file_path)

        return df.fillna("").to_dict("records")

    def save(self, data):

        if not data:
            return False

        df_old = pd.read_excel(self.file_path)

        df_new = pd.DataFrame(data)

        start_id = 1

        if not df_old.empty:
            start_id = int(pd.to_numeric(df_old["ID"], errors="coerce").max()) + 1

        df_new["ID"] = range(start_id, start_id + len(df_new))

        df = pd.concat([df_old, df_new], ignore_index=True)

        df.to_excel(self.file_path, index=False)

        return True

    def analytics(self):

        df = pd.read_excel(self.file_path)

        if "العمولة قبل إجمالي" not in df.columns:
            return {}

        df["العمولة قبل إجمالي"] = pd.to_numeric(
            df["العمولة قبل إجمالي"], errors="coerce"
        ).fillna(0)

        total = df["العمولة قبل إجمالي"].sum()

        top_items = {}

        if "الصنف" in df.columns:
            top_items = (
                df.groupby("الصنف")["العمولة قبل إجمالي"]
                .sum()
                .sort_values(ascending=False)
                .head(5)
                .to_dict()
            )

        return {
            "total_commission": total,
            "top_items": top_items,
            "records": len(df)
        }

    # ── Aliases used by app.py ─────────────────────────────────────────────

    def get_uploaded_data(self):
        return self.get_all()

    def save_uploaded_data(self, data):
        return self.save(data)

    def get_reports_analytics(self):
        return self.analytics()

    def update_report(self, record_id, data):
        df = pd.read_excel(self.file_path)
        idx = df.index[df['ID'] == int(record_id)]
        if idx.empty:
            return False
        for k, v in data.items():
            if k in df.columns:
                df.loc[idx, k] = v
        df.to_excel(self.file_path, index=False)
        return True

    def delete_report(self, record_id):
        df = pd.read_excel(self.file_path)
        df = df[df['ID'] != int(record_id)]
        df.to_excel(self.file_path, index=False)
        return True

    def delete_all_reports(self):
        df = pd.read_excel(self.file_path)
        df = df.iloc[0:0]
        df.to_excel(self.file_path, index=False)

    def parse_uploaded_report(self, file_path):
        """Parse an uploaded report Excel file and return list of dicts."""
        try:
            import re
            date_pattern = r'\d{4}-\d{2}-\d{2}'
            df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
            header_row_idx = -1
            report_period = ''
            for i, row in df_raw.iterrows():
                row_text = ' '.join(row.astype(str).tolist())
                if report_period == '':
                    dates = re.findall(date_pattern, row_text)
                    if len(dates) >= 2:
                        report_period = f"{dates[0]} to {dates[1]}"
                    elif len(dates) == 1:
                        report_period = dates[0]
                if 'الصنف' in row_text and header_row_idx == -1:
                    header_row_idx = i
            if header_row_idx == -1:
                return []
            df = pd.read_excel(file_path, header=header_row_idx, engine='openpyxl')
            df.columns = df.columns.astype(str).str.strip()
            if 'الصنف' in df.columns:
                df = df.dropna(subset=['الصنف'])
                df = df[~df['الصنف'].astype(str).str.contains(r'^nan$|^الصنف$', regex=True, na=False)]
            df['ReportPeriod'] = report_period or 'Unknown'
            return df.fillna('').to_dict('records')
        except Exception as e:
            print(f"Parse Report Error: {e}")
            return []
