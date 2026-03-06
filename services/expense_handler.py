import pandas as pd
import os
from datetime import datetime


class ExpenseHandler:

    REQUIRED_COLUMNS = [
        'ID', 'Amount', 'Description', 'ExpenseType', 'Date', 'Time', 'ExpenseYear'
    ]

    def __init__(self, file_path='data/expenses.xlsx'):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.file_path = os.path.join(base_dir, file_path)
        self.ensure_file_exists()

    def ensure_file_exists(self):

        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

        if not os.path.exists(self.file_path):
            pd.DataFrame(columns=self.REQUIRED_COLUMNS).to_excel(self.file_path, index=False)
            return

        df = pd.read_excel(self.file_path)

        for col in self.REQUIRED_COLUMNS:
            if col not in df.columns:
                if col == "ExpenseType":
                    df[col] = "Essential"
                elif col == "ExpenseYear":
                    df[col] = datetime.now().year
                else:
                    df[col] = ""

        df.to_excel(self.file_path, index=False)

    def _read(self):
        try:
            return pd.read_excel(self.file_path)
        except:
            return pd.DataFrame(columns=self.REQUIRED_COLUMNS)

    def _write(self, df):
        df.to_excel(self.file_path, index=False)

    def get_next_id(self, df):
        if df.empty:
            return 1
        return int(pd.to_numeric(df['ID'], errors='coerce').max()) + 1

    def add_expense(self, amount, description, expense_type='Essential', expense_year=None):

        df = self._read()

        now = datetime.now()

        new = {
            "ID": self.get_next_id(df),
            "Amount": float(amount),
            "Description": description,
            "ExpenseType": expense_type or "Essential",
            "Date": now.strftime("%Y-%m-%d"),
            "Time": now.strftime("%H:%M:%S"),
            "ExpenseYear": int(expense_year) if expense_year else now.year
        }

        df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)

        self._write(df)

        return new

    def get_all(self, year=None):

        df = self._read().fillna("")

        if year:
            df = df[df['ExpenseYear'] == int(year)]

        return df.sort_values('ID', ascending=False).to_dict("records")

    def delete(self, record_id):

        df = self._read()

        df = df[df['ID'] != int(record_id)]

        self._write(df)

        return True

    # ── Aliases used by app.py ────────────────────────────────────────────

    def get_all_expenses(self, year=None):
        return self.get_all(year)

    def delete_expense(self, record_id):
        return self.delete(record_id)

    def delete_all_expenses(self):
        df = self._read()
        df = df.iloc[0:0]  # empty, keep columns
        self._write(df)

    def update_expense(self, record_id, data):
        df = self._read()
        idx = df.index[df['ID'] == int(record_id)]
        if idx.empty:
            return False
        for k, v in data.items():
            if k in df.columns:
                df.loc[idx, k] = v
        self._write(df)
        return True

    def get_available_years(self):
        df = self._read()
        if 'ExpenseYear' not in df.columns or df.empty:
            return []
        return sorted(df['ExpenseYear'].dropna().astype(int).unique().tolist(), reverse=True)

    def get_stats(self, year=None):
        df = self._read().fillna(0)
        
        available_years = self.get_available_years()
        selected_year = int(year) if year else ""

        if year:
            df = df[df['ExpenseYear'] == int(year)]

        amount_col = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
        total_amount = float(amount_col.sum())

        df['ExpenseType'] = df['ExpenseType'].replace(0, 'Essential').fillna('Essential')
        
        essential_total = float(amount_col[df['ExpenseType'] == 'Essential'].sum())
        side_total = float(amount_col[df['ExpenseType'] == 'Side'].sum())

        # by_type
        by_type = {}
        for t, g in df.groupby('ExpenseType'):
            by_type[str(t)] = float(pd.to_numeric(g['Amount'], errors='coerce').sum())

        # by_description
        by_desc_series = df.groupby('Description')['Amount'].apply(lambda x: pd.to_numeric(x, errors='coerce').sum())
        by_desc_series = by_desc_series[by_desc_series > 0].sort_values(ascending=False).head(10)
        by_description = {str(k): float(v) for k, v in by_desc_series.items()}

        # by_date
        by_date_series = df.groupby('Date')['Amount'].apply(lambda x: pd.to_numeric(x, errors='coerce').sum())
        by_date_series = by_date_series[by_date_series > 0].sort_index().tail(30)
        by_date = {str(k): float(v) for k, v in by_date_series.items()}

        # top_expenses
        df['numeric_amount'] = amount_col
        top_df = df.sort_values(by='numeric_amount', ascending=False).head(10)
        top_expenses = []
        for _, row in top_df.iterrows():
            if row['numeric_amount'] > 0:
                top_expenses.append({
                    'Description': str(row.get('Description', '')),
                    'Amount': round(float(row['numeric_amount']), 2)
                })

        return {
            'total_amount': round(total_amount, 2),
            'essential_total': round(essential_total, 2),
            'side_total': round(side_total, 2),
            'total_entries': len(df),
            'by_type': by_type,
            'by_description': by_description,
            'by_date': by_date,
            'top_expenses': top_expenses,
            'available_years': available_years,
            'selected_year': selected_year
        }

    @staticmethod
    def infer_expense_type(description):
        desc = (description or '').lower()
        if any(w in desc for w in ['fuel', 'petrol', 'benzine', 'وقود', 'بنزين']):
            return 'Transport'
        if any(w in desc for w in ['food', 'meal', 'اكل', 'طعام', 'وجبة']):
            return 'Food'
        return 'Essential'
