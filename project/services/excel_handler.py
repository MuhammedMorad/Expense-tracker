import pandas as pd
import os
from datetime import datetime

class ExcelHandler:
    def __init__(self, file_path='data/expenses.xlsx'):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.file_path = os.path.join(base_dir, file_path)
        self.ensure_file_exists()

    def ensure_file_exists(self):
        directory = os.path.dirname(self.file_path)
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        required_columns = ['ID', 'Amount', 'Description', 'ExpenseType', 'Date', 'Time', 'ExpenseYear']
        
        if not os.path.exists(self.file_path):
            df = pd.DataFrame(columns=required_columns)
            df.to_excel(self.file_path, index=False)
        else:
            try:
                df = pd.read_excel(self.file_path)
                changed = False
                
                # Backfill logic
                if 'ExpenseYear' not in df.columns:
                    # If ExpenseYear missing, infer from Date
                    if 'Date' in df.columns and not df.empty:
                        df['DateDt'] = pd.to_datetime(df['Date'], errors='coerce')
                        df['ExpenseYear'] = df['DateDt'].dt.year.fillna(datetime.now().year).astype(int)
                        # Drop temp
                        df = df.drop(columns=['DateDt'])
                    else:
                        df['ExpenseYear'] = datetime.now().year
                    changed = True
                
                for col in required_columns:
                    if col not in df.columns:
                        if col == 'ExpenseType':
                            df['ExpenseType'] = 'Essential'
                        else:
                            df[col] = ''
                        changed = True
                        
                if changed:
                    df.to_excel(self.file_path, index=False)
            except:
                pass

    def get_next_id(self, df):
        if df.empty: return 1
        if 'ID' not in df.columns: return 1 
        try:
            ids = pd.to_numeric(df['ID'], errors='coerce').fillna(0)
            return int(ids.max()) + 1
        except: return 1

    def infer_expense_type(self, description):
        description = description.lower()
        sides = ['فسحة', 'سينما', 'ترفيه', 'gift', 'cafe', 'كافيه', 'خروجة', 'هدايا', 'لعبة']
        if any(x in description for x in sides): return 'Side'
        return 'Essential'

    def add_expense(self, amount, description, expense_type=None, expense_year=None, custom_date=None):
        try:
            df = pd.read_excel(self.file_path)
        except Exception:
            df = pd.DataFrame(columns=['ID', 'Amount', 'Description', 'ExpenseType', 'Date', 'Time', 'ExpenseYear'])

        new_id = self.get_next_id(df)
        now = datetime.now()
        
        if not expense_type:
            expense_type = self.infer_expense_type(description)

        # Date is creation date (or custom if provided, but typically system now)
        # ExpenseYear is the accounting year
        expense_date = custom_date if custom_date else now.strftime('%Y-%m-%d')
        
        if not expense_year:
            # Default to year of the expense_date
            try:
                dt = datetime.strptime(expense_date, '%Y-%m-%d')
                expense_year = dt.year
            except:
                expense_year = now.year

        new_record = {
            'ID': new_id,
            'Amount': float(amount),
            'Description': description,
            'ExpenseType': expense_type,
            'Date': expense_date,
            'Time': now.strftime('%H:%M:%S'),
            'ExpenseYear': int(expense_year)
        }
        
        new_df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
        new_df.to_excel(self.file_path, index=False)
        return new_record

    def get_all_expenses(self, year=None):
        try:
            if not os.path.exists(self.file_path):
                self.ensure_file_exists()
            df = pd.read_excel(self.file_path)
            df = df.fillna('')
            
            if not df.empty and 'ID' in df.columns:
                if year:
                    try: 
                        # Filter by ExpenseYear instead of Date
                        df = df[df['ExpenseYear'] == int(year)]
                    except: pass
                
                df = df.sort_values(by='ID', ascending=False)
            return df.to_dict('records')
        except Exception as e:
            print(f"Error reading excel: {e}")
            return []
            
    def get_available_years(self):
        try:
            if not os.path.exists(self.file_path): return []
            df = pd.read_excel(self.file_path)
            if df.empty or 'ExpenseYear' not in df.columns: return []
            return sorted(df['ExpenseYear'].dropna().unique().astype(int).tolist(), reverse=True)
        except: return []

    def delete_expense(self, record_id):
        try:
            df = pd.read_excel(self.file_path)
            df = df[df['ID'] != int(record_id)]
            df.to_excel(self.file_path, index=False)
            return True
        except Exception as e:
            return False

    def delete_all_expenses(self):
        try:
            required_columns = ['ID', 'Amount', 'Description', 'ExpenseType', 'Date', 'Time', 'ExpenseYear']
            df = pd.DataFrame(columns=required_columns)
            df.to_excel(self.file_path, index=False)
            return True
        except Exception:
            return False

    def update_expense(self, record_id, data):
        try:
            df = pd.read_excel(self.file_path)
            mask = df['ID'] == int(record_id)
            if not mask.any(): return False
            idx = df[mask].index[0]
            
            for key, value in data.items():
                if key in df.columns:
                    if key == 'Amount':
                        try: value = float(value)
                        except: pass
                    # If updating ExpenseYear, ensure int
                    if key == 'ExpenseYear':
                        try: value = int(value)
                        except: pass
                    df.at[idx, key] = value
            
            df.to_excel(self.file_path, index=False)
            return True
        except Exception as e:
            return False

    def get_stats(self, selected_year=None):
        try:
            df = pd.read_excel(self.file_path)
            if df.empty:
                return {
                    'total_entries': 0, 'total_amount': 0, 'average_amount': 0,
                    'by_type': {}, 'by_description': {}, 'by_date': {}, 'top_expenses': [],
                    'essential_total': 0, 'side_total': 0, 'available_years': [], 'selected_year': selected_year
                }

            # Backfill ExpenseYear in memory if missing (though ensure_file_exists should handle it)
            if 'ExpenseYear' not in df.columns:
                 df['TempDate'] = pd.to_datetime(df['Date'], errors='coerce')
                 df['ExpenseYear'] = df['TempDate'].dt.year.fillna(datetime.now().year).astype(int)

            available_years = sorted(df['ExpenseYear'].dropna().unique().astype(int).tolist(), reverse=True)
            
            # Filter by ExpenseYear
            if selected_year:
                try:
                    selected_year = int(selected_year)
                    df = df[df['ExpenseYear'] == selected_year]
                except: pass
            
            df['DateStr'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
            
            total_entries = len(df)
            total_amount = df['Amount'].sum()
            average_amount = round(total_amount / total_entries, 2) if total_entries > 0 else 0
            
            if 'ExpenseType' not in df.columns: df['ExpenseType'] = 'Essential'
            df['ExpenseType'] = df['ExpenseType'].replace('', 'Essential')
            
            by_type = df.groupby('ExpenseType')['Amount'].sum().to_dict()
            essential_total = by_type.get('Essential', 0)
            side_total = by_type.get('Side', 0)
            
            by_description_all = df.groupby('Description')['Amount'].sum().sort_values(ascending=False)
            by_description = by_description_all.head(15).to_dict()
            
            by_date = df.groupby('DateStr')['Amount'].sum().to_dict()
            
            top_expenses = df.nlargest(5, 'Amount')[['Amount', 'Description', 'DateStr', 'ExpenseType']].rename(columns={'DateStr': 'Date'}).to_dict('records')
            
            return {
                'total_entries': total_entries,
                'total_amount': total_amount,
                'average_amount': average_amount,
                'by_type': by_type,
                'by_description': by_description,
                'by_date': by_date,
                'top_expenses': top_expenses,
                'essential_total': essential_total,
                'side_total': side_total,
                'available_years': available_years,
                'selected_year': selected_year
            }
        except Exception as e:
            print(f"Error stats: {e}")
            return {}

    def parse_uploaded_report(self, file_path):
        try:
            # Read all as string to find header and date
            df_raw = pd.read_excel(file_path, header=None)
            header_row_idx = -1
            report_period = ''
            
            # Simple Regex for dates YYYY-MM-DD
            import re
            date_pattern = r'\d{4}-\d{2}-\d{2}'
            
            for i, row in df_raw.iterrows():
                row_text = ' '.join(row.astype(str).tolist())
                
                # Check for Period/Date
                # Assuming format like "الفترة من 2024-01-01 الى 2024-12-31" from the user image
                if 'الفترة' in row_text or re.search(date_pattern, row_text):
                    dates = re.findall(date_pattern, row_text)
                    if len(dates) >= 2:
                        report_period = f"{dates[0]} to {dates[1]}"
                    elif len(dates) == 1:
                        report_period = dates[0]
                
                # Check for Header
                if 'الصنف' in row_text:
                    header_row_idx = i
                    break
            
            if header_row_idx != -1:
                df = pd.read_excel(file_path, header=header_row_idx)
                if 'الصنف' in df.columns:
                    df = df.dropna(subset=['الصنف'])
                    df = df[df['الصنف'].astype(str) != 'nan']
                
                # Add Report Period Column
                if report_period:
                    df['ReportPeriod'] = report_period
                else:
                    df['ReportPeriod'] = 'Unknown'
                    
                return df.fillna('').to_dict('records')
            return []
        except Exception as e:
            print(f"Parse Error: {e}")
            return []

    def ensure_reports_file_exists(self):
        reports_path = os.path.join(os.path.dirname(self.file_path), 'uploaded_reports.xlsx')
        if not os.path.exists(reports_path):
            df = pd.DataFrame(columns=['ID'])
            df.to_excel(reports_path, index=False)
            return
        try:
            df = pd.read_excel(reports_path)
            if 'ID' not in df.columns:
                df['ID'] = range(1, len(df) + 1)
                df.to_excel(reports_path, index=False)
        except Exception:
            pass

    def save_uploaded_data(self, data):
        if not data: return False
        try:
            self.ensure_reports_file_exists()
            reports_path = os.path.join(os.path.dirname(self.file_path), 'uploaded_reports.xlsx')
            
            new_df = pd.DataFrame(data)
            existing_df = pd.read_excel(reports_path)
            
            next_id = 1
            if not existing_df.empty and 'ID' in existing_df.columns:
                try:
                    next_id = int(pd.to_numeric(existing_df['ID'], errors='coerce').max()) + 1
                except:
                    next_id = len(existing_df) + 1
                    
            new_df['ID'] = range(next_id, next_id + len(new_df))
            
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                
            combined_df.to_excel(reports_path, index=False)
            return True
        except Exception as e:
            print(f"Error saving uploaded report: {e}")
            return False

    def get_uploaded_data(self):
        try:
            self.ensure_reports_file_exists()
            reports_path = os.path.join(os.path.dirname(self.file_path), 'uploaded_reports.xlsx')
            df = pd.read_excel(reports_path)
            # Ensure ID is treated as int
            if 'ID' in df.columns:
                df['ID'] = pd.to_numeric(df['ID'], errors='coerce').fillna(0).astype(int)
            return df.fillna('').to_dict('records')
        except Exception:
            return []

    def delete_report(self, record_id):
        try:
            reports_path = os.path.join(os.path.dirname(self.file_path), 'uploaded_reports.xlsx')
            df = pd.read_excel(reports_path)
            df = df[df['ID'] != int(record_id)]
            df.to_excel(reports_path, index=False)
            return True
        except Exception as e:
            return False

    def delete_all_reports(self):
        try:
            reports_path = os.path.join(os.path.dirname(self.file_path), 'uploaded_reports.xlsx')
            df = pd.DataFrame(columns=['ID'])
            df.to_excel(reports_path, index=False)
            return True
        except Exception:
            return False

    def update_report(self, record_id, data):
        try:
            reports_path = os.path.join(os.path.dirname(self.file_path), 'uploaded_reports.xlsx')
            df = pd.read_excel(reports_path)
            mask = df['ID'] == int(record_id)
            if not mask.any(): return False
            idx = df[mask].index[0]
            
            for key, value in data.items():
                if key in df.columns and key != 'ID':
                    try:
                        # minimal numeric attempt
                        if str(value).replace('.','',1).isdigit():
                            value = float(value)
                    except: pass
                    df.at[idx, key] = value
            
            df.to_excel(reports_path, index=False)
            return True
        except Exception as e:
            return False

    def get_reports_analytics(self):
        try:
            data = self.get_uploaded_data()
            if not data:
                return {
                    'commission_by_period': {},
                    'top_items': [],
                    'total_vs_converted': {'total': 0, 'converted': 0}
                }
            
            df = pd.DataFrame(data)
            
            # Ensure numeric
            df['العمولة قبل إجمالي'] = pd.to_numeric(df['العمولة قبل إجمالي'], errors='coerce').fillna(0)
            
            # Handle potential column name mismatch for converted
            converted_col = 'العمولة قيمة' if 'العمولة قيمة' in df.columns else 'العمولة المحولة'
            if converted_col in df.columns:
                df[converted_col] = pd.to_numeric(df[converted_col], errors='coerce').fillna(0)
            else:
                df[converted_col] = 0
                
            # 1. Commission by Period
            if 'ReportPeriod' in df.columns:
                by_period = df.groupby('ReportPeriod')['العمولة قبل إجمالي'].sum().to_dict()
            else:
                by_period = {'Unknown': df['العمولة قبل إجمالي'].sum()}
                
            # 2. Top Items by Commission
            if 'الصنف' in df.columns:
                top_items = df.groupby('الصنف')['العمولة قبل إجمالي'].sum().sort_values(ascending=False).head(5).to_dict()
            else:
                top_items = {}
                
            # 3. Total vs Converted
            total_comm = df['العمولة قبل إجمالي'].sum()
            total_conv = df[converted_col].sum()
            
            # 4. Quantity and Records
            if 'عدد' in df.columns:
                df['عدد'] = pd.to_numeric(df['عدد'], errors='coerce').fillna(0)
                total_quantity = int(df['عدد'].sum())
            else:
                total_quantity = 0
            
            total_records = len(df)

            return {
                'commission_by_period': by_period,
                'top_items': top_items,
                'total_vs_converted': {'total': total_comm, 'converted': total_conv},
                'scalar': {
                    'records': total_records,
                    'quantity': total_quantity,
                    'commission': round(total_comm, 2),
                    'converted': round(total_conv, 2)
                }
            }
        except Exception as e:
            print(f"Analytics Error: {e}")
            return {}
