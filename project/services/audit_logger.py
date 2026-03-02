import pandas as pd
import os
from datetime import datetime

class AuditLogger:
    def __init__(self, file_path='data/user_audit_log.xlsx'):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.file_path = os.path.join(base_dir, file_path)
        self.ensure_file_exists()

    def ensure_file_exists(self):
        directory = os.path.dirname(self.file_path)
        if not os.path.exists(directory):
            try: os.makedirs(directory)
            except: pass
        
        if not os.path.exists(self.file_path):
            try:
                df = pd.DataFrame(columns=['Timestamp', 'Username', 'Role', 'Action', 'Details'])
                df.to_excel(self.file_path, index=False)
            except: pass

    def log_action(self, username, role, action, details=None):
        try:
            if os.path.exists(self.file_path):
                df = pd.read_excel(self.file_path)
            else:
                df = pd.DataFrame(columns=['Timestamp', 'Username', 'Role', 'Action', 'Details'])

            new_entry = {
                'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Username': username,
                'Role': role,
                'Action': action,
                'Details': str(details) if details else ''
            }
            
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            df.to_excel(self.file_path, index=False)
            df.to_excel(self.file_path, index=False)
        except Exception as e:
            print(f"Audit Log Error: {e}")

    def get_logs(self):
        try:
            if not os.path.exists(self.file_path): return []
            df = pd.read_excel(self.file_path)
            # Sort by Timestamp descending if possible
            if 'Timestamp' in df.columns:
                df = df.sort_values(by='Timestamp', ascending=False)
            return df.fillna('').to_dict('records')
        except Exception:
            return []
