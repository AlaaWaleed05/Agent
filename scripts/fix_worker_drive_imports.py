from pathlib import Path
p = Path('worker.py')
s = p.read_text(encoding='utf-8')
s = s.replace('import json\n', 'import io\nimport json\n', 1) if 'import io\n' not in s else s
s = s.replace('from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload\n\nfrom cryptography.fernet import Fernet\n', 'from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload\n\nimport openpyxl\nfrom cryptography.fernet import Fernet\n', 1) if 'import openpyxl\n' not in s else s
p.write_text(s, encoding='utf-8')
