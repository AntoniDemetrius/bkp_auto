import logging
import sys
import os
from dotenv import load_dotenv
import signal
import sys

# .env config email
load_dotenv()
smtp_server = os.getenv("SMTP_SERVER")
smtp_port = int(os.getenv("SMTP_PORT", "587"))
smtp_user = os.getenv("SMTP_USER")
smtp_pass = os.getenv("SMTP_PASS")
from_email = os.getenv("FROM_EMAIL")
to_email = os.getenv("TO_EMAIL")

# LOG na pasta raiz
if getattr(sys, "frozen", False):
    # executável
    main_dir = os.path.dirname(sys.executable)
else:
    # script
    main_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

# Caminho completo do log
log_file_path = os.path.join(main_dir, "log.log")

logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8",
)

MAX_RETRIES = 3
RETRY_DELAY = 2  # segundos
NETWORK_TIMEOUT = 30  # segundos
MAX_THREADS = 5

EXCEL_PATH = r"C:\Users\antoni.demetrius\OneDrive\Documentos\backup_auto\excel\Acompanhamento de Backups 2025.xlsx"
STORAGE_BASE = r"\\192.168.0.36\bkp\VSC"
BACKUP_EXT = [".rar", ".zip", ".lscx"]
DATA_FORMATO = "%d/%m/%Y"

STOP_VERIFICATION = False
