import logging
import os
from storage_verificar import main
from storage_settings import EXCEL_PATH, log_file_path
from datetime import datetime

# Configurar logging
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

def run_daily():
    logging.info(f"Iniciando verificação diária. Diretório atual: {os.getcwd()}")
    logging.info(f"Tentando acessar: \\\\192.168.0.36\\bkp\\VSC")
    try:
        # Verificar se o caminho remoto existe
        if os.path.exists(r"\\192.168.0.36\bkp\VSC"):
            logging.info("Diretório remoto acessível.")
            # Listar conteúdo para depuração
            dir_contents = os.listdir(r"\\192.168.0.36\bkp\VSC")
            logging.info(f"Conteúdo do diretório: {dir_contents}")
        else:
            logging.error("Diretório remoto não existe ou não está acessível.")
        success = main(
            excel_path=EXCEL_PATH,
            send_email=True,
            stop_event=None,
            progress_callback=None,
        )
        if success:
            logging.info("Verificação concluída com sucesso.")
        else:
            logging.error("Falha na verificação.")
    except Exception as e:
        logging.error(f"Erro durante a verificação: {e}", exc_info=True)

if __name__ == "__main__":
    run_daily()