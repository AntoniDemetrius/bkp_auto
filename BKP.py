import customtkinter as ctk
import threading
import os
import logging
import tkinter.filedialog as filedialog
import pystray
from PIL import Image
import time
from storage_verificar import main
from storage_settings import *

class BackupCheckerGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("BKP - Verificador de Backup")
        self.geometry("360x640")
        self.resizable(False, True)
        self.configure(padx=0, pady=0)

        # Definir ícone da janela principal
        script_dir = os.path.dirname(os.path.abspath(__file__))
        ico_path = os.path.join(script_dir, "BKP.ico")
        try:
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
                logging.info(f"Ícone da janela principal definido com sucesso: {ico_path}")
            else:
                logging.warning(f"Arquivo BKP.ico não encontrado em {ico_path}, ícone padrão será usado.")
        except Exception as e:
            logging.error(f"Erro ao definir ícone da janela: {e}")

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # Variáveis para os switches
        self.send_email_var = ctk.BooleanVar(value=True)
        self.theme_var = ctk.BooleanVar(value=True)
        self._toggling_theme = False  # evitar chamadas simultâneas
        self.log_open = False  # Rastrear o estado do log
        self.excel_open = False  # Rastrear o estado do Excel

        # Menu suspenso com opções dinâmicas baseadas no estado atual
        email_option = "Desativar E-mail" if self.send_email_var.get() else "Ativar E-mail"
        log_option = "Abrir Log" if not self.log_open else "Fechar Log"
        excel_option = "Abrir Excel" if not self.excel_open else "Fechar Excel"
        theme_option = "Tema Claro" if self.theme_var.get() else "Tema Escuro"
        self.optionmenu = ctk.CTkOptionMenu(
            self,
            values=[email_option, log_option, excel_option, theme_option],
            command=self.optionmenu_callback,
            font=("Arial", 14),
            width=100,
            height=10,
            corner_radius=10,
            fg_color="#1E90FF",
            button_color="#1E90FF",
            button_hover_color="#3B3B3B",
        )
        self.optionmenu.set("≡")  # Valor inicial
        self.optionmenu.pack(pady=10, padx=10)

        self.excel_path = ctk.StringVar(value=EXCEL_PATH or "Nenhum arquivo")
        self.excel_label = ctk.CTkLabel(
            self, textvariable=self.excel_path, font=("Arial", 12), wraplength=320
        )
        self.excel_label.pack(pady=5, fill="x", padx=20)
        self.select_excel_button = ctk.CTkButton(
            self,
            text="Buscar Excel",
            command=self.select_excel_file,
            width=200,
            height=40,
            font=("Arial", 14),
            corner_radius=10,
            hover_color="#1E90FF",
        )
        self.select_excel_button.pack(pady=10, fill="x", padx=20)

        self.action_button = ctk.CTkButton(
            self,
            text="Iniciar",
            command=self.start_verification,
            width=200,
            height=40,
            font=("Arial", 14),
            corner_radius=10,
            hover_color="#1E90FF",
        )
        self.action_button.pack(pady=10, fill="x", padx=20)

        self.status_label = ctk.CTkLabel(
            self, text="Pronto para verificar!", font=("Arial", 16, "bold")
        )
        self.status_label.pack(pady=0, fill="x", padx=20)

        self.progressbar = ctk.CTkProgressBar(
            self,
            orientation="Horizontal",
            width=300,
            height=20,
            corner_radius=10,
            progress_color="#1E90FF",
            mode="determinate",
        )
        self.progressbar.pack(pady=0, fill="x", padx=20)
        self.progressbar.set(0.0)

        self.progress_label = ctk.CTkLabel(self, text="0%")
        self.progress_label.pack(pady=0)

        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(pady=10, fill="x", padx=20)

        self.icon = None
        self.icon_running = False
        self.setup_system_tray()
        self.start_system_tray()

        self.protocol("WM_ICONIFY", self.minimize_to_tray)
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        self.pending_events = []
        self.stop_event = threading.Event()
        self.verification_thread = None

    def optionmenu_callback(self, choice):
        logging.info(f"Opção selecionada: {choice}")
        try:
            if choice == "Ativar E-mail":
                self.send_email_var.set(True)
            elif choice == "Desativar E-mail":
                self.send_email_var.set(False)
            elif choice == "Tema Escuro":
                self.theme_var.set(True)
                self.toggle_theme()
            elif choice == "Tema Claro":
                self.theme_var.set(False)
                self.toggle_theme()
            elif choice == "Abrir Log":
                self.log_open = True
                self.open_log()
            elif choice == "Fechar Log":
                self.log_open = False
                self.status_label.configure(text="Log fechado!")
            elif choice == "Abrir Excel":
                self.excel_open = True
                self.open_excel()
            elif choice == "Fechar Excel":
                self.excel_open = False
                self.status_label.configure(text="Excel fechado!")

            # Atualizar as opções do menu com base nos novos estados
            email_option = "Desativar E-mail" if self.send_email_var.get() else "Ativar E-mail"
            log_option = "Abrir Log" if not self.log_open else "Fechar Log"
            excel_option = "Abrir Excel" if not self.excel_open else "Fechar Excel"
            theme_option = "Tema Claro" if self.theme_var.get() else "Tema Escuro"
            self.optionmenu.configure(
                values=[email_option, log_option, excel_option, theme_option]
            )
            self.optionmenu.set("≡")  # Reseta para o valor padrão
        except Exception as e:
            logging.error(f"Erro ao processar opção do menu: {e}")
            self.status_label.configure(text="Erro ao aplicar configuração!")

    def toggle_theme(self):
        if self._toggling_theme:
            logging.info("Troca de tema já em andamento, ignorando nova solicitação")
            return
        try:
            self._toggling_theme = True
            if self.theme_var.get():
                ctk.set_appearance_mode("dark")
                logging.info("Tema alterado para escuro")
            else:
                ctk.set_appearance_mode("light")
                logging.info("Tema alterado para claro")
            self.update()
        except Exception as e:
            logging.error(f"Erro ao alternar tema: {e}")
            self.status_label.configure(text="Erro ao alternar tema!")
        finally:
            self._toggling_theme = False

    def center_window(self):
        try:
            self.update_idletasks()
            window_width = self.winfo_width()
            window_height = self.winfo_height()
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            self.geometry(f"360x640+{x}+{y}")
        except Exception as e:
            logging.error(f"Erro ao centralizar a janela: {e}")

    def setup_system_tray(self):
        try:
            image = Image.open(
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "BKP.ico")
            )
        except FileNotFoundError:
            logging.warning(f"Arquivo BKP.ico não encontrado para a bandeja, usando ícone padrão.")
            image = Image.new("RGB", (64, 64), color="red")
        menu = (
            pystray.MenuItem("Restaurar", self.restore_window),
            pystray.MenuItem("Sair", self.quit_application),
        )
        self.icon = pystray.Icon("BackupChecker", image, "Verificador de Backups", menu)
        self.icon.on_clicked = lambda icon, event: (
            self.restore_window() if event.button == 1 and event.clicks >= 1 else None
        )

    def start_system_tray(self):
        if self.icon_running:
            return
        try:
            time.sleep(0.1)
            self.icon_running = True
            threading.Thread(target=self.icon.run, daemon=True).start()
        except Exception as e:
            logging.error(f"Erro ao iniciar o ícone da bandeja automaticamente: {e}")
            self.icon_running = False
            self.status_label.configure(text="Erro ao iniciar o ícone da bandeja!")

    def minimize_to_tray(self):
        if not self.icon_running:
            self.start_system_tray()
        try:
            self.withdraw()
        except Exception as e:
            logging.error(f"Erro ao minimizar para a bandeja: {e}")
            self.deiconify()
            self.center_window()
            self.status_label.configure(text="Erro ao minimizar para a bandeja!")

    def restore_window(self):
        try:
            self.deiconify()
            self.center_window()
        except Exception as e:
            logging.error(f"Erro ao restaurar a janela: {e}")
            self.status_label.configure(text="Erro ao restaurar a janela!")

    def quit_application(self):
        self.stop_event.set()
        try:
            for event_id in self.pending_events:
                try:
                    self.after_cancel(event_id)
                except:
                    pass
            self.pending_events.clear()
            if self.verification_thread and self.verification_thread.is_alive():
                self.verification_thread.join(timeout=2.0)
            if self.icon_running:
                try:
                    self.icon.stop()
                    self.icon_running = False
                except Exception as e:
                    logging.error(f"Erro ao parar o ícone da bandeja ao fechar: {e}")
            self.destroy()
        except Exception as e:
            logging.error(f"Erro ao fechar a aplicação: {e}")

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path.set(file_path)
            self.status_label.configure(text="Excel selecionado!")
            logging.info(f"Excel selecionado: {file_path}")

    def open_excel(self):
        excel_path = self.excel_path.get()
        if os.path.exists(excel_path):
            try:
                os.startfile(excel_path)
                self.status_label.configure(text="Excel aberto!")
                logging.info(f"Excel aberto: {excel_path}")
            except Exception as e:
                self.status_label.configure(text="Erro ao abrir Excel!")
                logging.error(f"Erro ao abrir Excel {excel_path}: {e}")
        else:
            self.status_label.configure(text="Erro: Excel não encontrado!")
            logging.error(f"Excel não encontrado: {excel_path}")

    def open_log(self):
        log_path = log_file_path
        if os.path.exists(log_path):
            try:
                os.startfile(log_path)
                self.status_label.configure(text="Log aberto no Notepad!")
                logging.info(f"Log aberto: {log_path}")
            except Exception as e:
                self.status_label.configure(text="Erro ao abrir log!")
                logging.error(f"Erro ao abrir log {log_path}: {e}")
        else:
            self.status_label.configure(text="Erro: Log não encontrado!")
            logging.error(f"Log não encontrado: {log_path}")

    def start_verification(self):
        excel_path = self.excel_path.get()
        if (
            not excel_path
            or not os.path.exists(excel_path)
            or not excel_path.endswith((".xlsx", ".xls"))
        ):
            self.status_label.configure(text="Selecione um arquivo Excel válido!")
            logging.error("Nenhum arquivo Excel válido selecionado.")
            return
        self.stop_event.clear()
        self.progressbar.set(0.0)
        self.progress_label.configure(text="0%")
        self.action_button.configure(
            text="Parar", command=self.stop_verification, hover_color="#8F2626"
        )
        self.select_excel_button.pack_forget()  # Alterado de open_excel_button para select_excel_button
        self.optionmenu.configure(state="disabled")
        self.status_label.configure(text="Processando...")
        self.verification_thread = threading.Thread(target=self.run_main, daemon=True)
        self.verification_thread.start()

    def stop_verification(self):
        self.stop_event.set()
        self.status_label.configure(text="Interrompendo verificação...")
        logging.info("Interrupção da verificação solicitada.")
        self.progressbar.set(0.0)
        self.progress_label.configure(text="0%")
        event_id = self.after(100, self.check_verification_stopped)
        self.pending_events.append(event_id)

    def check_verification_stopped(self):
        if self.verification_thread and self.verification_thread.is_alive():
            event_id = self.after(100, self.check_verification_stopped)
            self.pending_events.append(event_id)
        else:
            self.progressbar.set(0.0)
            self.progress_label.configure(text="0%")
            self.status_label.configure(
                text="Verificação interrompida! Pronto para nova verificação."
            )
            self.action_button.configure(
                text="Iniciar", command=self.start_verification, hover_color="#1E90FF"
            )
            self.optionmenu.configure(state="normal")
            self.select_excel_button.pack(pady=10, fill="x", padx=20)  # Restaurar botão
            self.select_excel_button.configure(state="normal")
            self.verification_thread = None

    def run_main(self):
        def update_progress(target_value):
            try:
                target_value = max(0.0, min(1.0, target_value))
                self.progressbar.set(target_value)
                self.progress_label.configure(text=f"{int(target_value * 100)}%")
                self.update()
                logging.info(f"Progresso atualizado: {int(target_value * 100)}%")
            except Exception as e:
                logging.error(f"Erro ao atualizar progresso: {e}")

        success = main(
            excel_path=self.excel_path.get(),
            send_email=self.send_email_var.get(),
            stop_event=self.stop_event,
            progress_callback=update_progress,
        )
        event_id = self.after(0, self.finish_verification, success)
        self.pending_events.append(event_id)

    def finish_verification(self, success):
        self.action_button.configure(
            text="Iniciar", command=self.start_verification, hover_color="#1E90FF"
        )
        self.optionmenu.configure(state="normal")
        if success:
            self.progressbar.set(1.0)
            self.progress_label.configure(text="100%")
            self.status_label.configure(text="Verificação concluída!")
            logging.info("Verificação concluída.")
            self.select_excel_button.pack(pady=10, fill="x", padx=20)  # Restaurar botão
            self.select_excel_button.configure(state="normal")
            excel_path = self.excel_path.get()
            if os.path.exists(excel_path):
                try:
                    logging.info(f"Excel pronto para abertura após verificação: {excel_path}")
                except Exception as e:
                    self.status_label.configure(text="Erro ao abrir Excel após verificação!")
                    logging.error(f"Erro ao abrir Excel após verificação {excel_path}: {e}")
        else:
            self.progressbar.set(0.0)
            self.progress_label.configure(text="0%")
            self.status_label.configure(text="Erro na verificação!")
            logging.error("Erro na verificação.")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    app = BackupCheckerGUI()
    app.center_window()
    app.mainloop()