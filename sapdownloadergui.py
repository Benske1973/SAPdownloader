import os
import time
import win32com.client as win32
from datetime import datetime
from playwright.sync_api import sync_playwright
import subprocess
import sys
import gc
import shutil
import threading
import pythoncom
import tkinter as tk
from tkinter import scrolledtext, ttk, messagebox

# --- INSTELLINGEN ---
DOWNLOAD_MAP = os.path.join(os.path.expanduser("~"), "EQUANS", "Projects Elek Wetteren - Documenten")
TIJDELIJKE_OPSLAG = os.path.join(os.getcwd(), "TEMP_BEX_DOWNLOADS")

SITES = [
    {"naam": "Ageing Balance", "url": "https://vheqnbpeci.sap.myequans.com:50001/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=CJ8INZ0NKWJE5I0NXFYRGX2XQ", "bestandsnaam": "Ageing Balance.xlsx", "vul_datum": True},
    {"naam": "Customer Financial Turnover", "url": "https://vheqnbpeci.sap.myequans.com:50001/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=CJ8INZ0NKWJFVSJC701F0I6VD", "bestandsnaam": "Customer Financial Turnover.xlsx", "vul_datum": False},
    {"naam": "Customer Invoicing", "url": "https://vheqnbpeci.sap.myequans.com:50001/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=CJ8INZ0NKWJFVSJBYYGNDSDZ1", "bestandsnaam": "Customer Invoicing.xlsx", "vul_datum": False},
    {"naam": "Internal Subcontracting", "url": "https://vheqnbpeci.sap.myequans.com:50001/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=CJ8INZ0NKWJE5I4S8H7CJGETU", "bestandsnaam": "Internal Subcontracting.xlsx", "vul_datum": False},
    {"naam": "PCB", "url": "https://vheqnbpeci.sap.myequans.com:50001/irj/servlet/prt/portal/prtroot/pcd!3aportal_content!2fcom.sap.pct!2fplatform_add_ons!2fcom.sap.ip.bi!2fiViews!2fcom.sap.ip.bi.bex?BOOKMARK=CJ8INZ0NKWJE5IRLO1DS2NNNA", "bestandsnaam": "PCB.xlsx", "vul_datum": False}
]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Equans SAP Downloader")
        self.root.geometry("600x450")
        
        self.lbl_info = tk.Label(root, text="SAP Automatisering - Klik op Start", font=("Arial", 12, "bold"))
        self.lbl_info.pack(pady=10)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=550, mode="determinate")
        self.progress.pack(pady=5)

        self.txt_log = scrolledtext.ScrolledText(root, height=15, width=70, state='disabled', font=("Consolas", 9))
        self.txt_log.pack(pady=10, padx=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        self.btn_start = tk.Button(btn_frame, text="START DOWNLOADS", command=self.start_thread, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=20)
        self.btn_start.pack(side=tk.LEFT, padx=10)

        self.btn_close = tk.Button(btn_frame, text="Afsluiten", command=root.destroy, state='disabled', width=15)
        self.btn_close.pack(side=tk.LEFT, padx=10)

    def log(self, message):
        self.root.after(0, self._log_impl, message)

    def _log_impl(self, message):
        self.txt_log.config(state='normal')
        self.txt_log.insert(tk.END, message + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    def update_progress(self, value):
        self.root.after(0, lambda: self.progress.configure(value=value))

    def enable_close(self):
        self.root.after(0, lambda: self.btn_close.config(state='normal'))
        self.root.after(0, lambda: self.btn_start.config(state='disabled'))

    def start_thread(self):
        self.btn_start.config(state='disabled')
        self.log("🚀 Proces gestart...")
        threading.Thread(target=self.run_process, daemon=True).start()

    # --- DE LOGICA FUNCTIES ---
    def kill_excel(self):
        try:
            subprocess.run("taskkill /F /IM excel.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(1)
        except: pass

    def ensure_playwright_browsers(self):
        self.log("🔍 Controleren op Playwright browsers...")
        
        # Check of we in een EXE zitten (Frozen)
        is_frozen = getattr(sys, 'frozen', False)

        try:
            # Manier 1: Probeer 'playwright' direct aan te roepen (werkt als het in PATH staat)
            subprocess.run(["playwright", "install", "chromium"], check=True, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            self.log("✔ Browsers gevonden/gereed.")
        except:
            if not is_frozen:
                # Manier 2: Alleen veilig als we NIET in een EXE zitten
                self.log("⚠️ Browsers worden geïnstalleerd (script mode)...")
                try:
                    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], shell=True)
                except:
                    self.log("❌ Kon browsers niet installeren.")
            else:
                # We zijn een EXE en 'playwright' command faalde.
                # We kunnen niet sys.executable aanroepen want dat start de GUI opnieuw.
                # We hopen dat de browsers er al zijn (meestal wel na 1x draaien op PC).
                self.log("ℹ️ In EXE-modus: automatische installatie overgeslagen. Hopelijk zijn browsers aanwezig.")

    def converteer_bestand(self, excel_app, bron_pad, doel_map, finale_naam):
        wb = None
        try:
            abs_bron = os.path.abspath(bron_pad)
            abs_doel = os.path.abspath(os.path.join(doel_map, finale_naam))
            
            wb = excel_app.Workbooks.Open(abs_bron)
            try: wb.Worksheets(1).Name = "YANALYSIS_PATTERN"
            except: pass

            if os.path.exists(abs_doel):
                os.remove(abs_doel)

            wb.SaveAs(abs_doel, FileFormat=51)
            wb.Close(SaveChanges=True)
            return True
        except Exception as e:
            self.log(f"  ❌ Excel-fout: {e}")
            if wb: 
                try: wb.Close(SaveChanges=False)
                except: pass
            return False

    def run_process(self):
        pythoncom.CoInitialize()

        try:
            vandaag = datetime.now().strftime("%d.%m.%Y")
            
            # 1. Voorbereiding
            self.ensure_playwright_browsers()
            self.kill_excel()

            if not os.path.exists(TIJDELIJKE_OPSLAG):
                os.makedirs(TIJDELIJKE_OPSLAG)
            else:
                for f in os.listdir(TIJDELIJKE_OPSLAG):
                    try: os.remove(os.path.join(TIJDELIJKE_OPSLAG, f))
                    except: pass

            # 2. Excel Starten
            self.log("⚙️ Excel opstarten op achtergrond...")
            excel_app = None
            try:
                excel_app = win32.DispatchEx('Excel.Application')
                excel_app.Visible = False
                excel_app.DisplayAlerts = False
            except Exception as e:
                self.log(f"❌ Kan Excel niet starten: {e}")
                self.enable_close()
                return

            # 3. Download Loop
            geslaagde_downloads = []
            totaal = len(SITES)
            step = 100 / (totaal + 1)
            current_progress = 0

            try:
                with sync_playwright() as p:
                    self.log("🌐 Browser starten...")
                    browser = p.chromium.launch(channel="msedge", headless=False)
                    context = browser.new_context(accept_downloads=True)

                    for i, site in enumerate(SITES):
                        self.log(f"📥 [{i+1}/{totaal}] {site['naam']} ophalen...")
                        page = context.new_page()
                        try:
                            page.goto(site['url'], timeout=60000)
                            bex_frame = page.frame_locator('iframe[name^="iframe_Roundtrip"]')
                            
                            if site["vul_datum"]:
                                try:
                                    datum_veld = bex_frame.locator("#DLG_VARIABLE_vsc_cvl_VAR_10_INPUT_inp")
                                    datum_veld.wait_for(state="visible", timeout=10000)
                                    datum_veld.fill(vandaag)
                                except:
                                    self.log("  ⚠️ Kon datumveld niet vinden (timeout)")

                            bex_frame.get_by_role("link", name="OK").click()
                            page.wait_for_timeout(8000)

                            with context.expect_page() as new_page_info:
                                bex_frame.get_by_role("link", name="Export to Microsoft Excel").click()
                            
                            launcher_page = new_page_info.value
                            with launcher_page.expect_download(timeout=60000) as download_info:
                                download = download_info.value
                            
                            temp_raw = os.path.join(TIJDELIJKE_OPSLAG, f"raw_{site['bestandsnaam']}.xls")
                            download.save_as(temp_raw)
                            
                            time.sleep(2) 

                            if self.converteer_bestand(excel_app, temp_raw, TIJDELIJKE_OPSLAG, site['bestandsnaam']):
                                geslaagde_downloads.append(site['bestandsnaam'])
                                self.log("  ✔ Gelukt")
                            
                            if os.path.exists(temp_raw):
                                try: os.remove(temp_raw)
                                except: pass

                            launcher_page.close()
                        except Exception as e:
                            self.log(f"  ❌ Fout: {e}")
                        finally:
                            page.close()
                        
                        current_progress += step
                        self.update_progress(current_progress)

                    browser.close()
            except Exception as e:
                self.log(f"❌ Kritieke Browser fout: {e}")

            # 4. Afsluiten Excel
            self.log("⚙️ Excel afsluiten...")
            if excel_app:
                excel_app.Quit()
                del excel_app
                gc.collect()

            # 5. Verplaatsen
            if geslaagde_downloads:
                self.log("\n📦 Verplaatsen naar SharePoint map...")
                self.kill_excel() 

                if not os.path.exists(DOWNLOAD_MAP):
                    self.log(f"⚠️ Doelmap niet gevonden:\n{DOWNLOAD_MAP}\nProbeer deze map aan te maken...")
                    try: os.makedirs(DOWNLOAD_MAP)
                    except: self.log("❌ Kon map niet aanmaken.")

                for bestand in geslaagde_downloads:
                    bron = os.path.join(TIJDELIJKE_OPSLAG, bestand)
                    doel = os.path.join(DOWNLOAD_MAP, bestand)
                    try:
                        shutil.move(bron, doel)
                        self.log(f"  ✔ Geupdate: {bestand}")
                    except PermissionError:
                        self.log(f"  ⚠️ FOUT: {bestand} is geopend door iemand anders!")
                    except Exception as e:
                        self.log(f"  ❌ Move fout: {e}")

            # Opruimen
            try: shutil.rmtree(TIJDELIJKE_OPSLAG)
            except: pass
            
            self.update_progress(100)
            self.log("\n✅ KLAAR! Je mag dit venster sluiten.")
            self.enable_close()
            messagebox.showinfo("Klaar", "Alle bestanden zijn bijgewerkt!")
        
        finally:
            pythoncom.CoUninitialize()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()