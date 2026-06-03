import json
import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import dotenv
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Playwright, TimeoutError, expect
import warnings
import pyxlsb
import getpass
import traceback


import time
from datetime import date, timedelta


dotenv.load_dotenv('.env')  # Load environment variables from .env file

from Tasks import process_planejamento



warnings.filterwarnings("ignore", category=UserWarning)
import sys

def get_playwright_browser_path():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
        chromium_path = os.path.join(base_path, "ms-playwright", "chromium-1187", "chrome-win", "chrome.exe")
    else:
        base_path = rf"C:\Users\{getpass.getuser()}\AppData\Local"

        # Join the rest of the Playwright folder path
        chromium_path = os.path.join(
            base_path,
            "ms-playwright",
            "chromium-1187",
            "chrome-win",
            "chrome.exe"
        )
   
    if chromium_path and not os.path.exists(chromium_path):
        raise FileNotFoundError(f"Chromium executable not found at {chromium_path}")

    return chromium_path


# --- GUI UPDATE FUNCTION ---
def update_gui(queue_instance, status_label, progress_bar, log_text, process_button=None):
    """Checks the queue for messages from the worker thread and updates the GUI."""
    try:
        while True:
            message_type, value = queue_instance.get_nowait()
            if message_type == "status":
                status_label.config(text=value)
                log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {value}\n")
                log_text.see(tk.END)
            elif message_type == "progress":
                progress_bar['value'] = value
            elif message_type == "done":
                status_label.config(text="Processo Concluído!")
                progress_bar['value'] = 100
                if process_button:
                    process_button.config(state="normal")
                return # Stop checking
    except queue.Empty:
        pass
    status_label.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text, process_button))


    
 


def run_automation(playwright: Playwright, q: queue.Queue, test_mode: bool = False):
    try:
        # EXPLICIT DEBUG at thread start
        debug_msg = f"{'='*70}\nDEBUG: run_automation iniciado com test_mode = {test_mode}\n{'='*70}"
        print(debug_msg)
        
        # 1. Load Credentials
        if test_mode:
            msg = "🧪 MODO DE TESTE ATIVADO"
            q.put(("status", msg))
            print(msg)
        else:
            msg = "🚀 MODO PRODUÇÃO ATIVADO"
            q.put(("status", msg))
            print(msg)
        
        q.put(("status", "Carregando credenciais..."))
        q.put(("progress", 5))
        
        # Process the planning and prepare emails
        email_list = process_planejamento(q=q)
        
        # Check if we have emails to send
        if not email_list or len(email_list) == 0:
            q.put(("status", "⚠️ Nenhum email foi preparado"))
            q.put(("progress", 100))
            q.put(("done", True))
            return
            
        q.put(("status", f"✅ {len(email_list)} emails preparados"))
        q.put(("progress", 90))
        
        # 2. Launch Browser to send emails via Outlook Web
        msg = "🌐 Abrindo Outlook Web para enviar emails..."
        q.put(("status", msg))
        print(msg)
        
        url_outlook = os.getenv("OUTLOOK_URL", "https://outlook.office.com/mail/")
        navegador = os.getenv("Nav", "chrome").lower()
        
        msg = f"🌐 Navegador selecionado: {navegador.upper()}"
        q.put(("status", msg))
        print(msg)
        
        # Determine browser profile path based on Nav setting
        if navegador == "edge":
            browser_user_data = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Edge", "User Data")
            browser_name = "Edge"
            channel = "msedge"
        else:  # chrome
            browser_user_data = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Google", "Chrome", "User Data")
            browser_name = "Chrome"
            channel = "chrome"
        
        # Check if browser default profile exists
        use_default_profile = os.path.exists(browser_user_data)
        
        if use_default_profile:
            msg = f"ℹ️ IMPORTANTE: Feche TODAS as janelas do {browser_name} antes de continuar!"
            q.put(("status", msg))
            print(msg)
            msg = "⏳ Aguardando 10 segundos para você fechar o navegador..."
            q.put(("status", msg))
            print(msg)
            import time
            time.sleep(10)
            
            msg = f"🔄 Encerrando processos em segundo plano do {browser_name} para liberar o perfil..."
            q.put(("status", msg))
            print(msg)
            
            # Força o encerramento de processos em background que "prendem" o perfil
            try:
                if navegador == "edge":
                    os.system("taskkill /F /IM msedge.exe /T >nul 2>&1")
                else:
                    os.system("taskkill /F /IM chrome.exe /T >nul 2>&1")
            except Exception:
                pass
            time.sleep(3) # Dá um tempinho para o Windows liberar os arquivos
            
            user_data_dir = browser_user_data
            msg = f"✅ Usando perfil do {browser_name} com suas credenciais salvas"
            q.put(("status", msg))
            print(msg)
        else:
            # Fallback to RPA-specific profile
            user_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "RPA_FIAPE_Browser_Data")
            os.makedirs(user_data_dir, exist_ok=True)
            msg = "ℹ️ Usando perfil RPA (você precisará fazer login na primeira vez)"
            q.put(("status", msg))
            print(msg)
        
        # Args to remove automation indicators and warnings
        browser_args = [
            "--start-maximized",
            "--disable-blink-features=AutomationControlled",
            "--disable-dev-shm-usage",
            "--disable-infobars",
            "--disable-extensions",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-timer-throttling",
            "--disable-backgrounding-occluded-windows",
            "--disable-renderer-backgrounding",
            "--test-type",
        ]
        
        msg = f"🚀 Iniciando {browser_name}..."
        q.put(("status", msg))
        print(msg)
        
        # Launch browser with persistent context (keeps login)
        try:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir,
                headless=False,
                channel=channel,
                args=browser_args,
                no_viewport=True,
                ignore_default_args=["--enable-automation", "--no-sandbox"]
            )
            msg = f"✅ {browser_name} iniciado com sucesso"
            q.put(("status", msg))
            print(msg)
        except Exception as e:
            error_msg = f"❌ Erro ao iniciar {browser_name}: {e}"
            q.put(("status", error_msg))
            print(error_msg)
            warning_msg = "⚠️ Tentando com perfil RPA..."
            q.put(("status", warning_msg))
            print(warning_msg)
            # Retry with RPA profile
            user_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "RPA_FIAPE_Browser_Data")
            os.makedirs(user_data_dir, exist_ok=True)
            context = playwright.chromium.launch_persistent_context(
                user_data_dir,
                headless=False,
                channel=channel,
                args=browser_args,
                no_viewport=True,
                ignore_default_args=["--enable-automation", "--no-sandbox"]
            )
        
        # Para garantir que o bot interaja com uma aba visível na tela:
        if len(context.pages) > 0:
            page = context.pages[0]
            # Se a primeira aba não for uma página em branco (ex: aba restaurada ou oculta), criamos uma nova
            if page.url not in ["about:blank", ""]:
                page = context.new_page()
        else:
            page = context.new_page()
            
        page.bring_to_front() # Força a aba e o navegador a virem para o primeiro plano
        
        # Hide webdriver property to remove automation detection
        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
        """)
        
        q.put(("status", f"🌐 Navegando para Outlook: {url_outlook}"))
        try:
            page.goto(url_outlook, timeout=90000)
            q.put(("status", "✅ Página Outlook carregada"))
        except Exception as e:
            q.put(("status", f"⚠️ Erro ao navegar: {e}"))
            q.put(("status", "Tentando continuar..."))
        
        q.put(("status", "⏳ Aguardando carregamento completo do Outlook..."))
        print("⏳ Aguardando carregamento completo do Outlook...")
        # Wait for Outlook to load - look for common elements
        try:
            # Wait for the page to load (try multiple selectors as Outlook UI may vary)
            page.wait_for_load_state("domcontentloaded", timeout=30000)
            page.wait_for_timeout(3000)  # Give UI time to render
            msg = "✅ Outlook carregado!"
            q.put(("status", msg))
            print(msg)
        except:
            msg = "⚠️ Outlook pode não estar totalmente carregado"
            q.put(("status", msg))
            print(msg)
        
        q.put(("progress", 95))
        
        # Check if user is logged in
        if "login" in page.url.lower() or "auth" in page.url.lower():
            msg = "⚠️ Você precisa fazer login no Outlook. Faça login e execute novamente."
            q.put(("status", msg))
            print(msg)
            page.wait_for_timeout(10000)  # Wait 10 seconds for user to see the message
        else:
            if test_mode:
                msg = "✅ Login detectado! Criando email de TESTE..."
            else:
                msg = "✅ Login detectado! Criando email de PRODUÇÃO..."
            q.put(("status", msg))
            print(msg)
            
            # Try to create a test email (first email from the list)
            if email_list and len(email_list) > 0:
                try:
                    test_email = email_list[0]
                    if test_mode:
                        msg = f"📧 TESTE: Criando email para {test_email['supplier_name']} (usando emails de teste)"
                    else:
                        msg = f"📧 PRODUÇÃO: Criando email para {test_email['supplier_name']}"
                    q.put(("status", msg))
                    print(msg)
                    
                    if not test_mode:
                        # Only print real email counts in production mode
                        print(f"   TO: {len(test_email['to_emails'])} emails (transportadora + fornecedor)")
                        print(f"   CC: {len([e for e in test_email['cc_emails'].split(';') if e.strip()])} emails Stellantis")
                    
                    # Wait a bit for page to be fully interactive
                    page.wait_for_timeout(3000)
                    
                    # Try to find and click "New Email" button
                    # Common selectors for Outlook Web
                    new_email_selectors = [
                        "button[aria-label*='New mail']",
                        "button[aria-label*='Nova mensagem']",
                        "button[aria-label*='New message']",
                        "button:has-text('New mail')",
                        "button:has-text('Nova mensagem')",
                        "[data-automation-id='newMessageButton']",
                        "div[role='button']:has-text('New')",
                    ]
                    
                    clicked = False
                    for selector in new_email_selectors:
                        try:
                            page.click(selector, timeout=2000)
                            clicked = True
                            msg = "✅ Botão 'Novo Email' clicado"
                            q.put(("status", msg))
                            print(msg)
                            break
                        except:
                            continue
                    
                    if not clicked:
                        msg = "⚠️ Não foi possível encontrar o botão 'Novo Email'. Clique manualmente."
                        q.put(("status", msg))
                        print(msg)
                        page.wait_for_timeout(5000)
                    else:
                        # Wait for compose window to open
                        page.wait_for_timeout(3000)
                        
                        # Try to fill in the TO field
                        msg = "📝 Preenchendo campo Para (TO)..."
                        q.put(("status", msg))
                        print(msg)
                        
                        # Conditional logic: Test mode vs Production mode
                        if test_mode:
                            # TEST MODE: Use emails from .env EMAIL_TESTE
                            test_emails_str = os.getenv("EMAIL_TESTE", "vpernarh@gmail.com")
                            to_emails_list = [email.strip() for email in test_emails_str.split(',') if email.strip()]
                            msg = f"🧪 MODO DE TESTE: Enviando para {', '.join(to_emails_list)}"
                            q.put(("status", msg))
                            print(msg)
                        else:
                            # PRODUCTION MODE: Use all real emails (transportadora + fornecedor)
                            to_emails_list = test_email['to_emails']
                            msg = f"📧 MODO PRODUÇÃO: {len(to_emails_list)} emails no campo TO"
                            q.put(("status", msg))
                            print(msg)
                        
                        to_filled = False
                        try:
                            # Find the Para div by aria-label
                            to_field = page.locator('div[aria-label="Para"]').first
                            to_field.click(timeout=3000)
                            page.wait_for_timeout(500)
                            
                            # Type each email and press Enter to create chip
                            for idx, email in enumerate(to_emails_list):
                                page.keyboard.type(email, delay=1)
                                page.wait_for_timeout(300)
                                page.keyboard.press("Enter")
                                page.wait_for_timeout(500)
                                # msg = f"   → Email {idx+1}/{len(to_emails_list)} adicionado: {email}"
                                # q.put(("status", msg))
                                # print(msg)
                            
                            msg = f"✅ Campo TO preenchido com {len(to_emails_list)} emails"
                            q.put(("status", msg))
                            # print(msg)
                            to_filled = True
                        except Exception as e:
                            error_msg = f"⚠️ Erro ao preencher TO: {str(e)[:80]}"
                            q.put(("status", error_msg))
                            print(error_msg)
                        
                        if not to_filled:
                            warning_msg = "⚠️ TO não preenchido automaticamente. Verifique os selectors."
                            q.put(("status", warning_msg))
                            print(warning_msg)
                        
                        # Conditional CC Field: Skip in test mode, fill in production mode
                        if test_mode:
                            # SKIP CC FIELD IN TEST MODE
                            msg = "🧪 MODO DE TESTE: Pulando campo CC (apenas testando TO e ENVIAR)"
                            q.put(("status", msg))
                            print(msg)
                        else:
                            # PRODUCTION MODE: Fill CC field with Stellantis emails
                            page.wait_for_timeout(1000)
                            msg = "📝 Preenchendo campo Cc..."
                            q.put(("status", msg))
                            print(msg)
                            
                            # CC emails are a semicolon-separated string
                            cc_emails_str = test_email['cc_emails']
                            cc_count = len([e for e in cc_emails_str.split(';') if e.strip()])
                            print(f"   Total de {cc_count} emails Stellantis para CC")
                            
                            # Click Cc button to show CC field if hidden
                            try:
                                page.get_by_text("Cc", exact=False).click(timeout=2000)
                                page.wait_for_timeout(500)
                                msg = "   → Campo Cc expandido"
                                q.put(("status", msg))
                                print(msg)
                            except:
                                msg = "   → Cc já visível ou botão não encontrado"
                                q.put(("status", msg))
                                print(msg)
                            
                            cc_filled = False
                            try:
                                # Find CC div by aria-label
                                cc_field = page.locator('div[aria-label="Cc"]').first
                                cc_field.click(timeout=2000)
                                page.wait_for_timeout(500)
                                
                                # Paste the entire semicolon-separated string at once
                                page.keyboard.type(cc_emails_str, delay=1)
                                page.wait_for_timeout(500)
                                
                                # Press Enter once to create all email chips
                                page.keyboard.press("Enter")
                                page.wait_for_timeout(1000)
                                
                                msg = f"✅ Campo CC preenchido com {cc_count} emails Stellantis"
                                q.put(("status", msg))
                                print(msg)
                                cc_filled = True
                            except Exception as e:
                                error_msg = f"⚠️ Erro ao preencher CC: {str(e)[:80]}"
                                q.put(("status", error_msg))
                                print(error_msg)
                            
                            if not cc_filled:
                                warning_msg = "⚠️ CC não preenchido automaticamente"
                                q.put(("status", warning_msg))
                                print(warning_msg)
                        
                        # Try to fill in Subject
                        page.wait_for_timeout(1000)
                        q.put(("status", "📝 Preenchendo assunto..."))
                        print("📝 Preenchendo assunto...")
                        
                        subject_selectors = [
                            "input[aria-label*='Assunto']",
                            "input[aria-label*='Subject']",
                            "[data-automation-id='subject-field'] input",
                            "input[placeholder*='Assunto']",
                        ]
                        
                        subject_filled = False
                        for selector in subject_selectors:
                            try:
                                page.fill(selector, test_email['subject'], timeout=2000)
                                msg = "✅ Assunto preenchido"
                                q.put(("status", msg))
                                print(msg)
                                subject_filled = True
                                break
                            except:
                                continue
                        
                        if not subject_filled:
                            warning_msg = "⚠️ Assunto não preenchido"
                            q.put(("status", warning_msg))
                            print(warning_msg)
                        
                        # Try to fill in Body with formatted HTML
                        page.wait_for_timeout(1000)
                        q.put(("status", "📝 Preenchendo corpo do email (HTML)..."))
                        print("📝 Preenchendo corpo do email (HTML)...")
                        
                        body_selectors = [
                            "div[aria-label*='Corpo da mensagem']",
                            "div[aria-label*='Message body']",
                            "[role='textbox'][aria-label*='mensagem']",
                            "div[contenteditable='true']",
                        ]
                        
                        body_filled = False
                        for selector in body_selectors:
                            try:
                                body_element = page.locator(selector).first
                                # Click to focus
                                body_element.click(timeout=2000)
                                page.wait_for_timeout(500)
                                
                                # Clear any existing content
                                page.keyboard.press("Control+A")
                                page.keyboard.press("Backspace")
                                
                                # Insert HTML content using JavaScript
                                escaped_html = test_email['content_html'].replace('`', '\\`').replace('$', '\\$')
                                body_element.evaluate(f"""
                                    (element) => {{
                                        element.innerHTML = `{escaped_html}`;
                                    }}
                                """)
                                
                                msg = "✅ Corpo do email preenchido com tabela formatada"
                                q.put(("status", msg))
                                print(msg)
                                body_filled = True
                                break
                            except Exception as e:
                                error_msg = f"⚠️ Tentativa de preencher body falhou: {str(e)[:50]}"
                                q.put(("status", error_msg))
                                print(error_msg)
                                continue
                        
                        if not body_filled:
                            warning_msg = "⚠️ Corpo do email não foi preenchido"
                            q.put(("status", warning_msg))
                            print(warning_msg)
                        
                        
                        if test_mode:
                            msg = "✅ Email de TESTE criado!"
                        else:
                            msg = "✅ Email de PRODUÇÃO criado!"
                        q.put(("status", msg))
                        print(msg)
                        
                        # Try to click SEND button
                        page.wait_for_timeout(2000)  # Wait 2 seconds before sending
                        msg = "📤 Tentando clicar no botão ENVIAR..."
                        q.put(("status", msg))
                        print(msg)
                        
                        send_selectors = [
                            "button[aria-label*='Enviar']",
                            "button[aria-label*='Send']",
                            "button:has-text('Enviar')",
                            "button:has-text('Send')",
                            "[data-automation-id='sendButton']",
                            "button[name='send']",
                        ]
                        
                        send_clicked = False
                        for selector in send_selectors:
                            try:
                                page.click(selector, timeout=2000)
                                send_clicked = True
                                msg = "✅ Botão ENVIAR clicado com sucesso!"
                                q.put(("status", msg))
                                print(msg)
                                break
                            except:
                                continue
                        
                        if not send_clicked:
                            msg = "⚠️ Não foi possível encontrar o botão ENVIAR automaticamente."
                            q.put(("status", msg))
                            print(msg)
                            msg = "🔍 Verifique a página - o email pode estar em rascunhos."
                            q.put(("status", msg))
                            print(msg)
                            msg = "⏸️ Aguardando 30 segundos para você verificar..."
                            q.put(("status", msg))
                            print(msg)
                            page.wait_for_timeout(30000)  # Wait 30 seconds if send failed
                        else:
                            # Wait to confirm send
                            page.wait_for_timeout(3000)
                            if test_mode:
                                test_emails_str = os.getenv("EMAIL_TESTE", "vpernarh@gmail.com")
                                msg = f"✅ Email de teste enviado! Verifique: {test_emails_str}"
                            else:
                                msg = f"✅ Email enviado para {test_email['supplier_name']}!"
                            q.put(("status", msg))
                            print(msg)
                            msg = "⏸️ Aguardando 10 segundos antes de fechar..."
                            q.put(("status", msg))
                            print(msg)
                            page.wait_for_timeout(10000)  # Keep browser open 10 seconds to verify
                        
                except Exception as e:
                    error_msg = f"⚠️ Erro ao criar email de teste: {e}"
                    q.put(("status", error_msg))
                    print(error_msg)
                    import traceback
                    print(traceback.format_exc())
                    msg = "🔍 Verifique o erro acima. Aguardando 30 segundos..."
                    q.put(("status", msg))
                    print(msg)
                    page.wait_for_timeout(30000)  # Wait 30 seconds on error
            
        if test_mode:
            test_emails_str = os.getenv("EMAIL_TESTE", "vpernarh@gmail.com")
            msg = f"🎉 TESTE CONCLUÍDO! Verifique {test_emails_str} para confirmar recebimento."
        else:
            msg = "🎉 PRODUÇÃO CONCLUÍDA! Emails enviados com sucesso."
        q.put(("status", msg))
        print(msg)
        q.put(("progress", 100))
        q.put(("done", True))
        
    except FileNotFoundError:
        q.put(("status", "❌ Erro: Arquivo '.env' não encontrado."))
        q.put(("done", True))
    except KeyError as e:
        q.put(("status", f"❌ Erro: Variável de ambiente inválida: {e}"))
        q.put(("done", True))
    except Exception as e:
        q.put(("status", f"❌ Erro inesperado: {e}"))
        traceback.print_exc()
        q.put(("done", True))
        

def main_process(q: queue.Queue, test_mode: bool = False):
    try:
        # Run with Playwright
        with sync_playwright() as playwright:
            run_automation(playwright, q, test_mode)  # Pass test_mode!
    except KeyboardInterrupt:
        q.put(("status", "⚠️ Processo interrompido pelo usuário."))
        q.put(("done", True))
    except Exception as e:
        q.put(("status", f"❌ Erro no processo principal: {e}"))
        q.put(("done", True))

# --- TKINTER APP SETUP ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Automação e Processamento RPA")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        
        # DHL & STELLANTIS Colors
        # DHL: Red (#FF0000), Yellow (#FFCC00)
        # STELLANTIS: Blue (#003DA5), Orange (#FF6600)
        dhl_red = "#FF0000"
        dhl_yellow = "#FFCC00"
        stellantis_blue = "#003DA5"
        stellantis_orange = "#FF6600"
        
        # Set modern color scheme with DHL/STELLANTIS theme
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure button style with STELLANTIS blue
        style.configure('TButton', background=stellantis_blue, foreground="white", relief="flat", padding=6)
        style.map('TButton', background=[('active', stellantis_orange)])
        
        # Configure progressbar with gradient effect using STELLANTIS colors
        style.configure('TProgressbar', background=stellantis_blue, troughcolor='#E8E8E8', bordercolor='#CCCCCC', lightcolor=stellantis_orange, darkcolor=stellantis_blue)
        
        # Configure labels with theme colors
        style.configure('Title.TLabel', font=("Segoe UI", 16, "bold"), foreground=stellantis_blue)
        
        self.queue = queue.Queue()
        self.download_var = tk.BooleanVar(value=True)  # Checkbox state: True = download enabled
        self.test_mode_var = tk.BooleanVar(value=False)  # Checkbox state: False = production, True = test mode

        # --- Main container ---
        container = tk.Frame(root, bg="white")
        container.pack(fill=tk.BOTH, expand=True)

        # --- Header with DHL/STELLANTIS accent ---
        header_frame = tk.Frame(container, bg=stellantis_blue, height=80)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Title section with colored background
        title_label = tk.Label(header_frame, text="🤖 Automação de Envio de Planejamento FIAPE", font=("Segoe UI", 16, "bold"), fg="white", bg=stellantis_blue)
        title_label.pack(anchor="w", padx=15, pady=(10, 2))
        
        subtitle_label = tk.Label(header_frame, text="Processamento Inteligente de Processos Manuais", font=("Segoe UI", 9), fg=dhl_yellow, bg=stellantis_blue)
        subtitle_label.pack(anchor="w", padx=15, pady=(0, 10))

        # --- Main content frame ---
        main_frame = ttk.Frame(container, padding="13")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Status section
        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar. Clique em 'Processar'.", font=("Segoe UI", 11), foreground=stellantis_blue)
        self.status_label.pack(pady=(2, 5), padx=1, fill=tk.X)

        # Progress bar with accent color
        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        # Button section with modern styling
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=4, fill=tk.X)
        
        self.process_button = ttk.Button(button_frame, text="▶ Processar", command=self.start_processing_thread, style='TButton')
        self.process_button.pack(side=tk.TOP, padx=5, pady=5)
        
        # Test mode checkbox
        test_mode_frame = ttk.Frame(button_frame)
        test_mode_frame.pack(side=tk.TOP, pady=5)
        
        self.test_mode_checkbox = ttk.Checkbutton(
            test_mode_frame,
            text="🧪 Modo de Teste (envia apenas para emails de teste)",
            variable=self.test_mode_var,
            command=self.update_mode_indicator,  # Update indicator when toggled
            style='TCheckbutton'
        )
        self.test_mode_checkbox.pack(side=tk.LEFT, padx=5)
        
        # Info label for test mode
        test_info_label = ttk.Label(
            test_mode_frame,
            text="← Marque para testar com EMAIL_TESTE do .env",
            font=("", 8),
            foreground="gray"
        )
        test_info_label.pack(side=tk.LEFT, padx=5)
        
        # Mode indicator label (shows current mode)
        mode_indicator_frame = ttk.Frame(button_frame)
        mode_indicator_frame.pack(side=tk.TOP, pady=(0, 5))
        
        self.mode_indicator_label = tk.Label(
            mode_indicator_frame,
            text="🚀 MODO ATUAL: PRODUÇÃO (todos os emails reais)",
            font=("Segoe UI", 10, "bold"),
            fg="white",
            bg="#FF6600",  # Orange for production
            padx=10,
            pady=5
        )
        self.mode_indicator_label.pack()
        
        # Log section with accent
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Atividades", padding="13")
        log_frame.pack(pady=0, padx=2, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=10, font=("Consolas", 11), bg="#F5F5F5", fg="#333333")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        
        # Footer section with DHL/STELLANTIS branding
        footer_frame = tk.Frame(container, bg=stellantis_blue, height=34)
        footer_frame.pack(fill=tk.X, padx=0, pady=0, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        # Left side - DHL -> STELLANTIS
        left_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        left_footer.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15, pady=10)
        
        # DHL Logo/Text (DHL Yellow)
        dhl_label = tk.Label(left_footer, text="🚚 DHL", font=("Segoe UI", 11, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        dhl_label.pack(side=tk.LEFT, padx=1)
        
        arrow_label = tk.Label(left_footer, text="→", font=("Segoe UI", 12, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        arrow_label.pack(side=tk.LEFT, padx=3)
        
        # STELLANTIS Logo/Text (STELLANTIS Orange accent)
        stellantis_label = tk.Label(left_footer, text="STELLANTIS 🏢", font=("Segoe UI", 11, "bold"), fg=stellantis_orange, bg=stellantis_blue)
        stellantis_label.pack(side=tk.LEFT, padx=3)
        
        # Right side - Developer credit
        right_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        right_footer.pack(side=tk.RIGHT, padx=15, pady=10)
        
        footer_label = tk.Label(right_footer, text="Desenvolvido por: Vincent Pernarh", font=("Segoe UI", 9), fg="white", bg=stellantis_blue)
        footer_label.pack(anchor="e")

    def toggle_download(self):
        """Toggle download mode on/off with visual feedback."""
        stellantis_blue = "#003DA5"
        stellantis_orange = "#FF6600"
        gray_disabled = "#888888"
        
        current_state = self.download_var.get()
        new_state = not current_state
        self.download_var.set(new_state)
        
        if new_state:
            # Download enabled
            self.download_button.config(
                text="✓ Baixar demandas",
                bg=stellantis_blue,
                relief=tk.FLAT
            )
        else:
            # Download disabled
            self.download_button.config(
                text="✗ Baixar demandas",
                bg=gray_disabled,
                relief=tk.FLAT
            )

    def update_mode_indicator(self):
        """Update the mode indicator label when checkbox is toggled."""
        test_mode = self.test_mode_var.get()
        
        if test_mode:
            # TEST MODE - Green background
            self.mode_indicator_label.config(
                text="🧪 MODO ATUAL: TESTE (apenas emails de teste)",
                bg="#28a745",  # Green
                fg="white"
            )
            print("✅ Modo alterado para: TESTE")
        else:
            # PRODUCTION MODE - Orange background
            self.mode_indicator_label.config(
                text="🚀 MODO ATUAL: PRODUÇÃO (todos os emails reais)",
                bg="#FF6600",  # Orange
                fg="white"
            )
            print("⚠️ Modo alterado para: PRODUÇÃO")



    def start_processing_thread(self):
        self.process_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        
        test_mode = self.test_mode_var.get()
        
        # EXPLICIT DEBUG: Show checkbox state
        print("=" * 70)
        print(f"DEBUG: Checkbox 'Modo de Teste' está: {'MARCADO ✓' if test_mode else 'DESMARCADO ✗'}")
        print(f"DEBUG: test_mode = {test_mode}")
        print("=" * 70)
        
        if test_mode:
            status_msg = "🧪 Iniciando processo em MODO DE TESTE..."
            self.status_label.config(text=status_msg)
            print(status_msg)
        else:
            status_msg = "🚀 Iniciando processo em MODO PRODUÇÃO..."
            self.status_label.config(text=status_msg)
            print(status_msg)
        
        self.thread = threading.Thread(target=main_process, args=(self.queue, test_mode))
        self.thread.daemon = True
        self.thread.start()
        
        # Start checking the queue for updates
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text, self.process_button)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
