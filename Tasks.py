import email
import json
import os

import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Playwright, TimeoutError, expect
from datetime import date, timedelta
import traceback
import getpass
import sys
# from App import load_config

if getattr(sys, 'frozen', False):
    # Quando rodando como .exe (PyInstaller)
    caminho_base = os.path.dirname(sys.executable)
else:
    # Quando rodando como script .py
    caminho_base = os.path.dirname(os.path.abspath(__file__))

df_geral = pd.DataFrame()
df_transportadoras = pd.DataFrame()
df_stellantis = pd.DataFrame()
df_fornecedores = pd.DataFrame()
df_Planejado = pd.DataFrame()



def process_planejamento(q=None):
    planejamento_path = os.path.join(caminho_base, "Base","Planejamento")
    Dados_email_path = os.path.join(caminho_base, "Base","Dados_email")
    
    
    try:
        if q:
            msg = "📂 Iniciando processamento de planejamento..."
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 15))
       
        current_date = date.today()
        weeknumber = current_date.isocalendar()[1]
        
        if q:
            msg = f"📅 Processando arquivos para a semana {weeknumber}..."
            q.put(("status", msg))
            print(msg)
       
        # Load email data files
        if q:
            msg = "📧 Carregando dados de emails..."
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 20))
            
        for file in os.listdir(Dados_email_path):
            if file.endswith(".xlsx") and  "email" in file.lower() and not file.startswith("~$"):
                email_path = os.path.join(Dados_email_path, file)
                df_geral = pd.read_excel(email_path, sheet_name='GERAL')
                df_transportadoras = pd.read_excel(email_path, sheet_name='TRANSPORTADORAS')
                df_stellantis = pd.read_excel(email_path, sheet_name='STELLANTIS')
                df_fornecedores = pd.read_excel(email_path, sheet_name='FORNECEDORES')
                if q:
                    msg = f"✅ Arquivo de emails carregado: {file}"
                    q.put(("status", msg))
                    print(msg)
            
        # Load planning file
        if q:
            msg = "📋 Carregando arquivo de planejamento..."
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 30))
            
        for file in os.listdir(planejamento_path):
            if file.endswith(".xlsx") and  "diretos" in file.lower() and f"w{weeknumber}" in file.lower() and not file.startswith("~$"):
                planejamento_file = os.path.join(planejamento_path, file)
                
                # First, read all columns to check what exists
                df_Planejado_temp = pd.read_excel(planejamento_file, sheet_name='Planejado', nrows=0, header=1)
                available_columns = df_Planejado_temp.columns.tolist()
                
                # if q:
                #     q.put(("status", f"📋 Colunas disponíveis: {', '.join(str(col) for col in available_columns)}"))
                
                # Define the columns we want to read
                desired_columns = ['SAP', 'Fornecedor', 'Veículos', 'Destino', 'Semana', 'Dia', 'Data Planejada de Coleta', 'Hora Planejada de Coleta']
                
                # Check which columns exist (case-insensitive and strip spaces)
                columns_to_read = []
                for desired_col in desired_columns:
                    for available_col in available_columns:
                        if str(desired_col).strip().lower() == str(available_col).strip().lower():
                            columns_to_read.append(available_col)
                            break
                
                if q:
                    q.put(("status", f"📋 Colunas mapeadas: {len(columns_to_read)}/{len(desired_columns)} - {', '.join(str(col) for col in columns_to_read)}"))
                
                if len(columns_to_read) == 0:
                    if q:
                        q.put(("status", "⚠️ Nenhuma coluna esperada foi encontrada, carregando todas as colunas..."))
                    df_Planejado = pd.read_excel(planejamento_file, sheet_name='Planejado', header=1)
                else:
                    # Use header=1 to match the columns we found
                    df_Planejado = pd.read_excel(planejamento_file, sheet_name='Planejado', usecols=columns_to_read, header=1)
                    
                if q:
                    msg = f"✅ Planejamento carregado: {file} ({len(df_Planejado)} linhas)"
                    q.put(("status", msg))
                    print(msg)
        
        if q:
            msg = "🔄 Processando mapeamentos de emails..."
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 40))
        
        # Process the data and prepare emails
        email_list = treat_data(df_geral, df_transportadoras, df_stellantis, df_fornecedores, df_Planejado, weeknumber, q)
        
        return email_list
        
    except Exception as e:
        if q:
            msg = f"❌ Ocorreu um erro: {e}"
            q.put(("status", msg))
            print(msg)
        else:
            print(f"❌ Ocorreu um erro: {e}")
        traceback.print_exc()
        return None




def treat_data(df_geral, df_transportadoras, df_stellantis, df_fornecedores, df_Planejado, weeknumber, q=None):
    """
    Process planning data and prepare email structures for each supplier.
    
    Steps:
    1. Group planning data by SAP + Fornecedor
    2. Map transportadora for each group
    3. Get transportadora emails (TO)
    4. Get fornecedor emails (TO)
    5. Get ALL Stellantis emails from CONTATOS column (CC) - returned as semicolon-separated string
    6. Format email content as HTML table
    7. Create email structure dictionary
    """
    
    try:
        if q:
            msg = "📊 Agrupando dados de planejamento por SAP + Fornecedor..."
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 50))
        
        # Step 1: Group planning data by SAP + Fornecedor
        if df_Planejado.empty:
            if q:
                msg = "⚠️ Nenhum dado de planejamento encontrado!"
                q.put(("status", msg))
                print(msg)
            return
        
        # Create a grouping key combining SAP and Fornecedor
        grouped = df_Planejado.groupby(['SAP', 'Fornecedor'])
        
        total_groups = len(grouped)
        if q:
            msg = f"✅ Total de {total_groups} grupos (fornecedores) encontrados"
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 55))
        
        # List to store all email structures
        email_list = []
        
        # Variable to store Stellantis emails (fetched once, used for all emails)
        stellantis_emails_str = ""
        
        # Step 2-6: Process each group
        for idx, ((sap_code, fornecedor_name), group_data) in enumerate(grouped, 1):
            if q:
                msg = f"🔍 Processando {idx}/{total_groups}: {fornecedor_name} (SAP: {sap_code})"
                q.put(("status", msg))
                print(msg)
            
            email_info = {
                'supplier_name': fornecedor_name,
                'sap_code': sap_code,
                'to_emails': [],
                'cc_emails': '',  # String with all Stellantis emails separated by semicolons
                'subject': '',
                'content_html': '',
                'planning_data': group_data
            }
            
            # Step 2: Map Transportadora from df_geral
            transportadora_name = None
            try:
                # Match Fornecedor in df_geral column C and get TRANSP from column B
                # Assuming column names: 'Fornecedor' (C) and 'TRANSP' (B)
                mask = df_geral['Fornecedor'].str.strip().str.upper() == fornecedor_name.strip().upper()
                if mask.any():
                    transportadora_name = df_geral.loc[mask, 'TRANSP'].iloc[0]
                    if q:
                        msg = f"   ✓ Transportadora mapeada: {transportadora_name}"
                        q.put(("status", msg))
                        print(msg)
                else:
                    if q:
                        msg = f"   ⚠️ Fornecedor '{fornecedor_name}' não encontrado na planilha GERAL"
                        q.put(("status", msg))
                        print(msg)
            except Exception as e:
                if q:
                    msg = f"   ⚠️ Erro ao mapear transportadora: {e}"
                    q.put(("status", msg))
                    print(msg)
            
            # Step 3: Get Transportadora emails (TO)
            if transportadora_name:
                try:
                    mask = df_transportadoras['TRANSPORTADORAS'].str.strip().str.upper() == transportadora_name.strip().upper()
                    if mask.any():
                        contatos = df_transportadoras.loc[mask, 'CONTATOS'].iloc[0]
                        if pd.notna(contatos):
                            # Split by both semicolon and newline, then clean up
                            contatos_str = str(contatos)
                            # Replace newlines with semicolons first
                            contatos_str = contatos_str.replace('\n', ';').replace('\r', ';')
                            # Split by semicolon and clean up whitespace and empty strings
                            transp_emails = [email.strip() for email in contatos_str.split(';') if email.strip()]
                            # Remove duplicates while preserving order
                            seen = set()
                            transp_emails_unique = []
                            for email in transp_emails:
                                if email not in seen and '@' in email:  # Ensure it's a valid email
                                    seen.add(email)
                                    transp_emails_unique.append(email)
                            
                            email_info['to_emails'].extend(transp_emails_unique)
                            if q:
                                msg = f"   ✓ {len(transp_emails_unique)} email(s) da transportadora adicionados"
                                q.put(("status", msg))
                                print(msg)
                        else:
                            if q:
                                msg = f"   ⚠️ CONTATOS vazio para transportadora {transportadora_name}"
                                q.put(("status", msg))
                                print(msg)
                    else:
                        if q:
                            msg = f"   ⚠️ Transportadora '{transportadora_name}' não encontrada na planilha TRANSPORTADORAS"
                            q.put(("status", msg))
                            print(msg)
                except Exception as e:
                    if q:
                        msg = f"   ⚠️ Erro ao buscar emails da transportadora: {e}"
                        q.put(("status", msg))
                        print(msg)
            
            # Step 4: Get Fornecedor emails (TO)
            try:
                mask = df_fornecedores['Fornecedor'].str.strip().str.upper() == fornecedor_name.strip().upper()
                if mask.any():
                    fornec_emails_str = df_fornecedores.loc[mask, 'EMAILS'].iloc[0]
                    if pd.notna(fornec_emails_str):
                        # Split by both semicolon and newline, then clean up
                        fornec_str = str(fornec_emails_str)
                        # Replace newlines with semicolons first
                        fornec_str = fornec_str.replace('\n', ';').replace('\r', ';')
                        # Split by semicolon and clean up whitespace and empty strings
                        fornec_emails = [email.strip() for email in fornec_str.split(';') if email.strip()]
                        # Remove duplicates and validate emails
                        seen = set()
                        fornec_emails_unique = []
                        for email in fornec_emails:
                            if email not in seen and '@' in email:  # Ensure it's a valid email
                                seen.add(email)
                                fornec_emails_unique.append(email)
                        
                        email_info['to_emails'].extend(fornec_emails_unique)
                        if q:
                            msg = f"   ✓ {len(fornec_emails_unique)} email(s) do fornecedor adicionados"
                            q.put(("status", msg))
                            print(msg)
                    else:
                        if q:
                            msg = f"   ⚠️ EMAILS vazio para fornecedor {fornecedor_name}"
                            q.put(("status", msg))
                            print(msg)
                else:
                    if q:
                        msg = f"   ⚠️ Fornecedor '{fornecedor_name}' não encontrado na planilha FORNECEDORES"
                        q.put(("status", msg))
                        print(msg)
            except Exception as e:
                if q:
                    msg = f"   ⚠️ Erro ao buscar emails do fornecedor: {e}"
                    q.put(("status", msg))
                    print(msg)
            
            # Remove duplicates from TO emails
            email_info['to_emails'] = list(set(email_info['to_emails']))
            
            if q:
                msg = f"   📧 Total de {len(email_info['to_emails'])} email(s) únicos no campo TO"
                q.put(("status", msg))
                print(msg)
            
            # Step 5: Get Stellantis emails (CC) - same for all emails
            if idx == 1:  # Only fetch once
                try:
                    # Get ALL non-empty values from CONTATOS column (column B)
                    stellantis_contacts = df_stellantis['CONTATOS'].dropna().tolist()
                    
                    # Collect all emails into a single list
                    stellantis_emails_list = []
                    for contact in stellantis_contacts:
                        # Each cell might have multiple emails separated by semicolon or newline
                        contact_str = str(contact)
                        # Replace newlines with semicolons first
                        contact_str = contact_str.replace('\n', ';').replace('\r', ';')
                        # Split by semicolon and clean up
                        emails = [email.strip() for email in contact_str.split(';') if email.strip() and '@' in email]
                        stellantis_emails_list.extend(emails)
                    
                    # Remove duplicates while preserving order
                    seen = set()
                    stellantis_emails_unique = []
                    for email in stellantis_emails_list:
                        if email not in seen:
                            seen.add(email)
                            stellantis_emails_unique.append(email)
                    
                    # Join all emails into a single semicolon-separated string
                    stellantis_emails_str = "; ".join(stellantis_emails_unique)
                    
                    if q:
                        msg = f"✅ {len(stellantis_emails_unique)} email(s) Stellantis para CC"
                        q.put(("status", msg))
                        print(msg)
                except Exception as e:
                    stellantis_emails_str = ""
                    if q:
                        msg = f"   ⚠️ Erro ao buscar emails Stellantis: {e}"
                        q.put(("status", msg))
                        print(msg)
            
            email_info['cc_emails'] = stellantis_emails_str
            
            # Step 6: Create email subject
            email_info['subject'] = f"Planejamento de Coletas - Semana {weeknumber} - {fornecedor_name}"
            
            # Step 7: Format email content as HTML table
            email_info['content_html'] = create_email_html_content(group_data, fornecedor_name, sap_code, weeknumber)
            
            # Add to email list
            email_list.append(email_info)
            
            # Update progress
            progress = 55 + int((idx / total_groups) * 30)
            if q:
                q.put(("progress", progress))
        
        if q:
            msg = "✅ Mapeamento de emails concluído!"
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 85))
        
        # Step 8: Display summary for validation
        # display_email_summary(email_list, q)
        
        if q:
            msg = "⏸️ Validação necessária antes de enviar emails"
            q.put(("status", msg))
            print(msg)
            q.put(("progress", 90))
        
        # Return the email list for future use (when sending emails)
        return email_list
        
    except Exception as e:
        if q:
            msg = f"❌ Erro ao processar dados: {e}"
            q.put(("status", msg))
            print(msg)
        traceback.print_exc()
        return None


def create_email_html_content(group_data, fornecedor_name, sap_code, weeknumber):
    """
    Convert planning data to HTML table format for email content with inline styles for Outlook compatibility.
    """
    # Build HTML with inline styles (required for Outlook Web)
    html = f"""<div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; padding: 10px;">
    <h2 style="color: #003DA5; font-size: 18px; margin-bottom: 15px;">Segue plano de coletas previsto para semana {weeknumber}</h2>
    <p style="margin: 8px 0;"><strong>Fornecedor:</strong> {fornecedor_name}</p>
    <p style="margin: 8px 0;"><strong>SAP:</strong> {sap_code}</p>
    <p style="margin: 15px 0; padding: 10px; background-color: #fff3cd; border-left: 4px solid #ffc107;"><strong>Obs:</strong> o plano abaixo é apenas uma prévia, podendo ser alterado de acordo com a necessidade Stellantis.</p>
    
    <table style="border-collapse: collapse; width: 100%; margin: 20px 0; border: 1px solid #ddd;">
        <thead>
            <tr style="background-color: #003DA5;">
                <th style="background-color: #003DA5; color: white; padding: 12px 8px; text-align: left; font-weight: bold; border: 1px solid #003DA5;">Fornecedor</th>
                <th style="background-color: #003DA5; color: white; padding: 12px 8px; text-align: left; font-weight: bold; border: 1px solid #003DA5;">Veículos</th>
                <th style="background-color: #003DA5; color: white; padding: 12px 8px; text-align: left; font-weight: bold; border: 1px solid #003DA5;">Destino</th>
                <th style="background-color: #003DA5; color: white; padding: 12px 8px; text-align: left; font-weight: bold; border: 1px solid #003DA5;">Semana</th>
                <th style="background-color: #003DA5; color: white; padding: 12px 8px; text-align: left; font-weight: bold; border: 1px solid #003DA5;">Dia</th>
                <th style="background-color: #003DA5; color: white; padding: 12px 8px; text-align: left; font-weight: bold; border: 1px solid #003DA5;">Data Planejada de Coleta</th>
                <th style="background-color: #003DA5; color: white; padding: 12px 8px; text-align: left; font-weight: bold; border: 1px solid #003DA5;">Hora Planejada de Coleta</th>
            </tr>
        </thead>
        <tbody>"""
    
    # Add data rows with alternating colors
    for idx, (_, row) in enumerate(group_data.iterrows()):
        bg_color = "#f9f9f9" if idx % 2 == 0 else "#ffffff"
        html += f'<tr style="background-color: {bg_color};">'
        html += f'<td style="padding: 10px 8px; border: 1px solid #ddd;">{row["Fornecedor"]}</td>'
        html += f'<td style="padding: 10px 8px; border: 1px solid #ddd;">{row["Veículos"]}</td>'
        html += f'<td style="padding: 10px 8px; border: 1px solid #ddd;">{row["Destino"]}</td>'
        html += f'<td style="padding: 10px 8px; border: 1px solid #ddd; text-align: center;">{row["Semana"]}</td>'
        html += f'<td style="padding: 10px 8px; border: 1px solid #ddd;">{row["Dia"]}</td>'
        
        # Format date
        data_coleta = row['Data Planejada de Coleta']
        if pd.notna(data_coleta):
            if isinstance(data_coleta, datetime):
                data_coleta = data_coleta.strftime('%d/%m/%Y')
            elif isinstance(data_coleta, str):
                data_coleta = data_coleta
        else:
            data_coleta = "-"
        html += f'<td style="padding: 10px 8px; border: 1px solid #ddd; text-align: center;">{data_coleta}</td>'
        
        # Format time
        hora_coleta = row['Hora Planejada de Coleta']
        if pd.notna(hora_coleta):
            if isinstance(hora_coleta, datetime):
                hora_coleta = hora_coleta.strftime('%H:%M')
            else:
                hora_coleta = str(hora_coleta)
        else:
            hora_coleta = "-"
        html += f'<td style="padding: 10px 8px; border: 1px solid #ddd; text-align: center;">{hora_coleta}</td>'
        html += "</tr>"
    
    html += """
        </tbody>
    </table>
    
    <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 12px; color: #666;">
        <p style="margin: 5px 0;">Este é um email automatizado. Para dúvidas, entre em contato com a equipe de planejamento.</p>
    </div>
</div>"""
    
    return html


def display_email_summary(email_list, q=None):
    """
    Display a summary of emails to be sent for validation.
    """
    if q:
        separator = "\n" + "="*70
        header = "📧 RESUMO DOS EMAILS A SEREM ENVIADOS"
        q.put(("status", separator))
        print(separator)
        q.put(("status", header))
        print(header)
        q.put(("status", "="*70))
        print("="*70)
    
    for idx, email_info in enumerate(email_list, 1):
        if q:
            email_header = f"\n✉️ Email {idx} de {len(email_list)}:"
            q.put(("status", email_header))
            print(email_header)
            
            fornecedor = f"   Fornecedor: {email_info['supplier_name']}"
            q.put(("status", fornecedor))
            print(fornecedor)
            
            sap = f"   SAP: {email_info['sap_code']}"
            q.put(("status", sap))
            print(sap)
            
            assunto = f"   Assunto: {email_info['subject']}"
            q.put(("status", assunto))
            print(assunto)
            
            to_emails = f"   Para (TO): {', '.join(email_info['to_emails']) if email_info['to_emails'] else 'NENHUM EMAIL ENCONTRADO'}"
            q.put(("status", to_emails))
            print(to_emails)
            
            # CC is now a string, count emails by splitting
            cc_count = len([e for e in email_info['cc_emails'].split(';') if e.strip()]) if email_info['cc_emails'] else 0
            cc_info = f"   CC: {cc_count} email(s) Stellantis"
            q.put(("status", cc_info))
            print(cc_info)
            
            planning = f"   Linhas de planejamento: {len(email_info['planning_data'])}"
            q.put(("status", planning))
            print(planning)
    
    if q:
        end_sep = "\n" + "="*70
        q.put(("status", end_sep))
        print(end_sep)
        total = f"✅ Total: {len(email_list)} emails preparados"
        q.put(("status", total))
        print(total)
        end_line = "="*70 + "\n"
        q.put(("status", end_line))
        print(end_line)