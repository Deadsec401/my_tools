##########################################################################################
import customtkinter as ctk
import os,tkinter,time,requests,webbrowser,random
from PIL import Image
from tkextrafont import Font
import threading,hashlib,datetime,base64,re
from email.message import EmailMessage
from tkinter import ttk
import smtplib
from CTkToolTip import *
import pandas as pd
import pdfkit
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from base64 import urlsafe_b64decode, urlsafe_b64encode
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from tkinter import filedialog
import uuid
##########################################################################################
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")
##########################################################################################
blue = '#333AA0'
black = '#0F100F'
dark = '#000000'
yellow = '#F8FA7C'
white='#B4B5BF'
gray='#0F100F'
green = '#00e3aa'
##########################################################################################
SCOPES = ['https://mail.google.com/']
stop_flag = threading.Event()
file_path = os.path.dirname(os.path.realpath(__file__))

##########################################################################################
exit_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/exit.png"),
                                 dark_image=Image.open(file_path + "/images/exit.png"),
                                 size=(20, 20))
home_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/home.png"),
                                 dark_image=Image.open(file_path + "/images/home.png"),
                                 size=(22, 22))
about_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/about.png"),
                                 dark_image=Image.open(file_path + "/images/about.png"),
                                 size=(23, 23))
mails_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/leads.png"),
                                 dark_image=Image.open(file_path + "/images/leads.png"),
                                 size=(20, 20))
csv_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/csv.png"),
                                 dark_image=Image.open(file_path + "/images/csv.png"),
                                 size=(20, 20))
logs_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/logs.png"),
                                 dark_image=Image.open(file_path + "/images/logs.png"),
                                 size=(20, 20))
auth_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/auth.png"),
                                 dark_image=Image.open(file_path + "/images/auth.png"),
                                 size=(20, 20))
att_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/attachment.png"),
                                 dark_image=Image.open(file_path + "/images/attachment.png"),
                                 size=(20, 20))
pdf_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/pdf.png"),
                                 dark_image=Image.open(file_path + "/images/pdf.png"),
                                 size=(20, 20))
sent_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/sent.png"),
                                 dark_image=Image.open(file_path + "/images/sent.png"),
                                 size=(20, 20))
failed_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/failed.png"),
                                 dark_image=Image.open(file_path + "/images/failed.png"),
                                 size=(20, 20))
authb_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/aut.png"),
                                 dark_image=Image.open(file_path + "/images/aut.png"),
                                 size=(20, 20))
impcsv_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/import_csv.png"),
                                 dark_image=Image.open(file_path + "/images/import_csv.png"),
                                 size=(20, 20))
import_attachment_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/add_attachment.png"),
                                 dark_image=Image.open(file_path + "/images/add_attachment.png"),
                                 size=(20, 20))
import_letter_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/import_letter.png"),
                                 dark_image=Image.open(file_path + "/images/import_letter.png"),
                                 size=(20, 20))
start_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/start.png"),
                                 dark_image=Image.open(file_path + "/images/start.png"),
                                 size=(20, 20))
stop_img = ctk.CTkImage(light_image=Image.open(file_path + "/images/stop.png"),
                                 dark_image=Image.open(file_path + "/images/stop.png"),
                                 size=(20, 20))
###################################################################################################################
def check_creds():
    creds_path = os.path.join(file_path, 'assets', 'data')
    authed_path = os.path.join(file_path, 'assets', 'data', 'tokens')
    total1 = len([item for item in os.listdir(creds_path) if os.path.isfile(os.path.join(creds_path, item))])
    total2 = len([item for item in os.listdir(authed_path) if os.path.isfile(os.path.join(authed_path, item))])
    total_logs.configure(text=total1)
    total_auth.configure(text=total2)

##################################################################################################################
def authorize_users():
    try:
        folder_path = os.path.join(file_path, 'assets', 'data')
        file_names = [item for item in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, item))]
        anum = len(file_names)
        print(anum)
        for i in range(anum):
            creds = None
            if os.path.exists(file_path + r'\assets\data\tokens\token{}.pickle'.format(i)):
                with open(file_path + r'\assets\data\tokens\token{}.pickle'.format(i), "rb") as token:
                    creds = pickle.load(token)
                    continue
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                    
                    continue
                else:
                    authorize_outbox.configure(state='normal')
                    authorize_outbox.insert("end",f"########################################\n")
                    authorize_outbox.insert("end",f"Waiting for user to authorize account cred{i}.json\n")
                    authorize_outbox.see('end')
                    authorize_outbox.update_idletasks()
                    authorize_outbox.insert("end",f"########################################\n")
                    authorize_outbox.configure(state='disabled')
                    flow = InstalledAppFlow.from_client_secrets_file(file_path + r'\assets\data\cred{}.json'.format(i), SCOPES)
                    creds = flow.run_local_server(port=0)
                with open(file_path + r'\assets\data\tokens\token{}.pickle'.format(i), "wb") as token:
                    pickle.dump(creds, token)
                    authorize_outbox.configure(state='normal')
                    authorize_outbox.insert("end",f"########################################\n")
                    authorize_outbox.insert("end",f"cred{i}.json have been authorized\n")
                    authorize_outbox.see('end')
                    authorize_outbox.update_idletasks()
                    authorize_outbox.insert("end",f"########################################\n")
                    authorize_outbox.configure(state='disabled')
                    check_creds()
                    continue
            
            return build('gmail', 'v1', credentials=creds)
    except Exception as e:
        authorize_outbox.configure(state='normal')
        authorize_outbox.insert("end",f"creed{i}.json failed to authorize\nReason :{e}\n")
        authorize_outbox.see('end')
        authorize_outbox.update_idletasks()
        authorize_outbox.insert("end",f"########################################\n")
        authorize_outbox.configure(state='disabled')
        
def start_authenticator():
    creds_path = os.path.join(file_path, 'assets', 'data')
    authed_path = os.path.join(file_path, 'assets', 'data', 'tokens')
    total1 = len([item for item in os.listdir(creds_path) if os.path.isfile(os.path.join(creds_path, item))])
    total2 = len([item for item in os.listdir(authed_path) if os.path.isfile(os.path.join(authed_path, item))])
    if total1 == total2:
        authorize_outbox.configure(state='normal')
        authorize_outbox.insert("end",f"########################################\n")
        authorize_outbox.insert("end",f"There's No credentials to authorize\nReason: All {total1} credentials have been authorized...\n")
        authorize_outbox.see('end')
        authorize_outbox.update_idletasks()
        authorize_outbox.configure(state='disabled')
    elif total1 == 0:
        authorize_outbox.configure(state='normal')
        authorize_outbox.insert("end",f"########################################\n")
        authorize_outbox.insert("end",f"There's No credentials to authorize\nReason: {total1} credentials found...\n")
        authorize_outbox.see('end')
        authorize_outbox.update_idletasks()
        authorize_outbox.configure(state='disabled')
    else:
        a = threading.Thread(target=authorize_users,daemon=True)
        a.start() 
##################################################################################################################
def open_attachment():
    global attachment_path_full
    try:
        f = []
        attachment_path.configure(state='normal')
        attachment_path.delete(0, 'end')
        attachment_path_full = filedialog.askopenfilename(initialdir="/", title="Select your html attachment", filetypes=(("Text files", "*.html"), ("All files", "*.*")))
        fname = attachment_path_full.split("/")[-1]
        f.append(attachment_path_full)
        attachment_path.insert(0, fname)
        print(attachment_path_full)
        attachment_path.configure(state='disabled')
        total_att.configure(text=len(f))
    except:
        pass

def open_letter():
    global letter_path_full
    try:
        letter_path
        letter_path.configure(state='normal')
        letter_path.delete(0, 'end')
        letter_path_full = filedialog.askopenfilename(initialdir="/", title="Select your letter", filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
        fname = letter_path_full.split("/")[-1]
        letter_path.insert(0, fname)
        print(letter_path_full)
        letter_path.configure(state='disabled')
    except:
        pass

def open_names():
    global names_path_full
    try:
        names_path.configure(state='normal')
        names_path.delete(0, 'end')
        names_path_full = filedialog.askopenfilename(initialdir="/", title="Select your from names", filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
        fname = names_path_full.split("/")[-1]
        names_path.insert(0, fname)
        print(names_path_full)
        names_path.configure(state='disabled')
    except:
        pass

def open_csv():
    global email_addresses
    global csv_path_full
    try:
        j=[]
        email_addresses = []
        csv_path.configure(state='normal')
        csv_path.delete(0, 'end')
        csv_path_full = filedialog.askopenfilename(initialdir="/", title="Select your from names [Csv only]", filetypes=(("Text files", "*.csv"), ("All files", "*.*")))
        fname = csv_path_full.split("/")[-1]
        csv_path.insert(0, fname)
        df = pd.read_csv(csv_path_full, encoding='ISO-8859-1')
        for idx,row in df.iterrows():
            email = row.get("Email", "")
            email_addresses.append(email)
            
        j.append(fname)
        total_mails.configure(text=len(email_addresses))
        total_csv.configure(text=len(j))
        csv_path.configure(state='disabled')
    except Exception as e:
        print(e)

##################################################################################################################
def gmail_authenticate(i):
    global creds
    try:
        ######### outbox ###########
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"Using Account Number Cred{i}\n")
        outbox.insert("end",f"########################################\n")
        # outbox.insert("end",f"########################################\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.configure(state='disabled')
        ######### outbox ###########
        creds = None
        if os.path.exists(file_path + r'\assets\data\tokens\token{}.pickle'.format(i)):
            with open(file_path + r'\assets\data\tokens\token{}.pickle'.format(i), "rb") as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(file_path + r'\assets\data\cred{}.json'.format(i), SCOPES)
                flow.authorization_url()
                creds = flow.run_local_server(port=0)
            with open(file_path + r'\assets\data\tokens\token{}.pickle'.format(i), "wb") as token:
                pickle.dump(creds, token)
       
        return build('gmail', 'v1', credentials=creds)
    except Exception as e:
        ######### outbox ############
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"Error occurred while authenticating with Gmail API.\nRerying with next account\nReason >>{e.args}\n")
        outbox.insert("end",f"########################################\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.configure(state='disabled')
        return None
        ######## outbox ###########




def add_attachment(message, filename):
    try:
        with open(filename, 'rb') as attachment_file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(filename)}"')
            message.attach(part)
    except Exception as e:
        ######### outbox ############
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"An error occoured (Attachment)\nReason >>{e}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')
        ######### outbox ###########

        
def build_message(destination, obj, body, attachments=[]):
    sender_email1 = from_email
    try:
        with open(names_path_full) as gg:
            sender_name1 = random.choice([name.strip() for name in gg])
        if not attachments:
            message = MIMEText(body)
            message['to'] = destination
            message['from'] = f'{sender_name1} <{sender_email1}>'
            message['subject'] = obj
        else:
            message = MIMEMultipart()
            message['to'] = destination
            message['from'] = f'{sender_name1} <{sender_email1}>'
            message['subject'] = obj
            message.attach(MIMEText(body))
            for filename in attachments:
                add_attachment(message, filename)
        return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}
    except Exception as e:
        ######### outbox ############
        outbox.configure(state='normal')
        outbox.insert("end",f"\n\n########################################\n")
        outbox.insert("end",f"An error occoured\nReason >>{e}\n")
        outbox.insert("end",f"########################################\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.configure(state='disabled')
        ######### outbox ###########

sent=0
failed = 0

def send_message(service, destination, obj,body, attachments=[]):
    global sent ,failed
    try:
        a = service.users().messages().send(
        userId="me",
        body=build_message(destination, obj ,body, attachments)
        ).execute()
        try:
            user_info = service.users().getProfile(userId='me').execute()
            from_email = (user_info['emailAddress'])
        except:
            from_email = 'Not fount'
        if 'SENT' in str(a):
            sent += 1
            ######### outbox ############
            outbox.configure(state='normal')
            outbox.insert("end",f"########################################\n")
            outbox.insert("end",f"Email sent successfully\nSent to >> {destination}\nSent with {from_email}\n")
            outbox.insert("end",f"########################################\n")
            outbox.see('end')
            outbox.update_idletasks()
            outbox.configure(state='disabled')
            total_sent.configure(text=sent)
            ######### outbox ###########
        else:
            failed +=1
            outbox.configure(state='normal')
            outbox.insert("end",f"########################################\n\n")
            outbox.insert("end",f"Failed to send email\nSent to >> {destination}\nReason >> {a}\n")
            outbox.insert("end",f"########################################\n")
            outbox.see('end')
            outbox.update_idletasks()
            outbox.configure(state='disabled')
            total_failed.configure(text=failed)
    except HttpError as e:
        failed +=1
        error_code = e.resp.status
        error_message = e._get_reason()
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"Failed to send email\nSent to >> {destination}\nStatus Code >> {error_code}\nReason >> {error_message}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')
        total_failed.configure(text=failed)
    except Exception as e:
        failed +=1
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"Failed to send email\nSent to >> {destination}\nReason >> {e}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')
        total_failed.configure(text=failed)

def replace_tags_with_data(content, data_row):
    for header_name in data_row.index:
        content = content.replace(f"{{{header_name}}}", str(data_row[header_name]))
    return content


pdf = 0
def convert_html_to_pdf(input_file_path, output_file_path ,pdf_name):
    global pdf
    try:
        options = {
            'page-size': "A5",
            'margin-top': '10mm',
            'margin-right': '0mm',
            'margin-bottom': '0mm',
            'margin-left': '0mm',
            'footer-center': 'Page [page] of [topage]',
        }
        
        with open(input_file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        
        # pdfkit.from_file(input_file_path, output_file_path, configuration=pdfkit.configuration(wkhtmltopdf=file_path + r'\assets\dont_delete_me\wkhtmltopdf.exe'),options=options)
        pdfkit.from_string(html_content, output_file_path, configuration=pdfkit.configuration(wkhtmltopdf=file_path + r'\assets\dont_delete_me\wkhtmltopdf.exe'), options=options)
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"Pdf conversion seccessfully\nName >> {pdf_name}\nType >> {sheet_type_var.get()}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')
        pdf +=1
        total_pdf.configure(text=pdf)
        time.sleep(1)
    except Exception as e:
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"Pdf conversion failed\nName >> {pdf_name}\nError Message >> {e}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')
        time.sleep(1)

progress_file_path = (file_path + r'\assets\logs_details\status_logs.txt')
def load_progress():
    try:
        if os.path.exists(progress_file_path):
            with open(progress_file_path, 'r') as progress_file:
                progress_data = progress_file.read().strip().split(',')
                if len(progress_data) == 2:
                    total_sent_emails, sent_per_account = map(int, progress_data)
                    return total_sent_emails, sent_per_account
        return 0, 0
    except Exception as e:
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"An error occoured\nReason >> {e}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')
        return 0, 0

def save_progress(total_sent_emails, sent_per_account):
    try:
        with open(progress_file_path, 'w') as progress_file:
            progress_file.write(f"{total_sent_emails},{sent_per_account}")
    except Exception as e:
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"\nAn error occoured\nReason >> {e}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')

# def sanitize_filename(filename):
#     return "".join(c if c.isalnum() or c in ['.', '-', '_'] else '_' for c in filename)

def main():
    global prog,from_email
    prog = 0
    output_folder_path = (file_path + r'\assets\logs_details')
    creds_path = os.path.join(file_path, 'assets', 'data')
    authed_path = os.path.join(file_path, 'assets', 'data', 'tokens')
    total1 = len([item for item in os.listdir(creds_path) if os.path.isfile(os.path.join(creds_path, item))])
    total2 = len([item for item in os.listdir(authed_path) if os.path.isfile(os.path.join(authed_path, item))])
    if total1 != total2:
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"An error occoured\nMessage >> Please authorize all your credentials before sending out, else remove the unauthorized accounts\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')
    else:
        delete_att = keep_pdf_var.get()
        try:
            if not os.path.exists(output_folder_path):
                os.makedirs(output_folder_path)
        except:
            pass
        #################################################
        txt_file_path = letter_path_full
        html_file_path = attachment_path_full
        delay = int(pause_time_var.get())
        ################################################
        max_emails_per_account = int(send_per_account_var.get())
        df = pd.read_csv(csv_path_full, encoding='ISO-8859-1')
        
        with open(txt_file_path, "r", encoding="utf-8") as f:
            email_body_template = f.read()
        total_sent_emails, _ = load_progress()
        for i in range(total1):
            if stop_flag.is_set():
                break
            sent_per_account = 0
            if total_sent_emails >=len(df):
                outbox.configure(state='normal')
                outbox.insert("end",f"########################################\n")
                outbox.insert("end",f"Message >> No mails to send, to restart sending , reset the status logs\n")
                outbox.see('end')
                outbox.update_idletasks()
                outbox.configure(state='disabled')
                outbox.insert("end",f"########################################\n")
                start_sender_btn.configure(state='normal')
                break
            service = gmail_authenticate(i)
            if service is None:
                continue
            user_info = service.users().getProfile(userId='me').execute()
            from_email = (user_info['emailAddress'])
            if not service:
                break
            df_to_send = df.iloc[total_sent_emails:]
            df_to_send = df_to_send.iloc[:min(max_emails_per_account, len(df_to_send))]
            if df_to_send.empty:
                continue
            for idx, row in df_to_send.iterrows():
                prog+=1
                prods = total1*max_emails_per_account
                iter_step = prog/prods
                progress_bar.set(iter_step)
                progress_bar.update_idletasks()
                progress_percentage = (prog / prods) * 100
                cent=(str(progress_percentage).split('.')[0])
                progress_count.configure(text=f'{cent}%')    
                if stop_flag.is_set():
                    break
                email_body = email_body_template
                email_body = replace_tags_with_data(email_body, row)
                pdf_file_name = row["PDF_File_Name"]
                email = row.get("Email", "")
                if not email:
                    continue
                subject1 = row.get("Subject", "")
                with open(html_file_path, "r") as f:
                    html_content = f.read()
                html_content = replace_tags_with_data(html_content, row)
                html_output_file = os.path.join(output_folder_path, f"{pdf_file_name}_{idx}.html")
                with open(html_output_file, "w",encoding='utf-8') as f:
                    f.write(html_content.replace('nan', ''))
                pdf_output_file = os.path.join(output_folder_path, f"{pdf_file_name}_{idx}.pdf")
                convert_html_to_pdf(html_output_file, pdf_output_file,pdf_file_name)
                send_message(service, destination=email, obj=subject1, body=str(email_body).replace('nan',''), attachments=[pdf_output_file])
                if delete_att == 'on':
                    pass
                else:
                    os.remove(html_output_file)
                    os.remove(pdf_output_file)
                sent_per_account += 1
                total_sent_emails += 1
                save_progress(total_sent_emails, sent_per_account)
                if pause_var.get() == 'on':
                    outbox.configure(state='normal')
                    outbox.insert("end",f"########################################\n")
                    outbox.insert("end",f"Waiting for {delay} Seconds\n")
                    outbox.see('end')
                    outbox.update_idletasks()
                    outbox.configure(state='disabled')
                    outbox.insert("end",f"########################################\n")
                    start_sender_btn.configure(state='normal')
                    time.sleep(delay)
                if sent_per_account >= max_emails_per_account:
                    break
            if prog == prods:
                outbox.configure(state='normal')
                outbox.insert("end",f"########################################\n")
                outbox.insert("end",f"Message >> Sendout Completed\nTotal Mails sent >> {prog}\nRemaining emails to send >> {len(email_addresses)-prods}\n")
                outbox.see('end')
                outbox.update_idletasks()
                outbox.configure(state='disabled')
                outbox.insert("end",f"########################################\n")
                start_sender_btn.configure(state='normal')
                break
            if total_sent_emails >= len(df):
                start_sender_btn.configure(state='normal')
                break
            
            
            
            





def start_main_sender():
    global sent,failed,pdf
    stop_flag.clear()
    outbox.configure(state='normal')
    outbox.delete("0.0", "end")
    outbox.configure(state='disabled')
    if sent != 0:
        sent = 0
    if failed != 0:
        failed = 0
    if pdf != 0:
        pdf = 0
    if csv_path.get() == '' or letter_path.get() == '' or attachment_path.get() =='' or names_path.get()=='':
        print('some fields is missing , reimport and continue')
    else:
        start_sender_btn.configure(state='disabled')
        a = threading.Thread(target=main,daemon=True)
        a.start()

def stop_main_sender():
    if csv_path.get() == '' or letter_path.get() == '' or attachment_path.get() =='' or names_path.get()=='':
        pass
    else:
        start_sender_btn.configure(state='normal')
        stop_flag.set()

def reset_logs_f():
    try:
        with open(progress_file_path, 'w') as progress_file:
            progress_file.write('')
            outbox.configure(state='normal')
            outbox.insert("end",f"########################################\n")
            outbox.insert("end",f"Logs Reset Successfuly,\nNote >> Sendout will not continue from where it ended\n")
            outbox.see('end')
            outbox.update_idletasks()
            outbox.insert("end",f"########################################\n")
            outbox.configure(state='disabled')
    except Exception as e:
        outbox.configure(state='normal')
        outbox.insert("end",f"########################################\n")
        outbox.insert("end",f"An error occoured\nReason >> {e}\n")
        outbox.see('end')
        outbox.update_idletasks()
        outbox.insert("end",f"########################################\n")
        outbox.configure(state='disabled')







##################################################################################################################
def home_page():
    global home_frame,progress_bar
    global total_mails,total_csv,total_att,total_pdf,total_logs,total_auth,total_sent,total_failed,outbox
    global start_authorize_btn,authorize_outbox,start_sender_btn,stop_sender_btn,use_send_per_account_var,progress_count
    global pause_var,keep_pdf_var,pause_time_var,send_per_account_var,sheet_type_var
    global csv_path,file_path,letter_path,attachment_path,names_path,total_att
    home_frame = ctk.CTkFrame(main_frame,fg_color=dark)
    home_frame.pack_propagate(False)
    ################################################################
    progress_frame = ctk.CTkFrame(home_frame,fg_color=dark,height=10)
    progress_frame.pack_propagate(False)
    progress_frame.pack(side='top',fill='x')
    progress_bar=ctk.CTkProgressBar(progress_frame,progress_color=yellow,width=660,height=10,fg_color=black)
    progress_bar.set(0)
    progress_bar.pack(padx=2,pady=3,side='left')
    progress_count = ctk.CTkLabel(progress_frame,text='0%',font=ctk.CTkFont(size=9),text_color=white)
    progress_count.pack(side='right')
    ################################################################
    top_frame = ctk.CTkFrame(home_frame,fg_color=dark)
    top_frame.pack_propagate(False)
    top_frame.pack(side='top',fill='x')
    ################################################################
    t1 = ctk.CTkFrame(top_frame,fg_color=dark,width=200)
    t1.pack_propagate(False)
    t1.pack(side='left',fill='x',padx=2)
    left_top = ctk.CTkFrame(t1,fg_color=dark,width=200,height=96)
    left_top.pack_propagate(False)
    left_top.pack(side='top',fill='x',padx=1,pady=1)
    mails_frame =  ctk.CTkFrame(left_top,fg_color=black,width=98,height=96)
    mails_frame.pack_propagate(False)
    mails_frame.pack(side='left',fill='x',padx=1,pady=1)
    mails_details = ctk.CTkLabel(mails_frame,text=' Mails',image=mails_img,compound='left',text_color=white)
    mails_details.place(x=6, y=3)
    total_mails = ctk.CTkLabel(mails_frame,text='0',compound='left',text_color=white)
    total_mails.place(x=6, y=65)
    csv_frame =  ctk.CTkFrame(left_top,fg_color=black,width=98,height=96)
    csv_frame.pack_propagate(False)
    csv_frame.pack(side='right',fill='x',padx=1,pady=1)
    csv_details = ctk.CTkLabel(csv_frame,text=' CSV',image=csv_img,compound='left',text_color=white)
    csv_details.place(x=6, y=3)
    total_csv = ctk.CTkLabel(csv_frame,text='0',compound='left',text_color=white)
    total_csv.place(x=6, y=65)
    left_down = ctk.CTkFrame(t1,fg_color=dark,width=200,height=96)
    left_down.pack_propagate(False)
    left_down.pack(side='bottom',fill='x',padx=1,pady=1)
    logs_frame =  ctk.CTkFrame(left_down,fg_color=black,width=98,height=96)
    logs_frame.pack_propagate(False)
    logs_frame.pack(side='left',fill='x',padx=1,pady=1)
    logs_details = ctk.CTkLabel(logs_frame,text=' Logs',image=logs_img,compound='left',text_color=white)
    logs_details.place(x=6, y=3)
    total_logs = ctk.CTkLabel(logs_frame,text='0',compound='left',text_color=white)
    total_logs.place(x=6, y=65)
    auth_frame =  ctk.CTkFrame(left_down,fg_color=black,width=98,height=96)
    auth_frame.pack_propagate(False)
    auth_frame.pack(side='right',fill='x',padx=1,pady=1)
    auth_details = ctk.CTkLabel(auth_frame,text=' Authed',image=auth_img,compound='left',text_color=white)
    auth_details.place(x=6, y=3)
    total_auth = ctk.CTkLabel(auth_frame,text='0',compound='left',text_color=white)
    total_auth.place(x=6, y=65)
    ################################################################
    t2= ctk.CTkFrame(top_frame,fg_color=dark,width=287)
    t2.pack_propagate(False)
    t2.pack(side='left',fill='x',padx=(1,1))
    outbox = ctk.CTkTextbox(t2,fg_color=dark,border_color='#353935',border_width=2,font=ctk.CTkFont(size=10),text_color=white)
    outbox.pack(fill='both',expand=True)
    outbox.configure(state='disabled')
    ################################################################
    t3= ctk.CTkFrame(top_frame,fg_color=dark,width=200)
    t3.pack_propagate(False)
    t3.pack(side='right',fill='x',padx=2)
    right_top = ctk.CTkFrame(t3,fg_color=dark,width=200,height=96)
    right_top.pack_propagate(False)
    right_top.pack(side='top',fill='x',padx=1,pady=1)
    att_frame =  ctk.CTkFrame(right_top,fg_color=black,width=98,height=96)
    att_frame.pack_propagate(False)
    att_frame.pack(side='left',fill='x',padx=1,pady=1)
    att_details = ctk.CTkLabel(att_frame,text=' Attach.',image=att_img,compound='left',text_color=white)
    att_details.place(x=6, y=3)
    total_att = ctk.CTkLabel(att_frame,text='0',compound='left',text_color=white)
    total_att.place(x=6, y=65)
    pdf_frame =  ctk.CTkFrame(right_top,fg_color=black,width=98,height=96)
    pdf_frame.pack_propagate(False)
    pdf_frame.pack(side='right',fill='x',padx=1,pady=1)
    pdf_details = ctk.CTkLabel(pdf_frame,text=' Pdf(s)',image=pdf_img,compound='left',text_color=white)
    pdf_details.place(x=6, y=3)
    total_pdf = ctk.CTkLabel(pdf_frame,text='0',compound='left',text_color=white)
    total_pdf.place(x=6, y=65)
    right_down = ctk.CTkFrame(t3,fg_color=dark,width=200,height=96)
    right_down.pack_propagate(False)
    right_down.pack(side='bottom',fill='x',padx=1,pady=1)
    sent_frame =  ctk.CTkFrame(right_down,fg_color=black,width=98,height=96)
    sent_frame.pack_propagate(False)
    sent_frame.pack(side='left',fill='x',padx=1,pady=1)
    sent_details = ctk.CTkLabel(sent_frame,text=' Sent',image=sent_img,compound='left',text_color=white)
    sent_details.place(x=6, y=3)
    total_sent = ctk.CTkLabel(sent_frame,text='0',compound='left',text_color=white)
    total_sent.place(x=6, y=65)
    failed_frame =  ctk.CTkFrame(right_down,fg_color=black,width=98,height=96)
    failed_frame.pack_propagate(False)
    failed_frame.pack(side='right',fill='x',padx=1,pady=1)
    failed_details = ctk.CTkLabel(failed_frame,text=' Failed',image=failed_img,compound='left',text_color=white)
    failed_details.place(x=6, y=3)
    total_failed = ctk.CTkLabel(failed_frame,text='0',compound='left',text_color=white)
    total_failed.place(x=6, y=65)
    ################################################################
    down_frame = ctk.CTkFrame(home_frame,fg_color=dark)
    down_frame.pack_propagate(False)
    down_frame.pack(side='bottom',fill='x')
    settings_tab = ctk.CTkTabview(down_frame,segmented_button_fg_color=black,segmented_button_selected_color=dark,fg_color=dark,
                             segmented_button_selected_hover_color=gray,segmented_button_unselected_color=black,text_color=white,
                             segmented_button_unselected_hover_color="#251843")
    settings_tab.pack(fill='both',expand=True)
    settings_tab.add(name='Authorize Accounts')
    settings_tab.add(name="Other Settings")
    settings_tab.set(name='Other Settings')
    authorize_account_frame = ctk.CTkFrame(settings_tab.tab(name='Authorize Accounts'),fg_color=dark)
    authorize_account_frame.pack_propagate(False)
    authorize_account_frame.pack(fill='both')
    i_frame = ctk.CTkFrame(authorize_account_frame,width=300,fg_color=dark,)
    i_frame.pack_propagate(False)
    i_frame.pack(side='left')
    authorize_outbox = ctk.CTkTextbox(i_frame,fg_color=dark,border_color='#353935',border_width=2,text_color=white,font=ctk.CTkFont(size=10))
    authorize_outbox.pack(fill='both',expand=True)
    authorize_outbox.configure(state='disabled')
    o_frame = ctk.CTkFrame(authorize_account_frame,fg_color=dark)
    o_frame.pack_propagate(False)
    o_frame.pack(side='right',fill='both',expand=True,padx=(3,0))
    message = 'Important Note:\n\nPlease, make sure you put all\nyour credentials(cred0.json) in the rightful\nfolder(assets\data), also, mame it in accordance,\nExample;cred0.json, cred1.json,\n\nWarning:\n\n Do Not Delete Any Folder'
    n =ctk.CTkFrame(o_frame,height=75,fg_color=dark)
    n.pack_propagate(False)
    n.pack(side='top',fill='x',)
    m=ctk.CTkScrollableFrame(n,fg_color=dark,border_color=black,border_width=2)
    m.pack(fill='both',expand=True)
    note_label = ctk.CTkLabel(m,text=message,text_color=white)
    note_label.pack()
    o =ctk.CTkFrame(o_frame,height=35,fg_color=dark)
    o.pack_propagate(False)
    o.pack(side='bottom',fill='x',pady=3)
    start_authorize_btn = ctk.CTkButton(o,image=authb_img,text='Authorize Accounts',fg_color=dark,
                                        hover_color=gray,border_color='#353935',border_width=2,command=start_authenticator)
    start_authorize_btn.pack(fill='both')
    CTkToolTip(start_authorize_btn, message="Authorize Accounts")
    other_frame = ctk.CTkFrame(settings_tab.tab(name="Other Settings"),fg_color=dark)
    other_frame.pack_propagate(False)
    other_frame.pack(fill='both')
    f1 = ctk.CTkFrame(other_frame,width=2,fg_color=yellow)
    f1.pack(side='left',padx=5)
    frame1 = ctk.CTkFrame(other_frame,width=240,fg_color=dark)
    frame1.pack_propagate(False)
    frame1.pack(side='left')
    f1 = ctk.CTkFrame(other_frame,width=2,fg_color=yellow)
    f1.pack(side='left',padx=5)
    frame2 = ctk.CTkFrame(other_frame,width=400,fg_color=dark)
    frame2.pack_propagate(False)
    frame2.pack(side='left')
    f1 = ctk.CTkFrame(other_frame,width=2,fg_color=yellow)
    f1.pack(side='left',padx=5)
    imports_frame1 = ctk.CTkFrame(frame1,fg_color=dark,height=23)
    imports_frame1.pack_propagate(False)
    imports_frame1.pack(side='top',fill='x',pady=2)
    
    csv_path = ctk.CTkEntry(imports_frame1,width=206,fg_color=dark,border_color='#353935',border_width=2,placeholder_text='import csv file')
    csv_path.pack(side='left')
    csv_path.configure(state='disable')
    
    import_csv_btn = ctk.CTkButton(imports_frame1,width=30,text='',
                                   fg_color=dark,image=impcsv_img,
                                   hover_color='#353935',command=open_csv)
    import_csv_btn.pack(side='right')
    CTkToolTip(import_csv_btn, message="Import Csv file")
    ############################################################################
    imports_frame2 = ctk.CTkFrame(frame1,fg_color=dark,height=23)
    imports_frame2.pack_propagate(False)
    imports_frame2.pack(side='top',fill='x',pady=2)
    
    attachment_path = ctk.CTkEntry(imports_frame2,width=206,fg_color=dark,border_color='#353935',border_width=2,placeholder_text='import attachment')
    attachment_path.pack(side='left')
    attachment_path.configure(state='disable')
    
    import_attachment_btn = ctk.CTkButton(imports_frame2,width=30,text='',
                                   fg_color=dark,image=import_attachment_img,
                                   hover_color='#353935',command=open_attachment)
    import_attachment_btn.pack(side='right')
    CTkToolTip(import_attachment_btn, message="Import attachment")
    ############################################################################
    
    imports_frame3 = ctk.CTkFrame(frame1,fg_color=dark,height=23)
    imports_frame3.pack_propagate(False)
    imports_frame3.pack(side='top',fill='x',pady=2)
    letter_path = ctk.CTkEntry(imports_frame3,width=206,fg_color=dark,border_color='#353935',border_width=2,placeholder_text='Select your letter')
    letter_path.pack(side='left')
    letter_path.configure(state='disable')
    import_letter_btn = ctk.CTkButton(imports_frame3,width=30,text='',
                                   fg_color=dark,image=import_letter_img,
                                   hover_color='#353935',command=open_letter)
    import_letter_btn.pack(side='right')
    
    CTkToolTip(import_letter_btn, message="Select your letter")
    
    imports_frame4 = ctk.CTkFrame(frame1,fg_color=dark,height=25)
    imports_frame4.pack_propagate(False)
    imports_frame4.pack(side='top',fill='x',pady=2)
    names_path = ctk.CTkEntry(imports_frame4,width=206,fg_color=dark,border_color='#353935',border_width=2,placeholder_text='Select From Names')
    names_path.pack(side='left')
    names_path.configure(state='disable')
    import_names_btn = ctk.CTkButton(imports_frame4,width=30,text='',
                                   fg_color=dark,image=import_letter_img,
                                   hover_color='#353935',command=open_names)
    import_names_btn.pack(side='right')
    
    CTkToolTip(import_names_btn, message="Select your subjects")
    #################################  CONTROLLER  ###########################################
    a1=ctk.CTkFrame(frame2,fg_color=dark,height=22)
    a1.pack_propagate(False)
    a1.pack(pady=1,fill='both')
    
    a2=ctk.CTkFrame(frame2,fg_color=dark,height=22)
    a2.pack_propagate(False)
    a2.pack(pady=1,fill='both')
    
    a3=ctk.CTkFrame(frame2,fg_color=dark,height=22)
    a3.pack_propagate(False)
    a3.pack(pady=1,fill='both')
    
    a4=ctk.CTkFrame(frame2,fg_color=dark,height=28)
    a4.pack_propagate(False)
    a4.pack(pady=1,fill='both')
    
    a5=ctk.CTkFrame(frame2,fg_color=yellow,height=2)
    a5.pack_propagate(False)
    a5.pack(pady=1)
    
    use_send_per_account_var = ctk.StringVar()
    use_send_per_account = ctk.CTkCheckBox(a1, text="Maximum limit Per Api (Default 50): ",command=None,
                                       height=20,variable=use_send_per_account_var  ,checkbox_height=20,
                                       checkmark_color=yellow,border_color='#353935',border_width=1,
                                       fg_color=gray,hover_color='#353935',
                                       onvalue="on",offvalue="off")
    use_send_per_account.pack(side='left',padx=(1,1))
    use_send_per_account_var.set('on')
    use_send_per_account.configure(state='disabled')
    
    rotate_vals = ['50','100','150','200','250','300','350','399']
    send_per_account_var = ctk.StringVar()
    send_per_account = ctk.CTkComboBox(a1,values=rotate_vals,fg_color=dark,border_width=1,width=320,
                                    border_color='#353935',button_color=black,button_hover_color='#353935',
                                    dropdown_fg_color=dark,dropdown_hover_color='#353935',
                                    command=None, variable=send_per_account_var)
    send_per_account.pack(side='left')
    send_per_account_var.set("50")
    ############################################################################
    pause_var = ctk.StringVar()
    pause_after = ctk.CTkCheckBox(a2, text="Pause For (In Seconds) ",command=None,
                                       height=20,variable=pause_var  ,checkbox_height=20,
                                       checkmark_color=yellow,border_color='#353935',border_width=1,
                                       fg_color=gray,hover_color='#353935',
                                       onvalue="on",offvalue="off")
    pause_after.pack(side='left',padx=(1,1))
    pause_time_var = ctk.StringVar()
    pause_time = ctk.CTkComboBox(a2,values=["3", "4", "5","6", "7", "8","9","10"],fg_color=dark,border_width=1,
                                    border_color='#353935',button_color=black,button_hover_color='#353935',width=70,
                                    dropdown_fg_color=black,dropdown_hover_color='#353935',
                                    command=None, variable=pause_time_var)
    pause_time.pack(side='left')
    pause_time_var.set("5")
    ###############################
    l1=ctk.CTkLabel(a2,text=' Sheet Type').pack(side='left',padx=(5,4))
    
    sheet_type_var = ctk.StringVar()
    sheet_type = ctk.CTkComboBox(a2,values=["A5", "A4", "B5","C4",],fg_color=dark,border_width=1,width=60,
                                    border_color='#353935',button_color=black,button_hover_color='#353935',
                                    dropdown_fg_color=black,dropdown_hover_color='#353935',
                                    command=None, variable=sheet_type_var)
    sheet_type.pack(side='left')
    sheet_type.set("A5")
    
    ################################################################################
    keep_pdf_var = ctk.StringVar()
    keep_pdf = ctk.CTkCheckBox(a3, text="Keep converted pdf(s) ",command=None,
                                       height=20,variable=keep_pdf_var  ,checkbox_height=20,
                                       checkmark_color=yellow,border_color='#353935',border_width=1,
                                       fg_color=gray,hover_color='#353935',
                                       onvalue="on",offvalue="off")
    keep_pdf.pack(side='left',padx=(1,1))
    
    reset_logs = ctk.CTkButton(a3,fg_color=dark,hover_color='#353935',text='Reset Status',image=None,width=150,command=reset_logs_f,border_color='#353935',border_width=1)
    reset_logs.pack(side='left',padx=5)
    
    ################################################################################
    start_sender_btn = ctk.CTkButton(a4,fg_color=dark,hover_color='#353935',text='Start',image=start_img,width=200,command=start_main_sender,border_color='#353935',border_width=1)
    start_sender_btn.pack(side='left',pady=3)
    
    stop_sender_btn = ctk.CTkButton(a4,fg_color=dark,hover_color='#353935',text='Stop',image=stop_img,width=200,command=stop_main_sender,border_color='#353935',border_width=1)
    stop_sender_btn.pack(side='right',pady=3)
    
    




##################################################################################################################
def about_page():
    global about_frame
    about_frame = ctk.CTkFrame(main_frame,fg_color=dark,corner_radius=1)
    about_frame.pack_propagate(False)
    ################################################################
    



##################################################################################################################
def hide_indicator():
    home_btn.configure(fg_color="transparent",bg_color="transparent",hover_color="#251843")
    about_btn.configure(fg_color="transparent",bg_color="transparent",hover_color="#251843")
    ################################################################
    


##################################################################################################################
def indicator(ex):
    hide_indicator()
    ex.configure(fg_color=dark,hover_color='#353935')
    ################################################################
    



##################################################################################################################
def get_values():
    home_page()
    about_page()
    check_creds()
    ################################################################
    


##################################################################################################################
def main_layout():
    global home_btn,about_btn,main_frame
    app=ctk.CTk(fg_color=dark)
    app.geometry("750x390")
    app.resizable(False,False)
    app.title("U-Society [order 27][Google api sender]")
    app.iconbitmap(file_path+'/images/icon.ico')
    side_menu_frame = ctk.CTkFrame(app,fg_color=black)
    side_menu_frame.pack(side='left',padx=(2,1),pady=2)
    side_menu_frame.pack_propagate(False)
    side_menu_frame.configure(width=40,height=390)
    
    main_frame = ctk.CTkFrame(app,fg_color=black)
    main_frame.pack(expand=True,fill='both',padx=2,pady=2)
    main_frame.pack_propagate(False)
    main_frame.configure(width=750,height=360)
    
    home_btn = ctk.CTkButton(side_menu_frame,command=lambda: select_page(home_frame, all_pages, home_btn),text='',width=25,image=home_img,fg_color='transparent',bg_color="transparent",hover_color="#353935")
    home_btn.pack(side='top',padx=2,pady=5)
    CTkToolTip(home_btn, message="Home")
    f=ctk.CTkFrame(side_menu_frame,width=20,height=2,fg_color=yellow)
    f.pack(side='top',padx=2,pady=(0,2))
    
    exit_btn = ctk.CTkButton(side_menu_frame,command=app.destroy,text='',width=25,image=exit_img,fg_color='transparent',bg_color="transparent",hover_color="#353935")
    exit_btn.pack(side='bottom',padx=2,pady=5)
    CTkToolTip(exit_btn, message="Exit")
    f=ctk.CTkFrame(side_menu_frame,width=20,height=2,fg_color=yellow)
    f.pack(side='bottom',padx=2,pady=(0,2))
    about_btn = ctk.CTkButton(side_menu_frame,command=lambda: select_page(about_frame, all_pages, about_btn),text='',width=25,image=about_img,fg_color='transparent',bg_color="transparent",hover_color="#353935")
    about_btn.pack(side='bottom',padx=2,pady=(0,5))
    CTkToolTip(about_btn, message="About")
    get_values()
    all_pages = [home_frame,about_frame]
    select_page(home_frame, all_pages, home_btn)
    
    app.mainloop()
##########################################################################################
def select_page(selected_page, all_pages, ex):
    indicator(ex)
    for page in all_pages:
        page.pack_forget()
    selected_page.pack(fill='both',expand=True,padx=(5,0),pady=(5,5))



main_layout()
