import os
import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog

# Initialize Outlook for sending emails
outlook = win32.Dispatch('outlook.application')

# Get the current date
today = datetime.today()

# Set the paths to the images (ensure these paths are correct)
imagen_compliance = os.path.abspath('Compliance.jpg')
imagen_faurecia = os.path.abspath('faurecia-inspiring.png')

# Function to validate image paths
def validate_image_paths():
    if not os.path.exists(imagen_compliance):
        raise FileNotFoundError(f"Image not found: {imagen_compliance}")
    if not os.path.exists(imagen_faurecia):
        raise FileNotFoundError(f"Image not found: {imagen_faurecia}")

# Function to send ECOI emails
def enviar_correo_ecoi(email, status, last_reminder, name=None):
    # Validate image paths
    validate_image_paths()

    # Calculate the deadline (1 month from today)
    deadline = today + timedelta(days=30)
    deadline_str = deadline.strftime("%B %d, %Y")

    # Select the email content based on status
    if status == "Invitar a declarar":
        asunto = "Mandatory Conflict of Interest Declaration to be performed"
        cuerpo_html = f"""
        <html>
            <body style="font-family: 'Century Gothic';">
                <p style="font-size: 24pt; font-weight: bold; color: #FF69B4;">Dear Faurecians,</p>
                <p>It is time to complete your Conflict of Interest Declaration. Please take a few minutes to complete your electronic declaration of conflict of interest by <strong style="color: #FF69B4;">{deadline_str}</strong>, to ensure a common ethic culture and strong compliance discipline. You can access it through the following link: <a href="https://e-coi.app.corp/">Complete your declaration</a>. Completion takes 5 minutes or less. Every Faurecian must complete a declaration, even if there is no conflict to disclose.</p>
                <p><strong>What is a Conflict of Interest?</strong></p>
                <p>As stated in our Code of Ethics and Article 4 of the Code of Conduct for the prevention of corruption, a Conflict of Interest refers to a situation in which you have a personal interest likely to influence or appear to influence your activities in your function. The interest may be direct or indirect and concern you or your close relations. This interest may be diverse (financial, economic, professional, political, personal, etc.) and may occur intentionally or unintentionally.</p>
                <p>Any conflict of interest must be identified, declared, and effectively managed. Managing conflicts of interest well is not only good practice but also protects Faurecia assets and the persons involved in the transactions.</p>
                <p><strong>Where can I go with questions?</strong></p>
                <p>You can refer to our policy <a href="https://faurus.ww.faurecia.com/external-link.jspa?url=https%3A%2F%2Fapps.faurecia%2Fsites%2Ffcp%2FLists%2FFCP%2FForms%2FID%2520Card%2Fdocsethomepage.aspx%3FID%3D9163%26FolderCTID%3D0x0120D5200009E984788F73457698654A755A26EF900BA834E384132C740B11DD0278C622E3%26List%3Dd2dd69ba-b701-446a-a53f-f10d556b28ae%26RootFolder%3D%252Fsites%252Ffcp%252FLists%252FFCP%252FFAU%2520DLSG%252D2406%26RecSrc%3D%252Fsites%252Ffcp%252FLists%252FFCP%252FFAU%2520DLSG%252D2406">here</a>, review the <a href="https://e-coi.app.corp/">FAQs section</a>, or contact your Compliance Officer:</p>
                <ul>
                    <li>Claudia Morales - Compliance Director: <a href="mailto:claudia.morales@faurecia.com">claudia.morales@faurecia.com</a></li>
                    <li>Evandro Rocha - Compliance Officer: <a href="mailto:evandro.rocha@faurecia.com">evandro.rocha@faurecia.com</a></li>
                </ul>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Thank you for acting with transparency and for your commitment to the Compliance culture!</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Group Compliance</p>
                <img src="cid:faurecia_image" style="width: 200px;" alt="Faurecia Logo">
            </body>
        </html>
        """
    elif status == "Recordatorio":
        asunto = "REMINDER: Complete your annual Conflict of Interest Declaration before the deadline"
        cuerpo_html = f"""
        <html>
            <body style="font-family: 'Century Gothic';">
                <p style="font-size: 24pt; font-weight: bold; color: #FF69B4;">Dear Faurecian,</p>
                <p>It is time to complete your annual Conflict of Interest Declaration. Every Faurecian must complete a declaration, even if there is no conflict to disclose. If you have not completed it yet, please click on the following link by <strong style="color: #FF69B4;">{deadline_str}</strong>: <a href="https://e-coi.app.corp/">Complete your declaration</a>. It takes 5 minutes or less. Please mind completing it regardless of any other recent declaration you have submitted in the previous tool.</p>
                <p><strong>What is a Conflict of Interest?</strong></p>
                <p>As stated in our Code of Ethics and Article 4 of the Code of Conduct for the prevention of corruption, a Conflict of Interest refers to a situation in which you have a personal interest likely to influence or appear to influence your activities in your function. The interest may be direct or indirect and concern you or your close relations. This interest may be diverse (financial, economic, professional, political, personal, etc.) and may occur intentionally or unintentionally.</p>
                <p>Any conflict of interest must be identified, declared, and effectively managed. Managing conflicts of interest well is not only good practice but also protects Faurecia assets and the persons involved in the transactions.</p>
                <p><strong>Where can I go with questions?</strong></p>
                <p>You can refer to our policy <a href="https://faurus.ww.faurecia.com/external-link.jspa?url=https%3A%2F%2Fapps.faurecia%2Fsites%2Ffcp%2FLists%2FFCP%2FForms%2FID%2520Card%2Fdocsethomepage.aspx%3FID%3D9163%26FolderCTID%3D0x0120D5200009E984788F73457698654A755A26EF900BA834E384132C740B11DD0278C622E3%26List%3Dd2dd69ba-b701-446a-a53f-f10d556b28ae%26RootFolder%3D%252Fsites%252Ffcp%252FLists%252FFCP%252FFAU%2520DLSG%252D2406%26RecSrc%3D%252Fsites%252Ffcp%252FLists%252FFCP%252FFAU%2520DLSG%252D2406">here</a>, review the <a href="https://e-coi.app.corp/">FAQs section</a>, or contact your Compliance Officer:</p>
                <ul>
                    <li>Claudia Morales - Compliance Director: <a href="mailto:claudia.morales@faurecia.com">claudia.morales@faurecia.com</a></li>
                    <li>Evandro Rocha - Compliance Officer: <a href="mailto:evandro.rocha@faurecia.com">evandro.rocha@faurecia.com</a></li>
                </ul>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Thank you for acting with transparency and for your commitment to the Compliance culture!</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Group Compliance</p>
                <img src="cid:faurecia_image" style="width: 200px;" alt="Faurecia Logo">
            </body>
        </html>
        """
    elif status == "Urgente":
        asunto = "REMINDER: Complete your annual Conflict of Interest Declaration DUE TODAY"
        cuerpo_html = f"""
        <html>
            <body style="font-family: 'Century Gothic';">
                <p style="font-size: 24pt; font-weight: bold; color: #FF69B4;">Dear {name},</p>
                <p>You have a mandatory Conflict of Interest Declaration due today. Every Faurecian must complete a declaration, even if there is no conflict to disclose. Please click on the following link by <strong style="color: #FF69B4;">{deadline_str}</strong>: <a href="https://e-coi.app.corp/">Complete your declaration</a>. It takes 5 minutes or less. Please mind completing it regardless of any other recent declaration you have submitted in the previous tool.</p>
                <p><strong>What is a Conflict of Interest?</strong></p>
                <p>As stated in our Code of Ethics and Article 4 of the Code of Conduct for the prevention of corruption, a Conflict of Interest refers to a situation in which you have a personal interest likely to influence or appear to influence your activities in your function. The interest may be direct or indirect and concern you or your close relations. This interest may be diverse (financial, economic, professional, political, personal, etc.) and may occur intentionally or unintentionally.</p>
                <p>Any conflict of interest must be identified, declared, and effectively managed. Managing conflicts of interest well is not only good practice but also protects Faurecia assets and the persons involved in the transactions.</p>
                <p><strong>Where can I go with questions?</strong></p>
                <p>You can refer to our policy <a href="https://faurus.ww.faurecia.com/external-link.jspa?url=https%3A%2F%2Fapps.faurecia%2Fsites%2Ffcp%2FLists%2FFCP%2FForms%2FID%2520Card%2Fdocsethomepage.aspx%3FID%3D9163%26FolderCTID%3D0x0120D5200009E984788F73457698654A755A26EF900BA834E384132C740B11DD0278C622E3%26List%3Dd2dd69ba-b701-446a-a53f-f10d556b28ae%26RootFolder%3D%252Fsites%252Ffcp%252FLists%252FFCP%252FFAU%2520DLSG%252D2406%26RecSrc%3D%252Fsites%252Ffcp%252FLists%252FFCP%252FFAU%2520DLSG%252D2406">here</a>, review the <a href="https://e-coi.app.corp/">FAQs section</a>, or contact your Compliance Officer:</p>
                <ul>
                    <li>Claudia Morales - Compliance Director: <a href="mailto:claudia.morales@faurecia.com">claudia.morales@faurecia.com</a></li>
                    <li>Evandro Rocha - Compliance Officer: <a href="mailto:evandro.rocha@faurecia.com">evandro.rocha@faurecia.com</a></li>
                </ul>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Thank you for acting with transparency and for your commitment to the Compliance culture!</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Group Compliance</p>
                <img src="cid:faurecia_image" style="width: 200px;" alt="Faurecia Logo">
            </body>
        </html>
        """
    else:
        # Handle unexpected status
        asunto = "Conflict of Interest Declaration"
        cuerpo_html = f"""
        <html>
            <body>
                <p>Dear {name},</p>
                <p>Please complete your Conflict of Interest Declaration by the deadline.</p>
            </body>
        </html>
        """

    # Create and send the email
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = asunto
    mail.HTMLBody = cuerpo_html
    mail.Importance = 2  # High importance
    
    # Attach images as inline content (without making them downloadable)
    attachment_faurecia = mail.Attachments.Add(imagen_faurecia)
    attachment_faurecia.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "faurecia_image")
    attachment_faurecia.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/png")

    attachment_compliance = mail.Attachments.Add(imagen_compliance)
    attachment_compliance.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "compliance_image")
    attachment_compliance.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/jpeg")
    
    mail.Send()

# Function to send MOOCs emails
def enviar_correo_moocs(email, moocs_pendientes, fase_recordatorio):
    # Validate image paths
    validate_image_paths()

    debe_ethics = 'Ethics' in moocs_pendientes
    debe_anticorruption = 'Anticorruption' in moocs_pendientes
    deadline = today + timedelta(days=30)
    deadline_str = deadline.strftime("%B %d, %Y")
    
    # Define subject and body based on what MOOCs are pending
    if debe_ethics and debe_anticorruption:
        asunto = "Reminder: Complete Your Ethics and Anti-corruption Online Trainings"
        cuerpo_html = f"""
        <html>
            <body style="font-family: 'Century Gothic';">
                <p style="font-size: 24pt; font-weight: bold; color: #FF69B4;">Dear Faurecians,</p>
                <p>You have mandatory <strong>Ethics and Anti-corruption Online Trainings</strong> due by <strong style="color: #FF69B4;">{deadline_str}</strong>.</p>
                <p>As part of the Compliance program, all employees must complete these trainings. Please be aware that the completion rate is a mandatory public information disclosed in Faurecia annual Registration document.</p>
                <p>Please take a few minutes to complete these trainings. Access the MOOCs using the following links:</p>
                <ul>
                    <li><a href='https://faurecia.coorpacademy.com/discipline/06'>Ethics</a></li>
                    <li><a href='https://faurecia.coorpacademy.com/discipline/disVkxVYVDFK'>Anti-Corruption</a></li>
                </ul>
                <p>If you have already done them in the last few days, you can safely ignore this reminder.</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Thank you for your commitment to the Compliance culture!</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Group Compliance</p>
                <img src="cid:faurecia_image" style="width: 200px;" alt="Faurecia Logo">
                <img src="cid:compliance_image" style="float: right; width: 200px;" alt="Compliance Logo">
            </body>
        </html>
        """
    elif debe_ethics:
        asunto = "Reminder: Complete Your Ethics Online Training"
        cuerpo_html = f"""
        <html>
            <body style="font-family: 'Century Gothic';">
                <p style="font-size: 24pt; font-weight: bold; color: #FF69B4;">Dear Faurecians,</p>
                <p>You have an <strong>Ethics Online Training</strong> due by <strong style="color: #FF69B4;">{deadline_str}</strong>.</p>
                <p>As part of the Compliance program, all employees must complete the mandatory Ethics Online Training.</p>
                <p>Please take a few minutes to complete it. Access the MOOC using the following link:</p>
                <p><a href='https://faurecia.coorpacademy.com/discipline/06'>Ethics</a></p>
                <p>If you have already done it in the last few days, you can safely ignore this reminder.</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Thank you for your commitment to the Compliance culture!</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Group Compliance</p>
                <img src="cid:faurecia_image" style="width: 200px;" alt="Faurecia Logo">
                <img src="cid:compliance_image" style="float: right; width: 200px;" alt="Compliance Logo">
            </body>
        </html>
        """
    elif debe_anticorruption:
        asunto = "Reminder: Complete Your Anti-corruption Online Training"
        cuerpo_html = f"""
        <html>
            <body style="font-family: 'Century Gothic';">
                <p style="font-size: 24pt; font-weight: bold; color: #FF69B4;">Dear Faurecians,</p>
                <p>You have an <strong>Anti-corruption Online Training</strong> due by <strong style="color: #FF69B4;">{deadline_str}</strong>.</p>
                <p>As part of the Compliance program, all employees must complete the mandatory Anti-corruption Online Training.</p>
                <p>Please take a few minutes to complete it. Access the MOOC using the following link:</p>
                <p><a href='https://faurecia.coorpacademy.com/discipline/disVkxVYVDFK'>Anti-Corruption</a></p>
                <p>If you have already done it in the last few days, you can safely ignore this reminder.</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Thank you for your commitment to the Compliance culture!</p>
                <p style="text-align: center; font-size: 16pt; font-weight: bold;">Group Compliance</p>
                <img src="cid:faurecia_image" style="width: 200px;" alt="Faurecia Logo">
                <img src="cid:compliance_image" style="float: right; width: 200px;" alt="Compliance Logo">
            </body>
        </html>
        """
    else:
        # No MOOCs to send reminders for
        return

    # Create and send the email
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = asunto
    mail.HTMLBody = cuerpo_html
    mail.Importance = 2  # High importance
    
    # Attach images as inline content (without making them downloadable)
    attachment_faurecia = mail.Attachments.Add(imagen_faurecia)
    attachment_faurecia.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "faurecia_image")
    attachment_faurecia.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/png")

    attachment_compliance = mail.Attachments.Add(imagen_compliance)
    attachment_compliance.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "compliance_image")
    attachment_compliance.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/jpeg")
    
    mail.Send()

# Function to process MOOCs emails
def procesar_moocs():
    file = filedialog.askopenfilename(title="Select MOOCs CSV", filetypes=(("CSV files", "*.csv"),))
    if not file:
        print("No file selected for MOOCs.")
        return
    
    df_moocs = pd.read_csv(file)
    
    for index, row in df_moocs.iterrows():
        email = row['Email']
        moocs_pendientes = row['Moocs Pendientes'].split(', ')
        fase_recordatorio = row['Fase de Recordatorio']
        
        enviar_correo_moocs(email, moocs_pendientes, fase_recordatorio)
    
    print("MOOCs emails processed.")

# Function to process ECOI emails
def procesar_ecoi():
    file = filedialog.askopenfilename(title="Select ECOI CSV", filetypes=(("CSV files", "*.csv"),))
    if not file:
        print("No file selected for ECOI.")
        return
    
    df_ecoi = pd.read_csv(file)
    
    for index, row in df_ecoi.iterrows():
        email = row['employee_email']
        status = row['Debe']
        last_reminder = row['last_reminder']
        name = row.get('Name', 'Faurecian')  # Default to "Faurecian" if no name is provided
        
        enviar_correo_ecoi(email, status, last_reminder, name)
    
    print("ECOI emails processed.")

# GUI setup for sending reminders
root = tk.Tk()
root.title('Send Reminders for MOOCs and ECOI')
root.geometry('400x300')
root.configure(bg="#2C3E50")

title_label = tk.Label(root, text="Send Reminders", font=('Helvetica', 18, 'bold'), bg="#2C3E50", fg="white")
title_label.pack(pady=20)

boton_moocs = tk.Button(root, text='Send MOOCs Reminders', command=procesar_moocs,
                        font=('Helvetica', 14), bg="#3498DB", fg="white", 
                        activebackground="#2980B9", activeforeground="white", padx=10, pady=10,
                        bd=0, relief="flat")
boton_moocs.pack(pady=10)

boton_ecoi = tk.Button(root, text='Send ECOI Reminders', command=procesar_ecoi,
                       font=('Helvetica', 14), bg="#FF6C63", fg="white", 
                       activebackground="#983B3B", activeforeground="white", padx=10, pady=10,
                       bd=0, relief="flat")
boton_ecoi.pack(pady=10)

root.mainloop()
