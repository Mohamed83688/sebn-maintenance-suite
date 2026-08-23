import smtplib
import ssl
import socket
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def get_smtp_config(email):
    """Detects provider and returns (host, port_tls, port_ssl)"""
    email = email.lower().strip()
    if "@gmail.com" in email:
        return "smtp.gmail.com", 587, 465
    elif "@outlook." in email or "@hotmail." in email or "@live." in email:
        return "smtp-mail.outlook.com", 587, 587
    elif "@yahoo." in email:
        return "smtp.mail.yahoo.com", 587, 465
    # Default to Gmail but maybe try common ports
    return "smtp.gmail.com", 587, 465


def _is_socket_blocked(err):
    """Detect WinError 10013 / Windows Firewall blocking outbound SMTP sockets."""
    err_str = str(err).lower()
    return "10013" in err_str or "access" in err_str or "permission" in err_str or "blocked" in err_str or "not permitted" in err_str


def test_email_connection(sender_email, app_password, recipient_email):
    """
    Tests the email connection without sending a real attachment.
    Returns (success: bool, details: str) with rich diagnostic info.
    """
    details = []
    host, port_tls, port_ssl = get_smtp_config(sender_email)
    
    # Step 1: Check Network (Ports)
    try:
        socket.create_connection((host, port_tls), timeout=5)
        details.append(f"✅ Réseau : Connexion TCP à {host}:{port_tls} réussie")
    except Exception as e:
        if _is_socket_blocked(e):
            details.append(f"❌ Réseau : PORT {port_tls} BLOQUÉ par le pare-feu Windows/entreprise")
            details.append(f"   → Solution : Demandez à votre administrateur réseau d'autoriser {host}:{port_tls}")
        else:
            details.append(f"❌ Réseau : Impossible de joindre {host}:{port_tls} → {e}")

    try:
        socket.create_connection((host, port_ssl), timeout=5)
        details.append(f"✅ Réseau : Connexion TCP à {host}:{port_ssl} réussie")
    except Exception as e:
        if _is_socket_blocked(e):
            details.append(f"❌ Réseau : PORT {port_ssl} aussi BLOQUÉ par le pare-feu")
        else:
            details.append(f"❌ Réseau : {host}:{port_ssl} inaccessible → {e}")

    # Step 2: Try SMTP auth
    app_password = app_password.replace(" ", "").strip()
    smtp_ok = False
    try:
        server = smtplib.SMTP(host, port_tls, timeout=10)
        server.starttls(context=ssl.create_default_context())
        server.login(sender_email, app_password)
        details.append(f"✅ Auth : Identifiants acceptés sur {host}:{port_tls}")
        smtp_ok = True
        # Send a real test email
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = "🔧 TEST EMAIL — SEBN Préventive App"
        msg.attach(MIMEText(
            "Ceci est un email de test envoyé depuis l'application SEBN-TN Préventive.\n"
            "Si vous recevez ce message, la configuration email est correcte.",
            "plain"
        ))
        server.sendmail(sender_email, [recipient_email], msg.as_string())
        server.quit()
        details.append(f"✅ Envoi : Email de test envoyé avec succès à {recipient_email}")
        return True, "\n".join(details)
    except Exception as e:
        if _is_socket_blocked(e):
            details.append(f"❌ SMTP 587 : BLOQUÉ par pare-feu ({e})")
        elif "Username and Password not accepted" in str(e) or "535" in str(e):
            details.append(f"❌ Auth : MOT DE PASSE D'APPLICATION INCORRECT → {e}")
            details.append("   → Solution : Créez un Mot de passe d'application Gmail (Sécurité > Connexion à Google)")
        else:
            details.append(f"❌ SMTP 587 : Échec → {e}")

    # Step 3: Try SMTP auth on 465 if 587 failed
    if not smtp_ok:
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=10) as server:
                server.login(sender_email, app_password)
                msg = MIMEMultipart()
                msg["From"] = sender_email
                msg["To"] = recipient_email
                msg["Subject"] = "🔧 TEST EMAIL — SEBN Préventive App"
                msg.attach(MIMEText("Email de test SEBN-TN.", "plain"))
                server.sendmail(sender_email, [recipient_email], msg.as_string())
            details.append(f"✅ Envoi via port 465 : Email de test envoyé à {recipient_email}")
            return True, "\n".join(details)
        except Exception as e:
            if _is_socket_blocked(e):
                details.append(f"❌ SMTP 465 : aussi BLOQUÉ par pare-feu ({e})")
            else:
                details.append(f"❌ SMTP 465 : Échec → {e}")

    # Step 4: Outlook COM fallback test
    try:
        import win32com.client as win32
        import pythoncom
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = recipient_email
        mail.Subject = "🔧 TEST EMAIL via Outlook — SEBN Préventive App"
        mail.Body = "Email de test envoyé via Microsoft Outlook (méthode de secours)."
        mail.Send()
        details.append("✅ Fallback : Email envoyé via Microsoft Outlook (COM)")
        return True, "\n".join(details)
    except Exception as e:
        details.append(f"❌ Outlook COM : Non disponible ou refusé → {e}")

    details.append("")
    details.append("⚠️  DIAGNOSTIC : Tous les canaux sont bloqués.")
    details.append("   Le pare-feu ou l'antivirus d'entreprise interdit les connexions SMTP sortantes.")
    details.append("   Contactez votre administrateur réseau pour débloquer smtp.gmail.com (ports 587/465).")
    return False, "\n".join(details)


def send_email_with_attachment(sender_email, app_password, recipient_email, subject, body, attachment_path=None):
    """
    Sends an email with a file attachment.
    Priority: 1. SMTP (Smart Detection) -> 2. Outlook COM (Fallback)
    """
    try:
        # Create the root message
        message = MIMEMultipart()
        message["From"] = sender_email
        message["Subject"] = subject
        
        if isinstance(recipient_email, list):
            recipients_str = ", ".join(recipient_email)
        else:
            recipients_str = recipient_email
            
        message["To"] = recipients_str
        message.attach(MIMEText(body, "plain"))

        if attachment_path and not os.path.exists(attachment_path):
            return False, f"Attachment not found: {attachment_path}"

        # Clean password
        app_password = app_password.replace(" ", "").strip()
        host, port_tls, _ = get_smtp_config(sender_email)

        # 1. ATTEMPT PRIMARY SMTP
        try:
            # Add attachment if provided
            if attachment_path:
                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
                message.attach(part)

            print(f"Attempting Primary SMTP ({host}:{port_tls})...")
            server = smtplib.SMTP(host, port_tls, timeout=12)
            server.starttls(context=ssl.create_default_context())
            server.login(sender_email, app_password)
            
            send_to = [r.strip() for r in recipients_str.split(',') if r.strip()]
            server.sendmail(sender_email, send_to, message.as_string())
            server.quit()
            return True, f"Email envoyé via SMTP ({host})"
            
        except Exception as smtp_err:
            print(f"SMTP Failed: {smtp_err}")
            
            # 2. ATTEMPT OUTLOOK FALLBACK (Guarantees delivery for most users)
            try:
                print("Attempting Outlook COM Fallback...")
                pythoncom.CoInitialize()
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                
                mail.To = recipients_str
                mail.Subject = subject
                mail.Body = body
                
                if attachment_path and os.path.exists(attachment_path):
                    mail.Attachments.Add(os.path.abspath(attachment_path))
                
                mail.Send()
                return True, "Email envoyé via Microsoft Outlook (Méthode de secours)"
            except Exception as outlook_err:
                print(f"Outlook Fallback Failed: {outlook_err}")
                return False, f"Échec de l'envoi (SMTP: {smtp_err} | Outlook: {outlook_err})"
            finally:
                try: pythoncom.CoUninitialize()
                except: pass

    except Exception as e:
        return False, f"Erreur critique lors de l'envoi : {str(e)}"

    except Exception as e:
        return False, f"Unexpected Email Error: {str(e)}"

def send_simple_alert(sender_email, app_password, recipient_email, subject, body):
    """
    Sends a simple text alert email.
    """
    try:
        # Create the root message
        message = MIMEMultipart()
        message["From"] = sender_email
        message["Subject"] = subject
        
        if isinstance(recipient_email, list):
            recipients_str = ", ".join(recipient_email)
        else:
            recipients_str = recipient_email
            
        message["To"] = recipients_str
        message.attach(MIMEText(body, "plain"))

        # Try SMTP (587)
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587, timeout=10)
            server.starttls(context=ssl.create_default_context())
            server.login(sender_email, app_password)
            send_to = [r.strip() for r in recipients_str.split(',') if r.strip()]
            server.sendmail(sender_email, send_to, message.as_string())
            server.quit()
            return True, "Alert sent successfully!"
        except Exception as smtp_err:
            # Fallback to Outlook if on Windows
            try:
                import win32com.client as win32
                import pythoncom
                pythoncom.CoInitialize()
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = recipients_str
                mail.Subject = subject
                mail.Body = body
                mail.Send()
                return True, "Alert sent via Outlook!"
            except:
                return False, f"Email failed: {smtp_err}"
    except Exception as e:
        return False, f"Error: {e}"

if __name__ == "__main__":
    # Test block
    SENDER = input("Enter Gmail: ")
    PWD = input("Enter App Password: ")
    RECEIVER = input("Enter Recipient Email: ")
    FILE = input("Enter PDF file path: ")
    
    success, msg = send_email_with_attachment(SENDER, PWD, RECEIVER, "Test Report", "Please find the report attached.", FILE)
    print(msg)
