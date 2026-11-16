"""
email_service.py
================
Enterprise Email & SMTP Notification Service for SEBN-TN Maintenance Portal.
Supports:
- Custom Factory SMTP Servers (IP / hostname, custom ports 25, 465, 587)
- Encryption modes: STARTTLS, SSL/TLS, or Unencrypted Internal Relay (None)
- Authenticated and unauthenticated (anonymous factory relay) configurations
- Rich responsive HTML templates with SEBN-TN branding
- Asynchronous notifications for intervention startup & critical events
- Preventive schedule Excel file dispatch to Team Leaders
- Diagnostic SMTP testing with detailed error logging
- Windows Outlook COM fallback when available
"""

import os
import re
import ssl
import socket
import smtplib
import logging
import threading
import datetime
from email import encoders
from email.utils import formatdate
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Tuple, Optional, Dict, Any

logger = logging.getLogger("sebn-email")


# ── HELPER UTILITIES ───────────────────────────────────────────────────────────

def parse_recipients(recipient_input: Any) -> List[str]:
    """Parse comma, semicolon, newline or space-separated recipients into a clean list of valid emails."""
    if not recipient_input:
        return []
    if isinstance(recipient_input, list):
        raw_items = recipient_input
    else:
        raw_items = re.split(r'[,;\n\r]+', str(recipient_input))
    
    clean = []
    for r in raw_items:
        r_str = str(r).strip()
        if r_str and '@' in r_str and '.' in r_str:
            clean.append(r_str)
    return list(dict.fromkeys(clean))  # remove duplicates while preserving order


def _is_socket_blocked(err: Exception) -> bool:
    """Detect Windows Firewall or enterprise network socket permission errors (e.g. WinError 10013)."""
    err_str = str(err).lower()
    return any(k in err_str for k in ["10013", "access", "permission", "blocked", "not permitted", "refused", "timed out"])


def get_default_smtp_config(email: str = "") -> Tuple[str, int, str]:
    """Guess default server, port and security from email domain if not specified."""
    email = (email or "").lower().strip()
    if "@gmail.com" in email:
        return "smtp.gmail.com", 587, "STARTTLS"
    elif any(d in email for d in ["@outlook.", "@hotmail.", "@live.", "@office365."]):
        return "smtp.office365.com", 587, "STARTTLS"
    elif "@yahoo." in email:
        return "smtp.mail.yahoo.com", 587, "STARTTLS"
    return "smtp.gmail.com", 587, "STARTTLS"


# ── CORE EMAIL SENDER ─────────────────────────────────────────────────────────

def send_email_core(
    smtp_server: str,
    smtp_port: int,
    security: str,
    sender_email: str,
    app_password: str,
    recipients: Any,
    subject: str,
    body_text: str,
    body_html: Optional[str] = None,
    attachment_path: Optional[str] = None,
    attachment_filename: Optional[str] = None,
    timeout: int = 15
) -> Tuple[bool, str]:
    """
    Core sender handling all SMTP flavors (Factory relay, STARTTLS, SSL/TLS, None)
    with optional attachment and HTML alternative.
    """
    recip_list = parse_recipients(recipients)
    if not recip_list:
        return False, "Aucun destinataire valide spécifié."
    
    sender_email = (sender_email or "").strip()
    smtp_server = (smtp_server or "").strip()
    try:
        smtp_port = int(smtp_port)
    except (ValueError, TypeError):
        smtp_port = 587

    security = (security or "STARTTLS").upper().strip()
    if security not in ("STARTTLS", "SSL", "NONE"):
        security = "SSL" if smtp_port == 465 else "STARTTLS"

    # Build MIME Structure
    msg_root = MIMEMultipart("mixed")
    msg_root["From"] = sender_email
    msg_root["To"] = ", ".join(recip_list)
    msg_root["Subject"] = subject
    msg_root["Date"] = formatdate(localtime=True)

    # Alternative body: plain text + optional HTML
    msg_alt = MIMEMultipart("alternative")
    if body_text:
        msg_alt.attach(MIMEText(body_text, "plain", "utf-8"))
    if body_html:
        msg_alt.attach(MIMEText(body_html, "html", "utf-8"))
    elif body_text:
        # Wrap plain text in clean HTML container if no HTML was passed
        fallback_html = f"<div style='font-family:Segoe UI,Helvetica,Arial,sans-serif;font-size:14px;color:#1E293B;white-space:pre-wrap;'>{body_text}</div>"
        msg_alt.attach(MIMEText(fallback_html, "html", "utf-8"))
    msg_root.attach(msg_alt)

    # Attach file if present
    if attachment_path:
        if os.path.isfile(attachment_path):
            try:
                fname = attachment_filename or os.path.basename(attachment_path)
                with open(attachment_path, "rb") as af:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(af.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{fname}"')
                msg_root.attach(part)
            except Exception as e:
                logger.error(f"Erreur attachement fichier {attachment_path}: {e}")
        else:
            logger.warning(f"Fichier joint introuvable : {attachment_path}")

    # 1. ATTEMPT PRIMARY SMTP
    smtp_error = None
    try:
        server = None
        # Mode SSL (ex: port 465)
        if security == "SSL":
            context = ssl.create_default_context()
            try:
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=timeout, context=context)
            except ssl.SSLError:
                # Factory relays might use self-signed certs
                context = ssl._create_unverified_context()
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=timeout, context=context)

        # Mode STARTTLS (ex: port 587 ou 25)
        elif security == "STARTTLS":
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=timeout)
            server.ehlo()
            try:
                context = ssl.create_default_context()
                server.starttls(context=context)
                server.ehlo()
            except ssl.SSLError:
                # Try unverified if self-signed cert in corporate LAN
                context = ssl._create_unverified_context()
                server.starttls(context=context)
                server.ehlo()
            except Exception as tls_err:
                logger.warning(f"STARTTLS non supporté sur {smtp_server}:{smtp_port} -> tentative en clair: {tls_err}")

        # Mode NONE (ex: port 25 relais interne usine sans chiffrement)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=timeout)
            server.ehlo()

        # Authenticate if password provided (factory relays often do not require authentication)
        cleaned_pw = (app_password or "").replace(" ", "").strip()
        if cleaned_pw and sender_email:
            try:
                server.login(sender_email, cleaned_pw)
            except smtplib.SMTPAuthenticationError as auth_err:
                server.quit()
                return False, f"Échec d'authentification SMTP ({auth_err.smtp_code}: {auth_err.smtp_error})"

        # Send
        server.sendmail(sender_email, recip_list, msg_root.as_string())
        try:
            server.quit()
        except Exception:
            pass

        return True, f"Email envoyé avec succès via {smtp_server}:{smtp_port} à {len(recip_list)} destinataire(s)."

    except Exception as e:
        smtp_error = e
        logger.error(f"SMTP Error on {smtp_server}:{smtp_port} -> {e}")

    # 2. ATTEMPT OUTLOOK COM FALLBACK (Windows clients)
    try:
        import win32com.client as win32
        import pythoncom
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = "; ".join(recip_list)
        mail.Subject = subject
        if body_html:
            mail.HTMLBody = body_html
        else:
            mail.Body = body_text
        if attachment_path and os.path.isfile(attachment_path):
            mail.Attachments.Add(os.path.abspath(attachment_path))
        mail.Send()
        pythoncom.CoUninitialize()
        return True, "Email envoyé avec succès via Microsoft Outlook (Méthode de secours locale)."
    except Exception as outlook_err:
        logger.warning(f"Outlook fallback unavailable or failed: {outlook_err}")

    return False, f"Échec de l'envoi email (SMTP: {smtp_error})"


# ── DIAGNOSTIC CONNECTION TEST ────────────────────────────────────────────────

def test_email_connection(
    sender_email: str,
    app_password: str,
    recipient_email: str,
    smtp_server: Optional[str] = None,
    smtp_port: Optional[int] = None,
    security: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Runs diagnostic tests against the configured SMTP server and sends a real test message.
    Returns (success: bool, detailed_log: str).
    """
    details = []
    
    sender_email = (sender_email or "").strip()
    recipient_email = (recipient_email or "").strip()
    
    # Defaults if missing
    def_host, def_port, def_sec = get_default_smtp_config(sender_email)
    host = (smtp_server or def_host).strip()
    try:
        port = int(smtp_port or def_port)
    except (ValueError, TypeError):
        port = 587
    sec = (security or def_sec).upper().strip()

    details.append(f"🔍 Démarrage du test SMTP :")
    details.append(f"   • Serveur   : {host}")
    details.append(f"   • Port      : {port}")
    details.append(f"   • Sécurité  : {sec}")
    details.append(f"   • Expéditeur: {sender_email or '(Non renseigné)'}")
    details.append(f"   • Cible     : {recipient_email or '(Non renseigné)'}")

    if not recipient_email:
        details.append("❌ Aucun destinataire de test renseigné.")
        return False, "\n".join(details)

    # Step 1: TCP Socket Reachability
    try:
        s = socket.create_connection((host, port), timeout=6)
        s.close()
        details.append(f"✅ Réseau : Connexion TCP vers {host}:{port} réussie.")
    except Exception as e:
        if _is_socket_blocked(e):
            details.append(f"❌ Réseau : Port {port} BLOQUÉ par le pare-feu ou routeur d'entreprise ({e})")
            details.append(f"   → Demandez à l'administrateur réseau d'autoriser le port {port} vers {host}.")
        else:
            details.append(f"❌ Réseau : Impossible de joindre {host}:{port} → {e}")
        return False, "\n".join(details)

    # Step 2: Try Send Test Email
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    subject = "⚙️ [TEST] Vérification Notification Email — SEBN-TN Maintenance"
    
    body_text = (
        f"Ceci est un message de test envoyé depuis le portail de maintenance SEBN-TN.\n\n"
        f"Date & Heure : {now_str}\n"
        f"Serveur SMTP : {host}:{port} ({sec})\n"
        f"Expéditeur   : {sender_email}\n\n"
        f"Si vous recevez cet email, la passerelle SMTP est fonctionnelle !"
    )

    body_html = f"""
    <div style="max-width:580px; margin:0 auto; font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif; border:1px solid #E2E8F0; border-radius:12px; overflow:hidden; box-shadow:0 4px 16px rgba(0,0,0,0.06); background:#FFFFFF;">
        <div style="background:linear-gradient(135deg, #2E008B 0%, #1A004F 100%); padding:22px 26px; color:#FFFFFF;">
            <div style="font-size:11px; font-weight:800; letter-spacing:1.5px; opacity:0.85; text-transform:uppercase;">SEBN-TN Enterprise Maintenance</div>
            <h2 style="margin:6px 0 0; font-size:18px; font-weight:800; color:#FFFFFF;">Test de Connexion Email Réussi</h2>
        </div>
        <div style="padding:24px 26px;">
            <p style="font-size:14px; color:#334155; line-height:1.6; margin-top:0;">
                Félicitations ! Les paramètres SMTP configurés sur votre portail de maintenance sont <strong>parfaitement opérationnels</strong>.
            </p>
            <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:16px; margin:18px 0; font-size:13px; line-height:1.7; color:#1E293B;">
                <div><strong>Serveur SMTP :</strong> <code>{host}:{port}</code></div>
                <div><strong>Mode de Sécurité :</strong> <span style="background:#EEF2FF; color:#4338CA; padding:2px 8px; border-radius:4px; font-weight:700;">{sec}</span></div>
                <div><strong>Expéditeur :</strong> {sender_email}</div>
                <div><strong>Destinataire :</strong> {recipient_email}</div>
                <div><strong>Horodatage :</strong> {now_str}</div>
            </div>
            <p style="font-size:12px; color:#64748B; margin-bottom:0;">
                Les alertes d'interventions et les rapports préventifs peuvent désormais être expédiés automatiquement.
            </p>
        </div>
        <div style="background:#F1F5F9; border-top:1px solid #E2E8F0; padding:12px 26px; font-size:11px; color:#94A3B8; text-align:center;">
            SEBN-TN Maintenance Portal &bull; Bizerte, Tunisie
        </div>
    </div>
    """

    ok, msg = send_email_core(
        smtp_server=host,
        smtp_port=port,
        security=sec,
        sender_email=sender_email,
        app_password=app_password,
        recipients=recipient_email,
        subject=subject,
        body_text=body_text,
        body_html=body_html
    )

    if ok:
        details.append(f"✅ Envoi : {msg}")
        return True, "\n".join(details)
    else:
        details.append(f"❌ Envoi échoué : {msg}")
        return False, "\n".join(details)


# ── HIGH-LEVEL WORKFLOW DISPATCHERS ───────────────────────────────────────────

def notify_intervention_started_async(
    intervention_data: Dict[str, Any],
    code: str,
    data_dir: str,
    base_url: str = "http://localhost:5000"
) -> None:
    """
    Sends an immediate email notification when an intervention is declared/started.
    Runs asynchronously in a background thread to prevent UI freezing.
    """
    def _run():
        try:
            cfg_path = os.path.join(data_dir, "email_config.json")
            if not os.path.isfile(cfg_path):
                return
            
            import json
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)

            # Check if notification is enabled
            if not cfg.get("notify_intervention_start", True):
                logger.info("Notification intervention désactivée dans la configuration.")
                return

            recipients = cfg.get("team_leader_emails") or cfg.get("recipient")
            recip_list = parse_recipients(recipients)
            if not recip_list:
                logger.info("Aucun destinataire Team Leader configuré pour notification d'intervention.")
                return

            sender = cfg.get("sender") or cfg.get("sender_email", "")
            pw = cfg.get("password") or cfg.get("sender_password", "")
            host = cfg.get("smtp_server", "smtp.gmail.com")
            port = int(cfg.get("smtp_port", 587))
            sec = cfg.get("security", "STARTTLS")

            mach_name = intervention_data.get('machine_name') or intervention_data.get('machine_id', 'Machine Inconnue')
            mach_id = intervention_data.get('machine_id', '')
            fault_desc = intervention_data.get('fault_description', 'Non spécifiée')
            priority = str(intervention_data.get('priority', 'MEDIUM')).upper()
            tech_name = intervention_data.get('technician_name', 'Technicien')
            tech_mat = intervention_data.get('technician_mat', '')
            shift = intervention_data.get('shift', 'A')
            start_time = intervention_data.get('start_time', datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))

            prio_color = "#DC2626" if priority in ("BLOQUANTE", "HIGH", "CRITIQUE") else "#F59E0B" if priority in ("MEDIUM", "MOYENNE") else "#2563EB"
            url_link = f"{base_url.rstrip('/')}/interventions/{code}"
            mat_html = f'<span style="color:#64748B; font-size:11px;">(Mat: {tech_mat})</span>' if tech_mat else ''

            subject = f"🚨 [INTERVENTION DÉMARRÉE] {mach_name} ({mach_id}) — Réf: #{code}"

            body_text = (
                f"NOUVELLE INTERVENTION DE MAINTENANCE CURATIVE\n"
                f"============================================\n"
                f"Réf: #{code}\n"
                f"Machine: {mach_name} ({mach_id})\n"
                f"Panne constatée: {fault_desc}\n"
                f"Priorité: {priority}\n"
                f"Technicien: {tech_name} ({tech_mat})\n"
                f"Shift: {shift}\n"
                f"Heure de début: {start_time}\n"
                f"Lien: {url_link}\n"
            )

            body_html = f"""
            <div style="max-width:600px; margin:0 auto; font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif; border:1px solid #E2E8F0; border-radius:12px; overflow:hidden; background:#FFFFFF; box-shadow:0 4px 20px rgba(0,0,0,0.06);">
                <div style="background:linear-gradient(135deg, #1E1B4B 0%, #2E008B 100%); padding:22px 28px; color:#FFFFFF;">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span style="font-size:11px; font-weight:800; letter-spacing:1.5px; opacity:0.8; text-transform:uppercase;">SEBN-TN Maintenance Curative</span>
                        <span style="background:{prio_color}; color:#FFFFFF; font-size:10px; font-weight:800; padding:3px 10px; border-radius:12px; text-transform:uppercase;">Priorité {priority}</span>
                    </div>
                    <h2 style="margin:10px 0 2px; font-size:20px; font-weight:800; color:#FFFFFF;">Intervention #{code} Démarrée</h2>
                    <div style="font-size:13px; opacity:0.9;">Équipement concerné : <strong>{mach_name}</strong></div>
                </div>
                
                <div style="padding:24px 28px;">
                    <table style="width:100%; border-collapse:collapse; font-size:13px; color:#1E293B; margin-bottom:20px;">
                        <tr style="border-bottom:1px solid #F1F5F9;">
                            <td style="padding:8px 0; color:#64748B; width:35%;">Machine / ID</td>
                            <td style="padding:8px 0; font-weight:700;">{mach_name} <span style="font-family:monospace; color:#64748B;">({mach_id})</span></td>
                        </tr>
                        <tr style="border-bottom:1px solid #F1F5F9;">
                            <td style="padding:8px 0; color:#64748B;">Description Panne</td>
                            <td style="padding:8px 0; font-weight:600; color:#0F172A;">{fault_desc}</td>
                        </tr>
                        <tr style="border-bottom:1px solid #F1F5F9;">
                            <td style="padding:8px 0; color:#64748B;">Déclarée par</td>
                            <td style="padding:8px 0; font-weight:600;">{tech_name} {mat_html}</td>
                        </tr>
                        <tr style="border-bottom:1px solid #F1F5F9;">
                            <td style="padding:8px 0; color:#64748B;">Équipe / Shift</td>
                            <td style="padding:8px 0; font-weight:700;"><span style="background:#F1F5F9; padding:2px 8px; border-radius:4px;">Équipe {shift}</span></td>
                        </tr>
                        <tr>
                            <td style="padding:8px 0; color:#64748B;">Heure de Début</td>
                            <td style="padding:8px 0; font-weight:600; color:#334155;">{start_time}</td>
                        </tr>
                    </table>

                    <div style="text-align:center; margin:26px 0 10px;">
                        <a href="{url_link}" style="background:#2E008B; color:#FFFFFF; padding:12px 28px; border-radius:8px; text-decoration:none; font-weight:700; font-size:14px; display:inline-block; box-shadow:0 3px 12px rgba(46,0,139,0.25);">
                            Consulter l'Intervention sur le Portail &rarr;
                        </a>
                    </div>
                </div>

                <div style="background:#F8FAFC; border-top:1px solid #E2E8F0; padding:14px 28px; font-size:11px; color:#94A3B8; text-align:center;">
                    Ce message a été généré automatiquement à la déclaration de l'intervention.
                </div>
            </div>
            """

            send_email_core(
                smtp_server=host,
                smtp_port=port,
                security=sec,
                sender_email=sender,
                app_password=pw,
                recipients=recip_list,
                subject=subject,
                body_text=body_text,
                body_html=body_html
            )
        except Exception as ex:
            logger.error(f"Erreur notification asynchrone intervention #{code}: {ex}")

    threading.Thread(target=_run, daemon=True).start()


def send_preventive_schedule_email(
    data_dir: str,
    custom_excel_path: Optional[str] = None,
    recipient_override: Optional[str] = None,
    custom_note: str = ""
) -> Tuple[bool, str]:
    """
    Sends the active PMA schedule Excel file to the configured Team Leader(s).
    """
    cfg_path = os.path.join(data_dir, "email_config.json")
    if not os.path.isfile(cfg_path):
        return False, "Configuration email non initialisée."

    import json
    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    recipients = recipient_override or cfg.get("team_leader_emails") or cfg.get("recipient")
    recip_list = parse_recipients(recipients)
    if not recip_list:
        return False, "Aucune adresse email de Team Leader configurée. Rendez-vous dans Paramètres Administrateur > Notifications Email."

    # Find the Excel file
    excel_file = None
    if custom_excel_path and os.path.isfile(custom_excel_path):
        excel_file = custom_excel_path
    else:
        last_file_record = os.path.join(data_dir, "last_excel.txt")
        if os.path.isfile(last_file_record):
            with open(last_file_record, 'r', encoding='utf-8') as lf:
                candidate = lf.read().strip()
                if candidate and os.path.isfile(candidate):
                    excel_file = candidate

    if not excel_file or not os.path.isfile(excel_file):
        # Look for default files
        candidates = ["current_schedule.xlsx", "calendrier.xlsx", "planning.xlsx"]
        for c in candidates:
            p = os.path.join(data_dir, c)
            if os.path.isfile(p):
                excel_file = p
                break

    if not excel_file:
        return False, "Aucun fichier Excel de planning préventif actif trouvé sur le serveur."

    sender = cfg.get("sender") or cfg.get("sender_email", "")
    pw = cfg.get("password") or cfg.get("sender_password", "")
    host = cfg.get("smtp_server", "smtp.gmail.com")
    port = int(cfg.get("smtp_port", 587))
    sec = cfg.get("security", "STARTTLS")

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    filename = os.path.basename(excel_file)
    filesize_kb = round(os.path.getsize(excel_file) / 1024, 1)

    subject = f"📋 [PLANNING PRÉVENTIF] Planning PMA SEBN-TN ({filename})"
    note_line = f"Note : {custom_note}\n" if custom_note else ""
    note_block = f'<div style="background:#F8FAFC; border-left:4px solid #10B981; padding:10px 14px; margin-bottom:20px; font-size:13px; color:#475569;"><strong>Note :</strong> {custom_note}</div>' if custom_note else ''

    body_text = (
        f"Bonjour Team Leader,\n\n"
        f"Veuillez trouver ci-joint le fichier Excel du planning préventif annuel/mensuel (PMA).\n\n"
        f"Fichier : {filename} ({filesize_kb} Ko)\n"
        f"Date d'envoi : {now_str}\n"
        f"{note_line}\n"
        f"Cordialement,\n"
        f"Système de Maintenance SEBN-TN"
    )

    body_html = f"""
    <div style="max-width:600px; margin:0 auto; font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif; border:1px solid #E2E8F0; border-radius:12px; overflow:hidden; background:#FFFFFF; box-shadow:0 4px 20px rgba(0,0,0,0.06);">
        <div style="background:linear-gradient(135deg, #10B981 0%, #047857 100%); padding:22px 28px; color:#FFFFFF;">
            <div style="font-size:11px; font-weight:800; letter-spacing:1.5px; opacity:0.85; text-transform:uppercase;">SEBN-TN Maintenance Préventive</div>
            <h2 style="margin:8px 0 2px; font-size:20px; font-weight:800; color:#FFFFFF;">Planning Préventif (PMA) — Fichier Excel</h2>
            <div style="font-size:13px; opacity:0.9;">Export automatique pour les Team Leaders</div>
        </div>
        
        <div style="padding:24px 28px;">
            <p style="font-size:14px; color:#334155; line-height:1.6; margin-top:0;">
                Bonjour Team Leader,<br>
                Vous trouverez ci-joint la dernière version du <strong>planning de maintenance préventive</strong> sous format Excel.
            </p>
            
            <div style="background:#F0FDF4; border:1px solid #BBF7D0; border-radius:10px; padding:16px 20px; margin:20px 0; display:flex; align-items:center; gap:16px;">
                <div style="font-size:28px;">📊</div>
                <div>
                    <div style="font-weight:700; color:#065F46; font-size:14px;">{filename}</div>
                    <div style="font-size:12px; color:#047857;">Taille du fichier : {filesize_kb} Ko &bull; Horodatage : {now_str}</div>
                </div>
            </div>

            {note_block}

            <p style="font-size:13px; color:#64748B; margin-bottom:0;">
                Pour consulter l'état d'avancement des interventions préventives en temps réel, connectez-vous au portail web de maintenance.
            </p>
        </div>

        <div style="background:#F8FAFC; border-top:1px solid #E2E8F0; padding:14px 28px; font-size:11px; color:#94A3B8; text-align:center;">
            SEBN-TN Maintenance Portal &bull; Envoi officiel sécurisé
        </div>
    </div>
    """

    return send_email_core(
        smtp_server=host,
        smtp_port=port,
        security=sec,
        sender_email=sender,
        app_password=pw,
        recipients=recip_list,
        subject=subject,
        body_text=body_text,
        body_html=body_html,
        attachment_path=excel_file,
        attachment_filename=filename
    )


# ── BACKWARDS COMPATIBILITY WRAPPERS ─────────────────────────────────────────

def send_simple_alert(sender_email, app_password, recipient_email, subject, body, smtp_server=None, smtp_port=None, security=None):
    """Backwards-compatible wrapper for sending simple text alerts."""
    def_host, def_port, def_sec = get_default_smtp_config(sender_email)
    host = smtp_server or def_host
    port = smtp_port or def_port
    sec = security or def_sec
    ok, msg = send_email_core(
        smtp_server=host,
        smtp_port=port,
        security=sec,
        sender_email=sender_email,
        app_password=app_password,
        recipients=recipient_email,
        subject=subject,
        body_text=body
    )
    return ok, msg


def send_email_with_attachment(sender_email, app_password, recipient_email, subject, body, attachment_path=None, smtp_server=None, smtp_port=None, security=None):
    """Backwards-compatible wrapper for sending attachments."""
    def_host, def_port, def_sec = get_default_smtp_config(sender_email)
    host = smtp_server or def_host
    port = smtp_port or def_port
    sec = security or def_sec
    ok, msg = send_email_core(
        smtp_server=host,
        smtp_port=port,
        security=sec,
        sender_email=sender_email,
        app_password=app_password,
        recipients=recipient_email,
        subject=subject,
        body_text=body,
        attachment_path=attachment_path
    )
    return ok, msg
