"""
checklist_pdf_generator.py
===========================
Generates a permanent, officially branded SEBN-TN PDF report for completed
preventive maintenance checklists. Saved once upon web checklist completion
and served directly without regeneration.
"""

import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

def generate_completed_checklist_pdf(exec_record: dict, items: list, output_pdf_path: str) -> bool:
    """
    Generates a high-quality PDF report for a completed web checklist.
    Returns True if generated successfully.
    """
    os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
    
    doc = SimpleDocTemplate(
        output_pdf_path,
        pagesize=A4,
        leftMargin=30,
        rightMargin=30,
        topMargin=28,
        bottomMargin=28
    )
    
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=18,
        textColor=colors.HexColor('#2E008B'),
        alignment=TA_LEFT
    )
    subtitle_style = ParagraphStyle(
        'DocSub',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#64748B'),
        alignment=TA_LEFT
    )
    badge_style = ParagraphStyle(
        'Badge',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor('#065F46'),
        alignment=TA_RIGHT
    )
    cell_desc = ParagraphStyle(
        'CellDesc',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#1E293B')
    )
    cell_method = ParagraphStyle(
        'CellMethod',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#475569')
    )
    cell_obs = ParagraphStyle(
        'CellObs',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#475569')
    )
    cell_bold_center = ParagraphStyle(
        'CellBoldCenter',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_CENTER
    )
    th_style = ParagraphStyle(
        'TH',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.white,
        alignment=TA_CENTER
    )
    meta_lbl = ParagraphStyle(
        'MetaLbl',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor('#64748B')
    )
    meta_val = ParagraphStyle(
        'MetaVal',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=10.5,
        textColor=colors.HexColor('#0F172A')
    )

    story = []
    
    # ── 1. HEADER SECTION ──
    header_data = [
        [
            Paragraph("<b>SEBN-TN MAINTENANCE SYSTEM</b><br/><font size=8.5 color='#2E008B'>Fiche d'Exécution — Maintenance Préventive</font>", title_style),
            Paragraph("<font color='#065F46'><b>RAPPORT OFFICIEL</b></font><br/><font size=8 color='#64748B'>Exécuté via Portail Web</font>", badge_style)
        ]
    ]
    t_header = Table(header_data, colWidths=[380, 155])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#2E008B'), spaceBefore=2, spaceAfter=8))

    # ── 2. METADATA BOX ──
    equip = exec_record.get('equipment') or '—'
    task_type = exec_record.get('task_type') or '—'
    sheet = exec_record.get('sheet') or '—'
    week = exec_record.get('week') or '—'
    tech_name = exec_record.get('technician_name') or 'Technicien'
    matricule = exec_record.get('technician_matricule') or '—'
    shift = exec_record.get('shift') or 'A'
    exec_at = exec_record.get('executed_at') or datetime.now().strftime('%Y-%m-%d %H:%M')
    chk_name = exec_record.get('checklist_name') or 'Checklist Standard'

    meta_grid = [
        [
            Paragraph("Équipement / Machine :", meta_lbl),
            Paragraph(f"<b>{equip}</b>", meta_val),
            Paragraph("Zone / Ligne :", meta_lbl),
            Paragraph(f"<b>{sheet}</b>", meta_val),
        ],
        [
            Paragraph("Périodicité & Semaine :", meta_lbl),
            Paragraph(f"<b>{task_type} — Semaine {week}</b>", meta_val),
            Paragraph("Shift d'Exécution :", meta_lbl),
            Paragraph(f"<b>Shift {shift}</b>", meta_val),
        ],
        [
            Paragraph("Technicien Déclarant :", meta_lbl),
            Paragraph(f"<b>{tech_name}</b> ({matricule})", meta_val),
            Paragraph("Date & Heure :", meta_lbl),
            Paragraph(f"<b>{exec_at}</b>", meta_val),
        ],
    ]
    t_meta = Table(meta_grid, colWidths=[120, 150, 110, 155])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 0.75, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 10))

    # ── 3. INSPECTION TABLE ──
    answers = exec_record.get('answers') or {}
    
    col_widths = [26, 214, 110, 50, 60, 75]
    table_rows = [
        [
            Paragraph("#", th_style),
            Paragraph("Point de Contrôle / Description", th_style),
            Paragraph("Méthode / Moyen", th_style),
            Paragraph("Résultat", th_style),
            Paragraph("Valeur", th_style),
            Paragraph("Observations", th_style),
        ]
    ]

    ok_count = 0
    nok_count = 0
    na_count = 0

    for idx, item in enumerate(items, 1):
        num = item.get('number') or item.get('item_number') or idx
        desc = item.get('description', '')
        method = item.get('method', 'Contrôle visuel / standard')
        
        ans = answers.get(str(num)) or answers.get(int(num)) or {}
        st = str(ans.get('status', '')).upper().strip()
        val = str(ans.get('val', '')).strip()
        obs = str(ans.get('obs', '')).strip()

        # Format status badge
        if st in ('OK', 'CONFORME', 'BON'):
            ok_count += 1
            st_cell = "<font color='#059669'><b>✔ OK</b></font>"
        elif st in ('NOK', 'NON CONFORME', 'DEFAUT'):
            nok_count += 1
            st_cell = "<font color='#DC2626'><b>✖ NOK</b></font>"
        elif st in ('N/A', 'NA'):
            na_count += 1
            st_cell = "<font color='#64748B'><b>— N/A</b></font>"
        else:
            st_cell = f"<font color='#64748B'>{st or '—'}</font>"

        table_rows.append([
            Paragraph(str(num), cell_bold_center),
            Paragraph(desc.replace('\n', '<br/>'), cell_desc),
            Paragraph(method.replace('\n', '<br/>'), cell_method),
            Paragraph(st_cell, cell_bold_center),
            Paragraph(val or '—', cell_bold_center),
            Paragraph(obs or '—', cell_obs),
        ])

    t_items = Table(table_rows, colWidths=col_widths, repeatRows=1)
    
    t_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E008B')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('BOX', (0,0), (-1,-1), 0.75, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
    ]
    # Alternating row colors
    for r in range(1, len(table_rows)):
        bg = colors.HexColor('#F8FAFC') if r % 2 == 0 else colors.white
        t_style.append(('BACKGROUND', (0, r), (-1, r), bg))

    t_items.setStyle(TableStyle(t_style))
    story.append(t_items)
    story.append(Spacer(1, 12))

    # ── 4. SUMMARY & SIGNATURE BLOCK ──
    total_pts = len(items)
    summary_html = f"Total Points : <b>{total_pts}</b> &nbsp;|&nbsp; Conformes (OK) : <font color='#059669'><b>{ok_count}</b></font> &nbsp;|&nbsp; Non Conformes (NOK) : <font color='#DC2626'><b>{nok_count}</b></font> &nbsp;|&nbsp; Non Applicables (N/A) : <b>{na_count}</b>"
    
    footer_table_data = [
        [
            Paragraph(summary_html, cell_desc),
            Paragraph("<b>VISA & VALIDATION SYSTÈME</b>", ParagraphStyle('FootH', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=TA_RIGHT, textColor=colors.HexColor('#2E008B')))
        ],
        [
            Paragraph(f"<font color='#64748B'>Fichier source : {os.path.basename(exec_record.get('filled_excel_path', ''))}</font>", subtitle_style),
            Paragraph(f"Validé électroniquement par <b>{tech_name}</b><br/><font size=7 color='#64748B'>Horodatage : {exec_at}</font>", ParagraphStyle('FootSig', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=TA_RIGHT))
        ]
    ]
    t_foot = Table(footer_table_data, colWidths=[335, 200])
    t_foot.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.75, colors.HexColor('#C7D2FE')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#EEF2FF')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    
    story.append(KeepTogether([t_foot]))

    try:
        doc.build(story)
        return True
    except Exception as e:
        import logging
        logging.getLogger("sebn-maintenance").error(f"[CHECKLIST-PDF] Error building PDF: {e}", exc_info=True)
        return False
