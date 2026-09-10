import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

def generate_technician_card_pdf(tech_details: dict) -> str:
    """
    Generates a professional printable A4 PDF containing the Technician ID card/badge.
    The layout uses SEBN corporate colors (Deep Purple #2E008B and clean white).
    """
    export_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'data', 'documents', 'certificates')
    os.makedirs(export_dir, exist_ok=True)
    
    mat = tech_details.get('matricule') or f"user_{tech_details.get('id')}"
    safe_mat = str(mat).replace("/", "_").replace("\\", "_").strip()
    
    filename = f"Carte_Technicien_{safe_mat}.pdf"
    filepath = os.path.join(export_dir, filename)
    
    # We want A4 page, portrait. The card will be centered and sized like a large badge (3.5 in x 5 in)
    doc = SimpleDocTemplate(
        filepath, 
        pagesize=A4, 
        leftMargin=1*inch, 
        rightMargin=1*inch, 
        topMargin=1.5*inch, 
        bottomMargin=1.5*inch
    )
    
    styles = getSampleStyleSheet()
    story = []
    
    # Custom styles
    title_style = ParagraphStyle(
        'CardTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=1, # Center
        spaceAfter=4
    )
    
    subtitle_style = ParagraphStyle(
        'CardSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        textColor=colors.HexColor('#E0E0FF'),
        alignment=1,
        spaceAfter=4
    )
    
    label_style = ParagraphStyle(
        'CardLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        textColor=colors.HexColor('#2E008B'),
        spaceAfter=2
    )
    
    val_style = ParagraphStyle(
        'CardVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        textColor=colors.HexColor('#333333'),
        spaceAfter=6
    )
    
    level_badge_style = ParagraphStyle(
        'CardLevel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=1
    )

    card_id_style = ParagraphStyle(
        'CardID',
        parent=styles['Normal'],
        fontName='Courier-Bold',
        fontSize=8,
        textColor=colors.HexColor('#666666'),
        alignment=2 # Right
    )

    # ── Let's design the card layout ──
    # Card Width: ~3.8 inches (270 points), Height: ~5.3 inches (380 points)
    card_width = 280
    
    # Header row flowables
    header_data = [
        [Paragraph("SEBN-TN MAINTENANCE", title_style)],
        [Paragraph("CARTE D'HABILITATION TECHNIQUE", subtitle_style)]
    ]
    header_table = Table(header_data, colWidths=[card_width])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#2E008B')),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ]))
    
    # User Photo Placeholder or actual photo
    photo_path = tech_details.get('photo')
    # Use placeholder if file doesn't exist
    if not photo_path or not os.path.exists(photo_path):
        photo_data = [
            [""],
            ["PHOTO"],
            [""]
        ]
        photo_table = Table(photo_data, colWidths=[90], rowHeights=[20, 40, 20])
        photo_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('INNERGRID', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#94A3B8')),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#64748B')),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
        ]))
    else:
        # Load and scale the image flowable
        photo_table = Image(photo_path, width=90, height=100)
        
    # User info side flowables
    info_data = [
        [Paragraph("NOM & PRÉNOM :", label_style)],
        [Paragraph(tech_details.get('name', 'N/A').upper(), val_style)],
        [Paragraph("MATRICULE :", label_style)],
        [Paragraph(tech_details.get('matricule', 'N/A'), val_style)],
        [Paragraph("ÉQUIPE (SHIFT) :", label_style)],
        [Paragraph(tech_details.get('shift', 'A'), val_style)]
    ]
    info_table = Table(info_data, colWidths=[150])
    info_table.setStyle(TableStyle([
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
    ]))
    
    # Combine photo + info side-by-side
    middle_data = [
        [photo_table, info_table]
    ]
    middle_table = Table(middle_data, colWidths=[100, 160])
    middle_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 14),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))

    # Level Banner flowable
    level_txt = tech_details.get('technician_level') or 'Level 0'
    # Extract level detail description
    lvl_desc = "Niveau Initial (Débutant)" if ("0" in level_txt) else "CSwin Basic Knowledge"
    if "2" in level_txt or "50%" in level_txt: lvl_desc = "CSwin Creation Hardware"
    elif "3" in level_txt or "75%" in level_txt: lvl_desc = "CSwin & Brainware & Vacuum"
    elif "4" in level_txt or "100%" in level_txt: lvl_desc = "CSwin & Brainware & Vacuum"

    level_data = [
        [Paragraph(f"{level_txt}", level_badge_style)],
        [Paragraph(f"{lvl_desc}", subtitle_style)]
    ]
    level_table = Table(level_data, colWidths=[card_width])
    level_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#2E008B')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ]))

    # Footer metrics flowable
    status_color = '#EF4444' if tech_details.get('status') == 'ÉCHOUÉ' else '#10B981'
    if tech_details.get('status') == 'Nouveau': status_color = '#3B82F6'

    status_badge_style = ParagraphStyle(
        'CardStatus',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        textColor=colors.HexColor(status_color)
    )

    footer_data = [
        [Paragraph("DERNIER TEST :", label_style), Paragraph("DÉCISION :", label_style)],
        [Paragraph(tech_details.get('exam_title', 'Aucun'), val_style), Paragraph(tech_details.get('status', 'Nouveau').upper(), status_badge_style)],
        [Paragraph("SCORE :", label_style), Paragraph("DATE DE DÉLIVRANCE :", label_style)],
        [Paragraph(f"{tech_details.get('score', 'N/A')} ({tech_details.get('percentage', 0.0)}%)", val_style), Paragraph(tech_details.get('date', '-'), val_style)],
        [Spacer(1, 10), Paragraph(tech_details.get('card_id', 'TECH-00000'), card_id_style)]
    ]
    footer_table = Table(footer_data, colWidths=[140, 120])
    footer_table.setStyle(TableStyle([
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('SPAN', (0,4), (0,4)), # Spacer span
    ]))

    # Final Combined Card Table representing the full cut-out badge
    card_elements = [
        [header_table],
        [middle_table],
        [level_table],
        [Spacer(1, 6)],
        [footer_table]
    ]
    
    # Outer box wrapping everything
    main_card_table = Table(card_elements, colWidths=[card_width])
    main_card_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 2, colors.HexColor('#2E008B')),
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    
    # Build Story
    story.append(Spacer(1, 15))
    story.append(Paragraph("SEBN-TN PORTAIL DE FORMATION", ParagraphStyle('PageTitle', parent=styles['Heading2'], alignment=1, textColor=colors.HexColor('#2E008B'))))
    story.append(Paragraph("HABILITATION & SUIVI DES NIVEAUX DE QUALIFICATION TECHNIQUE", ParagraphStyle('PageSub', parent=styles['Normal'], alignment=1, fontSize=8, textColor=colors.gray, spaceAfter=20)))
    story.append(Spacer(1, 10))
    story.append(main_card_table)
    story.append(Spacer(1, 30))
    story.append(Paragraph("Instructions d'impression :", ParagraphStyle('InstHead', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, spaceAfter=4)))
    story.append(Paragraph("1. Imprimez ce document en taille réelle à 100% sur du papier cartonné A4.", ParagraphStyle('Inst1', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#475569'))))
    story.append(Paragraph("2. Découpez le badge le long de la ligne de contour violette solide.", ParagraphStyle('Inst2', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#475569'))))
    story.append(Paragraph("3. Plastifiez le badge découpé pour assurer sa durabilité lors de son utilisation sur le terrain.", ParagraphStyle('Inst3', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#475569'))))
    
    doc.build(story)
    return filepath
