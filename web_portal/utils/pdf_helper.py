import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def export_dashboard_pdf(ima_db, pma_engine):
    export_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'exports')
    os.makedirs(export_dir, exist_ok=True)
    
    filename = f"Rapport_Maintenance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    filepath = os.path.join(export_dir, filename)
    
    doc = SimpleDocTemplate(filepath, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    story.append(Paragraph("Rapport de Maintenance SEBN-TN", styles['Title']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Date: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 24))
    
    # IMA Stats
    ima_stats = ima_db.get_analytics_bundle()
    story.append(Paragraph("1. Curatif (IMA)", styles['Heading2']))
    
    data = [
        ["Indicateur", "Valeur"],
        ["Total Interventions", str(ima_stats.get('total_interventions', 0))],
        ["Ouvertes", str(ima_stats.get('open_interventions', 0))],
        ["Clôturées", str(ima_stats.get('closed_interventions', 0))],
        ["Temps total d'arrêt (H)", str(ima_stats.get('total_downtime_hours', 0))]
    ]
    t = Table(data, colWidths=[300, 150])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-1), colors.beige),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    story.append(t)
    story.append(Spacer(1, 24))
    
    # PMA Stats
    pma_stats = pma_engine.get_stats()
    story.append(Paragraph("2. Préventif (PMA)", styles['Heading2']))
    
    data_pma = [
        ["Indicateur", "Valeur"],
        ["Total Tâches", str(pma_stats.get('total', 0))],
        ["Complétées", str(pma_stats.get('done', 0))],
        ["En Attente", str(pma_stats.get('pending', 0))],
        ["Taux de Complétion", f"{pma_stats.get('rate', 0)}%"]
    ]
    t2 = Table(data_pma, colWidths=[300, 150])
    t2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-1), colors.lightgreen),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    story.append(t2)
    
    doc.build(story)
    return filepath
