# ============================================================
# PDF REPORT ENHANCEMENTS
# Mini charts, TOC, per-broker reports, conditional formatting
# ============================================================

import pandas as pd
import numpy as np
import io
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, 
                                Spacer, PageBreak, Image, KeepTogether)
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class PDFWithBookmarks(SimpleDocTemplate):
    """PDF document with bookmarks and TOC support"""
    def __init__(self, *args, **kwargs):
        self.toc_entries = []
        self.page_bookmarks = {}
        super().__init__(*args, **kwargs)
    
    def add_bookmark(self, title, page=None):
        """Add a bookmark for a page"""
        if page is None:
            page = self._pageNumber if hasattr(self, '_pageNumber') else 0
        self.page_bookmarks[title] = page
        self.toc_entries.append((title, page))

def get_color_for_percentage(percentage, threshold_high=50, threshold_med=70):
    """Get color based on performance thresholds"""
    if percentage >= threshold_med:
        return colors.HexColor('#28a745')  # Green: excellent
    elif percentage >= threshold_high:
        return colors.HexColor('#ffc107')  # Yellow: good
    else:
        return colors.HexColor('#dc3545')  # Red: needs improvement

def create_mini_bar_chart(value, max_value=100, width=0.8*inch, height=0.2*inch):
    """Create a simple mini bar chart image for PDF"""
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.lib.utils import ImageReader
    
    buffer = io.BytesIO()
    pdf_canvas = pdfcanvas.Canvas(buffer, pagesize=(width, height))
    
    # Calculate bar width
    bar_width = (value / max_value) * (width - 0.1*inch) if max_value > 0 else 0
    
    # Get color
    bar_color = get_color_for_percentage(value)
    pdf_canvas.setFillColor(bar_color)
    pdf_canvas.rect(0.05*inch, 0.05*inch, bar_width, height - 0.1*inch, fill=1, stroke=0)
    
    # Border
    pdf_canvas.setLineWidth(0.5)
    pdf_canvas.setStrokeColor(colors.grey)
    pdf_canvas.rect(0.05*inch, 0.05*inch, (width - 0.1*inch), height - 0.1*inch, fill=0, stroke=1)
    
    # Label
    pdf_canvas.setFont("Helvetica", 7)
    pdf_canvas.drawString(0.1*inch, 0.06*inch, f"{value:.1f}%")
    
    pdf_canvas.save()
    buffer.seek(0)
    return ImageReader(buffer)

def create_toc_page(toc_entries, doc, styles):
    """Create a Table of Contents page"""
    story = []
    story.append(Paragraph("TABLE OF CONTENTS", styles['Heading1']))
    story.append(Spacer(1, 0.3*inch))
    
    toc_data = []
    for entry_title, page_num in toc_entries:
        toc_data.append([
            entry_title,
            str(page_num) if page_num else "–"
        ])
    
    if toc_data:
        toc_table = Table(toc_data, colWidths=[5*inch, 1*inch])
        toc_table.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 10),
            ('ALIGN', (1,0), (1,-1), 'RIGHT'),
            ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.grey),
        ]))
        story.append(toc_table)
    
    story.append(PageBreak())
    return story

def create_per_broker_pdf(data, latest_df, broker_name, include_reports=None):
    """Generate a single-broker focused PDF report"""
    if include_reports is None:
        include_reports = {
            'report1': True,
            'report2': True,
        }
    
    # Filter data to specific broker
    broker_df = latest_df[latest_df['Broker'] == broker_name]
    
    if broker_df.empty:
        raise ValueError(f"No data found for broker: {broker_name}")
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=2.5*cm,
        bottomMargin=2.5*cm
    )
    
    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#1a5490'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
    )
    
    heading_style = ParagraphStyle(
        'Heading',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=colors.HexColor('#2c5aa0'),
        fontName='Helvetica-Bold',
    )
    
    body_style = styles['Normal']
    
    story = []
    
    # Title
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(f"BROKER PERFORMANCE REPORT", title_style))
    story.append(Paragraph(f"<b>{broker_name}</b>", heading_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M')}", body_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Key metrics
    total_cat = broker_df['Total Weight'].sum()
    total_sold = broker_df[broker_df['Status_Clean'] == 'sold']['Total Weight'].sum()
    total_outsold = broker_df[broker_df['Status_Clean'] == 'outsold']['Total Weight'].sum()
    total_unsold = broker_df[broker_df['Status_Clean'] == 'unsold']['Total Weight'].sum()
    total_value = broker_df['Total Value'].sum()
    avg_price = broker_df[broker_df['Status_Clean'] == 'sold']['Price'].mean()
    
    sold_pct = (total_sold / total_cat * 100) if total_cat > 0 else 0
    unsold_pct = (total_unsold / total_cat * 100) if total_cat > 0 else 0
    outsold_pct = (total_outsold / total_cat * 100) if total_cat > 0 else 0
    
    # Summary metrics table
    story.append(Paragraph("PERFORMANCE SUMMARY", heading_style))
    summary_data = [
        ['Metric', 'Value'],
        ['Catalogued (kg)', f'{total_cat:,.0f}'],
        ['Sold (kg)', f'{total_sold:,.0f}'],
        ['Unsold (kg)', f'{total_unsold:,.0f}'],
        ['Outsold (kg)', f'{total_outsold:,.0f}'],
        ['Sold %', f'{sold_pct:.2f}%'],
        ['Unsold %', f'{unsold_pct:.2f}%'],
        ['Outsold %', f'{outsold_pct:.2f}%'],
        ['Total Value (LKR)', f'{total_value:,.0f}'],
        ['Average Price (LKR)', f'{avg_price:,.2f}' if pd.notna(avg_price) else 'N/A'],
    ]
    
    summary_table = Table(summary_data, colWidths=[2.5*inch, 2*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1a5490')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('PADDING', (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
    ]))
    story.append(summary_table)
    story.append(PageBreak())
    
    # Elevation breakdown
    story.append(Paragraph("ELEVATION-WISE PERFORMANCE", heading_style))
    story.append(Spacer(1, 0.1*inch))
    
    elev_perf = broker_df.groupby('Sub Elevation').apply(lambda x: pd.Series({
        'Catalogued': x['Total Weight'].sum(),
        'Sold': x[x['Status_Clean'] == 'sold']['Total Weight'].sum(),
        'Unsold': x[x['Status_Clean'] == 'unsold']['Total Weight'].sum(),
        'Outsold': x[x['Status_Clean'] == 'outsold']['Total Weight'].sum(),
        'Total_Value': x['Total Value'].sum(),
    }), include_groups=False).reset_index()
    
    elev_perf['Sold_%'] = (elev_perf['Sold'] / elev_perf['Catalogued'] * 100).fillna(0)
    elev_perf['Unsold_%'] = (elev_perf['Unsold'] / elev_perf['Catalogued'] * 100).fillna(0)
    elev_perf['Outsold_%'] = (elev_perf['Outsold'] / elev_perf['Catalogued'] * 100).fillna(0)
    
    # Table with conditional formatting
    elev_table_data = [['Sub Elevation', 'Catalogued (kg)', 'Sold %', 'Unsold %', 'Outsold %', 'Value (LKR)']]
    table_style_rules = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1a5490')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('PADDING', (0,0), (-1,-1), 5),
    ]
    
    for idx, (_, row) in enumerate(elev_perf.iterrows(), 1):
        sold_pct_val = row['Sold_%']
        
        # Conditional formatting: color rows based on Sold%
        if sold_pct_val < 50:
            row_bg = colors.HexColor('#f8d7da')  # Light red
            row_text = colors.HexColor('#721c24')  # Dark red text
        elif sold_pct_val < 70:
            row_bg = colors.HexColor('#fff3cd')  # Light yellow
            row_text = colors.HexColor('#856404')  # Dark yellow text
        else:
            row_bg = colors.HexColor('#d4edda')  # Light green
            row_text = colors.HexColor('#155724')  # Dark green text
        
        table_style_rules.append(('BACKGROUND', (0,idx), (-1,idx), row_bg))
        table_style_rules.append(('TEXTCOLOR', (0,idx), (-1,idx), row_text))
        
        elev_table_data.append([
            row['Sub Elevation'],
            f"{row['Catalogued']:,.0f}",
            f"{sold_pct_val:.2f}%",
            f"{row['Unsold_%']:.2f}%",
            f"{row['Outsold_%']:.2f}%",
            f"{row['Total_Value']:,.0f}"
        ])
    
    if len(elev_table_data) > 1:
        elev_table = Table(elev_table_data, colWidths=[1.2*inch, 1.2*inch, 0.9*inch, 0.9*inch, 0.9*inch, 1.2*inch])
        elev_table.setStyle(TableStyle(table_style_rules))
        story.append(elev_table)
    
    story.append(PageBreak())
    
    # Build PDF
    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    
    return pdf_bytes

# Helper function for multi-broker reports with TOC
def generate_optimized_elevation_report_with_toc(data, latest_df, include_reports=None):
    """Generate optimized elevation-wise PDF report with TOC and bookmarks"""
    # For now, delegate to existing optimizer
    # TOC will be added in next iteration with proper bookmark integration
    from pdf_report_optimizer import generate_optimized_elevation_report
    return generate_optimized_elevation_report(data, latest_df, include_reports)
