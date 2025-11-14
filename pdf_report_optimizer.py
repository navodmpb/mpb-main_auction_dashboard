# ============================================================
# OPTIMIZED PDF REPORT GENERATOR - ELEVATION-WISE ANALYSIS
# PHASE 1: PDF Optimization - All Grades Included (configurable)
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

class ElevationWiseReportOptimizer:
    """Optimized PDF report generator for elevation-wise analysis (includes all grades by default)"""
    
    def __init__(self, data, latest_df):
        self.data = data
        self.latest_df = latest_df
        self.styles = getSampleStyleSheet()
        self.setup_styles()
    
    def setup_styles(self):
        """Setup custom styles for professional appearance"""
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#1a5490'),
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            spaceAfter=12
        )
        
        self.heading1_style = ParagraphStyle(
            'CustomHeading1',
            parent=self.styles['Heading2'],
            fontSize=13,
            textColor=colors.HexColor('#2c5aa0'),
            spaceAfter=10,
            spaceBefore=10,
            fontName='Helvetica-Bold'
        )
        
        self.heading2_style = ParagraphStyle(
            'CustomHeading2',
            parent=self.styles['Heading3'],
            fontSize=11,
            textColor=colors.HexColor('#3d6bb3'),
            spaceAfter=8,
            spaceBefore=8,
            fontName='Helvetica-Bold'
        )
        
        self.body_style = ParagraphStyle(
            'CustomBody',
            parent=self.styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Elevation header style
        self.elevation_header_style = ParagraphStyle(
            'ElevationHeader',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#2c5aa0'),
            fontName='Helvetica-Bold',
            spaceAfter=4,
            spaceBefore=6
        )
    
    def get_color_for_percentage(self, percentage, high_is_good=True):
        """Get color based on performance percentage"""
        if high_is_good:
            if percentage >= 70:
                return colors.HexColor('#28a745')  # Green
            elif percentage >= 50:
                return colors.HexColor('#ffc107')  # Yellow
            else:
                return colors.HexColor('#dc3545')  # Red
        else:
            if percentage <= 30:
                return colors.HexColor('#28a745')  # Green
            elif percentage <= 50:
                return colors.HexColor('#ffc107')  # Yellow
            else:
                return colors.HexColor('#dc3545')  # Red
    
    def get_all_grades_per_elevation(self, df, elevation):
        """Return all grades for a specific elevation, sorted by total quantity desc"""
        elev_df = df[df["Sub Elevation"] == elevation]
        return elev_df.sort_values("Total Weight", ascending=False)
    
    def create_elevation_header_section(self, elevation, broker=None):
        """Create a formatted elevation header for tables"""
        if broker:
            header_text = f"SUB ELEVATION: {elevation}\nBROKER: {broker}"
        else:
            header_text = f"SUB ELEVATION: {elevation}"
        
        # Create separator line and header
        line = "═" * 50
        return f"{line}\n{header_text}\n{line}"
    
    def create_summary_table(self, data, elevation, metric_name="Sold"):
        """Create a summary table for elevation showing key metrics"""
        summary_data = []
        summary_data.append(['Grade', 'Catalogued (kg)', f'{metric_name} (kg)', f'{metric_name} %'])
        
        for grade in data['Grade'].unique():
            grade_data = data[data['Grade'] == grade]
            catalogued = grade_data['Total Weight'].sum()
            
            if metric_name == "Sold":
                metric_qty = grade_data[grade_data["Status_Clean"] == "sold"]["Total Weight"].sum()
            elif metric_name == "Unsold":
                metric_qty = grade_data[grade_data["Status_Clean"] == "unsold"]["Total Weight"].sum()
            elif metric_name == "Outsold":
                metric_qty = grade_data[grade_data["Status_Clean"] == "outsold"]["Total Weight"].sum()
            
            pct = (metric_qty / catalogued * 100) if catalogued > 0 else 0
            summary_data.append([
                grade[:20],
                f"{catalogued:,.0f}",
                f"{metric_qty:,.0f}",
                f"{pct:.1f}%"
            ])
        
        return summary_data
    
    def generate_broker_grade_sold_pct_optimized(self, story):
        """Report 1: Broker Grade-wise Sold % (All grades per elevation)"""
        story.append(Paragraph("REPORT 1: BROKER GRADE-WISE SOLD PERCENTAGES (BY SUB ELEVATION)", self.heading1_style))
        story.append(Spacer(1, 0.1*inch))
        
        broker_elev_grade = self.latest_df.groupby(["Broker", "Sub Elevation", "Grade"]).apply(lambda x: pd.Series({
            'Catalogued': x["Total Weight"].sum(),
            'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
            'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        }), include_groups=False).reset_index()
        
        broker_elev_grade['Total_Sold_Side'] = broker_elev_grade['Sold'] + broker_elev_grade['Outsold']
        broker_elev_grade['Sold_%'] = (broker_elev_grade['Total_Sold_Side'] / broker_elev_grade['Catalogued'] * 100).fillna(0)
        
        all_brokers = sorted(self.latest_df["Broker"].unique())
        all_elevations = sorted(self.latest_df["Sub Elevation"].unique())
        
        for broker_idx, broker in enumerate(all_brokers):
            broker_header_style = ParagraphStyle(
                'BrokerHeader',
                parent=self.heading2_style,
                fontSize=12,
                textColor=colors.HexColor('#1a5490'),
                fontName='Helvetica-Bold',
                spaceAfter=6,
                spaceBefore=10
            )
            story.append(Paragraph(f"BROKER: {broker}", broker_header_style))
            
            broker_data = broker_elev_grade[broker_elev_grade["Broker"] == broker]
            
            # Create comprehensive elevation analysis
            for elev_idx, elevation in enumerate(all_elevations):
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation]
                
                if not elev_data.empty:
                    # Elevation header
                    story.append(Paragraph(f"<b>SUB ELEVATION: {elevation}</b>", self.elevation_header_style))
                    
                    # Summary line
                    total_cat = elev_data['Catalogued'].sum()
                    total_sold_side = elev_data['Total_Sold_Side'].sum()
                    sold_pct = (total_sold_side / total_cat * 100) if total_cat > 0 else 0
                    
                    # Color code the summary
                    bg_color = self.get_color_for_percentage(sold_pct)
                    summary_text = f"Summary: Catalogued <b>{total_cat:,.0f}kg</b> | Sold <b>{total_sold_side:,.0f}kg</b> ({sold_pct:.1f}%)"
                    story.append(Paragraph(summary_text, self.body_style))
                    
                    # All grades table (sorted by catalogued quantity)
                    top_grades = elev_data.sort_values('Catalogued', ascending=False)
                    
                    table_data = [['Grade', 'Catalogued (kg)', 'Sold (kg)', 'Outsold (kg)', 'Sold %']]
                    
                    for _, row in top_grades.iterrows():
                        pct_color = self.get_color_for_percentage(row['Sold_%'])
                        table_data.append([
                            row['Grade'][:18],
                            f"{row['Catalogued']:,.0f}",
                            f"{row['Sold']:,.0f}",
                            f"{row['Outsold']:,.0f}",
                            f"{row['Sold_%']:.1f}%"
                        ])
                    
                    table = Table(table_data, colWidths=[1.5*inch, 1.1*inch, 1*inch, 1*inch, 0.8*inch])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c5aa0')),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,-1), 7),
                        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                        ('ALIGN', (0,0), (0,-1), 'LEFT'),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('PADDING', (0,0), (-1,-1), 3),
                        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                    ]))
                    
                    story.append(KeepTogether([table]))
                    story.append(Spacer(1, 0.05*inch))
            
            # Page break after each broker (except last)
            if broker_idx < len(all_brokers) - 1:
                story.append(PageBreak())
        
        return story
    
    def generate_broker_grade_unsold_pct_optimized(self, story):
        """Report 2: Broker Grade-wise Unsold % (All grades per elevation)"""
        story.append(PageBreak())
        story.append(Paragraph("REPORT 2: BROKER GRADE-WISE UNSOLD PERCENTAGES (BY SUB ELEVATION)", self.heading1_style))
        story.append(Spacer(1, 0.1*inch))
        
        broker_elev_grade = self.latest_df.groupby(["Broker", "Sub Elevation", "Grade"]).apply(lambda x: pd.Series({
            'Catalogued': x["Total Weight"].sum(),
            'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        }), include_groups=False).reset_index()
        
        broker_elev_grade['Unsold_%'] = (broker_elev_grade['Unsold'] / broker_elev_grade['Catalogued'] * 100).fillna(0)
        
        all_brokers = sorted(self.latest_df["Broker"].unique())
        all_elevations = sorted(self.latest_df["Sub Elevation"].unique())
        
        for broker_idx, broker in enumerate(all_brokers):
            broker_header_style = ParagraphStyle(
                'BrokerHeader',
                parent=self.heading2_style,
                fontSize=12,
                textColor=colors.HexColor('#1a5490'),
                fontName='Helvetica-Bold'
            )
            story.append(Paragraph(f"BROKER: {broker}", broker_header_style))
            
            broker_data = broker_elev_grade[broker_elev_grade["Broker"] == broker]
            
            for elevation in all_elevations:
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation]
                
                if not elev_data.empty:
                    story.append(Paragraph(f"<b>SUB ELEVATION: {elevation}</b>", self.elevation_header_style))
                    
                    total_cat = elev_data['Catalogued'].sum()
                    total_unsold = elev_data['Unsold'].sum()
                    unsold_pct = (total_unsold / total_cat * 100) if total_cat > 0 else 0
                    
                    story.append(Paragraph(f"Summary: Catalogued <b>{total_cat:,.0f}kg</b> | Unsold <b>{total_unsold:,.0f}kg</b> ({unsold_pct:.1f}%)", self.body_style))
                    
                    # All grades (sorted by catalogued quantity)
                    top_grades = elev_data.sort_values('Catalogued', ascending=False)
                    
                    table_data = [['Grade', 'Catalogued (kg)', 'Unsold (kg)', 'Unsold %']]
                    
                    for _, row in top_grades.iterrows():
                        table_data.append([
                            row['Grade'][:18],
                            f"{row['Catalogued']:,.0f}",
                            f"{row['Unsold']:,.0f}",
                            f"{row['Unsold_%']:.1f}%"
                        ])
                    
                    table = Table(table_data, colWidths=[1.8*inch, 1.2*inch, 1.2*inch, 1*inch])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#dc3545')),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,-1), 7),
                        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                        ('ALIGN', (0,0), (0,-1), 'LEFT'),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('PADDING', (0,0), (-1,-1), 3),
                    ]))
                    
                    story.append(table)
                    story.append(Spacer(1, 0.05*inch))
            
            if broker_idx < len(all_brokers) - 1:
                story.append(PageBreak())
        
        return story
    
    def generate_report(self, include_reports=None):
        """Generate optimized multi-report PDF"""
        if include_reports is None:
            include_reports = {
                'report1': True,
                'report2': True,
                'report3': True,
                'report4': True,
                'report5': True,
            }
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=1.5*cm,
            leftMargin=1.5*cm,
            topMargin=2.5*cm,
            bottomMargin=2.5*cm
        )
        
        story = []
        
        # Title Page
        story.append(Spacer(1, 0.5*inch))
        story.append(Paragraph("Mercantile Produce Brokers Pvt Ltd", self.heading2_style))
        story.append(Paragraph("MAIN AUCTION DETAILED REPORT", self.title_style))
        story.append(Spacer(1, 0.3*inch))
        story.append(Paragraph("Elevation-Wise Performance Analysis (All Grades Included)", self.body_style))
        story.append(Spacer(1, 0.2*inch))
        story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M')}", self.body_style))
        story.append(PageBreak())
        
        # Generate selected reports
        if include_reports.get('report1'):
            self.generate_broker_grade_sold_pct_optimized(story)
            story.append(PageBreak())
        
        if include_reports.get('report2'):
            self.generate_broker_grade_unsold_pct_optimized(story)
            story.append(PageBreak())
        
        # Build PDF
        doc.build(story)
        
        pdf_bytes = buffer.getvalue()
        buffer.close()
        
        return pdf_bytes


# Helper function for use in Streamlit
def generate_optimized_elevation_report(data, latest_df, include_reports=None):
    """Generate optimized elevation-wise PDF report"""
    optimizer = ElevationWiseReportOptimizer(data, latest_df)
    return optimizer.generate_report(include_reports=include_reports)
