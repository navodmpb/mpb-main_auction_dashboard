# ============================================================
# IMPORTS
# ============================================================

import warnings
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.filterwarnings('ignore', message='.*Kaleido.*')

# Core libraries
import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import io
import base64
from datetime import datetime

# Plotly for visualizations
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio  # For chart exports to PDF

# ReportLab for PDF generation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, 
                                Spacer, PageBreak, Image, KeepTogether)
from reportlab.pdfgen import canvas
from reportlab.pdfgen.canvas import Canvas


from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import math

# ============================================================
# STREAMLIT PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Tea Auction Intelligence Dashboard", 
    page_icon="", 
    layout="wide"
)

class NumberedCanvas(canvas.Canvas):
    """Custom canvas with page numbers and company branding"""
    
    def __init__(self, *args, **kwargs):
        self.logo_path = kwargs.pop('logo_path', None)
        self.footer_text = kwargs.pop('footer_text', 'MPBL IT')
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []
        
    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()
        
    def save(self):
        page_count = len(self.pages)
        for page_num, page in enumerate(self.pages, 1):
            self.__dict__.update(page)
            self.draw_page_elements(page_num, page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)
        
    def draw_page_elements(self, page_num, page_count):
        """Draw header, footer, page numbers"""
        # Footer
        self.saveState()
        self.setFont('Helvetica', 8)
        self.setFillColorRGB(0.5, 0.5, 0.5)
        
        # Left footer - Company name
        self.drawString(1.5*cm, 1*cm, self.footer_text)
        
        # Center footer - Generated date
        date_str = f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}"
        self.drawCentredString(A4[0]/2, 1*cm, date_str)
        
        # Right footer - Page numbers
        page_str = f"Page {page_num} of {page_count}"
        self.drawRightString(A4[0] - 1.5*cm, 1*cm, page_str)
        
        self.restoreState()

# ============================================================
# HELPER FUNCTIONS FOR PDF CHARTS
# ============================================================

def plotly_fig_to_image(fig, width=800, height=500):
    """
    Convert Plotly figure to PIL Image for PDF embedding
    
    Args:
        fig: Plotly figure object
        width: Image width in pixels
        height: Image height in pixels
    
    Returns:
        PIL Image object
    """
    try:
        import PIL.Image as PILImage
        from io import BytesIO
        import warnings
        
        # Suppress Kaleido deprecation warnings
        warnings.filterwarnings('ignore', category=DeprecationWarning)
        warnings.filterwarnings('ignore', message='.*Kaleido.*')
        
        # Update figure layout for better PDF display
        fig.update_layout(
            width=width,
            height=height,
            margin=dict(l=50, r=50, t=50, b=50),
            font=dict(size=10)
        )
        
        # Convert to PNG bytes (suppress warnings)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            img_bytes = fig.to_image(format="png", width=width, height=height)
        
        # Convert to PIL Image
        img = PILImage.open(BytesIO(img_bytes))
        return img
    except ImportError:
        # Try alternative method with PIL
        try:
            from PIL import Image as PILImage
            from io import BytesIO
            import warnings
            warnings.filterwarnings('ignore', category=DeprecationWarning)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                img_bytes = fig.to_image(format="png", width=width, height=height, scale=2)
            img = PILImage.open(BytesIO(img_bytes))
            return img
        except:
            return None
    except Exception as e:
        # Silent fail for PDF generation - charts are optional
        return None

def create_market_share_chart(latest_df):
    """Create market share pie chart"""
    market_share = latest_df.groupby("Broker")["Total Value"].sum().sort_values(ascending=False).reset_index()
    fig = px.pie(market_share, names="Broker", values="Total Value", 
                 title="Market Share by Broker (Value)",
                 color_discrete_sequence=px.colors.qualitative.Pastel)
    fig.update_traces(textinfo="percent+label")
    return fig

def create_status_distribution_chart(latest_df):
    """Create status distribution pie chart"""
    sold_mask = latest_df["Status_Clean"] == "sold"
    unsold_mask = latest_df["Status_Clean"] == "unsold"
    outsold_mask = latest_df["Status_Clean"] == "outsold"
    
    sold_weight = latest_df.loc[sold_mask, "Total Weight"].sum()
    unsold_weight = latest_df.loc[unsold_mask, "Total Weight"].sum()
    outsold_weight = latest_df.loc[outsold_mask, "Total Weight"].sum()
    
    status_dist = pd.DataFrame({
        'Status': ['Sold', 'Unsold', 'Outsold'],
        'Weight': [sold_weight, unsold_weight, outsold_weight]
    })
    status_dist = status_dist[status_dist['Weight'] > 0]
    
    fig = px.pie(status_dist, values='Weight', names='Status',
                 title='Overall Sale Status Distribution',
                 color='Status',
                 color_discrete_map={'Sold': '#28a745', 'Unsold': '#dc3545', 'Outsold': '#ffc107'})
    fig.update_traces(textposition='inside', textinfo='percent+label')
    return fig

def create_broker_performance_chart(broker_performance):
    """Create broker performance bar chart"""
    fig = px.bar(broker_performance, x='Broker', y='Sold_Percentage',
                 title='Broker Performance - Sold % (Sold+Outsold)',
                 color='Sold_Percentage',
                 color_continuous_scale='Greens',
                 text='Sold_Percentage')
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig.update_layout(
        xaxis_tickangle=-45,
        yaxis_title='Sold %',
        showlegend=False
    )
    return fig

def create_elevation_performance_chart(elev_summary_data):
    """Create elevation performance stacked bar chart"""
    fig = px.bar(elev_summary_data, x='Sub Elevation', 
                 y=['Sold_Percentage', 'Unsold_Percentage', 'Outsold_Percentage'],
                 title='Elevation Performance - Status Percentages',
                 labels={'value': 'Percentage', 'variable': 'Status'},
                 barmode='stack',
                 color_discrete_map={'Sold_Percentage': '#28a745', 
                                    'Unsold_Percentage': '#dc3545', 
                                    'Outsold_Percentage': '#ffc107'})
    fig.update_layout(xaxis_tickangle=-45)
    return fig

# ============================================================
# OPTIMIZED PDF GENERATOR - ELEVATION-WISE REPORTS (NO CHARTS)
# ============================================================

def generate_broker_grade_sold_pct(latest_df, story, heading1_style, heading2_style, body_style):
    """Report 1: Each Broker's grade wise sold percentages (Sub Elevation wise) - ALL GRADES"""
    story.append(Paragraph("REPORT 1: BROKER GRADE-WISE SOLD PERCENTAGES (BY SUB ELEVATION)", heading1_style))
    story.append(Spacer(1, 0.1*inch))
    
    # Calculate data: Broker -> Sub Elevation -> Grade -> Sold %
    broker_elev_grade = latest_df.groupby(["Broker", "Sub Elevation", "Grade"]).apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
    }), include_groups=False).reset_index()
    
    broker_elev_grade['Total_Sold_Side'] = broker_elev_grade['Sold'] + broker_elev_grade['Outsold']
    broker_elev_grade['Sold_%'] = (broker_elev_grade['Total_Sold_Side'] / broker_elev_grade['Catalogued'] * 100).fillna(0)
    
    all_brokers = sorted(latest_df["Broker"].unique())
    
    for broker_idx, broker in enumerate(all_brokers):
        broker_header_style = ParagraphStyle(
            'BrokerHeader',
            parent=heading2_style,
            fontSize=12,
            textColor=colors.HexColor('#1a5490'),
            fontName='Helvetica-Bold',
            spaceAfter=6,
            spaceBefore=10
        )
        story.append(Paragraph(f"BROKER: {broker}", broker_header_style))
        
        broker_data = broker_elev_grade[broker_elev_grade["Broker"] == broker]
        
        if not broker_data.empty:
            elevations = sorted(broker_data["Sub Elevation"].unique())
            
            # Summary table for all elevations
            summary_data = []
            for elevation in elevations:
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation]
                total_cat = elev_data['Catalogued'].sum()
                total_sold_side = elev_data['Total_Sold_Side'].sum()
                sold_pct = (total_sold_side / total_cat * 100) if total_cat > 0 else 0
                summary_data.append({
                    'Elevation': elevation,
                    'Catalogued': total_cat,
                    'Sold_%': sold_pct
                })
            
            # Add summary table
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_table_data = [['Sub Elevation', 'Catalogued (kg)', 'Sold %']]
                for _, row in summary_df.iterrows():
                    summary_table_data.append([
                        row['Elevation'],
                        f"{row['Catalogued']:,.0f}",
                        f"{row['Sold_%']:.2f}%"
                    ])
                
                summary_table = Table(summary_table_data, colWidths=[1.5*inch, 1.5*inch, 1*inch])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c5aa0')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                    ('ALIGN', (0,0), (0,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('PADDING', (0,0), (-1,-1), 5),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 0.15*inch))
            
            # Show ALL grades per elevation (not top 10)
            for elev_idx, elevation in enumerate(elevations):
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation].sort_values('Catalogued', ascending=False)
                
                if not elev_data.empty:
                    elev_header_style = ParagraphStyle(
                        'ElevHeader',
                        parent=body_style,
                        fontSize=9,
                        textColor=colors.HexColor('#2c5aa0'),
                        fontName='Helvetica-Bold',
                        spaceAfter=4,
                        spaceBefore=6
                    )
                    story.append(Paragraph(f"Sub Elevation: {elevation}", elev_header_style))
                    
                    table_data = [['Grade', 'Catalogued (kg)', 'Sold (kg)', 'Outsold (kg)', 'Sold %']]
                    
                    for _, row in elev_data.iterrows():
                        table_data.append([
                            row['Grade'][:18],
                            f"{row['Catalogued']:,.0f}",
                            f"{row['Sold']:,.0f}",
                            f"{row['Outsold']:,.0f}",
                            f"{row['Sold_%']:.2f}%"
                        ])
                    
                    table = Table(table_data, colWidths=[1.5*inch, 1*inch, 0.9*inch, 0.9*inch, 0.9*inch])
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
                    story.append(Spacer(1, 0.1*inch))
        
        # Page break after each broker
        story.append(PageBreak())

def generate_broker_grade_unsold_pct(latest_df, story, heading1_style, heading2_style, body_style):
    """Report 2: Each Broker's grade wise unsold percentages (Sub Elevation wise) - ALL GRADES"""
    story.append(Paragraph("REPORT 2: BROKER GRADE-WISE UNSOLD PERCENTAGES (BY SUB ELEVATION)", heading1_style))
    story.append(Spacer(1, 0.1*inch))
    
    broker_elev_grade = latest_df.groupby(["Broker", "Sub Elevation", "Grade"]).apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
    }), include_groups=False).reset_index()
    
    broker_elev_grade['Unsold_%'] = (broker_elev_grade['Unsold'] / broker_elev_grade['Catalogued'] * 100).fillna(0)
    
    all_brokers = sorted(latest_df["Broker"].unique())
    
    for broker_idx, broker in enumerate(all_brokers):
        broker_header_style = ParagraphStyle(
            'BrokerHeader',
            parent=heading2_style,
            fontSize=12,
            textColor=colors.HexColor('#1a5490'),
            fontName='Helvetica-Bold',
            spaceAfter=6,
            spaceBefore=10
        )
        story.append(Paragraph(f"BROKER: {broker}", broker_header_style))
        
        broker_data = broker_elev_grade[broker_elev_grade["Broker"] == broker]
        
        if not broker_data.empty:
            elevations = sorted(broker_data["Sub Elevation"].unique())
            
            # Summary table
            summary_data = []
            for elevation in elevations:
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation]
                total_cat = elev_data['Catalogued'].sum()
                total_unsold = elev_data['Unsold'].sum()
                unsold_pct = (total_unsold / total_cat * 100) if total_cat > 0 else 0
                summary_data.append({
                    'Elevation': elevation,
                    'Catalogued': total_cat,
                    'Unsold_%': unsold_pct
                })
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_table_data = [['Sub Elevation', 'Catalogued (kg)', 'Unsold %']]
                for _, row in summary_df.iterrows():
                    summary_table_data.append([
                        row['Elevation'],
                        f"{row['Catalogued']:,.0f}",
                        f"{row['Unsold_%']:.2f}%"
                    ])
                
                summary_table = Table(summary_table_data, colWidths=[1.5*inch, 1.5*inch, 1*inch])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#dc3545')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                    ('ALIGN', (0,0), (0,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('PADDING', (0,0), (-1,-1), 5),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 0.15*inch))
            
            # Show ALL grades per elevation
            for elev_idx, elevation in enumerate(elevations):
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation].sort_values('Catalogued', ascending=False)
                
                if not elev_data.empty:
                    elev_header_style = ParagraphStyle(
                        'ElevHeader',
                        parent=body_style,
                        fontSize=9,
                        textColor=colors.HexColor('#2c5aa0'),
                        fontName='Helvetica-Bold',
                        spaceAfter=4,
                        spaceBefore=6
                    )
                    story.append(Paragraph(f"Sub Elevation: {elevation}", elev_header_style))
                    
                    table_data = [['Grade', 'Catalogued (kg)', 'Unsold (kg)', 'Unsold %']]
                    for _, row in elev_data.iterrows():
                        table_data.append([
                            row['Grade'][:18],
                            f"{row['Catalogued']:,.0f}",
                            f"{row['Unsold']:,.0f}",
                            f"{row['Unsold_%']:.2f}%"
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
                        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                    ]))
                    story.append(KeepTogether([table]))
                    story.append(Spacer(1, 0.1*inch))
        
        story.append(PageBreak())

def generate_broker_grade_outsold_pct(latest_df, story, heading1_style, heading2_style, body_style):
    """Report 3: Each Broker's grade wise outsold percentages (Sub Elevation wise) - ALL GRADES"""
    story.append(Paragraph("REPORT 3: BROKER GRADE-WISE OUTSOLD PERCENTAGES (BY SUB ELEVATION)", heading1_style))
    story.append(Spacer(1, 0.1*inch))
    
    broker_elev_grade = latest_df.groupby(["Broker", "Sub Elevation", "Grade"]).apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
    }), include_groups=False).reset_index()
    
    broker_elev_grade['Outsold_%'] = (broker_elev_grade['Outsold'] / broker_elev_grade['Catalogued'] * 100).fillna(0)
    
    all_brokers = sorted(latest_df["Broker"].unique())
    
    for broker_idx, broker in enumerate(all_brokers):
        broker_header_style = ParagraphStyle(
            'BrokerHeader',
            parent=heading2_style,
            fontSize=12,
            textColor=colors.HexColor('#1a5490'),
            fontName='Helvetica-Bold',
            spaceAfter=6,
            spaceBefore=10
        )
        story.append(Paragraph(f"BROKER: {broker}", broker_header_style))
        
        broker_data = broker_elev_grade[broker_elev_grade["Broker"] == broker]
        
        if not broker_data.empty:
            elevations = sorted(broker_data["Sub Elevation"].unique())
            
            # Summary table
            summary_data = []
            for elevation in elevations:
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation]
                total_cat = elev_data['Catalogued'].sum()
                total_outsold = elev_data['Outsold'].sum()
                outsold_pct = (total_outsold / total_cat * 100) if total_cat > 0 else 0
                summary_data.append({
                    'Elevation': elevation,
                    'Catalogued': total_cat,
                    'Outsold_%': outsold_pct
                })
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_table_data = [['Sub Elevation', 'Catalogued (kg)', 'Outsold %']]
                for _, row in summary_df.iterrows():
                    summary_table_data.append([
                        row['Elevation'],
                        f"{row['Catalogued']:,.0f}",
                        f"{row['Outsold_%']:.2f}%"
                    ])
                
                summary_table = Table(summary_table_data, colWidths=[1.5*inch, 1.5*inch, 1*inch])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ffc107')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#000000')),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                    ('ALIGN', (0,0), (0,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('PADDING', (0,0), (-1,-1), 5),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 0.15*inch))
            
            # Show ALL grades per elevation
            for elev_idx, elevation in enumerate(elevations):
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation].sort_values('Catalogued', ascending=False)
                
                if not elev_data.empty:
                    elev_header_style = ParagraphStyle(
                        'ElevHeader',
                        parent=body_style,
                        fontSize=9,
                        textColor=colors.HexColor('#2c5aa0'),
                        fontName='Helvetica-Bold',
                        spaceAfter=4,
                        spaceBefore=6
                    )
                    story.append(Paragraph(f"Sub Elevation: {elevation}", elev_header_style))
                    
                    table_data = [['Grade', 'Catalogued (kg)', 'Outsold (kg)', 'Outsold %']]
                    for _, row in elev_data.iterrows():
                        table_data.append([
                            row['Grade'][:18],
                            f"{row['Catalogued']:,.0f}",
                            f"{row['Outsold']:,.0f}",
                            f"{row['Outsold_%']:.2f}%"
                        ])
                    
                    table = Table(table_data, colWidths=[1.8*inch, 1.2*inch, 1.2*inch, 1*inch])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ffc107')),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#000000')),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,-1), 7),
                        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                        ('ALIGN', (0,0), (0,-1), 'LEFT'),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('PADDING', (0,0), (-1,-1), 3),
                        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                    ]))
                    story.append(KeepTogether([table]))
                    story.append(Spacer(1, 0.1*inch))
        
        story.append(PageBreak())

def generate_broker_grade_sold_qty_price(latest_df, story, heading1_style, heading2_style, body_style):
    """Report 4: Each Broker's grade wise sold quantities and Avg. Prices (Sub Elevation wise) - SUMMARIZED"""
    story.append(Paragraph("REPORT 4: BROKER GRADE-WISE SOLD QUANTITIES & AVG PRICES (BY SUB ELEVATION)", heading1_style))
    story.append(Spacer(1, 0.1*inch))
    
    broker_elev_grade = latest_df.groupby(["Broker", "Sub Elevation", "Grade"]).apply(lambda x: pd.Series({
        'Sold_Qty': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Outsold_Qty': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean(),
    }), include_groups=False).reset_index()
    
    broker_elev_grade['Total_Sold_Side'] = broker_elev_grade['Sold_Qty'] + broker_elev_grade['Outsold_Qty']
    
    all_brokers = sorted(latest_df["Broker"].unique())
    
    for broker_idx, broker in enumerate(all_brokers):
        broker_header_style = ParagraphStyle(
            'BrokerHeader',
            parent=heading2_style,
            fontSize=12,
            textColor=colors.HexColor('#1a5490'),
            fontName='Helvetica-Bold',
            spaceAfter=6,
            spaceBefore=10
        )
        story.append(Paragraph(f"BROKER: {broker}", broker_header_style))
        
        broker_data = broker_elev_grade[broker_elev_grade["Broker"] == broker]
        
        if not broker_data.empty:
            elevations = sorted(broker_data["Sub Elevation"].unique())
            
            # Summary table
            summary_data = []
            for elevation in elevations:
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation]
                total_sold = elev_data['Total_Sold_Side'].sum()
                avg_price = elev_data[elev_data['Avg_Price'].notna()]['Avg_Price'].mean()
                summary_data.append({
                    'Elevation': elevation,
                    'Total_Sold': total_sold,
                    'Avg_Price': avg_price if pd.notna(avg_price) else 0
                })
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_table_data = [['Sub Elevation', 'Total Sold+Outsold (kg)', 'Avg Price (LKR)']]
                for _, row in summary_df.iterrows():
                    summary_table_data.append([
                        row['Elevation'],
                        f"{row['Total_Sold']:,.0f}",
                        f"{row['Avg_Price']:,.2f}" if row['Avg_Price'] > 0 else 'N/A'
                    ])
                
                summary_table = Table(summary_table_data, colWidths=[1.5*inch, 1.8*inch, 1.2*inch])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#28a745')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                    ('ALIGN', (0,0), (0,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('PADDING', (0,0), (-1,-1), 5),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 0.15*inch))
            
            # Show ALL grades per elevation
            for elev_idx, elevation in enumerate(elevations):
                elev_data = broker_data[broker_data["Sub Elevation"] == elevation].sort_values('Total_Sold_Side', ascending=False)
                elev_data = elev_data[elev_data['Total_Sold_Side'] > 0]
                
                if not elev_data.empty:
                    elev_header_style = ParagraphStyle(
                        'ElevHeader',
                        parent=body_style,
                        fontSize=9,
                        textColor=colors.HexColor('#2c5aa0'),
                        fontName='Helvetica-Bold',
                        spaceAfter=4,
                        spaceBefore=6
                    )
                    story.append(Paragraph(f"Sub Elevation: {elevation}", elev_header_style))
                    
                    table_data = [['Grade', 'Sold (kg)', 'Outsold (kg)', 'Total Sold+Outsold (kg)', 'Avg Price (LKR)']]
                    for _, row in elev_data.iterrows():
                        table_data.append([
                            row['Grade'][:18],
                            f"{row['Sold_Qty']:,.0f}",
                            f"{row['Outsold_Qty']:,.0f}",
                            f"{row['Total_Sold_Side']:,.0f}",
                            f"{row['Avg_Price']:,.2f}" if pd.notna(row['Avg_Price']) else 'N/A'
                        ])
                    
                    table = Table(table_data, colWidths=[1.2*inch, 0.9*inch, 0.9*inch, 1.1*inch, 1*inch])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#28a745')),
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
                    story.append(Spacer(1, 0.1*inch))
        
        story.append(PageBreak())

def generate_buyer_grade_profiles(latest_df, story, heading1_style, heading2_style, body_style):
    """Report 5: Outlots purchased buyer profiles (Grade wise, Sub Elevation wise) - ALL BUYERS & GRADES"""
    story.append(Paragraph("REPORT 5: OUTLOTS PURCHASED BUYER PROFILES (GRADE WISE BY SUB ELEVATION)", heading1_style))
    story.append(Spacer(1, 0.1*inch))
    
    # Get sold data (outlots = sold lots)
    sold_df = latest_df[latest_df["Status_Clean"] == "sold"]
    
    if sold_df.empty:
        story.append(Paragraph("No sold lots available for buyer analysis.", body_style))
        return
    
    # Calculate buyer-elevation-grade data
    buyer_elev_grade = sold_df.groupby(["Buyer", "Sub Elevation", "Grade"]).agg({
        "Total Weight": "sum",
        "Price": "mean",
        "Total Value": "sum",
        "Lot No": "count"
    }).reset_index()
    
    # Get ALL buyers (sorted by total value)
    all_buyers = sold_df.groupby("Buyer")["Total Value"].sum().sort_values(ascending=False).index.tolist()
    
    for buyer_idx, buyer in enumerate(all_buyers):
        buyer_header_style = ParagraphStyle(
            'BuyerHeader',
            parent=heading2_style,
            fontSize=12,
            textColor=colors.HexColor('#1a5490'),
            fontName='Helvetica-Bold',
            spaceAfter=6,
            spaceBefore=10
        )
        story.append(Paragraph(f"BUYER: {buyer}", buyer_header_style))
        
        buyer_data = buyer_elev_grade[buyer_elev_grade["Buyer"] == buyer]
        
        if not buyer_data.empty:
            elevations = sorted(buyer_data["Sub Elevation"].unique())
            
            # Summary table
            summary_data = []
            for elevation in elevations:
                elev_data = buyer_data[buyer_data["Sub Elevation"] == elevation]
                total_qty = elev_data['Total Weight'].sum()
                total_value = elev_data['Total Value'].sum()
                avg_price = elev_data['Price'].mean()
                summary_data.append({
                    'Elevation': elevation,
                    'Quantity': total_qty,
                    'Total_Value': total_value,
                    'Avg_Price': avg_price
                })
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                summary_table_data = [['Sub Elevation', 'Quantity (kg)', 'Total Value (LKR)', 'Avg Price (LKR)']]
                for _, row in summary_df.iterrows():
                    summary_table_data.append([
                        row['Elevation'],
                        f"{row['Quantity']:,.0f}",
                        f"{row['Total_Value']:,.0f}",
                        f"{row['Avg_Price']:,.2f}"
                    ])
                
                summary_table = Table(summary_table_data, colWidths=[1.3*inch, 1.2*inch, 1.3*inch, 1*inch])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3d6bb3')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                    ('ALIGN', (0,0), (0,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('PADDING', (0,0), (-1,-1), 5),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 0.15*inch))
            
            # Show ALL grades per elevation
            for elev_idx, elevation in enumerate(elevations):
                elev_data = buyer_data[buyer_data["Sub Elevation"] == elevation].sort_values('Total Weight', ascending=False)
                
                if not elev_data.empty:
                    elev_header_style = ParagraphStyle(
                        'ElevHeader',
                        parent=body_style,
                        fontSize=9,
                        textColor=colors.HexColor('#2c5aa0'),
                        fontName='Helvetica-Bold',
                        spaceAfter=4,
                        spaceBefore=6
                    )
                    story.append(Paragraph(f"Sub Elevation: {elevation}", elev_header_style))
                    
                    table_data = [['Grade', 'Quantity (kg)', 'Total Value (LKR)', 'Avg Price (LKR)', 'No. of Lots']]
                    for _, row in elev_data.iterrows():
                        table_data.append([
                            row['Grade'][:18],
                            f"{row['Total Weight']:,.0f}",
                            f"{row['Total Value']:,.0f}",
                            f"{row['Price']:,.2f}",
                            f"{int(row['Lot No'])}"
                        ])
                    
                    table = Table(table_data, colWidths=[1.2*inch, 1*inch, 1.1*inch, 1*inch, 0.8*inch])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3d6bb3')),
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
                    story.append(Spacer(1, 0.1*inch))
        
        if buyer_idx < len(all_buyers) - 1:
            story.append(PageBreak())

def generate_overall_market_summary(latest_df, story, heading1_style, heading2_style, body_style):
    """Summary Report: Overall Market Performance with MPB Highlighting"""
    story.append(Paragraph("SUMMARY REPORT: OVERALL MARKET PERFORMANCE", heading1_style))
    story.append(Spacer(1, 0.15*inch))
    
    # Overall market statistics
    total_catalogued = latest_df["Total Weight"].sum()
    total_sold = latest_df[latest_df["Status_Clean"] == "sold"]["Total Weight"].sum()
    total_unsold = latest_df[latest_df["Status_Clean"] == "unsold"]["Total Weight"].sum()
    total_outsold = latest_df[latest_df["Status_Clean"] == "outsold"]["Total Weight"].sum()
    total_sold_side = total_sold + total_outsold
    
    sold_pct = (total_sold_side / total_catalogued * 100) if total_catalogued > 0 else 0
    unsold_pct = (total_unsold / total_catalogued * 100) if total_catalogued > 0 else 0
    outsold_pct = (total_outsold / total_catalogued * 100) if total_catalogued > 0 else 0
    
    total_value = latest_df["Total Value"].sum()
    avg_price = latest_df[latest_df["Status_Clean"] == "sold"]["Price"].mean()
    
    # Overall summary table
    overall_data = [
        ['Metric', 'Value'],
        ['Total Catalogued (kg)', f"{total_catalogued:,.0f}"],
        ['Total Sold (kg)', f"{total_sold:,.0f}"],
        ['Total Outsold (kg)', f"{total_outsold:,.0f}"],
        ['Total Sold+Outsold (kg)', f"{total_sold_side:,.0f}"],
        ['Total Unsold (kg)', f"{total_unsold:,.0f}"],
        ['Sold %', f"{sold_pct:.2f}%"],
        ['Unsold %', f"{unsold_pct:.2f}%"],
        ['Outsold %', f"{outsold_pct:.2f}%"],
        ['Total Value (LKR)', f"{total_value:,.0f}"],
        ['Average Price (LKR)', f"{avg_price:,.2f}" if pd.notna(avg_price) else 'N/A']
    ]
    
    overall_table = Table(overall_data, colWidths=[2.5*inch, 2*inch])
    overall_table.setStyle(TableStyle([
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
    story.append(overall_table)
    story.append(Spacer(1, 0.2*inch))
    
    # Broker performance summary (with MPB highlighting)
    story.append(Paragraph("BROKER PERFORMANCE SUMMARY", heading2_style))
    story.append(Spacer(1, 0.1*inch))
    
    broker_summary = latest_df.groupby("Broker").apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Total_Value': x["Total Value"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean()
    }), include_groups=False).reset_index()
    
    broker_summary['Total_Sold_Side'] = broker_summary['Sold'] + broker_summary['Outsold']
    broker_summary['Sold_%'] = (broker_summary['Total_Sold_Side'] / broker_summary['Catalogued'] * 100).fillna(0)
    broker_summary['Market_Share_%'] = (broker_summary['Total_Value'] / total_value * 100).fillna(0)
    broker_summary = broker_summary.sort_values('Total_Value', ascending=False)
    
    # Identify MPB (highest market share)
    mpb = broker_summary.iloc[0]['Broker'] if not broker_summary.empty else None
    
    broker_table_data = [['Broker', 'Catalogued (kg)', 'Sold %', 'Market Share %', 'Total Value (LKR)']]
    for _, row in broker_summary.iterrows():
        is_mpb = row['Broker'] == mpb
        broker_table_data.append([
            f"{row['Broker']} {'(MPB)' if is_mpb else ''}",
            f"{row['Catalogued']:,.0f}",
            f"{row['Sold_%']:.2f}%",
            f"{row['Market_Share_%']:.2f}%",
            f"{row['Total_Value']:,.0f}"
        ])
    
    broker_table = Table(broker_table_data, colWidths=[1.5*inch, 1.2*inch, 1*inch, 1*inch, 1.2*inch])
    table_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1a5490')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('PADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
    ]
    
    # Highlight MPB row
    for i in range(1, len(broker_table_data)):
        if mpb and broker_table_data[i][0].startswith(mpb):
            table_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor('#fff3cd')))
            table_style.append(('TEXTCOLOR', (0,i), (-1,i), colors.HexColor('#856404')))
            table_style.append(('FONTNAME', (0,i), (-1,i), 'Helvetica-Bold'))
    
    broker_table.setStyle(TableStyle(table_style))
    story.append(broker_table)
    story.append(PageBreak())

def generate_broker_performance_summary(latest_df, story, heading1_style, heading2_style, body_style):
    """Summary Report: Detailed Broker Performance Comparison with MPB Highlighting"""
    story.append(Paragraph("SUMMARY REPORT: BROKER PERFORMANCE COMPARISON", heading1_style))
    story.append(Spacer(1, 0.15*inch))
    
    # Calculate broker performance by sub elevation
    broker_elev_perf = latest_df.groupby(["Broker", "Sub Elevation"]).apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Total_Value': x["Total Value"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean()
    }), include_groups=False).reset_index()
    
    broker_elev_perf['Total_Sold_Side'] = broker_elev_perf['Sold'] + broker_elev_perf['Outsold']
    broker_elev_perf['Sold_%'] = (broker_elev_perf['Total_Sold_Side'] / broker_elev_perf['Catalogued'] * 100).fillna(0)
    broker_elev_perf['Unsold_%'] = (broker_elev_perf['Unsold'] / broker_elev_perf['Catalogued'] * 100).fillna(0)
    
    # Identify MPB
    total_value = latest_df["Total Value"].sum()
    broker_totals = latest_df.groupby("Broker")["Total Value"].sum()
    mpb = broker_totals.idxmax() if not broker_totals.empty else None
    
    all_brokers = sorted(latest_df["Broker"].unique())
    
    for broker in all_brokers:
        is_mpb = broker == mpb
        broker_header_style = ParagraphStyle(
            'BrokerHeader',
            parent=heading2_style,
            fontSize=11,
            textColor=colors.HexColor('#856404') if is_mpb else colors.HexColor('#1a5490'),
            fontName='Helvetica-Bold',
            spaceAfter=6,
            spaceBefore=10
        )
        mpb_label = " (MPB - Market Performance Leader)" if is_mpb else ""
        story.append(Paragraph(f"BROKER: {broker}{mpb_label}", broker_header_style))
        
        broker_data = broker_elev_perf[broker_elev_perf["Broker"] == broker].sort_values('Sub Elevation')
        
        if not broker_data.empty:
            table_data = [['Sub Elevation', 'Catalogued (kg)', 'Sold %', 'Unsold %', 'Total Value (LKR)']]
            for _, row in broker_data.iterrows():
                table_data.append([
                    row['Sub Elevation'],
                    f"{row['Catalogued']:,.0f}",
                    f"{row['Sold_%']:.2f}%",
                    f"{row['Unsold_%']:.2f}%",
                    f"{row['Total_Value']:,.0f}"
                ])
            
            table = Table(table_data, colWidths=[1.5*inch, 1.2*inch, 1*inch, 1*inch, 1.2*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#856404') if is_mpb else colors.HexColor('#1a5490')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                ('ALIGN', (0,0), (0,-1), 'LEFT'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('PADDING', (0,0), (-1,-1), 5),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f5f5')])
            ]))
            story.append(table)
            story.append(Spacer(1, 0.15*inch))
        
        story.append(PageBreak())

def generate_fast_pdf_report(data, latest_df, output_filename="report.pdf", 
                            include_sections=None, highlight_broker="MPB"):
    """
    OPTIMIZED: Fast PDF generation with section selection - NO CHARTS, TABLES ONLY
    
    Args:
        data: Full dataset
        latest_df: Latest sale data
        output_filename: Output PDF filename
        include_sections: Dict of sections to include:
            - 'report1_sold_pct': Broker grade-wise sold percentages (Sub Elevation) - ALL GRADES
            - 'report2_unsold_pct': Broker grade-wise unsold percentages (Sub Elevation) - ALL GRADES
            - 'report3_outsold_pct': Broker grade-wise outsold percentages (Sub Elevation) - ALL GRADES
            - 'report4_sold_qty_price': Broker grade-wise sold quantities & avg prices (Sub Elevation) - ALL GRADES
            - 'report5_buyer_profiles': Outlots purchased buyer profiles (Grade wise, Sub Elevation) - ALL BUYERS & GRADES
            - 'summary_market': Overall market performance summary with MPB highlighting
            - 'summary_broker_perf': Broker performance comparison with MPB highlighting
        highlight_broker: Broker to highlight (default: "MPB")
    """
    
    # Default sections if none specified - include all 5 required reports
    if include_sections is None:
        include_sections = {
            'report1_sold_pct': True,
            'report2_unsold_pct': True,
            'report3_outsold_pct': True,
            'report4_sold_qty_price': True,
            'report5_buyer_profiles': True
        }
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=2.5*cm,
        bottomMargin=2.5*cm,
        title="MPBL MAIN TEA AUCTION Report"
    )
    
    # Styles
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1a5490'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
        spaceAfter=12
    )
    
    heading1_style = ParagraphStyle(
        'CustomHeading1',
        parent=styles['Heading2'],
        fontSize=13,
        textColor=colors.HexColor('#2c5aa0'),
        spaceAfter=10,
        spaceBefore=10,
        fontName='Helvetica-Bold'
    )
    
    heading2_style = ParagraphStyle(
        'CustomHeading2',
        parent=styles['Heading3'],
        fontSize=11,
        textColor=colors.HexColor('#3d6bb3'),
        spaceAfter=8,
        spaceBefore=8,
        fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=9,
        spaceAfter=6
    )
    
    story = []
    latest_sale = latest_df['Sale_No'].max()
    
    # ============================================================
    # TITLE PAGE
    # ============================================================
    
    story.append(Spacer(1, 0.5*inch))
    story.append(Paragraph("Mercantile Produce Brokers Pvt Ltd", heading2_style))
    story.append(Paragraph("MAIN AUCTION DETAILED REPORT", title_style))
    story.append(Spacer(1, 0.3*inch))
    
    # Count selected reports
    selected_reports = [k for k, v in include_sections.items() if v]
    report_type = f"Elevation-wise Analysis ({len(selected_reports)} report{'s' if len(selected_reports) != 1 else ''})"
    
    report_info = [
        ['Sale Number:', f"Sale {latest_sale}"],
        ['Report Date:', datetime.now().strftime('%B %d, %Y at %H:%M')],
        ['Report Type:', report_type],
        ['Data Period:', f"Sales {data['Sale_No'].min()} - {data['Sale_No'].max()}"]
    ]
    
    report_table = Table(report_info, colWidths=[3*inch, 3*inch], hAlign='LEFT')
    report_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#e8f4f8')),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#1a5490')),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#b8d4e8')),
        ('PADDING', (0,0), (-1,-1), 8)
    ]))
    
    story.append(report_table)
    story.append(PageBreak())
    
    # ============================================================
    # GENERATE SELECTED REPORTS (NO CHARTS - TABLES ONLY)
    # ============================================================
    
    # Report 1: Broker Grade-wise Sold Percentages (Sub Elevation)
    if include_sections.get('report1_sold_pct', False):
        generate_broker_grade_sold_pct(latest_df, story, heading1_style, heading2_style, body_style)
    
    # Report 2: Broker Grade-wise Unsold Percentages (Sub Elevation)
    if include_sections.get('report2_unsold_pct', False):
        generate_broker_grade_unsold_pct(latest_df, story, heading1_style, heading2_style, body_style)
    
    # Report 3: Broker Grade-wise Outsold Percentages (Sub Elevation)
    if include_sections.get('report3_outsold_pct', False):
        generate_broker_grade_outsold_pct(latest_df, story, heading1_style, heading2_style, body_style)
    
    # Report 4: Broker Grade-wise Sold Quantities & Avg Prices (Sub Elevation)
    if include_sections.get('report4_sold_qty_price', False):
        generate_broker_grade_sold_qty_price(latest_df, story, heading1_style, heading2_style, body_style)
    
    # Report 5: Outlots Purchased Buyer Profiles (Grade wise, Sub Elevation)
    if include_sections.get('report5_buyer_profiles', False):
        generate_buyer_grade_profiles(latest_df, story, heading1_style, heading2_style, body_style)
    
    # Summary Report 1: Overall Market Performance (with MPB highlighting)
    if include_sections.get('summary_market', False):
        generate_overall_market_summary(latest_df, story, heading1_style, heading2_style, body_style)
    
    # Summary Report 2: Broker Performance Comparison (with MPB highlighting)
    if include_sections.get('summary_broker_perf', False):
        generate_broker_performance_summary(latest_df, story, heading1_style, heading2_style, body_style)
    
    # Build PDF
    try:
        doc.build(story, canvasmaker=NumberedCanvas)
    except Exception as e:
        # Fallback without custom canvas
        doc.build(story)
    
    pdf_bytes = buffer.getvalue()
    buffer.close()
    
    return pdf_bytes

@st.cache_data
def load_all_sales(data_folder="sales_data"):
    all_data = []
    
    if not os.path.exists(data_folder):
        st.error(f"Folder '{data_folder}' not found. Please create it and add your sales data files.")
        return pd.DataFrame()
    
    for file in sorted(os.listdir(data_folder)):
        if file.startswith("~") or file.startswith("."):
            continue
            
        if (file.endswith(".csv") or file.endswith(".xlsx")) and re.search(r"Sale_(\d+)", file):
            try:
                sale_no = int(re.search(r"Sale_(\d+)", file).group(1))
                file_path = os.path.join(data_folder, file)
                
                if file.endswith(".csv"):
                    df = pd.read_csv(file_path)
                else:
                    df = pd.read_excel(file_path, sheet_name=0)
                
                df["Sale_No"] = sale_no
                all_data.append(df)
                
            except Exception as e:
                st.warning(f"Could not read file '{file}': {str(e)}")
                continue
    
    if not all_data:
        return pd.DataFrame()
    
    full_df = pd.concat(all_data, ignore_index=True)
    return full_df

def format_currency(val):
    """Enhanced currency formatting with M/B notation"""
    if pd.isna(val) or val == 0:
        return "LKR 0"
    
    if abs(val) >= 1_000_000_000:
        return f"LKR {val/1_000_000_000:.2f}B"
    elif abs(val) >= 1_000_000:
        return f"LKR {val/1_000_000:.2f}M"
    elif abs(val) >= 1_000:
        return f"LKR {val/1_000:.1f}K"
    else:
        return f"LKR {val:,.0f}"

def format_large_number(val):
    """Format large numbers with M/B notation"""
    if pd.isna(val) or val == 0:
        return "0"
    
    if abs(val) >= 1_000_000_000:
        return f"{val/1_000_000_000:.2f}B"
    elif abs(val) >= 1_000_000:
        return f"{val/1_000_000:.2f}M"
    elif abs(val) >= 1_000:
        return f"{val/1_000:.1f}K"
    else:
        return f"{val:,.0f}"

def format_number(val):
    return f"{val:,.2f}"

def calculate_sell_percentage(sold_qty, catalogued_qty):
    if catalogued_qty > 0:
        return (sold_qty / catalogued_qty) * 100
    return 0

def generate_ai_summary(df, broker="MPB"):
    broker_df = df[df["Broker"] == broker]
    total_value = (broker_df["Total Weight"] * broker_df["Price"]).sum()
    avg_price = broker_df["Price"].mean()
    overall_avg = df["Price"].mean()
    diff = avg_price - overall_avg
    market_share = total_value / (df["Total Weight"] * df["Price"]).sum() * 100
    return (
        f"In the latest sale, **{broker}** achieved a market share of **{market_share:.2f}%** "
        f"with an average price of **LKR {avg_price:,.2f}/kg**, which is "
        f"{'higher' if diff > 0 else 'lower'} than the market average by **LKR {abs(diff):,.2f}/kg**. "
        f"Total value sold: **{format_currency(total_value)}**."
    )

def get_base_trade_mark(tm):
    """Extract base trade mark by removing trailing single letter"""
    if tm and len(tm) > 1 and tm[-1].isalpha() and tm[-1].isupper():
        if not tm[-2].isalpha():
            return tm[:-1]
    return tm

# --- Load Data ---
data = load_all_sales("sales_data")

if data.empty:
    st.warning("No sale files found in `sales_data/` folder. Please add files like `Sale_42.csv` or `Sale_42.xlsx`.")
    st.stop()

# --- Data Preprocessing ---
data["Broker"] = data["Broker"].astype(str).str.strip()
data["Price"] = pd.to_numeric(data["Price"], errors="coerce")
data["Total Weight"] = pd.to_numeric(data["Total Weight"], errors="coerce")
data["Total Value"] = data["Total Weight"] * data["Price"]
data["Category"] = data["Category"].astype(str)
data["Grade"] = data["Grade"].astype(str)
data["Sub Elevation"] = data["Sub Elevation"].astype(str)

if "Selling Mark" in data.columns:
    data["Selling Mark"] = data["Selling Mark"].fillna("").astype(str).str.strip()
else:
    data["Selling Mark"] = ""

data["Status"] = data["Status"].fillna("").astype(str).str.strip() if "Status" in data.columns else ""
data["Status_Clean"] = data["Status"].str.lower()

if "Trade Mark" in data.columns:
    data["Trade Mark"] = data["Trade Mark"].fillna("").astype(str).str.strip()
    data["Trade Mark"] = data["Trade Mark"].replace(['nan', 'NaN', 'None', ''], pd.NA)
else:
    data["Trade Mark"] = pd.NA

data["Valuation price"] = pd.to_numeric(data["Valuation price"], errors="coerce") if "Valuation price" in data.columns else 0
data["Asking Price"] = pd.to_numeric(data["Asking Price"], errors="coerce") if "Asking Price" in data.columns else 0

latest_sale = data["Sale_No"].max()
latest_df = data[data["Sale_No"] == latest_sale]

# --- Sidebar ---
st.sidebar.title(" Filters")
selected_broker = st.sidebar.multiselect("Select Broker(s)", sorted(data["Broker"].unique()), default=sorted(data["Broker"].unique()))
selected_category = st.sidebar.multiselect("Select Category", sorted(data["Category"].unique()), default=sorted(data["Category"].unique()))
filtered = data[data["Broker"].isin(selected_broker) & data["Category"].isin(selected_category)]

# --- Dashboard Tabs - Merged Overview and MPB Intelligence ---
tabs = st.tabs([
    "Market Overview & MPBL Metrics",
    "Broker Performance",
    "Elevation & Category",
    "Buyer Insights",
    "Selling Mark Analysis",
    "Price Trends"
])

# OVERVIEW & MPB INTELLIGENCE TAB
with tabs[0]:
    st.header(f"Sale {latest_sale} - Market Overview & MPBL Metrics")
    
    # Top Level Metrics with enhanced formatting
    total_val = latest_df["Total Value"].sum()
    avg_price = latest_df["Price"].mean()
    total_weight = latest_df["Total Weight"].sum()
    
    # Calculate sold/unsold
    sold_df = latest_df[latest_df["Status_Clean"] == "sold"]
    unsold_df = latest_df[latest_df["Status_Clean"] == "unsold"]
    outsold_df = latest_df[latest_df["Status_Clean"] == "outsold"]
    
    sold_weight = sold_df["Total Weight"].sum()
    unsold_weight = unsold_df["Total Weight"].sum()
    outsold_weight = outsold_df["Total Weight"].sum()
    
    # Calculate Sold % as (Sold + Outsold) / Total
    total_sold_side_weight = sold_weight + outsold_weight
    sold_pct = (total_sold_side_weight / total_weight * 100) if total_weight > 0 else 0
    avg_price = avg_price if pd.notna(avg_price) else 0
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Market Value", format_currency(total_val) if pd.notna(total_val) else "N/A")
    col2.metric("Average Price/kg", f"LKR {avg_price:,.2f}" if avg_price > 0 else "N/A")
    col3.metric("Total Weight", f"{format_large_number(total_weight)} kg")
    col4.metric("Sold %", f"{sold_pct:.1f}%")
    col5.metric("Total Lots", f"{len(latest_df):,}")
    
    st.markdown("---")
    
    # Market Share and Status Distribution
    col1, col2 = st.columns(2)
    
    with col1:
        market_share = latest_df.groupby("Broker")["Total Value"].sum().sort_values(ascending=False).reset_index()
        fig = px.pie(market_share, names="Broker", values="Total Value", 
                     title="Market Share by Broker (Value)",
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_traces(textinfo="percent+label", pull=[0.1 if b == "MPB" else 0 for b in market_share["Broker"]])
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        status_dist = pd.DataFrame({
            'Status': ['Sold', 'Unsold', 'Outsold'],
            'Weight': [sold_weight, unsold_weight, outsold_weight]
        })
        status_dist = status_dist[status_dist['Weight'] > 0]
        
        fig2 = px.pie(status_dist, values='Weight', names='Status',
                      title='Overall Sale Status Distribution',
                      color='Status',
                      color_discrete_map={'Sold': '#28a745', 'Unsold': '#dc3545', 'Outsold': '#ffc107'})
        fig2.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig2, use_container_width=True)
    
    st.markdown("---")
    
    # MPB Intelligence Section
    st.subheader("MPBL MAIN SALE DASHBOARD")
    st.info(generate_ai_summary(latest_df, broker="MPB"))

    mpb_df = data[data["Broker"] == "MPB"]
    
    # MPB Performance Metrics
    mpb_latest = latest_df[latest_df["Broker"] == "MPB"]
    if not mpb_latest.empty:
        mpb_sold = mpb_latest[mpb_latest["Status_Clean"] == "sold"]
        mpb_total_value = mpb_latest["Total Value"].sum()
        mpb_avg_price = mpb_sold["Price"].mean() if not mpb_sold.empty else 0
        mpb_sold_pct = (mpb_sold["Total Weight"].sum() / mpb_latest["Total Weight"].sum() * 100) if mpb_latest["Total Weight"].sum() > 0 else 0
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("MPB Total Value", format_currency(mpb_total_value))
        col2.metric("MPB Avg Price", f"LKR {mpb_avg_price:,.2f}")
        col3.metric("MPB Sold %", f"{mpb_sold_pct:.1f}%")
        col4.metric("MPB Market Share", f"{(mpb_total_value/total_val*100):.1f}%")
    
    # MPB Trend Analysis
    st.subheader(" MPBL Performance")
    
    col1, col2 = st.columns(2)
    
    with col1:
        avg_price_trend = data.groupby(["Sale_No", "Broker"], as_index=False)["Price"].mean()
        fig = px.line(avg_price_trend, x="Sale_No", y="Price", color="Broker", markers=True,
                      title="Average Price Trend by Broker", color_discrete_sequence=px.colors.qualitative.Vivid)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        top_buyers = mpb_df.groupby("Buyer", as_index=False)["Total Value"].sum().sort_values("Total Value", ascending=False).head(10)
        fig2 = px.bar(top_buyers, x="Buyer", y="Total Value", title="Top 10 Buyers - MPB",
                      color_discrete_sequence=["#007bff"])
        fig2.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig2, use_container_width=True)
    
    # MPB Grade Performance
    st.subheader(" MPBL Grade Performance")
    if not mpb_latest.empty:
        mpb_grade_perf = mpb_latest.groupby("Grade").agg({
            "Total Weight": "sum",
            "Price": "mean",
            "Total Value": "sum",
            "Lot No": "count"
        }).sort_values("Total Value", ascending=False).head(15).reset_index()
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig_grade_value = px.bar(mpb_grade_perf, x="Grade", y="Total Value",
                                   title="MPB Top Grades by Value",
                                   color="Total Value",
                                   color_continuous_scale="Viridis")
            fig_grade_value.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_grade_value, use_container_width=True)
        
        with col2:
            fig_grade_price = px.bar(mpb_grade_perf, x="Grade", y="Price",
                                   title="MPB Average Price by Grade",
                                   color="Price",
                                   color_continuous_scale="Plasma")
            fig_grade_price.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_grade_price, use_container_width=True)
    
    st.markdown("---")
    
    # Elevation-wise Performance Summary
    st.subheader(" Elevation-wise Performance Summary")
    
    elev_summary = latest_df.groupby("Sub Elevation").apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean(),
        'Total_Lots': len(x)
    }), include_groups=False).reset_index()
    
    # Calculate Sold % as (Sold + Outsold) / Total
    elev_summary['Total_Sold_Side'] = elev_summary['Sold'] + elev_summary['Outsold']
    elev_summary['Sold %'] = (elev_summary['Total_Sold_Side'] / elev_summary['Catalogued'] * 100).fillna(0)
    elev_summary['Unsold %'] = (elev_summary['Unsold'] / elev_summary['Catalogued'] * 100).fillna(0)
    elev_summary['Outsold %'] = (elev_summary['Outsold'] / elev_summary['Catalogued'] * 100).fillna(0)
    
    # Display elevation metrics with enhanced formatting
    elev_display = elev_summary.copy()
    elev_display['Catalogued'] = elev_display['Catalogued'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
    elev_display['Sold'] = elev_display['Sold'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
    elev_display['Unsold'] = elev_display['Unsold'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
    elev_display['Outsold'] = elev_display['Outsold'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
    elev_display['Avg_Price'] = elev_display['Avg_Price'].fillna(0).apply(lambda x: f"LKR {x:,.2f}" if x > 0 else "N/A")
    elev_display['Sold %'] = elev_display['Sold %'].fillna(0).apply(lambda x: f"{x:.1f}%")
    elev_display['Unsold %'] = elev_display['Unsold %'].fillna(0).apply(lambda x: f"{x:.1f}%")
    elev_display['Outsold %'] = elev_display['Outsold %'].fillna(0).apply(lambda x: f"{x:.1f}%")
    
    st.dataframe(elev_display, use_container_width=True)
    
    # Elevation comparison charts
    col1, col2 = st.columns(2)
    
    with col1:
        fig_elev_sold = px.bar(elev_summary, x='Sub Elevation', y=['Sold %', 'Unsold %', 'Outsold %'],
                               title='Elevation-wise Sale Status %',
                               labels={'value': 'Percentage', 'variable': 'Status'},
                               barmode='stack',
                               color_discrete_map={'Sold %': '#28a745', 'Unsold %': '#dc3545', 'Outsold %': '#ffc107'})
        st.plotly_chart(fig_elev_sold, use_container_width=True)
    
    with col2:
        fig_elev_price = px.bar(elev_summary, x='Sub Elevation', y='Avg_Price',
                                title='Average Price by Elevation',
                                color='Avg_Price',
                                color_continuous_scale='Viridis')
        st.plotly_chart(fig_elev_price, use_container_width=True)

# BROKER PERFORMANCE
    st.subheader(" Quick Report Generation")

    col1, col2, col3 = st.columns(3)

    with col1:
        report_scope = st.selectbox(
            " Data Scope",
            ["Current Sale Only", "Last 3 Sales", "Last 5 Sales", "All Available Sales"],
            index=0,
            help="Current sale is fastest (5-8 seconds)"
        )

    with col2:
        report_format = st.selectbox(
            " Report Format",
            ["Standard (Fast)", "Detailed (More Data)"],
            index=0,
            help="Standard format is optimized for speed"
        )

    with col3:
        st.markdown("#####  Report Selection")
        st.info("Select which reports to include in PDF")

    # Quick templates
    st.markdown("---")
    st.subheader(" Quick Generation Templates")

    template_col1, template_col2, template_col3 = st.columns(3)

    with template_col1:
        quick_report_btn = st.button(
            " Quick Report - Current Sale",
            use_container_width=True,
            type="primary",
            help="Generate report for current sale only (5-8 seconds)"
        )

    with template_col2:
        weekly_report_btn = st.button(
            " Weekly Report - Last 3 Sales",
            use_container_width=True,
            help="Generate report for last 3 sales (10-12 seconds)"
        )

    with template_col3:
        full_report_btn = st.button(
            " Full Report - All Sales",
            use_container_width=True,
            help="Generate comprehensive report (15-20 seconds)"
        )

    # Report Selection Checkboxes
    st.markdown("---")
    st.subheader(" Select Reports to Include")

    col1, col2, col3 = st.columns(3)

    with col1:
        report1 = st.checkbox("Report 1: Broker Grade-wise Sold % (Sub Elevation)", value=True, 
                             help="Each broker's grade wise sold percentages by Sub Elevation")
        report2 = st.checkbox("Report 2: Broker Grade-wise Unsold % (Sub Elevation)", value=True,
                             help="Each broker's grade wise unsold percentages by Sub Elevation")

    with col2:
        report3 = st.checkbox("Report 3: Broker Grade-wise Outsold % (Sub Elevation)", value=True,
                             help="Each broker's grade wise outsold percentages by Sub Elevation")
        report4 = st.checkbox("Report 4: Broker Grade-wise Sold Qty & Avg Prices (Sub Elevation)", value=True,
                             help="Each broker's grade wise sold quantities and average prices by Sub Elevation")

    with col3:
        report5 = st.checkbox("Report 5: Outlots Purchased Buyer Profiles (Grade wise, Sub Elevation)", value=True,
                             help="Outlots purchased buyer profiles, grade wise by Sub Elevation")
        st.markdown("---")
        st.markdown("** Summary Reports (Optional):**")
        summary_market = st.checkbox(" Overall Market Performance Summary", value=False,
                                    help="Overall market statistics with MPB highlighting")
        summary_broker_perf = st.checkbox(" Broker Performance Comparison", value=False,
                                          help="Detailed broker performance by Sub Elevation with MPB highlighting")

    # Main generate button
    st.markdown("---")
    generate_col1, generate_col2, generate_col3 = st.columns([1, 2, 1])

    with generate_col2:
        generate_button = st.button(
            " GENERATE PROFESSIONAL PDF REPORT",
            type="primary",
            use_container_width=True,
            help="Generate PDF with selected reports only (faster generation)"
        )

    # Report generation logic
    if generate_button or quick_report_btn or weekly_report_btn or full_report_btn:
        try:
            # Check reportlab installation
            try:
                from reportlab.lib import colors
                from reportlab.lib.pagesizes import A4
            except ImportError:
                st.error("""
                 **ReportLab library not found!**
            
                Please install it using:
                ```bash
                pip install reportlab
                ```
            
                Then restart your Streamlit application.
                """)
                st.stop()
        
            # Determine scope based on button clicked
            if quick_report_btn:
                report_scope = "Current Sale Only"
            elif weekly_report_btn:
                report_scope = "Last 3 Sales"
            elif full_report_btn:
                report_scope = "All Available Sales"
        
            # Generate report with progress tracking
            with st.spinner(f" Generating {report_scope} report..."):
                import time
            
                # Progress indicators
                progress_bar = st.progress(0)
                status_text = st.empty()
                time_estimate = st.empty()
            
                # Estimate completion time
                if report_scope == "Current Sale Only":
                    estimated_time = "8-12 seconds"
                elif report_scope in ["Last 3 Sales", "Last 5 Sales"]:
                    estimated_time = "15-20 seconds"
                else:
                    estimated_time = "25-30 seconds"
            
                time_estimate.info(f" Estimated completion time: {estimated_time}")
            
                # Step 1: Filter data
                status_text.text(" Processing sale data...")
                progress_bar.progress(15)
                time.sleep(0.2)
            
                if report_scope == "Current Sale Only":
                    report_data = data[data["Sale_No"] == latest_sale]
                elif report_scope == "Last 3 Sales":
                    recent_sales = sorted(data["Sale_No"].unique())[-3:]
                    report_data = data[data["Sale_No"].isin(recent_sales)]
                elif report_scope == "Last 5 Sales":
                    recent_sales = sorted(data["Sale_No"].unique())[-5:]
                    report_data = data[data["Sale_No"].isin(recent_sales)]
                else:
                    report_data = data
            
                # Step 2: Calculate metrics
                status_text.text(" Calculating performance metrics...")
                progress_bar.progress(30)
                time.sleep(0.2)
            
                # Step 3: Generate broker analysis
                status_text.text("Analyzing all 8 brokers...")
                progress_bar.progress(50)
                time.sleep(0.3)
            
                # Step 4: Generate PDF
                status_text.text(" Creating PDF document...")
                progress_bar.progress(70)
            
                # Import the optimized function
                from reportlab.lib import colors
                from reportlab.lib.pagesizes import A4
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.units import inch, cm
                from reportlab.lib.enums import TA_CENTER
                from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
                from reportlab.pdfgen import canvas
                import io
            
                # Prepare selected sections
                include_sections = {
                    'report1_sold_pct': report1,
                    'report2_unsold_pct': report2,
                    'report3_outsold_pct': report3,
                    'report4_sold_qty_price': report4,
                    'report5_buyer_profiles': report5,
                    'summary_market': summary_market,
                    'summary_broker_perf': summary_broker_perf
                }
            
                # Check if at least one report is selected (excluding charts)
                report_selections = [report1, report2, report3, report4, report5]
                if not any(report_selections):
                    st.warning(" Please select at least one report to generate!")
                    st.stop()
            
                # Call the optimized PDF generator
                pdf_data = generate_fast_pdf_report(
                    report_data,
                    latest_df,
                    output_filename=f"report_sale_{latest_sale}.pdf",
                    include_sections=include_sections
                )
            
                # Step 5: Finalize
                status_text.text(" Finalizing document...")
                progress_bar.progress(95)
                time.sleep(0.2)
            
                progress_bar.progress(100)
                status_text.text(" Report generation complete!")
                time.sleep(0.3)
            
                # Clear progress indicators
                progress_bar.empty()
                status_text.empty()
                time_estimate.empty()
            
                # Success message
                st.success(f"""
                 **PDF Report Generated Successfully!**
            
                 Scope: {report_scope}  
                 Pages: Approximately 12-15 pages  
                 Size: {len(pdf_data) / 1024:.1f} KB  
                 Generated: {datetime.now().strftime('%H:%M:%S')}
                """)
            
                # Download section
                st.markdown("---")
                st.subheader(" Download Your Report")
            
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"Tea_Auction_Report_Sale_{latest_sale}_{timestamp}.pdf"
            
                col1, col2, col3 = st.columns([1, 2, 1])
            
                with col2:
                    st.download_button(
                        label=" DOWNLOAD COMPLETE PDF REPORT",
                        data=pdf_data,
                        file_name=filename,
                        mime="application/pdf",
                        use_container_width=True,
                        help="Download your comprehensive PDF report"
                    )
            
                # Report summary
                st.markdown("---")
                st.subheader(" Report Contents Summary")
            
                summary_col1, summary_col2, summary_col3 = st.columns(3)
            
                with summary_col1:
                    st.markdown("""
                    ** Market Overview**
                    - Executive summary
                    - Market metrics
                    - Key performance indicators
                    - Overall elevation performance
                    """)
            
                with summary_col2:
                    st.markdown("""
                    **All 8 Brokers**
                    - Performance by elevation
                    - Grade breakdown per elevation
                    - Sold/Unsold/Outsold % per elevation
                    - Top 8 grades per elevation
                    """)
            
                with summary_col3:
                    st.markdown("""
                    ** Buyer Analysis**
                    - Top 5 buyers
                    - Elevation preferences
                    - Grade patterns by elevation
                    - Price analysis by elevation
                    """)
            
                # Additional details
                with st.expander(" Detailed Report Breakdown"):
                    st.markdown(f"""
                    **Report Details:**
                
                    1. **Title Page**
                       - Company branding
                       - Sale number: {latest_sale}
                       - Generation timestamp
                       - Data coverage period
                
                    2. **Executive Summary (Page 2)**
                       - Total market value
                       - Average prices
                       - Sell-through rates
                       - Key insights
                
                    3. **Broker Performance (Page 3)**
                       - All 8 brokers comparison table
                       - Total quantities and percentages
                       - Average prices achieved
                       - Lot counts
                
                    4. **Broker Elevation & Grade Analysis (Pages 4-12)**
                       - Each broker gets dedicated pages
                       - Breakdown by elevation
                       - Each elevation shows:
                         * Top 8 grades
                         * Sold/Unsold/Outsold percentages
                         * Average prices by grade
                         * Lot counts
                
                    5. **Overall Elevation Analysis (Page 13)**
                       - Market-wide elevation performance
                       - Quantity and percentage breakdowns
                       - Price comparisons across elevations
                
                    6. **Buyer Profiles by Elevation (Pages 14-17)**
                       - Top 5 buyers detailed analysis
                       - Breakdown by elevation for each buyer
                       - Grade preferences per elevation
                       - Average prices paid by elevation
                       - Purchase quantities by elevation & grade
                
                    7. **Recommendations (Page 18)**
                       - Elevation-based insights
                       - Grade recommendations by elevation
                       - Broker performance insights
                       - Strategic suggestions
                
                    **Footer on Every Page:**
                    - Company name
                    - Generation date/time
                    - Page numbers (Page X of Y)
                    """)
            
                # Save to history (optional)
                if 'report_history' not in st.session_state:
                    st.session_state.report_history = []
            
                st.session_state.report_history.append({
                    'title': f"Sale {latest_sale} Report",
                    'date': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'scope': report_scope,
                    'size': f"{len(pdf_data) / 1024:.1f} KB"
                })
            
        except Exception as e:
            st.error(f"""
             **Error generating PDF report:**
        
            {str(e)}
        
            **Troubleshooting Steps:**
            1. Ensure `reportlab` is installed: `pip install reportlab`
            2. Check that data files are accessible
            3. Verify sufficient memory available
            4. Try "Current Sale Only" for faster generation
            """)
        
            # Show detailed error for debugging
            with st.expander(" Technical Error Details"):
                import traceback
                st.code(traceback.format_exc())

    # Report history section
    if 'report_history' in st.session_state and st.session_state.report_history:
        st.markdown("---")
        st.subheader(" Recent Report Generation History")
    
        # Show last 5 reports
        history_data = []
        for report in st.session_state.report_history[-5:]:
            history_data.append([
                report.get('title', 'Report'),
                report.get('date', 'N/A'),
                report.get('scope', 'N/A'),
                report.get('size', 'N/A')
            ])
    
        history_df = pd.DataFrame(
            history_data,
            columns=['Report', 'Generated', 'Scope', 'Size']
        )
    
        st.dataframe(history_df, use_container_width=True, hide_index=True)
    
        if st.button(" Clear History"):
            st.session_state.report_history = []
            st.rerun()

    # Tips and best practices
    st.markdown("---")
    with st.expander(" PDF Generation Tips & Best Practices"):
        st.markdown("""
        ** For Fastest Generation:**
        - Use "Current Sale Only" (8-12 seconds)
        - Standard format is optimized for speed
        - All 8 brokers with elevation breakdown included
    
        ** Understanding Elevation-Based Reports:**
        - Each broker's data is organized by elevation first
        - Within each elevation, top 8 grades are shown
        - **Sold %**: Percentage of catalogued quantity sold per elevation
        - **Unsold %**: Percentage that remained unsold per elevation
        - **Outsold %**: Percentage that went to competing brokers per elevation
        - Buyers show their purchasing patterns by elevation
    
        ** Best Use Cases:**
        - **Daily Operations**: Current Sale Only (elevation focus)
        - **Weekly Reviews**: Last 3 Sales (elevation trends)
        - **Monthly Analysis**: Last 5 Sales (elevation performance)
        - **Comprehensive Audit**: All Sales (complete elevation history)
    
        ** Report Format:**
        - Professional A4 size (portrait)
        - 15-20 pages typical (with elevation breakdown)
        - Clean tables organized by elevation
        - Page numbers on every page
        - Automatic timestamp
    
        ** File Management:**
        - Reports are generated fresh each time
        - Download and save important reports
        - Filename includes sale number and timestamp
        - Average file size: 75-200 KB (more data with elevations)
    
        ** Technical Notes:**
        - Requires `reportlab` library
        - Elevation-grade calculations optimized
        - All 8 brokers processed in single pass
        - Buyer-elevation data pre-calculated
        - Can generate hundreds of reports per hour
        """)

    # Performance information
    st.markdown("---")
    st.info("""
    ** Performance Optimized - Elevation-Based Analysis:**  
    This PDF generator provides comprehensive elevation-wise breakdown:
    -  All 8 brokers with elevation analysis
    -  Each elevation shows grade performance
    -  Sold/Unsold/Outsold % per elevation
    -  Buyer preferences by elevation
    -  Pre-calculated metrics for speed
    -  Progress tracking with time estimates
    -  15-20 pages of detailed insights
    """)

    # Footer
    st.markdown("---")
    st.caption(" Professional PDF Reports | Optimized for Tea Auction Business Intelligence")

    # Remove the problematic pie chart at the end that was causing the error
    st.success(" Dashboard Loaded Successfully with OKLO MAIN AUCTION DATA")


with tabs[1]:
    st.header("Broker Performance Analysis")
    
    # Broker-wise Grade Performance

    
    broker_grade_analysis = latest_df.groupby(["Broker", "Grade"]).apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean(),
        'Lots': len(x)
    }), include_groups=False).reset_index()
    
    # Calculate Sold % as (Sold + Outsold) / Total
    broker_grade_analysis['Total_Sold_Side'] = broker_grade_analysis['Sold'] + broker_grade_analysis['Outsold']
    broker_grade_analysis['Sold %'] = (broker_grade_analysis['Total_Sold_Side'] / broker_grade_analysis['Catalogued'] * 100).fillna(0)
    broker_grade_analysis['Unsold %'] = (broker_grade_analysis['Unsold'] / broker_grade_analysis['Catalogued'] * 100).fillna(0)
    broker_grade_analysis['Outsold %'] = (broker_grade_analysis['Outsold'] / broker_grade_analysis['Catalogued'] * 100).fillna(0)
    
    # Broker selector for detailed view
    selected_broker_view = st.selectbox("Select Broker for Detailed Grade Analysis", 
                                        sorted(latest_df["Broker"].unique()),
                                        key="broker_grade_view")
    
    broker_data = broker_grade_analysis[broker_grade_analysis["Broker"] == selected_broker_view]
    
    if not broker_data.empty:
        # Top performing grades
        top_grades = broker_data.nlargest(15, 'Sold')
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Stacked bar for sold/unsold/outsold by grade
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(name='Sold', x=top_grades['Grade'], y=top_grades['Sold %'],
                                 marker_color='#28a745'))
            fig1.add_trace(go.Bar(name='Unsold', x=top_grades['Grade'], y=top_grades['Unsold %'],
                                 marker_color='#dc3545'))
            fig1.add_trace(go.Bar(name='Outsold', x=top_grades['Grade'], y=top_grades['Outsold %'],
                                 marker_color='#ffc107'))
            fig1.update_layout(barmode='stack', 
                              title=f'{selected_broker_view} - Grade-wise Status % (Top 15 Grades)',
                              xaxis_title='Grade', yaxis_title='Percentage',
                              xaxis_tickangle=-45)
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            # Average price by grade
            fig2 = px.bar(top_grades, x='Grade', y='Avg_Price',
                         title=f'{selected_broker_view} - Average Price by Grade',
                         color='Avg_Price',
                         color_continuous_scale='Blues')
            fig2.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig2, use_container_width=True)
        
        # Detailed grade table with enhanced formatting
        st.markdown(f"###  Detailed Grade Performance - {selected_broker_view}")
        display_data = broker_data[['Grade', 'Catalogued', 'Sold', 'Unsold', 'Outsold', 
                                     'Sold %', 'Unsold %', 'Outsold %', 'Avg_Price', 'Lots']].copy()
        
        # Format the display data with null handling
        display_data_formatted = display_data.copy()
        display_data_formatted['Catalogued'] = display_data_formatted['Catalogued'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
        display_data_formatted['Sold'] = display_data_formatted['Sold'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
        display_data_formatted['Unsold'] = display_data_formatted['Unsold'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
        display_data_formatted['Outsold'] = display_data_formatted['Outsold'].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
        display_data_formatted['Avg_Price'] = display_data_formatted['Avg_Price'].fillna(0).apply(lambda x: f"LKR {x:,.2f}" if x > 0 else "N/A")
        display_data_formatted['Sold %'] = display_data_formatted['Sold %'].fillna(0).apply(lambda x: f"{x:.1f}%")
        display_data_formatted['Unsold %'] = display_data_formatted['Unsold %'].fillna(0).apply(lambda x: f"{x:.1f}%")
        display_data_formatted['Outsold %'] = display_data_formatted['Outsold %'].fillna(0).apply(lambda x: f"{x:.1f}%")
        display_data_formatted['Lots'] = display_data_formatted['Lots'].fillna(0).apply(lambda x: f"{x:,.0f}")
        
        st.dataframe(display_data_formatted, use_container_width=True)
    
    st.markdown("---")
    
    # Broker comparison by category
    st.subheader(" Broker Performance by Category")
    broker_perf = filtered.groupby(["Broker", "Category"], as_index=False).agg({
        "Total Weight": "sum",
        "Price": "mean",
        "Total Value": "sum"
    })
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig = px.bar(broker_perf, x="Broker", y="Total Value", color="Category",
                     title=f"Broker Performance by Category - Sale {latest_sale}",
                     color_discrete_sequence=px.colors.qualitative.Safe)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        fig = px.bar(broker_perf, x="Broker", y="Total Weight", color="Category",
                     title=f"Broker Weight Distribution by Category",
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig, use_container_width=True)


# BROKER PERFORMANCE
with tabs[2]:
    st.header("Elevation & Category Performance")
    
    # Broker quantity performance summary

    
    # Calculate broker performance metrics
    broker_performance = latest_df.groupby("Broker").apply(lambda x: pd.Series({
        'Total_Quantity': x["Total Weight"].sum(),
        'Sold_Quantity': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Unsold_Quantity': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Outsold_Quantity': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Total_Lots': len(x),
        'Sold_Lots': len(x[x["Status_Clean"] == "sold"]),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean()
    }), include_groups=False).reset_index()
    
    # Calculate percentages and additional metrics
    # Sold % should include both Sold + Outsold
    broker_performance['Total_Sold_Side_Quantity'] = broker_performance['Sold_Quantity'] + broker_performance['Outsold_Quantity']
    broker_performance['Sold_Percentage'] = (broker_performance['Total_Sold_Side_Quantity'] / broker_performance['Total_Quantity'] * 100).fillna(0)
    broker_performance['Unsold_Percentage'] = (broker_performance['Unsold_Quantity'] / broker_performance['Total_Quantity'] * 100).fillna(0)
    broker_performance['Outsold_Percentage'] = (broker_performance['Outsold_Quantity'] / broker_performance['Total_Quantity'] * 100).fillna(0)
    broker_performance['Sold_Side_Percentage'] = broker_performance['Sold_Percentage']  # Same as Sold_Percentage now
    
    # Highlight MPB in the data
    broker_performance['Is_MPB'] = broker_performance['Broker'] == 'MPB'
    
    # Top metrics cards for all brokers
    st.markdown("###  Key Quantity Metrics - All Brokers")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        total_market_quantity = broker_performance['Total_Quantity'].sum()
        st.metric("Total Market Quantity", f"{format_large_number(total_market_quantity)} kg")
    with col2:
        total_sold_quantity = broker_performance['Sold_Quantity'].sum()
        st.metric("Sold Quantity (Sold Only)", f"{format_large_number(total_sold_quantity)} kg")
    with col3:
        total_sold_side_quantity = broker_performance['Total_Sold_Side_Quantity'].sum()
        st.metric("Sold + Outsold Quantity", f"{format_large_number(total_sold_side_quantity)} kg")
    with col4:
        # Market Sold % should include outsold
        market_sold_percentage = (total_sold_side_quantity / total_market_quantity * 100) if total_market_quantity > 0 else 0
        st.metric("Market Sold %", f"{market_sold_percentage:.1f}%")
    with col5:
        market_sold_side_percentage = (total_sold_side_quantity / total_market_quantity * 100) if total_market_quantity > 0 else 0
        st.metric("Sold Side %", f"{market_sold_side_percentage:.1f}%")
    
    st.markdown("---")
    
    # Broker Quantity Distribution Charts
    st.subheader(" Broker Quantity Distribution")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Total quantity by broker with MPB highlight
        fig_total_qty = px.bar(broker_performance, x='Broker', y='Total_Quantity',
                             title='Total Quantity by Broker (kg) - MPB Highlighted',
                             color='Is_MPB',
                             color_discrete_map={True: '#FF6B6B', False: '#4ECDC4'},
                             text='Total_Quantity')
        fig_total_qty.update_traces(texttemplate='%{text:.2s}kg', textposition='outside')
        fig_total_qty.update_layout(showlegend=False, xaxis_tickangle=-45)
        st.plotly_chart(fig_total_qty, use_container_width=True)
    
    with col2:
        # Sold side quantity (Sold + Outsold)
        fig_sold_side = px.bar(broker_performance, x='Broker', y='Total_Sold_Side_Quantity',
                             title='Sold + Outsold Quantity by Broker (kg)',
                             color='Is_MPB',
                             color_discrete_map={True: '#FF6B6B', False: '#45B7D1'},
                             text='Total_Sold_Side_Quantity')
        fig_sold_side.update_traces(texttemplate='%{text:.2s}kg', textposition='outside')
        fig_sold_side.update_layout(showlegend=False, xaxis_tickangle=-45)
        st.plotly_chart(fig_sold_side, use_container_width=True)
    
    # Stacked quantity breakdown
    st.subheader(" Quantity Status Breakdown by Broker")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Stacked bar chart for quantity status
        status_quantity = broker_performance[['Broker', 'Sold_Quantity', 'Unsold_Quantity', 'Outsold_Quantity']].melt(
            id_vars=['Broker'], 
            var_name='Status', 
            value_name='Quantity'
        )
        
        fig_stacked_qty = px.bar(status_quantity, x='Broker', y='Quantity', color='Status',
                               title='Quantity Status Distribution by Broker (kg)',
                               color_discrete_map={
                                   'Sold_Quantity': '#28a745',
                                   'Unsold_Quantity': '#dc3545', 
                                   'Outsold_Quantity': '#ffc107'
                               },
                               barmode='stack')
        fig_stacked_qty.update_layout(xaxis_tickangle=-45, legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ))
        st.plotly_chart(fig_stacked_qty, use_container_width=True)
    
    with col2:
        # Percentage stacked bar chart
        status_percentage = broker_performance[['Broker', 'Sold_Percentage', 'Unsold_Percentage', 'Outsold_Percentage']].melt(
            id_vars=['Broker'], 
            var_name='Status', 
            value_name='Percentage'
        )
        
        fig_stacked_pct = px.bar(status_percentage, x='Broker', y='Percentage', color='Status',
                               title='Quantity Status Percentage by Broker (%)',
                               color_discrete_map={
                                   'Sold_Percentage': '#28a745',
                                   'Unsold_Percentage': '#dc3545',
                                   'Outsold_Percentage': '#ffc107'
                               },
                               barmode='stack')
        fig_stacked_pct.update_layout(xaxis_tickangle=-45, legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ))
        st.plotly_chart(fig_stacked_pct, use_container_width=True)
    
    st.markdown("---")
    
    # Broker Performance Metrics Table
    st.subheader(" Detailed Broker Quantity Performance")
    
    # Create display table with formatted quantities
    display_table = broker_performance[[
        'Broker', 'Total_Quantity', 'Sold_Quantity', 'Unsold_Quantity', 
        'Outsold_Quantity', 'Total_Sold_Side_Quantity',
        'Sold_Percentage', 'Unsold_Percentage', 'Outsold_Percentage', 'Sold_Side_Percentage',
        'Total_Lots', 'Sold_Lots', 'Avg_Price'
    ]].copy()
    
    # Format the quantities for display with null handling
    display_table_formatted = display_table.copy()
    quantity_columns = ['Total_Quantity', 'Sold_Quantity', 'Unsold_Quantity', 'Outsold_Quantity', 'Total_Sold_Side_Quantity']
    
    for col in quantity_columns:
        display_table_formatted[col] = display_table_formatted[col].fillna(0).apply(lambda x: f"{format_large_number(x)} kg")
    
    display_table_formatted['Sold_Percentage'] = display_table_formatted['Sold_Percentage'].fillna(0).apply(lambda x: f"{x:.1f}%")
    display_table_formatted['Unsold_Percentage'] = display_table_formatted['Unsold_Percentage'].fillna(0).apply(lambda x: f"{x:.1f}%")
    display_table_formatted['Outsold_Percentage'] = display_table_formatted['Outsold_Percentage'].fillna(0).apply(lambda x: f"{x:.1f}%")
    display_table_formatted['Sold_Side_Percentage'] = display_table_formatted['Sold_Side_Percentage'].fillna(0).apply(lambda x: f"{x:.1f}%")
    display_table_formatted['Total_Lots'] = display_table_formatted['Total_Lots'].fillna(0).apply(lambda x: f"{x:,}")
    display_table_formatted['Sold_Lots'] = display_table_formatted['Sold_Lots'].fillna(0).apply(lambda x: f"{x:,}")
    display_table_formatted['Avg_Price'] = display_table_formatted['Avg_Price'].fillna(0).apply(lambda x: f"LKR {x:,.2f}" if x > 0 else "N/A")
    
    # Style the table to highlight MPB
    def highlight_mpb(row):
        if row['Broker'] == 'MPB':
            return ['background-color: #FFF3CD'] * len(row)
        return [''] * len(row)
    
    st.dataframe(
        display_table_formatted.style.apply(highlight_mpb, axis=1),
        use_container_width=True
    )
    
    st.markdown("---")
    
    # Grade-wise Broker Performance
    st.subheader(" Grade-wise Quantity Performance by Broker")
    
    # Get top grades by total quantity
    top_grades = latest_df.groupby('Grade')['Total Weight'].sum().nlargest(10).index
    
    grade_broker_performance = latest_df[latest_df['Grade'].isin(top_grades)].groupby(['Grade', 'Broker']).apply(lambda x: pd.Series({
        'Total_Quantity': x["Total Weight"].sum(),
        'Sold_Quantity': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Unsold_Quantity': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Outsold_Quantity': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
    }), include_groups=False).reset_index()
    
    grade_broker_performance['Total_Sold_Side_Quantity'] = grade_broker_performance['Sold_Quantity'] + grade_broker_performance['Outsold_Quantity']
    grade_broker_performance['Is_MPB'] = grade_broker_performance['Broker'] == 'MPB'
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Heatmap of total quantity by grade and broker
        pivot_total = grade_broker_performance.pivot_table(
            index='Grade', columns='Broker', values='Total_Quantity', aggfunc='sum', fill_value=0
        )
        
        fig_heatmap_total = px.imshow(
            pivot_total,
            title='Total Quantity by Grade and Broker (kg) - Heatmap',
            labels=dict(x="Broker", y="Grade", color="Quantity (kg)"),
            color_continuous_scale="Blues",
            aspect="auto"
        )
        fig_heatmap_total.update_xaxes(side="bottom")
        st.plotly_chart(fig_heatmap_total, use_container_width=True)
    
    with col2:
        # Heatmap of sold side quantity
        pivot_sold_side = grade_broker_performance.pivot_table(
            index='Grade', columns='Broker', values='Total_Sold_Side_Quantity', aggfunc='sum', fill_value=0
        )
        
        fig_heatmap_sold_side = px.imshow(
            pivot_sold_side,
            title='Sold + Outsold Quantity by Grade and Broker (kg) - Heatmap',
            labels=dict(x="Broker", y="Grade", color="Quantity (kg)"),
            color_continuous_scale="Greens",
            aspect="auto"
        )
        fig_heatmap_sold_side.update_xaxes(side="bottom")
        st.plotly_chart(fig_heatmap_sold_side, use_container_width=True)
    
    # Broker performance in top grades
    st.subheader(" Broker Performance in Top 10 Grades")
    
    # Create a visualization for broker share in top grades
    top_grades_broker_share = grade_broker_performance.groupby(['Grade', 'Broker'])['Total_Quantity'].sum().unstack(fill_value=0)
    top_grades_broker_share_percentage = top_grades_broker_share.div(top_grades_broker_share.sum(axis=1), axis=0) * 100
    
    fig_grade_share = px.bar(top_grades_broker_share_percentage.reset_index().melt(id_vars=['Grade'], var_name='Broker', value_name='Percentage'),
                           x='Grade', y='Percentage', color='Broker',
                           title='Broker Market Share in Top 10 Grades (%)',
                           barmode='stack')
    fig_grade_share.update_layout(xaxis_tickangle=-45, legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ))
    st.plotly_chart(fig_grade_share, use_container_width=True)
    
    st.markdown("---")
    
    # Elevation-wise Broker Performance
    st.subheader(" Elevation-wise Quantity Performance by Broker")
    
    elevation_broker_performance = latest_df.groupby(['Sub Elevation', 'Broker']).apply(lambda x: pd.Series({
        'Total_Quantity': x["Total Weight"].sum(),
        'Sold_Quantity': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Unsold_Quantity': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Outsold_Quantity': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
    }), include_groups=False).reset_index()
    
    elevation_broker_performance['Total_Sold_Side_Quantity'] = elevation_broker_performance['Sold_Quantity'] + elevation_broker_performance['Outsold_Quantity']
    elevation_broker_performance['Is_MPB'] = elevation_broker_performance['Broker'] == 'MPB'
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Total quantity by elevation and broker
        fig_elev_total = px.bar(elevation_broker_performance, x='Sub Elevation', y='Total_Quantity', color='Broker',
                              title='Total Quantity by Elevation and Broker (kg)',
                              barmode='group')
        fig_elev_total.update_layout(xaxis_tickangle=-45, legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ))
        st.plotly_chart(fig_elev_total, use_container_width=True)
    
    with col2:
        # Sold side quantity by elevation and broker
        fig_elev_sold_side = px.bar(elevation_broker_performance, x='Sub Elevation', y='Total_Sold_Side_Quantity', color='Broker',
                                  title='Sold + Outsold Quantity by Elevation and Broker (kg)',
                                  barmode='group')
        fig_elev_sold_side.update_layout(xaxis_tickangle=-45, legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ))
        st.plotly_chart(fig_elev_sold_side, use_container_width=True)
    
    # Broker Performance Efficiency
    st.markdown("---")
    st.subheader(" Broker Efficiency Analysis - Quantity Focus")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Sold side percentage vs total quantity
        fig_efficiency = px.scatter(broker_performance, x='Total_Quantity', y='Sold_Side_Percentage',
                                  size='Sold_Quantity', color='Is_MPB',
                                  color_discrete_map={True: '#FF6B6B', False: '#4ECDC4'},
                                  title='Broker Efficiency: Total Quantity vs Sold Side %',
                                  hover_data=['Broker', 'Sold_Quantity', 'Unsold_Quantity'],
                                  size_max=40,
                                  labels={
                                      'Total_Quantity': 'Total Quantity (kg)',
                                      'Sold_Side_Percentage': 'Sold + Outsold %'
                                  })
        st.plotly_chart(fig_efficiency, use_container_width=True)
    
    with col2:
        # Lot success rate vs quantity
        broker_performance['Lot_Success_Rate'] = (broker_performance['Sold_Lots'] / broker_performance['Total_Lots'] * 100).fillna(0)
        
        fig_lot_efficiency = px.scatter(broker_performance, x='Total_Quantity', y='Lot_Success_Rate',
                                      size='Sold_Lots', color='Is_MPB',
                                      color_discrete_map={True: '#FF6B6B', False: '#45B7D1'},
                                      title='Broker Efficiency: Total Quantity vs Lot Success Rate %',
                                      hover_data=['Broker', 'Sold_Lots', 'Total_Lots'],
                                      size_max=40,
                                      labels={
                                          'Total_Quantity': 'Total Quantity (kg)',
                                          'Lot_Success_Rate': 'Lot Success Rate %'
                                      })
        st.plotly_chart(fig_lot_efficiency, use_container_width=True)

# BUYER INSIGHTS
with tabs[3]:
    st.header("Buyer Insights & Profiles")
    
    # Add MPB filter for buyer insights
    col1, col2 = st.columns([3, 1])
    with col2:
        buyer_mpb_only = st.checkbox(" Show MPB Buyers Only", value=False, key="buyer_mpb_filter")
    
    # Filter data based on MPB selection
    buyer_analysis_df = latest_df[latest_df["Status_Clean"] == "sold"]
    if buyer_mpb_only:
        buyer_analysis_df = buyer_analysis_df[buyer_analysis_df["Broker"] == "MPB"]
        st.info(f" Showing MPB buyers only - {len(buyer_analysis_df)} records")
    
    # Buyer grade-wise purchasing analysis
    st.subheader(" Buyer Purchased Profiles (Grade-wise)")
    
    buyer_grade_profile = buyer_analysis_df.groupby(["Buyer", "Grade"]).agg({
        "Total Weight": "sum",
        "Price": "mean",
        "Total Value": "sum",
        "Lot No": "count"
    }).reset_index()
    
    buyer_grade_profile.columns = ["Buyer", "Grade", "Quantity", "Avg_Price", "Total_Value", "Lots"]
    
    # Top buyers selector
    top_buyers_list = buyer_analysis_df.groupby("Buyer")["Total Value"].sum().nlargest(20).index.tolist()
    
    selected_buyer = st.selectbox("Select Buyer for Detailed Profile", top_buyers_list, key="buyer_profile_select")
    
    buyer_data = buyer_grade_profile[buyer_grade_profile["Buyer"] == selected_buyer]
    
    if not buyer_data.empty:
        st.markdown(f"###  {selected_buyer} - Purchase Profile")
        
        # Buyer summary metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Purchase Value", format_currency(buyer_data["Total_Value"].sum()))
        with col2:
            st.metric("Total Quantity", f"{format_large_number(buyer_data['Quantity'].sum())} kg")
        with col3:
            st.metric("Avg Purchase Price", f"LKR {buyer_data['Avg_Price'].mean():,.2f}")
        with col4:
            st.metric("Total Lots Purchased", f"{int(buyer_data['Lots'].sum()):,}")
        
        # Grade-wise breakdown
        col1, col2 = st.columns(2)
        
        with col1:
            # Grade distribution by value
            fig1 = px.pie(buyer_data.nlargest(10, 'Total_Value'), 
                         values='Total_Value', names='Grade',
                         title=f'{selected_buyer} - Purchase Value by Grade (Top 10)')
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            # Grade distribution by quantity
            fig2 = px.bar(buyer_data.nlargest(15, 'Quantity'), 
                         x='Grade', y='Quantity',
                         title=f'{selected_buyer} - Quantity by Grade (Top 15)',
                         color='Quantity',
                         color_continuous_scale='Blues')
            fig2.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig2, use_container_width=True)
        
        # Price paid by grade
        fig3 = px.bar(buyer_data.nlargest(15, 'Total_Value'), 
                     x='Grade', y='Avg_Price',
                     title=f'{selected_buyer} - Average Price Paid by Grade',
                     color='Avg_Price',
                     color_continuous_scale='Reds')
        fig3.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig3, use_container_width=True)
        
        # Detailed grade table with enhanced formatting
        st.markdown(f"###  Grade-wise Purchase Details - {selected_buyer}")
        display_buyer = buyer_data.sort_values('Total_Value', ascending=False).copy()
        
        # Format the display data
        display_buyer_formatted = display_buyer.copy()
        display_buyer_formatted['Quantity'] = display_buyer_formatted['Quantity'].apply(lambda x: f"{format_large_number(x)} kg")
        display_buyer_formatted['Avg_Price'] = display_buyer_formatted['Avg_Price'].apply(lambda x: f"LKR {x:,.2f}")
        display_buyer_formatted['Total_Value'] = display_buyer_formatted['Total_Value'].apply(lambda x: format_currency(x))
        display_buyer_formatted['Lots'] = display_buyer_formatted['Lots'].apply(lambda x: f"{x:,.0f}")
        
        st.dataframe(display_buyer_formatted, use_container_width=True)
    
    st.markdown("---")
    
    # Top buyers comparison
    st.subheader(" Top 20 Buyers Comparison")
    
    buyers = buyer_analysis_df.groupby("Buyer", as_index=False).agg({
        "Total Value": "sum",
        "Total Weight": "sum",
        "Price": "mean",
        "Lot No": "count"
    })
    buyers = buyers.sort_values("Total Value", ascending=False).head(20)
    buyers.columns = ["Buyer", "Total_Value", "Total_Weight", "Avg_Price", "Lots"]
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig = px.bar(buyers, x="Buyer", y="Total_Value", 
                    color="Total_Value",
                    title="Top 20 Buyers by Purchase Value",
                    color_continuous_scale="Blues")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        fig = px.bar(buyers, x="Buyer", y="Total_Weight",
                    color="Avg_Price",
                    title="Top 20 Buyers by Quantity",
                    color_continuous_scale="Greens")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    
    # Buyer loyalty analysis
    st.subheader(" Buyer Loyalty & Historical Participation")
    
    loyalty = data.groupby("Buyer").agg({
        "Sale_No": "nunique",
        "Total Value": "sum",
        "Total Weight": "sum"
    }).reset_index()
    loyalty.columns = ["Buyer", "Sales_Participated", "Total_Value", "Total_Weight"]
    loyalty = loyalty.sort_values("Sales_Participated", ascending=False).head(15)
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig = px.bar(loyalty, x="Buyer", y="Sales_Participated", 
                    title="Top 15 Loyal Buyers (by Sales Participated)",
                    color="Sales_Participated",
                    color_continuous_scale="Teal")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        fig = px.scatter(loyalty, x="Sales_Participated", y="Total_Value",
                        size="Total_Weight", hover_data=["Buyer"],
                        title="Buyer Loyalty vs Purchase Value",
                        color="Total_Value",
                        color_continuous_scale="Viridis")
        st.plotly_chart(fig, use_container_width=True)

# SELLING MARK ANALYSIS
with tabs[4]:
    st.header("Selling Mark Analysis")

    if "Selling Mark" not in data.columns:
        st.warning(" 'Selling Mark' column not found in your dataset.")
    else:
        has_trade_mark = ("Trade Mark" in data.columns and 
                         data["Trade Mark"].notna().any() and 
                         len(data[data["Trade Mark"].notna()]) > 0)
        
        if has_trade_mark:
            st.info(f" Found {data['Trade Mark'].notna().sum()} rows with Trade Mark data")
        else:
            st.info("ℹ No Trade Mark column found or it's empty - showing all Selling Marks")
        
        col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
        
        with col1:
            if has_trade_mark:
                valid_trade_marks = data[data["Trade Mark"].notna()]["Trade Mark"].unique()
                trade_marks_raw = sorted([str(tm).strip() for tm in valid_trade_marks if str(tm).strip()])
                
                # Create base trade marks dictionary
                base_trade_marks = {}
                for tm in trade_marks_raw:
                    base = get_base_trade_mark(tm)
                    if base not in base_trade_marks:
                        base_trade_marks[base] = []
                    base_trade_marks[base].append(tm)
                
                # Create display list with proper grouping
                trade_marks_display = []
                trade_marks_mapping = {}
                for base, sub_marks in sorted(base_trade_marks.items()):
                    display_name = f"{base}" if len(sub_marks) == 1 else f"{base} ({len(sub_marks)} variants)"
                    trade_marks_display.append(display_name)
                    trade_marks_mapping[display_name] = sub_marks
                
                st.write(f"**Trade Mark Groups Found:** {len(trade_marks_display)}")
                
                if not trade_marks_display:
                    st.warning(" No valid Trade Marks found. Showing all Selling Marks.")
                    has_trade_mark = False
                    available_marks = sorted([str(m).strip() for m in data["Selling Mark"].unique() 
                                            if str(m).strip() and str(m) not in ['nan', 'NaN', 'None', '']])
                    if not available_marks:
                        st.error(" No valid Selling Marks found in your dataset.")
                        st.stop()
                    selected_mark = st.selectbox(" Select Selling Mark", ["All Marks"] + available_marks, key="selling_mark_select")
                    analysis_mode = "single" if selected_mark != "All Marks" else "all"
                else:
                    selected_trade_mark_display = st.selectbox(
                        " Select Trade Mark Group", 
                        trade_marks_display, 
                        key="trade_mark_select"
                    )
                    
                    selected_trade_marks = trade_marks_mapping[selected_trade_mark_display]
                    trade_mark_data = data[data["Trade Mark"].isin(selected_trade_marks)]
                    available_marks = sorted([str(m).strip() for m in trade_mark_data["Selling Mark"].unique() 
                                            if str(m).strip() and str(m) not in ['nan', 'NaN', 'None', '']])
                    
                    if len(selected_trade_marks) > 1:
                        st.info(f" **{len(selected_trade_marks)} Trade Marks Selected:** {', '.join(selected_trade_marks)}")
                    
                    if not available_marks:
                        st.warning(f" No Selling Marks found under Trade Mark '{selected_trade_mark_display}'")
                        st.stop()
                    
                    # Add option to view all marks under this trade mark group
                    mark_options = [" All Marks (Combined View)"] + available_marks
                    selected_mark = st.selectbox(
                        f" Select Selling Mark ({len(available_marks)} available)", 
                        mark_options, 
                        key="selling_mark_select"
                    )
                    
                    analysis_mode = "combined" if selected_mark == " All Marks (Combined View)" else "single"
            else:
                available_marks = sorted([str(m).strip() for m in data["Selling Mark"].unique() 
                                        if str(m).strip() and str(m) not in ['nan', 'NaN', 'None', '']])
                if not available_marks:
                    st.error(" No valid Selling Marks found in your dataset.")
                    st.stop()
                
                st.write(f"** Selling Marks Found:** {len(available_marks)}")
                selected_mark = st.selectbox(
                    " Select Selling Mark", 
                    available_marks, 
                    key="selling_mark_select"
                )
                analysis_mode = "single"
        
        with col2:
            all_sales = sorted(data["Sale_No"].unique())
            sale_range = st.select_slider(
                " Sale Range",
                options=all_sales,
                value=(min(all_sales), max(all_sales))
            )
        
        with col3:
            show_mpb_only = st.checkbox(" MPB Only View", value=False, key="mpb_filter")
        
        with col4:
            st.markdown("###  Export")
            if st.button(" Download PDF", type="primary", help="Generate comprehensive PDF report"):
                st.info(" PDF report generation is available below. CSV exports have been removed.")

        # Filter data based on selection
        if has_trade_mark and 'selected_trade_marks' in locals():
            if analysis_mode == "combined":
                # Show all marks under selected trade mark group
                mark_df = data[
                    (data["Trade Mark"].isin(selected_trade_marks)) &
                    (data["Sale_No"] >= sale_range[0]) & 
                    (data["Sale_No"] <= sale_range[1])
                ]
                selected_marks_list = available_marks
                display_title = f"All Selling Marks under {selected_trade_mark_display}"
            else:
                # Show single selected mark
                mark_df = data[
                    (data["Selling Mark"] == selected_mark) & 
                    (data["Sale_No"] >= sale_range[0]) & 
                    (data["Sale_No"] <= sale_range[1])
                ]
                selected_marks_list = [selected_mark]
                display_title = selected_mark
        elif analysis_mode == "all":
            # Show all selling marks
            mark_df = data[
                (data["Sale_No"] >= sale_range[0]) & 
                (data["Sale_No"] <= sale_range[1])
            ]
            selected_marks_list = available_marks
            display_title = "All Selling Marks"
        else:
            # Single mark without trade mark
            mark_df = data[
                (data["Selling Mark"] == selected_mark) & 
                (data["Sale_No"] >= sale_range[0]) & 
                (data["Sale_No"] <= sale_range[1])
            ]
            selected_marks_list = [selected_mark]
            display_title = selected_mark
        
        if show_mpb_only:
            mark_df = mark_df[mark_df["Broker"] == "MPB"]

        if mark_df.empty:
            st.info(" No data available for the selected filters.")
        else:
            if has_trade_mark and 'selected_trade_marks' in locals():
                if analysis_mode == "combined":
                    st.markdown(f"""
                    <div style="background-color: #e3f2fd; padding: 15px; border-radius: 10px; border-left: 5px solid #2196F3;">
                        <h4 style="margin: 0; color: #1976D2;"> Trade Mark Group: <strong>{selected_trade_mark_display}</strong></h4>
                        <p style="margin: 5px 0 0 0; color: #424242;"> Analyzing <strong>{len(selected_marks_list)} Selling Marks</strong> combined</p>
                        <p style="margin: 5px 0 0 0; font-size: 0.9em; color: #666;">Marks: {', '.join(selected_marks_list[:10])}{' ...' if len(selected_marks_list) > 10 else ''}</p>
                    </div>
                    """, unsafe_allow_html=True)
                elif len(selected_trade_marks) > 1:
                    trade_marks_str = ", ".join(selected_trade_marks)
                    st.markdown(f"""
                    <div style="background-color: #e3f2fd; padding: 15px; border-radius: 10px; border-left: 5px solid #2196F3;">
                        <h4 style="margin: 0; color: #1976D2;"> Trade Mark(s): <strong>{trade_marks_str}</strong></h4>
                        <p style="margin: 5px 0 0 0; color: #424242;"> Selling Mark: <strong>{selected_mark}</strong></p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="background-color: #e8f5e9; padding: 15px; border-radius: 10px; border-left: 5px solid #4CAF50;">
                        <h4 style="margin: 0; color: #2E7D32;"> Selling Mark: <strong>{selected_mark}</strong></h4>
                        <p style="margin: 5px 0 0 0; color: #424242;">Trade Mark: {selected_trade_marks[0]}</p>
                    </div>
                    """, unsafe_allow_html=True)
            elif analysis_mode == "all":
                st.markdown(f"""
                <div style="background-color: #fff3e0; padding: 15px; border-radius: 10px; border-left: 5px solid #FF9800;">
                    <h4 style="margin: 0; color: #E65100;"> Analyzing All Selling Marks</h4>
                    <p style="margin: 5px 0 0 0; color: #424242;">Total: <strong>{len(selected_marks_list)} marks</strong></p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="background-color: #fff3e0; padding: 15px; border-radius: 10px; border-left: 5px solid #FF9800;">
                    <h4 style="margin: 0; color: #E65100;"> Selling Mark: <strong>{display_title}</strong></h4>
                </div>
                """, unsafe_allow_html=True)
            
            st.write("")
            
            # Show breakdown information if applicable
            if has_trade_mark and 'selected_trade_marks' in locals() and len(selected_trade_marks) > 1:
                with st.expander(f" Trade Mark Breakdown ({len(selected_trade_marks)} variants)"):
                    breakdown_data = []
                    for tm in selected_trade_marks:
                        tm_data = mark_df[mark_df["Trade Mark"] == tm]
                        if not tm_data.empty:
                            tm_marks = sorted([str(m).strip() for m in tm_data["Selling Mark"].unique() 
                                             if str(m).strip() and str(m) not in ['nan', 'NaN', 'None', '']])
                            sold_data = tm_data[tm_data["Status_Clean"] == "sold"]
                            
                            breakdown_data.append({
                                "Trade Mark": tm,
                                "Selling Marks": len(tm_marks),
                                "Total Volume": tm_data["Total Weight"].sum(),
                                "Sold Volume": sold_data["Total Weight"].sum(),
                                "Avg Price": sold_data["Price"].mean() if not sold_data.empty else 0,
                                "Total Value": (sold_data["Total Weight"] * sold_data["Price"]).sum()
                            })
                    
                    if breakdown_data:
                        breakdown_df = pd.DataFrame(breakdown_data)
                        st.dataframe(
                            breakdown_df.style.format({
                                "Total Volume": "{:,.2f}",
                                "Sold Volume": "{:,.2f}",
                                "Avg Price": "{:,.2f}",
                                "Total Value": "{:,.0f}"
                            }),
                            use_container_width=True,
                            hide_index=True
                        )
            
            st.markdown("---")
            st.subheader(" Quick Performance Comparison - Last 3 Sales")
            
            recent_sales = sorted(mark_df["Sale_No"].unique())[-3:]
            comparison_data = []
            
            for sale in recent_sales:
                sale_df = mark_df[mark_df["Sale_No"] == sale]
                sold_df = sale_df[sale_df["Status_Clean"] == "sold"]
                
                comparison_data.append({
                    "Sale": f"Sale {sale}",
                    "Catalogued (kg)": sale_df["Total Weight"].sum(),
                    "Sold (kg)": sold_df["Total Weight"].sum(),
                    "Unsold (kg)": sale_df[sale_df["Status_Clean"] == "unsold"]["Total Weight"].sum(),
                    "Avg Price": sold_df["Price"].mean() if not sold_df.empty else 0,
                    "Total Proceeds": (sold_df["Total Weight"] * sold_df["Price"]).sum(),
                    "Sold %": (sold_df["Total Weight"].sum() / sale_df["Total Weight"].sum() * 100) if sale_df["Total Weight"].sum() > 0 else 0,
                    "Brokers": sale_df["Broker"].nunique(),
                    "Top Broker": sale_df.groupby("Broker")["Total Value"].sum().idxmax() if not sale_df.empty else "N/A"
                })
            
            comparison_df = pd.DataFrame(comparison_data)
            
            cols = st.columns(len(recent_sales))
            for idx, sale in enumerate(recent_sales):
                sale_data = comparison_df[comparison_df["Sale"] == f"Sale {sale}"].iloc[0]
                with cols[idx]:
                    st.markdown(f"### Sale {sale}")
                    st.metric("Catalogued", f"{sale_data['Catalogued (kg)']:,.0f} kg")
                    st.metric("Sold %", f"{sale_data['Sold %']:.1f}%", 
                             delta=f"{sale_data['Sold %'] - comparison_df['Sold %'].mean():.1f}%" if idx > 0 else None)
                    st.metric("Avg Price", f"LKR {sale_data['Avg Price']:,.0f}")
                    st.metric("Proceeds", f"LKR {sale_data['Total Proceeds']:,.0f}")
                    st.info(f" Top: **{sale_data['Top Broker']}**")
            
            col1, col2 = st.columns(2)
            with col1:
                fig_trend = px.line(comparison_df, x="Sale", y="Avg Price", 
                                   markers=True, title="Price Trend - Last 3 Sales",
                                   labels={"Avg Price": "Average Price (LKR)"})
                fig_trend.update_traces(line_color='#007bff', line_width=3, marker_size=10)
                st.plotly_chart(fig_trend, use_container_width=True)
            
            with col2:
                fig_sold = px.bar(comparison_df, x="Sale", y="Sold %",
                                 title="Sell Percentage Trend - Last 3 Sales",
                                 color="Sold %",
                                 color_continuous_scale=['#dc3545', '#ffc107', '#28a745'])
                st.plotly_chart(fig_sold, use_container_width=True)
            
            st.markdown("---")
            st.subheader(f" Current Sale Summary - {display_title}")
            
            current_sale_df = mark_df[mark_df["Sale_No"] == latest_sale]
            
            if not current_sale_df.empty:
                if "broker_filter" not in st.session_state:
                    st.session_state.broker_filter = sorted(current_sale_df["Broker"].unique())
                
                st.markdown("###  Broker Participation")
                broker_summary = current_sale_df.groupby("Broker").agg({
                    "Total Weight": "sum",
                    "Price": "mean",
                    "Lot No": "count"
                }).reset_index()
                broker_summary["Total Value"] = broker_summary["Total Weight"] * broker_summary["Price"]
                broker_summary = broker_summary.sort_values("Total Value", ascending=False)
                
                broker_cols = st.columns(min(len(broker_summary), 4))
                for idx, (_, broker) in enumerate(broker_summary.iterrows()):
                    if idx < 4:
                        with broker_cols[idx]:
                            st.markdown(f"**{broker['Broker']}**")
                            st.metric("Lots", int(broker['Lot No']))
                            st.metric("Weight", f"{broker['Total Weight']:,.0f} kg")
                            st.metric("Value", f"LKR {broker['Total Value']:,.0f}")
                
                fig_broker = px.bar(broker_summary, x="Broker", y="Total Value",
                                   title="Broker Performance by Value",
                                   color="Broker",
                                   text="Total Value")
                fig_broker.update_traces(texttemplate='LKR %{text:,.0f}', textposition='outside')
                st.plotly_chart(fig_broker, use_container_width=True)
                
                st.markdown("---")
                
                catalogued_qty = current_sale_df["Total Weight"].sum()
                sold_df = current_sale_df[current_sale_df["Status_Clean"] == "sold"]
                unsold_df = current_sale_df[current_sale_df["Status_Clean"] == "unsold"]
                
                sold_qty = sold_df["Total Weight"].sum()
                unsold_qty = unsold_df["Total Weight"].sum()
                withdrawn_qty = catalogued_qty - sold_qty - unsold_qty if catalogued_qty > sold_qty + unsold_qty else 0
                
                avg_price = sold_df["Price"].mean() if not sold_df.empty else 0
                total_proceeds = (sold_df["Total Weight"] * sold_df["Price"]).sum()
                
                # Calculate Sold % as (Sold + Outsold) / Total
                outsold_df = current_sale_df[current_sale_df["Status_Clean"] == "outsold"]
                outsold_qty = outsold_df["Total Weight"].sum()
                total_sold_side_qty = sold_qty + outsold_qty
                sell_pct = calculate_sell_percentage(total_sold_side_qty, catalogued_qty)
                unsold_pct = calculate_sell_percentage(unsold_qty, catalogued_qty)
                
                st.markdown("###  Sale Metrics")
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.metric(" Catalogued Qty", f"{catalogued_qty:,.2f} kg")
                    st.metric(" Sold Qty", f"{sold_qty:,.2f} kg")
                
                with col2:
                    st.metric(" Average Price", f"LKR {avg_price:,.2f}")
                    st.metric(" Sold %", f"{sell_pct:.2f}%")
                
                with col3:
                    st.metric(" Total Proceeds", format_currency(total_proceeds))
                    st.metric(" Unsold %", f"{unsold_pct:.2f}%")
                
                with col4:
                    st.metric(" Unsold Qty", f"{unsold_qty:,.2f} kg")
                    st.metric(" Withdrawn Qty", f"{withdrawn_qty:,.2f} kg")
                
                with col5:
                    num_lots = len(current_sale_df)
                    num_sold = len(sold_df)
                    st.metric(" Total Lots", num_lots)
                    st.metric(" Lots Sold", num_sold)
                
                st.markdown("###  Sale Performance Breakdown")
                col1, col2 = st.columns(2)
                
                with col1:
                    qty_dist = pd.DataFrame({
                        'Status': ['Sold', 'Unsold', 'Withdrawn'],
                        'Quantity': [sold_qty, unsold_qty, withdrawn_qty]
                    })
                    qty_dist = qty_dist[qty_dist['Quantity'] > 0]
                    
                    fig_qty = px.pie(
                        qty_dist, values='Quantity', names='Status',
                        title=f'Quantity Distribution - Sale {latest_sale}',
                        color='Status',
                        color_discrete_map={'Sold': '#28a745', 'Unsold': '#dc3545', 'Withdrawn': '#ffc107'}
                    )
                    fig_qty.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(fig_qty, use_container_width=True)
                
                with col2:
                    metrics_df = pd.DataFrame({
                        'Metric': ['Catalogued', 'Sold', 'Unsold'],
                        'Quantity (kg)': [catalogued_qty, sold_qty, unsold_qty]
                    })
                    
                    fig_bar = px.bar(
                        metrics_df, x='Metric', y='Quantity (kg)',
                        title='Quantity Comparison',
                        color='Metric',
                        color_discrete_map={'Catalogued': '#007bff', 'Sold': '#28a745', 'Unsold': '#dc3545'}
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
            
            st.markdown("---")
            st.subheader(" Detailed Lot Information with Broker Details")
            
            if not current_sale_df.empty:
                st.markdown("###  Quick Filters")
                filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 2])
                
                with filter_col1:
                    if st.button(" Show All Brokers", use_container_width=True):
                        st.session_state.broker_filter = sorted(current_sale_df["Broker"].unique())
                        st.rerun()
                
                with filter_col2:
                    if st.button(" Show MPB Only", use_container_width=True):
                        st.session_state.broker_filter = ["MPB"]
                        st.rerun()
                
                with filter_col3:
                    available_brokers = sorted(current_sale_df["Broker"].unique())
                    valid_filter = [b for b in st.session_state.broker_filter if b in available_brokers]
                    if not valid_filter:
                        valid_filter = available_brokers
                    
                    broker_filter = st.multiselect(
                        "Or Select Specific Brokers:",
                        options=available_brokers,
                        default=valid_filter,
                        key="broker_multiselect"
                    )
                    
                    if broker_filter:
                        st.session_state.broker_filter = broker_filter
                    else:
                        st.session_state.broker_filter = available_brokers
                        broker_filter = available_brokers
                
                filtered_lots = current_sale_df[current_sale_df["Broker"].isin(broker_filter)]
                
                base_columns = ["Lot No", "Broker", "Grade", "Invoice No", "Total Weight", "Buyer", "Price", "Status"]
                
                if "Valuation price" in filtered_lots.columns:
                    base_columns.insert(7, "Valuation price")
                if "Asking Price" in filtered_lots.columns:
                    base_columns.insert(8, "Asking Price")
                
                lot_details = filtered_lots[base_columns].copy()
                lot_details["Proceeds"] = lot_details["Total Weight"] * lot_details["Price"]
                
                if "Valuation price" in lot_details.columns and "Asking Price" in lot_details.columns:
                    lot_details["Val vs Sale"] = lot_details["Price"] - lot_details["Valuation price"]
                    lot_details["Ask vs Sale"] = lot_details["Price"] - lot_details["Asking Price"]
                
                lot_details = lot_details.sort_values(["Broker", "Lot No"])
                
                st.markdown("###  Summary by Broker")
                summary_data = []
                for broker in sorted(filtered_lots["Broker"].unique()):
                    broker_data = filtered_lots[filtered_lots["Broker"] == broker]
                    sold_data = broker_data[broker_data["Status_Clean"] == "sold"]
                    unsold_data = broker_data[broker_data["Status_Clean"] == "unsold"]
                    
                    summary_data.append({
                        "Broker": broker,
                        "Total Lots": len(broker_data),
                        "Sold": len(sold_data),
                        "Unsold": len(unsold_data),
                        "Sold %": f"{(len(sold_data)/len(broker_data)*100):.1f}%" if len(broker_data) > 0 else "0%",
                        "Total Weight": f"{broker_data['Total Weight'].sum():,.2f} kg",
                        "Avg Price": f"LKR {sold_data['Price'].mean():,.2f}" if not sold_data.empty else "N/A"
                    })
                
                summary_df = pd.DataFrame(summary_data)
                st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
                st.markdown("###  Lot Details")
                st.caption(" **Broker column shows which broker sold each lot**")
                
                lot_details_display = lot_details.copy()
                lot_details_display["Total Weight"] = lot_details_display["Total Weight"].apply(lambda x: f"{x:,.2f}")
                lot_details_display["Price"] = lot_details_display["Price"].apply(lambda x: f"{x:,.2f}")
                lot_details_display["Proceeds"] = lot_details_display["Proceeds"].apply(lambda x: f"{x:,.2f}")
                
                if "Valuation price" in lot_details_display.columns:
                    lot_details_display["Valuation price"] = lot_details_display["Valuation price"].apply(lambda x: f"{x:,.2f}")
                if "Asking Price" in lot_details_display.columns:
                    lot_details_display["Asking Price"] = lot_details_display["Asking Price"].apply(lambda x: f"{x:,.2f}")
                if "Val vs Sale" in lot_details_display.columns:
                    lot_details_display["Val vs Sale"] = lot_details_display["Val vs Sale"].apply(lambda x: f"{x:+,.2f}")
                if "Ask vs Sale" in lot_details_display.columns:
                    lot_details_display["Ask vs Sale"] = lot_details_display["Ask vs Sale"].apply(lambda x: f"{x:+,.2f}")
                
                st.dataframe(lot_details_display, use_container_width=True, hide_index=True)
                
                if "Valuation price" in filtered_lots.columns and "Asking Price" in filtered_lots.columns:
                    st.markdown("###  Price Analysis")
                    
                    sold_lots = filtered_lots[filtered_lots["Status_Clean"] == "sold"]
                    if not sold_lots.empty:
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            avg_val = sold_lots["Valuation price"].mean()
                            avg_sale = sold_lots["Price"].mean()
                            diff = avg_sale - avg_val
                            st.metric("Avg Valuation Price", f"LKR {avg_val:,.2f}")
                            st.metric("Valuation vs Sale", f"LKR {diff:+,.2f}", 
                                     delta=f"{(diff/avg_val*100):.1f}%" if avg_val > 0 else "N/A")
                        
                        with col2:
                            avg_ask = sold_lots["Asking Price"].mean()
                            diff_ask = avg_sale - avg_ask
                            st.metric("Avg Asking Price", f"LKR {avg_ask:,.2f}")
                            st.metric("Asking vs Sale", f"LKR {diff_ask:+,.2f}",
                                     delta=f"{(diff_ask/avg_ask*100):.1f}%" if avg_ask > 0 else "N/A")
                        
                        with col3:
                            st.metric("Avg Sale Price", f"LKR {avg_sale:,.2f}")
                            lots_above_val = len(sold_lots[sold_lots["Price"] > sold_lots["Valuation price"]])
                            st.metric("Lots Above Valuation", f"{lots_above_val}/{len(sold_lots)}")
                        
                        price_comp = sold_lots.groupby("Broker").agg({
                            "Valuation price": "mean",
                            "Asking Price": "mean",
                            "Price": "mean"
                        }).reset_index()
                        
                        fig_price_comp = go.Figure()
                        fig_price_comp.add_trace(go.Bar(name="Valuation", x=price_comp["Broker"], y=price_comp["Valuation price"]))
                        fig_price_comp.add_trace(go.Bar(name="Asking", x=price_comp["Broker"], y=price_comp["Asking Price"]))
                        fig_price_comp.add_trace(go.Bar(name="Sale", x=price_comp["Broker"], y=price_comp["Price"]))
                        fig_price_comp.update_layout(title="Price Comparison by Broker", barmode="group")
                        st.plotly_chart(fig_price_comp, use_container_width=True)
                
                # CSV export options removed - PDF export is available in the report section
            
            st.markdown("---")
            st.subheader(" Grade-Wise Comparative Analysis")
            
            if analysis_mode in ["combined", "all"]:
                # Multi-mark comparison like Excel report
                st.info(f" Comparing **{len(selected_marks_list)}** selling marks across grades")
                
                # Add filters for comparative analysis
                col1, col2, col3 = st.columns([2, 1, 1])
                with col1:
                    # Multi-select for specific marks if needed
                    if len(selected_marks_list) > 5:
                        marks_for_analysis = st.multiselect(
                            " Select specific marks to compare (or leave empty for all)",
                            options=selected_marks_list,
                            default=[],
                            key="marks_filter"
                        )
                        if marks_for_analysis:
                            selected_marks_list = marks_for_analysis
                            st.success(f" Analyzing {len(marks_for_analysis)} selected marks")
                
                with col2:
                    min_qty_filter = st.number_input(
                        "Min Quantity (kg)",
                        min_value=0.0,
                        value=0.0,
                        step=10.0,
                        key="min_qty"
                    )
                
                with col3:
                    top_n_grades = st.selectbox(
                        "Show Top N Grades",
                        options=[10, 15, 20, 25, "All"],
                        index=0,
                        key="top_n"
                    )
                
                # Create comprehensive grade analysis
                grade_comparison = []
                
                for mark in selected_marks_list:
                    mark_data = mark_df[mark_df["Selling Mark"] == mark]
                    sold_data = mark_data[mark_data["Status_Clean"] == "sold"]
                    
                    if not sold_data.empty:
                        for grade in sold_data["Grade"].unique():
                            grade_data = sold_data[sold_data["Grade"] == grade]
                            total_qty = grade_data["Total Weight"].sum()
                            
                            # Apply quantity filter
                            if total_qty >= min_qty_filter:
                                grade_comparison.append({
                                    "Selling Mark": mark,
                                    "Grade": grade,
                                    "QTY (kg)": total_qty,
                                    "% of Total": (total_qty / sold_data["Total Weight"].sum() * 100) if sold_data["Total Weight"].sum() > 0 else 0,
                                    "AVG Price": grade_data["Price"].mean(),
                                    "Total Value": (grade_data["Total Weight"] * grade_data["Price"]).sum(),
                                    "Lots": len(grade_data),
                                    "Min Price": grade_data["Price"].min(),
                                    "Max Price": grade_data["Price"].max(),
                                    "Std Dev": grade_data["Price"].std()
                                })
                
                if grade_comparison:
                    grade_comp_df = pd.DataFrame(grade_comparison)
                    
                    # Apply top N filter
                    if top_n_grades != "All":
                        top_grades = grade_comp_df.groupby("Grade")["QTY (kg)"].sum().sort_values(ascending=False).head(top_n_grades).index
                        grade_comp_df = grade_comp_df[grade_comp_df["Grade"].isin(top_grades)]
                    
                    if grade_comp_df.empty:
                        st.warning("No data matches the selected filters. Try adjusting the minimum quantity.")
                    else:
                        # Summary cards for combined analysis
                        col1, col2, col3, col4, col5 = st.columns(5)
                        with col1:
                            st.metric("Total Marks", grade_comp_df["Selling Mark"].nunique())
                            st.metric("Total Grades", grade_comp_df["Grade"].nunique())
                        with col2:
                            st.metric("Total Volume", f"{grade_comp_df['QTY (kg)'].sum():,.0f} kg")
                            st.metric("Avg Price", f"LKR {grade_comp_df['AVG Price'].mean():,.0f}")
                        with col3:
                            st.metric("Total Value", format_currency(grade_comp_df['Total Value'].sum()))
                            st.metric("Total Lots", f"{grade_comp_df['Lots'].sum():,.0f}")
                        with col4:
                            st.metric("Highest Price", f"LKR {grade_comp_df['Max Price'].max():,.0f}")
                            st.metric("Lowest Price", f"LKR {grade_comp_df['Min Price'].min():,.0f}")
                        with col5:
                            price_range = grade_comp_df['Max Price'].max() - grade_comp_df['Min Price'].min()
                            st.metric("Price Range", f"LKR {price_range:,.0f}")
                            st.metric("Avg Std Dev", f"LKR {grade_comp_df['Std Dev'].mean():,.0f}")
                        
                        # Pivot table for grade comparison (like Excel)
                        st.markdown("###  Grade Comparison Matrix (Excel-Style)")
                        
                        # Create pivot for QTY with error handling
                        try:
                            pivot_qty = grade_comp_df.pivot_table(
                                index='Grade',
                                columns='Selling Mark',
                                values='QTY (kg)',
                                aggfunc='sum',
                                fill_value=0
                            )
                            
                            # Create pivot for AVG Price
                            pivot_price = grade_comp_df.pivot_table(
                                index='Grade',
                                columns='Selling Mark',
                                values='AVG Price',
                                aggfunc='mean',
                                fill_value=0
                            )
                            
                            # Create pivot for Percentage
                            pivot_pct = grade_comp_df.pivot_table(
                                index='Grade',
                                columns='Selling Mark',
                                values='% of Total',
                                aggfunc='sum',
                                fill_value=0
                            )
                            
                            # Create pivot for Total Value
                            pivot_value = grade_comp_df.pivot_table(
                                index='Grade',
                                columns='Selling Mark',
                                values='Total Value',
                                aggfunc='sum',
                                fill_value=0
                            )
                        except Exception as e:
                            st.error(f"Error creating pivot tables: {str(e)}")
                            st.stop()
                        
                        tab1, tab2, tab3, tab4 = st.tabs([" Quantity (kg)", " Avg Price (LKR)", " Percentage (%)", " Total Value"])
                        
                        with tab1:
                            # Add totals row and column
                            pivot_qty_display = pivot_qty.copy()
                            pivot_qty_display['TOTAL'] = pivot_qty_display.sum(axis=1)
                            pivot_qty_display.loc['TOTAL'] = pivot_qty_display.sum()
                            
                            st.dataframe(
                                pivot_qty_display.style.format("{:,.2f}")
                                .background_gradient(cmap='Greens', subset=pd.IndexSlice[pivot_qty_display.index[:-1], pivot_qty_display.columns[:-1]])
                                .highlight_max(axis=1, subset=pivot_qty_display.columns[:-1], color='lightgreen')
                                .set_properties(**{'text-align': 'right', 'font-weight': 'bold'}, subset=['TOTAL'])
                                .set_properties(**{'text-align': 'right', 'font-weight': 'bold'}, subset=pd.IndexSlice['TOTAL', :]),
                                use_container_width=True
                            )
                            
                            # Visualization
                            fig_qty_comp = px.bar(
                                grade_comp_df,
                                x="Grade",
                                y="QTY (kg)",
                                color="Selling Mark",
                                title="Quantity Distribution by Grade and Selling Mark",
                                barmode="group"
                            )
                            fig_qty_comp.update_layout(xaxis={'categoryorder': 'total descending'})
                            st.plotly_chart(fig_qty_comp, use_container_width=True)
                            
                            # Stacked bar chart
                            fig_qty_stacked = px.bar(
                                grade_comp_df,
                                x="Grade",
                                y="QTY (kg)",
                                color="Selling Mark",
                                title="Quantity Distribution (Stacked)",
                                barmode="stack"
                            )
                            fig_qty_stacked.update_layout(xaxis={'categoryorder': 'total descending'})
                            st.plotly_chart(fig_qty_stacked, use_container_width=True)
                        
                        with tab2:
                            # Add average row and column
                            pivot_price_display = pivot_price.copy()
                            pivot_price_display['AVG'] = pivot_price_display.mean(axis=1)
                            pivot_price_display.loc['AVERAGE'] = pivot_price_display.mean()
                            
                            st.dataframe(
                                pivot_price_display.style.format("{:,.2f}")
                                .background_gradient(cmap='RdYlGn', subset=pd.IndexSlice[pivot_price_display.index[:-1], pivot_price_display.columns[:-1]])
                                .highlight_max(axis=1, subset=pivot_price_display.columns[:-1], color='lightgreen')
                                .highlight_min(axis=1, subset=pivot_price_display.columns[:-1], color='lightcoral')
                                .set_properties(**{'text-align': 'right', 'font-weight': 'bold'}, subset=['AVG'])
                                .set_properties(**{'text-align': 'right', 'font-weight': 'bold'}, subset=pd.IndexSlice['AVERAGE', :]),
                                use_container_width=True
                            )
                            
                            # Multi-line chart
                            fig_price_comp = px.line(
                                grade_comp_df,
                                x="Grade",
                                y="AVG Price",
                                color="Selling Mark",
                                title="Average Price Comparison by Grade",
                                markers=True,
                                line_shape='spline'
                            )
                            fig_price_comp.update_layout(xaxis={'categoryorder': 'category ascending'})
                            st.plotly_chart(fig_price_comp, use_container_width=True)
                            
                            # Box plot for price distribution
                            sold_df_for_box = mark_df[
                                (mark_df["Status_Clean"] == "sold") &
                                (mark_df["Selling Mark"].isin(selected_marks_list))
                            ]
                            
                            if not sold_df_for_box.empty and len(sold_df_for_box["Selling Mark"].unique()) > 1:
                                fig_box = px.box(
                                    sold_df_for_box,
                                    x="Selling Mark",
                                    y="Price",
                                    title="Price Distribution by Selling Mark (Box Plot)",
                                    color="Selling Mark"
                                )
                                st.plotly_chart(fig_box, use_container_width=True)
                        
                        with tab3:
                            pivot_pct_display = pivot_pct.copy()
                            pivot_pct_display['AVG'] = pivot_pct_display.mean(axis=1)
                            
                            st.dataframe(
                                pivot_pct_display.style.format("{:.2f}")
                                .background_gradient(cmap='Blues', subset=pd.IndexSlice[:, pivot_pct_display.columns[:-1]])
                                .set_properties(**{'text-align': 'right'}),
                                use_container_width=True
                            )
                            
                            # Heatmap visualization
                            fig_heatmap = px.imshow(
                                pivot_pct,
                                title="Grade Distribution Heatmap (% of Total)",
                                labels=dict(x="Selling Mark", y="Grade", color="Percentage"),
                                color_continuous_scale="Blues",
                                aspect="auto",
                                text_auto='.2f'
                            )
                            fig_heatmap.update_xaxes(side="bottom")
                            st.plotly_chart(fig_heatmap, use_container_width=True)
                        
                        with tab4:
                            pivot_value_display = pivot_value.copy()
                            pivot_value_display['TOTAL'] = pivot_value_display.sum(axis=1)
                            pivot_value_display.loc['TOTAL'] = pivot_value_display.sum()
                            
                            st.dataframe(
                                pivot_value_display.style.format("{:,.0f}")
                                .background_gradient(cmap='Purples', subset=pd.IndexSlice[pivot_value_display.index[:-1], pivot_value_display.columns[:-1]])
                                .set_properties(**{'text-align': 'right', 'font-weight': 'bold'}, subset=['TOTAL'])
                                .set_properties(**{'text-align': 'right', 'font-weight': 'bold'}, subset=pd.IndexSlice['TOTAL', :]),
                                use_container_width=True
                            )
                            
                            # Treemap for value distribution
                            treemap_data = grade_comp_df[grade_comp_df['Total Value'] > 0].copy()
                            if not treemap_data.empty and len(treemap_data) > 0:
                                fig_treemap = px.treemap(
                                    treemap_data,
                                    path=['Selling Mark', 'Grade'],
                                    values='Total Value',
                                    title='Value Distribution Treemap',
                                    color='AVG Price',
                                    color_continuous_scale='RdYlGn'
                                )
                                st.plotly_chart(fig_treemap, use_container_width=True)
                            else:
                                st.info("No data available for treemap visualization")
                        
                        # Top performing grades
                        st.markdown("###  Top Performing Grades Analysis")
                        
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            top_by_qty = grade_comp_df.groupby("Grade")["QTY (kg)"].sum().sort_values(ascending=False).head(10)
                            fig_top_qty = px.bar(
                                x=top_by_qty.index,
                                y=top_by_qty.values,
                                title="Top 10 Grades by Quantity",
                                labels={"x": "Grade", "y": "Quantity (kg)"},
                                color=top_by_qty.values,
                                color_continuous_scale="Greens"
                            )
                            fig_top_qty.update_layout(showlegend=False)
                            st.plotly_chart(fig_top_qty, use_container_width=True)
                        
                        with col2:
                            top_by_price = grade_comp_df.groupby("Grade")["AVG Price"].mean().sort_values(ascending=False).head(10)
                            fig_top_price = px.bar(
                                x=top_by_price.index,
                                y=top_by_price.values,
                                title="Top 10 Grades by Average Price",
                                labels={"x": "Grade", "y": "Average Price (LKR)"},
                                color=top_by_price.values,
                                color_continuous_scale="Reds"
                            )
                            fig_top_price.update_layout(showlegend=False)
                            st.plotly_chart(fig_top_price, use_container_width=True)
                        
                        with col3:
                            top_by_value = grade_comp_df.groupby("Grade")["Total Value"].sum().sort_values(ascending=False).head(10)
                            fig_top_value = px.bar(
                                x=top_by_value.index,
                                y=top_by_value.values,
                                title="Top 10 Grades by Total Value",
                                labels={"x": "Grade", "y": "Total Value (LKR)"},
                                color=top_by_value.values,
                                color_continuous_scale="Blues"
                            )
                            fig_top_value.update_layout(showlegend=False)
                            st.plotly_chart(fig_top_value, use_container_width=True)
                        
                        # Price vs Quantity scatter
                        st.markdown("###  Price vs Quantity Analysis")
                        fig_scatter = px.scatter(
                            grade_comp_df,
                            x="QTY (kg)",
                            y="AVG Price",
                            color="Selling Mark",
                            size="Total Value",
                            hover_data=["Grade", "Lots"],
                            title="Price vs Quantity Scatter (Bubble size = Total Value)",
                            labels={"QTY (kg)": "Quantity (kg)", "AVG Price": "Average Price (LKR)"}
                        )
                        fig_scatter.update_traces(marker=dict(line=dict(width=1, color='DarkSlateGrey')))
                        st.plotly_chart(fig_scatter, use_container_width=True)
                        
                        # Selling Mark performance summary
                        st.markdown("###  Selling Mark Performance Summary")
                        mark_summary = grade_comp_df.groupby("Selling Mark").agg({
                            "QTY (kg)": "sum",
                            "AVG Price": "mean",
                            "Total Value": "sum",
                            "Lots": "sum",
                            "Grade": "nunique",
                            "Min Price": "min",
                            "Max Price": "max"
                        }).reset_index()
                        mark_summary.columns = ["Selling Mark", "Total QTY (kg)", "Avg Price (LKR)", "Total Value (LKR)", "Total Lots", "Grades Count", "Min Price", "Max Price"]
                        mark_summary["Price Range"] = mark_summary["Max Price"] - mark_summary["Min Price"]
                        mark_summary["Value per Lot"] = mark_summary["Total Value (LKR)"] / mark_summary["Total Lots"]
                        mark_summary = mark_summary.sort_values("Total Value (LKR)", ascending=False)
                        
                        # Add ranking
                        mark_summary.insert(0, "Rank", range(1, len(mark_summary) + 1))
                        
                        st.dataframe(
                            mark_summary.style.format({
                                "Total QTY (kg)": "{:,.2f}",
                                "Avg Price (LKR)": "{:,.2f}",
                                "Total Value (LKR)": "{:,.0f}",
                                "Total Lots": "{:.0f}",
                                "Grades Count": "{:.0f}",
                                "Min Price": "{:,.0f}",
                                "Max Price": "{:,.0f}",
                                "Price Range": "{:,.0f}",
                                "Value per Lot": "{:,.0f}"
                            })
                            .background_gradient(subset=["Total Value (LKR)"], cmap='Greens')
                            .background_gradient(subset=["Avg Price (LKR)"], cmap='RdYlGn')
                            .set_properties(**{'text-align': 'right'}),
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        # Performance comparison charts
                        col1, col2 = st.columns(2)
                        with col1:
                            fig_mark_value = px.bar(
                                mark_summary,
                                x="Selling Mark",
                                y="Total Value (LKR)",
                                title="Total Value by Selling Mark",
                                color="Total Value (LKR)",
                                color_continuous_scale="Viridis"
                            )
                            fig_mark_value.update_layout(xaxis_tickangle=-45, showlegend=False)
                            st.plotly_chart(fig_mark_value, use_container_width=True)
                        
                        with col2:
                            fig_mark_price = px.scatter(
                                mark_summary,
                                x="Total QTY (kg)",
                                y="Avg Price (LKR)",
                                size="Total Value (LKR)",
                                color="Selling Mark",
                                title="Avg Price vs Quantity by Selling Mark",
                                hover_data=["Grades Count", "Total Lots"]
                            )
                            st.plotly_chart(fig_mark_price, use_container_width=True)
                        
                        # Grade diversity analysis
                        st.markdown("###  Grade Diversity Analysis")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            fig_diversity = px.bar(
                                mark_summary.sort_values("Grades Count", ascending=False),
                                x="Selling Mark",
                                y="Grades Count",
                                title="Number of Grades per Selling Mark",
                                color="Grades Count",
                                color_continuous_scale="Rainbow"
                            )
                            fig_diversity.update_layout(xaxis_tickangle=-45, showlegend=False)
                            st.plotly_chart(fig_diversity, use_container_width=True)
                        
                        with col2:
                            # Grade concentration - show which marks dominate which grades
                            grade_leaders = grade_comp_df.loc[grade_comp_df.groupby("Grade")["Total Value"].idxmax()]
                            grade_leader_count = grade_leaders.groupby("Selling Mark").size().reset_index(name="Grades Led")
                            
                            fig_leaders = px.pie(
                                grade_leader_count,
                                values="Grades Led",
                                names="Selling Mark",
                                title="Grade Leadership (Highest Value per Grade)",
                                hole=0.4
                            )
                            st.plotly_chart(fig_leaders, use_container_width=True)
                        
                        # CSV export options removed - PDF export is available in the report section
                else:
                    st.warning(" No grade comparison data available for the selected filters. Try adjusting:")
                    st.info("• Decrease minimum quantity filter\n• Select more selling marks\n• Expand the sale range")
            else:
                st.info("ℹ Grade-wise comparative analysis is available when viewing multiple selling marks together. Select 'All Marks (Combined View)' to see detailed comparisons.")
            
            st.markdown("---")
            st.subheader(" Historical Performance Analysis")
            
            historical = mark_df.groupby("Sale_No").apply(lambda x: pd.Series({
                'Catalogued': x["Total Weight"].sum(),
                'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
                'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
                'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean(),
                'Total_Proceeds': (x[x["Status_Clean"] == "sold"]["Total Weight"] * 
                                  x[x["Status_Clean"] == "sold"]["Price"]).sum(),
                'Num_Lots': len(x)
            }), include_groups=False).reset_index()
            
            # Calculate Outsold for historical data
            historical_outsold = mark_df.groupby("Sale_No").apply(lambda x: pd.Series({
                'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum()
            }), include_groups=False).reset_index()
            
            # Merge outsold data into historical
            if not historical_outsold.empty and 'Outsold' in historical_outsold.columns:
                historical = historical.merge(historical_outsold[['Sale_No', 'Outsold']], on='Sale_No', how='left')
                historical['Outsold'] = historical['Outsold'].fillna(0)
            else:
                historical['Outsold'] = 0
            
            # Calculate Sold % as (Sold + Outsold) / Total
            historical['Total_Sold_Side'] = historical['Sold'] + historical['Outsold']
            historical['Sell_Pct'] = (historical['Total_Sold_Side'] / historical['Catalogued'] * 100).fillna(0)
            historical['Unsold_Pct'] = (historical['Unsold'] / historical['Catalogued'] * 100).fillna(0)
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig_price_trend = px.line(
                    historical, x='Sale_No', y='Avg_Price',
                    title='Average Price Trend',
                    markers=True,
                    labels={'Avg_Price': 'Average Price (LKR)', 'Sale_No': 'Sale Number'}
                )
                fig_price_trend.update_traces(line_color='#007bff', line_width=3)
                st.plotly_chart(fig_price_trend, use_container_width=True)
            
            with col2:
                fig_qty_trend = go.Figure()
                fig_qty_trend.add_trace(go.Scatter(
                    x=historical['Sale_No'], y=historical['Sold'],
                    mode='lines+markers', name='Sold',
                    line=dict(color='#28a745', width=3)
                ))
                fig_qty_trend.add_trace(go.Scatter(
                    x=historical['Sale_No'], y=historical['Unsold'],
                    mode='lines+markers', name='Unsold',
                    line=dict(color='#dc3545', width=3)
                ))
                fig_qty_trend.update_layout(
                    title='Sold vs Unsold Quantity Trend',
                    xaxis_title='Sale Number',
                    yaxis_title='Quantity (kg)',
                    hovermode='x unified'
                )
                st.plotly_chart(fig_qty_trend, use_container_width=True)
            
            fig_sell_pct = px.bar(
                historical, x='Sale_No', y='Sell_Pct',
                title='Sell Percentage Trend',
                labels={'Sell_Pct': 'Sell %', 'Sale_No': 'Sale Number'},
                color='Sell_Pct',
                color_continuous_scale=['#dc3545', '#ffc107', '#28a745']
            )
            fig_sell_pct.update_layout(showlegend=False)
            st.plotly_chart(fig_sell_pct, use_container_width=True)
            
            st.markdown("---")
            st.subheader(" Buyer Analysis with Broker Breakdown")
            
            buyer_analysis = mark_df[mark_df["Status_Clean"] == "sold"].groupby(["Buyer", "Broker"]).agg({
                "Total Weight": "sum",
                "Price": "mean",
                "Sale_No": "nunique",
                "Lot No": "count"
            }).reset_index()
            
            buyer_analysis.columns = ["Buyer", "Broker", "Total_Qty", "Avg_Price", "Num_Sales", "Num_Lots"]
            buyer_analysis["Total_Value"] = buyer_analysis["Total_Qty"] * buyer_analysis["Avg_Price"]
            
            top_buyers_overall = buyer_analysis.groupby("Buyer").agg({
                "Total_Value": "sum",
                "Total_Qty": "sum"
            }).sort_values("Total_Value", ascending=False).head(15).reset_index()
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig_buyers = px.bar(
                    top_buyers_overall.head(10), x="Total_Value", y="Buyer",
                    orientation='h',
                    title='Top 10 Buyers by Total Value',
                    labels={'Total_Value': 'Total Value (LKR)'},
                    color='Total_Value',
                    color_continuous_scale='Blues'
                )
                st.plotly_chart(fig_buyers, use_container_width=True)
            
            with col2:
                buyer_broker_top = buyer_analysis.nlargest(15, "Total_Value")
                fig_buyer_broker = px.bar(
                    buyer_broker_top, x="Buyer", y="Total_Value",
                    color="Broker",
                    title='Top Buyers with Broker Breakdown',
                    labels={'Total_Value': 'Total Value (LKR)'}
                )
                fig_buyer_broker.update_xaxes(tickangle=-45)
                st.plotly_chart(fig_buyer_broker, use_container_width=True)
            
            if "MPB" in buyer_analysis["Broker"].values:
                st.markdown("###  MPB Performance vs Other Brokers")
                
                mpb_buyers = buyer_analysis[buyer_analysis["Broker"] == "MPB"].groupby("Buyer")["Total_Value"].sum().reset_index()
                other_buyers = buyer_analysis[buyer_analysis["Broker"] != "MPB"].groupby("Buyer")["Total_Value"].sum().reset_index()
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("MPB Total Buyers", len(mpb_buyers))
                    st.metric("MPB Total Value", format_currency(mpb_buyers["Total_Value"].sum()))
                
                with col2:
                    st.metric("Other Brokers Total Buyers", len(other_buyers))
                    st.metric("Other Brokers Total Value", format_currency(other_buyers["Total_Value"].sum()))
            
            st.markdown("---")
            st.subheader(" Grade Performance Analysis")
            
            grade_analysis = mark_df[mark_df["Status_Clean"] == "sold"].groupby("Grade").agg({
                "Total Weight": "sum",
                "Price": "mean",
                "Lot No": "count"
            }).reset_index()
            
            grade_analysis.columns = ["Grade", "Total_Qty", "Avg_Price", "Num_Lots"]
            grade_analysis["Total_Value"] = grade_analysis["Total_Qty"] * grade_analysis["Avg_Price"]
            grade_analysis = grade_analysis.sort_values("Total_Value", ascending=False)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if not grade_analysis.empty and grade_analysis['Total_Value'].sum() > 0:
                    fig_grade_value = px.treemap(
                        grade_analysis[grade_analysis['Total_Value'] > 0], 
                        path=['Grade'], 
                        values='Total_Value',
                        title='Grade Distribution by Value',
                        color='Avg_Price',
                        color_continuous_scale='RdYlGn'
                    )
                    st.plotly_chart(fig_grade_value, use_container_width=True)
                else:
                    st.info("No data available for grade value treemap")
            
            with col2:
                fig_grade_price = px.bar(
                    grade_analysis.head(15), x="Grade", y="Avg_Price",
                    title='Average Price by Grade (Top 15)',
                    color='Avg_Price',
                    color_continuous_scale='Plasma'
                )
                fig_grade_price.update_layout(showlegend=False)
                st.plotly_chart(fig_grade_price, use_container_width=True)
            
            st.markdown("---")
            st.subheader(" Month-to-Date Performance")
            
            mtd_catalogued = historical['Catalogued'].sum()
            mtd_sold = historical['Sold'].sum()
            mtd_unsold = historical['Unsold'].sum()
            mtd_avg_price = mark_df[mark_df["Status_Clean"] == "sold"]["Price"].mean()
            mtd_proceeds = historical['Total_Proceeds'].sum()
            mtd_sell_pct = (mtd_sold / mtd_catalogued * 100) if mtd_catalogued > 0 else 0
            mtd_unsold_pct = (mtd_unsold / mtd_catalogued * 100) if mtd_catalogued > 0 else 0
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric(" MTD Catalogued", f"{mtd_catalogued:,.2f} kg")
                st.metric(" MTD Sold", f"{mtd_sold:,.2f} kg")
            
            with col2:
                st.metric(" MTD Avg Price", f"LKR {mtd_avg_price:,.2f}")
                st.metric(" MTD Sold %", f"{mtd_sell_pct:.2f}%")
            
            with col3:
                st.metric(" MTD Proceeds", format_currency(mtd_proceeds))
                st.metric(" MTD Unsold %", f"{mtd_unsold_pct:.2f}%")
            
            with col4:
                st.metric(" MTD Unsold", f"{mtd_unsold:,.2f} kg")
                st.metric(" Total Sales", len(historical))
            
            st.markdown("###  Sale-by-Sale Comparison")
            historical_display = historical.copy()
            historical_display['Avg_Price'] = historical_display['Avg_Price'].apply(lambda x: f"{x:,.2f}")
            historical_display['Total_Proceeds'] = historical_display['Total_Proceeds'].apply(lambda x: f"{x:,.0f}")
            historical_display['Catalogued'] = historical_display['Catalogued'].apply(lambda x: f"{x:,.2f}")
            historical_display['Sold'] = historical_display['Sold'].apply(lambda x: f"{x:,.2f}")
            historical_display['Unsold'] = historical_display['Unsold'].apply(lambda x: f"{x:,.2f}")
            historical_display['Sell_Pct'] = historical_display['Sell_Pct'].apply(lambda x: f"{x:.2f}%")
            
            st.dataframe(historical_display, use_container_width=True, hide_index=True)

# PRICE TRENDS
with tabs[5]:
    st.header("Price Trends Analysis")
    
    # Top level price metrics
    st.subheader("Key Price Metrics")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        overall_avg_price = latest_df[latest_df["Status_Clean"] == "sold"]["Price"].mean()
        st.metric("Overall Avg Price", f"LKR {overall_avg_price:,.2f}")
    with col2:
        max_price = latest_df["Price"].max()
        st.metric("Highest Price", f"LKR {max_price:,.2f}")
    with col3:
        min_sold_price = latest_df[latest_df["Status_Clean"] == "sold"]["Price"].min()
        st.metric("Lowest Sold Price", f"LKR {min_sold_price:,.2f}")
    with col4:
        total_sold_value = latest_df[latest_df["Status_Clean"] == "sold"]["Total Value"].sum()
        st.metric("Total Sold Value", format_currency(total_sold_value))
    
    st.markdown("---")
    
    # Main scatter plot
    st.subheader(" Price vs Weight Distribution")
    
    col1, col2 = st.columns([3, 1])
    with col2:
        show_unsold = st.checkbox("Show Unsold Lots", value=False)
        size_by_value = st.checkbox("Size by Total Value", value=True)
    
    # Prepare data for scatter plot
    scatter_df = latest_df.dropna(subset=["Total Weight", "Price"]).copy()
    if not show_unsold:
        scatter_df = scatter_df[scatter_df["Status_Clean"] == "sold"]
    
    fig_scatter = px.scatter(scatter_df, x="Total Weight", y="Price", color="Broker",
                           size="Total Value" if size_by_value else "Total Weight",
                           hover_data=["Grade", "Selling Mark", "Category", "Sub Elevation"],
                           title=f"Price vs Weight Distribution - Sale {latest_sale}",
                           color_discrete_sequence=px.colors.qualitative.Bold,
                           opacity=0.7)
    fig_scatter.update_layout(
        xaxis_title="Total Weight (kg)",
        yaxis_title="Price (LKR/kg)",
        hovermode='closest'
    )
    st.plotly_chart(fig_scatter, use_container_width=True)
    
    st.markdown("---")
    
    # Top Prices Analysis by Broker
    st.subheader(" Top 10 Highest Prices by Broker")
    
    # Get top 10 prices for each broker
    top_prices_by_broker = []
    for broker in latest_df["Broker"].unique():
        broker_df = latest_df[
            (latest_df["Broker"] == broker) & 
            (latest_df["Status_Clean"] == "sold")
        ].nlargest(10, "Price")
        
        for _, row in broker_df.iterrows():
            top_prices_by_broker.append({
                'Broker': broker,
                'Price': row['Price'],
                'Total Weight': row['Total Weight'],
                'Grade': row['Grade'],
                'Selling Mark': row.get('Selling Mark', 'N/A'),
                'Category': row['Category'],
                'Sub Elevation': row['Sub Elevation'],
                'Buyer': row.get('Buyer', 'N/A'),
                'Lot No': row.get('Lot No', 'N/A')
            })
    
    top_prices_df = pd.DataFrame(top_prices_by_broker)
    
    if not top_prices_df.empty:
        # Display top prices table
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("###  Top Prices Details")
            display_top_prices = top_prices_df[[
                'Broker', 'Price', 'Total Weight', 'Grade', 'Selling Mark'
            ]].copy()
            
            display_top_prices['Price'] = display_top_prices['Price'].apply(lambda x: f"LKR {x:,.2f}")
            display_top_prices['Total Weight'] = display_top_prices['Total Weight'].apply(lambda x: f"{x:,.2f} kg")
            
            st.dataframe(display_top_prices, use_container_width=True)
        
        with col2:
            # Top prices visualization
            fig_top_prices = px.box(top_prices_df, x='Broker', y='Price', color='Broker',
                                  title='Distribution of Top 10 Prices by Broker',
                                  points="all",
                                  hover_data=['Grade', 'Selling Mark'])
            fig_top_prices.update_layout(xaxis_tickangle=-45, showlegend=False)
            st.plotly_chart(fig_top_prices, use_container_width=True)
        
        # Selling Marks analysis for top prices
        st.markdown("###  Selling Marks in Top Prices")
        
        # Count selling marks in top prices
        selling_mark_top_counts = top_prices_df[top_prices_df['Selling Mark'] != 'N/A'].groupby('Selling Mark').agg({
            'Price': ['count', 'mean', 'max'],
            'Broker': 'nunique'
        }).round(2)
        
        selling_mark_top_counts.columns = ['Count', 'Avg_Price', 'Max_Price', 'Broker_Count']
        selling_mark_top_counts = selling_mark_top_counts.sort_values('Count', ascending=False)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Top selling marks by frequency in high prices
            fig_selling_marks_freq = px.bar(
                selling_mark_top_counts.head(15).reset_index(),
                x='Selling Mark', y='Count',
                title='Top Selling Marks in Highest Prices (Frequency)',
                color='Count',
                color_continuous_scale='Viridis'
            )
            fig_selling_marks_freq.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_selling_marks_freq, use_container_width=True)
        
        with col2:
            # Top selling marks by average price
            fig_selling_marks_avg = px.bar(
                selling_mark_top_counts.nlargest(15, 'Avg_Price').reset_index(),
                x='Selling Mark', y='Avg_Price',
                title='Top Selling Marks by Average Price in High-Priced Lots',
                color='Avg_Price',
                color_continuous_scale='Plasma'
            )
            fig_selling_marks_avg.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_selling_marks_avg, use_container_width=True)
    
    st.markdown("---")
    
    # Price Distribution Analysis
    st.subheader(" Price Distribution Analysis")
    
    col1, col2 = st.columns(2)
    
    sold_df = latest_df[latest_df["Status_Clean"] == "sold"]

    with col1:
        # Price distribution by broker
        fig_price_dist = px.violin(sold_df, x='Broker', y='Price', color='Broker',
                                 title='Price Distribution by Broker (Violin Plot)',
                                 box=True, points=False)
        fig_price_dist.update_layout(xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig_price_dist, use_container_width=True)
    
    with col2:
        # Price distribution by category
        fig_cat_price = px.box(sold_df, x='Category', y='Price', color='Category',
                             title='Price Distribution by Category')
        fig_cat_price.update_layout(xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig_cat_price, use_container_width=True)
    
    st.markdown("---")
    
    # Grade-wise Price Analysis
    st.subheader(" Grade-wise Price Performance")
    
    # Get top grades by average price
    grade_price_analysis = sold_df.groupby('Grade').agg({
        'Price': ['mean', 'median', 'std', 'count'],
        'Total Weight': 'sum',
        'Total Value': 'sum'
    }).round(2)
    
    grade_price_analysis.columns = ['Avg_Price', 'Median_Price', 'Std_Price', 'Count', 'Total_Weight', 'Total_Value']
    grade_price_analysis = grade_price_analysis[grade_price_analysis['Count'] >= 3]  # Filter for meaningful stats
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Top grades by average price
        top_grades_price = grade_price_analysis.nlargest(15, 'Avg_Price')
        fig_grade_avg_price = px.bar(top_grades_price.reset_index(), 
                                   x='Grade', y='Avg_Price',
                                   title='Top 15 Grades by Average Price',
                                   color='Avg_Price',
                                   color_continuous_scale='RdYlGn',
                                   hover_data=['Count', 'Total_Weight'])
        fig_grade_avg_price.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_grade_avg_price, use_container_width=True)
    
    with col2:
        # Grade price stability (low std deviation)
        stable_grades = grade_price_analysis.nsmallest(15, 'Std_Price')
        fig_grade_stable = px.bar(stable_grades.reset_index(),
                                x='Grade', y='Std_Price',
                                title='Most Price-Stable Grades (Lowest Std Deviation)',
                                color='Std_Price',
                                color_continuous_scale='Blues',
                                hover_data=['Avg_Price', 'Count'])
        fig_grade_stable.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_grade_stable, use_container_width=True)
    
    st.markdown("---")
    
    # Broker Price Performance Comparison
    st.subheader(" Broker Price Performance Comparison")
    
    broker_price_stats = sold_df.groupby('Broker').agg({
        'Price': ['mean', 'median', 'std', 'min', 'max', 'count'],
        'Total Weight': 'sum',
        'Total Value': 'sum'
    }).round(2)
    
    broker_price_stats.columns = ['Avg_Price', 'Median_Price', 'Std_Price', 'Min_Price', 'Max_Price', 'Count', 'Total_Weight', 'Total_Value']
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Broker average price comparison
        fig_broker_avg = px.bar(broker_price_stats.reset_index(),
                              x='Broker', y='Avg_Price',
                              title='Average Price by Broker',
                              color='Avg_Price',
                              color_continuous_scale='Viridis',
                              hover_data=['Count', 'Total_Weight'])
        fig_broker_avg.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_broker_avg, use_container_width=True)
    
    with col2:
        # Broker price range
        broker_price_stats['Price_Range'] = broker_price_stats['Max_Price'] - broker_price_stats['Min_Price']
        fig_broker_range = px.bar(broker_price_stats.reset_index(),
                                x='Broker', y='Price_Range',
                                title='Price Range by Broker (Max - Min)',
                                color='Price_Range',
                                color_continuous_scale='RdBu',
                                hover_data=['Avg_Price', 'Std_Price'])
        fig_broker_range.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_broker_range, use_container_width=True)
    
    # Detailed broker price statistics table
    st.markdown("###  Detailed Broker Price Statistics")
    
    broker_display_stats = broker_price_stats.reset_index()
    broker_display_stats['Avg_Price'] = broker_display_stats['Avg_Price'].apply(lambda x: f"LKR {x:,.2f}")
    broker_display_stats['Median_Price'] = broker_display_stats['Median_Price'].apply(lambda x: f"LKR {x:,.2f}")
    broker_display_stats['Std_Price'] = broker_display_stats['Std_Price'].apply(lambda x: f"LKR {x:,.2f}")
    broker_display_stats['Min_Price'] = broker_display_stats['Min_Price'].apply(lambda x: f"LKR {x:,.2f}")
    broker_display_stats['Max_Price'] = broker_display_stats['Max_Price'].apply(lambda x: f"LKR {x:,.2f}")
    broker_display_stats['Price_Range'] = broker_display_stats['Price_Range'].apply(lambda x: f"LKR {x:,.2f}")
    broker_display_stats['Total_Weight'] = broker_display_stats['Total_Weight'].apply(lambda x: f"{format_large_number(x)} kg")
    broker_display_stats['Total_Value'] = broker_display_stats['Total_Value'].apply(lambda x: format_currency(x))
    broker_display_stats['Count'] = broker_display_stats['Count'].apply(lambda x: f"{x:,}")
    
    st.dataframe(broker_display_stats, use_container_width=True)
    
    st.markdown("---")
    
    # Price Trends Over Time (if multiple sales data available)
    if len(data['Sale_No'].unique()) > 1:
        st.subheader(" Price Trends Across Sales")
        
        price_trends = data[data["Status_Clean"] == "sold"].groupby(['Sale_No', 'Broker']).agg({
            'Price': 'mean',
            'Total Weight': 'sum',
            'Total Value': 'sum'
        }).reset_index()
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig_trend_avg = px.line(price_trends, x='Sale_No', y='Price', color='Broker',
                                  title='Average Price Trend by Broker',
                                  markers=True)
            st.plotly_chart(fig_trend_avg, use_container_width=True)
        
        with col2:
            fig_trend_value = px.area(price_trends, x='Sale_No', y='Total Value', color='Broker',
                                    title='Total Value Trend by Broker',
                                    markers=True)
            st.plotly_chart(fig_trend_value, use_container_width=True)
    
    # Price vs Quantity Correlation Analysis
    st.markdown("---")
    st.subheader(" Price-Quantity Correlation Analysis")
    
    correlation_df = sold_df[['Price', 'Total Weight', 'Total Value']].corr()
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Correlation heatmap
        fig_corr = px.imshow(correlation_df,
                           title='Price-Quantity-Value Correlation Matrix',
                           color_continuous_scale='RdBu',
                           aspect="auto",
                           text_auto=True)
        st.plotly_chart(fig_corr, use_container_width=True)
    
    with col2:
        # Price density by broker
        fig_density = px.density_heatmap(sold_df, x='Total Weight', y='Price', facet_col='Broker',
                                       title='Price-Weight Density by Broker',
                                       color_continuous_scale='Viridis',
                                       facet_col_wrap=3)
        st.plotly_chart(fig_density, use_container_width=True)
    
    # CSV export options removed - PDF export is available in the report section


st.markdown("---")
st.header(" Professional PDF Report Generation")
st.markdown("Generate elevation-wise PDF reports with **selective report options** - **3-15 seconds** based on selection")

# Quick info section
with st.expander("ℹ What's Included in Reports"):
    st.markdown("""
    ** 5 Required Reports (Sub Elevation Wise):**
    
     **Report 1: Broker Grade-wise Sold Percentages** - Each broker's grade wise sold percentages by Sub Elevation  
     **Report 2: Broker Grade-wise Unsold Percentages** - Each broker's grade wise unsold percentages by Sub Elevation  
     **Report 3: Broker Grade-wise Outsold Percentages** - Each broker's grade wise outsold percentages by Sub Elevation  
     **Report 4: Broker Grade-wise Sold Quantities & Avg Prices** - Sold quantities and average prices by Sub Elevation  
     **Report 5: Outlots Purchased Buyer Profiles** - Buyer profiles grade wise by Sub Elevation  
    
    ** Optimized Performance:**
    - **Tables only** - Fast generation without charts (default)
    - **Optional Charts** - Add summary charts/graphs if needed (slower but visual)
    - **Selective generation** - Choose only reports you need
    - Current sale only: **3-5 seconds** (1 report) to **10-15 seconds** (all 5 reports)
    - With charts: **+5-8 seconds** additional time
    
    ** Report Structure (SUMMARIZED):**
    - All reports organized by **Sub Elevation** (L, M, UH, UM, WH, WM)
    - Each broker shown separately with **summary table + bar chart** by elevation
    - **Top 10 grades only** per elevation (not all grades - reduces pages significantly)
    - **Bar charts for percentages** included in each report
    - Accurate calculations based on real data
    - Professional A4 format with page numbers
    - **Estimated 15-25 pages** (much shorter than before)
    
    ** Summary Charts (Optional):**
    - Market Share by Broker (Pie Chart)
    - Overall Sale Status Distribution (Pie Chart)
    - Broker Performance - Sold Percentage (Bar Chart)
    - Elevation Performance - Status Percentages (Stacked Bar Chart)
    
    ** Bar Charts in Reports:**
    - Each report includes bar charts showing percentages by Sub Elevation
    - Mini bar charts for top grades within each elevation
    - Visual representation for easy reference
    """)

# Report configuration
st.subheader(" Quick Report Generation")

col1, col2, col3 = st.columns(3)

with col1:
    report_scope = st.selectbox(
        " Data Scope",
        ["Current Sale Only", "Last 3 Sales", "Last 5 Sales", "All Available Sales"],
        index=0,
        help="Current sale is fastest (5-8 seconds)"
    )

with col2:
    report_format = st.selectbox(
        " Report Format",
        ["Standard (Fast)", "Detailed (More Data)"],
        index=0,
        help="Standard format is optimized for speed"
    )

with col3:
    st.markdown("#####  Report Selection")
    st.info("Select which reports to include in PDF")

# Quick templates
st.markdown("---")
st.subheader(" Quick Generation Templates")

template_col1, template_col2, template_col3 = st.columns(3)

with template_col1:
    quick_report_btn = st.button(
        " Quick Report - Current Sale",
        use_container_width=True,
        type="primary",
        help="Generate report for current sale only (5-8 seconds)"
    )

with template_col2:
    weekly_report_btn = st.button(
        " Weekly Report - Last 3 Sales",
        use_container_width=True,
        help="Generate report for last 3 sales (10-12 seconds)"
    )

with template_col3:
    full_report_btn = st.button(
        " Full Report - All Sales",
        use_container_width=True,
        help="Generate comprehensive report (15-20 seconds)"
    )

# Report Selection Checkboxes
st.markdown("---")
st.subheader(" Select Reports to Include")

col1, col2, col3 = st.columns(3)

with col1:
    report1 = st.checkbox("Report 1: Broker Grade-wise Sold % (Sub Elevation)", value=True, 
                         help="Each broker's grade wise sold percentages by Sub Elevation")
    report2 = st.checkbox("Report 2: Broker Grade-wise Unsold % (Sub Elevation)", value=True,
                         help="Each broker's grade wise unsold percentages by Sub Elevation")

with col2:
    report3 = st.checkbox("Report 3: Broker Grade-wise Outsold % (Sub Elevation)", value=True,
                         help="Each broker's grade wise outsold percentages by Sub Elevation")
    report4 = st.checkbox("Report 4: Broker Grade-wise Sold Qty & Avg Prices (Sub Elevation)", value=True,
                         help="Each broker's grade wise sold quantities and average prices by Sub Elevation")

with col3:
    report5 = st.checkbox("Report 5: Outlots Purchased Buyer Profiles (Grade wise, Sub Elevation)", value=True,
                         help="Outlots purchased buyer profiles, grade wise by Sub Elevation")
    st.markdown("---")
    st.markdown("** Summary Reports (Optional):**")
    summary_market = st.checkbox(" Overall Market Performance Summary", value=False,
                                help="Overall market statistics with MPB highlighting")
    summary_broker_perf = st.checkbox(" Broker Performance Comparison", value=False,
                                      help="Detailed broker performance by Sub Elevation with MPB highlighting")

# Main generate button
st.markdown("---")
generate_col1, generate_col2, generate_col3 = st.columns([1, 2, 1])

with generate_col2:
    generate_button = st.button(
        " GENERATE PROFESSIONAL PDF REPORT",
        type="primary",
        use_container_width=True,
        help="Generate PDF with selected reports only (faster generation)"
    )

# Report generation logic
if generate_button or quick_report_btn or weekly_report_btn or full_report_btn:
    try:
        # Check reportlab installation
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
        except ImportError:
            st.error("""
             **ReportLab library not found!**
            
            Please install it using:
            ```bash
            pip install reportlab
            ```
            
            Then restart your Streamlit application.
            """)
            st.stop()
        
        # Determine scope based on button clicked
        if quick_report_btn:
            report_scope = "Current Sale Only"
        elif weekly_report_btn:
            report_scope = "Last 3 Sales"
        elif full_report_btn:
            report_scope = "All Available Sales"
        
        # Generate report with progress tracking
        with st.spinner(f" Generating {report_scope} report..."):
            import time
            
            # Progress indicators
            progress_bar = st.progress(0)
            status_text = st.empty()
            time_estimate = st.empty()
            
            # Estimate completion time
            if report_scope == "Current Sale Only":
                estimated_time = "8-12 seconds"
            elif report_scope in ["Last 3 Sales", "Last 5 Sales"]:
                estimated_time = "15-20 seconds"
            else:
                estimated_time = "25-30 seconds"
            
            time_estimate.info(f" Estimated completion time: {estimated_time}")
            
            # Step 1: Filter data
            status_text.text(" Processing sale data...")
            progress_bar.progress(15)
            time.sleep(0.2)
            
            if report_scope == "Current Sale Only":
                report_data = data[data["Sale_No"] == latest_sale]
            elif report_scope == "Last 3 Sales":
                recent_sales = sorted(data["Sale_No"].unique())[-3:]
                report_data = data[data["Sale_No"].isin(recent_sales)]
            elif report_scope == "Last 5 Sales":
                recent_sales = sorted(data["Sale_No"].unique())[-5:]
                report_data = data[data["Sale_No"].isin(recent_sales)]
            else:
                report_data = data
            
            # Step 2: Calculate metrics
            status_text.text(" Calculating performance metrics...")
            progress_bar.progress(30)
            time.sleep(0.2)
            
            # Step 3: Generate broker analysis
            status_text.text("Analyzing all 8 brokers...")
            progress_bar.progress(50)
            time.sleep(0.3)
            
            # Step 4: Generate PDF
            status_text.text(" Creating PDF document...")
            progress_bar.progress(70)
            
            # Import the optimized function
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch, cm
            from reportlab.lib.enums import TA_CENTER
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
            from reportlab.pdfgen import canvas
            import io
            
            # Prepare selected sections
            include_sections = {
                'report1_sold_pct': report1,
                'report2_unsold_pct': report2,
                'report3_outsold_pct': report3,
                'report4_sold_qty_price': report4,
                'report5_buyer_profiles': report5,
                'summary_market': summary_market,
                'summary_broker_perf': summary_broker_perf
            }
            
            # Check if at least one report is selected (excluding charts)
            report_selections = [report1, report2, report3, report4, report5]
            if not any(report_selections):
                st.warning(" Please select at least one report to generate!")
                st.stop()
            
            # Call the optimized PDF generator
            pdf_data = generate_fast_pdf_report(
                report_data,
                latest_df,
                output_filename=f"report_sale_{latest_sale}.pdf",
                include_sections=include_sections
            )
            
            # Step 5: Finalize
            status_text.text(" Finalizing document...")
            progress_bar.progress(95)
            time.sleep(0.2)
            
            progress_bar.progress(100)
            status_text.text(" Report generation complete!")
            time.sleep(0.3)
            
            # Clear progress indicators
            progress_bar.empty()
            status_text.empty()
            time_estimate.empty()
            
            # Success message
            st.success(f"""
             **PDF Report Generated Successfully!**
            
             Scope: {report_scope}  
             Pages: Approximately 12-15 pages  
             Size: {len(pdf_data) / 1024:.1f} KB  
             Generated: {datetime.now().strftime('%H:%M:%S')}
            """)
            
            # Download section
            st.markdown("---")
            st.subheader(" Download Your Report")
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"Tea_Auction_Report_Sale_{latest_sale}_{timestamp}.pdf"
            
            col1, col2, col3 = st.columns([1, 2, 1])
            
            with col2:
                st.download_button(
                    label=" DOWNLOAD COMPLETE PDF REPORT",
                    data=pdf_data,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True,
                    help="Download your comprehensive PDF report"
                )
            
            # Report summary
            st.markdown("---")
            st.subheader(" Report Contents Summary")
            
            summary_col1, summary_col2, summary_col3 = st.columns(3)
            
            with summary_col1:
                st.markdown("""
                ** Market Overview**
                - Executive summary
                - Market metrics
                - Key performance indicators
                - Overall elevation performance
                """)
            
            with summary_col2:
                st.markdown("""
                **All 8 Brokers**
                - Performance by elevation
                - Grade breakdown per elevation
                - Sold/Unsold/Outsold % per elevation
                - Top 8 grades per elevation
                """)
            
            with summary_col3:
                st.markdown("""
                ** Buyer Analysis**
                - Top 5 buyers
                - Elevation preferences
                - Grade patterns by elevation
                - Price analysis by elevation
                """)
            
            # Additional details
            with st.expander(" Detailed Report Breakdown"):
                st.markdown(f"""
                **Report Details:**
                
                1. **Title Page**
                   - Company branding
                   - Sale number: {latest_sale}
                   - Generation timestamp
                   - Data coverage period
                
                2. **Executive Summary (Page 2)**
                   - Total market value
                   - Average prices
                   - Sell-through rates
                   - Key insights
                
                3. **Broker Performance (Page 3)**
                   - All 8 brokers comparison table
                   - Total quantities and percentages
                   - Average prices achieved
                   - Lot counts
                
                4. **Broker Elevation & Grade Analysis (Pages 4-12)**
                   - Each broker gets dedicated pages
                   - Breakdown by elevation
                   - Each elevation shows:
                     * Top 8 grades
                     * Sold/Unsold/Outsold percentages
                     * Average prices by grade
                     * Lot counts
                
                5. **Overall Elevation Analysis (Page 13)**
                   - Market-wide elevation performance
                   - Quantity and percentage breakdowns
                   - Price comparisons across elevations
                
                6. **Buyer Profiles by Elevation (Pages 14-17)**
                   - Top 5 buyers detailed analysis
                   - Breakdown by elevation for each buyer
                   - Grade preferences per elevation
                   - Average prices paid by elevation
                   - Purchase quantities by elevation & grade
                
                7. **Recommendations (Page 18)**
                   - Elevation-based insights
                   - Grade recommendations by elevation
                   - Broker performance insights
                   - Strategic suggestions
                
                **Footer on Every Page:**
                - Company name
                - Generation date/time
                - Page numbers (Page X of Y)
                """)
            
            # Save to history (optional)
            if 'report_history' not in st.session_state:
                st.session_state.report_history = []
            
            st.session_state.report_history.append({
                'title': f"Sale {latest_sale} Report",
                'date': datetime.now().strftime('%Y-%m-%d %H:%M'),
                'scope': report_scope,
                'size': f"{len(pdf_data) / 1024:.1f} KB"
            })
            
    except Exception as e:
        st.error(f"""
         **Error generating PDF report:**
        
        {str(e)}
        
        **Troubleshooting Steps:**
        1. Ensure `reportlab` is installed: `pip install reportlab`
        2. Check that data files are accessible
        3. Verify sufficient memory available
        4. Try "Current Sale Only" for faster generation
        """)
        
        # Show detailed error for debugging
        with st.expander(" Technical Error Details"):
            import traceback
            st.code(traceback.format_exc())

# Report history section
if 'report_history' in st.session_state and st.session_state.report_history:
    st.markdown("---")
    st.subheader(" Recent Report Generation History")
    
    # Show last 5 reports
    history_data = []
    for report in st.session_state.report_history[-5:]:
        history_data.append([
            report.get('title', 'Report'),
            report.get('date', 'N/A'),
            report.get('scope', 'N/A'),
            report.get('size', 'N/A')
        ])
    
    history_df = pd.DataFrame(
        history_data,
        columns=['Report', 'Generated', 'Scope', 'Size']
    )
    
    st.dataframe(history_df, use_container_width=True, hide_index=True)
    
    if st.button(" Clear History"):
        st.session_state.report_history = []
        st.rerun()

# Tips and best practices
st.markdown("---")
with st.expander(" PDF Generation Tips & Best Practices"):
    st.markdown("""
    ** For Fastest Generation:**
    - Use "Current Sale Only" (8-12 seconds)
    - Standard format is optimized for speed
    - All 8 brokers with elevation breakdown included
    
    ** Understanding Elevation-Based Reports:**
    - Each broker's data is organized by elevation first
    - Within each elevation, top 8 grades are shown
    - **Sold %**: Percentage of catalogued quantity sold per elevation
    - **Unsold %**: Percentage that remained unsold per elevation
    - **Outsold %**: Percentage that went to competing brokers per elevation
    - Buyers show their purchasing patterns by elevation
    
    ** Best Use Cases:**
    - **Daily Operations**: Current Sale Only (elevation focus)
    - **Weekly Reviews**: Last 3 Sales (elevation trends)
    - **Monthly Analysis**: Last 5 Sales (elevation performance)
    - **Comprehensive Audit**: All Sales (complete elevation history)
    
    ** Report Format:**
    - Professional A4 size (portrait)
    - 15-20 pages typical (with elevation breakdown)
    - Clean tables organized by elevation
    - Page numbers on every page
    - Automatic timestamp
    
    ** File Management:**
    - Reports are generated fresh each time
    - Download and save important reports
    - Filename includes sale number and timestamp
    - Average file size: 75-200 KB (more data with elevations)
    
    ** Technical Notes:**
    - Requires `reportlab` library
    - Elevation-grade calculations optimized
    - All 8 brokers processed in single pass
    - Buyer-elevation data pre-calculated
    - Can generate hundreds of reports per hour
    """)

# Performance information
st.markdown("---")
st.info("""
** Performance Optimized - Elevation-Based Analysis:**  
This PDF generator provides comprehensive elevation-wise breakdown:
-  All 8 brokers with elevation analysis
-  Each elevation shows grade performance
-  Sold/Unsold/Outsold % per elevation
-  Buyer preferences by elevation
-  Pre-calculated metrics for speed
-  Progress tracking with time estimates
-  15-20 pages of detailed insights
""")

# Footer
st.markdown("---")
st.caption(" Professional PDF Reports | Optimized for Tea Auction Business Intelligence")

# Remove the problematic pie chart at the end that was causing the error
st.success(" Dashboard Loaded Successfully with OKLO MAIN AUCTION DATA")


