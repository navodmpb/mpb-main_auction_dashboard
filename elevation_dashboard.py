# ============================================================
# ELEVATION PERFORMANCE DASHBOARD MODULE
# NEW VISUALIZATIONS FOR ELEVATION-WISE ANALYSIS
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from streamlit_autorefresh import rerun

@st.cache_data
def calculate_elevation_performance(df):
    """Calculate comprehensive elevation performance metrics"""
    elev_perf = df.groupby('Sub Elevation').apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Total_Value': x["Total Value"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean(),
        'Num_Lots': len(x),
        'Num_Sold': len(x[x["Status_Clean"] == "sold"])
    }), include_groups=False).reset_index()
    
    elev_perf['Total_Sold_Side'] = elev_perf['Sold'] + elev_perf['Outsold']
    elev_perf['Sold_Pct'] = (elev_perf['Total_Sold_Side'] / elev_perf['Catalogued'] * 100).fillna(0)
    elev_perf['Unsold_Pct'] = (elev_perf['Unsold'] / elev_perf['Catalogued'] * 100).fillna(0)
    elev_perf['Outsold_Pct'] = (elev_perf['Outsold'] / elev_perf['Catalogued'] * 100).fillna(0)
    
    return elev_perf

@st.cache_data
def calculate_broker_elevation_performance(df):
    """Calculate broker performance across all elevations"""
    broker_elev = df.groupby(['Broker', 'Sub Elevation']).apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean(),
        'Num_Lots': len(x)
    }), include_groups=False).reset_index()
    
    broker_elev['Total_Sold_Side'] = broker_elev['Sold'] + broker_elev['Outsold']
    broker_elev['Sold_Pct'] = (broker_elev['Total_Sold_Side'] / broker_elev['Catalogued'] * 100).fillna(0)
    
    return broker_elev

def create_elevation_performance_dashboard(latest_df, data):
    """Create comprehensive elevation performance dashboard section"""
    
    st.subheader("🏔️ Elevation Performance Dashboard")
    
    # Calculate elevation metrics
    elev_perf = calculate_elevation_performance(latest_df)
    
    # Key metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Total Elevations", len(elev_perf))
        st.metric("Total Catalogued", f"{elev_perf['Catalogued'].sum():,.0f} kg")
    
    with col2:
        total_market_sold = elev_perf['Total_Sold_Side'].sum()
        st.metric("Total Sold+Outsold", f"{total_market_sold:,.0f} kg")
        avg_sold_pct = (total_market_sold / elev_perf['Catalogued'].sum() * 100)
        st.metric("Market Sold %", f"{avg_sold_pct:.1f}%")
    
    with col3:
        st.metric("Highest Avg Price", f"LKR {elev_perf['Avg_Price'].max():,.2f}")
        st.metric("Lowest Avg Price", f"LKR {elev_perf['Avg_Price'].min():,.2f}")
    
    with col4:
        best_elev = elev_perf.loc[elev_perf['Sold_Pct'].idxmax(), 'Sub Elevation']
        best_sold_pct = elev_perf['Sold_Pct'].max()
        st.metric("Best Performing Elev", f"{best_elev} ({best_sold_pct:.1f}%)")
    
    with col5:
        st.metric("Total Lots", f"{elev_perf['Num_Lots'].sum():,}")
        st.metric("Total Value", f"LKR {elev_perf['Total_Value'].sum()/1e6:.2f}M")
    
    st.markdown("---")
    
    # Stacked bar chart - Status by elevation
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Elevation Sale Status Distribution")
        
        fig_stacked = px.bar(
            elev_perf,
            x='Sub Elevation',
            y=['Sold', 'Unsold', 'Outsold'],
            title='Quantity by Status per Elevation (kg)',
            labels={'value': 'Quantity (kg)', 'variable': 'Status'},
            barmode='stack',
            color_discrete_map={
                'Sold': '#28a745',
                'Unsold': '#dc3545',
                'Outsold': '#ffc107'
            }
        )
        fig_stacked.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_stacked, use_container_width=True)
    
    with col2:
        st.markdown("### 📈 Elevation Sold % Performance")
        
        # Color code based on performance
        colors_list = ['#28a745' if x >= 70 else '#ffc107' if x >= 50 else '#dc3545' for x in elev_perf['Sold_Pct']]
        
        fig_sold_pct = px.bar(
            elev_perf,
            x='Sub Elevation',
            y='Sold_Pct',
            title='Sold % by Elevation (Sold+Outsold)',
            labels={'Sold_Pct': 'Sold %', 'Sub Elevation': 'Elevation'},
            text='Sold_Pct',
            color='Sold_Pct',
            color_continuous_scale=[[0, '#dc3545'], [0.5, '#ffc107'], [1, '#28a745']]
        )
        fig_sold_pct.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_sold_pct.update_layout(xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig_sold_pct, use_container_width=True)
    
    # Heatmap - Broker performance across elevations
    st.markdown("---")
    st.markdown("### 🔥 Heatmap: Broker Performance Across Elevations")
    
    broker_elev = calculate_broker_elevation_performance(latest_df)
    
    # Create pivot table for heatmap
    heatmap_data = broker_elev.pivot_table(
        index='Broker',
        columns='Sub Elevation',
        values='Sold_Pct',
        aggfunc='mean'
    )
    
    fig_heatmap = px.imshow(
        heatmap_data,
        title='Broker Sold % by Elevation (Green=High, Red=Low)',
        labels=dict(x="Elevation", y="Broker", color="Sold %"),
        color_continuous_scale='RdYlGn',
        aspect="auto",
        text_auto='.1f',
        zmin=0,
        zmax=100
    )
    fig_heatmap.update_xaxes(side="bottom")
    st.plotly_chart(fig_heatmap, use_container_width=True)
    
    # Price trends by elevation
    st.markdown("---")
    st.markdown("### 💹 Price Trends by Elevation")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_price = px.bar(
            elev_perf,
            x='Sub Elevation',
            y='Avg_Price',
            title='Average Price by Elevation (LKR/kg)',
            color='Avg_Price',
            color_continuous_scale='Viridis',
            text='Avg_Price'
        )
        fig_price.update_traces(texttemplate='LKR %{text:,.0f}', textposition='outside')
        fig_price.update_layout(xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig_price, use_container_width=True)
    
    with col2:
        # Price range by elevation
        broker_elev_stats = broker_elev.groupby('Sub Elevation').agg({
            'Avg_Price': ['min', 'max', 'mean']
        }).reset_index()
        
        broker_elev_stats.columns = ['Sub Elevation', 'Min_Price', 'Max_Price', 'Avg_Price']
        broker_elev_stats['Price_Range'] = broker_elev_stats['Max_Price'] - broker_elev_stats['Min_Price']
        
        fig_range = px.bar(
            broker_elev_stats,
            x='Sub Elevation',
            y=['Min_Price', 'Avg_Price', 'Max_Price'],
            title='Price Range by Elevation',
            labels={'value': 'Price (LKR/kg)', 'variable': 'Price Type'},
            barmode='group'
        )
        fig_range.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_range, use_container_width=True)
    
    # Elevation summary table
    st.markdown("---")
    st.markdown("### 📋 Elevation Performance Summary Table")
    
    display_elev = elev_perf.copy()
    display_elev['Catalogued'] = display_elev['Catalogued'].apply(lambda x: f"{x:,.0f} kg")
    display_elev['Total_Sold_Side'] = display_elev['Total_Sold_Side'].apply(lambda x: f"{x:,.0f} kg")
    display_elev['Sold_Pct'] = display_elev['Sold_Pct'].apply(lambda x: f"{x:.1f}%")
    display_elev['Avg_Price'] = display_elev['Avg_Price'].apply(lambda x: f"LKR {x:,.2f}")
    display_elev['Total_Value'] = display_elev['Total_Value'].apply(lambda x: f"LKR {x:,.0f}")
    
    st.dataframe(
        display_elev[['Sub Elevation', 'Catalogued', 'Total_Sold_Side', 'Sold_Pct', 'Avg_Price', 'Total_Value', 'Num_Lots']],
        use_container_width=True,
        hide_index=True
    )

def create_grade_performance_matrix(latest_df):
    """Create interactive grade performance matrix by elevation"""
    
    st.markdown("---")
    st.subheader("📊 Grade Performance Matrix (By Elevation)")
    
    # Elevation filter
    elevations = sorted(latest_df['Sub Elevation'].unique())
    selected_elevation = st.selectbox(
        "Select Elevation to Analyze",
        elevations,
        key="grade_matrix_elev"
    )
    
    # Get data for selected elevation
    elev_df = latest_df[latest_df['Sub Elevation'] == selected_elevation]
    
    # Calculate grade performance
    grade_perf = elev_df.groupby('Grade').apply(lambda x: pd.Series({
        'Catalogued': x["Total Weight"].sum(),
        'Sold': x[x["Status_Clean"] == "sold"]["Total Weight"].sum(),
        'Unsold': x[x["Status_Clean"] == "unsold"]["Total Weight"].sum(),
        'Outsold': x[x["Status_Clean"] == "outsold"]["Total Weight"].sum(),
        'Avg_Price': x[x["Status_Clean"] == "sold"]["Price"].mean(),
        'Num_Lots': len(x)
    }), include_groups=False).reset_index()
    
    grade_perf['Total_Sold_Side'] = grade_perf['Sold'] + grade_perf['Outsold']
    grade_perf['Sold_Pct'] = (grade_perf['Total_Sold_Side'] / grade_perf['Catalogued'] * 100).fillna(0)
    grade_perf = grade_perf.sort_values('Catalogued', ascending=False).head(15)
    
    # Display metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Grades", len(grade_perf))
        st.metric("Total Quantity", f"{grade_perf['Catalogued'].sum():,.0f} kg")
    with col2:
        st.metric("Avg Sold %", f"{grade_perf['Sold_Pct'].mean():.1f}%")
        st.metric("Avg Price", f"LKR {grade_perf['Avg_Price'].mean():,.2f}")
    with col3:
        st.metric("Top Grade", grade_perf.iloc[0]['Grade'])
        st.metric("Highest Avg Price", f"LKR {grade_perf['Avg_Price'].max():,.2f}")
    with col4:
        st.metric("Lowest Avg Price", f"LKR {grade_perf['Avg_Price'].min():,.2f}")
    
    # Visualizations
    col1, col2 = st.columns(2)
    
    with col1:
        fig_grade_sold = px.bar(
            grade_perf,
            x='Grade',
            y='Sold_Pct',
            title=f'Grade Sold % - {selected_elevation}',
            color='Sold_Pct',
            color_continuous_scale='RdYlGn',
            text='Sold_Pct'
        )
        fig_grade_sold.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_grade_sold.update_layout(xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig_grade_sold, use_container_width=True)
    
    with col2:
        fig_grade_price = px.bar(
            grade_perf,
            x='Grade',
            y='Avg_Price',
            title=f'Average Price by Grade - {selected_elevation}',
            color='Avg_Price',
            color_continuous_scale='Viridis',
            text='Avg_Price'
        )
        fig_grade_price.update_traces(texttemplate='LKR %{text:,.0f}', textposition='outside')
        fig_grade_price.update_layout(xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig_grade_price, use_container_width=True)
    
    # Detailed table
    st.markdown("### 📋 Top Grades Details")
    display_grade = grade_perf.copy()
    display_grade['Catalogued'] = display_grade['Catalogued'].apply(lambda x: f"{x:,.0f} kg")
    display_grade['Total_Sold_Side'] = display_grade['Total_Sold_Side'].apply(lambda x: f"{x:,.0f} kg")
    display_grade['Sold_Pct'] = display_grade['Sold_Pct'].apply(lambda x: f"{x:.1f}%")
    display_grade['Avg_Price'] = display_grade['Avg_Price'].apply(lambda x: f"LKR {x:,.2f}")
    
    st.dataframe(
        display_grade[['Grade', 'Catalogued', 'Total_Sold_Side', 'Sold_Pct', 'Avg_Price', 'Num_Lots']],
        use_container_width=True,
        hide_index=True
    )

def create_broker_comparison_view(latest_df):
    """Create side-by-side broker comparison with elevation breakdown"""
    
    st.markdown("---")
    st.subheader("🢠 Broker Comparison View (Elevation-wise)")
    
    brokers = sorted(latest_df['Broker'].unique())
    selected_brokers = st.multiselect(
        "Select Brokers to Compare",
        brokers,
        default=brokers[:2] if len(brokers) >= 2 else brokers,
        key="broker_compare"
    )
    
    if not selected_brokers:
        st.warning("Please select at least one broker")
        return
    
    # Create comparison data
    comparison_data = []
    for broker in selected_brokers:
        broker_df = latest_df[latest_df['Broker'] == broker]
        
        for elev in sorted(latest_df['Sub Elevation'].unique()):
            elev_df = broker_df[broker_df['Sub Elevation'] == elev]
            
            if not elev_df.empty:
                comparison_data.append({
                    'Broker': broker,
                    'Elevation': elev,
                    'Catalogued': elev_df['Total Weight'].sum(),
                    'Sold': elev_df[elev_df['Status_Clean'] == 'sold']['Total Weight'].sum(),
                    'Avg_Price': elev_df[elev_df['Status_Clean'] == 'sold']['Price'].mean(),
                    'Num_Lots': len(elev_df)
                })
    
    comparison_df = pd.DataFrame(comparison_data)
    comparison_df['Sold_Pct'] = (comparison_df['Sold'] / comparison_df['Catalogued'] * 100).fillna(0)
    
    # Comparison charts
    col1, col2 = st.columns(2)
    
    with col1:
        fig_comp_qty = px.bar(
            comparison_df,
            x='Elevation',
            y='Catalogued',
            color='Broker',
            title='Total Quantity by Elevation - Broker Comparison',
            barmode='group'
        )
        st.plotly_chart(fig_comp_qty, use_container_width=True)
    
    with col2:
        fig_comp_sold = px.bar(
            comparison_df,
            x='Elevation',
            y='Sold_Pct',
            color='Broker',
            title='Sold % by Elevation - Broker Comparison',
            barmode='group'
        )
        st.plotly_chart(fig_comp_sold, use_container_width=True)
    
    # Comparison table
    st.markdown("### 📋 Broker Performance Comparison Table")
    display_comp = comparison_df.copy()
    display_comp['Catalogued'] = display_comp['Catalogued'].apply(lambda x: f"{x:,.0f} kg")
    display_comp['Sold'] = display_comp['Sold'].apply(lambda x: f"{x:,.0f} kg")
    display_comp['Sold_Pct'] = display_comp['Sold_Pct'].apply(lambda x: f"{x:.1f}%")
    display_comp['Avg_Price'] = display_comp['Avg_Price'].apply(lambda x: f"LKR {x:,.2f}" if pd.notna(x) else "N/A")
    
    st.dataframe(
        display_comp[['Broker', 'Elevation', 'Catalogued', 'Sold', 'Sold_Pct', 'Avg_Price', 'Num_Lots']],
        use_container_width=True,
        hide_index=True
    )

def create_buyer_analysis_by_elevation(latest_df):
    """Create buyer analysis by elevation"""
    
    st.markdown("---")
    st.subheader("👥 Buyer Analysis by Elevation")
    
    sold_df = latest_df[latest_df['Status_Clean'] == 'sold']
    
    if sold_df.empty:
        st.info("No sold lots available for analysis")
        return
    
    # Top buyers selector
    top_buyers = sold_df.groupby('Buyer')['Total Value'].sum().nlargest(10).index.tolist()
    selected_buyer = st.selectbox("Select Buyer to Analyze", top_buyers, key="buyer_elev")
    
    buyer_df = sold_df[sold_df['Buyer'] == selected_buyer]
    
    # Buyer elevation analysis
    buyer_elev_analysis = buyer_df.groupby('Sub Elevation').apply(lambda x: pd.Series({
        'Quantity': x['Total Weight'].sum(),
        'Avg_Price': x['Price'].mean(),
        'Num_Lots': len(x),
        'Total_Value': x['Total Value'].sum()
    }), include_groups=False).reset_index()
    
    # Metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Elevations Purchased", len(buyer_elev_analysis))
        st.metric("Total Quantity", f"{buyer_df['Total Weight'].sum():,.0f} kg")
    with col2:
        st.metric("Total Value", f"LKR {buyer_df['Total Value'].sum():,.0f}")
        st.metric("Avg Price Paid", f"LKR {buyer_df['Price'].mean():,.2f}")
    with col3:
        fav_elev = buyer_elev_analysis.loc[buyer_elev_analysis['Quantity'].idxmax(), 'Sub Elevation']
        st.metric("Favorite Elevation", fav_elev)
    with col4:
        highest_price_elev = buyer_elev_analysis.loc[buyer_elev_analysis['Avg_Price'].idxmax(), 'Sub Elevation']
        st.metric("Highest Price Paid In", highest_price_elev)
    
    # Visualizations
    col1, col2 = st.columns(2)
    
    with col1:
        fig_qty_elev = px.pie(
            buyer_elev_analysis,
            values='Quantity',
            names='Sub Elevation',
            title=f'{selected_buyer} - Quantity Distribution by Elevation'
        )
        st.plotly_chart(fig_qty_elev, use_container_width=True)
    
    with col2:
        fig_price_elev = px.bar(
            buyer_elev_analysis,
            x='Sub Elevation',
            y='Avg_Price',
            title=f'{selected_buyer} - Average Price by Elevation',
            color='Avg_Price',
            color_continuous_scale='RdYlGn'
        )
        st.plotly_chart(fig_price_elev, use_container_width=True)
    
    # Grade preferences by elevation
    st.markdown("### 🎯 Grade Preferences by Elevation")
    
    buyer_grade_elev = buyer_df.groupby(['Sub Elevation', 'Grade']).agg({
        'Total Weight': 'sum',
        'Price': 'mean',
        'Lot No': 'count'
    }).reset_index()
    
    buyer_grade_elev.columns = ['Sub Elevation', 'Grade', 'Quantity', 'Avg_Price', 'Num_Lots']
    buyer_grade_elev = buyer_grade_elev.sort_values('Quantity', ascending=False).head(15)
    
    display_grade_elev = buyer_grade_elev.copy()
    display_grade_elev['Quantity'] = display_grade_elev['Quantity'].apply(lambda x: f"{x:,.0f} kg")
    display_grade_elev['Avg_Price'] = display_grade_elev['Avg_Price'].apply(lambda x: f"LKR {x:,.2f}")
    
    st.dataframe(display_grade_elev, use_container_width=True, hide_index=True)

# Export all dashboard functions
__all__ = [
    'create_elevation_performance_dashboard',
    'create_grade_performance_matrix',
    'create_broker_comparison_view',
    'create_buyer_analysis_by_elevation',
    'calculate_elevation_performance',
    'calculate_broker_elevation_performance'
]
