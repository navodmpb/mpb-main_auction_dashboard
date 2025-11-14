# MPBL Tea Auction Intelligence Dashboard
Streamlit-based analytics dashboard for MPBL tea auction data.

## Installation & Running Locally

### Prerequisites
- Python 3.9+
- pip3

### Setup
```bash
# Clone repository
git clone https://github.com/navodmpb/mpb-main_auction_dashboard.git
cd mpb-main_auction_dashboard

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run bid_dashboard_up.py
```

App will be available at `http://localhost:8501`

## Deployment to Streamlit Cloud

### Quick Deploy
1. Go to [Streamlit Cloud](https://share.streamlit.io/)
2. Click "New app"
3. Select repository: `navodmpb/mpb-main_auction_dashboard`
4. Select branch: `main`
5. Set main file path: `bid_dashboard_up.py`
6. Click "Deploy"

### Configuration
The app uses:
- Main entry point: `bid_dashboard_up.py`
- Configuration: `.streamlit/config.toml`
- Dependencies: `requirements.txt`

## Features
- Market Overview & MPBL Metrics
- Broker Performance Analysis
- Elevation & Category Performance
- Buyer Insights & Profiles
- Selling Mark Analysis
- Price Trends Analysis

### Advanced Features
- Per-broker PDF report generation (optimized, all grades included)
- Elevation-wise performance dashboards
- Conditional formatting in PDF tables (color-coded by performance)
- Professional UI (no emojis, clean headers)

## Data Sources
- Sales data: `sales_data/Sale_*.csv`
- Format: CSV with broker, grade, elevation, price, status columns

## Testing
Run headless PDF smoke test:
```bash
python3 test_generate_pdf_headless.py
```

## Support
For issues or questions, please open an issue in the GitHub repository.
