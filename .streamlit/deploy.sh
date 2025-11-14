#!/usr/bin/env bash
# Streamlit deployment script
# This is executed by Streamlit Cloud during deployment

set -e

echo "Installing dependencies..."
pip install -r requirements.txt

echo "Running validation tests..."
python3 -m py_compile bid_dashboard_up.py \
                      pdf_report_enhancements.py \
                      pdf_report_optimizer.py \
                      elevation_dashboard.py

echo "Streamlit app ready for deployment!"
