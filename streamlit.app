# Streamlit Configuration for Cloud Deployment
# This file tells Streamlit Cloud which file is the main app entry point

# Main app entry point
client.mainPath = "bid_dashboard_up.py"

# Streamlit Cloud settings
[client]
showErrorDetails = true
logger.level = "info"

[server]
headless = true
enableXsrfProtection = true
enableCORS = true
runOnSave = true
