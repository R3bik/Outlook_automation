import os

# Configuration settings
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DISTRIBUTOR_DATA = os.path.join(BASE_DIR, 'data', 'distributor_data.csv')
REVENUE_FOLDER = os.path.join(BASE_DIR, 'data', 'revenue')

# Create folders if they don't exist
os.makedirs(REVENUE_FOLDER, exist_ok=True)