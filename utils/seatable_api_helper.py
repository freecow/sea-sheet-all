from seatable_api import Base
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def get_seatable_config():
    """Load SeaTable configuration from environment variables."""
    return {
        'server_url': os.getenv('SEATABLE_SERVER_URL'),
        'api_token': os.getenv('SEATABLE_API_TOKEN')
    }

def get_seatable_base(config):
    """Get the SeaTable Base object."""
    base = Base(config['api_token'], config['server_url'])
    base.auth()
    base.use_api_gateway = False
    return base

def fetch_data_from_seatable(base, table_name, view_name):
    """Fetch data from SeaTable based on table and view name."""
    print(f"Fetching data from SeaTable view '{view_name}'...")
    rows = base.list_rows(table_name, view_name=view_name)
    if not rows:
        print(f"No data found for view '{view_name}'. Skipping...")
    return rows
