import sys
import importlib

# Define the root path
email_root = r'C:\Users\wb539289\python_projects\wbhrdata\emails'

# Import the variables
if email_root not in sys.path:
    sys.path.append(email_root)
config = importlib.import_module('email_config')
