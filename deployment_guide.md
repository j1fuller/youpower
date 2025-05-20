# Deployment Guide: YouPower PG&E Tool on DigitalOcean

This guide explains how to deploy the YouPower PG&E Green Button Data tool on a DigitalOcean server while maintaining a hybrid approach that allows non-technical users to use Excel on the front end.

## Architecture Overview

The solution consists of:

1. **Client-side component**: A Windows executable (.exe) that users run to download GBD data and generate Excel files locally
2. **Server-side component**: A Python-based processing engine deployed on DigitalOcean for advanced data processing

## DigitalOcean Server Setup

### 1. Initial Server Configuration

Based on your existing DigitalOcean Droplet:

```
Image: Ubuntu LEMP on Ubuntu 22.04
Size: 2 vCPUs, 4GB RAM, 120GB Disk
Region: NYC1
```

### 2. Install Required Packages

SSH into your DigitalOcean server and install necessary dependencies:

```bash
# Update system packages
sudo apt-get update
sudo apt-get upgrade -y

# Install Python 3 and development tools
sudo apt-get install -y python3 python3.10-venv python3-pip python3-dev build-essential

# Install additional libraries
sudo apt-get install -y libxml2-dev libxslt1-dev
```

### 3. Set Up Python Environment

```bash
# Create a dedicated directory
mkdir -p /var/www/youpower_gbd
cd /var/www/youpower_gbd

# Create a virtual environment
python3 -m venv venv
source venv/bin/activate

# Install required Python packages
pip install Flask pandas openpyxl lxml requests selenium webdriver-manager
```

### 4. Deploy the Backend Processing Scripts

Upload the following files to your server:

- `gbd_parser.py` - XML parser for Green Button Data
- `pge_calculator.py` - PG&E rate calculation logic
- `api_server.py` - Simple Flask API to handle requests

### 5. Create the Flask API Server

Create a file named `api_server.py` with the following content:

```python
from flask import Flask, request, jsonify, send_file
import os
import pandas as pd
import tempfile
from gbd_parser import GBDXMLParser
from pge_calculator import PGECalculator

app = Flask(__name__)

@app.route('/api/process-gbd', methods=['POST'])
def process_gbd():
    """API endpoint to process GBD data and return Excel file."""
    try:
        # Check if file was included in the request
        if 'gbd_file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
            
        gbd_file = request.files['gbd_file']
        utility = request.form.get('utility', 'PG&E')
        
        # Save the uploaded file to a temporary location
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, gbd_file.filename)
        gbd_file.save(input_path)
        
        # Process the file based on its type
        if input_path.lower().endswith('.xml'):
            # Parse XML using the GBD parser
            parser = GBDXMLParser(input_path)
            df = parser.parse()
            account_info = parser.get_account_info()
        elif input_path.lower().endswith('.csv'):
            # Read CSV directly
            df = pd.read_csv(input_path)
            account_info = {}
        else:
            return jsonify({"error": "Unsupported file format"}), 400
            
        # Generate the Excel output
        output_path = os.path.join(temp_dir, 'processed_data.xlsx')
        calculator = PGECalculator(df, account_info)
        calculator.create_excel_output(output_path)
        
        # Return the Excel file
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f"{os.path.splitext(gbd_file.filename)[0]}_{utility}_processed.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    """Simple health check endpoint."""
    return jsonify({"status": "healthy"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
```

### 6. Configure Nginx as Reverse Proxy

```bash
# Create a new Nginx site configuration
sudo nano /etc/nginx/sites-available/youpower_gbd

# Add the following configuration
server {
    listen 80;
    server_name gbd.youpower.com;

    location / {
        proxy_pass http://localhost:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}

# Enable the site
sudo ln -s /etc/nginx/sites-available/youpower_gbd /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx
```

### 7. Set Up Service for Automatic Startup

```bash
# Create a systemd service file
sudo nano /etc/systemd/system/youpower_gbd.service

# Add the following content
[Unit]
Description=YouPower GBD Processing API
After=network.target

[Service]
User=www-data
Group=www-data
WorkingDirectory=/var/www/youpower_gbd
ExecStart=/var/www/youpower_gbd/venv/bin/python /var/www/youpower_gbd/api_server.py
Restart=always

[Install]
WantedBy=multi-user.target

# Enable and start the service
sudo systemctl enable youpower_gbd
sudo systemctl start youpower_gbd
```

## Client-Side Integration

Modify the desktop application to:

1. Process GBD data locally for simple calculations
2. Offer an option to submit data to the server for advanced processing

Add this function to the client application:

```python
def process_with_server(self, file_path, utility_provider):
    """Send the GBD file to the server for processing."""
    try:
        # Display status message
        QMessageBox.information(self, "Server Processing", "Sending file to server for advanced processing...")
        
        # Prepare the file for upload
        url = "https://gbd.youpower.com/api/process-gbd"
        files = {'gbd_file': open(file_path, 'rb')}
        data = {'utility': utility_provider}
        
        # Send the request
        response = requests.post(url, files=files, data=data)
        
        if response.status_code == 200:
            # Save the returned Excel file
            output_path = os.path.join(os.path.dirname(file_path), f"{os.path.basename(file_path)}_server_processed.xlsx")
            with open(output_path, 'wb') as f:
                f.write(response.content)
            return True, f"Server processing complete. File saved as: {output_path}"
        else:
            # Handle server error
            error_data = response.json()
            return False, f"Server processing failed: {error_data.get('error', 'Unknown error')}"
            
    except Exception as e:
        return False, f"Error with server processing: {e}"
```

## Security Considerations

1. **HTTPS**: Configure Let's Encrypt SSL for secure communication:

```bash
sudo apt-get install -y certbot python3-certbot-nginx
sudo certbot --nginx -d gbd.youpower.com
```

2. **API Authentication**: Add a simple API key mechanism:

```python
# In api_server.py
@app.route('/api/process-gbd', methods=['POST'])
def process_gbd():
    # Check API key
    api_key = request.headers.get('X-API-Key')
    if api_key != os.environ.get('YOUPOWER_API_KEY'):
        return jsonify({"error": "Unauthorized"}), 401
        
    # Rest of the function...
```

3. **Rate Limiting**: Add Flask-Limiter to prevent abuse:

```bash
pip install Flask-Limiter
```

## Monitoring and Maintenance

1. Set up basic monitoring using DigitalOcean's built-in monitoring tools
2. Configure log rotation for the Flask application
3. Create a backup script for the Python code and MongoDB data
4. Schedule regular updates for the server

## Conclusion

This deployment provides a hybrid approach that:

1. Allows non-technical users to work with Excel files locally
2. Leverages DigitalOcean server for advanced processing
3. Maintains a clean separation between frontend and backend components

The desktop application can work independently for basic functionality while having the option to connect to the server for more complex operations, providing the best of both worlds.
