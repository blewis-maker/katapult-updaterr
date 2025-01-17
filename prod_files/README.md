# KatapultPro Data Processing

This script processes data from the KatapultPro API, generating reports and geographic data files.

## Features

- Fetches job data from KatapultPro API
- Processes nodes, connections, and anchors
- Generates shapefiles and GeoPackage files
- Creates Excel reports with job status
- Uploads data to SharePoint
- Sends email notifications with reports

## Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in your credentials:
   ```bash
   cp .env.example .env
   ```

## Environment Variables

The following environment variables are required:

- `EMAIL_USER`: Email address for notifications
- `EMAIL_PASSWORD`: Email password
- `AZURE_CLIENT_ID`: Azure AD client ID
- `AZURE_CLIENT_SECRET`: Azure AD client secret
- `AZURE_TENANT_ID`: Azure AD tenant ID

## Usage

Run the script with:
```bash
python main.py
```

The script will:
1. Fetch all jobs from KatapultPro
2. Process the job data
3. Generate geographic data files
4. Create an Excel report
5. Upload data to SharePoint
6. Send email notifications

## Output Files

- `*.gpkg`: GeoPackage files containing geographic data
- `*.shp`, `*.shx`, `*.dbf`, `*.prj`: Shapefile components
- Excel reports with job status
- Log files with processing details

## Notes

- Sensitive files (`.env`, tokens, logs) are excluded from git
- SharePoint integration requires valid Azure AD credentials
- Email notifications require valid email credentials 