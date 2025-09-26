import requests
from urllib3.exceptions import InsecureRequestWarning
import xml.etree.ElementTree as ET
from dateutil import parser
from datetime import datetime
import os
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import urllib3

# Monkey-patch requests.Session to always use verify=False
os.environ['O365_DISABLE_SSL'] = 'true'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
old_request = requests.Session.request
def new_request(self, *args, **kwargs):
    kwargs['verify'] = False
    return old_request(self, *args, **kwargs)
requests.Session.request = new_request

USERNAME = "ntern3ks"
PASSWORD = "-rechi=*D3SPO0a04C3i?1be"
API_URL = "https://qualysguard.qg3.apps.qualys.com/api/2.0/fo/report/"
OUTPUT_DIR = r"C:\Users\aritra.gautam\OneDrive - International SOS\Documents\Qualys\Reports"
SHAREPOINT_USERNAME = "cybersecurity.ops@internationalsos.com"
SHAREPOINT_PASSWORD = "zOBk+)~?Ov5B"
SHAREPOINT_SITE_URL = "https://internationalsosms.sharepoint.com/sites/GLBITSECCyberSecOps"
SHAREPOINT_PATHS = {
    "Overall_Infra_Report": "/sites/GLBITSECCyberSecOps/Shared Documents/General/PowerBI/Raw Vulnerability Reports/Server Vulnerability Report",
    "Overall_Endpoint_Report": "/sites/GLBITSECCyberSecOps/Shared Documents/General/PowerBI/Raw Vulnerability Reports/Endpoint Vulnerability Report"
}

def get_latest_report_ids(report_names):
    headers = {"Accept": "application/xml", "X-Requested-With": "QualysAPI"}
    data = {"action": "list"}
    try:
        resp = requests.post(API_URL, auth=(USERNAME, PASSWORD), headers=headers, data=data)
        resp.raise_for_status()
        root = ET.fromstring(resp.text)
        all_titles = [r.findtext("TITLE") for r in root.findall(".//REPORT")]
        print("All report titles found:", all_titles)
        reports = [
            {
                "title": r.findtext("TITLE").strip(),
                "id": r.findtext("ID").strip(),
                "launch_date": r.findtext("LAUNCH_DATETIME").strip()
            }
            for r in root.findall(".//REPORT")
            if r.findtext("TITLE") and r.findtext("ID") and r.findtext("LAUNCH_DATETIME")
               and r.findtext("TITLE").strip() in report_names
               and r.findtext("STATUS/STATE") and r.findtext("STATUS/STATE").strip().lower() == "finished"
        ]
        print("Reports found after status filter:", reports)
        reports.sort(key=lambda r: parser.parse(r["launch_date"]), reverse=True)
        latest = {name: next((r["id"] for r in reports if r["title"] == name), None) for name in report_names}
        return latest
    except Exception as e:
        print(f"Error fetching report IDs: {e}")
        return {name: None for name in report_names}

def download_report_csv(report_id, filename):
    headers = {"Accept": "application/csv", "X-Requested-With": "QualysAPI"}
    data = {
        "action": "fetch",
        "id": report_id
    }
    output_path = os.path.join(OUTPUT_DIR, filename)
    try:
        resp = requests.post(API_URL, auth=(USERNAME, PASSWORD), headers=headers, data=data, stream=True)
        resp.raise_for_status()
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(output_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        print(f"Report {report_id} downloaded as {output_path}")
    except Exception as e:
        print(f"Error downloading report {report_id}: {e}")

def convert_csv_to_xlsx_and_rename(csv_path):
    if not os.path.exists(csv_path):
        print(f"CSV file not found: {csv_path}")
        return
    base_filename = os.path.splitext(os.path.basename(csv_path))[0]
    today_str = datetime.today().strftime("%Y%m%d")
    new_filename = f"Scan_Report_{base_filename}_({today_str}).xlsx"
    xlsx_path = os.path.join(OUTPUT_DIR, new_filename)
    # Determine sheet name
    if "Infra" in base_filename:
        sheet_name = "Scan_Report_Overall_Infra_Repor"
    elif "Endpoint" in base_filename:
        sheet_name = "Scan_Report_Overall_Endpoint_Re"
    else:
        sheet_name = "Scan_Report"

    try:
        date_columns = ['First Detected', 'Last Detected', 'Date Last Fixed', 'First Reopened', 'Last Reopened']
        df = pd.read_csv(csv_path, skiprows=10, low_memory=False, dtype=str)
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], format='%m/%d/%Y %H:%M:%S', errors='coerce')
        with pd.ExcelWriter(xlsx_path, date_format='mm/dd/yyyy hh:mm:ss') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        print(f"Converted {csv_path} to {xlsx_path} with sheet name '{sheet_name}'")
    except Exception as e:
        print(f"Error converting {csv_path} to XLSX: {e}")

def upload_xlsx_to_sharepoint(local_xlsx_path, report_type):
    target_folder = SHAREPOINT_PATHS.get(report_type)
    if not target_folder:
        print(f"No SharePoint path configured for report type: {report_type}")
        return

    ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(
        UserCredential(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD)
    )
    with open(local_xlsx_path, "rb") as content_file:
        file_name = os.path.basename(local_xlsx_path)
        try:
            ctx.web.get_folder_by_server_relative_url(target_folder).upload_file(file_name, content_file.read()).execute_query()
            print(f"Uploaded {file_name} to SharePoint folder: {target_folder}")
        except Exception as e:
            import traceback
            print(f"Error uploading {file_name} to SharePoint: {e}")
            traceback.print_exc()

def ensure_archive_folder(ctx, folder):
    """Ensure the Archive subfolder exists in the given SharePoint folder."""
    archive_url = f"{folder}/Archive"
    try:
        archive_folder = ctx.web.get_folder_by_server_relative_url(archive_url)
        ctx.load(archive_folder)
        ctx.execute_query()
        print(f"Archive folder exists: {archive_url}")
    except Exception:
        parent_folder = ctx.web.get_folder_by_server_relative_url(folder)
        parent_folder.folders.add('Archive').execute_query()
        print(f"Created Archive in {folder}")
    return archive_url

def cleanup_files():
    """Remove local CSVs and move the oldest SharePoint file in each path to Archive."""
    # Remove all .csv files from OUTPUT_DIR
    for fname in os.listdir(OUTPUT_DIR):
        if fname.lower().endswith('.csv'):
            try:
                os.remove(os.path.join(OUTPUT_DIR, fname))
                print(f"Removed CSV: {fname}")
            except Exception as e:
                print(f"Error removing {fname}: {e}")

    # Move oldest file in each SharePoint directory to Archive
    ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(
        UserCredential(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD)
    )
    for key, folder in SHAREPOINT_PATHS.items():
        try:
            folder_obj = ctx.web.get_folder_by_server_relative_url(folder)
            files = folder_obj.files
            ctx.load(files)
            ctx.execute_query()
            file_list = [
                f for f in files
                if not f.properties['Name'].startswith('.')
            ]

            # Parse creation dates
            for f in file_list:
                time_created = f.properties.get('TimeCreated')
                if isinstance(time_created, str):
                    try:
                        f._created_dt = datetime.fromisoformat(time_created.replace('Z', '+00:00'))
                    except Exception as ex:
                        print(f"Failed to parse TimeCreated for {f.properties['Name']}: {ex}")
                elif isinstance(time_created, datetime):
                    f._created_dt = time_created
                else:
                    print(f"Warning: TimeCreated is not a string or datetime for file {f.properties['Name']}")

            # Only keep files with a valid _created_dt
            file_list = [f for f in file_list if hasattr(f, "_created_dt")]
            if not file_list:
                print(f"No files with valid TimeCreated in {folder}")
                continue

            # Find the oldest file
            oldest_file = min(file_list, key=lambda f: f._created_dt)
            oldest_file_name = oldest_file.properties['Name']
            print(f"Oldest file in {folder}: {oldest_file_name}")

            # Ensure Archive exists
            archive_url = ensure_archive_folder(ctx, folder)

            # Move the file (use full server-relative URL and overwrite flag)
            dst_url = f"{archive_url}"
            print(f"Moving file: {oldest_file_name} to {dst_url}")
            try:
                oldest_file.moveto(dst_url, 1).execute_query()
                print(f"Moved {oldest_file_name} to {archive_url}")
            except AttributeError:
                print("moveto() not available, falling back to copyto and delete.")
                oldest_file.copyto(dst_url, True).execute_query()
                oldest_file.delete_object().execute_query()
                print(f"Copied and deleted {oldest_file_name} to {archive_url}")
        except Exception as e:
            import traceback
            print(f"Error processing {folder}: {e}")
            traceback.print_exc()

if __name__ == "__main__":
    report_names = ["Overall Endpoint Report", "Overall Infra Report"]
    headers = {"Accept": "application/xml", "X-Requested-With": "QualysAPI"}
    data = {"action": "list"}
    resp = requests.post(API_URL, auth=(USERNAME, PASSWORD), headers=headers, data=data)
    latest_ids = get_latest_report_ids(report_names)
    print("Latest report IDs:", latest_ids)
    for name, report_id in latest_ids.items():
        if report_id:
            filename = f"{name.replace(' ', '_').replace('/', '_')}.csv"
            download_report_csv(report_id, filename)
            csv_path = os.path.join(OUTPUT_DIR, filename)
            convert_csv_to_xlsx_and_rename(csv_path)
            base_filename = os.path.splitext(os.path.basename(csv_path))[0]
            today_str = datetime.today().strftime("%Y%m%d")
            xlsx_filename = f"Scan_Report_{base_filename}_({today_str}).xlsx"
            xlsx_path = os.path.join(OUTPUT_DIR, xlsx_filename)
            report_key = name.replace(" ", "_").replace("/", "_")
            print(f"Looking for XLSX file: {xlsx_path}")
            if os.path.exists(xlsx_path):
                upload_xlsx_to_sharepoint(xlsx_path, report_key)
            else:
                print(f"XLSX file not found for upload: {xlsx_path}")
        else:
            print(f"No finished report")
    # cleanup_files()