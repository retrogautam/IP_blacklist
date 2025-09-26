import requests
import os
import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

def download_content(base_url, day, output_folder):
    url = f"{base_url}/{day}/"
    date = (datetime.now() - timedelta(days=1)).strftime("%Y/%m/%d")
    date = "2025/09/09" # Test string
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html5lib')
        # Find all article links and titles
        articles = soup.find_all('a', attrs={'rel': 'bookmark'})
        titles = [a.get_text(strip=True) for a in articles]
        urls = [a['href'] for a in articles]

        # Prepare SharePoint context once
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)

        dir_name = os.path.join(output_folder, date)
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)

        for idx, (vuln, sourceURL) in enumerate(zip(titles, urls)):
            try:
                newresponse = requests.get(sourceURL, headers=headers)
                newresponse.raise_for_status()
                soup_detail = BeautifulSoup(newresponse.content, 'html5lib')

                # Extract vulnerability content
                match = re.search(r"Live Threat Intelligence Feed([\s\S]+)Leave a Reply", soup_detail.text)
                if not match:
                    print(f"Could not extract content for {vuln}")
                    continue
                soupy = match.group(1)
                soupy = "".join([s for s in soupy.strip().splitlines(True) if s.strip("\r\n").strip()])
                soupy = re.sub('\t', '', soupy)
                soupy = re.sub(r"Posted on \w+\s\d+\,\s\d{4}", "", soupy)
                soupy = re.sub(r"Posted by Author[\w\s\,]+\n", "\n", soupy)
                soupy = re.sub(r"Qualys\sDetection\n[\s\S]+\nReferences", "References", soupy)
                soupy = re.sub(r"\d+\sthought\son[\s\S]+","",soupy)

                # Safe and unique filename
                safe_vuln = re.sub(r'[\\/*?:"<>|]', "_", vuln)
                output_file = os.path.join(dir_name, f"{safe_vuln}.txt")
                # If file exists, append index to avoid overwrite
                if os.path.exists(output_file):
                    output_file = os.path.join(dir_name, f"{safe_vuln}_{idx}.txt")
                with open(output_file, 'w', encoding='utf-8') as file:
                    file.write(soupy)

                # Upload to SharePoint
                with open(output_file, 'rb') as content_file:
                    file_name = os.path.basename(output_file)
                    target_folder.upload_file(file_name, content_file.read())
                    ctx.execute_query()

                print(f"Uploaded {output_file} to SharePoint folder {target_folder_url}")
            except Exception as e:
                print(f"Error processing {vuln}: {e}")

    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
    except requests.exceptions.RequestException as err:
        print(f"Error: {err}")

# Main function
base_url = "https://threatprotect.qualys.com"
day = (datetime.now() - timedelta(days=1)).strftime("%Y/%m/%d")
day = "2025/09/09" # Test string
output_folder = "C:\\Users\\aritra.gautam\\Downloads\\sample"
site_url = "https://internationalsosms.sharepoint.com/sites/GLBITSECCyberSecOps"
username = os.environ.get("SHAREPOINT_USERNAME", "cybersecurity.ops@internationalsos.com")
password = os.environ.get("SHAREPOINT_PASSWORD", "zOBk+)~?Ov5B")
target_folder_url = "/sites/GLBITSECCyberSecOps/Shared Documents/Threat Advisory/"
download_content(base_url, day, output_folder)