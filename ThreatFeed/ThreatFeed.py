import requests
import os
import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

def safe_filename(title, idx, dir_name):
    # Sanitize and ensure uniqueness
    safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
    filename = f"{safe_title}.txt"
    filepath = os.path.join(dir_name, filename)
    if os.path.exists(filepath):
        filename = f"{safe_title}_{idx}.txt"
        filepath = os.path.join(dir_name, filename)
    return filepath

def extract_vuln_content(soup_text):
    match = re.search(r"Live Threat Intelligence Feed([\s\S]+)Leave a Reply", soup_text)
    if not match:
        return None
    content = match.group(1)
    content = "".join([s for s in content.strip().splitlines(True) if s.strip("\r\n").strip()])
    content = re.sub('\t', '', content)
    content = re.sub(r"Posted on \w+\s\d+\,\s\d{4}", "", content)
    content = re.sub(r"Posted by Author[\w\s\,]+\n", "\n", content)
    content = re.sub(r"Qualys\sDetection\n[\s\S]+\nReferences", "References", content)
    content = re.sub(r"\d+\sthought\son[\s\S]+","", content)
    return content

def download_content(base_url, day, output_folder):
    url = f"{base_url}/{day}/"
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html5lib')
        articles = soup.find_all('a', attrs={'rel': 'bookmark'})
        titles = [a.get_text(strip=True) for a in articles]
        urls = [a['href'] for a in articles]

        dir_name = os.path.join(output_folder, day)
        os.makedirs(dir_name, exist_ok=True)

        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)

        for idx, (title, source_url) in enumerate(zip(titles, urls)):
            try:
                detail_resp = requests.get(source_url, headers=headers)
                detail_resp.raise_for_status()
                content = extract_vuln_content(BeautifulSoup(detail_resp.content, 'html5lib').text)
                if not content:
                    print(f"Could not extract content for {title}")
                    continue
                output_file = safe_filename(title, idx, dir_name)
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                with open(output_file, 'rb') as f:
                    target_folder.upload_file(os.path.basename(output_file), f.read())
                    ctx.execute_query()
                print(f"Uploaded {output_file} to SharePoint folder {target_folder_url}")
            except Exception as e:
                print(f"Error processing {title}: {e}")
    except requests.exceptions.RequestException as err:
        print(f"Request Error: {err}")

# Main function
base_url = "https://threatprotect.qualys.com"
day = (datetime.now() - timedelta(days=1)).strftime("%Y/%m/%d")
output_folder = "C:\\Users\\aritra.gautam\\Downloads\\sample"
site_url = "https://internationalsosms.sharepoint.com/sites/GLBITSECCyberSecOps"
username = os.environ.get("SHAREPOINT_USERNAME", "cybersecurity.ops@internationalsos.com")
password = os.environ.get("SHAREPOINT_PASSWORD", "zOBk+)~?Ov5B")
target_folder_url = "/sites/GLBITSECCyberSecOps/Shared Documents/Threat Advisory/"
download_content(base_url, day, output_folder)