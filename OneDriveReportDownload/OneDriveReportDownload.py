import os
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

username = "cybersecurity.ops@internationalsos.com"
password = "zOBk+)~?Ov5B"
site_url = "https://internationalsosms-my.sharepoint.com/personal/aritra_gautam_internationalsos_com"
folder_url = "/personal/aritra_gautam_internationalsos_com/Documents/Documents/Crowdstrike Falcon/Falcon Reports"
local_folder = "C:\\Users\\aritra.gautam\\Downloads\\sample\\"

os.makedirs(local_folder, exist_ok=True)

ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
folder = ctx.web.get_folder_by_server_relative_url(folder_url)
files = folder.files
ctx.load(files)
ctx.execute_query()

if not files:
    print("No files found in the folder.")
else:
    for file in files:
        local_path = os.path.join(local_folder, file.properties["Name"])
        file_response = file.open_binary(ctx, file.properties["ServerRelativeUrl"])
        with open(local_path, "wb") as f:
            f.write(file_response.content)
        print(f"Downloaded {file.properties['Name']} to {local_path}")