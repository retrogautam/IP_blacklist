import requests
print(requests.certs.where())
print(requests.get("https://www.google.com"), verify=False)