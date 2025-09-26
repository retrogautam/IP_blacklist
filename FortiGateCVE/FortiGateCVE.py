import requests
from bs4 import BeautifulSoup

# Example list of IR numbers (Fortinet format: FG-IR-YY-XXX)
ir_list = [
    "FG-IR-24-196",
    "FG-IR-23-112",
    "FG-IR-22-455"
]

def get_fortinet_ir_severity(ir_number):
    url = f"https://fortiguard.fortinet.com/psirt/{ir_number}"
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    try:
        response = requests.get(url, headers=headers, timeout=10)
        print(response.text)  # Add this after fetching the response
        if response.status_code == 404:
            return "Not Found"

        soup = BeautifulSoup(response.text, 'html.parser')

        # Try to find any element containing "Severity"
        for tag in soup.find_all(text=lambda t: "Severity" in t):
            parent = tag.parent
            # If the next sibling is a tag, get its text
            if parent and parent.next_sibling:
                next_sib = parent.next_sibling
                if hasattr(next_sib, 'text'):
                    value = next_sib.text.strip()
                    if value:
                        return value
            # If the parent is a <th>, try to get the next <td>
            if parent.name == "th":
                td = parent.find_next_sibling("td")
                if td:
                    return td.text.strip()

        return "Unknown"

    except Exception as e:
        return f"Error: {e}"

if __name__ == "__main__":
    print(f"{'IR Number':<20}Severity")
    print("="*45)
    for ir in ir_list:
        severity = get_fortinet_ir_severity(ir)
        print(f"{ir:<20}{severity}")