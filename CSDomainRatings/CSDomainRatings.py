import requests
import json
import csv
import time
import re

# --- CONFIGURATION ---
ENDPOINT = "https://api.us-2.crowdstrike.com/identity-protection/combined/graphql/v1"
OAUTH_URL = "https://api.us-2.crowdstrike.com/oauth2/token"
CLIENT_ID = "3c5f15835a3042e9a97eae41658f1cd7"
CLIENT_SECRET = "3VE91u46x0UOlGr5mZYfI2gin8WPN7CsBzkMeSXy"
DOMAINS_FILE = r"C:\Users\aritra.gautam\OneDrive - International SOS\Scripts\cs_domain_list.txt"
EXCLUDE_LOW = True
OUT_PREFIX = "out"
TIMEOUT = 30
RETRIES = 2

SELECTION_SET = """
  overallScore
  overallScoreLevel
  assessmentFactors {
    riskFactorType
    severity
    likelihood
  }
"""

def fetch_crowdstrike_token(client_id, client_secret, oauth_url):
    resp = requests.post(
        oauth_url,
        headers={"accept": "application/json", "Content-Type": "application/x-www-form-urlencoded"},
        data={"client_id": client_id, "client_secret": client_secret}
    )
    resp.raise_for_status()
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError(f"Token not found in response: {resp.text}")
    return token

def slugify_alias(s):
    s = re.sub(r"[^a-z0-9]+", "_", s.strip().lower())
    s = re.sub(r"^_+|_+$", "", s) or "domain"
    return f"d_{s}" if s[0].isdigit() else s

def build_query_with_aliases(domains):
    alias_map, body_parts, used = {}, [], set()
    for domain in domains:
        base_alias = slugify_alias(domain)
        alias = base_alias
        suffix = 2
        while alias in used:
            alias = f"{base_alias}_{suffix}"
            suffix += 1
        used.add(alias)
        alias_map[alias] = domain
        body_parts.append(f'{alias}: securityAssessment(domain: "{domain}") {{ {SELECTION_SET} }}')
    return "query {\n  " + "\n  ".join(body_parts) + "\n}", alias_map

def post_graphql(endpoint, query, headers, timeout=30, retries=2):
    for attempt in range(retries + 1):
        try:
            resp = requests.post(endpoint, json={"query": query}, headers=headers, timeout=timeout)
            resp.raise_for_status()
            payload = resp.json()
            if payload.get("errors"):
                raise RuntimeError(json.dumps(payload["errors"], indent=2))
            return payload.get("data", {})
        except Exception as e:
            if attempt < retries:
                time.sleep(1.5 * (attempt + 1))
            else:
                raise RuntimeError(f"GraphQL request failed after retries: {e}")

def flatten_assessments(data, alias_map, exclude_low=False):
    assessments_summary, assessment_factors = [], []
    for alias, assessment in data.items():
        domain = alias_map.get(alias, alias)
        if not assessment:
            assessments_summary.append({
                "domain": domain, "overallScore": None, "overallScoreLevel": None, "note": "No data (null)"
            })
            continue
        assessments_summary.append({
            "domain": domain,
            "overallScore": assessment.get("overallScore"),
            "overallScoreLevel": assessment.get("overallScoreLevel"),
        })
        for f in assessment.get("assessmentFactors") or []:
            sev = (f.get("severity") or "").upper()
            if exclude_low and sev == "LOW":
                continue
            assessment_factors.append({
                "domain": domain,
                "riskFactorType": f.get("riskFactorType"),
                "likelihood": f.get("likelihood"),
                "severity": sev,
            })
    return assessments_summary, assessment_factors

def load_domains_from_file(path):
    with open(path, encoding="utf-8") as f:
        domains = [line.strip() for line in f if line.strip() and not line.lstrip().startswith("#")]
    if not domains:
        raise SystemExit("No domains provided in file.")
    return domains

def write_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)

def write_csv(path, rows, fieldnames):
    with open(path, "w", encoding="utf-8", newline="") as f:
        csv.DictWriter(f, fieldnames=fieldnames).writeheader()
        csv.DictWriter(f, fieldnames=fieldnames).writerows(rows)

def main():
    print("Fetching OAuth2 token...")
    token = fetch_crowdstrike_token(CLIENT_ID, CLIENT_SECRET, OAUTH_URL)
    print("Token acquired.")

    print(f"Loading domains from {DOMAINS_FILE} ...")
    domains = load_domains_from_file(DOMAINS_FILE)
    print(f"Loaded {len(domains)} domains.")

    query, alias_map = build_query_with_aliases(domains)
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {token}"}

    print("Querying GraphQL endpoint...")
    data = post_graphql(ENDPOINT, query, headers, timeout=TIMEOUT, retries=RETRIES)

    assessments_summary, assessment_factors = flatten_assessments(data, alias_map, exclude_low=EXCLUDE_LOW)

    summary_path = r"C:\Users\aritra.gautam\OneDrive - International SOS\Documents\Crowdstrike Falcon\Falcon Domain Ratings\assessments_summary.csv"
    factors_path = r"C:\Users\aritra.gautam\OneDrive - International SOS\Documents\Crowdstrike Falcon\Falcon Domain RiskFactors\assessment_factors.csv"

    write_csv(summary_path, assessments_summary, ["domain", "overallScore", "overallScoreLevel"])
    write_csv(factors_path, assessment_factors, ["domain", "riskFactorType", "likelihood", "severity"])

    print(f"\nQueried {len(domains)} domain(s). Files written:")
    print(f"  - {summary_path}")
    print(f"  - {factors_path}")
    if EXCLUDE_LOW:
        print("  (LOW severity factors excluded)")
    print("\nSample summary:")
    for row in assessments_summary[:10]:
        print(f'  {row["domain"]}: score={row["overallScore"]}, level={row["overallScoreLevel"]}')

if __name__ == "__main__":
    main()