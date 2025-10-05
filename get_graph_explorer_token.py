import json
import requests
import browser_cookie3

def get_graph_explorer_token():
    """
    Extracts Microsoft Graph Explorer token from your logged-in browser session.
    Works with Chrome/Edge (local use only).
    """
    try:
        # Try Chrome cookies first
        cj = browser_cookie3.chrome(domain_name=".microsoft.com")
    except Exception:
        # Fallback to Edge
        cj = browser_cookie3.edge(domain_name=".microsoft.com")

    # Check cookie names for Graph Explorer auth tokens
    tokens = {}
    for cookie in cj:
        if "graph" in cookie.domain and ("access_token" in cookie.name or "microsoft" in cookie.name):
            tokens[cookie.name] = cookie.value

    # You may not see token directly — Graph Explorer uses localStorage
    # Try hitting Graph Explorer API to extract the current token
    resp = requests.get(
        "https://developer.microsoft.com/en-us/graph/graph-explorer/api/proxy",
        cookies=cj
    )

    if resp.status_code == 200 and "accessToken" in resp.text:
        # Attempt to parse token if available
        try:
            token_data = resp.json()
            token = token_data.get("accessToken")
        except Exception:
            # Fallback to raw search
            start = resp.text.find("accessToken")
            if start != -1:
                token = resp.text[start:].split('"')[2]
            else:
                token = None
    else:
        token = None

    if not token:
        print("❌ Could not find Graph Explorer token. Please ensure you're logged into Graph Explorer in your browser.")
        return None

    # Save token for use in your scripts
    with open("access_token.txt", "w") as f:
        f.write(token)
    print("✅ Token extracted and saved to access_token.txt")

    return token


if __name__ == "__main__":
    get_graph_explorer_token()
