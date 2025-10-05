import subprocess
import json
import sys

def get_graph_token():
    """
    Get Microsoft Graph delegated access token using Azure CLI login.
    Requires: `az login` to have been done once manually.
    """
    try:
        # Run Azure CLI command to fetch token for Microsoft Graph
        result = subprocess.run(
            ["az", "account", "get-access-token", "--resource-type", "ms-graph"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True
        )

        # Parse the JSON output
        token_data = json.loads(result.stdout)
        access_token = token_data.get("accessToken")

        if not access_token:
            print("❌ No access token found in Azure CLI output.")
            sys.exit(1)

        print("✅ Successfully retrieved Microsoft Graph access token.")
        return access_token

    except subprocess.CalledProcessError as e:
        print("⚠️ Error while fetching token using Azure CLI.")
        print(e.stderr)
        print("➡️ Try running `az login` in your terminal first.")
        sys.exit(1)
    except json.JSONDecodeError:
        print("❌ Failed to parse Azure CLI token output.")
        sys.exit(1)

if __name__ == "__main__":
    token = get_graph_token()
    with open("access_token.txt", "w") as f:
        f.write(token)
    print("\n🔐 Access Token (truncated):", token[:100], "...")
