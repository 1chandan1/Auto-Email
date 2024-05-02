import os

import requests


def main():
    token = os.getenv("GITHUB_TOKEN")
    headers = {
        "Authorization": f"token {token}",
        "Content-Type": "application/octet-stream",
    }
    repo = "ChandanHans/NotaireCiclade"
    tag = os.getenv("RELEASE_TAG")
    file_path = "output/NotaireCiclade.exe"  # Adjust based on your output directory and file naming

    # Fetch the release by tag
    response = requests.get(
        f"https://api.github.com/repos/{repo}/releases/tags/{tag}", headers=headers
    )
    if response.status_code != 200:
        print(f"Failed to fetch release: {response.json()}")
        return
    release_info = response.json()
    release_id = release_info["id"]

    # Find if asset already exists
    existing_asset_id = None
    for asset in release_info.get("assets", []):
        if asset["name"] == os.path.basename(file_path):
            existing_asset_id = asset["id"]
            break

    # Delete old asset if exists
    if existing_asset_id:
        delete_response = requests.delete(
            f"https://api.github.com/repos/{repo}/releases/assets/{existing_asset_id}",
            headers={"Authorization": f"token {token}"},
        )
        if delete_response.status_code != 204:
            print(f"Failed to delete asset: {delete_response.text}")
            return
        print(f"Deleted existing asset: {existing_asset_id}")

    # Upload new asset
    upload_url = f"https://uploads.github.com/repos/{repo}/releases/{release_id}/assets?name={os.path.basename(file_path)}"
    with open(file_path, "rb") as f:
        upload_response = requests.post(upload_url, headers=headers, data=f.read())
        if upload_response.status_code not in range(200, 300):
            print(f"Failed to upload asset: {upload_response.json()}")
            return
        print("Asset uploaded successfully:", upload_response.json())


if __name__ == "__main__":
    main()
