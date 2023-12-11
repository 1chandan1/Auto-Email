import requests
import os
import sys

def get_release_id(repo, tag, token):
    url = f"https://api.github.com/repos/{repo}/releases/tags/{tag}"
    headers = {'Authorization': f'token {token}'}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()['id']
    else:
        print(f"Error fetching release: {response.status_code}")
        return None

def delete_release_assets(repo, release_id, token):
    url = f"https://api.github.com/repos/{repo}/releases/{release_id}/assets"
    headers = {'Authorization': f'token {token}'}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        assets = response.json()
        for asset in assets:
            delete_url = f"https://api.github.com/repos/{repo}/releases/assets/{asset['id']}"
            delete_response = requests.delete(delete_url, headers=headers)
            if delete_response.status_code == 204:
                print(f"Deleted asset {asset['id']}")
            else:
                print(f"Error deleting asset {asset['id']}: {delete_response.status_code}")
    else:
        print(f"Error fetching assets: {response.status_code}")

def main():
    repo = sys.argv[1]
    tag = sys.argv[2]
    token = os.environ['GITHUB_TOKEN']

    release_id = get_release_id(repo, tag, token)
    if release_id:
        delete_release_assets(repo, release_id, token)
    else:
        print("Release not found.")

if __name__ == "__main__":
    main()
