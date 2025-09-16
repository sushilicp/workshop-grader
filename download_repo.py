import requests
import csv
import os
import sys
from dotenv import load_dotenv

load_dotenv()

def fetch_assignment_grades(assignment_id, token):
    """
    Fetch grades for a given GitHub Classroom assignment.
    Returns a list of grade dicts.
    """
    base = "https://api.github.com"
    endpoint = f"/assignments/{assignment_id}/grades"
    url = base + endpoint
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28"
    }
    all_grades = []
    params = {
        "per_page": 100,
        "page": 1
    }
    while True:
        try:
            resp = requests.get(url, headers=headers, params=params)
            resp.raise_for_status()  # raises HTTPError for bad responses
            if resp.status_code != 200:
                print(f"Error {resp.status_code}")
                print(f"Incorrect assignment id: {assignment_id} provided or check your GITHUB_TOKEN permissions.")
                print(f"Please verify assignment id: {assignment_id} using cmd -> gh classroom assignment")
                sys.exit(1)
            page_data = resp.json()
            if not isinstance(page_data, list):
                print("Unexpected data format:", page_data)
                sys.exit(1)
            if not page_data:
                break
            all_grades.extend(page_data)

            if "Link" in resp.headers:
                links = resp.headers["Link"]
                # Simple check for 'rel="next"'
                if 'rel="next"' in links:
                    params["page"] += 1
                    continue
            break
        except ConnectionError as e:
            print(f"[ERROR] Could not connect to {url}. Reason: {e}")
            return []  # or None, depending on your logic
        except requests.Timeout:
            print(f"[ERROR] Request to {url} timed out.")
            return []
        except requests.RequestException as e:
            print(f"[ERROR] Request failed...Connect to Internet")
            return []   
    return all_grades

def save_grades_to_csv(grades, section, workshop_number):
    """
    Save the grades list (list of dicts) to a CSV file.
    """
    fieldnames = [
        "roster_identifier",
        "student_repository_name",
        "student_repository_url"
    ]
    try:
        output_path = os.getenv("CLASSROOM_DIR") + f"\\L2C{section}\\workshop_{workshop_number}.csv"
    
        
        with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for grade in grades:
                row = {fieldname: grade.get(fieldname, "") for fieldname in fieldnames}
                writer.writerow(row)
        print(f"Saved {len(grades)} grades to {output_path}")
    except FileNotFoundError:
        print(f"Error: Classroom directory: {output_path} path not found")
        print(f"Please create the section directory L2C{section} or check the CLASSROOM_DIR environment variable.")
        sys.exit(1)

def start_download(section, workshop):
    token = os.getenv("GITHUB_TOKEN")
    if not token:
        print("Please set the environment variable GITHUB_TOKEN to your GitHub PAT or fine-grained token.")
        sys.exit(1)

    # Take assignment ID and CSV path from user input
    assignment_id = input("Enter the assignment ID: ").strip()
    try:
        grades = fetch_assignment_grades(assignment_id, token)
        save_grades_to_csv(grades, section, workshop)
    except ConnectionError:
        print("connect to internet...")        
