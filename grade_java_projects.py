import pandas as pd
import os
import subprocess
import shutil
import tempfile
import re

# TODO work with different sheets for submissions and results
# TODO section wise in the tabs in sheet
# TODO work in onedrive sheets, get details from tab -> sections {choose the tab name}

# --- CONFIGURATION ---
# IMPORTANT: Update these values in env to match your setup

# 1. Path to your Excel file in your local OneDrive folder
# ONEDRIVE_FILE_PATH = "C:/Users/YourUsername/OneDrive/StudentProjects.xlsx" # Example for Windows
ONEDRIVE_FILE_PATH = os.path.expanduser(os.getenv("ONEDRIVE_FILE_PATH")) # Example for Mac/Linux

INPUT_SHEET_NAME = "Submissions"
OUTPUT_SHEET_NAME = "Results"
STUDENT_NAME_COLUMN = "Student Name"
REPO_URL_COLUMN = "Repo URL"
PROGRAM_TIMEOUT = 15 # A shorter timeout is fine for simple programs

# NEW: Input to provide to the Java program's standard input.
# Use '\n' to simulate the user pressing the Enter key.
#TODO set the input based on the question for different assignments
PROGRAM_INPUT = "25.0\n0.0\n" # Provides 25.0 for the first prompt, 10.0 for the second.

# --- END OF CONFIGURATION ---

def find_file(directory, filename):
    for root, dirs, files in os.walk(directory):
        if filename in files:
            return os.path.join(root, filename)
    return None

def run_command(command, working_dir, input_data=None):
    """
    Runs a command, optionally providing input data to its stdin.
    Returns its status and output.
    """
    try:
        result = subprocess.run(
            command,
            cwd=working_dir,
            check=False,
            capture_output=True,
            text=True,
            timeout=PROGRAM_TIMEOUT,
            input=input_data # Pass input to the process
        )
        return result
    except FileNotFoundError:
        print(f"Error: Command '{command[0]}' not found. Is it in your system's PATH?")
        return None
    except subprocess.TimeoutExpired:
        print(f"  Error: Command timed out after {PROGRAM_TIMEOUT} seconds.")
        return "Timeout"

# look for main method in the code to find the main class
def detect_main_class(java_files):
    """Detect the main class with package if present."""
    for file in java_files:
        with open(file, "r", encoding="utf-8") as f:
            content = f.read()

            if re.search(r'public\s+static\s+void\s+main\s*\(\s*String\s*\[\]\s*\w+\)', content):
                class_name = os.path.splitext(os.path.basename(file))[0]

                # Check for package declaration
                package_match = re.search(r'package\s+([\w\.]+);', content)
                if package_match:
                    return package_match.group(1) + "." + class_name
                else:
                    return class_name
    return None


# def run_java_program(main_class, out_dir):
#     """Run the Java program interactively in terminal."""
#     if not main_class:
#         print("No class with main method found.")
#         return False

#     print(f"Running Java code (Main class: {main_class})...")
#     run_proc = subprocess.run(
#         ["java", "-cp", out_dir, main_class]  # interactive mode
#     )

#     return run_proc.returncode == 0

def process_student_repo(repo_url):
    """
    Clones, compiles, and runs a student's Java project.
    Returns a status string and any relevant error messages.
    """
    if not repo_url or pd.isna(repo_url):
        return "Absent", "No repository URL provided."

    with tempfile.TemporaryDirectory() as temp_dir:
        clone_path = os.path.join(temp_dir, "repo")
        print(f"  Cloning {repo_url}...")

        clone_command = ["git", "clone", repo_url, clone_path]
        clone_result = run_command(clone_command, temp_dir)

        if clone_result == "Timeout":
            return "Git Clone Error", "Git clone timed out."
        if clone_result is None or clone_result.returncode != 0:
            error_message = clone_result.stderr if clone_result else "Git command failed."
            return "Git Clone Error", f"Failed to clone repo.\n{error_message}"
        
        # Look for .java files
        java_files = []
        for root, dirs, files in os.walk(clone_path):
            for fname in files:
                if fname.endswith(".java"):
                    java_files.append(os.path.join(root, fname))
        if not java_files:
            return "Incomplete", "No .java files found in the repository."

        main_class = detect_main_class(java_files)
        print(f"  Detected main class: {main_class}")
        if main_class is None:
            return "Incomplete", "Could not find a class with a main method."

        # compile all java files from repo root so package structure is preserved
        print(f"  Compiling Java files ({len(java_files)} files)...")
        compile_command = ["javac"] + java_files
        compile_result = run_command(compile_command, clone_path)

        if compile_result == "Timeout":
            return "Compile Error", "Compiler timed out."
        if compile_result is None or compile_result.returncode != 0:
            error_message = compile_result.stderr if compile_result else "Javac command failed."
            return "Compile Error", f"Code did not compile.\n{error_message}"

        # run using classpath = clone_path and the detected main_class (may include package)
        print(f"  Running '{main_class}' with input...")
        run_command_list = ["java", "-cp", clone_path, main_class]
        # *** THE IMPORTANT CHANGE IS HERE ***
        run_result = run_command(run_command_list, clone_path, input_data=PROGRAM_INPUT)

        if run_result == "Timeout":
             return "Runtime Error", "Program timed out. It might have an infinite loop or requested more input than provided."
        if run_result is None or run_result.returncode != 0:
            error_message = run_result.stderr if run_result else "Java command failed."
            return "Runtime Error", f"Program ran but failed with an error (e.g., wrong input format).\n{error_message}"

        return "Completed", f"Program compiled and ran successfully.\nOutput:\n{run_result.stdout}"

def main():
    """Main function to drive the script."""
    print("--- Starting Student Project Grader ---")

    if not os.path.exists(ONEDRIVE_FILE_PATH):
        print(f"Error: The file '{ONEDRIVE_FILE_PATH}' was not found.")
        return

    try:
        df = pd.read_excel(ONEDRIVE_FILE_PATH, sheet_name=INPUT_SHEET_NAME)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    results = []
    for index, row in df.iterrows():
        student_name = row[STUDENT_NAME_COLUMN]
        repo_url = row.get(REPO_URL_COLUMN)

        print(f"\nProcessing {student_name}...")
        
        status, details = process_student_repo(repo_url)
        
        final_status = {
            "Absent": "Absent",
            "Git Clone Error": "Incomplete",
            "Compile Error": "Incomplete",
            "Runtime Error": "Partial Complete",
            "Completed": "Completed"
        }.get(status, "Unknown Error")

        print(f"  Status: {status} -> {final_status}")
        
        results.append({
            STUDENT_NAME_COLUMN: student_name,
            "Status": final_status,
            "Details": details
        })

    results_df = pd.DataFrame(results)

    print(f"\nWriting results to sheet '{OUTPUT_SHEET_NAME}' in '{ONEDRIVE_FILE_PATH}'...")
    try:
        with pd.ExcelWriter(ONEDRIVE_FILE_PATH, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            results_df.to_excel(writer, sheet_name=OUTPUT_SHEET_NAME, index=False)
        print("--- Script finished successfully! ---")
    except Exception as e:
        print(f"\nError writing to Excel file: {e}")

if __name__ == "__main__":
    main()