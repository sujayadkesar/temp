import os
import shutil
import sys
import glob

# Define paths to artifacts (based on the provided drive letter)
def get_artifact_paths(drive_letter):
    return {
        "Registry": {
            "BAM": os.path.join(drive_letter, "Windows", "System32", "config", "BAM"),
            "ShimCache": os.path.join(drive_letter, "Windows", "AppCompat", "Programs", "Amcache.hve"),
            "SYSTEM": os.path.join(drive_letter, "Windows", "System32", "config", "SYSTEM"),
            "SAM": os.path.join(drive_letter, "Windows", "System32", "config", "SAM"),
            "SECURITY": os.path.join(drive_letter, "Windows", "System32", "config", "SECURITY"),
            "SOFTWARE": os.path.join(drive_letter, "Windows", "System32", "config", "SOFTWARE")
        },
        "User_Profiles": os.path.join(drive_letter, "Users"),
        "LNK_files": os.path.join(drive_letter, "Users", "*", "AppData", "Roaming", "Microsoft", "Windows", "Recent", "*.lnk"),
        "Prefetch_files": os.path.join(drive_letter, "Windows", "Prefetch", "*.pf"),
        "Event_Logs": os.path.join(drive_letter, "Windows", "System32", "winevt", "Logs", "*.evtx"),
        "Browser_Artifacts": [
            os.path.join(drive_letter, "Users", "*", "AppData", "Local", "Microsoft", "Edge", "User Data", "Default", "History"),
            os.path.join(drive_letter, "Users", "*", "AppData", "Local", "Google", "Chrome", "User Data", "Default", "History"),
            os.path.join(drive_letter, "Users", "*", "AppData", "Roaming", "Microsoft", "Internet Explorer", "UserData", "index.dat"),
        ],
        "Recycle_Bin": os.path.join(drive_letter, "$Recycle.Bin", "*"),
        "Scheduled_Tasks": os.path.join(drive_letter, "Windows", "System32", "Tasks", "*")
    }

# Function to copy files or directories into individual subfolders
def copy_artifacts(artifacts, destination_folder):
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    for category, paths in artifacts.items():
        # Create subdirectory for each category
        category_folder = os.path.join(destination_folder, category)
        if not os.path.exists(category_folder):
            os.makedirs(category_folder)

        if isinstance(paths, dict):  # For nested dictionaries like Registry
            for name, path in paths.items():
                try:
                    if os.path.exists(path):
                        shutil.copy2(path, category_folder)
                        print(f"Copied {path} to {category_folder}")
                    else:
                        print(f"{name} not found at {path}")
                except Exception as e:
                    print(f"Failed to copy {name}: {e}")
        elif isinstance(paths, list):  # For lists like Browser_Artifacts
            for subpath in paths:
                try:
                    for match in glob.glob(subpath):
                        shutil.copy2(match, category_folder)
                        print(f"Copied {match} to {category_folder}")
                except Exception as e:
                    print(f"Failed to copy {subpath}: {e}")
        else:  # For single paths, including those with wildcards
            try:
                if "*" in paths:  # Handling wildcards in paths (like user directories or Recycle Bin)
                    for match in glob.glob(paths):
                        if os.path.isfile(match):
                            shutil.copy2(match, category_folder)
                            print(f"Copied {match} to {category_folder}")
                elif os.path.isdir(paths):  # If it's a directory, copy relevant files
                    for root, dirs, files in os.walk(paths):
                        for file in files:
                            src_file = os.path.join(root, file)
                            shutil.copy2(src_file, category_folder)
                            print(f"Copied {src_file} to {category_folder}")
                elif os.path.exists(paths):
                    shutil.copy2(paths, category_folder)
                    print(f"Copied {paths} to {category_folder}")
                else:
                    print(f"{category} not found at {paths}")
            except Exception as e:
                print(f"Failed to copy {category}: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python collect_windows_artifacts.py <drive_letter> <destination_folder>")
        sys.exit(1)

    drive_letter = sys.argv[1].rstrip(":") + ":"
    destination_folder = sys.argv[2]

    # Get paths to artifacts
    artifacts = get_artifact_paths(drive_letter)

    # Copy artifacts to the destination
    copy_artifacts(artifacts, destination_folder)
