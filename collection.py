# import os
# import shutil
# import sys
# import glob
# import ctypes

# # Function to check if script is run with administrator privileges
# def is_admin():
#     try:
#         return ctypes.windll.shell32.IsUserAnAdmin()
#     except:
#         return False

# # Define paths to artifacts (based on the provided base path, which could be a drive letter or a folder)
# def get_artifact_paths(base_path):
#     return {
#         "Registry": {
#             "BAM": os.path.join(base_path, "Windows", "System32", "config", "BAM"),
#             "ShimCache": os.path.join(base_path, "Windows", "AppCompat", "Programs", "Amcache.hve"),
#             "SYSTEM": os.path.join(base_path, "Windows", "System32", "config", "SYSTEM"),
#             "SAM": os.path.join(base_path, "Windows", "System32", "config", "SAM"),
#             "SECURITY": os.path.join(base_path, "Windows", "System32", "config", "SECURITY"),
#             "SOFTWARE": os.path.join(base_path, "Windows", "System32", "config", "SOFTWARE")
#         },
#         "LNK_files": os.path.join(base_path, "Users", "*", "AppData", "Roaming", "Microsoft", "Windows", "Recent", "*.lnk"),
#         "Prefetch_files": os.path.join(base_path, "Windows", "Prefetch", "*.pf"),
#         "Event_Logs": os.path.join(base_path, "Windows", "System32", "winevt", "Logs", "*.evtx"),
#         "Browser_Artifacts": [
#             {"browser": "Edge", "path": os.path.join(base_path, "Users", "*", "AppData", "Local", "Microsoft", "Edge", "User Data", "Default", "History")},
#             {"browser": "Chrome", "path": os.path.join(base_path, "Users", "*", "AppData", "Local", "Google", "Chrome", "User Data", "Default", "History")},
#             {"browser": "IE", "path": os.path.join(base_path, "Users", "*", "AppData", "Roaming", "Microsoft", "Internet Explorer", "UserData", "index.dat")}
#         ],
#         "Recycle_Bin": os.path.join(base_path, "$Recycle.Bin", "*"),
#         "Scheduled_Tasks": os.path.join(base_path, "Windows", "System32", "Tasks", "*"),
#         "NTUSER_&_UserClass": os.path.join(base_path, "Users", "*", "NTUSER.dat")
#     }

# # Function to copy browser artifacts for each user into respective folders
# def copy_browser_artifacts(browser_artifacts, destination_folder, base_path):
#     browser_folder = os.path.join(destination_folder, "Browser_Artifacts")
#     if not os.path.exists(browser_folder):
#         os.makedirs(browser_folder)

#     users_folder = os.path.join(base_path, "Users")
#     user_dirs = [d for d in os.listdir(users_folder) if os.path.isdir(os.path.join(users_folder, d))]

#     for user in user_dirs:
#         user_folder_path = os.path.join(users_folder, user)
#         user_folder = os.path.join(browser_folder, user)
#         if not os.path.exists(user_folder):
#             os.makedirs(user_folder)

#         for browser_info in browser_artifacts:
#             browser_name = browser_info["browser"]
#             browser_path = browser_info["path"].replace("*", user)

#             browser_user_folder = os.path.join(user_folder, browser_name)
#             if not os.path.exists(browser_user_folder):
#                 os.makedirs(browser_user_folder)

#             try:
#                 for match in glob.glob(browser_path):
#                     if os.path.isfile(match):
#                         shutil.copy2(match, browser_user_folder)
#                         print(f"Copied {match} to {browser_user_folder}")
#             except Exception as e:
#                 print(f"Failed to copy {browser_path} for {user}: {e}")

# # Function to copy NTUSER.dat and UserClass.dat files with the required renaming
# def copy_ntuser_artifacts(users_folder, destination_folder):
#     ntuser_folder = os.path.join(destination_folder, "NTUSER_&_UserClass")
#     if not os.path.exists(ntuser_folder):
#         os.makedirs(ntuser_folder)

#     user_dirs = [d for d in os.listdir(users_folder) if os.path.isdir(os.path.join(users_folder, d))]

#     for user in user_dirs:
#         ntuser_file = os.path.join(users_folder, user, "NTUSER.dat")
#         userclass_file = os.path.join(users_folder, user, "AppData", "Local", "Microsoft", "Windows", "UserClass.dat")
#         if os.path.exists(ntuser_file):
#             shutil.copy2(ntuser_file, os.path.join(ntuser_folder, f"{user}_NTUSER.dat"))
#             print(f"Copied {ntuser_file} to {ntuser_folder}")
#         if os.path.exists(userclass_file):
#             shutil.copy2(userclass_file, os.path.join(ntuser_folder, f"{user}_UserClass.dat"))
#             print(f"Copied {userclass_file} to {ntuser_folder}")

# # Function to handle locked files like registry hives using error handling
# def copy_other_artifacts(artifacts, destination_folder):
#     for category, paths in artifacts.items():
#         if category == "Browser_Artifacts":
#             continue

#         category_folder = os.path.join(destination_folder, category)
#         if not os.path.exists(category_folder):
#             os.makedirs(category_folder)

#         if isinstance(paths, dict):  # For nested dictionaries like Registry
#             for name, path in paths.items():
#                 try:
#                     if os.path.exists(path):
#                         try:
#                             shutil.copy2(path, category_folder)
#                             print(f"Copied {path} to {category_folder}")
#                         except PermissionError:
#                             print(f"PermissionError: {path} is locked by the system.")
#                     else:
#                         print(f"{name} not found at {path}")
#                 except Exception as e:
#                     print(f"Failed to copy {name}: {e}")
#         else:
#             try:
#                 if "*" in paths:
#                     for match in glob.glob(paths):
#                         if os.path.isfile(match):
#                             shutil.copy2(match, category_folder)
#                             print(f"Copied {match} to {category_folder}")
#                 elif os.path.isdir(paths):
#                     for root, dirs, files in os.walk(paths):
#                         for file in files:
#                             src_file = os.path.join(root, file)
#                             shutil.copy2(src_file, category_folder)
#                             print(f"Copied {src_file} to {category_folder}")
#                 elif os.path.exists(paths):
#                     shutil.copy2(paths, category_folder)
#                     print(f"Copied {paths} to {category_folder}")
#                 else:
#                     print(f"{category} not found at {paths}")
#             except Exception as e:
#                 print(f"Failed to copy {category}: {str(e)}")

# # Function to delete empty folders
# def delete_empty_folders(folder_path):
#     for root, dirs, files in os.walk(folder_path, topdown=False):
#         for dir in dirs:
#             dir_path = os.path.join(root, dir)
#             if not os.listdir(dir_path):  # Check if directory is empty
#                 os.rmdir(dir_path)
#                 print(f"Deleted empty folder: {dir_path}")

# if __name__ == "__main__":
#     if len(sys.argv) < 3:
#         print("Usage: python collect_windows_artifacts.py <base_path> <destination_folder>")
#         sys.exit(1)

#     if not is_admin():
#         print("Script must be run as an administrator.")
#         sys.exit(1)

#     base_path = sys.argv[1].rstrip("\\")
#     if len(base_path) == 1 and base_path.isalpha():
#         base_path = base_path + ":\\" 

#     destination_folder = sys.argv[2]
#     artifacts = get_artifact_paths(base_path)

#     copy_other_artifacts(artifacts, destination_folder)
#     copy_browser_artifacts(artifacts["Browser_Artifacts"], destination_folder, base_path)
#     copy_ntuser_artifacts(os.path.join(base_path, "Users"), destination_folder)
    
#     # delete_empty_folders(destination_folder)


import os
import shutil
import sys
import glob
import ctypes
import win32file, win32con  # Requires pywin32 library

# Function to check if script is run with administrator privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Function to check if a file is hidden
def is_hidden(filepath):
    try:
        attrs = win32file.GetFileAttributes(filepath)
        return attrs & win32con.FILE_ATTRIBUTE_HIDDEN
    except Exception as e:
        print(f"Error checking attributes for {filepath}: {e}")
        return False

# Define paths to artifacts (based on the provided base path, which could be a drive letter or a folder)
def get_artifact_paths(base_path):
    return {
        "Registry": {
            "BAM": os.path.join(base_path, "Windows", "System32", "config", "BAM"),
            "ShimCache": os.path.join(base_path, "Windows", "AppCompat", "Programs", "Amcache.hve"),
            "SYSTEM": os.path.join(base_path, "Windows", "System32", "config", "SYSTEM"),
            "SAM": os.path.join(base_path, "Windows", "System32", "config", "SAM"),
            "SECURITY": os.path.join(base_path, "Windows", "System32", "config", "SECURITY"),
            "SOFTWARE": os.path.join(base_path, "Windows", "System32", "config", "SOFTWARE")
        },
        "LNK_files": os.path.join(base_path, "Users", "*", "AppData", "Roaming", "Microsoft", "Windows", "Recent", "*.lnk"),
        "Prefetch_files": os.path.join(base_path, "Windows", "Prefetch", "*.pf"),
        "Event_Logs": os.path.join(base_path, "Windows", "System32", "winevt", "Logs", "*.evtx"),
        "Browser_Artifacts": [
            {"browser": "Edge", "path": os.path.join(base_path, "Users", "*", "AppData", "Local", "Microsoft", "Edge", "User Data", "Default", "History")},
            {"browser": "Chrome", "path": os.path.join(base_path, "Users", "*", "AppData", "Local", "Google", "Chrome", "User Data", "Default", "History")},
            {"browser": "IE", "path": os.path.join(base_path, "Users", "*", "AppData", "Roaming", "Microsoft", "Internet Explorer", "UserData", "index.dat")}
        ],
        "Recycle_Bin": os.path.join(base_path, "$Recycle.Bin", "*"),
        "Scheduled_Tasks": os.path.join(base_path, "Windows", "System32", "Tasks", "*"),
        "NTUSER_&_UserClass": os.path.join(base_path, "Users", "*", "NTUSER.dat")
    }

# Function to copy browser artifacts for each user into respective folders
def copy_browser_artifacts(browser_artifacts, destination_folder, base_path):
    browser_folder = os.path.join(destination_folder, "Browser_Artifacts")
    if not os.path.exists(browser_folder):
        os.makedirs(browser_folder)

    users_folder = os.path.join(base_path, "Users")
    user_dirs = [d for d in os.listdir(users_folder) if os.path.isdir(os.path.join(users_folder, d))]

    for user in user_dirs:
        user_folder_path = os.path.join(users_folder, user)
        user_folder = os.path.join(browser_folder, user)
        if not os.path.exists(user_folder):
            os.makedirs(user_folder)

        for browser_info in browser_artifacts:
            browser_name = browser_info["browser"]
            browser_path = browser_info["path"].replace("*", user)

            browser_user_folder = os.path.join(user_folder, browser_name)
            if not os.path.exists(browser_user_folder):
                os.makedirs(browser_user_folder)

            try:
                for match in glob.glob(browser_path):
                    if os.path.isfile(match):
                        shutil.copy2(match, browser_user_folder)
                        print(f"Copied {match} to {browser_user_folder}")
            except Exception as e:
                print(f"Failed to copy {browser_path} for {user}: {e}")


def copy_ntuser_artifacts(user_dir, destination_folder):
    for user in os.listdir(user_dir):
        user_path = os.path.join(user_dir, user)
        userclass_file = os.path.join(user_path, "AppData", "Local", "Microsoft", "Windows", "UserClass.dat")
        if os.path.exists(userclass_file):
            dest_path = os.path.join(destination_folder, f"{user}_UserClass.dat")
            shutil.copy2(userclass_file, dest_path)
            print(f"Copied {userclass_file} to {dest_path}")
        else:
            print(f"UserClass.dat file not found for {user}. Skipping.")

# Function to handle locked files like registry hives using error handling
def copy_other_artifacts(artifacts, destination_folder):
    for category, paths in artifacts.items():
        if category == "Browser_Artifacts":
            continue

        category_folder = os.path.join(destination_folder, category)
        if not os.path.exists(category_folder):
            os.makedirs(category_folder)

        if isinstance(paths, dict):  # For nested dictionaries like Registry
            for name, path in paths.items():
                try:
                    if os.path.exists(path):
                        try:
                            shutil.copy2(path, category_folder)
                            print(f"Copied {path} to {category_folder}")
                        except PermissionError:
                            print(f"PermissionError: {path} is locked by the system.")
                    else:
                        print(f"{name} not found at {path}")
                except Exception as e:
                    print(f"Failed to copy {name}: {e}")
        else:
            try:
                if "*" in paths:
                    for match in glob.glob(paths):
                        if os.path.isfile(match):
                            shutil.copy2(match, category_folder)
                            print(f"Copied {match} to {category_folder}")
                elif os.path.isdir(paths):
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

# Function to delete empty folders
def delete_empty_folders(folder_path):
    for root, dirs, files in os.walk(folder_path, topdown=False):
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            if not os.listdir(dir_path):  # Check if directory is empty
                os.rmdir(dir_path)
                print(f"Deleted empty folder: {dir_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python collect_windows_artifacts.py <base_path> <destination_folder>")
        sys.exit(1)

    if not is_admin():
        print("Script must be run as an administrator.")
        sys.exit(1)

    base_path = sys.argv[1].rstrip("\\")
    if len(base_path) == 1 and base_path.isalpha():
        base_path = base_path + ":\\" 

    destination_folder = sys.argv[2]
    artifacts = get_artifact_paths(base_path)

    copy_other_artifacts(artifacts, destination_folder)
    copy_browser_artifacts(artifacts["Browser_Artifacts"], destination_folder, base_path)
    copy_ntuser_artifacts(os.path.join(base_path, "Users"), destination_folder)
    
    # Optionally delete empty folders
    # delete_empty_folders(destination_folder)
