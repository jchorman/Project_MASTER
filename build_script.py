import os
import shutil
import glob

def determine_location():
    while True:
        location = input("Enter location (h/w for home/work): ").lower()
        if location == 'h':
            return "home"
        elif location == 'w':
            return "work"
        else:
            print("Invalid input. Please enter 'h' for home or 'w' for work.")

def cleanup_after_build():
    # Delete the Project_MASTER directory
    if os.path.exists("Project_MASTER"):
        shutil.rmtree("Project_MASTER")

    # Delete the build directory
    if os.path.exists("build"):
        shutil.rmtree("build")

    # Delete the Project_MASTER.spec file
    spec_file = "Project_MASTER.spec"
    if os.path.isfile(spec_file):
        os.remove(spec_file)

def determine_zip_name(directory):
    base_name = "Project_MASTER"
    zip_files = glob.glob(os.path.join(directory, f"{base_name}*.zip"))
    if not zip_files:
        return f"{base_name}.zip"
    
    # Find the maximum existing index of the zip files
    indices = [0]
    for zip_file in zip_files:
        stripped_name = os.path.basename(zip_file).replace(base_name, '').replace('.zip', '')
        if stripped_name.isdigit():
            indices.append(int(stripped_name))
    next_index = max(indices) + 1
    return f"{base_name}{next_index}.zip"

def build_app(location):
    # Step 0: Prepare the appropriate config file
    if location == "home":
        shutil.copy("config.ini", "config_temp.ini")
    elif location == "work":
        shutil.copy("config_work.ini", "config.ini")

    # Step 1: Build the PyInstaller executable
    command = "pyinstaller --add-data=\"config.ini;.\" --add-data=\"email_body.html;.\" --add-data=\"Project_MASTER.py;.\" --distpath=Project_MASTER Project_MASTER.py"
    os.system(command)

    # Restore original config.ini if we're in home location
    if location == "home":
        shutil.move("config_temp.ini", "config.ini")

    # Determine the unique zip_name before zipping
    if location == "work":
        directory = "C:\\Users\\jimch\\New Jersey Office of Information Technology\\OIT-SolArch - General\\Scripts\\"
        zip_name = determine_zip_name(directory)
    else:
        directory = "C:\\Users\\jimch\\Desktop\\Scripts\\"
        zip_name = determine_zip_name(directory)  # Use the function to determine the name even in the "home" location
    
    # Step 2: Zip the Project_MASTER directory
    shutil.make_archive(zip_name.replace(".zip", ""), 'zip', "Project_MASTER")
    print(f"Moving {zip_name} to {directory}")

    current_directory = os.getcwd()

    # Check if the zip file is already in the target directory
    if os.path.abspath(directory) == os.path.abspath(current_directory):
        print(f"{zip_name} is already in the desired directory.")
        # We don't need to do anything else in this case
    else:
        # Move the zip file to the desired location
        try:
            shutil.move(zip_name, directory)
        except Exception as e:
            print(f"Error occurred while moving the file: {e}")

    # Cleanup build directories and files for "work" location
    if location == "work":
        cleanup_after_build()

location = determine_location()
build_app(location)