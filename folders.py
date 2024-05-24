import os
import shutil

# List all files in the current directory
files = os.listdir()

for file in files:
    # Split the filename by dashes and check if it starts with a date
    parts = file.split('-')
    if len(parts) >= 3 and parts[0].isdigit() and parts[1].isdigit() and parts[2][:2].isdigit():
        # Extract year, month, and day from the filename
        year, month, day = parts[0], parts[1], parts[2][:2]  # Assuming 'day' is always two digits
        
        # Format for the new folder name (year-month-day)
        new_folder = f'{month}-{day}'
        
        # Create new directory if it doesn't exist
        if not os.path.exists(new_folder):
            os.makedirs(new_folder)
        
        # Construct new path for the file
        new_path = os.path.join(new_folder, file)
        
        # Move the file to the new directory
        shutil.move(file, new_path)
