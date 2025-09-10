import os

# Define the target directory
downloads_path = r"/home/user/coastrev/data/downloads"

# Loop through each file in the directory
for filename in os.listdir(downloads_path):
    if 'expedia_price_grid' in filename.lower():
        file_path = os.path.join(downloads_path, filename)
        try:
            os.remove(file_path)
            print(f"Deleted: {file_path}")
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")
