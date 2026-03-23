import os

folder = "files"

files = os.listdir(folder)

count = 1

for file in files:

    old_path = os.path.join(folder, file)

    new_name = f"photo_{count}.png"

    new_path = os.path.join(folder, new_name)

    os.rename(old_path, new_path)

    print(f"{file} renamed to {new_name}")

    count += 1

print("All files renamed successfully!")