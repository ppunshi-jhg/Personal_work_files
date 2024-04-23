import re

# Define the file paths as raw strings or use double backslashes
input_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\ratm-export.txt'
output_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\decode1-ratm.txt'

try:
    with open(input_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # Replace octal codes with a comma
    converted_content = re.sub(r'\\[0-7]{3}', ',', content)

    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(converted_content)

except Exception as e:
    print(f"An error occurred: {e}")