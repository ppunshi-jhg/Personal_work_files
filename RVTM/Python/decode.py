import re

# This function converts octal to its corresponding character
def octal_to_char(match):
    return chr(int(match.group(0)[1:], 8))

# Define the file paths as raw strings or use double backslashes
input_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\ratm-export.txt'
output_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\decode-ratm.txt'

try:
    with open(input_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # Replace octal codes with corresponding ASCII characters
    converted_content = re.sub(r'\\[0-7]{3}', octal_to_char, content)

    if converted_content:
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(converted_content)
    else:
        print("No octal codes were found or the file was empty.")
except Exception as e:
    print(f"An error occurred: {e}")