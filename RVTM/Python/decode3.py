

# Define the file paths as raw strings or use double backslashes
input_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\3307239REQIF.csv'
output_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\3307239REQIF3.csv'

try:
    with open(input_path, 'r', encoding='utf-8') as file:
        content = file.read()


    content = content.replace('\\044', ',')  # Replace \044 with comma
    content = content.replace('\\034', '"')  # Replace \034 with double quotes
    content = content.replace('\\010', '\n') # Replace \010 with new line
    content = content.replace('\\009', '\t') # Replace \009 with tab character

    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(content)

except Exception as e:
    print(f"An error occurred: {e}")
