import time
import openpyxl
import re
import os
import json
import pandas as pd
import subprocess


def returnid(id):
    pattern = r"^(?:01-)?(\d+)@.*$"
    match = re.search(pattern, id)
    
    if match:
        # Extract the ID from the capturing group
        id_number = match.group(1)
        return id_number
    


def remove_suffix_1(input_string):
    """
    Removes the suffix '1' from the input string if it ends with '1'.
    
    Args:
        input_string (str): The input string to process.
    
    Returns:
        str: The input string without the suffix '1' if it ends with '1'.
    """
    # Check if the input string ends with '1'
    if input_string.endswith('1'):
        # Remove the suffix '1' and return the resulting string
        return input_string[:-1]
    else:
        # Return the original string if it doesn't end with '1'
        return input_string


#split user.name to user name
def splitusername(username):
    if username.endswith('1'):
        # Remove the suffix '1' and return the resulting string
        new_user = username[:-1]
        return new_user.split('.')
    splited = username.split('.')
    return splited

def run_powershell_script(*args):
    script_path='user2json.ps1'
    print("runing powershell..")
    command = ['powershell', '-ExecutionPolicy', 'Bypass', '-File', script_path]
    
    # Add the arguments to the command
    command.extend(args)
    
    # Run the PowerShell script with the provided arguments
    # subprocess.run will execute the command and wait for it to complete
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    
    # Get the output and errors
    output = result.stdout
    errors = result.stderr
    
    # Check the return code to see if the script executed successfully
    if result.returncode == 0:
        print(f"PowerShell script executed successfully.")
        print(f"Output:\n{output}")
    else:
        print(f"PowerShell script execution failed with return code {result.returncode}.")
        print(f"Errors:\n{errors}")



def filexcel(username):
    # Load the JSON file
    filepath = f"json\{username}.json"
    print(f"json file :  {filepath}")
    
    try:
        # Open the JSON file with 'utf-8-sig' encoding to handle BOM
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            person = json.load(f)
            print("JSON file loaded successfully")
    except FileNotFoundError:
        print(f"Error: JSON file '{filepath}' not found.")
        return
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON file: {e}")
        return
    # Load the Excel workbook and select the desired sheet
    workbook = openpyxl.load_workbook('template.xlsx')
    sheet = workbook['גיליון1']  # Replace 'Sheet1' with the name of the sheet you want to work on

    
    sheet['B8'] = person['GivenName']  # Firstname in heb
    sheet['D8'] = person['Surname']  # last name heb
    sheet['B7'] = returnid(person['UserPrincipalName'])  # T.z number
    sheet['B9'] = f"{splitusername(person['SamAccountName'])[0]}" # firstname in english
    sheet['D9'] = f"{splitusername(person['SamAccountName'])[1]}" # last name in english
    sheet['B13'] = person['mail']  # Fill in the email address in B13
    sheet['D12'] = person['telephoneassistant']  # Fill in the phone number 

    # Save the updated workbook
    if not os.path.exists("excel"):
        # Create the directory if it doesn't exist
        os.makedirs("excel")
    workbook.save(f"excel\{username}.xlsx")


# Function to parse arguments and call filexcel function
def main():
    import argparse    
    # Create an argument parser
    username = argparse.ArgumentParser(description='Process JSON data into Excel file')
    username.add_argument('username', help='username')
    # Parse the arguments
    args = username.parse_args()
    user = args.username

    run_powershell_script(args.username)
    time.sleep(2)
    print(args.username)
    filexcel(args.username)
    

# This is the entry point of the script
if __name__ == '__main__':
    main()
