import time 
start_time = time.time()

import_start = time.time()
import re
import os
import pandas as pd
import subprocess
import argparse    
print(f"\n⏱ Heavy imports loaded in {time.time() - import_start:.2f} seconds\n")




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


def run_powershell_script(usernames=None, userfile_path=None, force_overwrite=False):
    script_path = 'user2json.ps1'
    print("Running PowerShell...")

    command = ['powershell', '-ExecutionPolicy', 'Bypass', '-File', script_path]

    if userfile_path:
        command += ['-UserListFile', userfile_path]
    elif usernames:
        command += ['-Usernames', usernames]

    if force_overwrite:
        command += ['-ForceOverwrite']

    print("PowerShell command:", " ".join(command))  # Debug

    result = subprocess.run(command, capture_output=True, text=True)

    if result.returncode == 0:
        print("✅ PowerShell script executed successfully.")
        print(result.stdout)
    else:
        print(f"❌ PowerShell script failed (code {result.returncode}).")
        print("---- PowerShell Error Output ----")
        print(result.stderr)
        print("---- End of Error Output ----")



def filexcel(username, skill):
    import openpyxl  # Lazy load
    import json
    import os

    filepath = f"json\\{username}.json"
    print(f"json file :  {filepath}")
    
    try:
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            person = json.load(f)
            print("JSON file loaded successfully")
    except FileNotFoundError:
        print(f"Error: JSON file '{filepath}' not found.")
        return
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON file: {e}")
        return

    # Load Excel workbook
    if skill == '4':
        workbook = openpyxl.load_workbook("skill4.xlsx")
    elif skill == '5':
        workbook = openpyxl.load_workbook("skill5.xlsx")
    else:
        print(f"skill has not being choosen:\n please choose skill ")
        return

    
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
    workbook.save(f"excel\{username}-{skill}.xlsx")


def main():
    print("Welcome to Excel filler")
    parse_start = time.time()
    
    parser = argparse.ArgumentParser(description='Process JSON data into Excel file(s)')
    parser.add_argument('-u', '--username', help='Single username (as in AD)')
    parser.add_argument('-f', '--userfile', help='Path to a text file with usernames, one per line')
    parser.add_argument('skill', help='Skill type: choose 4 or 5')
    parser.add_argument('--force-overwrite', action='store_true', help='Force regenerate JSON files')

    args = parser.parse_args()

    if not args.username and not args.userfile:
        parser.error("You must provide either --username or --userfile")

    print(f"\n⏱ Args parsed in {time.time() - parse_start:.2f} seconds\n")

    # Prepare list of usernames
    if args.userfile:
        if not os.path.exists(args.userfile):
            print(f"❌ Error: File not found - {args.userfile}")
            return
        with open(args.userfile, 'r', encoding='utf-8') as f:
            usernames = [line.strip() for line in f if line.strip()]
    else:
        usernames = [args.username]
        
    if args.force_overwrite:
        print("⚠️ Force overwrite enabled. Regenerating JSON for all users.")
        run_powershell_script(
            usernames=args.username if args.username else None,
            userfile_path=args.userfile if args.userfile else None,
            force_overwrite=True
        )
    else:
        usernames_to_generate = [
            u for u in usernames if not os.path.exists(f"json\\{u}.json")
        ]
        if usernames_to_generate:
            print(f"🔄 Generating JSON for {len(usernames_to_generate)} user(s)...")
            if args.userfile:
                run_powershell_script(userfile_path=args.userfile)
            else:
                run_powershell_script(usernames=args.username)
        else:
            print("📂 All JSON files already exist. Skipping PowerShell call.")


    # Now generate Excel for each user
    for username in usernames:
        print(f"\n➡️ Processing user: {username}")
        filexcel(username, args.skill)

    print(f"\n✅ All done! Processed {len(usernames)} user(s).")



# This is the entry point of the script
if __name__ == '__main__':
    main()
    print(f"\n⏱ Total time elapsed: {time.time() - start_time:.2f} seconds\n")
 
