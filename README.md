# ExcelFIller
---

```markdown
# Excel Filler from Active Directory

This tool automates the process of exporting Active Directory user data to JSON files and populating Excel templates (Skill 4 or Skill 5) using that data.

---

## 📁 Project Structure

```
excelFiller/
│
├── main.py                # Python script that drives the whole process
├── user2json.ps1          # PowerShell script to fetch AD user data
├── users.txt              # Example text file with usernames (one per line)
├── json/                  # Folder to store generated JSON files
├── excel/                 # Folder to store final Excel files
├── skill4.xlsx            # Excel template for skill 4
└── skill5.xlsx            # Excel template for skill 5
```

---

## ⚙️ Requirements

- Python 3.x
- PowerShell (on Windows)
- Required Python packages:
  ```bash
  pip install pandas openpyxl
  ```
- AD PowerShell module (for `Get-ADUser` command to work)

---

## 📝 Usage

### Single User

Generate Excel for a single user:

```bash
python main.py -u username 4
```

Where `username` is the Active Directory account (e.g., `john.doe`) and `4` is the skill type (`4` or `5`).

---

### Multiple Users via Text File

Create a `users.txt` file with one username per line:

```
john.doe
jane.smith
admin.user
```

Then run:

```bash
python main.py -f users.txt 5
```

This will:
- Generate JSON files using PowerShell for each user (if missing)
- Fill in the appropriate Excel template (`skill5.xlsx`) with their AD data
- Save results in the `excel/` folder

---

## 🔄 Force Regeneration

If you want to **re-generate all JSONs**, regardless of whether they already exist:

```bash
python main.py -f users.txt 4 --force-overwrite
```

---

## 🛠 Notes

- The PowerShell script fetches the following fields from AD:
  - `GivenName`
  - `Surname`
  - `UserPrincipalName`
  - `SamAccountName`
  - `mail`
  - `telephoneassistant`

- Excel templates must be named `skill4.xlsx` and `skill5.xlsx` and placed in the root folder.

---

## ✅ Output

You will find the filled Excel files inside the `excel/` directory:

```
excel/
├── john.doe-4.xlsx
├── jane.smith-4.xlsx
└── ...
```

---

## 📞 Support

If you're stuck or see errors from PowerShell, make sure:
- You're running the script with permission to query Active Directory.
- PowerShell's execution policy allows script execution.
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```
```

---

Let me know if you'd like a version with screenshots or badges for GitHub.
