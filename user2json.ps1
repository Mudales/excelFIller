param (
    [string[]] $Usernames,
    [string] $UserListFile,
    [switch] $ForceOverwrite
)

Write-Host "runing powershell script form powesrtshell`n`n`n`n"
function Export-UserJson {
    param (
        [string] $username
    )

    # Ensure json folder exists
    if (-not (Test-Path "json")) {
        New-Item -Path "json" -ItemType Directory -Force | Out-Null
    }

    $outputFilePath = "json\\$username.json"

    if (-not $ForceOverwrite -and (Test-Path $outputFilePath)) {
        Write-Host " Skipping '$username': JSON already exists at $outputFilePath" -ForegroundColor Yellow
        return
    }

    try {
        $user = Get-ADUser $username -Properties GivenName, Surname, UserPrincipalName, SamAccountName, telephoneassistant, mail |
            Select-Object GivenName, Surname, UserPrincipalName, SamAccountName, telephoneassistant, mail

        if ($null -eq $user) {
            Write-Host " User '$username' not found in AD." -ForegroundColor Red
            return
        }

        $jsonData = $user | ConvertTo-Json -Depth 2
        $jsonData | Out-File -FilePath $outputFilePath -Encoding utf8
        Write-Host " Exported $username to $outputFilePath"
    } catch {
        Write-Host " Error processing user '$username': $_" -ForegroundColor Red
    }
}  # <-- Closing brace for Export-UserJson

# Main Execution Block
if ($UserListFile) {
    Write-Host "the userlist: $UserListFile" 
    if (-Not (Test-Path $UserListFile)) {
        Write-Host " User list file '$UserListFile' not found." -ForegroundColor Red
        exit 1
    }

    $Usernames = Get-Content -Path $UserListFile | Where-Object { $_.Trim() -ne "" }
}

if ($Usernames) {
    foreach ($user in $Usernames) {
        Export-UserJson -username $user
    }
} else {
    Write-Host " No users specified. Use -Usernames or -UserListFile." -ForegroundColor Red
    exit 1
}
 
