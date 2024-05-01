function user2json {
    param (
        [Parameter(Mandatory = $true)]
        [string] $username
    )
    # Check if the "json" folder exists
    if (-not (Test-Path "json")) {
        # If the "json" folder does not exist, create it
        New-Item -Path "json" -ItemType Directory -Force
    }
    $outputFilePath = "json\\$username.json"

    # Get AD users and select the desired properties
    $user = Get-ADUser $username -Properties GivenName, Surname, UserPrincipalName, SamAccountName, telephoneassistant, mail | 
        Select-Object GivenName, Surname, UserPrincipalName, SamAccountName, telephoneassistant, mail 

    # Convert the users to JSON format
    $jsonData = $user | ConvertTo-Json -Depth 2

    # Save the JSON data to a file
    $jsonData | Out-File -FilePath $outputFilePath -Encoding utf8

    Write-Host "AD user data saved to $outputFilePath"
}

# This block executes when the script is run from the command line.
# It checks for the username parameter and calls the function with that username.
if ($args.Count -gt 0) {
    # The first argument is the username
    $username = $args[0]
    # Call the function with the provided username
    user2json -username $username
} else {
    Write-Host "Please provide a username as a command-line argument."
}
