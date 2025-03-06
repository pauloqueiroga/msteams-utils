<#
.SYNOPSIS
    Sends Microsoft Teams messages to a list of users.

.DESCRIPTION
    This script sends a specified message to multiple Microsoft Teams users.
    The message is sent individually to each user (not as a group).
    Uses Microsoft Graph API to send the messages.
    
.PARAMETER UsersFile
    Path to a file containing a list of user emails (one per line).
    This is a mandatory parameter.

.PARAMETER Message
    The message to send to all users. Defaults to "Hello!" if not specified.

.EXAMPLE
    .\Send-TeamsMessage.ps1 -UsersFile "C:\users.txt" -Message "Meeting at 3 PM today!"

.EXAMPLE
    .\Send-TeamsMessage.ps1 -UsersFile "C:\users.txt"
    (Will send "Hello!" to all users in the file)
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$UsersFile,
    
    [Parameter(Mandatory=$false)]
    [string]$Message = "Hello!"
)

# Check if the users file exists
if (-not (Test-Path $UsersFile)) {
    Write-Error "Error: User file not found at path: $UsersFile"
    exit 1
}

# Verify Microsoft Graph PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Host "Installing Microsoft Graph PowerShell modules..."
    Install-Module -Name Microsoft.Graph.Users -Force -AllowClobber -Scope CurrentUser
    Install-Module -Name Microsoft.Graph.Authentication -Force -AllowClobber -Scope CurrentUser
    Install-Module -Name Microsoft.Graph.Beta.Users -Force -AllowClobber -Scope CurrentUser
}

# Import the Microsoft Graph modules
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Beta.Users

# Define required permissions for Microsoft Graph
$requiredScopes = @(
    "Chat.ReadWrite",
    "User.Read.All"
)

# Connect to Microsoft Graph with the required permissions
try {
    Connect-MgGraph -Scopes $requiredScopes
    Write-Host "Successfully connected to Microsoft Graph API"
} catch {
    Write-Error "Failed to connect to Microsoft Graph API: $_"
    exit 1
}

# Read user list from file
try {
    $users = Get-Content $UsersFile -ErrorAction Stop
} catch {
    Write-Error "Error reading user list file: $_"
    exit 1
}

# Check if user list is empty
if ($users.Count -eq 0) {
    Write-Error "Error: User list file is empty"
    exit 1
}

# Information about what's happening
Write-Host "Preparing to send message: '$Message'"
Write-Host "Number of users to message: $($users.Count)"

# Send message to each user
$successCount = 0
$failCount = 0

foreach ($userEmail in $users) {
    # Trim any whitespace
    $userEmail = $userEmail.Trim()
    
    # Skip empty lines
    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        continue
    }
    
    try {
        Write-Host "Sending message to $userEmail..." -NoNewline
        
        # Get the user ID from their email address
        $userId = (Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop).Id
        
        if (-not $userId) {
            Write-Host "Failed! User not found." -ForegroundColor Red
            $failCount++
            continue
        }
        
        # Create a new chat with the user
        $chatBody = @{
            "chatType" = "oneOnOne"
            "members" = @(
                @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    "roles" = @("owner")
                    "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($env:USERINFO_ID)')"
                },
                @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    "roles" = @("owner")
                    "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$userId')"
                }
            )
        }
        
        # Create or get existing chat
        $chatId = $null
        $existingChats = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/me/chats" -OutputType PSObject
        
        foreach ($chat in $existingChats.value) {
            if ($chat.chatType -eq "oneOnOne") {
                $members = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/chats/$($chat.id)/members" -OutputType PSObject
                foreach ($member in $members.value) {
                    if ($member.userId -eq $userId) {
                        $chatId = $chat.id
                        break
                    }
                }
            }
            if ($chatId) { break }
        }
        
        # If no existing chat, create a new one
        if (-not $chatId) {
            $newChat = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/chats" -Body ($chatBody | ConvertTo-Json -Depth 10) -ContentType "application/json" -OutputType PSObject
            $chatId = $newChat.id
        }
        
        # Create the message body
        $messageBody = @{
            "body" = @{
                "content" = $Message
                "contentType" = "text"
            }
        }
        
        # Send the message to the chat
        $messageUri = "https://graph.microsoft.com/beta/chats/$chatId/messages"
        Invoke-MgGraphRequest -Method POST -Uri $messageUri -Body ($messageBody | ConvertTo-Json) -ContentType "application/json"
        
        Write-Host "Success!" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "Failed!" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        $failCount++
    }
}

# Display summary
Write-Host "`nMessage sending complete!"
Write-Host "Summary:"
Write-Host "  Successful: $successCount"
Write-Host "  Failed: $failCount"
Write-Host "  Total: $($users.Count)"

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Host "Disconnected from Microsoft Graph API"