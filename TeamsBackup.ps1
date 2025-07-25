# Load Microsoft Graph module
Import-Module Microsoft.Graph -MinimumVersion 2.0.0 -Force

# Authenticate your login
Connect-MgGraph -Scopes "Chat.Read"

# Create attachments root folder
$rootAttachments = ".\Attachments"
New-Item -ItemType Directory -Path $rootAttachments -Force | Out-Null

# Get all your chat threads
$chats = Get-MgChat -All

foreach ($chat in $chats) {
    $messages = Get-MgChatMessage -ChatId $chat.Id -All
    $chatResults = @()

    # Folder-safe chat ID
    $safeChatId = $chat.Id -replace '[^a-zA-Z0-9]', '_'
    $chatFolder = Join-Path $rootAttachments $safeChatId
    New-Item -ItemType Directory -Path $chatFolder -Force | Out-Null

    foreach ($msg in $messages) {
        $attachments = @()

        if ($msg.Attachments.Count -gt 0) {
            foreach ($att in $msg.Attachments) {
                if ($att.ContentUrl -ne $null) {
                    $attachments += $att.ContentUrl

                    # Download attachment
                    $fileExt = [System.IO.Path]::GetExtension($att.Name)
                    $fileName = "$($msg.Id)_$($att.Name)" -replace '[^a-zA-Z0-9._-]', '_'
                    $filePath = Join-Path $chatFolder $fileName

                    try {
                        Invoke-WebRequest -Uri $att.ContentUrl `
                                          -Headers @{ Authorization = "Bearer $((Get-MgContext).AccessToken)" } `
                                          -OutFile $filePath -ErrorAction Stop
                    } catch {
                        Write-Warning "Failed to download $($att.ContentUrl)"
                    }
                }
            }
        }

        $chatResults += [PSCustomObject]@{
            ChatType    = $chat.ChatType
            MessageId   = $msg.Id
            Sender      = $msg.From?.User?.DisplayName
            Timestamp   = $msg.CreatedDateTime
            Content     = $msg.Body.Content
            Attachments = ($attachments -join "`n")
        }
    }

    # Export messages to CSV
    $csvPath = ".\ChatBackup_$safeChatId.csv"
    $chatResults | Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation

    Write-Host "Exported chat: $csvPath"
}
