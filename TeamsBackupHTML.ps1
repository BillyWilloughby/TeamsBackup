# Load Microsoft Graph module
Import-Module Microsoft.Graph -MinimumVersion 2.0.0 -Force

# Authenticate user
Connect-MgGraph -Scopes "Chat.Read"

# Create base folders
$exportFolder = ".\ChatHTML"
$attachmentsRoot = ".\Attachments"
New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
New-Item -ItemType Directory -Path $attachmentsRoot -Force | Out-Null

# Get all chats
$chats = Get-MgChat -All

foreach ($chat in $chats) {
    $messages = Get-MgChatMessage -ChatId $chat.Id -All

    $safeChatId = $chat.Id -replace '[^a-zA-Z0-9]', '_'
    $chatFile = Join-Path $exportFolder "Chat_$safeChatId.html"
    $chatFolder = Join-Path $attachmentsRoot $safeChatId
    New-Item -ItemType Directory -Path $chatFolder -Force | Out-Null

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body { font-family: sans-serif; }
        .msg { margin: 10px 0; padding: 10px; border: 1px solid #ccc; border-radius: 8px; }
        .sender { font-weight: bold; color: #2A2A2A; }
        .timestamp { font-size: small; color: gray; }
        .content { margin-top: 5px; }
        .attachment { font-size: small; margin-top: 5px; color: #444; }
    </style>
</head>
<body>
<h2>Chat Type: $($chat.ChatType)</h2>
<hr>
"@

    foreach ($msg in $messages) {
        $chatSender = $msg.From?.User?.DisplayName ?? "Unknown"
        $time = $msg.CreatedDateTime
        $body = $msg.Body.Content
        $attachmentsHtml = ""

        if ($msg.Attachments.Count -gt 0) {
            foreach ($att in $msg.Attachments) {
                if ($att.ContentUrl -ne $null) {
                    $fileName = "$($msg.Id)_$($att.Name)" -replace '[^a-zA-Z0-9._-]', '_'
                    $filePath = Join-Path $chatFolder $fileName
                    try {
                        Invoke-WebRequest -Uri $att.ContentUrl `
                                          -Headers @{ Authorization = "Bearer $((Get-MgContext).AccessToken)" } `
                                          -OutFile $filePath -ErrorAction Stop
                        $relativePath = "..\Attachments\$safeChatId\$fileName"
                        $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                    } catch {
                        $attachmentsHtml += "<div class='attachment'>Failed to download attachment: $($att.Name)</div>"
                    }
                }
            }
        }

        $html += @"
<div class='msg'>
    <div class='sender'>$chatSender</div>
    <div class='timestamp'>$time</div>
    <div class='content'>$body</div>
    $attachmentsHtml
</div>
"@
    }

    $html += "</body></html>"

    # Save HTML file
    Set-Content -Path $chatFile -Value $html -Encoding UTF8
    Write-Host "Exported: $chatFile"
}
