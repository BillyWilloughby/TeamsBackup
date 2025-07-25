# Load Graph Module and authenticate
#Import-Module Microsoft.Graph -MinimumVersion 2.0.0 

$LogPath = ".\DownloadTeams.log"
$exportFolder = ".\ChatHTML"
$attachmentsRoot = ".\Attachments"
$retryDelaySec = 10
$throttleDelayMs = 500

function Log {
    param([string]$msg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $msg" | Tee-Object -FilePath $LogPath -Append
}

Log "Authenticating with Microsoft Graph..."
Connect-MgGraph -Scopes "Chat.Read"

# Ensure folders
New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
New-Item -ItemType Directory -Path $attachmentsRoot -Force | Out-Null

try {
    Log "Fetching all chats..."
    $chats = Get-MgChat -All -ErrorAction Stop
    Log "Found $($chats.Count) chats."
} catch {
    Log "ERROR retrieving chats: $_"
    return
}

foreach ($chat in $chats) {
    $chatId = $chat.Id
    $participants = ($chat.Members | ForEach-Object {
        $_.DisplayName
    }) -join ", "

    $baseName = if ($chat.Topic) {
        $chat.Topic
    } elseif ($participants) {
        $participants -replace '[^a-zA-Z0-9 _-]', '_'
    } else {
        "Chat_" + (Get-Date).ToString("yyyy-MM-dd")
    }

    $safeName = $baseName -replace '[^a-zA-Z0-9 _-]', '_'
    $chatFile = Join-Path $exportFolder "$safeName.html"
    $chatFolder = Join-Path $attachmentsRoot $safeName
    New-Item -ItemType Directory -Path $chatFolder -Force | Out-Null

    $messages = $null
    $retry = $true
    while ($retry) {
        try {
            $messages = Get-MgChatMessage -ChatId $chatId -All -ErrorAction Stop
            $retry = $false
        } catch {
            if ($_.Exception.Response.StatusCode -eq 403) {
                Log "403 received. Retrying after $retryDelaySec seconds for chat $chatId..."
                Start-Sleep -Seconds $retryDelaySec
            } else {
                Log "Error retrieving messages for chat $chatId0 : $_"
                break
            }
        }
    }

    if (-not $messages) { continue }

    Log "Processing $($messages.Count) messages for chat $chatId"

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

        Log "Message from $chatSender at $time"
        Start-Sleep -Milliseconds $throttleDelayMs

        if ($msg.Attachments.Count -gt 0) {
            foreach ($att in $msg.Attachments) {
                if ($att.ContentUrl) {
                    $fileName = "$($msg.Id)_$($att.Name)" -replace '[^a-zA-Z0-9._-]', '_'
                    $filePath = Join-Path $chatFolder $fileName
                    try {
                        Invoke-WebRequest -Uri $att.ContentUrl `
                            -Headers @{ Authorization = "Bearer $((Get-MgContext).AccessToken)" } `
                            -OutFile $filePath -ErrorAction Stop
                        $relativePath = "..\Attachments\$safeName\$fileName"
                        $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                    } catch {
                        Log "Failed to download attachment $($att.Name): $_"
                        $attachmentsHtml += "<div class='attachment'>Failed: $($att.Name)</div>"
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
    Set-Content -Path $chatFile -Value $html -Encoding UTF8
    Log "Exported chat to: $chatFile"
}

Log "Export completed."
