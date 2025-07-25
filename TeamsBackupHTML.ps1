##########################################################################
# Backup Teams messages from your teams
# to HTML files with attachments
# Requires Microsoft Graph PowerShell SDK
# https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0
##########################################################################


# Timestamped log path
$logTimestamp = (Get-Date).ToString("yyyy-MM-dd HH.mm.ss")
$LogPath = ".\DownloadTeams_$logTimestamp.log"

$exportFolder = ".\ChatHTML"
$attachmentsRoot = ".\Attachments"
$retryDelaySec = 10
$throttleDelayMs = 75

function Log {
    param([string]$msg, [switch]$ErrorMsg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp - $msg"
    if ($ErrorMsg) {
        Write-Host $line -ForegroundColor Red
    }
    else {
        Write-Host $line
    }
    $line | Tee-Object -FilePath $LogPath -Append | Out-Null
}


Log "Authenticating with Microsoft Graph..."
Connect-MgGraph -Scopes "Chat.Read", "ChatMessage.Read", "Files.Read"

# Ensure folders
New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
New-Item -ItemType Directory -Path $attachmentsRoot -Force | Out-Null

try {
    Log "Fetching all chats..."
    $chats = Get-MgChat -All -ErrorAction Stop
    Log "Found $($chats.Count) chats."
}
catch {
    Log "ERROR retrieving chats: $_" -Error
    return
}

foreach ($chat in $chats) {
    $chatId = $chat.Id
    $participants = ($chat.Members | ForEach-Object {
            $_.DisplayName
        }) -join ", "

    $baseName = if ($chat.Topic) {
        $chat.Topic
    }
    elseif ($participants) {
        $participants -replace '[^a-zA-Z0-9 _-]', '_'
    }
    else {
        "Chat_" + (Get-Date).ToString("yyyy-MM-dd")
    }

    $safeName = $baseName -replace '[^a-zA-Z0-9 _-]', '_'
    $timestampSuffix = (Get-Date).ToString("yyyy-MM-dd.HH-mm")
 
    $baseFileName = "$safeName`_$timestampSuffix"
    $chatFile = Join-Path $exportFolder "$baseFileName.html"
    $index = 1
    while (Test-Path $chatFile) {
        $chatFile = Join-Path $exportFolder "$baseFileName`_$index.html"
        $index++
    }

    $chatFolder = Join-Path $attachmentsRoot $safeName
    $attachmentsDownloaded = $false

    $messages = $null
    $retry = $true
    while ($retry) {
        try {
            $messages = Get-MgChatMessage -ChatId $chatId -All -ErrorAction Stop
            $retry = $false
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 403) {
                Log "403 received. Retrying after $retryDelaySec seconds for chat $chatId..."
                Start-Sleep -Seconds $retryDelaySec
            }
            else {
                Log "Error retrieving messages for chat $chatId : $_" -Error
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
        $chatSender = "Unknown"
        if ($msg.From -and $msg.From.User) {
            $chatSender = $msg.From.User.DisplayName
        }
        elseif ($msg.From -and $msg.From.Application) {
            $chatSender = "System: $($msg.From.Application.DisplayName)"
        }

        $time = $msg.CreatedDateTime
        $body = $null

        if ($msg.MessageType -eq "systemEventMessage" -and $msg.EventDetail) {
            $eventType = $msg.EventDetail.'@odata.type'
            $eventJson = $msg.EventDetail | ConvertTo-Json -Depth 5 -Compress
            $body = "<pre>SYSTEM EVENT: $eventType`n$eventJson</pre>"
        }
        elseif ($msg.Body -and $msg.Body.Content -ne "") {
            $body = $msg.Body.Content
        }

        # Skip blank or deleted messages
        if (-not $body -or $body.Trim() -eq "") {
            Log "Skipped empty message $($msg.Id) at $time"
            continue
        }

        $attachmentsHtml = ""
        Log "Message from $chatSender at $time"
        Start-Sleep -Milliseconds $throttleDelayMs

        if ($msg.Attachments.Count -gt 0) {
            foreach ($att in $msg.Attachments) {
                if ($att.Name -and $att.ContentBytes) {
                    $fileName = "$($msg.Id)_$($att.Name)" -replace '[^a-zA-Z0-9._-]', '_'
                    $filePath = Join-Path $chatFolder $fileName
                    try {
                        Write-Host "        Downloading: $($att.Name)"
        
                        $bytes = [System.Convert]::FromBase64String($att.ContentBytes)
                        if (-not (Test-Path $chatFolder)) {
                            New-Item -ItemType Directory -Path $chatFolder -Force | Out-Null
                        }
        
                        [System.IO.File]::WriteAllBytes($filePath, $bytes)
        
                        $relativePath = "..\Attachments\$safeName\$fileName"
                        $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                        $attachmentsDownloaded = $true
                    }
                    catch {
                        Log "        Failed to write: $($att.Name) - $_" -Error
                        $attachmentsHtml += "<div class='attachment'>Failed to download attachment: $($att.Name)</div>"
                    }
                }
                else {
                    $type = $att.'@odata.type'
                    Log "        Attachment $($att.Name) has no base64 content. Type: $type" -Error
                    $attachmentsHtml += "<div class='attachment'>Not downloadable (cloud/reference): $($att.Name)</div>"
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

    if (-not $attachmentsDownloaded -and (Test-Path $chatFolder)) {
        Remove-Item -Recurse -Force -Path $chatFolder
    }
}

Log "Export completed."
