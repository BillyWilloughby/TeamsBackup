# Updated script with performance improvements
##########################################################################
# Backup Teams messages from your teams
# to HTML files with attachments and inline images
# Requires Microsoft Graph PowerShell SDK
#
# Billy Willoughby 2025/07/25
##########################################################################

# Timestamped log path
$logTimestamp = (Get-Date).ToString("yyyy-MM-dd HH.mm.ss")
$LogPath = ".\DownloadTeams_$logTimestamp.log"

# Folders for export
$exportFolder = ".\TeamsChatHTML"
$attachmentsFolder = ".\TeamsAttachments"
$retryDelaySec = 10 # seconds to wait before retrying after 403
# Removed throttle delay to improve performance

# Create log function
function Log {
    param([string]$msg, [switch]$ErrorMsg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp - $msg"
    if ($ErrorMsg) { Write-Host $line -ForegroundColor Red }
    else { Write-Host $line }
    $line | Tee-Object -FilePath $LogPath -Append | Out-Null
}

Log "Authenticating with Microsoft Graph..."
# Connect to the Scopes needed for reading chats and files
Connect-MgGraph -Scopes "User.Read", "Chat.Read", "Chat.ReadWrite", "Files.Read"
if (-not (Get-MgContext)) {
    Log "Failed to authenticate with Microsoft Graph. Please check your credentials and permissions." -Error
    return
}

# Create export directories
New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
New-Item -ItemType Directory -Path $attachmentsFolder -Force | Out-Null

# Get current user's ID using /me endpoint directly
try {
    $me = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/me").Id
    $meName = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/me").DisplayName
}
catch {
    Log "Failed to resolve current user identity: $_" -Error
    return
}

# Ensure we have a valid user ID
if (-not $me) {
    Log "Current user ID is null or empty. Cannot proceed." -Error
    return
}
else {
    Log ""
    Log "Current user ID: $me"
    Log "$meName"
    Log "Exporting chats to HTML files in $exportFolder"
    Log "Attachments will be saved in $attachmentsFolder"
    Log ""
}
$ctx = Get-MgContext
Write-Host "Authenticated with Microsoft Graph:"

# Fetch all chats
try {
    Log "Fetching all chats..."
    $chats = Get-MgChat -All -ErrorAction Stop
    Log "Found $($chats.Count) chats."
}
catch {
    Log "ERROR retrieving chats: $_" -Error
    return
}

# Cache for attachments to avoid duplicate downloads
$attachmentCache = @{}

# Process each chat and export messages
foreach ($chat in $chats) {
    $chatId = $chat.Id
    $participants = ($chat.Members | ForEach-Object { $_.DisplayName }) -join ", "
    if ($chat.ChatType -eq "oneOnOne" -and -not $chat.Topic) {
        $baseName = ($chat.Members | Where-Object { $_.UserId -ne $me }).DisplayName
    } elseif ($chat.Topic) {
        $baseName = $chat.Topic
    } else {
        $baseName = "Chat_" + (Get-Date).ToString("yyyy-MM-dd")
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
    $chatFolder = Join-Path $attachmentsFolder $safeName
    if (-not (Test-Path $chatFolder)) { New-Item -ItemType Directory -Path $chatFolder -Force | Out-Null }

    $messageFound = 0
    try {
        Log "Retrieving messages for chat $chatId..."
        $messages = Get-MgChatMessage -ChatId $chatId -All -ErrorAction Stop
    }
    catch {
        Log "ERROR retrieving messages: $_" -Error
        continue
    }
    if (-not $messages) { continue }
    Log "Processing $($messages.Count) messages for chat $chatId"

    # Build HTML using StringBuilder
    $htmlSB = New-Object System.Text.StringBuilder
    $htmlSB.AppendLine("<!DOCTYPE html>")
    $htmlSB.AppendLine("<html><head><meta charset='UTF-8'><style>body{font-family:sans-serif;} .msg{margin:10px 0;padding:10px;border:1px solid #ccc;border-radius:8px;} .sender{font-weight:bold;color:#2A2A2A;} .timestamp{font-size:small;color:gray;} .content{margin-top:5px;} .attachment{font-size:small;margin-top:5px;color:#444;}</style></head><body>")
    $htmlSB.AppendLine("<h2>Chat Type: $($chat.ChatType)</h2>")
    $htmlSB.AppendLine("<p><strong>Chat ID:</strong> $chatId</p><hr>")

    foreach ($msg in $messages) {
        $chatSender = "Unknown"
        if ($msg.From -and $msg.From.User) { $chatSender = $msg.From.User.DisplayName }
        elseif ($msg.From -and $msg.From.Application) { $chatSender = "System: $($msg.From.Application.DisplayName)" }

        $time = $msg.CreatedDateTime
        $body = $null

        if ($msg.MessageType -eq "systemEventMessage" -and $msg.EventDetail) {
            $eventType = $msg.EventDetail.'@odata.type'
            $eventJson = $msg.EventDetail | ConvertTo-Json -Depth 5 -Compress
            $body = "<pre>SYSTEM EVENT: $eventType`n$eventJson</pre>"
        }
        elseif ($msg.Body -and $msg.Body.Content) {
            $body = $msg.Body.Content
            # Hosted content handling
            try {
                $hostedItems = Get-MgChatMessageHostedContent -ChatId $chatId -ChatMessageId $msg.Id -ErrorAction Stop
                foreach ($hc in $hostedItems) {
                    if (-not $hc.ContentUrl) { continue }
                    if ($attachmentCache.ContainsKey($hc.Id)) {
                        $base64 = $attachmentCache[$hc.Id]
                    } else {
                        try {
                            $hcStream = Invoke-MgGraphRequest -Uri $hc.ContentUrl -Method GET
                            $bytes = $hcStream.Content.ReadAsByteArrayAsync().Result
                            $base64 = [System.Convert]::ToBase64String($bytes)
                            $attachmentCache[$hc.Id] = $base64
                        } catch {
                            Log "Failed to download hostedContent ID: $($hc.Id) for msg $($msg.Id): $_" -Error
                            continue
                        }
                    }
                    $hostedContentMap[$hc.Id] = "data:image/png;base64,$base64"
                }
            } catch {
                Log "Failed to get hostedContent for msg $($msg.Id): $_" -Error
            }
            if ($hostedContentMap.Count -gt 0) {
                $body = $body -replace 'src="cid:(.+?)"', { param($m)
                    $cid = $m.Groups[1].Value
                    if ($hostedContentMap.ContainsKey($cid)) { "src='0'" -f $hostedContentMap[$cid] } else { "src=''" }
                }
                $messageFound++
            }
        }

        if (-not $body -or $body.Trim() -eq "") {
            Log "Skipped empty message $($msg.Id) at $time"
            continue
        }

        $attachmentsHtml = ""
        Log "Message from $chatSender at $time"

        if ($msg.Attachments.Count -gt 0) {
            foreach ($att in $msg.Attachments) {
                $attName = if ($att.Name) { $att.Name } else { "[Unnamed]" }
                $fileName = "$($msg.Id)_$attName" -replace '[^a-zA-Z0-9._-]', '_'
                $filePath = Join-Path $chatFolder $fileName
                $relativePath = "..\Attachments\$safeName\$fileName"
                if (-not (Test-Path $chatFolder)) { New-Item -ItemType Directory -Path $chatFolder -Force | Out-Null }
                if ($att.ContentUrl) {
                    Log "Downloading attachment: $attName from ContentUrl"
                    try {
                        Invoke-WebRequest -Uri $att.ContentUrl -OutFile $filePath -ErrorAction Stop
                        $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                    } catch {
                        Log "Direct download failed. Attempting Graph fallback for $attName..."
                        try {
                            $uri = $att.ContentUrl
                            $path = $uri -replace '^https://[^/]+/[^/]+/([^?]+).*$', '/$1'
                            $path = [System.Web.HttpUtility]::UrlPathEncode($path)
                            $graphUri = "https://graph.microsoft.com/v1.0/me/drive/root:$path`:/content"
                            $token = (Get-MgContext).AccessToken
                            Invoke-RestMethod -Uri $graphUri -Headers @{ Authorization = "Bearer $token" } -OutFile $filePath -ErrorAction Stop
                            $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                        } catch {
                            Log "Failed Graph fallback download for $attName : $_" -Error
                            $attachmentsHtml += "<div class='attachment'>Download failed: $attName</div>"
                        }
                    }
                }
                elseif ($att.ContentBytes) {
                    try {
                        $bytes = [System.Convert]::FromBase64String($att.ContentBytes)
                        [System.IO.File]::WriteAllBytes($filePath, $bytes)
                        $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                    } catch {
                        Log "Failed to write: $attName - $_" -Error
                        $attachmentsHtml += "<div class='attachment'>Failed to download attachment: $attName</div>"
                    }
                }
                else {
                    $type = if ($att.'@odata.type') { $att.'@odata.type' } else { "[Unknown]" }
                    Log "Attachment $attName has no base64 content. Type: $type" -Error
                    $attachmentsHtml += "<div class='attachment'>Not downloadable (cloud/reference): $attName</div>"
                }
            }
        }

        $htmlSB.AppendLine("<div class='msg'><div class='sender'>$chatSender</div><div class='timestamp'>$time</div><div class='content'>$body</div>$attachmentsHtml</div>")
    }

    $htmlSB.AppendLine("</body></html>")
    if ($messageFound -eq 0) {
        Log "No messages with content found for chat $chatId. Skipping export."
        continue
    }
    Set-Content -Path $chatFile -Value $htmlSB.ToString() -Encoding UTF8
    Log "Exported chat to: $chatFile"
}
Disconnect-MgGraph
Remove-Variable -Name MgContext -Scope Global -ErrorAction SilentlyContinue
Log "Export completed."

##########################################################################
# THIS SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,       #
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF     #
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.#
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY  #
# CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,  #
# TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE     #
# SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.                #
##########################################################################
