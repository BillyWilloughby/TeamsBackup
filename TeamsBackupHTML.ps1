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

$exportFolder = ".\TeamsChatHTML"
$attachmentsFolder = ".\TeamsAttachments"
$retryDelaySec = 10
$throttleDelayMs = 75

function Log {
    param([string]$msg, [switch]$ErrorMsg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp - $msg"
    if ($ErrorMsg) { Write-Host $line -ForegroundColor Red }
    else { Write-Host $line }
    $line | Tee-Object -FilePath $LogPath -Append | Out-Null
}

Log "Authenticating with Microsoft Graph..."
#Connect-MgGraph -Scopes "User.Read", "Chat.Read", "ChatMessage.Read", "Files.Read"
Connect-MgGraph -Scopes "User.Read", "Chat.Read", "Chat.ReadWrite", "Files.Read"
if (-not (Get-MgContext)) {
    Log "Failed to authenticate with Microsoft Graph. Please check your credentials and permissions." -Error
    return
}
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

Write-Host "    Account              :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.Account)"
Write-Host "    ClientId             :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.ClientId)"
Write-Host "    TenantId             :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.TenantId)"
Write-Host "    AppName              :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.AppName)"
Write-Host "    AuthType             :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.AuthType)"
Write-Host "    Scopes               :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.Scopes -join ', ')"
Write-Host "    Environment          :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.Environment)"
Write-Host "    TokenCredentialType  :" -ForegroundColor Green -NoNewline
Write-Host " $($ctx.TokenCredentialType)"

Read-Host -Prompt "Press Enter to continue..."


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
    $participants = ($chat.Members | ForEach-Object { $_.DisplayName }) -join ", "
    $baseName = if ($chat.ChatType -eq "oneOnOne" -and -not $chat.Topic) {
        ($chat.Members | Where-Object { $_.UserId -ne $me }).DisplayName
    }
    elseif ($chat.Topic) { $chat.Topic } elseif ($participants) {
        $participants -replace '[^a-zA-Z0-9 _-]', '_'
    }
    else { "Chat_" + (Get-Date).ToString("yyyy-MM-dd") }

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
        if ($msg.From -and $msg.From.User) { $chatSender = $msg.From.User.DisplayName }
        elseif ($msg.From -and $msg.From.Application) { $chatSender = "System: $($msg.From.Application.DisplayName)" }

        $time = $msg.CreatedDateTime
        $body = $null

        if ($msg.MessageType -eq "systemEventMessage" -and $msg.EventDetail) {
            $eventType = $msg.EventDetail.'@odata.type'
            $eventJson = $msg.EventDetail | ConvertTo-Json -Depth 5 -Compress
            $body = "<pre>SYSTEM EVENT: $eventType`n$eventJson</pre>"
        }
        elseif ($msg.Body -and $msg.Body.Content -ne "") {
            $body = $msg.Body.Content

            $hostedContentMap = @{}
            try {
                try {
                    $hostedItems = Get-MgChatMessageHostedContent -ChatId $chatId -ChatMessageId $msg.Id -ErrorAction Stop
                }
                catch {
                    if ($_.Exception.Response.StatusCode.Value__ -eq 404) {
                        Log "    Hosted content not found (404) for msg $($msg.Id)" -Error
                        continue
                    }
                    Log "    General failure retrieving hostedContent for msg $($msg.Id): $_" -Error
                    continue
                }
            
                foreach ($hc in $hostedItems) {
                    if (-not $hc.ContentUrl) {
                        Log "    Ignored hostedContent with null ContentUrl for msg $($msg.Id), Id: $($hc.Id)" -Error
                        continue
                    }
            
                    try {
                        $hcStream = Invoke-MgGraphRequest -Uri $hc.ContentUrl -Method GET
                        $bytes = $hcStream.Content.ReadAsByteArrayAsync().Result
                        $base64 = [System.Convert]::ToBase64String($bytes)
                        $hostedContentMap[$hc.Id] = "data:image/png;base64,$base64"
                    }
                    catch {
                        Log "    Failed to download hostedContent ID: $($hc.Id) for msg $($msg.Id): $_" -Error
                    }
                }
            }
            catch {
                Log "    Failed to get hostedContent for msg $($msg.Id): $_" -Error
            }
            

            if ($hostedContentMap.Count -gt 0) {
                $body = $body -replace 'src="cid:(.+?)"', { param($m)
                    $cid = $m.Groups[1].Value
                    if ($hostedContentMap.ContainsKey($cid)) { "src='0'" -f $hostedContentMap[$cid] } else { "src=''" }
                }
            }
        }

        if (-not $body -or $body.Trim() -eq "") {
            Log "Skipped empty message $($msg.Id) at $time"
            continue
        }

        $attachmentsHtml = ""
        Log "Message from $chatSender at $time"
        Start-Sleep -Milliseconds $throttleDelayMs

        if ($msg.Attachments.Count -gt 0) {
            foreach ($att in $msg.Attachments) {
                $attName = if ($att.Name) { $att.Name } else { "[Unnamed]" }
                $fileName = "$($msg.Id)_$attName" -replace '[^a-zA-Z0-9._-]', '_'
                $filePath = Join-Path $chatFolder $fileName
                $relativePath = "..\Attachments\$safeName\$fileName"
            
                if (-not (Test-Path $chatFolder)) {
                    New-Item -ItemType Directory -Path $chatFolder -Force | Out-Null
                }
            
                if ($att.ContentUrl) {
                    Log "    Downloading attachment: $attName from ContentUrl: $($att.ContentUrl)"
                    try {
                        # Attempt direct download (works only for public links)
                        Invoke-WebRequest -Uri $att.ContentUrl -OutFile $filePath -ErrorAction Stop
                        $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                        $attachmentsDownloaded = $true
                    }
                    catch {
                        Log "        Direct ContentUrl failed. Attempting Graph fallback for $attName..."
            
                        # Build fallback Graph API URI using OneDrive path
                        try {
                            # Only works if file is in your personal OneDrive (based on observed path structure)
                            $uri = $att.ContentUrl
                            $path = $uri -replace '^https://[^/]+/[^/]+/([^?]+).*$', '/$1'
                            $path = $path -replace '%20', ' '  # optional
                            $path = [System.Web.HttpUtility]::UrlPathEncode($path)
            
                            $graphUri = "https://graph.microsoft.com/v1.0/me/drive/root:$path`:/content"
                            $token = (Get-MgContext).AccessToken
            
                            Invoke-RestMethod -Uri $graphUri -Headers @{ Authorization = "Bearer $token" } -OutFile $filePath -ErrorAction Stop
                            $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                            $attachmentsDownloaded = $true
                        }
                        catch {
                            Log "        Failed Graph fallback download for $attName : $_" -Error
                            $attachmentsHtml += "<div class='attachment'>Download failed: $attName</div>"
                        }
                    }
                }
                elseif ($att.ContentBytes) {
                    try {
                        $bytes = [System.Convert]::FromBase64String($att.ContentBytes)
                        [System.IO.File]::WriteAllBytes($filePath, $bytes)
                        $attachmentsHtml += "<div class='attachment'>Attachment: <a href='$relativePath'>$fileName</a></div>"
                        $attachmentsDownloaded = $true
                    } catch {
                        Log "        Failed to write: $attName - $_" -Error
                        $attachmentsHtml += "<div class='attachment'>Failed to download attachment: $attName</div>"
                    }
                }
                else {
                    $type = if ($att.'@odata.type') { $att.'@odata.type' } else { "[Unknown]" }
                    Log "        Attachment $attName has no base64 content. Type: $type" -Error
                    $attachmentsHtml += "<div class='attachment'>Not downloadable (cloud/reference): $attName</div>"
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

    # if (-not $attachmentsDownloaded -and (Test-Path $chatFolder)) {
    #     Remove-Item -Recurse -Force -Path $chatFolder
    # }
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
