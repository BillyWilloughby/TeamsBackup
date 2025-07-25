# Install full Graph module with all features (v2+ required)
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Ensure Chat-specific module is available
Import-Module Microsoft.Graph.Chat -Force

# Authenticate with correct scopes
Connect-MgGraph -Scopes "Chat.Read"