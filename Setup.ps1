# Install unified Graph SDK (v2+)
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Import the core module (includes all cmdlets)
Import-Module Microsoft.Graph -Force

# Set to beta or v1.0 profile (beta includes more Teams chat features)
Select-MgProfile -Name "beta"

# Connect with delegated user auth
Connect-MgGraph -Scopes "Chat.Read"
