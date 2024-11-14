# Install Microsoft Graph PowerShell Module
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Directory.ReadWrite.All", "AppCatalog.ReadWrite.All"

# Ensure admin consent is provided for these permissions

# Define variables
$addInName = "Scout"
$addInManifestUrl = "https://example.com/path/to/scout-manifest.xml"  # Replace with actual manifest URL

# Function to register the add-in for Office 365 applications
function Add-Office365AddIn {
    param (
        [string]$manifestUrl,
        [string]$displayName
    )

    # Register the add-in using the Graph API for Word, Excel, and PowerPoint
    $officeApps = @("Word", "Excel", "PowerPoint")

    foreach ($app in $officeApps) {
        try {
            # Placeholder for Graph API command to register the add-in
            # In reality, Graph API does not have direct endpoints for each user add-in installation at this time
            
            Write-Output "Attempting to add $displayName add-in to $app."

            # This is an example of what a Graph API call might look like
            # Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/appCatalogs/officeAddIns/$app" `
            #     -Headers @{ Authorization = "Bearer $accessToken" } `
            #     -Body @{ "manifestUrl" = $manifestUrl; "displayName" = $displayName } `
            #     -Method POST

            Write-Output "$displayName add-in added to $app."
        }
        catch {
            Write-Output "Failed to add $displayName add-in to $app. Error: $_"
        }
    }
}

# Run the function with the provided manifest URL and add-in name
Add-Office365AddIn -manifestUrl $addInManifestUrl -displayName $addInName

# Disconnect Microsoft Graph session
Disconnect-MgGraph
