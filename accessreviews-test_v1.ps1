# Stop on error
$ErrorActionPreference = 'Stop'

# Validate required modules
$requiredModules = @('Microsoft.Graph', 'ImportExcel')
foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Error "Required module '$module' is not installed. Install it using: Install-Module $module -Force"
        exit 1
    }
}

# Import Excel file
$excelFilePath = "C:\Users\binuverghese\OneDrive - Microsoft\AccessReviewTest\PIMTesting\InputFile.xlsx"  # Update this path
if (!(Test-Path $excelFilePath)) {
    Write-Error "Excel file not found at: $excelFilePath"
    exit 1
}

# Read Excel data
$inputData = Import-Excel -Path $excelFilePath

# Securely Get Credentials 
$credential = Get-Credential -Message "Enter Azure App Client ID and Secret" -UserName "Client ID"

# Extract Secure Credentials
$clientId = $credential.UserName
$clientSecret = $credential.GetNetworkCredential().Password

# Set Tenant ID Securely (Replace with environment variable if needed)
$tenantId = Read-Host "Enter Tenant ID"

# Get Access Token Securely
$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
} 

try {
    Write-Verbose "Retrieving Access Token..."
    $connection = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $body
    $token = $connection.access_token | ConvertTo-SecureString -AsPlainText -Force
} catch {
    Write-Error "Failed to retrieve access token. Ensure your credentials are correct."
    exit 1
}

# Connect to Microsoft Graph securely
Connect-MgGraph -AccessToken $token

foreach ($entry in $inputData) {
    try {
        # Read parameters from Excel
        $groupName = $entry.GroupName
        $accessReviewName = $entry.AccessReviewName

        Write-Verbose "Processing Access Review for Group: $groupName"

        # Get Entra ID Group
        Write-Verbose "Fetching group '$groupName' from Entra ID..."
        $group = Get-MgGroup -Filter "DisplayName eq '$groupName'"
        if (!$group) {
            throw "Group '$groupName' does not exist."
        }
        $groupId = $group.id

        # Check if Access Review already exists
        Write-Verbose "Checking if access review '$accessReviewName' exists..."
        $existingAccessReview = Get-MgIdentityGovernanceAccessReviewDefinition -Filter "DisplayName eq '$accessReviewName'"

        if ($existingAccessReview) {
            Write-Verbose "Access review '$accessReviewName' already exists. Skipping..."
            continue
        }

        # Define Access Review parameters
        $params = @{
            displayName                = $accessReviewName
            descriptionForAdmins       = "Periodic review for Entra ID group $groupName"
            descriptionForReviewers    = "Review members' access rights."
            scope                      = @{
                "@odata.type" = "#microsoft.graph.accessReviewQueryScope"
                query         = "/groups/$groupId/transitiveMembers"
                queryType     = "MicrosoftGraph"
            }
            reviewers                  = @(
                @{
                    query     = "/groups/$groupId/owners"
                    queryType = "MicrosoftGraph"
                }
            )
            settings                   = @{
                mailNotificationsEnabled         = $true
                reminderNotificationsEnabled     = $true
                justificationRequiredOnApproval  = $true
                defaultDecisionEnabled           = $false
                defaultDecision                  = "None"
                instanceDurationInDays           = 15
                recommendationsEnabled           = $true
                autoApplyDecisionsEnabled        = $true
                recurrence = @{
                    pattern = @{
                        type     = "absoluteMonthly"
                        interval = 6
                    }
                    range   = @{
                        type      = "noEnd"
                        startDate = "$(Get-Date).AddMonths(6).ToString('yyyy-MM-dd')T12:02:30.667Z"
                    }
                }
            }
        }

        # Create Access Review
        Write-Verbose "Creating access review '$accessReviewName'..."
        New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params
        Write-Verbose "Access review '$accessReviewName' created successfully."

    } catch {
        Write-Error "Error processing '$groupName': $_"
    }
}

exit 0
