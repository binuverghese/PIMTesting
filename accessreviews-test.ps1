# Stop on error
$ErrorActionPreference = 'Stop'

# Ensure required module is installed
$requiredModule = 'Microsoft.Graph'
if (!(Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Error "Required module '$requiredModule' is not installed. Install it using: Install-Module $requiredModule -Force"
    exit 1
}

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

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$GroupName,

    [Parameter(Mandatory = $true)]
    [string]$AccessReviewName,

    [bool]$MailNotificationsEnabled = $true,
    [bool]$ReminderNotificationsEnabled = $true,
    [bool]$JustificationRequiredOnApproval = $true,
    [bool]$DefaultDecisionEnabled = $false,

    [ValidateSet("None", "Accept", "Deny")]
    [string]$DefaultDecision = "None",

    [ValidateRange(1, 180)]
    [int]$InstanceDurationInDays = 15,

    [bool]$RecommendationsEnabled = $true,
    [bool]$AutoApplyDecisionsEnabled = $true,

    [ValidateRange(1, 12)]
    [int]$RecurrenceIntervalMonths = 6,

    [ValidateRange(0, 12)]
    [int]$StartDateOffsetMonths = 6
)

try {
    Write-Verbose "Getting group $GroupName from Entra ID..."
    $group = Get-MgGroup -Filter "DisplayName eq '$GroupName'"

    if (!$group) {
        throw "Group '$GroupName' does not exist."
    }

    $groupId = $group.id

    Write-Verbose "Checking if access review '$AccessReviewName' already exists..."
    $existingAccessReview = Get-MgIdentityGovernanceAccessReviewDefinition -Filter "DisplayName eq '$AccessReviewName'"

    if ($existingAccessReview) {
        Write-Verbose "Access review '$AccessReviewName' already exists."
        exit 0
    }

    $params = @{
        displayName                = $AccessReviewName
        descriptionForAdmins       = "Semi-annual review for Entra ID group $GroupName"
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
            mailNotificationsEnabled         = $MailNotificationsEnabled
            reminderNotificationsEnabled     = $ReminderNotificationsEnabled
            justificationRequiredOnApproval  = $JustificationRequiredOnApproval
            defaultDecisionEnabled           = $DefaultDecisionEnabled
            defaultDecision                  = $DefaultDecision
            instanceDurationInDays           = $InstanceDurationInDays
            recommendationsEnabled           = $RecommendationsEnabled
            autoApplyDecisionsEnabled        = $AutoApplyDecisionsEnabled
            recurrence = @{
                pattern = @{
                    type     = "absoluteMonthly"
                    interval = $RecurrenceIntervalMonths
                }
                range   = @{
                    type      = "noEnd"
                    startDate = "$(Get-Date).AddMonths($StartDateOffsetMonths).ToString('yyyy-MM-dd')T12:02:30.667Z"
                }
            }
        }
    }

    Write-Verbose "Creating access review '$AccessReviewName'..."
    New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params
    Write-Verbose "Access review created successfully."

} catch {
    Write-Error "Error occurred: $_"
    exit 1
}

exit 0
