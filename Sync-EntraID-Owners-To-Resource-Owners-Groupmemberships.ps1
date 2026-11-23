#####################################################
# HelloID-Conn-SA-Sync-EntraID-Owners-To-ResourceOwner-Groupmemberships
#
# Version: 1.1.0
#####################################################

# Set to false to actually perform actions - Only run as DryRun when testing/troubleshooting!
$dryRun = $false
# Set to true to log each individual action - May cause lots of logging, so use with cause, Only run testing/troubleshooting!
$verboseLogging = $false

switch ($verboseLogging) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}
$informationPreference = "Continue"
$WarningPreference = "Continue"

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

######################################################################################
# Connection Configuration - Set up connection credentials
######################################################################################
# HelloID API connection (required)
$helloIDPortalBaseUrl = $portalBaseUrl # When running from HelloID, set from default Global Variable
$helloIDPortalApiKey = $portalApiKey # When running from HelloID, set from default Global Variable
$helloIDPortalApiSecret = $portalApiSecret # When running from HelloID, set from default Global Variable

# Entra ID / Microsoft Graph connection (required)
# Used to connect to Microsoft Graph API
$MSGraphBaseUri = "https://graph.microsoft.com/" # Fixed value
# $EntraIdTenantId = "" # Set from Global Variable
# $EntraIdAppId = "" # Set from Global Variable
# $EntraIdCertificateBase64String = "" # Set from Global Variable
# $EntraIdCertificatePassword = "" # Set from Global Variable
######################################################################################

# Filter which groups to sync using Microsoft Graph search filter (optional)
# $search is not supported by default. Use $filter options
# Examples:
# - All groups: $null
# - By displayName: "startsWith(displayName,'department_')"
# Note: Only displayName and description support the search filter
# Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter?tabs=http#using-search-on-directory-object-collections
$entraIDGroupsSearchFilter = $null
######################################################################################

#HelloID Configuration
$resourceOwnerGroupSource = "Local" # Specify the source of the groups - if source is any other than "Local", the sync of the target system itself might overwrite the memberships set form this sync
# The HelloID Resource owner group will be queried based on the distribution group name and the specified prefix and suffix
$resourceOwnerGroupPrefix = "" # Specify prefix to recognize the resource owner group
$resourceOwnerGroupSuffix = " Resource Owner" # Specify suffix to recognize the resource owner group
$removeMembers = $true # If true, existing members will be removed if they no longer have full access to the corresponding mailbox - This will overwrite manual added users

#region functions
function Write-StatusMessage {
    <#
    .SYNOPSIS
    Writes a status message to the appropriate logging system.
    
    .DESCRIPTION
    When running locally: Uses native PowerShell cmdlets (Write-Information, Write-Warning, Write-Error)
    When running in HelloID: Uses HelloID's native Hid-Write-Status function
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Message,

        [Parameter(Mandatory = $true)]
        [String]
        $Event
    )
    
    if ($null -eq $portalBaseUrl) {
        # Running locally - use native PowerShell cmdlets
        switch ($Event) {
            "Information" { Write-Information ($Message) -InformationAction Continue }
            "Warning" { Write-Warning ($Message) -WarningAction Continue }
            "Success" { Write-Information ($Message) -InformationAction Continue }
            "Error" { Write-Error ($Message) -ErrorAction Continue }
            "Critical" { Write-Error ($Message) -ErrorAction Continue }
            "Failed" { Write-Error ($Message) -ErrorAction Continue }
        }
    }
    else {
        # Running in HelloID - use native HelloID function
        Hid-Write-Status -Message $Message -Event $Event
    }
}

function Write-SummaryMessage {
    <#
    .SYNOPSIS
    Writes a summary message to the appropriate logging system.
    
    .DESCRIPTION
    When running locally: Uses native PowerShell cmdlets (Write-Information, Write-Warning, Write-Error)
    When running in HelloID: Uses HelloID's native Hid-Write-Summary function
    #>
    [cmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Message,

        [Parameter(Mandatory = $true)]
        [String]
        $Event
    )
    
    if ($null -eq $portalBaseUrl) {
        # Running locally - use native PowerShell cmdlets
        switch ($Event) {
            "Information" { Write-Information ($Message) -InformationAction Continue }
            "Warning" { Write-Warning ($Message) -WarningAction Continue }
            "Success" { Write-Information ($Message) -InformationAction Continue }
            "Error" { Write-Error ($Message) -ErrorAction Continue }
            "Critical" { Write-Error ($Message) -ErrorAction Continue }
            "Failed" { Write-Error ($Message) -ErrorAction Continue }
        }
    }
    else {
        # Running in HelloID - use native HelloID function
        Hid-Write-Summary -Message $Message -Event $Event
    }
}

function Remove-StringLatinCharacters {
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}

function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Resolve-HelloIDError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq "System.Net.WebException") {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json)
            # error message can be either in [resultMsg] or [message] or [textResult] or [error]
            if ([bool]($errorDetailsObject.PSobject.Properties.name -eq "resultMsg")) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.resultMsg
            }
            elseif ([bool]($errorDetailsObject.PSobject.Properties.name -eq "message")) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.message
            }
            elseif ([bool]($errorDetailsObject.PSobject.Properties.name -eq "textResult")) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.textResult
            }
            elseif ([bool]($errorDetailsObject.PSobject.Properties.name -eq "error")) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails # Temporarily assignment
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Resolve-HTTPError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId
            MyCommand             = $ErrorObject.InvocationInfo.MyCommand
            RequestUri            = $ErrorObject.TargetObject.RequestUri
            ScriptStackTrace      = $ErrorObject.ScriptStackTrace
            ErrorMessage          = ""
        }

        if ($ErrorObject.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") {
            # $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message # Does not show the correct error message for the Raet IAM API calls
            $httpErrorObj.ErrorMessage = $ErrorObject.Exception.Message

        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq "System.Net.WebException") {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }

        Write-Output $httpErrorObj
    }
}

function Get-ErrorMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $errorMessage = [PSCustomObject]@{
            VerboseErrorMessage = $null
            AuditErrorMessage   = $null
        }

        if ( $($ErrorObject.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") -or $($ErrorObject.Exception.GetType().FullName -eq "System.Net.WebException")) {
            $httpErrorObject = Resolve-HTTPError -Error $ErrorObject

            $errorMessage.VerboseErrorMessage = $httpErrorObject.ErrorMessage

            $errorMessage.AuditErrorMessage = $httpErrorObject.ErrorMessage
        }

        # If error message empty, fall back on $ex.Exception.Message
        if ([String]::IsNullOrEmpty($errorMessage.VerboseErrorMessage)) {
            $errorMessage.VerboseErrorMessage = $ErrorObject.Exception.Message
        }
        if ([String]::IsNullOrEmpty($errorMessage.AuditErrorMessage)) {
            $errorMessage.AuditErrorMessage = $ErrorObject.Exception.Message
        }

        Write-Output $errorMessage
    }
}

function Invoke-HelloIDRestMethod {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Method,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Uri,

        [object]
        $Body,

        [string]
        $ContentType = "application/json",

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]
        $Headers,

        [Parameter()]
        [Boolean]
        $UsePaging = $false,

        [Parameter()]
        [Int]
        $Skip = 0,

        [Parameter()]
        [Int]
        $Take = 1000
    )

    process {
        try {
            $splatParams = @{
                Uri             = $Uri
                Headers         = $Headers
                Method          = $Method
                ContentType     = $ContentType
                UseBasicParsing = $true
                Verbose         = $false
                ErrorAction     = "Stop"
            }

            if ($Body) {
                $splatParams["Body"] = ([System.Text.Encoding]::UTF8.GetBytes($Body))
            }

            if ($UsePaging -eq $true) {
                $result = [System.Collections.ArrayList]@()
                $startUri = $splatParams.Uri
                do {
                    # Determine separator based on whether URI already contains query parameters
                    $separator = if ($startUri -match '\?') { '&' } else { '?' }
                    $splatParams["Uri"] = $startUri + "$($separator)take=$($take)&skip=$($skip)"
                    $response = (Invoke-RestMethod @splatParams)
                    if ([bool]($response.PSobject.Properties.name -eq "data")) {
                        $response = $response.data
                    }
                    if ($response -is [array]) {
                        [void]$result.AddRange($response)
                    }
                    else {
                        [void]$result.Add($response)
                    }
        
                    $skip += $take
                } while (($response | Measure-Object).Count -eq $take)
            }
            else {
                $result = Invoke-RestMethod @splatParams
            }

            Write-Output $result
        }
        catch {
            throw $_
        }
    }
}


function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions

#region script
Write-StatusMessage -Event Information -Message "Starting synchronization of Entra ID Owners to HelloID Groups"
Write-StatusMessage -Event Information -Message "-----------[Entra ID]-----------"
# Get Entra ID Groups
try {  
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $entraIDHeaders = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }


    $entraIDGroupPropertiesToRetrieve = @(
        "id"
        , "displayName"
        , "description"
        , "onPremisesSyncEnabled"
        , "groupTypes"
        , "mail"
    )
    $select = "`$select=$($entraIDGroupPropertiesToRetrieve -join ",")"
    
    # Get Entra ID groups
    $actionMessage = "querying Entra ID groups that match filter [$entraIDGroupsSearchFilter] and retrieving properties [$($entraIDGroupPropertiesToRetrieve -join ", ")]"
    $m365GroupFilter = "groupTypes/any(c:c+eq+'Unified')"
    $securityGroupFilter = "NOT(groupTypes/any(c:c+eq+'DynamicMembership')) and onPremisesSyncEnabled eq null and mailEnabled eq false and securityEnabled eq true"
    $managableGroupsFilter = "($m365GroupFilter or $securityGroupFilter)"
    # Example how to filter on teams
    # $managableGroupsFilter = "`$filter=($m365GroupFilter or $securityGroupFilter) and resourceProvisioningOptions/Any(x:x eq 'Team')"
  
    # Add additional filter if specified
    if (-not [string]::IsNullOrEmpty($entraIDGroupsSearchFilter)) {
        $managableGroupsFilter = "$managableGroupsFilter and $entraIDGroupsSearchFilter"
    }
    
    $select = "`$select=$($entraIDGroupPropertiesToRetrieve -join ",")"
    $entraIDQuerySplatParams = @{
        Uri         = "$($MSGraphBaseUri)/v1.0/groups?`$filter=$managableGroupsFilter&$select&`$top=999&`$count=true"
        Headers     = $entraIDHeaders
        Method      = 'GET'
        ErrorAction = 'Stop'
    }

    $entraIdGroups = [System.Collections.ArrayList]@()
    $getEntraIDGroupsResponse = $null
    $getEntraIDGroupsResponse = Invoke-RestMethod @entraIDQuerySplatParams -Verbose:$false
    if ($getEntraIDGroupsResponse.value -is [array]) {
        [void]$entraIdGroups.AddRange($getEntraIDGroupsResponse.value)
    }
    else {
        [void]$entraIdGroups.Add($getEntraIDGroupsResponse.value)
    }

    while (![string]::IsNullOrEmpty($getEntraIDGroupsResponse.'@odata.nextLink')) {
        $entraIDQuerySplatParams = @{
            Uri         = $getEntraIDGroupsResponse.'@odata.nextLink'
            Headers     = $entraIDHeaders
            Method      = 'GET'
            ErrorAction = 'Stop'
        }
        $getEntraIDGroupsResponse = $null
        $getEntraIDGroupsResponse = Invoke-RestMethod @entraIDQuerySplatParams -Verbose:$false
        if ($getEntraIDGroupsResponse.value -is [array]) {
            [void]$entraIdGroups.AddRange($getEntraIDGroupsResponse.value)
        }
        else {
            [void]$entraIdGroups.Add($getEntraIDGroupsResponse.value)
        }
    }

    Write-StatusMessage -Event Success -Message "Successfully queried Entra ID groups that match filter [$entraIDGroupsSearchFilter]. Result count: $(($entraIdGroups | Measure-Object).Count)"
    $entraIdGroups | Add-Member -MemberType NoteProperty -Name Users -Value $null

    $entraIdGroupsInScope = [System.Collections.Generic.List[Object]]::New()
    foreach ($entraIdGroup in $entraIdGroups) {	
        $actionMessage = "querying Entra ID group [$($entraIdGroup.displayName) ($($entraIdGroup.id))] memberships"
        $entraIDOwnerQuerySplatParams = @{
            Uri         = "$($MSGraphBaseUri)/v1.0/groups/$($entraIdGroup.id)/owners"
            Headers     = $entraIDHeaders
            Method      = 'GET'
            ErrorAction = 'Stop'
        }

        $getOwnersResponse = $null
        $getOwnersResponse = Invoke-RestMethod @entraIDOwnerQuerySplatParams -Verbose:$false
        $entraIdGroup.Users = $getOwnersResponse.value

        if ($verboseLogging -eq $true) {
            Write-Verbose "Successfully queried Entra ID group [$($entraIdGroup.displayName) ($($entraIdGroup.id))] with owners. Result count: $(($entraIdGroup.Users | Measure-Object).Count)"
        }
		        
        [void]$entraIdGroupsInScope.Add($entraIdGroup)
    }

    if (($entraIdGroupsInScope | Measure-Object).Count -eq 0) {
        throw "No Entra ID Groups have been found"
    }

    Write-StatusMessage -Event Success -Message "Successfully queried Entra ID groups with members. Result count: $(($entraIdGroupsInScope | Measure-Object).Count)"
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $errorMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $errorMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-StatusMessage -Event "Error" -Message $warningMessage
    Write-SummaryMessage -Event "Failed" -Message $errorMessage
    exit
}

Write-StatusMessage -Event Information -Message "------[HelloID]------"
#region Get HelloID
try {
    # Create authorization headers with HelloID API key
    $actionMessage = "creating authorization headers with HelloID API key"

    $pair = "$($helloIDPortalApiKey):$($helloIDPortalApiSecret)"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $helloIDHeaders = @{"authorization" = $Key }

    $actionMessage = "querying HelloID users"
    $getHelloIDUsersSplatParams = @{
        Uri       = "$($helloIDPortalBaseUrl)/api/v1/users"
        Method    = 'GET'
        Headers   = $helloIDHeaders
        UsePaging = $true
    }

    $helloIDUsers = Invoke-HelloIDRestMethod @getHelloIDUsersSplatParams
    $helloIDUsersGroupedOnUserName = $helloIDUsers | Group-Object -Property "userName" -AsHashTable -AsString
    $helloIDUsersGroupedOnUserGUID = $helloIDUsers | Group-Object -Property "userGUID" -AsHashTable -AsString

    Write-StatusMessage -Event Success -Message "Successfully queried users from HelloID. Result count: $(($helloIDUsers | Measure-Object).Count)"

    $actionMessage = "querying HelloID groups"
    $getHelloIDGroupsSplatParams = @{
        Uri       = "$($helloIDPortalBaseUrl)/api/v1/groups"
        Method    = 'GET'
        Headers   = $helloIDHeaders
        UsePaging = $true
    }
  
    $helloIDGroups = Invoke-HelloIDRestMethod @getHelloIDGroupsSplatParams

    Write-StatusMessage -Event Success -Message "Successfully queried groups from HelloID. Result count: $(($helloIDGroups | Measure-Object).Count)"

    [System.Collections.ArrayList]$helloIDGroupsWithMembers = @()
    foreach ($helloIDGroup in $helloIDGroups) {
        $actionMessage = "querying HelloID group [$($helloIDGroup.name) ($($helloIDGroup.groupGuid))] memberships"

        $getHelloIDGroupsMembershipsSplatParams = @{
            Uri       = "$($helloIDPortalBaseUrl)/api/v1/groups/$($helloIDGroup.groupGuid)"
            Method    = 'GET'
            Headers   = $helloIDHeaders
            UsePaging = $true
        }

        $helloIDGroupMembers = Invoke-HelloIDRestMethod @getHelloIDGroupsMembershipsSplatParams

        [void]$helloIDGroupsWithMembers.Add($helloIDGroupMembers)

        if ($verboseLogging -eq $true) {
            Write-Verbose "Successfully queried HelloID group [$($helloIDGroup.name) ($($helloIDGroup.groupGuid))] with members. Result count: $(($helloIDGroup.users | Measure-Object).Count)"
        }
        
    }

    $helloIDGroupsWithMembers | Add-Member -MemberType NoteProperty -Name SourceAndName -Value $null
    $helloIDGroupsWithMembers | ForEach-Object {
        if ([string]::IsNullOrEmpty($_.source)) {
            $_.source = "Local"
        }
        $_.SourceAndName = "$($_.source)/$($_.name)"
    }

    $helloIDGroupsWithMembers = $helloIDGroupsWithMembers | Where-Object { $_.SourceAndName -like "$($resourceOwnerGroupSource)/*" }
    $helloIDGroupsWithMembersGroupedBySourceAndName = $helloIDGroupsWithMembers | Group-Object -Property "SourceAndName" -AsHashTable -AsString

    Write-StatusMessage -Event Success -Message "Successfully queried HelloID groups with members. Result count: $(($helloIDGroupsWithMembers.users | Measure-Object).Count)"
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") -or
        $($ex.Exception.GetType().FullName -eq "System.Net.WebException")) {
        $errorObj = Resolve-HelloIDError -ErrorObject $ex
        $warningMessage = "Error at Line '$($errorObj.ScriptLineNumber)': $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
        $errorMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
    }
    else {
        $warningMessage = "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
        $errorMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
    }
    Write-StatusMessage -Event "Error" -Message $warningMessage
    Write-SummaryMessage -Event "Failed" -Message $errorMessage
    exit
}
#endregion Get HelloID

#endregionGet members of HelloID groups

Write-StatusMessage -Event Information -Message "------[Calculations of combined data]------"
# Calculate new and obsolete groupmemberships
try {
    # Define existing & new groupmemberships
    $actionMessage = "calculating new and obsolete groupmemberships"
    $existingGroupMembershipObjects = [System.Collections.ArrayList]@()
    $newGroupMembershipObjects = [System.Collections.ArrayList]@()
    foreach ($entraIdGroupInScope in $entraIdGroupsInScope) {
		
        # Define Resource owner Group
        $resourceOwnerGroupName = "$($resourceOwnerGroupSource)/" + "$($resourceOwnerGroupPrefix)" + "$($entraIdGroupInScope.displayName)" + "$($resourceOwnerGroupSuffix)"

        # Get HelloID Resource Owner Group
        $helloIDResourceOwnerGroup = $null
        if (-not[string]::IsNullOrEmpty($resourceOwnerGroupName)) {
            $resourceOwnerGroupName = Remove-StringLatinCharacters $resourceOwnerGroupName
            $helloIDResourceOwnerGroup = $helloIDGroupsWithMembersGroupedBySourceAndName["$($resourceOwnerGroupName)"]
            if ($null -eq $helloIDResourceOwnerGroup) {
                if ($verboseLogging -eq $true) {
                    Write-StatusMessage -Event Information -Message "Resource owner group [$($resourceOwnerGroupName)] for Group [$($entraIdGroupInScope.displayName)] not found in HelloID"
                }

                # Skip further actions for this record
                Continue
            }
        }
        else {
            if ($verboseLogging -eq $true) {
                Write-StatusMessage -Event Information -Message "No Resource owner group name provided for Group [$($entraIdGroupInScope.displayName)]"
            }
        }

        # Define existing groupmemberships
        foreach ($helloIDResourceOwnerGroupUser in $helloIDResourceOwnerGroup.Users) {
            # Get HelloID User
            $helloIDUser = $null
            $helloIDUser = $helloIDUsersGroupedOnUserGUID["$($helloIDResourceOwnerGroupUser)"]
            if ($null -eq $helloIDUser) {
                if ($verboseLogging -eq $true) {
                    Write-StatusMessage -Event Information -Message "No HelloID user found for Group Resource owner group [$($helloIDResourceOwnerGroupUser)]"
                }

                # Skip further actions for this record
                Continue
            }
            
            $existingGroupMembershipObject = [PSCustomObject]@{
                GroupName    = "$($helloIDResourceOwnerGroup.name)"
                GroupId      = "$($helloIDResourceOwnerGroup.groupGuid)"
                UserUsername = "$($helloIDUser.userName)"
                UserId       = "$($helloIDUser.userGUID)"
                CombinedId   = "$($helloIDResourceOwnerGroup.groupGuid)" + "_" + "$($helloIDUser.userGUID)"
            }
			
            [void]$existingGroupMembershipObjects.Add($existingGroupMembershipObject)
        }

        # Define new groupmemberships
        foreach ($owner in $entraIdGroupInScope.Users) {
			
            # Get HelloID User
            $helloIDUser = $null
            if (-not[string]::IsNullOrEmpty($owner.UserPrincipalName)) {
                $helloIDUser = $helloIDUsersGroupedOnUserName["$($owner.UserPrincipalName)"]
                if ($null -eq $helloIDUser) {
                    if ($verboseLogging -eq $true) {
                        Write-StatusMessage -Event Information -Message "No HelloID user found for Teams Resource owner group [$($owner.UserPrincipalName)]"
                    }

                    # Skip further actions for this record
                    Continue
                }
            }
            else {
                if ($verboseLogging -eq $true) {
                    Write-StatusMessage -Event Information -Message "No UserPrincipalName provided for full access user [$($owner.Id)]"
                }
            }
			

            $newGroupMembershipObject = [PSCustomObject]@{
                GroupName    = "$($helloIDResourceOwnerGroup.name)"
                GroupId      = "$($helloIDResourceOwnerGroup.groupGuid)"
                UserUsername = "$($helloIDUser.userName)"
                UserId       = "$($helloIDUser.userGUID)"
                CombinedId   = "$($helloIDResourceOwnerGroup.groupGuid)" + "_" + "$($helloIDUser.userGUID)"
            }

            [void]$newGroupMembershipObjects.Add($newGroupMembershipObject)
        }
    }

    # Define new groupmemberships
    $newGroupMemberships = [System.Collections.ArrayList]@()
    $newGroupMemberships = $newGroupMembershipObjects | Where-Object { $_.CombinedId -notin $existingGroupMembershipObjects.CombinedId }

    # Define obsolete groupmemberships
    $obsoleteGroupMemberships = [System.Collections.ArrayList]@()
    $obsoleteGroupMemberships = $existingGroupMembershipObjects | Where-Object { $_.CombinedId -notin $newGroupMembershipObjects.CombinedId }

    # Define existing groupmemberships
    $existingGroupMemberships = [System.Collections.ArrayList]@()
    $existingGroupMemberships = $existingGroupMembershipObjects | Where-Object { $_.CombinedId -notin $obsoleteGroupMemberships.CombinedId }

    # Define total groupmemberships (existing + new)
    $totalGroupMemberships = ($(($existingGroupMemberships | Measure-Object).Count) + $(($newGroupMemberships | Measure-Object).Count))
}
catch {
    $ex = $PSItem
    $errorMessage = Get-ErrorMessage -ErrorObject $ex
    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"
    $errorMessage = "Error $($actionMessage). Error: $($errorMessage.AuditErrorMessage)"
    Write-StatusMessage -Event "Error" -Message $warningMessage
    Write-SummaryMessage -Event "Failed" -Message $errorMessage
    exit
}

Write-StatusMessage -Event Information -Message "------[Summary]------"
Write-StatusMessage -Event Information -Message "Total EntraId Groups in scope [$(($entraIdGroupsInScope.Users | Measure-Object).Count)]"

Write-StatusMessage -Event Information -Message "New HelloID Resource Owner Groupmembership(s) that will be granted [$(($newGroupMemberships | Measure-Object).Count)]"

if ($removeMembers) {
    Write-StatusMessage -Event Information "Obsolete HelloID Resource Owner Groupmembership(s) that will be revoked [$(($obsoleteGroupMemberships | Measure-Object).Count)]"
}
else {
    Write-StatusMessage -Event Information -Message "Obsolete HelloID Resource Owner Groupmembership(s) that won't be revoked [$(($obsoleteGroupMemberships | Measure-Object).Count)]"
}

Write-StatusMessage -Event Information -Message "------[Processing]------------------"
try {
    $addUserToGroupSuccess = 0
    $addUserToGroupError = 0
    foreach ($newGroupMembership in $newGroupMemberships) {
        # Add HelloID User to HelloID Group
        $actionMessage = "adding memberships to group [$($newGroupMembership.GroupId)]"

        if ($verboseLogging -eq $true) {
            Write-Verbose "Adding HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]"
        }

        $addUserToGroupBody = [PSCustomObject]@{
            UserGUID = "$($newGroupMembership.UserId)"
        }
        $body = ($addUserToGroupBody | ConvertTo-Json -Depth 10)

        $postHelloIDAddMembershipsSplatParams = @{
            Uri     = "$($helloIDPortalBaseUrl)/api/v1/groups/$($newGroupMembership.GroupId)/users"
            Method  = 'POST'
            Headers = $helloIDHeaders
            Body    = $body
        }

        if ($dryRun -eq $false) {
            $addUserToGroupResult = Invoke-HelloIDRestMethod @postHelloIDAddMembershipsSplatParams
            $addUserToGroupSuccess++

            if ($verboseLogging -eq $true) {
                Write-Verbose "Successfully added HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]"
            }
        }
        else {
            if ($verboseLogging -eq $true) {
                Write-StatusMessage -Event Information -Message "DryRun: Would add HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]"
            }
        }
    }
    if ($dryRun -eq $false) {
        if ($addUserToGroupSuccess -ge 1 -or $addUserToGroupError -ge 1) {
            Write-StatusMessage -Event Information -Message "Added HelloID users to HelloID groups. Success: $($addUserToGroupSuccess). Error: $($addUserToGroupError)"
            Write-SummaryMessage -Event Information -Message "Added HelloID users to HelloID groups. Success: $($addUserToGroupSuccess). Error: $($addUserToGroupError)"
        }
    }
    else {
        Write-StatusMessage -Event Warning -Message "DryRun: Would add [$(($newGroupMemberships | Measure-Object).Count)] HelloID users to HelloID groups"
    }

    if ($removeMembers -eq $true) {
        $removeUserFromGroupSuccess = 0
        $removeUserFromGroupError = 0
        foreach ($obsoleteGroupMembership in $obsoleteGroupMemberships) {
            # Remove HelloID User from HelloID Group
            $actionMessage = "removing user [$($obsoleteGroupMembership.UserId)] from group [$($obsoleteGroupMembership.GroupId)]"

            if ($verboseLogging -eq $true) {
                Write-Verbose "Removing HelloID user [$($obsoleteGroupMembership.UserUsername) ($($obsoleteGroupMembership.UserId))] to HelloID group [$($obsoleteGroupMembership.GroupName) ($($obsoleteGroupMembership.GroupId))]"
            }

            $deleteHelloIDRemoveMembershipsSplatParams = @{
                Uri     = "$($helloIDPortalBaseUrl)/api/v1/groups/$($obsoleteGroupMembership.GroupId)/users/$($obsoleteGroupMembership.UserId)"
                Method  = 'DELETE'
                Headers = $helloIDHeaders
            }

            if ($dryRun -eq $false) {
                $removeUserToGroupResult = Invoke-HelloIDRestMethod @deleteHelloIDRemoveMembershipsSplatParams
                $removeUserFromGroupSuccess++

                if ($verboseLogging -eq $true) {
                    Write-Verbose "Successfully removed HelloID user [$($obsoleteGroupMembership.UserUsername) ($($obsoleteGroupMembership.UserId))] to HelloID group [$($obsoleteGroupMembership.GroupName) ($($obsoleteGroupMembership.GroupId))]"
                }
            }
            else {
                if ($verboseLogging -eq $true) {
                    Write-StatusMessage -Event Information -Message "DryRun: Would remove HelloID user [$($obsoleteGroupMembership.UserUsername) ($($obsoleteGroupMembership.UserId))] to HelloID group [$($obsoleteGroupMembership.GroupName) ($($obsoleteGroupMembership.GroupId))]"
                }
            }
        }
        if ($dryRun -eq $false) {
            if ($removeUserFromGroupSuccess -ge 1 -or $removeUserFromGroupError -ge 1) {
                Write-StatusMessage -Event Information -Message "Removed HelloID users from HelloID groups. Success: $($removeUserFromGroupSuccess). Error: $($removeUserFromGroupError)"
                Write-SummaryMessage -Event Information -Message "Removed HelloID users from HelloID groups. Success: $($removeUserFromGroupSuccess). Error: $($removeUserFromGroupError)"
            }
        }
        else {
            Write-StatusMessage -Event Warning -Message "DryRun: Would remove [$(($obsoleteGroupMemberships | Measure-Object).Count)] HelloID users from HelloID groups"
        }
    }
    else {
        Write-StatusMessage -Event Warning -Message "Option to remove members is set to [$removeMembers]. Skipped removing [$(($obsoleteGroupMemberships | Measure-Object).Count)] HelloID users to HelloID groups"
    }

    if ($dryRun -eq $false) {
        Write-StatusMessage -Event Success -Message "Successfully synchronized [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Groups to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
        Write-SummaryMessage -Event Success -Message "Successfully synchronized [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Groups to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
    }
    else {
        Write-StatusMessage -Event Success -Message "DryRun: Would synchronize [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Groups to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
        Write-SummaryMessage -Event Success -Message "DryRun: Would synchronize [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Groups to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") -or
        $($ex.Exception.GetType().FullName -eq "System.Net.WebException")) {
        $errorObj = Resolve-HelloIDError -ErrorObject $ex
        $warningMessage = "Error at Line '$($errorObj.ScriptLineNumber)': $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
        $errorMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
    }
    else {
        $warningMessage = "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
        $errorMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
    }
    Write-StatusMessage -Event "Error" -Message $warningMessage
    Write-SummaryMessage -Event "Failed" -Message $errorMessage
    Write-StatusMessage -Event Error -Message "Error synchronization of [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Groups to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
    Write-SummaryMessage -Event Failed -Message "Error synchronization of [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Groups to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
}
#endregion