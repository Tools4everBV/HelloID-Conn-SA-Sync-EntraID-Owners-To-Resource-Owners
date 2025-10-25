#####################################################
# HelloID-Conn-SA-Sync-Teams-Owners-To-ResourceOwner-Groupmemberships
#
# Version: 1.0.0
#####################################################
# Set to false to acutally perform actions - Only run as DryRun when testing/troubleshooting!
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

# Make sure to create the Global variables defined below in HelloID
#HelloID Connection Configuration
# $script:PortalBaseUrl = "" # Set from Global Variable
# $portalApiKey = "" # Set from Global Variable
# $portalApiSecret = "" # Set from Global Variable

######################################################################################
# Entra ID Connection Configuration
$MSGraphBaseUri = "https://graph.microsoft.com/" # Fixed value
# $EntraTenantId = "" # Set from Global Variable
# $EntraAppID = "" # Set from Global Variable
# $EntraAppSecret = "" # Set from Global Variable
$entraIDGroupsSearchFilter = "`$search=`"displayName:department_`"" # Optional, when no filter is provided ($entraIDGroupsSearchFilter = $null), all groups will be queried - Only displayName and description are supported with the search filter. Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter?tabs=http#using-search-on-directory-object-collections
######################################################################################

#HelloID Configuration
$resourceOwnerGroupSource = "Local" # Specify the source of the groups - if source is any other than "Local", the sync of the target system itself might overwrite the memberships set form this sync
# The HelloID Resource owner group will be queried based on the shared mailbox name and the specified prefix and suffix
$resourceOwnerGroupPrefix = "Team eigenaar " # Specify prefix to recognize the resource owner group
$resourceOwnerGroupSuffix = "" # Specify suffix to recognize the resource owner group
$removeMembers = $true # If true, existing members will be removed if they no longer have full access to the corresponding mailbox - This will overwrite manual added users

#region functions

function Remove-StringLatinCharacters {
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
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

function Invoke-HIDRestmethod {
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

        [Parameter(Mandatory = $false)]
        $PageSize,

        [string]
        $ContentType = "application/json"
    )

    try {
        Write-Verbose "Switching to TLS 1.2"
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

        Write-Verbose "Setting authorization headers"
        $apiKeySecret = "$($portalApiKey):$($portalApiSecret)"
        $base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($apiKeySecret))
        $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
        $headers.Add("Authorization", "Basic $base64")
        $headers.Add("Content-Type", $ContentType)
        $headers.Add("Accept", $ContentType)

        $splatWebRequest = @{
            Uri             = "$($script:PortalBaseUrl)/api/v1/$($Uri)"
            Headers         = $headers
            Method          = $Method
            UseBasicParsing = $true
            ErrorAction     = "Stop"
        }
        
        if (-not[String]::IsNullOrEmpty($PageSize)) {
            $data = [System.Collections.ArrayList]@()

            $skip = 0
            $take = $PageSize
            Do {
                $splatWebRequest["Uri"] = "$($script:PortalBaseUrl)/api/v1/$($Uri)?skip=$($skip)&take=$($take)"

                Write-Verbose "Invoking [$Method] request to [$Uri]"
                $response = $null
                $response = Invoke-RestMethod @splatWebRequest -Verbose:$false
                if (($response.PsObject.Properties.Match("pageData") | Measure-Object).Count -gt 0) {
                    $dataset = $response.pageData
                }
                else {
                    $dataset = $response
                }

                if ($dataset -is [array]) {
                    [void]$data.AddRange($dataset)
                }
                else {
                    [void]$data.Add($dataset)
                }
            
                $skip += $take
            }until(($dataset | Measure-Object).Count -ne $take)

            return $data
        }
        else {
            if ($Body) {
                Write-Verbose "Adding body to request"
                $splatWebRequest["Body"] = ([System.Text.Encoding]::UTF8.GetBytes($body))
            }

            Write-Verbose "Invoking [$Method] request to [$Uri]"
            $response = $null
            $response = Invoke-RestMethod @splatWebRequest -Verbose:$false

            return $response
        }

    }
    catch {
        throw $_
    }
}


function New-AuthorizationHeaders {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[[String], [String]]])]
    param(
        [parameter(Mandatory)]
        [string]
        $TenantId,

        [parameter(Mandatory)]
        [string]
        $ClientId,

        [parameter(Mandatory)]
        [string]
        $ClientSecret
    )
    try {
        Write-Verbose "Creating Access Token"
        $authUri = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
    
        $body = @{
            grant_type    = "client_credentials"
            client_id     = "$ClientId"
            client_secret = "$ClientSecret"
            resource      = "https://graph.microsoft.com"
        }
    
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
        $accessToken = $Response.access_token
    
        #Add the authorization header to the request
        Write-Verbose 'Adding Authorization headers'

        $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
        $headers.Add('Authorization', "Bearer $accesstoken")
        $headers.Add('Accept', 'application/json')
        $headers.Add('Content-Type', 'application/json')
        # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
        $headers.Add('ConsistencyLevel', 'eventual')

        Write-Output $headers  
    }
    catch {
        throw $_
    }
}
#endregion functions

#region script
Hid-Write-Status -Event Information -Message "Starting synchronization of Entra ID to HelloID Self service Producs"
Hid-Write-Status -Event Information -Message "-----------[Entra ID]-----------"
# Get Entra ID Groups
try {  
    $headers = New-AuthorizationHeaders -TenantId $EntraTenantId -ClientId $EntraAppID -ClientSecret $EntraAppSecret

    $properties = @(
        "id"
        , "displayName"
        , "description"
        , "onPremisesSyncEnabled"
        , "groupTypes"
        , "mail"
    )
    $select = "`$select=$($properties -join ",")"

    $m365GroupFilter = "groupTypes/any(c:c+eq+'Unified')"
    $securityGroupFilter = "NOT(groupTypes/any(c:c+eq+'DynamicMembership')) and onPremisesSyncEnabled eq null and mailEnabled eq false and securityEnabled eq true"
    $managableGroupsFilter = "`$filter=($m365GroupFilter or $securityGroupFilter) and resourceProvisioningOptions/Any(x:x eq 'Team')"
  
    $entraIDQuerySplatParams = @{
        Uri         = "$($MSGraphBaseUri)/v1.0/groups?$managableGroupsFilter and $entraIDGroupsSearchFilter&$select&`$top=999&`$count=true"
        Headers     = $headers
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
            Headers     = $headers
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

	$entraIdGroups | Add-Member -MemberType NoteProperty -Name Users -Value $null

    $entraIdGroupsInScope = [System.Collections.Generic.List[Object]]::New()
    foreach ($entraIdGroup in $entraIdGroups) {
		
		#Hid-Write-Status -Event Information -Message "Searching owners for $($entraIdGroup.displayName) with id $($entraIdGroup.id)"
		
		$entraIDOwnerQuerySplatParams = @{
			Uri         = "$($MSGraphBaseUri)/v1.0/groups/$($entraIdGroup.id)/owners"
			Headers     = $headers
			Method      = 'GET'
			ErrorAction = 'Stop'
		}

		#$entraIdGroups = [System.Collections.ArrayList]@()
		$getOwnersResponse = $null
		$getOwnersResponse = Invoke-RestMethod @entraIDOwnerQuerySplatParams -Verbose:$false
		
		$entraIdGroup.Users = $getOwnersResponse.value
		
        # Custom - Only process groups with a description
        if ([string]::IsNullOrEmpty($entraIdGroup.description)) {
            if ($verboseLogging -eq $true) {
                Write-Verbose "No description set in Entra ID for Entra ID group [$($entraIdGroup)]"
            }

        }
        else {
            [void]$entraIdGroupsInScope.Add($entraIdGroup)
        }
    }

    if (($entraIdGroupsInScope | Measure-Object).Count -eq 0) {
        throw "No Entra ID Groups have been found"
    }

    Hid-Write-Status -Event Success -Message "Successfully queried Entra ID groups. Result count: $(($entraIdGroupsInScope | Measure-Object).Count)"
}
catch {
    $ex = $PSItem
    $errorMessage = Get-ErrorMessage -ErrorObject $ex

    Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

    throw "Error querying Entra ID groups that match filter [$($entraIDGroupsSearchFilter)]. Error Message: $($errorMessage.AuditErrorMessage)"
}

Hid-Write-Status -Event Information -Message "------[HelloID]------"
#region Get HelloID Users
try {
    Write-Verbose "Querying Users from HelloID"

    $splatWebRequest = @{
        Method   = "GET"
        Uri      = "users"
        PageSize = 1000
    }
    $helloIDUsers = Invoke-HIDRestMethod @splatWebRequest

    $helloIDUsersGroupedOnUserName = $helloIDUsers | Group-Object -Property "userName" -AsHashTable -AsString
    $helloIDUsersGroupedOnUserGUID = $helloIDUsers | Group-Object -Property "userGUID" -AsHashTable -AsString

    Hid-Write-Status -Event Success -Message "Successfully queried Users from HelloID. Result count: $(($helloIDUsers | Measure-Object).Count)"
}
catch {
    $ex = $PSItem
    $errorMessage = Get-ErrorMessage -ErrorObject $ex

    Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

    throw "Error querying Users from HelloID. Error Message: $($errorMessage.AuditErrorMessage)"
}
#endregion Get HelloID Users

#region Get HelloID Groups
try {
    Write-Verbose "Querying Groups from HelloID"

    $splatWebRequest = @{
        Method   = "GET"
        Uri      = "groups"
        PageSize = 1000
    }
    $helloIDGroups = Invoke-HIDRestMethod @splatWebRequest

    Hid-Write-Status -Event Success -Message "Successfully queried Groups from HelloID. Result count: $(($helloIDGroups | Measure-Object).Count)"
}
catch {
    $ex = $PSItem
    $errorMessage = Get-ErrorMessage -ErrorObject $ex

    Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

    throw "Error querying Groups from HelloID. Error Message: $($errorMessage.AuditErrorMessage)"
}
#endregion Get HelloID Groups

#region Get members of HelloID groups
try {
    [System.Collections.ArrayList]$helloIDGroupsWithMembers = @()
    Write-Verbose "Querying HelloID groups with members"
    foreach ($helloIDGroup in $helloIDGroups) {
        #region Get HelloID users that are member of HelloID group
        try {
            Write-Verbose "Querying HelloID group [$($helloIDGroup.name) ($($helloIDGroup.groupGuid))] with members"

            $splatWebRequest = @{
                Method   = "GET"
                Uri      = "groups/$($helloIDGroup.groupGuid)"
                PageSize = 1000
            }
            $helloIDGroup = Invoke-HIDRestMethod @splatWebRequest

            [void]$helloIDGroupsWithMembers.Add($helloIDGroup)

            if ($verboseLogging -eq $true) {
                Write-Verbose "Successfully queried HelloID group [$($helloIDGroup.name) ($($helloIDGroup.groupGuid))] with members. Result count: $(($helloIDGroup.users | Measure-Object).Count)"
            }
        }
        catch {
            $ex = $PSItem
            $errorMessage = Get-ErrorMessage -ErrorObject $ex
        
            Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"
        
            throw "Error querying HelloID group [$($helloIDGroup.name) ($($helloIDGroup.groupGuid))] with members. Error Message: $($errorMessage.AuditErrorMessage)"
        }
        #endregion Get HelloID users that are member of HelloID group
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

    Hid-Write-Status -Event Success -Message "Successfully queried HelloID groups with members. Result count: $(($helloIDGroupsWithMembers.users | Measure-Object).Count)"
}
catch {
    $ex = $PSItem
    $errorMessage = Get-ErrorMessage -ErrorObject $ex

    Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

    throw "Error querying HelloID users that are member of HelloID groups. Error Message: $($errorMessage.AuditErrorMessage)"
}
#endregionGet members of HelloID groups

Hid-Write-Status -Event Information -Message "------[Calculations of combined data]------"
# Calculate new and obsolete groupmemberships
try {
    # Define existing & new groupmemberships
    $existingGroupMembershipObjects = [System.Collections.ArrayList]@()
    $newGroupMembershipObjects = [System.Collections.ArrayList]@()
    foreach ($entraIdGroupInScope in $entraIdGroupsInScope) {
		
        # Define Resource owner Group
        $resourceOwnerGroupName = "$($resourceOwnerGroupSource)/" + "$($resourceOwnerGroupPrefix)" + ($entraIdGroupInScope.mail).split("@")[0] + "$($resourceOwnerGroupSuffix)"

        # Get HelloID Resource Owner Group
        $helloIDResourceOwnerGroup = $null
        if (-not[string]::IsNullOrEmpty($resourceOwnerGroupName)) {
            $resourceOwnerGroupName = Remove-StringLatinCharacters $resourceOwnerGroupName
            $helloIDResourceOwnerGroup = $helloIDGroupsWithMembersGroupedBySourceAndName["$($resourceOwnerGroupName)"]
            if ($null -eq $helloIDResourceOwnerGroup) {
                if ($verboseLogging -eq $true) {
                    Hid-Write-Status -Event Information -Message "Resource owner group [$($resourceOwnerGroupName)] for Shared Mailbox [$($entraIdGroupInScope.displayName)] not found in HelloID"
                }

                # Skip further actions for this record
                Continue
            }
        }
        else {
            if ($verboseLogging -eq $true) {
                Hid-Write-Status -Event Information -Message "No Resource owner group name provided for Team [$($entraIdGroupInScope.displayName)]"
            }
        }

        # Define existing groupmemberships
        foreach ($helloIDResourceOwnerGroupUser in $helloIDResourceOwnerGroup.Users) {
            # Get HelloID User
            $helloIDUser = $null
            $helloIDUser = $helloIDUsersGroupedOnUserGUID["$($helloIDResourceOwnerGroupUser)"]
            if ($null -eq $helloIDUser) {
                if ($verboseLogging -eq $true) {
                    Hid-Write-Status -Event Information -Message "No HelloID user found for Teams Resource owner group [$($helloIDResourceOwnerGroupUser)]"
                }

                # Skip further actions for this record
                Continue
            }
            
            $existingGroupMembershipObject = [PSCustomObject]@{
                GroupName    = "$($helloIDResourceOwnerGroup.name)"
                GroupId      = "$($helloIDResourceOwnerGroup.groupGuid)"
                UserUsername = "$($helloIDUser.userName)"
                UserId       = "$($helloIDUser.userGUID)"
				CombinedId  	 = "$($helloIDResourceOwnerGroup.groupGuid)" + "_" + "$($helloIDUser.userGUID)"
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
                        Hid-Write-Status -Event Information -Message "No HelloID user found for Teams Resource owner group [$($owner.UserPrincipalName)]"
                    }

                    # Skip further actions for this record
                    Continue
                }
            }
            else {
                if ($verboseLogging -eq $true) {
                    Hid-Write-Status -Event Information -Message "No UserPrincipalName provided for full access user [$($owner.Id)]"
                }
            }
			

            $newGroupMembershipObject = [PSCustomObject]@{
                GroupName    = "$($helloIDResourceOwnerGroup.name)"
                GroupId      = "$($helloIDResourceOwnerGroup.groupGuid)"
                UserUsername = "$($helloIDUser.userName)"
                UserId       = "$($helloIDUser.userGUID)"
				CombinedId  	 = "$($helloIDResourceOwnerGroup.groupGuid)" + "_" + "$($helloIDUser.userGUID)"
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

    Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

    throw "Error calculating new and obsolete groupmemberships. Error Message: $($errorMessage.AuditErrorMessage)"
}

Hid-Write-Status -Event Information -Message "------[Summary]------"
Hid-Write-Status -Event Information -Message "Total EntraId Groups in scope [$(($entraIdGroupsInScope.Users | Measure-Object).Count)]"

Hid-Write-Status -Event Information -Message "New HelloID Resource Owner Groupmembership(s) that will be granted [$(($newGroupMemberships | Measure-Object).Count)]"

if ($removeMembers) {
    Hid-Write-Status -Event Information "Obsolete HelloID Resource Owner Groupmembership(s) that will be revoked [$(($obsoleteGroupMemberships | Measure-Object).Count)]"
}
else {
    Hid-Write-Status -Event Information -Message "Obsolete HelloID Resource Owner Groupmembership(s) that won't be revoked [$(($obsoleteGroupMemberships | Measure-Object).Count)]"
}

Hid-Write-Status -Event Information -Message "------[Processing]------------------"
try {
    $addUserToGroupSuccess = 0
    $addUserToGroupError = 0
    foreach ($newGroupMembership in $newGroupMemberships) {
        # Add HelloID User to HelloID Group
        try {
            if ($verboseLogging -eq $true) {
                Write-Verbose "Adding HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]"
            }

            $addUserToGroupBody = [PSCustomObject]@{
                UserGUID = "$($newGroupMembership.UserId)"
            }
            $body = ($addUserToGroupBody | ConvertTo-Json -Depth 10)
            $splatWebRequest = @{
                Uri    = "groups/$($newGroupMembership.GroupId)/users"
                Method = 'POST'
                Body   = $body
            }

            if ($dryRun -eq $false) {
                $addUserToGroupResult = Invoke-HIDRestMethod @splatWebRequest
                $addUserToGroupSuccess++

                if ($verboseLogging -eq $true) {
                    Write-Verbose "Successfully added HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]"
                }
            }
            else {
                if ($verboseLogging -eq $true) {
                    Write-Verbose "DryRun: Would add HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]"
					Hid-Write-Status -Event Information -Message "DryRun: Would add HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]"
                }
            }
        }
        catch {
            $addUserToGroupError++

            $ex = $PSItem
            $errorMessage = Get-ErrorMessage -ErrorObject $ex
            
            Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"
            
            throw "Error adding HelloID user [$($newGroupMembership.UserUsername) ($($newGroupMembership.UserId))] to HelloID group [$($newGroupMembership.GroupName) ($($newGroupMembership.GroupId))]. Error Message: $($errorMessage.AuditErrorMessage)"
        }
    }
    if ($dryRun -eq $false) {
        if ($addUserToGroupSuccess -ge 1 -or $addUserToGroupError -ge 1) {
            Hid-Write-Status -Event Information -Message "Added HelloID users to HelloID groups. Success: $($addUserToGroupSuccess). Error: $($addUserToGroupError)"
            Hid-Write-Summary -Event Information -Message "Added HelloID users to HelloID groups. Success: $($addUserToGroupSuccess). Error: $($addUserToGroupError)"
        }
    }
    else {
        Hid-Write-Status -Event Warning -Message "DryRun: Would add [$(($newGroupMemberships | Measure-Object).Count)] HelloID users to HelloID groups"
    }

    if ($removeMembers -eq $true) {
        $removeUserFromGroupSuccess = 0
        $removeUserFromGroupError = 0
        foreach ($obsoleteGroupMembership in $obsoleteGroupMemberships) {
            # Remove HelloID User from HelloID Group
            try {
                if ($verboseLogging -eq $true) {
                    Write-Verbose "Removing HelloID user [$($obsoleteGroupMembership.UserUsername) ($($obsoleteGroupMembership.UserId))] to HelloID group [$($obsoleteGroupMembership.GroupName) ($($obsoleteGroupMembership.GroupId))]"
                }

                $splatWebRequest = @{
                    Uri    = "groups/$($obsoleteGroupMembership.GroupId)/users/$($obsoleteGroupMembership.UserId)"
                    Method = 'DELETE'
                }

                if ($dryRun -eq $false) {
                    $removeUserToGroupResult = Invoke-HIDRestMethod @splatWebRequest
                    $removeUserFromGroupSuccess++

                    if ($verboseLogging -eq $true) {
                        Write-Verbose "Successfully removed HelloID user [$($obsoleteGroupMembership.UserUsername) ($($obsoleteGroupMembership.UserId))] to HelloID group [$($obsoleteGroupMembership.GroupName) ($($obsoleteGroupMembership.GroupId))]"
                    }
                }
                else {
                    if ($verboseLogging -eq $true) {
                        Write-Verbose "DryRun: Would remove HelloID user [$($obsoleteGroupMembership.UserUsername) ($($obsoleteGroupMembership.UserId))] to HelloID group [$($obsoleteGroupMembership.GroupName) ($($obsoleteGroupMembership.GroupId))]"
                    }
                }
            }
            catch {
                $removeUserFromGroupError++

                $ex = $PSItem
                $errorMessage = Get-ErrorMessage -ErrorObject $ex
            
                Hid-Write-Status -Event Error -Message "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"
            
                throw "Error removing HelloID user [$($obsoleteGroupMembership.UserUsername) ($($obsoleteGroupMembership.UserId))] to HelloID group [$($obsoleteGroupMembership.GroupName) ($($obsoleteGroupMembership.GroupId))]. Error Message: $($errorMessage.AuditErrorMessage)"
            }
        }
        if ($dryRun -eq $false) {
            if ($removeUserFromGroupSuccess -ge 1 -or $removeUserFromGroupError -ge 1) {
                Hid-Write-Status -Event Information -Message "Removed HelloID users from HelloID groups. Success: $($removeUserFromGroupSuccess). Error: $($removeUserFromGroupError)"
                Hid-Write-Summary -Event Information -Message "Removed HelloID users from HelloID groups. Success: $($removeUserFromGroupSuccess). Error: $($removeUserFromGroupError)"
            }
        }
        else {
            Hid-Write-Status -Event Warning -Message "DryRun: Would remove [$(($obsoleteGroupMemberships | Measure-Object).Count)] HelloID users from HelloID groups"
        }
    }
    else {
        Hid-Write-Status -Event Warning -Message "Option to remove members is set to [$removeMembers]. Skipped removing [$(($obsoleteGroupMemberships | Measure-Object).Count)] HelloID users to HelloID groups"
    }

    if ($dryRun -eq $false) {
        Hid-Write-Status -Event Success -Message "Successfully synchronized [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Teams to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
        Hid-Write-Summary -Event Success -Message "Successfully synchronized [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Teams to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
    }
    else {
        Hid-Write-Status -Event Success -Message "DryRun: Would synchronize [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Teams to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
        Hid-Write-Summary -Event Success -Message "DryRun: Would synchronize [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Teams to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
    }
}
catch {
    Hid-Write-Status -Event Error -Message "Error synchronization of [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Teams to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
    Hid-Write-Status -Event Error -Message "Error at Line [$($_.InvocationInfo.ScriptLineNumber)]: $($_.InvocationInfo.Line)."
    Hid-Write-Status -Event Error -Message "Exception message: $($_.Exception.Message)"
    Hid-Write-Status -Event Error -Message "Exception details: $($_.errordetails)"
    Hid-Write-Summary -Event Failed -Message "Error synchronization of [$(($entraIdGroupsInScope.Users | Measure-Object).Count)] Owners of Teams to [$totalGroupMemberships] HelloID ResourceOwner Groupmemberships"
}
#endregion