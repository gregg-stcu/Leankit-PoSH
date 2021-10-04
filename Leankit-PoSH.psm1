#requires -Version 3.0
$script:leankiturl = $null
$script:headers = $null
$script:LKUsers = $null
$script:LKBoards = $null
$script:cachePath = $null
$script:config = $null



# tell powershell not to default to insecure protocols!
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Set-LKConfig {
    param
    (
        #Provide a fullpath\Filename for the config file (recommended config.json)
        [Parameter(Mandatory = $true,
            ParameterSetName = 'ManualConfig')]
        [Parameter(Mandatory = $true,
            ParameterSetName = 'PassthruConfig')]
        [String]
        $LKconfigFile,
        [Parameter(Mandatory = $true,
            ParameterSetName = 'PassthruConfig')]
        [string]
        $cacheFolder,
        [Parameter(Mandatory = $true,
            ParameterSetName = 'PassthruConfig')]
        [pscredential]
        $LKCreds,
        [Parameter(Mandatory = $true,
            ParameterSetName = 'PassthruConfig')]
        [string]
        $commentsynclimit,
        [Parameter(Mandatory = $true,
            ParameterSetName = 'PassthruConfig')]
        [string]
        $leankitDomain
    )
    if ((Test-Path -Path $LKconfigFile) -and $null -eq $LKCreds) {
        $script:config = Get-Content -Path $LKconfigFile | ConvertFrom-Json
        $CacheFolder = Read-Host -Prompt "Cache Folder (Default: $($script:config.cacheFolder))"
        if (!($CacheFolder)) {
            $CacheFolder = $script:config.cacheFolder
        }
        $keepCreds = Read-Host -Prompt "Keep existing LK Token? (y/n)"
        $LKToken = $script:config.LKToken
        if ($keepCreds.ToLower() -eq 'n') {
            $LeankitIntegrationAccount = Get-Credential -Message "Enter Leankit Credentials"
        }
        $leankitDomain = Read-Host -Prompt "Leankit Domain (Default: $($script:config.LKURL))"
        if (!($leankitDomain)) {
            $leankitDomain = ($script:config.LKURL -split '/')[2]
        }
        $commentsynclimit = Read-Host -Prompt "Comment retrieval size default (default: $($script:config.commentsynclimit))"
        if (!($commentsynclimit)) {
            $commentsynclimit = $script:config.commentsynclimit
        }
    }
    elseif (!$LKCreds) {
        $CacheFolder = Read-Host -Prompt "Cache Folder"
        $LeankitIntegrationAccount = Get-Credential -Message "Enter Leankit Credentials"
        $leankitDomain = Read-Host -Prompt "Leankit Domain (ex:company.leankit.com)"
        $commentsynclimit = read-host -Prompt "How many comments will be retreived at a time by default (recommended: 5)"
    }
    else {
        if ($LKCreds -and $leankitDomain -and $commentsynclimit -and $cacheFolder) {
            $LeankitIntegrationAccount = $LKCreds
        }
        else {
            throw "Config file doesn't exist, you must provide LKCreds, Leankitdomain, cachefolder and commentsynclimit variables"
        }
    }

    #Get LK token and save
    if ($LeankitIntegrationAccount) {
        $creds = "$($LeankitIntegrationAccount.username):$($LeankitIntegrationAccount.GetNetworkCredential().Password)"

        $encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($creds))
        $basicAuthValue = "Basic $encodedCreds"
        $headers = @{
            Authorization = $basicAuthValue
        }
        $command = @{
            description = $env:COMPUTERNAME
        }

        $jsonCommand = $command | ConvertTo-Json

        try {
            $Token = (Invoke-RestMethod -Headers $headers -Method Post -Body "$jsonCommand" -Uri "https://$LeankitDomain/io/auth/token" -ContentType 'application/json').token
        }
        catch {
            $errortext = ConvertFrom-Json -InputObject $_
            throw "$($errortext.message)"
        }
        $LKToken = (ConvertTo-SecureString $token -AsPlainText -Force | ConvertFrom-SecureString)
    }
    try {
        @{
            cacheFolder      = $cacheFolder
            LKURL            = "https://$LeankitDomain/io"
            LKToken          = $LKToken
            commentsynclimit = $commentsynclimit
        } | ConvertTo-Json | Out-File -FilePath $LKconfigFile
    }
    catch {
        throw $_
    }

}

function Initialize-LKToken {
    param (
        [parameter(Mandatory = $true,
            ParameterSetName = 'NoConfigFile')]
        [string]
        $cachePath,
        [parameter(Mandatory = $true,
            ParameterSetName = 'NoConfigFile')]
        [hashtable]
        $headers,
        [parameter(Mandatory = $true,
            ParameterSetName = 'NoConfigFile')]
        [string]
        $URL,
        [parameter(Mandatory = $true,
            ParameterSetName = 'ConfigFile')]
        [string]
        $configFile
    )
    if (Test-Path -Path $configFile) {
        $config = Get-Content -Path $configfile | ConvertFrom-Json
        $script:Config = $config
        $script:cachePath = $config.cacheFolder
        $script:leankiturl = $config.LKURL
        #set LK headers
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR(($config.LKToken | ConvertTo-SecureString))
        $LKToken = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $tokenhead = "Bearer $LKToken"
        $script:headers = @{
            Authorization = $tokenhead

        }
        #test LK token
        try {
            Invoke-RestMethod -Headers $script:headers -Method Get -Uri "$($config.LKURL)/auth/token" -ContentType 'application/json' | Out-Null
        }
        catch {
            if (($_.ToString() | ConvertFrom-Json).message -eq "Unauthorized") {
                Write-Error -Message "Bad LK Token in config file ($configfile), run set-lkconfig to fix!"
            }
        }

    }
    else {

        if ($headers) {
            $script:leankiturl = $URL
            $script:cachePath = $cachePath
            $script:headers = $headers
        }

    }

}

function Get-LKBoardList {
    try {
        $boards = (Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/board" -ContentType 'application/json').boards
        return $boards
    }
    catch {
        $errortext = ConvertFrom-Json -InputObject $_
        return "$($errortext.message)"
    }

}
function Get-LKCardByID {
    param
    (
        [String]
        [Parameter(Mandatory)]
        $CardID
    )

    $card = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/$CardID" -ContentType 'application/json'

    return $card
}

function Get-LKtaskCardsByParent {
    param
    (
        [String]
        [Parameter(Mandatory)]
        $CardID
    )

    $cards = (Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/$CardID/tasks" -ContentType 'application/json').cards

    return $cards
}

function Expand-LKParentCard {
    #This function takes all task cards within a card and makes them children. If you set the 'setHeaders' variable to true it will copy the header from the parent to all children.
    param
    (
        [String]
        [Parameter(Mandatory)]
        $Parent,

        [bool]
        $setHeaders = $false
    )

    $tasks = Get-LKtaskCardsByParent -CardID $Parent
    $ParentCard = (Get-LKCardByID -CardID $Parent)
    $lane = $ParentCard.lane.id

    $taskIDarray = $tasks | ForEach-Object -Process {
        $_.id
    }

    $command = @{
        cardIds            = @($taskIDarray)
        destination        = @{
            laneId = "$lane"
        }
        wipOverrideComment = 'Parent Card Expansion'
    }

    $jsonCommand = $command | ConvertTo-Json
    try {
        if ($setHeaders -and ($null -ne $ParentCard.customId.value)) {
            $result = Set-LKCardHeader -CardIDs $taskIDarray -HeaderText $ParentCard.customId.value
        }
        $result = Invoke-RestMethod -Headers $script:headers -Method Post -Uri "$leankiturl/card/move" -Body $jsonCommand -ContentType 'application/json'
    }
    catch {
        $errortext = ConvertFrom-Json -InputObject $_
        return "$($errortext.message)"
    }
    $result = Update-LKParentCard -CardIDs $taskIDarray -ParentID $Parent
    return $result
}

function Update-LKParentCard {
    #This function takes all card IDs passed in and associates them to the parent card ID as a child.

    param
    (
        [array]
        [Parameter(Mandatory)]
        $CardIDs,

        [String]
        [Parameter(Mandatory)]
        $ParentID
    )

    $command = @{
        cardId           = $ParentID
        connectedCardIds = @($CardIDs)
    }

    $jsonCommand = ConvertTo-Json -InputObject $command
    $response = Invoke-RestMethod -Headers $script:headers -Method Post -Uri "$leankiturl/card/$ParentID/connection/many" -Body $jsonCommand -ContentType 'application/json'

    return $response
}


function Set-LKCardFinishDate {
    #This function sets the Card Finish Date on a list of cards
    param
    (
        [array]
        [Parameter(Mandatory)]
        $CardIDs,

        [string]
        $date = $null
    )
    $response = @()
    $operation = 'replace'
    if ($date -eq '') {
        $finishDate = 'null'
        $operation = 'remove'
    }
    else {
        $finishDate = "$(([datetime]$date).ToString('yyyy-MM-dd'))"
    }
    $i = 0
    foreach ($card in $CardIDs) {
        $command = @{
            op    = "$operation"
            path  = '/plannedFinish'
            value = "$finishDate"
        }
        $cardupdaterequest = @($command)
        $jsonCommand = ConvertTo-Json -InputObject $cardupdaterequest

        $i++
        Write-Progress -activity "Setting Finish Date" -status "Card: $card" -percentComplete (($i / $CardIDs.Count) * 100)
        $jsonCommand = ConvertTo-Json -InputObject $cardupdaterequest

        try {
            # Content
            $response += Invoke-RestMethod -Headers $script:headers -Method Patch -Uri "$leankiturl/card/$($card.id)" -Body $jsonCommand -ContentType 'application/json'
        }
        catch {
            #$errortext = ConvertFrom-Json -InputObject $_
            Write-Output $_
            return
        }
    }
    return $response
}
function Set-LKCardStartDate {
    #This Function sets the Start Date on a list of cards
    param
    (
        [array]
        [Parameter(Mandatory)]
        $CardIDs,

        [string]
        $Date = $null
    )
    $response = @()
    $operation = 'replace'
    if ($Date -eq '') {
        $startDate = 'null'
        $operation = 'remove'
    }
    else {
        $startDate = "$(([datetime]$Date).ToString('yyyy-MM-dd'))"
    }
    $i = 0
    foreach ($card in $CardIDs) {
        $command = @{
            op    = $operation
            path  = '/plannedStart'
            value = $startDate
        }
        $cardupdaterequest = @($command)
        $i++
        Write-Progress -activity "Setting Start Date" -status "Current Card: $($card.id)" -percentComplete (($i / $CardIDs.Count) * 100)
        try {
            # Content
            $jsonCommand = ConvertTo-Json -InputObject $cardupdaterequest
            $response += Invoke-RestMethod -Headers $script:headers -Method Patch -Uri "$leankiturl/card/$($card.id)" -Body $jsonCommand -ContentType 'application/json'
        }
        catch {
            $errortext = ConvertFrom-Json -InputObject $_
            return "$($errortext.message)"
        }
    }
    return $response
}
function Set-LKCardHeader {
    param
    (
        [array]
        [Parameter(Mandatory)]
        $CardIDs,

        [String]
        [Parameter(Mandatory)]
        $HeaderText
    )
    $response = @()


    foreach ($card in $CardIDs) {
        $command = @{
            op    = 'replace'
            path  = '/customId'
            value = "$HeaderText"
        }
        Start-Sleep -Milliseconds ($CardIDs.Count * 5)

        $cardupdaterequest = @($command)
        $jsonCommand = ConvertTo-Json -InputObject $cardupdaterequest
        try {
            # Content
            $response += Invoke-RestMethod -Headers $script:headers -Method Patch -Uri "$leankiturl/card/$card" -Body $jsonCommand -ContentType 'application/json'
        }
        catch {
            $errortext = ConvertFrom-Json -InputObject $_
            $errortext
            return "$($errortext.message)"
        }
    }
    #return $response
}

function Get-LKCardsOnBoard {
    param
    (
        [String]
        [Parameter(Mandatory)]
        $BoardID,

        [ValidateSet('active', 'backlog', 'archive')]
        [string]$laneClass = 'active',
        <#
        Valid Fields:
        assignedUsers
        id
        index
        version
        title
        description
        priority
        size
        plannedStart
        plannedFinish
        actualFinish
        actualStart
        createdOn
        archivedOn
        updatedOn
        movedOn
        tags
        color
        iconPath
        customIconLabel
        customIcon
        blockedStatus
        board
        externalLinks
        lane
        type
    #>
        [array]
        $fields = $null,
        [int]
        $limit = 2000,
        [int]
        $offset = 0,
        [switch]
        $includeTaskCards,
        [string]
        $search = ""

    )
    if ($includeTaskCards) {
        $select = 'both'
    }
    else {
        $select = 'cards'
    }
    if ($null -eq $fields) {
        $cards = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/?board=$BoardID&search=`"$search`"&lane_class_types=$laneClass&limit=$limit&offset=$offset&select=$select" -ContentType 'application/json'
    }
    else {
        $fields | ForEach-Object -Process {
            $fieldString += ($(if ($fieldString) {
                        ', '
                    }
                ) + $_)
        }
        $cards = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/?board=$BoardID&search=`"$search`"&lane_class_types=$laneClass&only=$fieldString&limit=$limit&offset=$offset&select=$select" -ContentType 'application/json'
    }

    return $cards.cards
}


function Get-LKCardsUpdatedSince {
    param
    (
        [String]
        $BoardID,
        [ValidateSet('active', 'backlog', 'archive')]
        [string]
        $laneClass = 'active',
        [array]
        $fields = $null,
        [string]
        $since,
        [string]
        $lanes = ""
    )
    if (!$since) {
        [string]$since = Get-CurrentDateTime -startOfDay $true
    }
    $lastupdate = $since
    if ($BoardID) {
        $Board = "board=$boardID&"
    }
    if ($null -eq $fields) {
        $cards = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/?$($Board)lanes=$lanes&lane_class_types=$laneClass&since=$since&limit=2000" -ContentType 'application/json'
    }
    else {
        $fields | ForEach-Object -Process {
            $fieldString += ($(if ($fieldString) {
                        ', '
                    }
                ) + $_)
        }
        $fieldString += ', updatedOn'
        $cards = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/?$($Board)lanes=$lanes&lane_class_types=$laneClass&since=$since&only=$fieldString&limit=2000" -ContentType 'application/json'
    }
    $lastupdate = $cards.cards[0].updatedOn

    $output = [psobject]@{
        cards      = $cards.cards
        lastUpdate = $lastupdate
    }
    return $output
}

function Get-CurrentDateTime {
    param
    (
        [bool]
        [Parameter(Mandatory)]
        $startOfDay

    )
    if ($startOfDay) {
        return (Get-Date -Format 'yyyy-MM-ddT00:00:00Z')
    }
    return (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')
}

function Move-LKCard {
    #This function will move a card or set of cards to a specified lane
    param
    (
        [array]
        [Parameter(Mandatory)]
        $CardID,

        [String]
        [Parameter(Mandatory)]
        $DestinationLane,

        [String]
        $WipOverrideReason = '',
        [int]
        $index
    )

    $response = @()
    if ($WipOverrideReason -eq '') {
        #No wip override


        if ($index) {
            $command = @{
                cardIds     = $CardID
                destination = @{
                    laneId = $DestinationLane
                    index  = $index
                }
            }
        }
        else {
            $command = @{
                cardIds     = $CardID
                destination = @{
                    laneId = $DestinationLane
                }
            }
        }

        $cardupdaterequest = $command

        $jsonCommand = ConvertTo-Json -InputObject $cardupdaterequest
        try {
            # Attempt to move card (fail if wip limit reached)
            $response = Invoke-RestMethod -Headers $script:headers -Method Post -Uri "$leankiturl/card/move" -Body $jsonCommand -ContentType 'application/json'
        }
        catch {
            $errortext = ConvertFrom-Json -InputObject $_
            return "$($errortext.message)"
        }
    }
    else {
        $command = @{
            cardIds            = @($CardID)
            destination        = @{
                laneId = $DestinationLane
            }
            wipOverrideComment = "$WipOverrideReason"
        }


        $cardupdaterequest = $command

        $jsonCommand = ConvertTo-Json -InputObject $cardupdaterequest
        try {
            # Attempt moving card with override
            $response = Invoke-RestMethod -Headers $script:headers -Method Post -Uri "$leankiturl/card/move" -Body $jsonCommand -ContentType 'application/json'
        }
        catch {
            $errortext = ConvertFrom-Json -InputObject $_
            return "$($errortext.message)"
        }
    }
    return $response
}
function Get-LKCardsInLane {
    param
    (
        [String]
        [Parameter(Mandatory)]
        $LaneID,
        [String]
        $BoardID,
        [array]
        $fields = $null,
        [int]
        $limit = 2000,
        [int]
        $offset = 0,
        [switch]
        $includeTaskCards,
        [string]
        $search = "",
        [ValidateSet('active', 'backlog', 'archive')]
        [string]$laneClass = ""
    )
    if ($includeTaskCards) {
        $select = 'both'
    }
    else {
        $select = 'cards'
    }
    if ($null -eq $fields) {
        $cards = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/?lanes=$LaneID&search=`"$search`"&lane_class_types=$laneClass&limit=$limit&offset=$offset&select=$select" -ContentType 'application/json'
    }
    else {
        $fields | ForEach-Object -Process {
            $fieldString += ($(if ($fieldString) {
                        ', '
                    }
                ) + $_)
        }
        $cards = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/?lanes=$LaneID&search=`"$search`"&lane_class_types=$laneClass&only=$fieldString&limit=$limit&offset=$offset&select=$select" -ContentType 'application/json'
    }
    return $cards.cards

}

function Set-LKCardDatesinLane {
    param
    (
        [String]
        [Parameter(Mandatory)]
        $LaneID,
        [string]
        [Parameter(Mandatory)]
        $BoardID,
        [string]
        $startDate = $null,
        [string]
        $finishDate = $null,
        [bool]
        $overwriteExistingFinishDate = $false,
        [bool]
        $overwriteExistingStartDate = $false
    )
    $Cards = Get-LKCardsInLane -LaneID $LaneID -BoardID $BoardID -fields id, plannedFinish, plannedStart | Select-Object -Property id, plannedFinish, plannedStart
    $CardIDs = $Cards | Select-Object -Property id

    if ($overwriteExistingFinishDate) {
        $null = Set-LKCardFinishDate -CardIDs $CardIDs -Date $finishDate
    }
    else {
        $filteredCards = $Cards | Where-Object -FilterScript {
            $null -eq $_.plannedFinish
        } | Select-Object -Property id
        if ($null -ne $filteredCards) {
            $null = Set-LKCardFinishDate -CardIDs $filteredCards -Date $finishDate
        }
    }
    if ($overwriteExistingStartDate) {
        $null = Set-LKCardStartDate -CardIDs $CardIDs -Date $startDate
    }
    else {
        $filteredCards = $Cards | Where-Object -FilterScript {
            ($null -eq $_.plannedStart)
        } | Select-Object -Property id
        if ($null -ne $filteredCards) {
            $null = Set-LKCardStartDate -CardIDs $filteredCards -Date $startDate
        }
    }
}

function Get-LKPointsForLane {
    param
    (
        [String]
        [Parameter(Mandatory)]
        $BoardID,
        [array]
        [Parameter(Mandatory)]
        $LaneIDs,
        [int]
        $DaysLeft = 10
    )
    $cards = @()
    foreach ($lane in $LaneIDs) {
        $cards += Get-LKCardsInLane -LaneID $lane -BoardID $BoardID -fields 'size', 'assignedUsers'
    }
    foreach ($card in $cards) {
        if (!$card.size) {
            $card.size = 1
        }
    }
    $TotalsPerUser = $cards |
    Select-Object -ExpandProperty assignedUsers -Property size |
    Group-Object -Property fullname |
    ForEach-Object -Process {
        New-Object -Property @{
            Item  = $_.Name
            Sum   = ($_.Group |
                Measure-Object -Property size -Sum).Sum
            Cards = ($_.Group |
                Measure-Object -Property size).Count
        } -TypeName PSObject
    }

    $Totals = @()

    foreach ($user in $TotalsPerUser) {
        #Write-Output "$($user.Item): has $($user.Sum) total points which is $($($user.Sum)/$DaysLeft) per day."
        $list = $null
        $list = New-Object -TypeName System.Object
        $list | Add-Member -TypeName NoteProperty -Name 'User' -Value $user.Item -MemberType NoteProperty
        $list | Add-Member -TypeName NoteProperty -Name 'Total Points' -Value $user.Sum -MemberType NoteProperty
        $list | Add-Member -TypeName NoteProperty -Name 'Average Points/Day' -Value ($user.sum / $DaysLeft).ToString("#.##") -MemberType NoteProperty
        $list | Add-Member -TypeName NoteProperty -Name 'Average Cards/Day' -Value ($user.Cards / $DaysLeft).ToString("#.##") -MemberType NoteProperty
        $Totals += $list
    }

    return $Totals
}

function Get-LKUsersOnBoard {
    #This should return all defined users on a board.
    param
    (
        [String]
        [Parameter(Mandatory)]
        $BoardID
    )

    $board = Get-LKBoard -BoardID $BoardID
    return $board.users
}

function Get-LKBoard {
    param
    (
        [String]
        [Parameter(Mandatory)]
        $BoardID
    )
    $board = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/board/$BoardID" -ContentType 'application/json'

    return $board
}

function Get-LKTokensForCurrentUser {
    $TokenList = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/auth/token" -ContentType 'application/json'
    return $TokenList.tokens
}

function revoke-LKTokensForCurrentUser {
    param
    (
        [array]
        [Parameter(Mandatory)]
        $TokenList
    )

    foreach ($Token in $TokenList) {
        Invoke-RestMethod -Headers $script:headers -Method DELETE -Uri "$leankiturl/auth/token/$Token" -ContentType 'application/json'
    }
}

function Get-LKCardComments {
    Param(
        # Parameter help description
        [Parameter(Mandatory)]
        [string]
        $CardID,
        [string]
        $limit = '',
        [validateset('newest', 'oldest')]
        [string]
        $sortby = 'newest'
    )
    $cards = Invoke-RestMethod -Headers $script:headers -Method Get -Uri "$leankiturl/card/$CardID/comment?limit=$limit&sortBy=$sortby" -ContentType 'application/json'

    if ($cards.comments) {
        return $cards.comments
    }
    else {
        return
    }
}

function get-LKAllUsers {
    $users = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/user" -ContentType 'application/json'

    return $users.users

}

function New-LKCard {
    Param(
        [parameter(Mandatory)]
        [string]
        $boardID,
        [string]
        $laneID,
        [parameter(Mandatory)]
        [string]
        $title,
        [string]
        $typeID,
        [array]
        $assignedUser,
        [string]
        $description,
        [string]
        $WipOverrideReason,
        [string]
        $plannedStart,
        [string]
        $plannedFinish,
        [string]
        $customId,
        [array]
        $tags,
        [array]
        $customFields,
        [hashtable]
        $externalLink,
        [string]
        $priority
    )
    $command = @{
        boardId = $boardID
        title   = $title
    }
    if ([array]$assignedUser.count -eq 1) {
        $command.Add('assignedUserIds', @($assignedUser))
    }
    elseif ([array]$assignedUser.Count -gt 1) {
        $command.Add('assignedUserIds', $assignedUser)
    }
    if ($typeID) {
        $command.Add('typeId', $typeID)
    }
    if ($laneID) {
        $command.Add('laneId', $laneID)
    }
    if ($description) {
        $command.Add('description', $description)
    }
    if ($WipOverrideReason) {
        $command.Add('wipOverrideComment', $WipOverrideReason)
    }
    if ($plannedStart) {
        $command.Add('plannedStart', $plannedStart)
    }
    if ($plannedFinish) {
        $command.Add('plannedFinish', $plannedFinish)
    }
    if ($customId) {
        $command.Add('customId', $customId)
    }
    if ($tags) {
        $command.Add('tags', $tags)
    }
    if ($customFields) {
        $command.add('customFields', $customFields)
    }
    if ($externalLink) {
        $command.add('externalLink', $externalLink)
    }
    if ($priority) {
        $command.add('priority', $priority)
    }

    $jsonCommand = ConvertTo-Json -InputObject $command

    $response = Invoke-RestMethod -Headers $script:headers -Body $jsonCommand -Method Post -Uri "$leankiturl/card" -ContentType 'application/json'
    return $response
}

function Get-LKUserID {
    $response = Invoke-RestMethod -Headers $script:headers -Method Get -Uri "$leankiturl/user/me" -ContentType 'application/json'
    return $response.id
}

function Get-LKLanes {
    param (
        [parameter(Mandatory)]
        [string]
        $boardID
    )

    $response = Invoke-RestMethod -Headers $script:headers -Method Get -Uri "$leankiturl/board/$boardID" -ContentType 'application/json'
    return $response.lanes
}

function Get-LKUpdatedCards {
    param (
        [parameter(Mandatory)]
        [string]
        $boardID,
        [parameter(Mandatory)]
        [string]
        $StartTime,
        [array]
        $fields = @('id', 'title', 'updatedOn', 'customId')
    )
    if ($fields) {
        $fields | ForEach-Object -Process {
            $fieldString += ($(if ($fieldString) {
                        ', '
                    }
                ) + $_)
        }
    }
    $response = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card?since=$StartTime&board=$boardID&only=$fieldstring"
    return $response.cards
}

function Get-LKNewComments {
    param (
        [parameter(Mandatory)]
        [string]
        $boardID,
        [parameter(Mandatory)]
        [string]
        $StartTime,
        [string]
        $filterstring,
        [switch]
        $includeTaskCards,
        [array]
        $updatedCards
    )
    $comments = @()
    if ($updatedCards) {
        $cards = $updatedCards
    }
    else {
        $cards = Get-LKUpdatedCards -boardID $boardID -StartTime $StartTime
    }
    ForEach ($card in $cards) {
        if ($filterstring) {

            $commentstoAdd = (Get-LKCardComments -CardID $card.id | Where-Object -FilterScript { $_.createdOn -gt $StartTime -and $_.text -imatch $filterstring -and $_.createdBy.id -ne '879254403' })
            $comments += @([pscustomobject]@{
                    CardID     = $card.id
                    Title      = $card.title
                    lastUpdate = $card.updatedOn
                    customId   = $card.customId.value
                    Comments   = $commentstoAdd
                })
            if ($includeTaskCards) {
                $tasks = get-LKTaskCards -cardID $card.id
                foreach ($task in $tasks) {
                    $comments += @([PSCustomObject]@{
                            CardID     = $task.id
                            Title      = $task.title
                            lastUpdate = $card.updatedOn
                            customId   = $task.customId.value
                            Comments   = $commentstoAdd
                        })
                }
            }
        }
        else {
            $commentstoAdd = (Get-LKCardComments -CardID $card.id | Where-Object -FilterScript { $_.createdOn -gt $StartTime })
            $comments += @([pscustomobject]@{
                    CardID     = $card.id
                    Title      = $card.title
                    lastUpdate = $card.updatedOn
                    customId   = $card.customId.value
                    Comments   = $commentstoAdd
                })
            if ($card.customId.value -contains 'Ticket') {
                $tasks = get-LKTaskCards -cardID $card.id
                foreach ($task in $tasks) {
                    $comments += @([PSCustomObject]@{
                            CardID     = $task.id
                            Title      = $task.title
                            lastUpdate = $card.updatedOn
                            customId   = $task.customId.value
                            Comments   = $commentstoAdd
                        })
                }
            }
        }
    }

    return $comments | Where-Object -FilterScript { $null -ne $_.Comments }

}

function New-LKComment {
    param (
        [parameter(Mandatory)]
        [string]
        $CardID,
        [parameter(Mandatory)]
        [string]
        $Comment
    )
    $command = @{
        text = $Comment
    }

    $jsonCommand = ConvertTo-Json -InputObject $command

    $response = Invoke-RestMethod -Headers $script:headers -Body $jsonCommand -Method Post -Uri "$leankiturl/card/$cardID/comment" -ContentType 'application/json'
    return $response
}

function get-LKTaskCards {
    param (
        [parameter(Mandatory)]
        [string]
        $cardID
    )
    $response = Invoke-RestMethod -Headers $script:headers -Method get -Uri "$leankiturl/card/$cardID/tasks"
    return $response.cards

}

function get-LKTicketIDs {
    param (
        [parameter(Mandatory)]
        [string]
        $boardID
    )
    $result = @()
    $ActiveCards = @(Get-LKCardsOnBoard -BoardID $boardID -laneClass active -fields id, customId, containingCardId, parentcards -search "Ticket")
    $ActiveCards += @(Get-LKCardsOnBoard -BoardID $boardID -laneClass backlog -fields id, customId, containingCardId, parentcards -search "Ticket")

    $ActiveCards = $ActiveCards | Where-Object -FilterScript { $_.customId.value -imatch 'ticket [0-9]+' }

    foreach ($activecard in $ActiveCards) {
        $result += @([PSCustomObject] @{
                CardID      = $activecard.id
                TicketID    = $activecard.customId.value -replace '[^0-9]', ''
                parentcards = $activecard.parentcards
            })

    }
    return $result

}

function Set-LKCardIcon {
    param (
        [parameter(Mandatory)]
        [string]
        $cardIDs,
        [parameter(Mandatory)]
        [string]
        $iconID
    )
    $response = @()


    foreach ($card in $CardIDs) {
        $command = @{
            op    = 'replace'
            path  = '/customIconId'
            value = "$iconID"
        }
        $cardupdaterequest = @($command)
        $jsonCommand = ConvertTo-Json -InputObject $cardupdaterequest
        try {
            # Content
            $response += Invoke-RestMethod -Headers $script:headers -Method Patch -Uri "$leankiturl/card/$card" -Body $jsonCommand -ContentType 'application/json'
        }
        catch {
            $errortext = ConvertFrom-Json -InputObject $_
            Write-Verbose "$($errortext.message)"
        }
    }
}

function optimize-LKCardOrder {
    param (

        [Parameter(Mandatory)]
        [string]
        $boardID,
        [Parameter(Mandatory)]
        [string]
        $laneID,
        [ValidateSet('plannedStart', 'plannedFinish')]
        $sortby = 'plannedStart'
    )
    if ($sortby -eq 'plannedStart') {
        $cards = Get-LKCardsInLane -BoardID $boardID -LaneID $laneID -fields 'id', 'plannedStart', 'plannedFinish', 'index' | Where-Object -FilterScript { $_.plannedStart } | Sort-Object -Property plannedStart, index
    }
    else {
        $cards = Get-LKCardsInLane -BoardID $boardID -LaneID $laneID -fields 'id', 'plannedStart', 'plannedFinish', 'index' | Where-Object -FilterScript { $_.plannedFinish } | Sort-Object -Property plannedFinish, index
    }
    [int]$index = 0
    foreach ($card in $cards) {
        if ($index -ne $card.index) {
            Move-LKCard -index $index -DestinationLane $laneID -CardID $card.id | Out-Null
        }
        $index += 1
    }

}

function get-LKCardsbyTicketNumber {
    param (
        [string]
        $boardID,
        [string]
        $TicketID,
        [array]
        $fields
    )

    $cards = @(Get-LKCardsOnBoard -BoardID $boardID -search "Ticket $TicketID" -fields $fields -includeTaskCards)
    if (!($cards)) {
        $cards = @(Get-LKCardsOnBoard -BoardID $boardID -search "Ticket $TicketID" -fields $fields -includeTaskCards -laneClass archive)
    }if (!($cards)) {
        $cards = @(Get-LKCardsOnBoard -BoardID $boardID -search "Ticket $TicketID" -fields $fields -includeTaskCards -laneClass backlog)
    }
    return $cards
}

function get-LKOpenTicketsDoneInLeankit {
    param (
        [string]
        $boardID,
        [array]
        $HelpstarQueues,
        [parameter(Mandatory)]
        [string]
        $SQLInstance,
        [parameter(Mandatory)]
        [string]
        $databaseName
    )

    $cards = @()
    $sqlQueues = $HelpstarQueues | ForEach-Object { "'$_'" }
    $sqlQueues = $sqlQueues -join ','
    $sqlSelect = "SELECT tblservicerequest.id as RequestNumber,[title],[status],[closedby],[name] FROM [HS2000CS].[dbo].[tblservicerequest] JOIN HS2000CS.dbo.tblqueue on tblservicerequest.queueid = tblqueue.id where tblqueue.name in ($sqlQueues) and [status] != '6'"
    Write-Verbose "SQL: $sqlselect"
    $openTickets = Invoke-sqlcommand -dataSource $SQLInstance -database $databaseName -sqlCommand $sqlSelect

    foreach ($openTicket in $openTickets) {
        $newCard = get-LKCardsbyTicketNumber -fields 'id', 'customId', 'actualFinish', 'title', 'assignedUsers' -boardID $boardID -TicketID $openTicket.RequestNumber
        if ($newCard) {
            $cards += [PSCustomObject]@{
                id           = @($newCard.id)[0]
                TicketNumber = @($newCard.customId.value -replace 'ticket ', '')[0]
                actualFinish = @($newCard.actualFinish)[0]
                title        = @($newCard.Title)[0]
            }

        }
        else {
            $cards += [PSCustomObject]@{
                id           = 'None'
                TicketNumber = "$($openTicket.RequestNumber)"
                actualFinish = $null
                title        = "$($openTicket.title)"
            }
        }
        $newCard = $null
    }
    return $cards | Where-Object { $_.id -eq 'None' -or $_.actualFinish }
}

function Move-LKProjectCard {
    param (
        [string]
        $lastrun,
        [string]
        $board,
        [string]
        $projectSourceLanes,
        [string]
        $droplane
    )
    <#
    $lastrun = '2020-06-10T20:01:12Z'
    $board = '747564708'
    $projectSourceLanes = '976056916'
    $droplane = '884212840'
    #>
    $updatedCards = Get-LKCardsUpdatedSince -BoardID $board -since $lastrun -fields 'actualFinish', 'id', 'tags'
    $completedcards = @($updatedCards.cards | Where-Object -FilterScript { $null -ne $_.actualFinish })
    $completedcards = @($completedcards | Where-Object -FilterScript { [datetime]$_.actualFinish -gt [datetime]$lastrun } )
    $projectTags = @($completedcards | Select-Object -Property tags -Unique) | Select-Object -ExpandProperty tags | Where-Object -FilterScript { $_ -match 'project-' }

    $projectCards = Get-LKCardsInLane -LaneID $projectSourceLanes -fields 'id', 'tags', 'index'
    foreach ($tag in $projectTags) {
        $nextCard = @($projectCards | Where-Object -FilterScript { $_.tags -contains $tag } | Sort-Object -Property index)
        if ($nextCard) {
            Move-LKCard -CardID $nextCard[0].id -DestinationLane $droplane -WipOverrideReason 'Next Project Task' | Out-Null
        }
    }
    if ($updatedCards.lastUpdate) {
        return $updatedCards.lastUpdate
    }
    else {
        return $lastrun
    }


}

function sync-cachedDataLK {
    # SingleDatatype to Sync
    param(
        [validateset('boards', 'users')]
        [string]
        $DatatypetoSync
    )

    switch ($DatatypetoSync) {
        'boards' {
            $script:LKboards = (Get-LKBoard -BoardID " ").boards
            $script:LKboards | ConvertTo-Json -Depth 10 | Out-File -Path "$script:cachePath\LKboards.json"

            break
        }
        'users' {
            $script:LKUsers = get-LKAllUsers
            $script:LKUsers | ConvertTo-Json -Depth 10 | Out-File "$script:cachePath\LKusers.json"
            break
        }

        Default {
            if (!(Test-Path -Path "$path\LKboards.json")) {
                $script:LKboards = (Get-LKBoard -BoardID " ").boards
                $script:LKboards | ConvertTo-Json -Depth 10 | Out-File -Path "$script:cachePath\LKboards.json"
            }
            else {
                $script:LKboards = Get-Content -Path "$script:cachePath\LKboards.json" | ConvertFrom-Json -Depth 10
            }
            if (!(Test-Path -path "$path\LKusers.json")) {
                $script:LKUsers = get-LKAllUsers
                $script:LKUsers | ConvertTo-Json -Depth 10 | Out-File "$script:cachePath\LKusers.json"
            }
            else {
                $script:LKUsers = Get-Content -Path "$script:cachePath\LKusers.json" | ConvertFrom-Json -Depth 10
            }
        }
    }
}

function Get-LKBoardbyProperty {
    param (
        [Parameter(Mandatory)]
        [validateset('title', 'id')]
        [string]
        $property,
        [parameter(Mandatory)]
        $valueToMatch
    )
    if (!($script:LKBoards)) {
        sync-CachedDataLK -path $script:cachePath
        $board = $script:LKBoards | Where-Object -FilterScript { $_.$property -eq $valueToMatch }
    }
    else {
        $board = $script:LKBoards | Where-Object -FilterScript { $_.$property -eq $valueToMatch }
        if (!($board)) {
            sync-CachedDataLK -path $script:cachePath -DatatypetoSync boards
            $board = $script:LKBoards | Where-Object -FilterScript { $_.$property -eq $valueToMatch }
        }
    }
    if (!($board)) {
        throw "Board doesn't exist"
    }
    return $board

}

function get-LKuserbyProperty {
    param (
        [Parameter(Mandatory)]
        [validateset('emailaddress', 'id')]
        [string]
        $property,
        [parameter(Mandatory)]
        $valueToMatch
    )

    if (!($script:LKUsers)) {
        sync-CachedDataLK -path $script:cachePath
        $user = $script:LKUsers | Where-Object -FilterScript { $_.$property -eq $valueToMatch }
    }
    else {
        $user = $script:LKUsers | Where-Object -FilterScript { $_.$property -eq $valueToMatch }
        if (!($user)) {
            sync-CachedDataLK -path $script:cachePath -DatatypetoSync users
            $user = $script:LKUsers | Where-Object -FilterScript { $_.$property -eq $valueToMatch }
        }
    }
    if (!($user)) {
        throw "User doesn't exist"
    }

    return $user
}

function remove-lkcardbyID {
    param (
        [parameter(Mandatory)]
        [string]
        $cardID
    )

    $response = Invoke-RestMethod -Headers $script:headers -Method Delete -Uri "$leankiturl/card/$cardID" -ContentType 'application/json'
    return $response

}

function update-lkCardProperties {
    param (
        [parameter(Mandatory)]
        [System.Object]
        $properties,
        [parameter(Mandatory)]
        [string]
        $cardID
    )
    <#
    Example:
    $properties = @(@{
            op    = 'replace'
            path  = '/customIconId'
            value = "$iconID"
        })
    Example of all fields that can be updated can be found here: https://success.planview.com/Planview_LeanKit/LeanKit_API/01_v2/card/update
    #>
    $jsonCommand = ConvertTo-Json -InputObject $properties
    $response = Invoke-RestMethod -Headers $script:headers -Method Patch -Body $jsonCommand -Uri "$leankiturl/card/$cardID" -ContentType 'application/json'
    return $response

}

function move-Lkcardtoboard {
    param (
        [parameter(Mandatory)]
        [string]
        $cardID,
        [parameter(Mandatory)]
        [string]
        $destinationBoardID
    )

    $command = @{
        cardIds     = @($cardID)
        destination = @{
            boardId = $destinationBoardID
        }
    }
    $jsonCommand = ConvertTo-Json -InputObject $command

    $response = Invoke-RestMethod -Headers $script:headers -Method post -Body $jsonCommand -Uri "$leankiturl/card/move" -ContentType 'application/json'
    return $response
}



Export-ModuleMember -Function *-LK*