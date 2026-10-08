#Tom Turner in WEbD Class


#region Dependencies
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
#endregion

#region options
$TEAMVISIBILITY = "Private"
$CHANNELMEMBERSHIPTYPE = "Private"
$script:currentLoginId = $null
$script:currentTeams = @()
#endregion


#Startup functions are used at the beginning to setup current session.
#region StartupFunctions

#function for installing necessary Powershell modules.
Function Install-TeamsPowerShell {
    if (Get-Module -ListAvailable -Name "MicrosoftTeams") {
        Write-Host "Powershell Microsoft Teams already installed" -ForegroundColor Green
    }
    else {
        Write-Host "Microsoft Teams powershell module not installed. installing..." -ForegroundColor DarkYellow
        if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
        }
        Install-Module -Name MicrosoftTeams -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module MicrosoftTeams
}

#function to connect to teams, prompting for credentials.
Function Connect-Teams {
    Write-Host "Initiating Teams connection..." -ForegroundColor DarkYellow
    $connection = Connect-MicrosoftTeams -ErrorAction Stop
    $script:currentLoginId = $connection.Account.Id
    Write-Host "Proceeding with user $script:currentLoginId"
    Update-CurrentTeams
}

#refreshes the cached list of teams the current user belongs to.
Function Update-CurrentTeams {
    $script:currentTeams = @(Get-Team -User $script:currentLoginId)
    Write-Host "User has $($script:currentTeams.Count) teams"
}

#Primary function for menu controlling
Function Show-Menu {
    do {
        Write-Host "###################################################"
        Write-Host "## Tools for Importing CSV to Microsoft Teams    ##"
        Write-Host "###################################################"

        Write-Host "What would you like to do now?" -ForegroundColor Green
        Write-Host "1: Import CSV to teams"
        Write-Host "2: Import CSV to Channel"
        Write-Host "3: Copy Members between Teams"
        Write-Host "4: Copy Members between Channels"
        Write-Host "5: Export Team Members to CSV"
        Write-Host "6: Remove All Team Members"
        Write-Host "[q]  Quit " -ForegroundColor Yellow
        Write-Host "Select 1-6 or q: " -NoNewline -ForegroundColor Green
        $answer = (Read-Host).Trim()

        try {
            switch ($answer) {
                '1' { Add-UsersToTeam -WithConfirm (Read-ConfirmEachPreference) }
                '2' { Add-UsersToTeamChannel -WithConfirm (Read-ConfirmEachPreference) }
                '3' { Invoke-TeamCopy }
                '4' { Invoke-ChannelCopy }
                '5' { Export-TeamMembers }
                '6' { Remove-AllTeamMembers }
                'q' { }
                default { Write-Host "Invalid Selection!" -BackgroundColor Black -ForegroundColor Red }
            }
        }
        catch {
            Write-Host "Something went wrong: $($_.Exception.Message)" -ForegroundColor Red
        }
    } until ($answer -eq 'q')
}

#endregion




#utility functions are reusable, multipurpose functions.
#region UtilityFunctions

#asks a yes/no question on the console. Anything other than y/yes/true counts as no.
Function Read-YesNo {
    Param (
        [Parameter(Mandatory, Position = 0)]
        [string]
        $Prompt
    )
    Write-Host "$Prompt Y or N: " -ForegroundColor Green -NoNewline
    (Read-Host).Trim() -in @("y", "yes", "true")
}

#asks whether every student addition should be confirmed individually.
Function Read-ConfirmEachPreference {
    if (Read-YesNo "Do you wish to confirm every addition?") {
        Write-Host "You WILL be asked before every student" -ForegroundColor DarkCyan
        return $true
    }
    Write-Host "You will NOT be asked before every student"
    $false
}

#shows a popup text box and returns what was typed (empty string if cancelled).
Function Read-InputBox {
    Param (
        [Parameter(Mandatory)]
        [string]
        $Prompt,
        [string]
        $Title = ""
    )
    [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title)
}

Function Write-OperationAborted {
    Write-Host "Operation Aborted!`n" -ForegroundColor Cyan
}

# a function that is passed objects and displays their DisplayName in a list to pick from.
# returns the selected object, or nothing if the dialog is cancelled.
Function Select-FromList {
    Param (
        [Parameter(ValueFromPipeline = $true, Position = 0)]
        [object[]]
        $InputObject,
        [string]
        $Instructions = "Make a selection"
    )
    BEGIN {
        $items = [System.Collections.Generic.List[object]]::new()
    }
    PROCESS {
        foreach ($item in $InputObject) { $items.Add($item) }
    }
    END {
        if ($items.Count -eq 0) {
            Write-Host "Nothing available to select." -ForegroundColor Red
            return
        }

        $form = New-Object System.Windows.Forms.Form -Property @{
            Text            = $Instructions
            StartPosition   = 'CenterScreen'
            ClientSize      = '380,400'
            FormBorderStyle = 'FixedDialog'
            MaximizeBox     = $false
            MinimizeBox     = $false
            TopMost         = $true
        }
        $lblInstructions = New-Object System.Windows.Forms.Label -Property @{
            Text     = $Instructions
            Location = '10,10'
            Size     = '360,20'
        }
        $listBox = New-Object System.Windows.Forms.ListBox -Property @{
            Location = '10,35'
            Size     = '360,320'
        }
        foreach ($item in $items) { [void]$listBox.Items.Add("$($item.DisplayName)") }
        $listBox.SelectedIndex = 0
        $listBox.Add_DoubleClick({ $form.DialogResult = 'OK' })

        $btnSelect = New-Object System.Windows.Forms.Button -Property @{
            Text         = 'Select'
            DialogResult = 'OK'
            Location     = '214,365'
        }
        $btnCancel = New-Object System.Windows.Forms.Button -Property @{
            Text         = 'Cancel'
            DialogResult = 'Cancel'
            Location     = '295,365'
        }
        $form.Controls.AddRange(@($lblInstructions, $listBox, $btnSelect, $btnCancel))
        $form.AcceptButton = $btnSelect
        $form.CancelButton = $btnCancel

        $result = $form.ShowDialog()
        $selectedIndex = $listBox.SelectedIndex
        $form.Dispose()

        if ($result -eq 'OK' -and $selectedIndex -ge 0) {
            $items[$selectedIndex]
        }
    }
}

Function Select-Team {
    Param (
        [string]
        $Instructions = "Select a Team"
    )
    $script:currentTeams | Select-FromList -Instructions $Instructions
}

Function Select-TeamChannel {
    Param (
        [Parameter(Mandatory)]
        [string]
        $GroupId,
        [string]
        $Instructions = "Select a Channel"
    )
    Get-TeamChannel -GroupId $GroupId | Select-FromList -Instructions $Instructions
}

#lets the user pick an existing team or create a new one. Returns the team, or nothing if cancelled.
Function Get-TargetTeam {
    if (Read-YesNo "Do you wish to use an existing team?") {
        return Select-Team -Instructions "Select a destination Team"
    }

    $teamName = Read-InputBox -Prompt 'Enter New Team Name:' -Title 'Team Name'
    if (-not $teamName) { return }
    $teamDescription = Read-InputBox -Prompt 'Enter New Team Description (be careful. This might be hard to change later):' -Title 'Team Description'

    $newTeam = @{
        DisplayName  = $teamName
        MailNickName = $teamName -replace '[^A-Za-z0-9]', ''
        Visibility   = $TEAMVISIBILITY
    }
    if ($teamDescription) { $newTeam.Description = $teamDescription }

    $team = New-Team @newTeam -ErrorAction Stop
    Write-Host "Created team $teamName" -ForegroundColor Green
    Update-CurrentTeams
    $team
}

#lets the user pick an existing channel or create a new one. Returns the channel, or nothing if cancelled.
Function Get-TargetChannel {
    Param (
        [Parameter(Mandatory)]
        [string]
        $GroupId
    )
    if (Read-YesNo "Do you wish to use an existing team Channel?") {
        return Select-TeamChannel -GroupId $GroupId -Instructions "Select a destination Channel"
    }

    $channelName = Read-InputBox -Prompt 'Enter New Channel Name:' -Title 'Channel Name'
    if (-not $channelName) { return }
    $channelDescription = Read-InputBox -Prompt 'Enter New Channel Description (be careful. This might be hard to change later):' -Title 'Channel Description'

    $newChannel = @{
        GroupId        = $GroupId
        DisplayName    = $channelName
        MembershipType = $CHANNELMEMBERSHIPTYPE
    }
    if ($channelDescription) { $newChannel.Description = $channelDescription }

    New-TeamChannel @newChannel -ErrorAction Stop
}

#returns a case-insensitive set of the UPNs of everyone currently in a team.
Function Get-TeamMemberSet {
    Param (
        [Parameter(Mandatory)]
        [string]
        $GroupId
    )
    $members = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    Get-TeamUser -GroupId $GroupId | ForEach-Object { [void]$members.Add($_.User) }
    #comma stops PowerShell from unrolling the set into individual strings
    , $members
}

#channel members must belong to the parent team, so add the user to the team first if they're missing.
#$TeamMembers comes from Get-TeamMemberSet and is updated as users are added.
Function Add-TeamUserIfMissing {
    Param (
        [Parameter(Mandatory)]
        [string]
        $GroupId,
        [Parameter(Mandatory)]
        [string]
        $User,
        [Parameter(Mandatory)]
        [System.Collections.Generic.HashSet[string]]
        $TeamMembers
    )
    if ($TeamMembers.Contains($User)) { return }

    Write-Host "$User is not in the team yet. Adding to team..." -ForegroundColor DarkYellow
    Add-TeamUser -GroupId $GroupId -User $User -ErrorAction Stop
    [void]$TeamMembers.Add($User)
}

#endregion




# Action Functions are for doing the things
#region ActionFunctions


################################################################
#Code for importing CSV to a Team or Team Channel              #
################################################################
#takes a teams GroupID (and optionally a channel) then prompts for CSV files to import.
Function Import-CsvToTeam {
    Param (
        [Parameter(Mandatory)]
        [String]
        $TeamGroupId,
        [String]
        $ChannelDisplayName,
        [bool]
        $WithConfirm = $false
    )
    $target = if ($ChannelDisplayName) { "team channel $ChannelDisplayName" } else { "team" }

    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'CSV (*.CSV)|*.CSV'
        MultiSelect      = $true
    }
    if ($fileBrowser.ShowDialog() -ne 'OK') {
        Write-Host "No CSV selected."
        Write-OperationAborted
        return
    }

    if ($ChannelDisplayName) {
        $teamMembers = Get-TeamMemberSet -GroupId $TeamGroupId
    }

    foreach ($curPath in $fileBrowser.FileNames) {
        $students = @(Import-Csv -Path $curPath)
        if ($students.Count -gt 0 -and -not ($students[0].PSObject.Properties.Name -contains 'Email')) {
            Write-Host "Skipping $curPath - it has no Email column" -ForegroundColor Red
            continue
        }

        foreach ($student in $students) {
            $email = "$($student.Email)".Trim()
            if (-not $email) {
                Write-Host "skipping empty email address"
                continue
            }
            if ($WithConfirm -and -not (Read-YesNo "Add $email to selected $target?")) {
                Write-Host "Skipping user $email"
                continue
            }

            Write-Host "Adding user $email to $target" -ForegroundColor Cyan
            try {
                if ($ChannelDisplayName) {
                    Add-TeamUserIfMissing -GroupId $TeamGroupId -User $email -TeamMembers $teamMembers
                    Add-TeamChannelUser -GroupId $TeamGroupId -DisplayName $ChannelDisplayName -User $email -ErrorAction Stop
                }
                else {
                    Add-TeamUser -GroupId $TeamGroupId -User $email -ErrorAction Stop
                }
                Write-Host "Successfully added user $email to $target" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to add user ${email}: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

Function Add-UsersToTeam {
    Param (
        [bool]
        $WithConfirm = $false
    )
    $team = Get-TargetTeam
    if (-not $team) { return Write-OperationAborted }

    Import-CsvToTeam -TeamGroupId $team.GroupId -WithConfirm $WithConfirm
}

Function Add-UsersToTeamChannel {
    Param (
        [bool]
        $WithConfirm = $false
    )
    $team = Get-TargetTeam
    if (-not $team) { return Write-OperationAborted }

    $channel = Get-TargetChannel -GroupId $team.GroupId
    if (-not $channel) { return Write-OperationAborted }

    Import-CsvToTeam -TeamGroupId $team.GroupId -ChannelDisplayName $channel.DisplayName -WithConfirm $WithConfirm
}


################################################################
#Code for Copying members between teams CHANNELS               #
################################################################
Function Copy-TeamsChannelMembers {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]
        $FromTeamId,
        [Parameter(Mandatory)]
        [String]
        $FromChannelName,
        [Parameter(Mandatory)]
        [String]
        $ToTeamId,
        [Parameter(Mandatory)]
        [String]
        $ToChannelName
    )
    $users = @(Get-TeamChannelUser -GroupId $FromTeamId -DisplayName $FromChannelName)
    Write-Host "Please Confirm that you want to add $($users.Count) users to $ToChannelName from $FromChannelName"
    $users | Format-Table | Out-Host

    if (-not (Read-YesNo "Do you wish to execute this?")) { return Write-OperationAborted }

    $teamMembers = Get-TeamMemberSet -GroupId $ToTeamId
    foreach ($user in $users) {
        try {
            Add-TeamUserIfMissing -GroupId $ToTeamId -User $user.User -TeamMembers $teamMembers
            Add-TeamChannelUser -GroupId $ToTeamId -DisplayName $ToChannelName -User $user.User -ErrorAction Stop
            Write-Host "Added $($user.User)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to add $($user.User): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Host "Channel copy Complete!`n" -ForegroundColor Cyan
}

Function Invoke-ChannelCopy {
    $sourceTeam = Select-Team -Instructions "Select a Source Team"
    if (-not $sourceTeam) { return Write-OperationAborted }
    $sourceChannel = Select-TeamChannel -GroupId $sourceTeam.GroupId -Instructions "Select a Source Channel"
    if (-not $sourceChannel) { return Write-OperationAborted }

    $destTeam = Select-Team -Instructions "Select a Destination Team"
    if (-not $destTeam) { return Write-OperationAborted }
    $destChannel = Select-TeamChannel -GroupId $destTeam.GroupId -Instructions "Select a Destination Channel"
    if (-not $destChannel) { return Write-OperationAborted }

    Copy-TeamsChannelMembers -FromTeamId $sourceTeam.GroupId -FromChannelName $sourceChannel.DisplayName -ToTeamId $destTeam.GroupId -ToChannelName $destChannel.DisplayName
}



################################################################
#Code for copy members between TEAMS                           #
################################################################
Function Copy-TeamMembers {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]
        $FromTeamId,
        [Parameter(Mandatory)]
        [String]
        $ToTeamId
    )
    $users = @(Get-TeamUser -GroupId $FromTeamId)
    Write-Host "Please Confirm that you want to Copy $($users.Count) users between teams"
    $users | Format-Table | Out-Host

    if (-not (Read-YesNo "Do you wish to execute this?")) { return Write-OperationAborted }

    Write-Host "Copying users..."
    foreach ($user in $users) {
        $role = if ($user.Role -eq 'owner') { 'Owner' } else { 'Member' }
        try {
            Add-TeamUser -GroupId $ToTeamId -User $user.User -Role $role -ErrorAction Stop
            Write-Host "Added $($user.User) as $role" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to add $($user.User): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Host "Team copy complete!`n" -ForegroundColor Cyan
}

Function Invoke-TeamCopy {
    $sourceTeam = Select-Team -Instructions "Select a Source Team"
    if (-not $sourceTeam) { return Write-OperationAborted }

    $destTeam = Select-Team -Instructions "Select a Destination Team"
    if (-not $destTeam) { return Write-OperationAborted }

    Copy-TeamMembers -FromTeamId $sourceTeam.GroupId -ToTeamId $destTeam.GroupId
}


################################################################
#Code for exporting team members to CSV                        #
################################################################
#the Email column matches what Import-CsvToTeam expects, so exports can be re-imported.
Function Export-TeamMembers {
    $team = Select-Team -Instructions "Select a Source Team"
    if (-not $team) { return Write-OperationAborted }

    $fileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'CSV (*.CSV)|*.CSV'
        FileName         = ($team.DisplayName -replace '[\\/:*?"<>|]', '_') + '.csv'
    }
    if ($fileBrowser.ShowDialog() -ne 'OK') { return Write-OperationAborted }

    $users = @(Get-TeamUser -GroupId $team.GroupId)
    $users |
        Select-Object @{ N = 'Email'; E = { $_.User } }, Name, Role |
        Export-Csv -Path $fileBrowser.FileName -NoTypeInformation
    Write-Host "Exported $($users.Count) members to $($fileBrowser.FileName)`n" -ForegroundColor Green
}


################################################################
#Code for removing all members from a team                     #
################################################################
#owners are kept so the team is never left without an owner.
Function Remove-AllTeamMembers {
    $team = Select-Team -Instructions "Select a Team to remove members from"
    if (-not $team) { return Write-OperationAborted }

    $users = @(Get-TeamUser -GroupId $team.GroupId | Where-Object { $_.Role -ne 'owner' })
    if ($users.Count -eq 0) {
        Write-Host "$($team.DisplayName) has no non-owner members to remove.`n" -ForegroundColor Cyan
        return
    }

    Write-Host "Please Confirm that you want to Remove $($users.Count) members from $($team.DisplayName) (owners will be kept)" -ForegroundColor Yellow
    $users | Format-Table | Out-Host

    if (-not (Read-YesNo "Do you wish to execute this?")) { return Write-OperationAborted }

    Write-Host "Removing users..."
    foreach ($user in $users) {
        try {
            Remove-TeamUser -GroupId $team.GroupId -User $user.User -ErrorAction Stop
            Write-Host "Removed $($user.User)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to remove $($user.User): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Host "Team member removal complete!`n" -ForegroundColor Cyan
}


#endregion

Clear-Host

Install-TeamsPowerShell

try {
    Connect-Teams
    Show-Menu
}
finally {
    Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
}





<#

Connect-MsolService

$azureSession = Connect-AzureAd

$azureSession | Get-Member
$teamsSesssion = Connect-MicrosoftTeams

get-team -user '100269208@durhamcollege.ca'

New-AzureADMSInvitation

Disconnect-AzureAD
Disconnect-MicrosoftTeams

$email = "tom@turnertechnology.ca"

Get-AzureADUser -ObjectId "$($email -replace "@", "_")#EXT#@dconline.onmicrosoft.com"

Get-AzureADUser -Filter "UserType eq 'Guest'"



Function Add-GuestToTeams {
    Param(
    [parameter()]
    [String]
    $email,
    [parameter()]
    [String]
    $tenant
    )
    BEGIN {}
    PROCESS{
        try
        {
           $guestUser = Get-AzureADUser -ObjectId "$($email -replace "@", "_")#EXT#@$tenant" -ErrorAction Ignore
           write-host "user $email was found"
        }
        catch{
           $guestUser = $null
           write-host "user $email was NOT found"
        }
    }
    END {}
}

Function ProcessListOfGuests {
    Param(
    [parameter(ValueFromPipeline,Position=0)]
    [object[]]
    $guest
    )
    BEGIN {
        Connect-AzureAD
        $tenant = Get-AzureADTenantDetail
        $tenantUrl = ($tenant.VerifiedDomains | where-object {$_._Default -eq $true}).Name
    }
    PROCESS{
        Add-GuestToTeams -email $guest.email -tenant $tenantUrl
    }
    END{}
}


$students = Import-Csv -Path "C:\Users\tturner\OneDrive - Turner Technology\Documents\turnerthomas\misc\testimport.csv"

$students | ProcessListOfGuests


#>
