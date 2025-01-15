#Tom Turner in WEbD Class


#region Dependencies
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

Add-Type -AssemblyName System.Windows.Forms
#endregion

#region options
$TEAMVISIBILITY = "private"
$DOMAINNAME = "@durhamcollege.ca"
REMOVE-Variable -Name "currentTeams"
New-Variable -Name "currentTeams" -Scope Script  -Value $null
#endregion


#Startup functions are used at the beginning to setup current session.
#region StartupFunctions

#functions for setting up the powershell module.
Function Startup-Stuff {
    Clear-Host
}

#function for installing necessary POwershell modules.
Function Setup-TeamsPowerShell {
    Register-PackageSource -provider NuGet -name nugetRepository -location https://www.nuget.org/api/v2 -ErrorAction SilentlyContinue
    Get-PackageProvider | where name -eq 'nuget' | Install-PackageProvider -Force

    #install Powershell get
    if (-not (Get-Module -Name "PowerShellGet")) {
        # module is not loaded
        Install-Module -Name PowerShellGet -Force -AllowClobber
    }else {write-host "Powershell Get already installed" -ForegroundColor Green}
    

    #install teams powershell
    if (-not (Get-Module -Name "MicrosoftTeams")) {
        # module is not loaded
        Write-host "Microsoft Teams powershell module not installed. installing..." -foregroundcolor DarkYellow
        Install-Module -Name MicrosoftTeams -Force -AllowClobber
    }else {write-host "Powershell Microsoft Teams already installed" -ForegroundColor Green}

    
}

#function to connect to to teams, prompting for credentials. 
Function Initiate-Teams {
    if($currentUser -eq $null)
    {
        write-host "No teams connection available. Initating Connection..." -foregroundcolor DarkYellow
        #connect to teams
        $currentUser = Connect-MicrosoftTeams


    }
    
    $currentLoginId = $currentUser.Account.Id
    write-host "Proceeding with user " $currentLoginId

            
    #get current user teams
    set-variable -name currentTeams -value (get-team -user $currentLoginId) -Scope Script
    Write-Host "User has $($currentTeams.count) teams"

}

#Primary function for menu controlling
Function Show-Menu
{
       #clear-host
        $answer=""
        do
        {
            Write-Host "###################################################"
            write-Host "## Tools for Importing CSV to Microsoft Teams    ##"
            write-Host "##                                               ##"
            write-Host "##                                               ##"
            write-Host "##                                               ##"
            write-Host "##                                               ##"
            Write-Host "###################################################"

            write-host "What would you like to do now?" -ForegroundColor Green
            write-host "1: Import CSV to teams"
            write-host "2: Import CSV to Channel"
            write-host "3: Copy Members between Teams"
            write-host "4: Copy Members between Channels"
            Write-host "5: Export Team Members to CSV"
            Write-host "6: Remove All Team Members"
            Write-host "[q]  Quit " -ForegroundColor Yellow
            write-host "Select 1-6 or q: " -nonewline -ForegroundColor Green
            $answer = Read-Host  
    
            #if you have selected a valid 0 through number of available options you are good
            if($answer -In 1..6)
            {
                switch($answer)
                {
                    1{
                        write-host "Do you wish to confirm every addition? Y or N:" -ForegroundColor Green -NoNewline
                        $asked =  Read-Host  | check-YN
                        if($asked -eq $false)
                        {
                            WRite-Host "You will NOT be asked before every student"
                            Add-UsersToTeams -withConfirm $asked
                        }
                        else
                        {
                            Write-Host "You WILL be asked before every student" -foregroundcolor DarkCyan
                            Add-UsersToTeams -withConfirm $asked
                        }
                    }
                    2{
                        write-host "Do you wish to confirm every addition? Y or N:" -ForegroundColor Green -NoNewline
                        $asked =  Read-Host  | check-YN
                        if($asked -eq $false)
                        {
                            WRite-Host "You will NOT be asked before every student"
                            Add-UsersToTeamsChannel -withConfirm $asked
                        }
                        else
                        {
                            Write-Host "You WILL be asked before every student" -foregroundcolor DarkCyan
                            Add-UsersToTeamsChannel -withConfirm $asked
                        }                  
                    }
                    3{
                        Execute-TeamCopy      
                    }
                    4{
                        Execute-ChannelCopy
                    }
                    5{
                        Export-TeamMembers
                    }
                    6{
                        #remove all team members
                    }
                }
        
            }
            elseif($answer -ne "q" -and $answer -notin 1..3)
            {
                #you done goofed.
                Write-Host "Invalid Selection!" -BackgroundColor black -ForegroundColor Red
            }

        } Until ($answer -eq "q")
        
}


#endregion




#utility functions are resuable, multipurpose functions.
#region UtilityFunctions
#function for converting Y or N or Yes and NO to boolean
Function check-YN {
    Param (
        [parameter(ValueFromPipeline,Position=0)]
        [String[]]
        $value
    )
    $result =  @("true","false","yes","no", "y", "n") -contains $value -and @("true","yes", "y") -contains $value
    Write-Output $result
}



#functon for asking a yes no question
Function ask-YesNo{
Param (
        [parameter(ValueFromPipeline,Position=0)]
        [string]
        $value)
        
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Yes'
        $no = New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'No'
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

        $result = $host.ui.PromptForChoice('Yes or No', $value, $options, 0)
        write-output $(if($result -eq 0) {$true}else{$false})
}


# a function that is passed objects and attempts to display them as radio buttons.
Function Ask-Options{
Param (
        [parameter(ValueFromPipeline=$true,Position=0)]
        [object[]]$value,
        [parameter()]
        [string]$instructions = ""
      )
      BEGIN{
            Add-Type -AssemblyName System.Windows.Forms
            $form1 = New-Object System.Windows.Forms.Form
            $form1.StartPosition = 'CenterScreen'
            $flp = New-Object System.Windows.Forms.FlowLayoutPanel
            
            $lblInstructions = New-Object System.Windows.Forms.Label
            $lblInstructions.Text = "$instructions"
            $lblInstructions.location = "10,$($form1.Height - 100)"
            $lblInstructions.AutoSize=$true
            $lblInstructions.Anchor = "Bottom,Left"

            $form1.Controls.Add($lblInstructions)
            $form1.Controls.Add($flp)
            $form1.AutoSize = $true
            
            $flp.Name = 'MyFlowPanel'
            $w = $form1.Width.ToString()
            $h = 10
            $flp.Size = "$w,$h"
            $flp.FlowDirection = 'TopDown'
            $flp.AutoSize = $true
            $flp.BackColor = "Red"
            $btn = New-Object System.Windows.Forms.Button
            $form1.Controls.Add($btn)
            $btn.Text = 'Select'
            $btn.DialogResult = 'OK'
            $btn.location="10,$($form1.Height - 50)"
            $btn.Anchor = "Bottom,Left"
            $btn.BringToFront()
        }
        Process {
                $rb = New-Object System.Windows.Forms.RadioButton
	                 $flp.Controls.Add($rb)
	                 $rb.Text = $_.DisplayName
	                 $rb.AutoSize = $true
        }
        END{


            $form1.ShowDialog()
            $form1.Controls['MyFlowPanel'].Controls | Where-Object{ $_.Checked } | Select-Object @{N='TeamName';E={$_.Text}} | Write-Output
        }
}


#endregion




# ACtion Functions are for doing the things
#region ActionFunctions


################################################################
#Code for importing CSV to a Team                              #
################################################################
#taking a teams GroupID then prompting to be fed CSV file paths
Function ImportCSV-ToTeams {
    Param (
        [parameter()]
        [String]
        $TeamGroupId,
        [parameter()]
        [bool]
        $withConfirm=$false
    )
    BEGIN{}
    PROCESS{
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Filter = 'CSV (*.CSV)|*.CSV'
            MultiSelect = $true
            }

        $null = $FileBrowser.ShowDialog()
        foreach($curpath in $FileBrowser.FileNames)
        {
            $students = Import-Csv -Path "$curpath"
            $students | foreach-object -Process { 
                if(-not $($_.Email) -eq "")
                {
                    if($withConfirm)
                    {
                        write-host "Add $($_.Email) to selected team? Y or N:" -ForegroundColor Green -NoNewline
                        $confirm =  Read-Host  | check-YN
                        if($confirm)
                        {
                            Add-TeamUser -GroupId "$TeamGroupId" -User "$($_.Email)" 
                            Write-host "User $($_.Email) added to team" -ForegroundColor Green
                        }
                        else
                        {
                            write-host "Skipping user $($_.Email)"
                        }
                    }
                    else
                    {
                        Write-host "Adding user $($_.Email) to teams" -ForegroundColor Cyan
                        Add-TeamUser -GroupId "$TeamGroupId" -User "$($_.Email)" 
                        Write-host "Successfully added user $($_.Email) to team" -ForegroundColor Green
                    }
                }
                else
                { write-host "skipping empty email address" }
                
            }
        }

    }
    END {}
}

Function Add-UsersToTeams {
    Param (
        [parameter()]
        [bool]
        $withConfirm=$false
    )

    if($("Do you wish to use an existing team?" | ask-YesNo | check-YN))
    {
        $answer = $currentTeams | Ask-Options 
        $selectedTeam = $currentTeams | Where-Object { $_.DisplayName -eq $answer.TeamName}
    }
    else
    {
        $teamName = [Microsoft.VisualBasic.Interaction]::InputBox('Team Name', 'Enter New Team Name:')
        $teamDescription = [Microsoft.VisualBasic.Interaction]::InputBox('Team Description', 'Enter New Team Description:')
        $selectedTeam = New-Team -MailNickName "$($teamName.Replace("" "", """").Replace(""/"", """"))" -DisplayName "$teamName" -Visibility "$TEAMVISIBILITY" -Description "$teamDescription"
    }
    ImportCSV-ToTeams -TeamGroupId $selectedTeam.GroupId -withConfirm $withConfirm

}

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
        $guestUser = Get-AzureADUser -ObjectId "$($email -replace "@", "_")#EXT#@$tenant.onmicrosoft.com"
    }
    END {}
}


################################################################
#Code for Copying memebers between teams CHANNELS              #
################################################################
Function Copy-TeamsChannelMembers {
[CmdletBinding()]
    Param (
        [parameter()]
        [String]
        $FromTeamId,
        [parameter()]
        [String]
        $FromChannelName,
        [parameter()]
        [String]
        $ToTeamId,
        [parameter()]
        [String]
        $ToChannelName
    )
    BEGIN {}
    PROCESS {
        $users = Get-TeamChannelUser -GroupId "$FromTeamId" -DisplayName "$FromChannelName"
        Write-Host "Please Confirm that you want to add $($users.count) users to $ToChannelName from $FromChannelName"
        $users | format-table 
        
        write-host "Do you wish to execute this? Y or N:" -ForegroundColor Green -NoNewline
        $answer =  Read-Host  | check-YN
        if($answer)
        {
            $users | Add-TeamChannelUser -GroupId "$ToTeamId" -DisplayName "$ToChannelName"
            Write-host "Channel copy Complete!" -ForegroundColor Cyan
            Write-Host ""
            Write-Host ""
            Write-Host ""
            Write-Host ""
        }
        else
        {
            Write-host "Operation Aborted!" -ForegroundColor Cyan
            Write-Host ""
            Write-Host ""
            Write-Host ""
            Write-Host ""

        }

        
    }
    END {}

}

Function Execute-ChannelCopy {

        $sourceteamanswer = $currentTeams | Ask-Options -instructions "Select a Source Team"
        $selectedSourceTeam = $currentTeams | Where-Object { $_.DisplayName -eq $sourceteamanswer.TeamName}
        $sourcechannels = Get-TeamChannel -GroupId $selectedSourceTeam.GroupId
        $sourcechannelanswer = $sourcechannels | Ask-Options -instructions "Select a Source Channel"
        $selectedSourceChannel = $sourcechannels | Where-Object { $_.DisplayName -eq $sourcechannelanswer.TeamName}


        $destTeamAnswer = $currentTeams | Ask-Options -instructions "Select a Destination Team"
        $selectedDestTeam = $currentTeams | Where-Object { $_.DisplayName -eq $destTeamAnswer.TeamName}
        $destChannels = Get-TeamChannel -GroupId $selectedDestTeam.GroupId
        $DestinationChannelAnswer = $destChannels | Ask-Options -instructions "Select a Destination Channel"
        $selectedDestChannel = $destChannels | Where-Object { $_.DisplayName -eq $DestinationChannelAnswer.TeamName}


        Copy-TeamsChannelMembers -FromTeamId $selectedSourceTeam.GroupId -FromChannelName $selectedSourceChannel.DisplayName -ToTeamId $selectedDestTeam.GroupId -ToChannelName $selectedDestChannel.DisplayName
}



################################################################
#Code for copy memebers between TEAMS                          #
################################################################
Function Copy-TeamMembers {
[CmdletBinding()]
    Param (
        [parameter()]
        [String]
        $FromTeamId,
        [parameter()]
        [String]
        $ToTeamId
    )
    BEGIN {}
    PROCESS {
        $users = Get-TeamUser -GroupId "$FromTeamId"
        Write-Host "Please Confirm that you want to Copy $($users.count) users between channels"
        $users | format-table 
        
        write-host "Do you wish to execute this? Y or N:" -ForegroundColor Green -NoNewline
        $answer =  Read-Host  | check-YN
        if($answer)
        {
            Write-host "Copying users..."
            $users | Add-TeamUser -GroupId "$ToTeamId" 
            Write-host "Team copy complete!" -ForegrounndColor Cyan
            write-host ""
            write-host ""
            write-host ""
            write-host ""
        }
        else
        {
            Write-host "Operation Aborted!" -ForegroundColor Cyan
            write-host ""
            write-host ""
            write-host ""
            write-host ""

        }

        
    }
    END {}

}

Function Execute-TeamCopy {

        $sourceteamanswer = $currentTeams | Ask-Options -instructions "Select a Source Team"
        $selectedSourceTeam = $currentTeams | Where-Object { $_.DisplayName -eq $sourceteamanswer.TeamName}
        
        $destTeamAnswer = $currentTeams | Ask-Options -instructions "Select a Destination Team"
        $selectedDestTeam = $currentTeams | Where-Object { $_.DisplayName -eq $destTeamAnswer.TeamName}
        
        Copy-TeamMembers -FromTeamId $selectedSourceTeam.GroupId -ToTeamId $selectedDestTeam.GroupId 
}


Function Export-TeamMembers{
        $sourceteamanswer = $currentTeams | Ask-Options -instructions "Select a Source Team"
        $selectedSourceTeam = $currentTeams | Where-Object { $_.DisplayName -eq $sourceteamanswer.TeamName}
       
        $users = Get-TeamUser -GroupId $selectedSourceTeam.GroupId

         $FileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Filter = 'CSV (*.CSV)|*.CSV'
            }

        $null = $FileBrowser.ShowDialog()

        $users | ConvertTo-csv | out-file -FilePath $FileBrowser.FileNames[0]

}


################################################################
#Code for importing CSV to a Teams Channel                     #
################################################################
Function ImportCSV-ToTeamsChannel {
    Param (
        [parameter()]
        [String]
        $TeamGroupId,
        [parameter()]
        [String]
        $ChannelDisplayName,
        [parameter()]
        [bool]
        $withConfirm=$false
    )
    BEGIN{}
    PROCESS{
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Filter = 'CSV (*.CSV)|*.CSV'
            MultiSelect = $true
            }

        $null = $FileBrowser.ShowDialog()
        foreach($curpath in $FileBrowser.FileNames)
        {
            $students = Import-Csv -Path "$curpath"
            $students | foreach-object -Process { 
                if(-not $($_.Email) -eq "")
                {
                    if($withConfirm)
                    {
                        write-host "Add $($_.Email) to selected team channel $ChannelDisplayName ? Y or N:" -ForegroundColor Green -NoNewline
                        $confirm =  Read-Host  | check-YN
                        if($confirm)
                        {
                            Add-TeamChannelUser -GroupId "$TeamGroupId" -DisplayName "$ChannelDisplayName"  -User "$($_.Email)" 
                            Write-host "User $($_.Email) added to team channel $ChannelDisplayName " -ForegroundColor Green
                        }
                        else
                        {
                            write-host "Skipping user $($_.Email)"
                        }
                    }
                    else
                    {
                        Write-host "Adding user $($_.Email) to team channel $TeamChannelDisplayName" -ForegroundColor Cyan
                        Add-TeamChannelUser -GroupId "$TeamGroupId" -DisplayName "$ChannelDisplayName"  -User "$($_.Email)" 
                        Write-host "Successfully added user $($_.Email) to team channel $ChannelDisplayName " -ForegroundColor Green
                    }
                }
                else
                { write-host "skipping empty email address" }
                
            }
        }

    }
    END {}
}

Remove-AllTeamMembers {
    Param (
        [parameter()]
        [String]
        $TeamGroupId,
        [parameter()]
        [bool]
        $withConfirm=$false
    )
    BEGIN{}
    PROCESS{
   #     $users = Get-TeamUser -GroupId "$TeamGroupId" | Where-Object { $_. -eq $answer.TeamName}
        Write-Host "Please Confirm that you want to Remove $($users.count) From TEAM"
        $users | format-table 
        
        write-host "Do you wish to execute this? Y or N:" -ForegroundColor Green -NoNewline
        $answer =  Read-Host  | check-YN
        if($answer)
        {
            Write-host "Copying users..."
            $users | Remove-TeamUser -GroupId "$ToTeamId" 
            Write-host "Team copy complete!" -ForegrounndColor Cyan
            write-host ""
            write-host ""
            write-host ""
            write-host ""
        }
        else
        {
            Write-host "Operation Aborted!" -ForegroundColor Cyan
            write-host ""
            write-host ""
            write-host ""
            write-host ""

        }

            
            
            
            Add-TeamUser -GroupId $selectedTeam.GroupId -User "$($_.Email)" 
            
    
        $answer = $currentTeams | Ask-Options 
        $selectedTeam = $currentTeams | Where-Object { $_.DisplayName -eq $answer.TeamName}
        Remove-TeamUser -GroupId "$TeamGroupId" -User "$($_.Email)" 
    }
    END{}
}
Function Add-UsersToTeamsChannel {
    Param (
        [parameter()]
        [bool]
        $withConfirm=$false
    )

    if($("Do you wish to use an existing team?" | ask-YesNo | check-YN))
    {
        $answer = $currentTeams | Ask-Options 
        $selectedTeam = $currentTeams | Where-Object { $_.DisplayName -eq $answer.TeamName}
    }
    else
    {
        $teamName = [Microsoft.VisualBasic.Interaction]::InputBox('Team Name', 'Enter New Team Name:')
        $teamDescription = [Microsoft.VisualBasic.Interaction]::InputBox('Team Description', 'Enter New Team Description (be careful. This might be hard to change later):')
        $selectedTeam = New-Team -MailNickName "$($teamName.Replace("" "", """").Replace(""/"", """"))" -DisplayName "$teamName" -Visibility "$TEAMVISIBILITY" -Description "$teamDescription"
    }

    if($("Do you wish to use an existing team Channel?" | ask-YesNo | check-YN))
    {
        $teamChannels = Get-TeamChannel -GroupId $selectedTeam.GroupId
        $teamChannleAnswer = $teamChannels | Ask-Options -instructions "Select a destination Channel"
        $selectedChannel = $teamChannels | Where-Object { $_.DisplayName -eq $teamChannleAnswer.TeamName}
    }
    else
    {
        $ChannelName = [Microsoft.VisualBasic.Interaction]::InputBox('Channel Name', 'Enter New Channel Name:')
        $ChannelDescription = [Microsoft.VisualBasic.Interaction]::InputBox('Channel Description', 'Enter New Channel Description (be careful. This might be hard to change later):')
        $selectedChannel = New-TeamChannel -GroupId $selectedTeam.GroupId -DisplayName "$ChannelName" -MembershipType "$TEAMVISIBILITY" -Description "$ChannelDescription"
    }



    ImportCSV-ToTeamsChannel -TeamGroupId $selectedTeam.GroupId -withConfirm $withConfirm -ChannelDisplayName $selectedChannel.DisplayName

}





#endregion

Startup-Stuff

Setup-TeamsPowerShell

Initiate-Teams

Show-Menu

Disconnect-MicrosoftTeams





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