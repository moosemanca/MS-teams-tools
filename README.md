# MS-teams-tools

A Powershell script for performing various CRUD and manipulation actions with Microsoft teams

Things it can do:
- Import a list of users via CSV. (CSV must have at least one field named Email)
- Copy Members from One Team to another
- Copy Members from one Team/Channel to another Team/Channel

Considering other features as needed:
- importing list of teams to create
- Importing list of users to Remove from a Team/sub channels
- Exporting Teams/Channels and their members.


Can easily be run with the following single line of code (REQUIRES ELEVATED PRIVILEGES):

`Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://github.com/moosemanca/MS-teams-tools/blob/main/importingTeamMembers.ps1?raw=true'))`
