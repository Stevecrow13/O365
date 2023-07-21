<#
.Synopsis
    Get's the date of last Windows Update install time and the last succesful search time.
.Description
    Get's the date of last Windows Update install time and the last succesful search time.  Creats a new-pssession, gets the data, then disconnects and closes the pssession.  If no parameters are used, it will get the local computer information.  Created by Steve Crowley.  Reachable at stevecrow13@gmail.com.
.Parameter computerName
    The remove server you want to get Windows Update information from.
.Example
    .\get-remotewinupdateinfo.ps1 -computerName DomainController.contoso.com
.Notes
    Created by Steve Crowley.  Reachable at stevecrow13@gmail.com.
#>

#Parameters
param
(
    [string]$computerName
)

#function to be put into a script block to run the script remotely.
function getlastwindowsupdate
{
    $windowsUpdateObject = New-Object -ComObject Microsoft.Update.AutoUpdate
    $windowsUpdateObject.Results
    return $updateoutput
}

#function to run the script locally.
function getlastwindowsupdatelocal
{
    $windowsUpdateObject = New-Object -ComObject Microsoft.Update.AutoUpdate
    $updateoutput = $windowsUpdateObject.Results
    return $updateoutput
}

#If $computerName is empty or null, run the command function locally
if([string]::IsNullOrEmpty($computerName))
    {
        $lastWindowsUpdateData = getlastwindowsupdatelocal
        $computerName = hostname.exe
        $lastWindowsUpdateData | add-member -NotePropertyName computerName -NotePropertyValue $computerName
        return $lastWindowsUpdateData
        break
    }

#Creates new PSSession with out creating a windows profile and using the remote computerName.
$PssessionOptions = new-pssessionoption -NoMachineProfile
$RemoteSession = New-PSSession -SessionOption $PssessionOptions -ComputerName $computerName -ErrorAction SilentlyContinue

#Runs the function as a script block on the remote PSSession and saves it to $results
$results = Invoke-Command -Session $RemoteSession -ScriptBlock ${function:getlastwindowsupdate} -ErrorAction SilentlyContinue
$printableresults = $results | Select-Object LastSearchSuccessDate, LastInstallationSuccessDate
$printableresults | add-member -NotePropertyName computerName -NotePropertyValue $computerName
return $printableresults

#Gets the remote PSSession by computerName, disconnects it cleanly, then removes it.
Get-PSSession -ComputerName "$computerName" | Disconnect-PSSession | Remove-PSSession
