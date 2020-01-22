
##*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*##
## Script Quick Guide                                                                                                    ##
## -- Open this script within "Windows Powershell ISE(x86)" to execute.                                                  ##
## -- This script is to be run on a local machine only. DO NOT RUN ON SERVERS.                                           ##
## -- Drop this script in an empty folder with exported data from both VSphere and Rubrik. (must be .csv)                ##
## -- VSphere report should only need generated ~every 2 weeks or when a new client backup configuration is complete     ##
## -- Rubrik data can be taken from the "Daily Detail Report" sent to the service desk mailbox each day.                 ##
## -- Name each data set "VSphereData.csv" and "RubrikData.csv" accordingly.                                             ##
## -- Highlight whole script and press F8 to run. Console will detail any errors.                                        ##
##*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*##

##Get Current Script Directory. Works in both ISE and executed script. 
function Get-ScriptDirectory {if ($psise) {Split-Path $psise.CurrentFile.FullPath}
    else {$global:PSScriptRoot}}

##Set scripts home location found in above function. 
$scriptHome = Get-ScriptDirectory

##Generate a path to export the report next to the script.
$exportPath = $scriptHome + "\ComparisonResult.csv"
$confirmedResult = $scriptHome + "\ConfirmedBackups.csv"
$failedResult = $scriptHome + "\NoBackups.csv"

##Importing Data Sets - Only needing the column listing Virtual Machines
$vSphereData = Import-CSV -Path $scriptHome"\VSphereData.csv" | Where-Object {($_.'State' -eq "Powered On") } | Select-Object 'Name'

##Importing RubrikData from "Daily Detail Report" received by ServiceDesk. Where we need only the objects with attributes "VSphere VM" and "Backup".
$rubrikData = Import-CSV -Path $scriptHome"\RubrikData.csv" | Where-Object {($_.'Object Type' -eq "vSphere VM") } | Select-Object 'Object Name', 'SLA Domain'  ##-and ($_.'Task Type' -like "Backup")

##Creating arrays for each list of virtual machines from VSphere and Rubrik data sets. 
$rVM = $rubrikData.'Object Name'
$vsVM = $vSphereData.Name

##We compare vSphere objects to Rubrik list and make sure all vSphere objects are present in Rubrik. 
##For each VM listed in VSphere data, make sure it is included in Rubrik data. If not, add to exported CSV. 
foreach ($VirtualMachine in $vsVM){
    if ($rVM -notcontains $VirtualMachine)
    {
       Write-Host "This machine is missing from Rubrik: " $VirtualMachine ##Output to confirm a difference is found.
       $exportObject = $VirtualMachine | Select-Object @{Name="VM's Missing from Rubrik";Expression={$_}} ##Create export object to handle and present data correctly
       $exportObject | Export-Csv -Path $failedResult -NoTypeInformation -append -Force ##Add the resulting VM to "Comparison Results.csv" 
    }
    else
    {
       Write-Host $VirtualMachine " is being backed up in Rubrik." 
       $exportObject = $VirtualMachine | Select-Object @{Name="VM's backed up in Rubrik";Expression={$_}}, @{Name="SLA";Expression={ $_}}  ##Create export object to handle and present data correctly
       $exportObject | Export-Csv -Path $confirmedResult -NoTypeInformation -append -Force ##Add the resulting VM to "Comparison Results.csv" 
    }
}

