##*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*##
## Script Quick Guide                                                                                                    ##
## -- Open this script within "Windows Powershell ISE to execute.                                                        ##
## -- This script is to be run on a local machine only. DO NOT RUN ON SERVERS.                                           ##
##                                                                                                                       ##
## Script Notes Still To Be Written Script Notes Still To Be Written Script Notes Still To Be Written Script Notes Still ##
##*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*##
##DV-All Customer Cards = Gallagher
##customer_contact = ServiceNow

###### Importing Data ######

##Get Current Script Directory. Works in both ISE and executed script. 
function Get-ScriptDirectory {if ($psise) {Split-Path $psise.CurrentFile.FullPath}
    else {$global:PSScriptRoot}}

##Set scripts home location found in above function. 
$scriptHome = Get-ScriptDirectory

##Generate a path to export the report next to the script.
$exportPath = $scriptHome + "\Result\ComparisonResult.csv"
$exportGal = $scriptHome + "\Result\GalComp.csv"
$exportSN = $scriptHome + "\Result\SNComp.csv"


$resultMatch = $scriptHome + "\Result\inBoth.csv"
$resultGal = $scriptHome + "\Result\MissingFromSN.csv"
$resultSN = $scriptHome + "\Result\MissingFromGal.csv"

##Importing Gallagher Data - Selecting First Name, Last Name, and 'Job Title' which holds users card number. 
$_GData = Import-CSV -Path $scriptHome"\DV-All Customer Cards.csv" | Select-Object 'First Name', 'Last Name', 'Job Title' 

##Importing Users from ServiceNow Data - Selecting Name and Pass Number
$_SNData = Import-CSV -Path $scriptHome"\customer_contact.csv" | Select-Object 'name', 'Data Centre Pass No'

###### Data Maintenance ######
#
##Create a new FullName field for easier comparison between data sets
$_GData | Add-Member -Membertype noteproperty -Name Name -Value 0
$_GData | Add-Member -MemberType NoteProperty -Name CardNo -Value 0
#
##Do the same for ServiceNow Data
$_SNData | Add-Member -MemberType NoteProperty -Name CardNo -Value 0
#
##############################

##Cleaning up Gallagher data
foreach ($_Person in $_GData){
    ##Filling in Full Name field.
    $_Person.Name = $_Person.'First Name' + " " + $_Person.'Last Name'
    ##Removing blank spaces from Card Numbers. 
    $_Person.'Job Title' = $_Person.'Job Title' -replace '\s','' 
    $_Person.CardNo = $_Person.'Job Title'
}

foreach ($_Card in $_SNData)
{
    $_Card.CardNo = $_Card.'Data Centre Pass No';
}

##Tidy up Usernames in Service Now data
foreach ($_User in $_SNData){
    $_User.name = $_User.name -replace '  ',' '
}

$_GComp = $_GData | Select-Object Name, CardNo   
$_SNComp = $_SNData | Select-Object Name, CardNo 

$_GComp | Export-Csv -Path $exportGal -NoTypeInformation -append -Force
$_SNComp| Export-Csv -Path $exportSN -NoTypeInformation -append -Force

Compare-Object $_GComp $_SNComp -Property Name, cardno -IncludeEqual | Export-Csv -Path $exportPath -NoTypeInformation -append -Force

$_CompResult = Import-CSV -Path $scriptHome"\Result\ComparisonResult.csv" | Select-Object name, cardno, sideindicator


foreach ($_result in $_CompResult)
{ 
    if("<=" -in $_result.SideIndicator){
       Write-Host "Only Gallagher: " $_result.name " " $_result.cardno ##Output to confirm a difference is found.
       $exportObject = $_result | Select-Object @{Name="User:";Expression={$_.name}},@{Name="Card Number: ";Expression={$_.cardno}},@{Name="Gallagher";Expression={'X'}},@{Name="ServiceNow";Expression={''}} ##Create export object to handle and present data correctly
       $exportObject | Export-Csv -Path $resultGal -NoTypeInformation -append -Force ##Add the resulting VM to "Comparison Results.csv" 
    }
    elseif("=>" -in $_result.SideIndicator){
       Write-Host "Only ServiceNow: " $_result.name " " $_result.cardno ##Output to confirm a difference is found.
       $exportObject = $_result | Select-Object @{Name="User:";Expression={$_.name}},@{Name="Card Number: ";Expression={$_.cardno}},@{Name="Gallagher";Expression={''}},@{Name="ServiceNow";Expression={'X'}} ##Create export object to handle and present data correctly
       $exportObject | Export-Csv -Path $resultGal -NoTypeInformation -append -Force ##Add the resulting VM to "Comparison Results.csv" 
    }
    elseif("==" -in $_result.SideIndicator){
       Write-Host "Present in both: " $_result.name " " $_result.cardno ##Output to confirm a difference is found.
       $exportObject = $_result | Select-Object @{Name="User:";Expression={$_.name}},@{Name="Card Number: ";Expression={$_.cardno}},@{Name="Gallagher";Expression={'X'}},@{Name="ServiceNow";Expression={'X'}} ##Create export object to handle and present data correctly
       $exportObject | Export-Csv -Path $resultGal -NoTypeInformation -append -Force ##Add the resulting VM to "Comparison Results.csv" 
    }
}