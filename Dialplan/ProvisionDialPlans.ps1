############## STATIC DATA ENTRY ###################

## Create the log output folder###
$outputfolder = ".\Logs\"

#If the folder does not exist, create it.
if (-not (Test-Path -Path $outputfolder))
{
	try
	{
		$null = New-Item -ItemType Directory -Path $outputfolder -Force -ErrorAction Stop
		Write-Host "The logging folder [$outputfolder] has been created."
	}
	catch
	{
		throw $_.Exception.Message
	}
}
# If the file already exists, show the message and do nothing.
else
{
	Write-Host "The logging folder [$outputfolder] already exists. We'll use it."
}
####################################################

############# FUNCTION DEFINITIONS #################

#File prompt to select the input file
Function Get-FileName($initialDirectory)
{
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
	If ($OpenFileDialog.filename -eq '')
	{
		Write-Host "You did not choose a file. Try again." -ForegroundColor White -BackgroundColor Red
		$importfile = Get-FileName
	}
} #end function Get-FileName

############## Connect To Microsoft Teams ####################
if (Get-Module -ListAvailable -Name MicrosoftTeams)
{
	Import-Module MicrosoftTeams
	Connect-MicrosoftTeams | Out-Null
	}
else
{
	#Install Microsoft Teams Module and connect to Microsoft Teams
	$title = 'Install Microsoft Teams Module and connect to Microsoft Teams'
	$msg = 'Do you want to install the Microsoft Teams Module and connect to Microsoft Teams?'
	$options = '&Yes', '&No'
	$default = '0' # Yes=Yes, No=No
	
	do
	{
		$response = $Host.UI.PromptForChoice($title, $msg, $options, $default)
		if ($response -eq '0')
		{
			Install-Module MicrosoftTeams
			Import-Module MicrosoftTeams
			Connect-MicrosoftTeams | Out-Null
			$TeamsStringQuit = 'y'
		}
		else
		{
			Write-Host "You will not be able to enable users until you restart this script and connect to Microsoft Teams"
			$TeamsStringQuit = 'Y'
		}
	}
	until ($TeamsStringQuit -eq 'Y')
}

# Import the CSV file containing user information, DIDs, and extensions
$importfile = Get-FileName
$outputcsvfile = ".\Logs\Dialplan_Enablement_" + $(get-date -f MM-dd-yy-hh-mm-ss) + ".csv" # so that you get a new file each time you run the script
$outputerrorfile = ".\Logs\Dialplan_Enablement_Errors_" + $(get-date -f MM-dd-yy-hh-mm-ss) + ".csv" # so that you get a new file each time you run the script
$subject = "The Dial Plan Enablement Script"


$dialplans = Import-Csv $importfile

# Process each user in the CSV

foreach ($dialplan in $dialplans) {

	$DPName = $dialplan.Name
	$DPDesc = $dialplan.Description
	
    $NRName01 = $dialplan.NRName01
	$NRPattern01 = $dialplan.NRPattern01
	$NRTranslation01 = $dialplan.NRTranslation01
	$NRDescription01 = $dialplan.NRDescription01
	
    $NRName02 = $dialplan.NRName02
	$NRPattern02 = $dialplan.NRPattern02
	$NRTranslation02 = $dialplan.NRTranslation02
	$NRDescription02 = $dialplan.NRDescription02

    $NRName03 = $dialplan.NRName03
	$NRPattern03 = $dialplan.NRPattern03
	$NRTranslation03 = $dialplan.NRTranslation03
	$NRDescription03 = $dialplan.NRDescription03

    $NRName04 = $dialplan.NRName04
	$NRPattern04 = $dialplan.NRPattern04
	$NRTranslation04 = $dialplan.NRTranslation04
	$NRDescription04 = $dialplan.NRDescription04

	$NRPattern05 = $dialplan.NRPattern05
	$NRTranslation05 = $dialplan.NRTranslation05
	$NRDescription05 = $dialplan.NRDescription05

    # Check if the dial plan exists

    $existingDialplan = Get-CsTenantDialPlan -Identity $DPName -ErrorAction SilentlyContinue

        if (!$existingDialplan) {

            try {
                # Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $TeamsPhoneNumber -PhoneNumberType DirectRoutingSet-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $TeamsPhoneNumber -PhoneNumberType DirectRouting
				New-CsTenantDialPlan $DPName -Description $DPDesc
            }
            catch {
                Write-Host "Error Adding Dial Plan with name: $upn" -foregroundcolor Red
					$failure = @()
					$FailedDialPlan = "" | Select "Dial Plan", "Failure_Reason"
					$FailedDialPlan.Name = $dialplan.Name
					$FailedDialPlan.Failure_Reason = $Error[0].Exception
					$failure += $FailedDialPlan
					$failure | Export-Csv $outputerrorfile -Append -NoTypeInformation
            }
            
			$NR = @()
			if ($NRName01) {
				$NR += New-CsVoiceNormalizationRule -Name $NRName01 -Parent $DPName -Pattern $NRPattern01 -Translation $NRTranslation01 -InMemory -Description $NRDescription01
			}
			if ($NRName02) {
				$NR += New-CsVoiceNormalizationRule -Name $NRName02 -Parent $DPName -Pattern $NRPattern02 -Translation $NRTranslation02 -InMemory -Description $NRDescription02
			}
			if ($NRName03) {
				$NR += New-CsVoiceNormalizationRule -Name $NRName03 -Parent $DPName -Pattern $NRPattern03 -Translation $NRTranslation03 -InMemory -Description $NRDescription03
			}
			if ($NRName04) {
				$NR += New-CsVoiceNormalizationRule -Name $NRName04 -Parent $DPName -Pattern $NRPattern04 -Translation $NRTranslation04 -InMemory -Description $NRDescription04
			}
			if ($NRName05) {
				$NR += New-CsVoiceNormalizationRule -Name $NRName05 -Parent $DPName -Pattern $NRPattern05 -Translation $NRTranslation05 -InMemory -Description $NRDescription05
			}
			Set-CsTenantDialPlan -Identity $DPName -NormalizationRules @{add=$NR}
            

        }
		Get-CsTenantDialPlan -Identity $DPName | Select-Object SimpleName, Description, NormalizationRules | Export-Csv -Path $outputcsvfile -Append -NoTypeInformation -Force -ErrorAction SilentlyContinue

    }