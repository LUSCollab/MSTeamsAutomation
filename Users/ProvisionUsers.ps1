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
	}
} #end function Get-FileName

###################################################
#################### Connect To Azure AD ###############################
$title = 'Connect to Azure AD for User Validation'
$msg = 'Do you want to connect to Azure AD to license users?'
$options = Write-Output Yes No
$default = 0 # Yes=Yes, No=No

do
{
	$response = $Host.UI.PromptForChoice($title, $msg, $options, $default)
	if ($response -eq '0')
	{
		if ((Get-Module -ListAvailable -Name Microsoft.Graph.Users) -and (Get-Module -ListAvailable Microsoft.Graph.Authentication))
		{
			Import-Module Microsoft.Graph.Authentication
            Import-Module Microsoft.Graph.Users
			Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All"
			$AzureStringQuit = 'y'
			$ConnectAzureAD = 'N'
		}
		else
		{
			$AzureADmsg = 'The Microsoft Graph module is not installed on this machine. Do you wish to install it now? [Y/N]'
			do
			{
				$AzureADResponse = Read-Host -Prompt $AzureADmsg
				if ($AzureADResponse -like 'y*')
				{
					Install-Module Microsoft.Graph.Users
                    Import-Module Install-Module Microsoft.Graph.Authentication
					Import-Module Install-Module Microsoft.Graph.Users
					$ConnectAzureAD = 'N'
					$AzureStringQuit = 'y'
				}
				else
				{
					Write-Host "You will not be able to Validate users until you install this module and connect to Azure AD"
					$ConnectAzureAD = 'N'
					$AzureStringQuit = 'y'
				}
			}
			until ($AzureStringQuit -eq 'y')
		}
	}
	else
	{
		Write-Host "You will not be able to validate users until you restart this script and connect to Azure AD"
		$ConnectAzureAD = 'N'
	}
}
until ($ConnectAzureAD -eq 'N')


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
$outputcsvfile = ".\Logs\User_Enablement_" + $(get-date -f MM-dd-yy-hh-mm-ss) + ".csv" # so that you get a new file each time you run the script
$outputerrorfile = ".\Logs\User_Enablement_Errors_" + $(get-date -f MM-dd-yy-hh-mm-ss) + ".csv" # so that you get a new file each time you run the script

$users = Import-Csv $importfile

# Process each user in the CSV
foreach ($user in $users) {

    $upn = $user.UserPrincipalName
    $phoneNumber = $user.PhoneNumber
	$extension = $user.Extension
	$phoneType = $user.Type
	$callHold = $user.CallHoldPolicy
	$callPark = $user.CallParkPolicy
	$callId = $user.CallerIdPolicy
	$calling = $user.CallingPolicy
	$voiceRoute = $user.VoiceRoutePolicy
	$emerRoute = $user.EmergencyCallRoutingPolicy
	$location = $user.Location
	$dialPlan = $user.DialPlan

    # Check if the user exists in Azure Active Directory
    # $existingUser = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue

    if ($existingUser) {
        # Set Phone number (DID)
        if ($phoneNumber) {
            try {
				
				# Formats the user's phone number if an extension is to be assigned
				if($extension) {
					$TeamsPhoneNumber = $phoneNumber + ";ext=" + $extension
				} else {
					$TeamsPhoneNumber = $phoneNumber
				}

				# Get location information from MS Teams
				$loc=Get-CsOnlineLisLocation -Description $location
				
				# If the location returns multiple results, assign the first result as the user's 'default' emergency location
				if ($loc.GetType().BaseType.Name -eq "Array") {
					Set-CsPhoneNumberAssignment -PhoneNumber $TeamsPhoneNumber -Identity $upn -PhoneNumberType $phoneType -LocationId $loc.LocationId[0] -ErrorAction Stop    
				} else {
					Set-CsPhoneNumberAssignment -PhoneNumber $TeamsPhoneNumber -Identity $upn -PhoneNumberType $phoneType -LocationId $loc.LocationId -ErrorAction Stop
				}
				
				# Assign a call hold policy
				if ($callHold) {
					Grant-CsTeamsCallHoldPolicy -Identity $upn -PolicyName $callHold
				}

				# Assign a call park policy
				if ($callPark) {
					Grant-CsTeamsCallParkPolicy -Identity $upn -PolicyName $callPark
				}

				# Assign a calling ID policy
				if ($callId) {
					Grant-CsCallingLineIdentity -Identity $upn -PolicyName $callId
				}

				# Assign a calling policy
				if ($calling) {
					Grant-CsTeamsCallingPolicy -Identity $upn -PolicyName $calling
				}

				# Assign a voice routing policy
				if ($voiceRoute) {
					Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $voiceRoute
				}
				
				# Assign a emergency call routing policy
				if ($emerRoute) {
					Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $upn -PolicyName $emerRoute
				}
				
				# Assign a dial plan
				if ($dialPlan) {
					Grant-CsTenantDialPlan -Identity $upn -PolicyName $dialPlan
				}

				Write-Host "Successfully assigned user: $upn" -ForegroundColor Green
				$success = @()
				$SuccessUser = "" | Select "UPN", "Result"
				$SuccessUser.UPN = $user.UserPrincipalName
				$SuccessUser.Result = "Successfully Configured"
				$success += $SuccessUser
				$success | Export-Csv $outputcsvfile -Append -NoTypeInformation
            }
            catch {
                Write-Host "Error Enabling Phone for user with UPN $upn" -foregroundcolor Red
					$failure = @()
					$FailedUser = "" | Select "UPN", "Error_Msg"
					$FailedUser.UPN = $user.UserPrincipalName
					$FailedUser.Error_Msg = $Error[0].Exception
					$failure += $FailedUser
					$failure | Export-Csv $outputcsvfile -Append -NoTypeInformation
            }
        }
	}
    else {
        
        Write-Host "User with UPN $upn not found in Azure Active Directory. Please ensure the user exists before running the script."
        $failure = @()
					$FailedUser = "" | Select "UPN", "Failure_Reason"
					$FailedUser.UPN = $user.UserPrincipalName
					$FailedUser.Failure_Reason = "User with UPN $upn not found in Azure Active Directory"
					$failure += $FailedUser
					$failure | Export-Csv $outputerrorfile -Append -NoTypeInformation

    }
}