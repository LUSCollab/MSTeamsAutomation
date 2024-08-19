## Create the log output folder###
$outputfolder = "C:\Users\" + $env:username + "\Documents\Logs\"

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

###################################################
#################### Connect To Azure AD ###############################
$title = 'Connect to Azure AD for User Validation'
$msg = 'Do you want to connect to Azure AD to license users?'
$options = echo Yes No
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

# Import the CSV file
$importfile = Get-FileName
$outputcsvfile = "C:\Users\" + $env:username + "\Documents\Logs\Delegates_" + $(get-date -f MM-dd-yy-hh-mm-ss) + ".csv" # so that you get a new file each time you run the script
$outputerrorfile = "C:\Users\" + $env:username + "\Documents\Logs\Delegates_Errors_" + $(get-date -f MM-dd-yy-hh-mm-ss) + ".csv" # so that you get a new file each time you run the script

$users = Import-Csv $importfile

# Process each user
foreach ( $user in $users ) {
    $upn = $user.UserPrincipalName
    $delegates = $user.Delegates
    $delegates = $delegates.split(",")
    # Verifies the user exists in Entra AD
    $existingUser = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue

    if ($existingUser) {
        try {
            $del = @()
            foreach ( $delegate in $delegates ){ 
                New-CsUserCallingDelegate -Identity $upn -Delegate $delegate -MakeCalls $true -ReceiveCalls $true -ManageSettings $true -ErrorAction Stop 
                $del += $delegate
            }
            # Set call forward to also ring delegates
            Set-CsUserCallingSettings -Identity $upn -IsForwardingEnabled $true -ForwardingType "Simultaneous" -ForwardingTargetType "MyDelegates"

            Write-Host "Successfully assigned delegates to user: $upn" -ForegroundColor Green
            $success = @()
            $SuccessUser = "" | Select "UPN", "Result"
            $SuccessUser.UPN = $user.UserPrincipalName
            $SuccessUser.Result = "Successfully added delegate(s) $del"
            $success += $SuccessUser
            $success | Export-Csv $outputcsvfile -Append -NoTypeInformation
        }
        catch {
            Write-Host "Error Enabling delegates for user with UPN $upn" -foregroundcolor Red
            $failure = @()
            $FailedUser = "" | Select "UPN", "Error_Msg"
            $FailedUser.UPN = $user.UserPrincipalName
            $FailedUser.Error_Msg = $Error[0].Exception
            $failure += $FailedUser
            $failure | Export-Csv $outputerrorfile -Append -NoTypeInformation
        } 
    } else {
        Write-Host "User with UPN $upn not found in Azure Active Directory. Please ensure the user exists before running the script."
        $failure = @()
        $FailedUser = "" | Select "UPN", "Error_Msg"
        $FailedUser.UPN = $user.UserPrincipalName
        $FailedUser.Error_Msg = "User with UPN $upn not found in Azure Active Directory"
        $failure += $FailedUser
        $failure | Export-Csv $outputerrorfile -Append -NoTypeInformation
    }
}