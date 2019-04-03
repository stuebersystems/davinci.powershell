# Restart DAVINCI Server if DAVINCI User File has changed

# Last updated: 03.04.2019
# Version: 0.0.1

# Copyright (c) 2019 STÜBER SYSTEMS 

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to
# deal in the Software without restriction, including without limitation the
# rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
# sell copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
# IN THE SOFTWARE.

param(
	[parameter(mandatory=$true,position=0,HelpMessage="File path to json file with additional params")]
	[ValidateNotNullOrEmpty()]
	[string]
	$ParamsFileName
)

# Send email

function SendEmail{

	param(
		[parameter(mandatory=$true,position=0)]
		[ValidateNotNullOrEmpty()]
		[string]
		$Subject,

		[parameter(mandatory=$true,position=1)]
		[ValidateNotNullOrEmpty()]
		[string]
		$Body,

		[parameter(mandatory=$true,position=2)]
		[ValidateNotNull()]
		[object]
		$Params
	)

	# Is Email configured?
	if (Get-Member -inputobject $Params -name "Email" -Membertype Properties) {
	
		# Init secure password
		$SmtpSecurePassword = $Params.Email.SmtpPassword | ConvertTo-SecureString -asPlainText -Force

		# Init credential
		$Cred = New-Object System.Management.Automation.PSCredential $Params.Email.SmtpUserName, $SmtpSecurePassword
		
		# Send message via SMTP
		Send-MailMessage -SmtpServer $Params.Email.SmtpServer -UseSsl:$Params.Email.UseSsl -Credential $Cred -to $Params.Email.To -from $Params.Email.From -Subject $Subject -body $Body
	}
}

# Main method

try
{
	# Establishes and enforces coding rules 
	set-strictmode -version latest

	# Powershell should stop on erros
	$ErrorActionPreference = 'Stop'

	# Read params from JSON file
	$params = Get-Content -Raw -Path $ParamsFileName | ConvertFrom-Json
	
	# Has DAVINCI User File changed?
	if ((!$params.DavUserFile.LastSucessfullRestart) -or 
	   ((Get-item $params.DavUserFile.LocalPath).LastWriteTime -gt [datetime]$params.DavUserFile.LastSucessfullRestart)) {
		
		# Console Output
		Write-Host [DAVSERVER] Restart DAVINCI Server

		# Restart DAVINCI Server
		Restart-Service -Name daVinciServerService
		
		# Store current timestamp
		$params.DavUserFile.LastSucessfullRestart = Get-Date -Format s
		
		# Update JSON file
		$params | ConvertTo-Json -depth 10 | Out-File $ParamsFileName

		# Console Output
		Write-Host [DAVSERVER] Sucessfull restarted
	} 
	else {

		# Console Output
		Write-Host [DAVSERVER] No changes detected
	}
}
catch
{
	# Extract error message
	$ErrorMessage = $_.Exception.Message
	
	# Console Output
	Write-Host $ErrorMessage -ForegroundColor Red
	
	# Report error
	SendEmail "[RESTARTDAVSERVER] Powershell script failed with error" $ErrorMessage $params
}