# Upload changed DAVINCI file / DAVINCI User file to FTP / WebDAV

# Last updated: 04.04.2019
# Version: 0.0.2

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
		
		$UseSsl = [System.Convert]::ToBoolean($Params.Email.UseSsl)

		# Init credential
		$Cred = New-Object System.Management.Automation.PSCredential $Params.Email.SmtpUserName, $SmtpSecurePassword
		
		# Send message via SMTP
		Send-MailMessage -SmtpServer $Params.Email.SmtpServer -UseSsl:$UseSsl -Credential $Cred -To $Params.Email.To -From $Params.Email.From -Subject $Subject -Body $Body
	}
}

# Uploads file to FTP

function UploadFileToFTP{

	param(
		[parameter(mandatory=$true,position=0)]
		[ValidateNotNullOrEmpty()]
		[string]
		$LocalFileName,

		[parameter(mandatory=$true,position=1)]
		[ValidateNotNull()]
		[object]
		$Params
	)

	# Is FTP configured?
	if (Get-Member -inputobject $Params -name "Ftp" -Membertype Properties) {
	
		# Init source and dest path
		$srcPath = $LocalFileName + ".tmp"
		$destUrl = $Params.Ftp.Url.TrimEnd('/') + "/" + (Split-Path $LocalFileName -leaf)

		# Create temporary copy of file
		Copy-Item $LocalFileName -Destination $srcPath
		
		# Create Uri object
		$uri = New-Object System.Uri($destUrl) 	
		
		# Upload via FTP
		$webclient = New-Object System.Net.WebClient 
		$webclient.Credentials = New-Object System.Net.NetworkCredential($Params.Ftp.Username, $Params.Ftp.Password)  
		$webclient.UploadFile($uri, $srcPath) 
		
		# Remove temporary copy file
		Remove-Item -path $srcPath
	}
}

# Uploads file to WebDAV

function UploadFileToWebDav{

	param(
		[parameter(mandatory=$true,position=0)]
		[ValidateNotNullOrEmpty()]
		[string]
		$LocalFileName,

		[parameter(mandatory=$true,position=1)]
		[ValidateNotNull()]
		[object]
		$Params
	)

	# Is WebDAV configured?
	if (Get-Member -inputobject $Params -name "WebDav" -Membertype Properties) {
	
		# Init source and dest path
		$srcPath = $LocalFileName + ".tmp"
		$destUrl = $Params.WebDav.Url.TrimEnd('/') + "/" + (Split-Path $LocalFileName -leaf)

		# Create temporary copy of file
		Copy-Item $LocalFileName -Destination $srcPath
		
		# Create Uri object
		$uri = New-Object System.Uri($destUrl) 	
		
		# Upload via FTP
		$webclient = New-Object System.Net.WebClient 
		$webclient.Credentials = New-Object System.Net.NetworkCredential($Params.WebDav.UserName, $Params.WebDav.Password)  
		$webclient.UploadFile($uri, "PUT", $srcPath) 

		# Remove temporary copy file
		Remove-Item -path $srcPath
	}
}

# Uploads file to shared folder

function UploadFileToSharedFolder{

	param(
		[parameter(mandatory=$true,position=0)]
		[ValidateNotNullOrEmpty()]
		[string]
		$LocalFileName,

		[parameter(mandatory=$true,position=1)]
		[ValidateNotNull()]
		[object]
		$Params
	)

	# Is shared folder configured?
	if (Get-Member -inputobject $Params -name "Shared" -Membertype Properties) {
	
		# Init source and dest path
		$srcPath = $LocalFileName + ".tmp"
		$destPath = $Params.Shared.FolderName.TrimEnd('\') + "\" + (Split-Path $LocalFileName -leaf)

		# Create temporary copy of file
		Copy-Item $LocalFileName -Destination $srcPath
		
		# Init secure password
		$SecurePassword = $Params.Shared.Password | ConvertTo-SecureString -asPlainText -Force
	
		# Init credential
		$Cred = New-Object System.Management.Automation.PSCredential $Params.Shared.UserName, $SecurePassword
		
		# Map shared folder to drive
		New-PSDrive -Name $Params.Shared.DriveName -PSProvider FileSystem -Root $Params.Shared.FolderName -Credential $Cred | Out-Null

		# Create temporary copy of file
		Copy-Item $srcPath -Destination $destPath

		# Unmap drive
		Remove-PSDrive -Name $Params.Shared.DriveName

		# Remove temporary copy file
		Remove-Item -path $srcPath
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
	
	# Has DAVINCI File changed?
	if ((!$params.DavFile.LastSucessfullUpload) -or 
	   ((Get-item $params.DavFile.LocalPath).LastWriteTime -gt [datetime]$params.DavFile.LastSucessfullUpload)) {
		
		# Console Output
		Write-Host [DAVFILE] Upload $params.DavFile.LocalPath
		
		# Upload DAVINCI File to FTP
		UploadFileToFtp $params.DavFile.LocalPath $params

		# Upload DAVINCI File to WebDAV
		UploadFileToWebDav $params.DavFile.LocalPath $params

		# Upload DAVINCI File to Shared Folder
		UploadFileToSharedFolder $params.DavFile.LocalPath $params

		# Store current timestamp
		$params.DavFile.LastSucessfullUpload = Get-Date -Format s
		
		# Update JSON file
		$params | ConvertTo-Json -depth 10 | Out-File $ParamsFileName

		# Console Output
		Write-Host [DAVFILE] Sucessfull uploaded
	} 
	else {

		# Console Output
		Write-Host [DAVFILE] No changes detected
		
	}
		
	# Has DAVINCI User File changed?
	if ((!$params.DavUserFile.LastSucessfullUpload) -or 
	   ((Get-item $params.DavUserFile.LocalPath).LastWriteTime -gt [datetime]$params.DavUserFile.LastSucessfullUpload)) {
		
		# Console Output
		Write-Host [DAVUSERFILE] Upload $params.DavUserFile.LocalPath

		# Upload DAVINCI User File to FTP
		UploadFileToFtp $params.DavUserFile.LocalPath $params

		# Upload DAVINCI User File to WebDAV
		UploadFileToWebDav $params.DavUserFile.LocalPath $params
		
		# Upload DAVINCI User File to Shared Folder
		UploadFileToSharedFolder $params.DavUserFile.LocalPath $params

		# Store current timestamp
		$params.DavUserFile.LastSucessfullUpload = Get-Date -Format s
		
		# Update JSON file
		$params | ConvertTo-Json -depth 10 | Out-File $ParamsFileName

		# Console Output
		Write-Host [DAVUSERFILE] Sucessfull uploaded
	} 
	else {

		# Console Output
		Write-Host [DAVUSERFILE] No changes detected
	}
}
catch
{
	# Extract error message
	$ErrorMessage = $_.Exception.Message
	
	# Console Output
	Write-Host $ErrorMessage -ForegroundColor Red
	
	# Report error
	SendEmail "[UPLOADDAVFILES] Powershell script failed with error" $ErrorMessage $params
}