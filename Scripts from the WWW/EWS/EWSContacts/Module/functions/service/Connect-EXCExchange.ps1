﻿function Connect-EXCExchange {
	<#
	.SYNOPSIS
		A brief description of the Connect-EXCExchange function.
	
	.DESCRIPTION
		A detailed description of the Connect-EXCExchange function.
	
	.PARAMETER MailboxName
		Target mailbox that the ContactList will be synced to; Email address of an Office 365 business user who is in the same directory as the admin credentials
	
	.PARAMETER Credentials
		Office 365 administrator credentials -- This account needs to have application impersonation permission

	.PARAMETER ModernAuth
		Enables Modern Authenication, See https://docs.microsoft.com/en-us/office365/enterprise/office-365-client-support-modern-authentication 
	
	.PARAMETER ClientId
		Used for ModernAuth/OAuth		
	
	.EXAMPLE
		PS C:\> Connect-EXCExchange -MailboxName 'value1' -Credentials $Credentials
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $False)]
		[System.Management.Automation.PSCredential]
		$Credentials,

		[Parameter(Position = 2, Mandatory = $False)]
		[bool]
		$ModernAuth,
		
		[Parameter(Position = 3, Mandatory = $False)]
		[String]
		$ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
	)
	Begin {
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		if (Test-Path ($script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll")) {
			Import-Module ($script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll")
			$Script:EWSDLL = $script:ModuleRoot + "/bin/Microsoft.Exchange.WebServices.dll"
			write-verbose ("Using EWS dll from Local Directory")
		}
		else {

			
			## Load Managed API dll  
			###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
			$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
			if (Test-Path $EWSDLL) {
				Import-Module $EWSDLL
				$Script:EWSDLL = $EWSDLL 
			}
			else {
				"$(get-date -format yyyyMMddHHmmss):"
				"This script requires the EWS Managed API 1.2 or later."
				"Please download and install the current version of the EWS Managed API from"
				"http://go.microsoft.com/fwlink/?LinkId=255472"
				""
				"Exiting Script."
				exit


			} 
		}
		
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
		
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
		
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential
		if ($ModernAuth) {
			Write-Verbose("Using Modern Auth")
			if ([String]::IsNullOrEmpty($ClientId)) {
				# This is the application ID for Microsoft Office; Source: https://docs.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2
				$ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
			}		
			Import-Module ($script:ModuleRoot + "/bin/Microsoft.IdentityModel.Clients.ActiveDirectory.dll") -Force
			$Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext("https://login.microsoftonline.com/common")
			if ($null -eq $Credentials) {
				$PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Auto  
				
				$token = ($Context.AcquireTokenAsync("https://outlook.office365.com", $ClientId , "urn:ietf:wg:oauth:2.0:oob", $PromptBehavior)).Result
				$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
			}else{
				$AADcredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList  $Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString()
				$token = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($Context,"https://outlook.office365.com",$ClientId,$AADcredential).result
				$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token.AccessToken)
			}
		}
		else {
			Write-Verbose("Using Negotiate Auth")
			if(!$Credentials){$Credentials = Get-Credential}
			$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(), $Credentials.GetNetworkCredential().password.ToString())
			$service.Credentials = $creds
		}

		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		#$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler = $Provider.CreateCompiler()
		$Params = New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable = $False
		$Params.GenerateInMemory = $True
		$Params.IncludeDebugInformation = $False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
		
		$TASource = @'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@
		$TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
		$TAAssembly = $TAResults.CompiledAssembly
		
		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
		
		## end code from http://poshcode.org/624
		
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		
		#CAS URL Option 1 Autodiscover  
		$service.AutodiscoverUrl($MailboxName, { $true })
		#Write-host ("Using CAS Server : " + $Service.url)
		
		#CAS URL Option 2 Hardcoded  
		
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		
		## Optional section for Exchange Impersonation
		$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if (!$service.URL) {
			throw "Error connecting to EWS"
		}
		else {
			return $service
		}
	}
}