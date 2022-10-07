<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 22.10.03.1829

<#
    .SYNOPSIS
        This script contains mitigations to help address the following vulnerabilities.
            CVE-2022-41040
    .DESCRIPTION
       This script has three operations it performs:
            Mitigation of CVE-2022-41040 via a URL Rewrite configuration. Note: this mitigates current known attacks.
    .PARAMETER RollbackMitigation
        If set, will only reverse the mitigations if present.
    .PARAMETER DoNotAutoUpdateEOMTv2
        If set, will not attempt to download and run latest EOMTv2 version from github.
    .EXAMPLE
		PS C:\> EOMTv2.ps1
		This will run the default mode which does the following:
            1. Checks if an updated version of EOMTv2 is available, downloads and runs latest version if so
            2. Downloads and installs the IIS URL rewrite tool.
            3. Applies the URL rewrite mitigation (only if vulnerable).
    .EXAMPLE
		PS C:\> EOMTv2.ps1 -RollbackMitigation
        This will only rollback the URL rewrite mitigation.
	.Link
        https://www.iis.net/downloads/microsoft/url-rewrite
        https://aka.ms/privacy
#>

[Cmdletbinding()]
param (
    [switch]$RollbackMitigation,
    [switch]$DoNotAutoUpdateEOMTv2
)

$ProgressPreference = "SilentlyContinue"
$EOMTv2Dir = Join-Path $env:TEMP "EOMTv2"
$EOMTv2LogFile = Join-Path $EOMTv2Dir "EOMTv2.log"
$SummaryFile = "$env:SystemDrive\EOMTv2Summary.txt"
$EOMTv2DownloadUrl = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/EOMTv2.ps1'
$versionsUrl = 'https://aka.ms/EOMTv2-VersionsUri'
$MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
$MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

#autopopulated by CSS-Exchange build
$BuildVersion = "22.10.03.1829"

# Force TLS1.2 to make sure we can download from HTTPS
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Run-Mitigate {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Invalid rule result')]
    param(
        [string]$WebSiteName = "Default Web Site",
        [string]$Stage = "MitigationProcess",
        [switch]$RollbackMitigation

    )

    function Get-MsiProductVersion {
        param (
            [string]$filename
        )

        try {
            $windowsInstaller = New-Object -com WindowsInstaller.Installer

            $database = $windowsInstaller.GetType().InvokeMember(
                "OpenDatabase", "InvokeMethod", $Null,
                $windowsInstaller, @($filename, 0)
            )

            $q = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"

            $View = $database.GetType().InvokeMember(
                "OpenView", "InvokeMethod", $Null, $database, ($q)
            )

            try {
                $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null) | Out-Null

                $record = $View.GetType().InvokeMember(
                    "Fetch", "InvokeMethod", $Null, $View, $Null
                )

                $productVersion = $record.GetType().InvokeMember(
                    "StringData", "GetProperty", $Null, $record, 1
                )

                return $productVersion
            } finally {
                if ($View) {
                    $View.GetType().InvokeMember("Close", "InvokeMethod", $Null, $View, $Null) | Out-Null
                }
            }
        } catch {
            throw "Failed to get MSI file version the error was: {0}." -f $_
        }
    }

    function Get-InstalledSoftwareVersion {
        param (
            [ValidateNotNullOrEmpty()]
            [string[]]$Name
        )

        try {
            $UninstallKeys = @(
                "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )

            New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null

            $UninstallKeys += Get-ChildItem HKU: | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object {
                "HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall"
            }

            foreach ($UninstallKey in $UninstallKeys) {
                $SwKeys = Get-ChildItem -Path $UninstallKey -ErrorAction SilentlyContinue
                foreach ($n in $Name) {
                    $SwKeys = $SwKeys | Where-Object { $_.GetValue('DisplayName') -like "$n" }
                }
                if ($SwKeys) {
                    foreach ($SwKey in $SwKeys) {
                        if ($SwKey.GetValueNames().Contains("DisplayVersion")) {
                            return $SwKey.GetValue("DisplayVersion")
                        }
                    }
                }
            }
        } catch {
            Write-Error -Message "Error: $($_.Exception.Message) - Line Number: $($_.InvocationInfo.ScriptLineNumber)"
        }
    }

    function Test-IIS10 {
        $iisRegPath = "hklm:\SOFTWARE\Microsoft\InetStp"

        if (Test-Path $iisRegPath) {
            $properties = Get-ItemProperty $iisRegPath
            if ($properties.MajorVersion -eq 10) {
                return $true
            }
        }

        return $false
    }

    function Get-URLRewriteLink {
        $DownloadLinks = @{
            "x86" = @{
                "de-DE" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_de-DE.msi"
                "en-US" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_en-US.msi"
                "es-ES" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_es-ES.msi"
                "fr-FR" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_fr-FR.msi"
                "it-IT" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_it-IT.msi"
                "ja-JP" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_ja-JP.msi"
                "ko-KR" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_ko-KR.msi"
                "ru-RU" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_ru-RU.msi"
                "zh-CN" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_zh-CN.msi"
                "zh-TW" = "https://download.microsoft.com/download/D/8/1/D81E5DD6-1ABB-46B0-9B4B-21894E18B77F/rewrite_x86_zh-TW.msi"
            }
            "x64" = @{
                "de-DE" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_de-DE.msi"
                "en-US" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"
                "es-ES" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_es-ES.msi"
                "fr-FR" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_fr-FR.msi"
                "it-IT" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_it-IT.msi"
                "ja-JP" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_ja-JP.msi"
                "ko-KR" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_ko-KR.msi"
                "ru-RU" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_ru-RU.msi"
                "zh-CN" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_zh-CN.msi"
                "zh-TW" = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_zh-TW.msi"
            }
        }

        if ([Environment]::Is64BitOperatingSystem) {
            $Architecture = "x64"
        } else {
            $Architecture = "x86"
        }

        if ((Get-Culture).Name -in @("de-DE", "en-US", "es-ES", "fr-FR", "it-IT", "ja-JP", "ko-KR", "ru-RU", "zn-CN", "zn-TW")) {
            $Language = (Get-Culture).Name
        } else {
            $Language = "en-US"
        }

        return $DownloadLinks[$Architecture][$Language]
    }

    #Configure Rewrite Rule consts
    $HttpRequestInput = '{REQUEST_URI}'
    $root = 'system.webServer/rewrite/rules'
    $inbound = '.*'
    $name = 'PowerShell - inbound'
    $pattern = '.*autodiscover\.json.*Powershell.*'
    $filter = "{0}/rule[@name='{1}']" -f $root, $name
    $site = "IIS:\Sites\$WebSiteName"
    Import-Module WebAdministration

    if ($RollbackMitigation) {
        $Message = "Starting rollback of mitigation on $env:computername"
        $RegMessage = "Starting rollback of mitigation"
        Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message

        $mitigationFound = $false
        if (Get-WebConfiguration -Filter $filter -PSPath $site) {
            $mitigationFound = $true
            Clear-WebConfiguration -Filter $filter -PSPath $site
        }

        if ($mitigationFound) {
            $Rules = Get-WebConfiguration -Filter 'system.webServer/rewrite/rules/rule' -Recurse
            if ($null -eq $Rules) {
                Clear-WebConfiguration -PSPath $site -Filter 'system.webServer/rewrite/rules'
            }

            $Message = "Rollback of mitigation complete on $env:computername"
            $RegMessage = "Rollback of mitigation complete"
            Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message
        } else {
            $Message = "Mitigation not present on $env:computername"
            $RegMessage = "Mitigation not present"
            Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message
        }
    } else {
        $Message = "Starting mitigation process on $env:computername"
        $RegMessage = "Starting mitigation process"
        Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message

        $RewriteModule = Get-InstalledSoftwareVersion -Name "*IIS*", "*URL*", "*2*"

        if ($RewriteModule) {
            $Message = "IIS URL Rewrite Module is already installed on $env:computername"
            $RegMessage = "IIS URL Rewrite Module already installed"
            Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message
        } else {
            $DownloadLink = Get-URLRewriteLink
            $DownloadPath = Join-Path $EOMTv2Dir "\$($DownloadLink.Split("/")[-1])"
            $RewriteModuleInstallLog = Join-Path $EOMTv2Dir "\RewriteModuleInstall.log"

            $response = Invoke-WebRequest $DownloadLink -UseBasicParsing
            [IO.File]::WriteAllBytes($DownloadPath, $response.Content)

            $MSIProductVersion = Get-MsiProductVersion -filename $DownloadPath

            if ($MSIProductVersion -lt "7.2.1993") {
                $Message = "Incorrect IIS URL Rewrite Module downloaded on $env:computername"
                $RegMessage = "Incorrect IIS URL Rewrite Module downloaded"
                Set-LogActivity -Error -Stage $Stage -RegMessage $RegMessage -Message $Message
                throw
            }
            #KB2999226 required for IIS Rewrite 2.1 on IIS ver under 10
            if (!(Test-IIS10) -and !(Get-HotFix -Id "KB2999226" -ErrorAction SilentlyContinue)) {
                $Message = "Did not detect the KB2999226 on $env:computername. Please review the pre-reqs for this KB and download from https://support.microsoft.com/en-us/topic/update-for-universal-c-runtime-in-windows-c0514201-7fe6-95a3-b0a5-287930f3560c"
                $RegMessage = "Did not detect KB299226"
                Set-LogActivity -Error -Stage $Stage -RegMessage $RegMessage -Message $Message
                throw
            }

            $Message = "Installing the IIS URL Rewrite Module on $env:computername"
            $RegMessage = "Installing IIS URL Rewrite Module"
            Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message

            $arguments = "/i `"$DownloadPath`" /quiet /log `"$RewriteModuleInstallLog`""
            $msiexecPath = $env:WINDIR + "\System32\msiexec.exe"

            if (!(Confirm-Signature -filepath $DownloadPath -Stage $stage)) {
                $Message = "File present at $DownloadPath does not seem to be signed as expected, stopping execution."
                $RegMessage = "File downloaded for UrlRewrite MSI does not seem to be signed as expected, stopping execution"
                Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message -Error
                Write-Summary -NoRemediation:$DoNotRemediate
                throw
            }

            Start-Process -FilePath $msiexecPath -ArgumentList $arguments -Wait
            Start-Sleep -Seconds 15
            $RewriteModule = Get-InstalledSoftwareVersion -Name "*IIS*", "*URL*", "*2*"

            if ($RewriteModule) {
                $Message = "IIS URL Rewrite Module installed on $env:computername"
                $RegMessage = "IIS URL Rewrite Module installed"
                Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message
            } else {
                $Message = "Issue installing IIS URL Rewrite Module $env:computername"
                $RegMessage = "Issue installing IIS URL Rewrite Module"
                Set-LogActivity -Error -Stage $Stage -RegMessage $RegMessage -Message $Message
                throw
            }
        }

        $Message = "Applying URL Rewrite configuration to $env:COMPUTERNAME :: $WebSiteName"
        $RegMessage = "Applying URL Rewrite configuration"
        Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message

        try {
            if ((Get-WebConfiguration -Filter $filter -PSPath $site).name -eq $name) {
                Clear-WebConfiguration -Filter $filter -PSPath $site
            }

            Add-WebConfigurationProperty -PSPath $site -filter $root -name '.' -value @{name = $name; patternSyntax = 'Regular Expressions'; stopProcessing = 'True' }
            Set-WebConfigurationProperty -PSPath $site -filter "$filter/match" -name 'url' -value $inbound
            Set-WebConfigurationProperty -PSPath $site -filter "$filter/conditions" -name '.' -value @{input = $HttpRequestInput; matchType = '0'; pattern = $pattern; ignoreCase = 'True'; negate = 'False' }
            Set-WebConfigurationProperty -PSPath $site -filter "$filter/action" -name 'type' -value 'AbortRequest'

            $Message = "Mitigation complete on $env:COMPUTERNAME :: $WebSiteName"
            $RegMessage = "Mitigation complete"
            Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message
        } catch {
            $Message = "Mitigation failed on $env:COMPUTERNAME :: $WebSiteName"
            $RegMessage = "Mitigation failed"
            Set-LogActivity -Error -Stage $Stage -RegMessage $RegMessage -Message $Message
            throw
        }
    }
}

function Write-Log {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Invalid rule result')]
    param
    (
        [string]$Message,
        [string]$Path = $EOMTv2LogFile,
        [string]$Level = "Info"
    )

    $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Write log entry to $Path
    "$FormattedDate $($Level): $Message" | Out-File -FilePath $Path -Append
}

function Set-LogActivity {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidAssignmentToAutomaticVariable', '', Justification = 'Invalid rule result')]
    [CmdletBinding(SupportsShouldProcess)]
    param (
        $Stage,
        $RegMessage,
        $Message,
        [switch]$Notice,
        [switch]$Error
    )
    if ($Notice) {
        $Level = "Notice"
    } elseif ($Error) {
        $Level = "Error"
    } else {
        $Level = "Info"
    }
    if ($Level -eq "Info") {
        Write-Verbose -Message $Message -Verbose
    } elseif ($Level -eq "Notice") {
        Write-Host -ForegroundColor Cyan -BackgroundColor black "NOTICE: $Message"
    } else {
        Write-Error -Message $Message
    }

    Write-Log -Message $Message -Level $Level
}

function Confirm-Signature {
    param(
        [string]$Filepath,
        [string]$Stage
    )

    $IsValid = $false
    $failMsg = "Signature of $Filepath not as expected. "
    try {
        if (!(Test-Path $Filepath)) {
            $IsValid = $false
            $failMsg += "Filepath does not exist"
            throw
        }

        $sig = Get-AuthenticodeSignature -FilePath $Filepath

        if ($sig.Status -ne 'Valid') {
            $IsValid = $false
            $failMsg += "Signature is not trusted by machine as Valid, status: $($sig.Status)"
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        $chainsCorrectly = $chain.Build($sig.SignerCertificate)

        if (!$chainsCorrectly) {
            $IsValid = $false
            $failMsg += "Signer certificate doesn't chain correctly"
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            $IsValid = $false
            $failMsg += "Certificate Chain shorter than expected"
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            $IsValid = $false
            $failMsg += "Top-level certifcate in chain is not a root certificate"
            throw
        }

        if ($rootCert.Certificate.Subject -eq $MicrosoftSigningRoot2010 -or $rootCert.Certificate.Subject -eq $MicrosoftSigningRoot2011) {
            $IsValid = $true
            $Message = "$Filepath is signed by Microsoft as expected, trusted by machine as Valid, signed by: $($sig.SignerCertificate.Subject), Issued by: $($sig.SignerCertificate.Issuer), with Root certificate: $($rootCert.Certificate.Subject)"
            $RegMessage = "$Filepath is signed by Microsoft as expected"
            Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message -Notice
        } else {
            $IsValid = $false
            $failMsg += "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)"
            throw
        }
    } catch {
        $IsValid = $false
        Set-LogActivity -Stage $Stage -RegMessage $failMsg -Message $failMsg -Error
    }

    return $IsValid
}
function Write-Summary {
    param(
        [switch]$Pass,
        [switch]$NoRemediation
    )

    $RemediationText = ""
    if (!$NoRemediation) {
        $RemediationText = " and clear malicious files"
    }

    $summary = @"
EOMTv2 mitigation summary
Message: Microsoft attempted to mitigate and protect your Exchange server from CVE-2022-41040 $RemediationText.
For more information on these vulnerabilities please visit (https://aka.ms/Exchangevulns2)
Please review locations and files as soon as possible and take the recommended action.
Microsoft saved several files to your system to "$EOMTv2Dir". The only files that should be present in this directory are:
    a - EOMTv2.log
    b - RewriteModuleInstall.log
    c - one of the following IIS URL rewrite MSIs:
        rewrite_amd64_[de-DE,en-US,es-ES,fr-FR,it-IT,ja-JP,ko-KR,ru-RU,zh-CN,zh-TW].msi
        rewrite_ x86_[de-DE,es-ES,fr-FR,it-IT,ja-JP,ko-KR,ru-RU,zh-CN,zh-TW].msi
        rewrite_x64_[de-DE,es-ES,fr-FR,it-IT,ja-JP,ko-KR,ru-RU,zh-CN,zh-TW].msi
        rewrite_2.0_rtw_x86.msi
        rewrite_2.0_rtw_x64.msi
1 - Confirm the IIS URL Rewrite Module is installed. This module is required for the mitigation of CVE-2022-41040, the module and the configuration (present or not) will not impact this system negatively.
    a - If installed, Confirm the following entry exists in the "$env:SystemDrive\inetpub\wwwroot\web.config". If this configuration is not present, your server is not mitigated. This may have occurred if the module was not successfully installed with a supported version for your system.
    <system.webServer>
        <rewrite>
            <rules>
                <rule name="PowerShell - inbound">
                    <match url=".*" />
                    <conditions>
                        <add input="{REQUEST_URI}" pattern=".*autodiscover\.json.*Powershell.*" />
                    </conditions>
                    <action type="AbortRequest" />
                </rule>
            </rules>
        </rewrite>
    </system.webServer>
"@

    if (Test-Path $SummaryFile) {
        Remove-Item $SummaryFile -Force
    }

    $summary = $summary.Replace("`r`n", "`n").Replace("`n", "`r`n")
    $summary | Out-File -FilePath $SummaryFile -Encoding ascii -Force
}

if (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Unable to launch EOMTv2.ps1: please re-run as administrator."
    exit
}

if ($PSVersionTable.PSVersion.Major -lt 3) {
    Write-Error "Unsupported version of PowerShell on $env:computername - The Exchange On-premises Mitigation Tool supports PowerShell 3 and later"
    exit
}

# Main
try {
    $Stage = "CheckEOMTv2Version"

    if (!(Test-Path $EOMTv2Dir)) {
        New-Item -ItemType Directory $EOMTv2Dir | Out-Null
    }

    try {
        $Message = "Checking if EOMTv2 is up to date with $versionsUrl"
        Set-LogActivity -Stage $Stage -RegMessage $Message -Message $Message
        $latestEOMTv2Version = $null
        $versionsData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequest $versionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
        $latestEOMTv2Version = ($versionsData | Where-Object -Property File -EQ "EOMTv2.ps1").Version
    } catch {
        $Message = "Cannot check version info at $versionsUrl to confirm EOMTv2.ps1 is latest version. Version currently running is $BuildVersion. Please download latest EOMTv2 from $EOMTv2DownloadUrl and re-run EOMTv2, unless you just did so. Exception: $($_.Exception)"
        $RegMessage = "Cannot check version info at $versionsUrl to confirm EOMTv2.ps1 is latest version. Version currently running is $BuildVersion. Continuing with execution"
        Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message -Notice
    }

    $DisableAutoupdateIfneeded = "If you are getting this error even with updated EOMTv2, re-run with -DoNotAutoUpdateEOMTv2 parameter";

    $Stage = "AutoupdateEOMTv2"
    if ($latestEOMTv2Version -and ($BuildVersion -ne $latestEOMTv2Version)) {
        if ($DoNotAutoUpdateEOMTv2) {
            $Message = "EOMTv2.ps1 is out of date. Version currently running is $BuildVersion, latest version available is $latestEOMTv2Version. We strongly recommend downloading latest EOMTv2 from $EOMTv2DownloadUrl and re-running EOMTv2. DoNotAutoUpdateEOMTv2 is set, so continuing with execution"
            $RegMessage = "EOMTv2.ps1 is out of date. Version currently running is $BuildVersion, latest version available is $latestEOMTv2Version.  DoNotAutoUpdateEOMTv2 is set, so continuing with execution"
            Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message -Notice
        } else {
            $Stage = "DownloadLatestEOMTv2"
            $EOMTv2LatestFilepath = Join-Path $EOMTv2Dir "EOMTv2_$latestEOMTv2Version.ps1"
            try {
                $Message = "Downloading latest EOMTv2 from $EOMTv2DownloadUrl"
                Set-LogActivity -Stage $Stage -RegMessage $Message -Message $Message
                Invoke-WebRequest $EOMTv2DownloadUrl -OutFile $EOMTv2LatestFilepath -UseBasicParsing
            } catch {
                $Message = "Cannot download latest EOMTv2.  Please download latest EOMTv2 yourself from $EOMTv2DownloadUrl, copy to necessary machine(s), and re-run. $DisableAutoupdateIfNeeded. Exception: $($_.Exception)"
                $RegMessage = "Cannot download latest EOMTv2 from $EOMTv2DownloadUrl. Stopping execution."
                Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message -Error
                throw
            }

            $Stage = "RunLatestEOMTv2"
            if (Confirm-Signature -Filepath $EOMTv2LatestFilepath -Stage $Stage) {
                $Message = "Running latest EOMTv2 version $latestEOMTv2Version downloaded to $EOMTv2LatestFilepath"
                Set-LogActivity -Stage $Stage -RegMessage $Message -Message $Message

                try {
                    & $EOMTv2LatestFilepath @PSBoundParameters
                    exit
                } catch {
                    $Message = "Run failed for latest EOMTv2 version $latestEOMTv2Version downloaded to $EOMTv2LatestFilepath, please re-run $EOMTv2LatestFilepath manually. $DisableAutoupdateIfNeeded. Exception: $($_.Exception)"
                    $RegMessage = "Run failed for latest EOMTv2 version $latestEOMTv2Version"
                    Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message -Error
                    throw
                }
            } else {
                $Message = "File downloaded to $EOMTv2LatestFilepath does not seem to be signed as expected, stopping execution."
                $RegMessage = "File downloaded for EOMTv2.ps1 does not seem to be signed as expected, stopping execution"
                Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message -Error
                Write-Summary -NoRemediation:$DoNotRemediate
                throw
            }
        }
    }

    $Stage = "EOMTv2Start"

    $Message = "Starting EOMTv2.ps1 version $BuildVersion on $env:computername"
    $RegMessage = "Starting EOMTv2.ps1 version $BuildVersion"
    Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message

    $Message = "EOMTv2 precheck complete on $env:computername"
    $RegMessage = "EOMTv2 precheck complete"
    Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message

    if ($RollbackMitigation) {
        Run-Mitigate -RollbackMitigation
    }

    else {
        $Message = "Applying mitigation on $env:computername"
        $RegMessage = ""
        Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message
        Run-Mitigate
    }

    $Message = "EOMTv2.ps1 complete on $env:computername, please review EOMTv2 logs at $EOMTv2LogFile and the summary file at $SummaryFile"
    $RegMessage = "EOMTv2.ps1 completed successfully"
    Set-LogActivity -Stage $Stage -RegMessage $RegMessage -Message $Message
    Write-Summary -Pass -NoRemediation:$DoNotRemediate #Pass
} catch {
    $Message = "EOMTv2.ps1 failed to complete on $env:computername, please review EOMTv2 logs at $EOMTv2LogFile and the summary file at $SummaryFile - $_"
    $RegMessage = "EOMTv2.ps1 failed to complete"
    Set-LogActivity -Error -Stage $Stage -RegMessage $RegMessage -Message $Message
    Write-Summary -NoRemediation:$DoNotRemediate #Fail
}

# SIG # Begin signature block
# MIIn4QYJKoZIhvcNAQcCoIIn0jCCJ84CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBlBJJ19vroQmKI
# vjIMTx6J9dppZ3GJhkFqRCpA+PfOWqCCDYEwggX/MIID56ADAgECAhMzAAACzI61
# lqa90clOAAAAAALMMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjIwNTEyMjA0NjAxWhcNMjMwNTExMjA0NjAxWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCiTbHs68bADvNud97NzcdP0zh0mRr4VpDv68KobjQFybVAuVgiINf9aG2zQtWK
# No6+2X2Ix65KGcBXuZyEi0oBUAAGnIe5O5q/Y0Ij0WwDyMWaVad2Te4r1Eic3HWH
# UfiiNjF0ETHKg3qa7DCyUqwsR9q5SaXuHlYCwM+m59Nl3jKnYnKLLfzhl13wImV9
# DF8N76ANkRyK6BYoc9I6hHF2MCTQYWbQ4fXgzKhgzj4zeabWgfu+ZJCiFLkogvc0
# RVb0x3DtyxMbl/3e45Eu+sn/x6EVwbJZVvtQYcmdGF1yAYht+JnNmWwAxL8MgHMz
# xEcoY1Q1JtstiY3+u3ulGMvhAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUiLhHjTKWzIqVIp+sM2rOHH11rfQw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDcwNTI5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAeA8D
# sOAHS53MTIHYu8bbXrO6yQtRD6JfyMWeXaLu3Nc8PDnFc1efYq/F3MGx/aiwNbcs
# J2MU7BKNWTP5JQVBA2GNIeR3mScXqnOsv1XqXPvZeISDVWLaBQzceItdIwgo6B13
# vxlkkSYMvB0Dr3Yw7/W9U4Wk5K/RDOnIGvmKqKi3AwyxlV1mpefy729FKaWT7edB
# d3I4+hldMY8sdfDPjWRtJzjMjXZs41OUOwtHccPazjjC7KndzvZHx/0VWL8n0NT/
# 404vftnXKifMZkS4p2sB3oK+6kCcsyWsgS/3eYGw1Fe4MOnin1RhgrW1rHPODJTG
# AUOmW4wc3Q6KKr2zve7sMDZe9tfylonPwhk971rX8qGw6LkrGFv31IJeJSe/aUbG
# dUDPkbrABbVvPElgoj5eP3REqx5jdfkQw7tOdWkhn0jDUh2uQen9Atj3RkJyHuR0
# GUsJVMWFJdkIO/gFwzoOGlHNsmxvpANV86/1qgb1oZXdrURpzJp53MsDaBY/pxOc
# J0Cvg6uWs3kQWgKk5aBzvsX95BzdItHTpVMtVPW4q41XEvbFmUP1n6oL5rdNdrTM
# j/HXMRk1KCksax1Vxo3qv+13cCsZAaQNaIAvt5LvkshZkDZIP//0Hnq7NnWeYR3z
# 4oFiw9N2n3bb9baQWuWPswG0Dq9YT9kb+Cs4qIIwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIZtjCCGbICAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAsyOtZamvdHJTgAAAAACzDAN
# BglghkgBZQMEAgEFAKCBxjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgRNQ60xZ6
# hj9fnQVIrKzDRaHBsQPdtrTn3YyIPuQVlFMwWgYKKwYBBAGCNwIBDDFMMEqgGoAY
# AEMAUwBTACAARQB4AGMAaABhAG4AZwBloSyAKmh0dHBzOi8vZ2l0aHViLmNvbS9t
# aWNyb3NvZnQvQ1NTLUV4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQALqwyPvd9l
# QK9x1opgRMeZm1aZl5jWNk0Qw1YDP/gBuCpyXAnw8AGLGwiGGBsjwPl1YXQI2E0m
# +z/L1N5XaMES4mNiYxomgrV2lJC0atD0G8Gech0s+44AKeIquI67seAOzwSP36nR
# wmDEKagrPU+KmMvjgC9/PtVvzR692LMaVDlPGbwME65gPT8sxmCImmCnSc86Irkv
# F4zdmFPWK5n+k3sBFuKidwYKiZ2Q511ZxL8Y5gvx5WOdHCVxo9oDuxbCOanytSn8
# CAIzrHbRAeBFpPkNRKTWs5LYlvmI80D3cSFDfrVseE3VMYaogXxq4orbSXWyUmyo
# AdXAc2cTMWzLoYIXKDCCFyQGCisGAQQBgjcDAwExghcUMIIXEAYJKoZIhvcNAQcC
# oIIXATCCFv0CAQMxDzANBglghkgBZQMEAgEFADCCAVgGCyqGSIb3DQEJEAEEoIIB
# RwSCAUMwggE/AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIOV7ZQAN
# 2AqYs0b+sShNkVa9/mulNyKM7xe7XA8Fao8wAgZjN1qp8l4YEjIwMjIxMDA0MTky
# NzAzLjg5WjAEgAIB9KCB2KSB1TCB0jELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9u
# cyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpBMjQwLTRCODItMTMw
# RTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCEXgwggcn
# MIIFD6ADAgECAhMzAAABuAjUwbh54FFJAAEAAAG4MA0GCSqGSIb3DQEBCwUAMHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIyMDkyMDIwMjIxNloXDTIz
# MTIxNDIwMjIxNlowgdIxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRl
# ZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046QTI0MC00QjgyLTEzMEUxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCcG7H8ERxEZ+QveUDxp9+7SebzomPWlo3iA8aYGl2m
# BoqrMNBtHLwit2Lpa6nphH9gpBknXDiQBfvhKoTkqqztd5xnI6sOoRTZ7BsCgDcs
# 3XcJzjtI89ZXwSCLu+023zb52lm8EkbtA0xhwO5+z0dDGCfWwvx6KUraWxX+tOiy
# nE1p8GlLhHQIcehQKoNa5ILoIrrmmsSssUZHYxWSyhVxFy/adNT7cEtHxur2YkVZ
# xCoHmpdk3i73QCKmqQcJbT2In4Kp6e2eFSTHD6NE1pkTVRTnB0TcFh3oTfcYOroC
# Blz6O4TKVqz4IOJqBztnqXNE9+q4JvYZQ65dvGxuU2pJlezSffJGxeZYLZa6ME+n
# nMBFbbi3eGtIM2KoWp1u82hWZw37ecjHscufYnxE3c6cAK49pyQMiLIqtT2N1z5t
# plDC5QApLKnpRyn1ooSvxXocP5s4qLwVcJJK/xtBdfTloEfjAVCon8KTskIiwMa6
# r1Z3wmOEIyQmYhyLkteRwkP0yLf0I+MdVr0WtVKo55qS7uP8gWFmbHb2aKvElyO4
# jrO5rnnACXhJGJTOojpDIoIcxIY0CvU4T6aALZbl2y+6iBxq3xZR5l/+LKHO449Y
# /Lgd9VN+ICZQdl6YYzGk/hPjvUiJpP+hkVfUpvAvPBZ5ptlXcDJbgjQFMzxY37x8
# dwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFMT9RyU3hcMzWKNZ437nF7lH0GluMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCfbpPJosKTKCsNw/feqYhM23oABsZAAR2J
# 9rz9oW6p5EvVPe7P+kJcmToRhbFbnWoi3kWWU7GhuYUdKArgSDWf5XpaOccx3Ppg
# TqQV6cUmltYaqMWgi7FR9RAbc+4p9uR48vno7gXIpR+hGdGbTcZlhiEM/EdALktE
# 4+FYCVxyJVz/XVToshqPVXpa5PhRsfwQvohLgminfyLMqRz4glAoedgxnPdbNku5
# XUMdR+ApYzgLWpw3270nowEli6P7N8NFzAGw7q1jZ7NF4nQBdkZy9T2saDss/VWG
# pDRiuBd/7iUWZ1YG7AmLsDV9QZksDNWWz0p4IDUhk2cfxUNtCY/7paytK8gG7zWz
# W+JJGkuGu+u4nwqr1DhS5VE/DzeN4YZXcSOshzzDliLSaRByGQYkzSm5QbGGyK4I
# W9WJvbX0rC2uWSQ81WTZ3UX4VKiTslxfglZvTrVZhQxvZCMDOkAF/ENIn+9tc+FT
# s2Tbw/JLYNZSPl414FwyVbN4eO7DLvRl0mM5Mvu3bYJnN4kT5HWt2FUXZjybTRTd
# DS8m3LLBO74RQoo++XgxkARatAOClRurgXa+lE1sBNFSh8Qc9gaRr5+wQrPucsZx
# fh1eglInJA6qj4vyCO2bLHfRGz/bs4+JbpXVwwD61hrXRqvsCtKHZRjVbgjMbC94
# Z/PiojvVIjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
# hvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# MjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAy
# MDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25Phdg
# M/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPF
# dvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
# GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBp
# Dco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
# yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3E
# XzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0
# lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1q
# GFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ
# +QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
# PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkw
# EgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxG
# NSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARV
# MFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAK
# BggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0x
# M7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmC
# VgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449
# xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
# nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDS
# PeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
# Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxn
# GSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+Crvs
# QWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokL
# jzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggLU
# MIICPQIBATCCAQChgdikgdUwgdIxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMg
# TGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046QTI0MC00QjgyLTEzMEUx
# JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUr
# DgMCGgMVAHBrXlahcfyG0ylxy3rlwj0TzA+7oIGDMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwDQYJKoZIhvcNAQEFBQACBQDm5nZGMCIYDzIwMjIx
# MDA0MTcwNzE4WhgPMjAyMjEwMDUxNzA3MThaMHQwOgYKKwYBBAGEWQoEATEsMCow
# CgIFAObmdkYCAQAwBwIBAAICApkwBwIBAAICEbIwCgIFAObnx8YCAQAwNgYKKwYB
# BAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGG
# oDANBgkqhkiG9w0BAQUFAAOBgQBb7oDnl674uCwYQIFNCuRWppAmjmCWVrZ/RfVR
# dEbep+Th3OoAnGaCxB8K3datxxTjyYxswexebD32KfUnGMUdn4pXKyvNISLhYM/x
# yLL3boFkBtGMfy4S/dZ+eAROZ3btyQZHe4hTMV/TtrsjuLCXOIx1MZPls+1X50Pc
# gKTJjjGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEw
# AhMzAAABuAjUwbh54FFJAAEAAAG4MA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG
# 9w0BCQMxDQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIHCFHHpQHa/NxEKm
# lS9Hs7CEL6tL+GzDiZe6WIz80u8qMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCB
# vQQgKOvWOKBwO0OKUvmNTbAX66SGE3lrD3hk2potF9Dw+x8wgZgwgYCkfjB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAbgI1MG4eeBRSQABAAABuDAi
# BCCaCx476aPCJCVs5EgyJeoi/tF2OIDvzYQZUWdCzCYWrjANBgkqhkiG9w0BAQsF
# AASCAgAE3rorYTszjyKUxfq1uyW/eiL8nm2YXY/8XDToIxWl176NvXInRAtdYQui
# yPl2Lzk3ebvZcnspy5FBY9WA3UCBXJlTjc+irghrz9czVRwfnHdaOttNDYV0mf2c
# DfX3OdMqQWMsCISlSUxS999jryA6lx39TeY8eSFd/XSml3mUBzxidLZzvsz2ZDF2
# tRY7RvcY/w+sZlqk/mb5mfyAGyZ/Vhfdqv6S02z3vw3RroYxvTkeHVUyXoancOYW
# RHIlRATwtq8aWLnxJw+ODRdjGap2ptD0RVqK1KoenRT87U5c3871HrDzBdC+/rt+
# D+xIT5Agb+gybwuo7WR6PFcb5fylFlnnzvaWfQn8571UNVd1q36DaZvWkXj6uZP+
# KOrKejeMg9kkUHF1fwYhRJgqDkbs4nMO6fJ8tqKYQwX3EkxgWAqnM2vD4l04wRlM
# JGneXbQ4KMOWirUNk+tuIkJP68AwHQm43vN4lO0PZGunbhp5Q1f0ClYUAKwL7duS
# +3gl430P3X75lTIFsmszWQv3egba3OMCiIzE0PUR7G+OLJXh6K0wlXWEEJyN+qkY
# tEH8kP2cmDXJaXKVhb/HPdcEhYgkl1X4oHPm50obkLldFxIBlo2/pRP0Sd8VurV/
# 4qiV4sRAlyW4Dm27Okbt8Ee4YJgDR9EmigWirmJrQR4dOkX4vw==
# SIG # End signature block
