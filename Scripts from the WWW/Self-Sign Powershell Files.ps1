# Generate a self-signed Authenticode certificate in the local computer's personal certificate store.
$authenticode = New-SelfSignedCertificate -Subject "DidITBetter Software" -CertStoreLocation Cert:\LocalMachine\My -Type CodeSigningCert



# Add the self-signed Authenticode certificate to the computer's root certificate store.
## Create an object to represent the LocalMachine\Root certificate store.
$rootStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("Root","LocalMachine")
## Open the root certificate store for reading and writing.
 $rootStore.Open("ReadWrite")
## Add the certificate stored in the $authenticode variable.
 $rootStore.Add($authenticode)
## Close the root certificate store.
 $rootStore.Close()
 
# Add the self-signed Authenticode certificate to the computer's trusted publishers certificate store.
## Create an object to represent the LocalMachine\TrustedPublisher certificate store.
 $publisherStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("TrustedPublisher","LocalMachine")
## Open the TrustedPublisher certificate store for reading and writing.
 $publisherStore.Open("ReadWrite")
## Add the certificate stored in the $authenticode variable.
 $publisherStore.Add($authenticode)
## Close the TrustedPublisher certificate store.
 $publisherStore.Close()



 # Confirm if the self-signed Authenticode certificate exists in the computer's Personal certificate store
 Get-ChildItem Cert:\LocalMachine\My | Where-Object {$_.Subject -eq "CN=DidITBetter Software"}
# Confirm if the self-signed Authenticode certificate exists in the computer's Root certificate store
 Get-ChildItem Cert:\LocalMachine\Root | Where-Object {$_.Subject -eq "CN=DidITBetter Software"}
# Confirm if the self-signed Authenticode certificate exists in the computer's Trusted Publishers certificate store
 Get-ChildItem Cert:\LocalMachine\TrustedPublisher | Where-Object {$_.Subject -eq "CN=DidITBetter Software"}







# Get the code-signing certificate from the local computer's certificate store with the name *DidITBetter Software* and store it to the $codeCertificate variable.
$codeCertificate = Get-ChildItem Cert:\LocalMachine\My | Where-Object {$_.Subject -eq "CN=DidITBetter Software"}

# Sign the PowerShell script
# PARAMETERS:
# FilePath - Specifies the file path of the PowerShell script to sign, eg. C:\ATA\myscript.ps1.
# Certificate - Specifies the certificate to use when signing the script.
# TimeStampServer - Specifies the trusted timestamp server that adds a timestamp to your script's digital signature. Adding a timestamp ensures that your code will not expire when the signing certificate expires.
Set-AuthenticodeSignature -FilePath "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell\Add2Exchange\Auto_Upgrade_Add2Exchange.ps1" -Certificate $codeCertificate -TimeStampServer "http://timestamp.digicert.com"