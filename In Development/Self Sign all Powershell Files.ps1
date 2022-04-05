$Filepath = Get-ChildItem -Recurse -Path "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell*" -Include *.ps1

ForEach ($File in $Filepath) {
    Set-AuthenticodeSignature -FilePath $Filepath -Certificate $codeCertificate -TimeStampServer "http://timestamp.digicert.com"
  
  }
