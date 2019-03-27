if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true


Set-ADUser lorih -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADUser MartyW -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADUser BradK -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADUser JimT -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADUser JimH -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADUser SteveP -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADUser EdB -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADUser dand -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser BarryL -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser RyanR -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser toolwatch -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser JimL -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser PaulW -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jeffg -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser guyb -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser JimmyL -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser CraigR -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser BrendaW -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser JeffB -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser TimS -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser amandam -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser joshm -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser SteveT -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser codetwo -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser areuter -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser dberthene -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser matts -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jgatherer -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser eharvey -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser ablel -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser dpohlmann -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser bvizanko -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser danderson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser ctabano -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser astawowy -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser lizs -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jpeper -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mmarchio -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser asamford -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser sbader -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cnewcombe -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rmorrish -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jbjelland -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser emalarek -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jhouser -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser johno -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser praymaker -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser scook -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jrowe -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cmusson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mhoxie -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mader -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser lidleman -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mlaflamme -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser ewidstrand -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser kstober -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser wmiley -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser pamm -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser xyang -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jerjavec -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser stracy -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser williaml -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rickv -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cfath -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser hoswald -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser shawnl -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser pblanton -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser aschwartz -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser pschrupp -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser brandonf -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser dhannu -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser JLucas -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser psweeney -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cseguin -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser ljensen -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser msaerr -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser crlakesuperior -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser expensereports -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser crvoyageur -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser accountspayable -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mwaters -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jficke -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser payments -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser crerickson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser AccountsReceivable -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser crstpaul -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser dingram -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser crglacier -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser ldavison -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cmelton -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser airwatch -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser helpdeskuser -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser insurancecerts -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser Contracts778888 -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser tpierce -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mlevy -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser apeterson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser dschoch -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mmitchell -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rbloom -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mheld -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mdeschamp -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser duluth -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser bwcorpcalendar -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser kharder -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser tharvin -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser knicholas -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser wayofwellbeing -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser payroll -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rshetrom -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser handlehelpuser -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser scobb -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jdownie -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser ngoldsmith -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jschmaedeke -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser klarson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser nkleist -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser areinarts -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jbradley -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rpeek -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rmartinez -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser bmaybee -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rusery -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser hking -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser wtindall -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jsternberg -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser aroski -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jculver -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser emurray -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cmsbd -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rbailey -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser asmith -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser safety -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jmurray -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser gjamieson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cyang -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser bhoover -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser njohnson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mrich -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser tharsch -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser csmith -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser eengstrom -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser kboumeester -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jhagley -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser pchurchill -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser thinz -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jphillips -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser vstallbaumer -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rramold -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rkrause -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser enimlos -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser acarlson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser mdraper -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser erlead -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser pouchida -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser bgipson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser apope -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser acarollo -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser nviikinsalo -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jlitchy -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser clarson -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser tsehlinger -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cmckay -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser phenwood -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser rspatz -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser jmeisner -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser bsanders -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser cmcclellan -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser ktaylor -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}
Set-ADuser eweaver -Remove @{msExchDelegateListLink = “CN=zadd2exchange,OU=Service Accounts,DC=baywest,DC=local”}


Write-Host Done
Write-Host Quitting
Get-PSSession | Remove-PSSession
Exit

# End Scripting