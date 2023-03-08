#######################################################
###  Script de maintenance de serveur Windows       ###
###     créé un rapport HTML avec les points        ###
### de controles défini pour les différents roles : ###
###          AD / FS / PRINT / TS                   ### 
###                                                 ###
###  Realise par GMU avec la collaboration de STA   ###
#######################################################

#explications :
#
#Le script va creer un rapport html dans lequel il fera les check suivants : 
#
#check commun : Version systeme, uptime, check des espaces disques et check de Windows defender (Windows > 2016) 
#
#check AD : Check des services AD, Check DCdiag (FR ou EN), Check DCdiag DNS (FR ou EN), check replication DC (si plusieurs DC) 
#           les checks Dcdiag creent également un fichier .txt pour afficher le test complet
#
#check ADconnect : affichage de la version de l'adconnect actuellement installe sur le serveur
#
#check du print server : affichage de la liste des imprimantes installees avec le nombre de travaux en attentes sur chaque imprimante
#
#check RDS : affichage de l'attribution des cal RDS (nombre et expiration)
#
#check backup VBR : check du repository, check des jobs, check de la licence
#
#check backup VBO : check du repository, check des jobs, check de la licence
#
#Check windows update : affichage des dernieres MAJ avec leur statut d'installation

<#PSScriptInfo

.VERSION 1.0

.AUTHOR greg@dfi.ch

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA

#>

<#
.DESCRIPTION
 Le script va creer un rapport html dans lequel il fera les check suivants : 
check commun : Version systeme, uptime, check des espaces disques et check de Windows defender (Windows > 2016) 
check AD : Check des services AD, Check DCdiag (FR ou EN), Check DCdiag DNS (FR ou EN), check replication DC (si plusieurs DC) 
          les checks Dcdiag creent également un fichier .txt pour afficher le test complet
check ADconnect : affichage de la version de l'adconnect actuellement installe sur le serveur
check du print server : affichage de la liste des imprimantes installees avec le nombre de travaux en attentes sur chaque imprimante
check RDS : affichage de l'attribution des cal RDS (nombre et expiration)
check backup VBR : check du repository, check des jobs, check de la licence
check backup VBO : check du repository, check des jobs, check de la licence
Check windows update : affichage des dernieres MAJ avec leur statut d'installation
#>



#CSS codes
$header = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }

    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }

    
    
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }
    


    #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #613098;
        font-size: 12px;

    }



    .StopStatus {

        color: #ff0000;
    }
    
  
    .RunningStatus {

        color: #008000;
    }


    .FailedWU {

        color: #ff0000;
    }

    .SuccessWU {

        color: #008000;
    }

    .RedColor {

        color: #ff0000;
    }

    .GreenColor {

        color: #008000;
    }

    .OrangeColor {

        color: #FBB917;
    }


</style>
"@

#Variables globales
$outfile = "c:\DFI-Maintenance\Maintenance_$($env:computername)_$(get-date -format dd-MM-yy ).html"
$reportPath = "c:\DFI-Maintenance\"
$hostname=hostname

#region test path
If ((Test-Path $reportPath) -eq $false) {
	New-Item "C:\DFI-Maintenance" -itemType Directory | out-null
}
If (Test-Path $outfile) {
	Remove-Item $outfile
}
#endregion test path

#region check commun
write-host "Check Commun"
#Variables
$WindowsFeature = Get-WindowsFeature | Where Installed
$WindowsDomain = (Get-WmiObject Win32_ComputerSystem).Domain
Function Get-UpTime {
	$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
	$CurrentDate = Get-Date
	$CurrentDate - $bootuptime
}

#The command below will get the name of the computer
$ComputerName = "<h1>Maintenance du serveur : $env:computername</h1>"

#The command below will get the Operating System information, convert the result to HTML code as table and store it to a variable
$OSinfo = Get-CimInstance -Class Win32_OperatingSystem | Select Version,Caption,BuildNumber| ConvertTo-Html -As List -Fragment -PreContent "<h2>Operating System Information</h2>"

$uptime = get-uptime | select @{N="Uptime (jrs)";E={(get-uptime).days}} | ConvertTo-Html -As List -Fragment -PreContent "<h2>Uptime</h2>"

#The command below will get the details of Disk, convert the result to HTML code as table and store it to a variable
$DiscInfo=Get-CimInstance Win32_LogicalDisk | Where { ($_.DriveType -eq "3") } |  Select DeviceID, 
        @{N="Total Size(GB)";E={[math]::Round(($_.Size/1GB),2)}}, 
        @{N="Free Space(GB)";E={[math]::Round(($_.FreeSpace/1GB),2)}},
        @{N="% Free";E={[math]::Round((($_.FreeSpace/$_.Size)*100),2)}} |
 ConvertTo-Html -Fragment -PreContent "<h2>Check Disk</h2>"

 if (((Get-Variable PSVersionTable -ValueOnly).PSVersion).Major -gt "5") {
 $DefenderStatus=Get-MpComputerStatus | ConvertTo-Html AMRunningMode  -Fragment -PreContent "<h2>Check Defender AV</h2>"
 if (($DefenderStatus).AMRunningMode -match "running" -or "Normal"){
$Tips="<p>Méthode pour activer le mode passif de Defender : </p>
<p>You can set Microsoft Defender Antivirus to passive mode using a registry key as follows :</p>
<p>Path: HKLM\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection</p>
<p>Name: ForceDefenderPassiveMode</p>
<p>Type: REG_DWORDValue: 1</p> "
 }
 }

$a=Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | select DisplayName,DisplayVersion | 
 where {$_.DisplayName -match "Firefox" -or $_.DisplayName -match "Chrome" -or $_.DisplayName -match "7-Zip" -or $_.DisplayName -match "lansweeper" -or $_.DisplayName -match "notepad" -or $_.DisplayName -eq "Microsoft Edge"}
$b=gp HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |Select DisplayName,DisplayVersion | 
 where {$_.DisplayName -match "Firefox" -or $_.DisplayName -match "Chrome" -or $_.DisplayName -match "7-Zip" -or $_.DisplayName -match "lansweeper" -or $_.DisplayName -match "notepad" -or $_.DisplayName -eq "Microsoft Edge"}
$c=@()
$c+=$a
$c+=$b

$CommunSoft= $c | ConvertTo-Html -Fragment -PreContent "<h2>Logiciels Commun</h2>"


write-host "Check Commun OK"
#endregion check commun

#region AD
if (($WindowsFeature).Name -eq "AD-Domain-Services") {
write-host "Check AD"
#variables
$dclist=Get-ADDomainController -Filter * | Select Name

#check AD
write-host ".......Check 1/3 (ServicesDC)"
$DCservice=Get-Service -name ntds,adws,dns,dnscache,kdc,w32time,netlogon | convertto-html Status,Name,DisplayName -Fragment -PreContent "<h2>Check Service AD</h2>"
$DCservice = $DCservice -replace '<td>Running</td>','<td class="GreenColor">Running</td>'
$DCservice = $DCservice -replace '<td>Stopped</td>','<td class="Redcolor">Stopped</td>'


function Invoke-DcDiag {
    param(
    )
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(437)
    $res=@()
    $regex = [regex]"(?sm)\.+\sLe\stest\s(\w+).+?de\s([A-Za-z0-9_-]+)\sa\s(réussi|échoué)"
    $result = dcdiag /s:$hostname

    $allmatches = $regex.Matches($result)

    foreach($line in $allmatches){
        $ObjectDiag = New-Object System.Object
        $ObjectDiag | Add-Member -Type NoteProperty -Name TestResult -Value $line.Groups[3].Value
        $ObjectDiag | Add-Member -Type NoteProperty -Name Entity -Value $line.Groups[2].Value
        $ObjectDiag | Add-Member -Type NoteProperty -Name TestName -Value $line.Groups[1].Value
        $res+=$ObjectDiag
    }
    return $res
}

write-host ".......Check 2/3 (DCDIAG)"

#dcdiag en FR
$resultfr=(Invoke-DcDiag).count
if(($resultfr).count -gt 0){
$dcdiag=Invoke-DcDiag | ConvertTo-Html -Fragment -PreContent "<h2>Check DCDIAG</h2>"
$dcdiag = $dcdiag -replace '<td>réussi</td>','<td class="GreenColor">réussi</td>'
$dcdiag = $dcdiag -replace '<td>échoué</td>','<td class="RedColor">échoué</td>'

Invoke-DcDiag > C:\DFI-Maintenance\dcdiag_$(get-date -format dd-MM-yy).txt
write-host "test dcdiag FR effectue"
} 

#dcdiag en EN
$result = dcdiag /s:$Hostname | select-string -pattern '\. (.*) \b(passed|failed)\b test (.*)'
if(($result).count -gt 0){
$dcdiag=$result | Select-Object -Property @{Name = 'TestResult'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[2].Value } },
     @{Name = 'TestName'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[3].Value } },
     @{Name = 'Entity'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[1].Value } } |
ConvertTo-Html -Fragment -PreContent "<h2>Check DCDIAG</h2>"
$dcdiag = $dcdiag -replace '<td>passed</td>','<td class="GreenColor">passed</td>'
$dcdiag = $dcdiag -replace '<td>failed</td>','<td class="RedColor">failed</td>'

$result | Select-Object -Property @{Name = 'TestResult'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[2].Value } },@{Name = 'TestName'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[3].Value } },@{Name = 'Entity'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[1].Value } } > C:\DFI-Maintenance\dcdiagdns_$(get-date -format dd-MM-yy ).txt
write-host "test dcdiag EN effectue"
}

dcdiag /s:$hostname >> C:\DFI-Maintenance\dcdiag_$(get-date -format dd-MM-yy).txt


#dcdiagDNS en FR
function Invoke-DcDiag-DNS {
    param(
    )
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(437)
    $res=@()
    $regex = [regex]"(?sm)\.+\sLe\stest\s(\w+).+?de\s([A-Za-z0-9_-]+)\sa\s(réussi|échoué)"
    $result = dcdiag /test:dns

    $allmatches = $regex.Matches($result)

    foreach($line in $allmatches){
        $ObjectDiag = New-Object System.Object
        $ObjectDiag | Add-Member -Type NoteProperty -Name TestResult -Value $line.Groups[3].Value
        $ObjectDiag | Add-Member -Type NoteProperty -Name Entity -Value $line.Groups[2].Value
        $ObjectDiag | Add-Member -Type NoteProperty -Name TestName -Value $line.Groups[1].Value
        $res+=$ObjectDiag
    }
    return $res
}

write-host ".......Check 3/3 (DCDIAG-DNS)"
$resultdnsfr=(Invoke-DcDiag-DNS).count
if ($resultdnsfr -gt "0") {
$dcdiagdns=Invoke-DcDiag-DNS | ConvertTo-Html -Fragment -PreContent "<h2>Check DCDIAG-DNS</h2>"
$dcdiagdns = $dcdiagdns -replace '<td>"réussi"</td>','<td class="GreenColor">"réussi"</td>'
$dcdiagdns = $dcdiagdns -replace '<td>"échoué"</td>','<td class="RedColor">"échoué"</td>'

Invoke-DcDiag-DNS > C:\DFI-Maintenance\dcdiagdns_$(get-date -format dd-MM-yy).txt
}

#dcdiagDNS en EN 
$resultdns=dcdiag /test:dns | select-string -pattern '\. (.*) \b(passed|failed)\b test (.*)'
if (($resultdns).count -gt "0") {
$dcdiagdns= $resultdns | Select-Object -Property @{Name = 'TestResult'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[2].Value } },
    @{Name = 'TestName'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[3].Value } },
    @{Name = 'Entity'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[1].Value } } |
    ConvertTo-Html -Fragment -PreContent "<h2>Check DCDIAG-DNS</h2>"
$dcdiagdns = $dcdiagdns -replace '<td>passed</td>','<td class="GreenColor">passed</td>'
$dcdiagdns = $dcdiagdns -replace '<td>failed</td>','<td class="RedColor">failed</td>'

$dcdiagdns > C:\DFI-Maintenance\dcdiagdns_$(get-date -format dd-MM-yy).txt

}

#rajout du test dns complet dans le fichier précédement cree (FR ou EN)
dcdiag /test:dns >> C:\DFI-Maintenance\dcdiagdns_$(get-date -format dd-MM-yy).txt

if ((($dclist).name).count -gt "1") {
write-host "Check AD Replication"
#check AD Replication metdata
write-host ".......Check 1/2 (Replication metadata)"
$ReplicationMetadata=Get-ADReplicationPartnerMetadata -target ($dclist).name | ConvertTo-Html Server,IntersiteTransportType,LastReplicationAttempt,LastReplicationResult,LastReplicationSuccess,Partner,SyncOnStartup -Fragment -PreContent "<h2>Replication AD metadata</h2>"
$ReplicationMetadata = $ReplicationMetadata -replace '<td>0</td>','<td class="GreenColor">success</td>'
$ReplicationMetadata = $ReplicationMetadata -replace '<td>1</td>','<td class="Redcolor">1</td>'

write-host ".......Check 2/2 (Replication schema)"
$ReplicationDC=Get-ADReplicationPartnerMetadata -Target ($dclist).name -Partition Schema -PartnerType Both | select Server,@{n="Partner";e={(Resolve-DnsName $_.PartnerAddress).NameHost}},Partition,LastReplicationResult,PartnerType | ConvertTo-Html -Fragment -PreContent "<h2>Check Replication DC</h2>"
$ReplicationDC = $ReplicationDC -replace '<td>0</td>','<td class="GreenColor">success</td>'

write-host "Check AD Replication OK"
}

write-host "Check AD OK"
}
#endregion AD

#region EventID

#check event ID 4625
write-host "Check eventID"
$events = Get-WinEvent -FilterHashtable @{logname="Security"; id=4625}
$resultArray = @() # Créer un tableau vide pour stocker tous les résultats
foreach ($event in $events) {
    $event = [xml]$events[0].ToXml()
    $eventArray = @{}
    $event.Event.EventData.Data | Where-Object {$_.name -eq "TargetUserName" -or $_.name -eq "Status" -or $_.name -eq "WorkstationName" -or $_.name -eq "IpAddress"} | ForEach-Object { $eventArray[$_.name] = $_.'#text' }
    $systemObject = $event.Event.system | Select-Object EventID,@{N="TimeCreated";E={($_.TimeCreated).SystemTime}}

    $combinedObject = New-Object -TypeName PSObject
    foreach ($property in $eventArray.Keys) {
        $combinedObject | Add-Member -MemberType NoteProperty -Name $property -Value $eventArray[$property]
    }
    foreach ($property in $systemObject.psobject.Properties) {
        $combinedObject | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
    }
    $resultArray += $combinedObject # Ajouter l'objet combiné au tableau des résultats
}
$Event4625=$resultArray | convertto-html -Fragment -PreContent "<h2>Check EventID</h2>" 
write-host "Check eventID OK"


#endregion EventID

#region ADConnect

$AdconnectInsatall=Get-WmiObject -Class Win32_Product | Select Name, Version | where { $_.Name -eq "Microsoft Azure AD Connect" } -ErrorAction SilentlyContinue
if ((($AdconnectInsatall).name).count -gt "0") { 
write-host "Check ADConnect"
$Adconnect=Get-WmiObject -Class Win32_Product | where { $_.Name -eq "Microsoft Azure AD Connect" }| select Name,Version | convertto-html -Fragment -PreContent "<h2>Check ADconnect</h2>"
write-host "Check ADconnect OK"
}

#endregion ADConnect

#region Print Server
#Printer
if (($WindowsFeature).Name -eq "Print-Server") {
write-host "Check Printer"
$Printers=Get-Printer | ConvertTo-Html Name,JobCount,Type,PortName,Shared,Published -Fragment -PreContent "<h2>Print Server</h2>"
write-host "Check Printer OK"
}
#endregion Print server

#region RDS
if (($WindowsFeature).Name -eq "RDS-Licensing") {
write-host "Check RDS"
$tsLicense = Get-CIMInstance -computername $server Win32_TSLicenseKeyPack -filter "TotalLicenses!=0" | ? { $_.TypeAndModel -ne 'Built-in TS Per Device CAL' }
$CALS=$tsLicense | select TypeAndModel, TotalLicenses, IssuedLicenses, AvailableLicenses, ProductVersion, ExpirationDate | convertto-html -Fragment -PreContent "<h2>Check RDS Licence</h2>"
write-host "Check RDS OK"
}

#endregion RDS

#region Backup VBR
$VBRInstall=Get-WmiObject -Class Win32_Product | Select Name, Version | where { $_.Name -match "Veeam Backup & Replication" } -ErrorAction SilentlyContinue

if ((($VBRInstall).name).count -gt "1") {
write-host "Check Backup VBR"
Connect-VBRServer
write-host ".......Check 1/3 (VBR Repository)"
$VBRRepo=Get-VBRBackupRepository | Select Name,Path,CloudProvider,IsAvailable,VersionOfCreation | ConvertTo-Html -Fragment -PreContent "<h2>Check VBR repository</h2>"

write-host ".......Check 2/3 (VBR Job)"
$VBRJob = Get-VBRJob | Select Name,JobType,TargetDir,TargetFile | ConvertTo-Html -Fragment -PreContent "<h2>Check VBR job</h2>"
$VBRJob = $VBRJob -replace '<td>success</td>','<td class="GreenColor">success</td>'
$VBRJob = $VBRJob -replace '<td>warning</td>','<td class="OrangeColor">warning</td>'
$VBRJob = $VBRJob -replace '<td>failed</td>','<td class="Redcolor">failed</td>'

write-host ".......Check 3/3 (VBR Licence)"
$VBRLicence=Get-VBRInstalledLicense | Select Status,Type,Edition,SupportID,SupportExpirationDate | ConvertTo-Html -Fragment -PreContent "<h2>Check VBR licence</h2>"
write-host "Check Backup VBR OK"
}
#endregion Backup VBR

#region Backup VBO365
Function Import-VBOModule {
Import-Module "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1"
Import-Module "C:\Program Files\Veeam\Backup and Replication\Explorers\Exchange\Veeam.Exchange.PowerShell\Veeam.Exchange.PowerShell.psd1"
Import-Module "C:\Program Files\Veeam\Backup and Replication\Explorers\SharePoint\Veeam.SharePoint.PowerShell\Veeam.SharePoint.PowerShell.psd1"
Import-Module "C:\Program Files\Veeam\Backup and Replication\Explorers\Teams\Veeam.Teams.PowerShell\Veeam.Teams.PowerShell.psd1"
Connect-VBOServer
}


$VBOInstall=Get-WmiObject -Class Win32_Product | Select Name, Version | where { $_.Name -eq "Veeam Backup for Microsoft 365" } -ErrorAction SilentlyContinue
if ((($VBOInstall).name).count -eq "1") {
write-host "Check Backup VBO"
Import-VBOModule

write-host ".......Check 1/3 (VBO Repository)"
$VBORepo=Get-VBORepository | Select Name,Path,IsOutdated,
@{n = "Capacity (GB)"; e ={ [math]::Round($_.Capacity / 1GB, 2) } }, 
@{n = "FreeSpace (GB)"; e = { [math]::Round($_.FreeSpace / 1GB, 2) } }, 
@{n = "UsedSpace (GB)"; e = { [math]::Round((($_.Capacity - $_.FreeSpace) / 1GB), 2) } }, 
@{name = "PercentFree (GB)"; expression = { [math]::Round(($_.FreeSpace / $_.Capacity) * 100, 2) } } |
ConvertTo-Html -Fragment -PreContent "<h2>Check VBO365 repository</h2>"

write-host ".......Check 2/3 (VBO Job)"
$VBOJob=Get-VBOJob | Select Name,Repository,LastStatus,NextRun,IsEnabled | ConvertTo-Html -Fragment -PreContent "<h2>Check VBO365 job</h2>"

write-host ".......Check 3/3 (VBO Licence)"
$VBOLicence=Get-VBOLicense | Select Status,Type,TotalNumber | ConvertTo-Html -Fragment -PreContent "<h2>Check VBO365 licence</h2>"
Disconnect-VBOServer
write-host "Check Backup VBO OK"
}

#endregion Backup VBO365

#region WindowsUpdate
write-host "Check Windows Update"
$session = (New-Object -ComObject 'Microsoft.Update.Session')
$history = $session.QueryHistory("", 0, 50) | where {$_.Title -notmatch "Defender"}
$ReportWU=$history | Select Date, Title,
@{N="Result";E={$_.ResultCode -replace '1', 'Success' -replace '2', 'success' -replace '3', 'Success with Warning' -replace '4', 'Failed' -replace '5', 'Last failed install attempt' }} |
ConvertTo-Html -Fragment -PreContent "<h2>Dernières mises à jours Windows Installées</h2>"
$ReportWU = $ReportWU -replace '<td>Success</td>','<td class="SuccessWU">Success</td>'
$ReportWU = $ReportWU -replace '<td>Failed</td>','<td class="FailedWU">Failed</td>'
write-host "Check Windows Update OK"
#endregion Windows Update

write-host "Rapport généré ici : " $outfile

#region Rapport  
#The command below will combine all the information gathered into a single HTML report
$Report = ConvertTo-HTML -Body "$ComputerName $OSinfo $uptime $DiscInfo $DefenderStatus $Tips $CommunSoft" -Head $header -Title "Computer Information Report"

if (($WindowsFeature).Name -eq "AD-Domain-Services") {
$ReportAD = ConvertTo-HTML -Body "$DCservice $dcdiag $dcdiagdns"
if ((($dclist).name).count -gt "1") {
$ReportADRep = ConvertTo-HTML -Body "$ReplicationMetadata $ReplicationDC"
}
}

$ReportEventID= ConvertTo-HTML -Body "$Event4625"

if ((($AdconnectInsatall).name).count -gt "0") { 
$ReportADConnect = ConvertTo-HTML -Body "$Adconnect"
}

if (($WindowsFeature).Name -eq "Print-Server") {
$ReportPrinter = ConvertTo-HTML -Body "$Printers"
}

if (($WindowsFeature).Name -eq "RDS-Licensing") {
$ReportRDS = ConvertTo-HTML -Body "$CALS"
}

if ((($VBRInstall).name).count -gt "1") {
$ReportVBR=ConvertTo-HTML -Body "$VBRRepo $VBRJob $VBRLicence"
}


if ((($VBOInstall).name).count -eq "1") {
$ReportVBO=ConvertTo-HTML -Body "$VBORepo $VBOJob $VBOLicence"
}

$ReportWU = ConvertTo-HTML -Body "$ReportWU" -PostContent "<p id='CreationDate'>Creation Date: $(Get-Date)</p>"
#endregion Rapport

#region Generate file
#The command below will generate the report to an HTML file
$Report | Out-File $outfile
$ReportAD | Add-Content $outfile
$ReportADRep | Add-Content $outfile
$ReportEventID | Add-Content $outfile
$ReportADConnect | Add-Content $outfile
$ReportPrinter | Add-Content $outfile
$ReportRDS | Add-Content $outfile
$ReportVBR | Add-Content $outfile
$ReportVBO | Add-Content $outfile
$ReportWU | Add-Content $outfile

#endregion Generate file


