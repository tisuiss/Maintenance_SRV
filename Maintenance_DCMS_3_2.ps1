<#
 Script de maintenance de serveur Windows       
 cree un rapport HTML avec les points        
 de controles dfini pour les diffrents roles : 
          AD / FS / PRINT / TS / Exchange                  
                                                 
 Realise par GMU avec la collaboration de STA   
#>

<#PSScriptInfo

.VERSION 3.2

.GUID 577f518e-7ba0-45d4-9358-1ebe1702e8e8

.AUTHOR greg

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


#> 

<# 

.DESCRIPTION 
Le script va creer un rapport html dans lequel il fera les check suivants : 

check commun : Version systeme, uptime, check des espaces disques et check de Windows defender (Windows > 2016) 

check AD : Check des services AD, Check DCdiag (FR ou EN), Check DCdiag DNS (FR ou EN), check replication DC (si plusieurs DC) 
          les checks Dcdiag creent galement un fichier .txt pour afficher le test complet

check ADconnect : affichage de la version de l'adconnect actuellement installe sur le serveur

check du print server : affichage de la liste des imprimantes installees avec le nombre de travaux en attentes sur chaque imprimante

check RDS : affichage de l'attribution des cal RDS (nombre et expiration)

check backup VBR : check du repository, check des jobs, check de la licence

check backup VBO : check du repository, check des jobs, check de la licence

check exchange : check services, check SSL, lancement des health check

Check windows update : affichage des dernieres MAJ avec leur statut d'installation

#> 
Param()

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


    h3 {

        font-family: Arial, Helvetica, sans-serif;
        color: #3333AD;
        font-size: 14px;

    }
    
    p {
        margin-left:40px
        font-family: Arial, Helvetica, sans-serif;
        font-size: 12px;
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

        td.Green{background: #7FFF00;}
        td.Yellow{background: #FFE600;}
        td.Red{background: #FF0000; color: #ffffff;}
        td.Info{background: #85D4FF;}

        td.pass{background: #7FFF00;}
        td.warn{background: #FFE600;}
        td.fail{background: #FF0000; color: #ffffff;}
        td.info{background: #85D4FF;}


</style>
"@

#Variables globales
write-host "Version du script de maintenance : 3.2"
$transcript = "yes"
$reportPath = "c:\DFI-Maintenance\"
$outfile = "$($reportPath)Maintenance_$($env:computername)_$(get-date -format yy-MM-dd).html"
$outfile4625 = "$($reportPath)events\event4625_$(get-date -format dd-MM-yy).html"
$outfile4769 = "$($reportPath)events\event4769_$(get-date -format dd-MM-yy).html"
$outfileMailboxFile = "$($reportPath)ExchangeSizeReport\MailboxSize_$(get-date -format dd-MM-yy).html"
$healthreportPath = "c:\DFI-Maintenance\HealthChecker"
$ExchangeSizeReport = "c:\DFI-Maintenance\ExchangeSizeReport"
$hostname=hostname
$CurrentDate = Get-Date

#region test path
If ((Test-Path $reportPath) -eq $false) {
	New-Item "C:\DFI-Maintenance" -itemType Directory | out-null
}
If (Test-Path $outfile) {
	Remove-Item $outfile
}
#endregion test path

#region activation de la transcription
If ($transcript -eq "yes"){
If ((Test-Path "C:\temp") -eq $false) {
	New-Item "C:\temp" -itemType Directory | out-null
}
Start-Transcript -Path "C:\temp\MaintenanceDCMS.txt" -ErrorAction SilentlyContinue
}
#endregion activation de la transcription

#region check commun
write-host "Recuperation des donnees preliminaires"
$Appslist=Get-WmiObject -Class Win32_Product | Select Name, Version
write-host "Check Commun"
#Variables
$WindowsFeature = Get-WindowsFeature | Where Installed
$WindowsDomain = (Get-WmiObject Win32_ComputerSystem).Domain
Function Get-UpTime {
	$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
	$CurrentDate - $bootuptime
}

#The command below will get the name of the computer
$ComputerName = "<h1>Maintenance du serveur : $env:computername</h1>"

#The command below will get the Operating System information, convert the result to HTML code as table and store it to a variable
$OS = Get-CimInstance -Class Win32_OperatingSystem | Select Version,Caption,BuildNumber
$OSinfo = $OS| ConvertTo-Html -As List -Fragment -PreContent "<h2>Operating System Information</h2>"
if (($OS).caption -match "Server 2012"){
$WarningOS="<p>Be careful, Windows Server 2012 will not be updated from June 2023!</p>"
}


# Obtenir les informations sur la RAM et les CPU
$ram = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | select Sum
$cpu = Get-WmiObject Win32_Processor | Measure-Object -Property NumberOfLogicalProcessors | select Count

$OSressources = @()
$OSressources += New-Object PSObject -Property @{
    "Composant" = "RAM"
    "Valeur" = "{0} Go" -f ($ram.Sum / 1GB)
}
$OSressources += New-Object PSObject -Property @{
    "Composant" = "CPU"
    "Valeur" = "{0}" -f $cpu.Count
}
$OSressources = ConvertTo-Html -As List -Fragment

#affichage uptime
$uptime = get-uptime | select @{N="Uptime (days)";E={$_.days}} | ConvertTo-Html -As List -Fragment -PreContent "<h2>Uptime</h2>" -Property "Uptime (days)"


#The command below will get the details of Disk, convert the result to HTML code as table and store it to a variable
$DiscInfo=Get-CimInstance Win32_LogicalDisk | Where { ($_.DriveType -eq "3") } |  Select DeviceID, 
        @{N="Total Size(GB)";E={[math]::Round(($_.Size/1GB),2)}}, 
        @{N="Free Space(GB)";E={[math]::Round(($_.FreeSpace/1GB),2)}},
        @{N="% Free";E={[math]::Round((($_.FreeSpace/$_.Size)*100),2)}} |
 ConvertTo-Html -Fragment -PreContent "<h2>Check Disk</h2>"

 if (((Get-Variable PSVersionTable -ValueOnly).PSVersion).Major -gt "5" -or ((Get-Variable PSVersionTable -ValueOnly).PSVersion).Major -eq "5") {
 $DefenderStatus=Get-MpComputerStatus -erroraction SilentlyContinue | ConvertTo-Html AMRunningMode -As List -Fragment -PreContent "<h2>Check Defender AV</h2>"
 if ((($DefenderStatus).AMRunningMode).count -gt "0") {
 if (($DefenderStatus).AMRunningMode -eq "running" -or ($DefenderStatus).AMRunningMode -eq "Normal"){
$Tips="<p>How to activate the passive mode of Defender : </p>
<p>You can set Microsoft Defender Antivirus to passive mode using a registry key as follows :</p>
<p>Path: HKLM\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection</p>
<p>Name: ForceDefenderPassiveMode</p>
<p>Type: REG_DWORDValue: 1</p> "
 }
 }
 }

#Check des logiciels
#Dernière versions des softs 
$software = "Greenshot","Firefox", "Chrome", "Edge", "Microsoft 365 Apps", "7-Zip", "Adobe", "lansweeper", "notepad++", "VMware Tools", "Microsoft Azure AD Connect","Veeam Backup & Replication", "Veeam Backup for Microsoft 365"

$Web = Invoke-WebRequest https://ninite.com/news
$Web.content > c:\temp\versions.txt
$versions = Get-Content c:\temp\versions.txt -raw

$versions = $versions -replace '<[^>]+>',''
$versions = $versions -replace 'about|press|updates|security|terms|privacy|jobs|&copy; Secure by Design Inc',''
$versions = $versions -replace '<[^>]+>',''
$versions = $versions -split "`r?`n" | Where-Object { $_.Trim() -ne '' } | Select-Object -Skip 7
$versions = $versions -replace 'updated to','%'
# Séparation des lignes en utilisant le caractère de saut de ligne
$lines = $versions -split "`r?`n"
# Filtrage des lignes contenant le caractère '%'
$filteredLines = $lines | Where-Object { $_ -match '%' }
# Reconstitution de la variable avec les lignes filtrées
$versions = $filteredLines -join "`r`n"

foreach ($version in $versions){
$versions = $version -split '%'
}

# Séparation des valeurs en utilisant le caractère de saut de ligne
$values = $versions -split "`r?`n"

# Tableau pour stocker les valeurs en deux colonnes
$tableau = @()

# Parcourir les valeurs et les ajouter au tableau en deux colonnes
for ($i = 0; $i -lt $values.Count; $i += 2) {
    $colonne1 = $values[$i]
    $colonne2 = $values[$i + 1]
    $tableau += [PsCustomObject]@{
        Soft = $colonne1
        Version = $colonne2
    }
}

$filteredVersions = @()

foreach ($logiciel in $software) {
    $filteredVersion = $tableau | Where-Object { $_.Soft -like "*$logiciel*" } | Sort-Object -Property Version -Descending | Select-Object -First 1
    $filteredVersions += $filteredVersion
}

$LastVersions = $filteredVersions | ConvertTo-Html -Fragment -PreContent "<h2>Check Logiciels</h2><h3>Derniere versions des logiciels</h3>"


#Soft installés sur le serveur
$registryObjects = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall" |
    Get-ItemProperty |
    Select-Object DisplayName, DisplayVersion

$filteredApps = @()  # Tableau vide

foreach ($soft in $software) {
    $filtered = $registryObjects | Where-Object { $_.DisplayName -like "*$soft*" }
    $filteredApps += $filtered  # Ajouter les résultats filtrés au tableau
}

$CommunSoft = $filteredApps | Select DisplayName, DisplayVersion | ConvertTo-Html -Fragment -PreContent "<h3>Logiciels installés</h3>"


#Report commun check
$Report = ConvertTo-HTML -Body "$ComputerName $OSinfo $WarningOS $OSressources $uptime $DiscInfo $DefenderStatus $Tips $LastVersions $CommunSoft" -Head $header -Title "Computer Information Report"
$Report | Out-File $outfile

write-host "Check Commun OK"
write-host "=====> Report can be open : " $outfile
#endregion check commun

#region AD
if (($WindowsFeature).Name -eq "AD-Domain-Services") {
write-host "Check AD"
#variables
$dclist=Get-ADDomainController -Filter * | Select Name

#check AD
#region service
write-host ".......Check 1/3 (ServicesDC)"
$DCservice=Get-Service -name ntds,adws,dns,dnscache,kdc,w32time,netlogon | convertto-html Status,Name,DisplayName -Fragment -PreContent "<h2>Check AD</h2><h3>Check Service AD</h3>"
$DCservice = $DCservice -replace '<td>Running</td>','<td class="GreenColor">Running</td>'
$DCservice = $DCservice -replace '<td>Stopped</td>','<td class="Redcolor">Stopped</td>'
#endregion service

write-host ".......Check 2/3 (DCDIAG)"

#region dcdiag 
$testDCdiag=dcdiag /e /v /c /skip:systemlog /skip:dns /s:$hostname
function Invoke-DcDiag {
    param(
    )
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(437)
    $res=@()
    $regex = [regex]"(?sm)\.+\sLe\stest\s(\w+).+?de\s([A-Za-z0-9_-]+)\sa\s(réussi|échoué)"
    $allmatches = $regex.Matches($testDCdiag)

    foreach($line in $allmatches){
        $ObjectDiag = New-Object System.Object
        $ObjectDiag | Add-Member -Type NoteProperty -Name TestResult -Value $line.Groups[3].Value
        $ObjectDiag | Add-Member -Type NoteProperty -Name Entity -Value $line.Groups[2].Value
        $ObjectDiag | Add-Member -Type NoteProperty -Name TestName -Value $line.Groups[1].Value
        $res+=$ObjectDiag
    }
    return $res
}

#region dcdiag en FR
$resultfr=Invoke-DcDiag
if(($resultfr).count -gt "0"){
$dcdiag = $resultfr | ConvertTo-Html -Fragment -PreContent "<h3>Check DCDIAG</h3>" -PostContent "<p>Details of the event can be found in C:\DFI-Maintenance\dcdiag </p>"
$dcdiag = $dcdiag -replace '<td>réussi</td>','<td class="GreenColor">réussi</td>'
$dcdiag = $dcdiag -replace '<td>échoué</td>','<td class="RedColor">échoué</td>'
} 
#endregion dcdiag en FR

#region dcdiag en EN
$resulten = $testDCdiag | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)'
if(($resulten).count -gt 0){
$dcdiag=$resulten | Select-Object -Property @{Name = 'TestResult'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[2].Value } },
     @{Name = 'TestName'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[3].Value } },
     @{Name = 'Entity'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[1].Value } } |
ConvertTo-Html -Fragment -PreContent "<h3>Check DCDIAG</h3>" -PostContent "<p>Details of the event can be found in C:\DFI-Maintenance\dcdiag </p>"
$dcdiag = $dcdiag -replace '<td>passed</td>','<td class="GreenColor">passed</td>'
$dcdiag = $dcdiag -replace '<td>failed</td>','<td class="RedColor">failed</td>'
}
#endregion dcdiag en EN
    If ((Test-Path "$($reportPath)dcdiag") -eq $false) {
        New-Item "$($reportPath)dcdiag" -itemType Directory | out-null
    }

$testDCdiag > "$($reportPath)dcdiag\dcdiag_$(get-date -format dd-MM-yy).txt"

#endregion dcdiag

#region dcdiagDNS
$testDCdiagDNS=dcdiag /test:dns /v /e
#region dcdiagDNS en FR
function Invoke-DcDiag-DNS {
    param(
    )
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(437)
    $res=@()
    $regex = [regex]"(?sm)\.+\sLe\stest\s(\w+).+?de\s([A-Za-z0-9_-]+)\sa\s(réussi|échoué)"
    $result = $testDCdiagDNS

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
$resultdnsfr=Invoke-DcDiag-DNS
if (($resultdnsfr).count -gt "0") {
$dcdiagdns = $resultdnsfr | ConvertTo-Html -Fragment -PreContent "<h3>Check DCDIAG-DNS</h3>" -PostContent "<p>Details of the event can be found in C:\DFI-Maintenance\dcdiag </p>"
$dcdiagdns = $dcdiagdns -replace '<td>réussi</td>','<td class="GreenColor">réussi</td>'
$dcdiagdns = $dcdiagdns -replace '<td>échoué</td>','<td class="RedColor">échoué</td>'
}

#endregion dcdiagDNS en FR

#region dcdiagDNS en EN 
$resultdns=$testDCdiagDNS | select-string -pattern '\. (.*) \b(passed|failed)\b test (.*)'
if (($resultdnsen).count -gt "0") {
$dcdiagdns= $resultdnsen | Select-Object -Property @{Name = 'TestResult'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[2].Value } },
    @{Name = 'TestName'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[3].Value } },
    @{Name = 'Entity'; Expression = { ($_.Line | Select-String -Pattern '\. (.*) \b(passed|failed)\b test (.*)').Matches.Groups[1].Value } } |
    ConvertTo-Html -Fragment -PreContent "<h3>Check DCDIAG-DNS</h3>" -PostContent "<p>Details of the event can be found in C:\DFI-Maintenance\dcdiag </p>"
$dcdiagdns = $dcdiagdns -replace '<td>passed</td>','<td class="GreenColor">passed</td>'
$dcdiagdns = $dcdiagdns -replace '<td>failed</td>','<td class="RedColor">failed</td>'
}
#endregion dcdiagDNS en EN

#creation du rapport complet dcdiagDNS
    If ((Test-Path "$($reportPath)dcdiag") -eq $false) {
        New-Item "$($reportPath)dcdiag" -itemType Directory | out-null
    }
$testDCdiagDNS > "$($reportPath)dcdiag\dcdiagdns_$(get-date -format dd-MM-yy).txt" 
#endregion dcdiagDNS

#region Kerberos
$krbtgt = Get-ADUser -Identity krbtgt -Properties PasswordLastSet
$lastPasswordSet = $krbtgt.PasswordLastSet
$accountName = $krbtgt.Name
$daysSinceLastChange = ($CurrentDate - $lastPasswordSet).Days

$result = New-Object -TypeName PSObject -Property @{
    'Nom du compte' = $accountName
    'Date de la dernière modification du mot de passe' = $lastPasswordSet
    'Nombre de jours écoulés depuis la dernière modification' = $daysSinceLastChange
}

if (($result).'Nombre de jours écoulés depuis la dernière modification' -gt "365"){
$KerberosAccount = $result | ConvertTo-Html -As Table -Fragment -PreContent "<h2>Kerberos Account</h2>" -PostContent "<p><span id='attention'>Attention</span> : Le mot de passe Kerberos a plus de 1 an et doit être changé</p><script>document.getElementById('attention').style.color = 'red'; document.getElementById('attention').style.fontWeight = 'bold';</script>"
}
if (($result).'Nombre de jours écoulés depuis la dernière modification' -gt "300"){
$KerberosAccount = $result | ConvertTo-Html -As Table -Fragment -PreContent "<h2>Kerberos Account</h2>" -PostContent "<p>Attention : Pensez à changer le mot de passe Kerberos</p>"
}
if (($result).'Nombre de jours écoulés depuis la dernière modification' -lt "300"){
$KerberosAccount = $result | ConvertTo-Html -As Table -Fragment -PreContent "<h2>Kerberos Account</h2>"
}

#endregion Kerberos

if ((($dclist).name).count -gt "1") {
write-host "Check AD Replication"
#check AD Replication metdata
write-host ".......Check 1/2 (Replication metadata)"
$ReplicationMetadata=Get-ADReplicationPartnerMetadata -target ($dclist).name | ConvertTo-Html Server,IntersiteTransportType,LastReplicationAttempt,LastReplicationResult,LastReplicationSuccess,Partner,SyncOnStartup -Fragment -PreContent "<h3>Replication AD metadata</h3>"
$ReplicationMetadata = $ReplicationMetadata -replace '<td>0</td>','<td class="GreenColor">success</td>'
$ReplicationMetadata = $ReplicationMetadata -replace '<td>1</td>','<td class="Redcolor">1</td>'

write-host ".......Check 2/2 (Replication schema)"
$ReplicationDC=Get-ADReplicationPartnerMetadata -Target ($dclist).name -Partition Schema -PartnerType Both | select Server,@{n="Partner";e={(Resolve-DnsName $_.PartnerAddress).NameHost}},Partition,LastReplicationResult,PartnerType | ConvertTo-Html -Fragment -PreContent "<h3>Check Replication DC</h3>"
$ReplicationDC = $ReplicationDC -replace '<td>0</td>','<td class="GreenColor">success</td>'

write-host "Check AD Replication OK"
}

#CHeck Backup AD
function Convert-TimeToDays {
    [CmdletBinding()]
    param (
        $StartTime,
        $EndTime,
        #[nullable[DateTime]] $StartTime, # can't use this just yet, some old code uses strings in StartTime/EndTime.
        #[nullable[DateTime]] $EndTime, # After that's fixed will change this.
        [string] $Ignore = '*1601*'
    )
    if ($null -ne $StartTime -and $null -ne $EndTime) {
        try {
            if ($StartTime -notlike $Ignore -and $EndTime -notlike $Ignore) {
                $Days = (NEW-TIMESPAN -Start $StartTime -End $EndTime).Days
            }
        } catch {}
    } elseif ($null -ne $EndTime) {
        if ($StartTime -notlike $Ignore -and $EndTime -notlike $Ignore) {
            $Days = (NEW-TIMESPAN -Start (Get-Date) -End ($EndTime)).Days
        }
    } elseif ($null -ne $StartTime) {
        if ($StartTime -notlike $Ignore -and $EndTime -notlike $Ignore) {
            $Days = (NEW-TIMESPAN -Start $StartTime -End (Get-Date)).Days
        }
    }
    return $Days
}
function Get-WinADLastBackup {
    [cmdletBinding()]
    param(
        [string[]] $Domains
    )
    $NameUsed = [System.Collections.Generic.List[string]]::new()
    [DateTime] $CurrentDate = Get-Date
    if (-not $Domains) {
        try {
            $Forest = Get-ADForest -ErrorAction Stop
            $Domains = $Forest.Domains
        } catch {
            Write-Warning "Get-WinADLastBackup - Failed to gather Forest Domains $($_.Exception.Message)"
        }
    }
    foreach ($Domain in $Domains) {
        try {
            [string[]]$Partitions = (Get-ADRootDSE -Server $Domain -ErrorAction Stop).namingContexts
            [System.DirectoryServices.ActiveDirectory.DirectoryContextType] $contextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain
            [System.DirectoryServices.ActiveDirectory.DirectoryContext] $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext($contextType, $Domain)
            [System.DirectoryServices.ActiveDirectory.DomainController] $domainController = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($context)
        } catch {
            Write-Warning "Get-WinADLastBackup - Failed to gather partitions information for $Domain with error $($_.Exception.Message)"
        }

        $Output = ForEach ($Name in $Partitions) {
            if ($NameUsed -contains $Name) {
                continue
            } else {
                $NameUsed.Add($Name)
            }
            $domainControllerMetadata = $domainController.GetReplicationMetadata($Name)
            $dsaSignature = $domainControllerMetadata.Item("dsaSignature")
            $LastBackup = [DateTime] $($dsaSignature.LastOriginatingChangeTime)
            [PSCustomObject] @{
                Domain            = $Domain
                NamingContext     = $Name
                LastBackup        = $LastBackup
                LastBackupDaysAgo = - (Convert-TimeToDays -StartTime ($CurrentDate) -EndTime ($LastBackup))
            }
        }
        $Output
    }
}

$BackupAD = Get-WinADLastBackup | ConvertTo-Html -Fragment -PreContent "<h3>Check Backup AD</h3>"

#Report check AD
if (($WindowsFeature).Name -eq "AD-Domain-Services") {
$ReportAD = ConvertTo-HTML -Body "$KerberosAccount $DCservice $dcdiag $dcdiagdns $BackupAD"
if ((($dclist).name).count -gt "1") {
$ReportADRep = ConvertTo-HTML -Body "$ReplicationMetadata $ReplicationDC"
}
}
$ReportAD | Add-Content $outfile
$ReportADRep | Add-Content $outfile

write-host "Check AD OK"
}
#endregion AD

#region EventID
write-host "Check eventID"
If ((Test-Path "$($reportPath)events") -eq $false) {
    New-Item "$($reportPath)events" -itemType Directory | out-null
    }

write-host ".......Check 1/4 (Resume Application Event)"
$ApplicationEvent=Get-WinEvent -LogName Application -FilterXPath "*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= '2592000000']]]" | Group-Object Id, ProviderName | Select-Object @{Name='ProviderName';Expression={$_.Group[0].ProviderName}}, @{Name='Id';Expression={$_.Group[0].Id}}, @{Name='Count';Expression={$_.Count}} | Sort-Object Count -Descending | convertto-html -Fragment -PreContent "<h2>Check EventID</h2><h3>Event Application</h3>" -PostContent "<p>Evenements sur les 30 derniers jours</p>"

write-host ".......Check 2/4 (Resume System Event)"
$SystemEvent=Get-WinEvent -LogName System -FilterXPath "*[System[(Level=1 or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= '2592000000']]]" | Group-Object Id, ProviderName | Select-Object @{Name='ProviderName';Expression={$_.Group[0].ProviderName}}, @{Name='Id';Expression={$_.Group[0].Id}}, @{Name='Count';Expression={$_.Count}} | Sort-Object Count -Descending | convertto-html -Fragment -PreContent "<h3>Event System</h3>" -PostContent "<p>Evenements sur les 30 derniers jours</p>"

#check event ID 4625
write-host ".......Check 3/4 (Check eventID 4625)"
$events4625 = Get-WinEvent -FilterHashtable @{logname="Security"; id=4625} -ErrorAction SilentlyContinue

if (($events4625).count -gt "0") {
$resultArray = @() # Creer un tableau vide pour stocker tous les resultats
foreach ($event in $events4625) {
    $eventXml = [xml]$event.ToXml()
    $eventArray = @{}
    $eventXml.Event.EventData.Data | Where-Object {$_.name -eq "TargetUserName" -or $_.name -eq "Status" -or $_.name -eq "WorkstationName" -or $_.name -eq "IpAddress" -or $_.name -eq "LogonType"} | ForEach-Object { $eventArray[$_.name] = $_.'#text' }
    $systemObject = $eventXml.Event.System | Select-Object EventID,@{N="TimeCreated";E={($_.TimeCreated).SystemTime}}

    $combinedObject = New-Object -TypeName PSObject
    foreach ($property in $eventArray.Keys) {
        $combinedObject | Add-Member -MemberType NoteProperty -Name $property -Value $eventArray[$property]
    }
    foreach ($property in $systemObject.psobject.Properties) {
        $combinedObject | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
    }
    $resultArray += $combinedObject # Ajouter l'objet combinÃ© au tableau des rÃ©sultats
}

$groupedResults = $resultArray | Group-Object TargetUserName
$event4625count = $groupedResults | Select Name,Count
$events4625=$event4625count | convertTo-html -Fragment -PreContent "<h3>Event 4625</h3>" -PostContent "<p>Details of the event can be found in C:\DFI-Maintenance\events </p>"
$Complete4625ReportHTML = $resultArray | convertto-html -Fragment -PreContent "<p>Logon Type 2: Interactive. A user logged on to this computer.</p><p>Logon type 3:  Network.  A user or computer logged on to this computer from the network.</p><p>Logon type 4: Batch.  Batch logon type is used by batch servers, where processes may be executing on behalf of a user without their direct intervention.</p><p>Logon type 5: Service.  A service was started by the Service Control Manager.<p>
"
$Report4625Title = "<h1>Evenements sur l'ID 4625</h1>"
$Complete4625Report = ConvertTo-HTML -Body "$Report4625Title $Complete4625ReportHTML" -Head $header
$Complete4625Report | Out-File $outfile4625
} 

if (($events4625).count -eq "0") {
$events4625 = "<h3>Event 4625</h3><p>No events are present during the requested period (30 days)</p>"
}

#Check Ressources serveur (id 2004)
write-host ".......Check 4/4 (Check eventID 2004)"
$event2004 = Get-WinEvent -FilterHashtable @{logname="System"; id=2004} -ErrorAction SilentlyContinue | select-object -unique | select TimeCreated,ID,LevelDisplayName,Message
$events2004 = $event2004 | convertTo-html -Fragment -PreContent "<h3>Event 2004</h3>"

if (($event2004).count -eq "0") {
$events2004 = "<h3>Event 2004</h3><p>No events are present regarding the server ressources</p>"
}

#Report Event ID Check
$ReportEventID= ConvertTo-HTML -Body "$ApplicationEvent $SystemEvent $events4625 $events4769 $events2004"
$ReportEventID | Add-Content $outfile

write-host "Check eventID OK"
#endregion EventID

#region Print Server
#Printer
if (($WindowsFeature).Name -eq "Print-Server") {
write-host "Check Printer"
$Printers=Get-Printer | where {$_.Name -notmatch "Microsoft"} | ConvertTo-Html Name,JobCount,Type,PortName,Shared,Published -Fragment -PreContent "<h2>Print Server</h2>"

#Report Print Server Check
$ReportPrinter = ConvertTo-HTML -Body "$Printers"
$ReportPrinter | Add-Content $outfile

write-host "Check Printer OK"
}
#endregion Print server

#region RDS
if (($WindowsFeature).Name -eq "RDS-Licensing") {
write-host "Check RDS"
$tsLicense = Get-CIMInstance -computername $server Win32_TSLicenseKeyPack -filter "TotalLicenses!=0" | ? { $_.TypeAndModel -ne 'Built-in TS Per Device CAL' }
$CALS=$tsLicense | select TypeAndModel, TotalLicenses, IssuedLicenses, AvailableLicenses, ProductVersion, ExpirationDate | convertto-html -Fragment -PreContent "<h2>Check RDS Licence</h2>"

#Report RDS Check
$ReportRDS = ConvertTo-HTML -Body "$CALS"
$ReportRDS | Add-Content $outfile

write-host "Check RDS OK"
}
#endregion RDS

#region Backup VBR
$VBRInstall = $Appslist | where { $_.Name -match "Veeam Backup & Replication" } -ErrorAction SilentlyContinue

if ((($VBRInstall).Name).count -gt "1") {
write-host "Check Backup VBR"
if (($VBRInstall).Version -match "10"){
Add-PSSnapin VeeamPSSnapin
}
Connect-VBRServer
write-host ".......Check 1/3 (VBR Repository)"
$VBRRepo=Get-VBRBackupRepository | Select Name,Path,CloudProvider,IsAvailable,VersionOfCreation | ConvertTo-Html -Fragment -PreContent "<h2>Check VBR repository</h2>"

write-host ".......Check 2/3 (VBR Job)"
$VBRJob = Get-VBRJob | Select Name,@{N="Last Result";E={$_.Info.LatestStatus}},JobType,TargetDir,TargetFile | ConvertTo-Html -Fragment -PreContent "<h2>Check VBR job</h2>"
$VBRJob = $VBRJob -replace '<td>success</td>','<td class="GreenColor">success</td>'
$VBRJob = $VBRJob -replace '<td>warning</td>','<td class="OrangeColor">warning</td>'
$VBRJob = $VBRJob -replace '<td>failed</td>','<td class="Redcolor">failed</td>'

write-host ".......Check 3/3 (VBR Licence)"
$VBRLicence=Get-VBRInstalledLicense | Select Status,Type,Edition,SupportID,SupportExpirationDate | ConvertTo-Html -Fragment -PreContent "<h2>Check VBR licence</h2>"
write-host "Check Backup VBR OK"

Disconnect-VBRServer

#Report VBR check
$ReportVBR=ConvertTo-HTML -Body "$VBRRepo $VBRJob $VBRLicence"
$ReportVBR | Add-Content $outfile

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


$VBOInstall= $Appslist | where { $_.Name -eq "Veeam Backup for Microsoft 365" } -ErrorAction SilentlyContinue
if ((($VBOInstall).name).count -eq "1") {
write-host "Check Backup VBO"
Import-VBOModule

write-host ".......Check 1/3 (VBO Repository)"
$VBORepo=Get-VBORepository | Select Name,Path,IsOutdated,
@{n = "Capacity (GB)"; e ={ [math]::Round($_.Capacity / 1GB, 2) } }, 
@{n = "FreeSpace (GB)"; e = { [math]::Round($_.FreeSpace / 1GB, 2) } }, 
@{n = "UsedSpace (GB)"; e = { [math]::Round((($_.Capacity - $_.FreeSpace) / 1GB), 2) } }, 
@{name = "PercentFree (GB)"; expression = { [math]::Round(($_.FreeSpace / $_.Capacity) * 100, 2) } } |
ConvertTo-Html -Fragment -PreContent "<h2>Check VBO365 repository</h2><h3>Repository local</h3>"

$VBORepoObject = Get-VBOObjectStorageRepository | Select Name,Type,SizeLimit,@{n = "UsedSpace (GB)"; e = { [math]::Round((($_.UsedSpace) / 1GB), 2) } },FreeSpace| ConvertTo-Html -Fragment -PreContent "<h3>Repository cloud</h3>"

write-host ".......Check 2/3 (VBO Job)"
$VBOJob=Get-VBOJob | Select Name,Repository,LastStatus,NextRun,IsEnabled | ConvertTo-Html -Fragment -PreContent "<h2>Check VBO365 job</h2>"

write-host ".......Check 3/3 (VBO Licence)"
$VBOLicence=get-vbolicense | select Status,ExpirationDate,TotalNumber | ConvertTo-Html -As List -Fragment -PreContent "<h2>Check VBO365 licence</h2>"

Disconnect-VBOServer

#Report VBO Check
$ReportVBO=ConvertTo-HTML -Body "$VBORepo $VBORepoObject $VBOJob $VBOLicence"
$ReportVBO | Add-Content $outfile

write-host "Check Backup VBO OK"
}

#endregion Backup VBO365

#region check Exchange
write-host "Check Exchange"
$ExchangeInstall = $Appslist | where { $_.Name -match "Exchange" } -ErrorAction SilentlyContinue
IF (($ExchangeInstall).count -gt "1"){
If ((Test-Path $healthreportPath) -eq $false) {
	New-Item "C:\DFI-Maintenance\HealthChecker" -itemType Directory | out-null
}
If ((Test-Path $ExchangeSizeReport) -eq $false) {
	New-Item "C:\DFI-Maintenance\ExchangeSizeReport" -itemType Directory | out-null
}

write-host ".......Check 1/3 (Lancement script Exchange Health)"
If ((Test-Path "c:\DFI-Maintenance\HealthChecker\HealthChecker.ps1") -eq $false) {
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1" -OutFile "c:\DFI-Maintenance\HealthChecker\HealthChecker.ps1"
} elseif ((Test-Path "c:\DFI-Maintenance\HealthChecker\HealthChecker.ps1") -eq $true) {
cd "c:\DFI-Maintenance\HealthChecker"
.\HealthChecker.ps1 -ScriptUpdateOnly
}

Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

cd "c:\DFI-Maintenance\HealthChecker"
.\HealthChecker.ps1 -SkipVersionCheck
.\HealthChecker.ps1 -BuildHtmlServersReport

write-host ".......Check 2/3 (Check Service Exchange)"
$EXCHService=Get-ServerComponentState -identity $hostname | select Component,State | ConvertTo-Html -Fragment -PreContent "<h2>Check Exchange Services</h2>"
$EXCHService = $EXCHService -replace '<td>Active</td>','<td class="SuccessWU">Active</td>'

write-host ".......Check 3/3 (Check SSL Exchange)"
$EXCHSSL= Get-ExchangeCertificate | where {$_.Services -match "iis" -and $_.friendlyname -notmatch "microsoft"} | ConvertTo-Html FriendlyName,Subject,Services,NotAfter,Status -Fragment -PreContent "<h2>Check SSL Exchange</h2>"
$EXCHSSL = $EXCHSSL -replace '<td>Valid</td>','<td class="SuccessWU">Valid</td>'

$htmlreportexchnage= get-content "C:\DFI-Maintenance\HealthChecker\ExchangeAllServersReport.html" -raw
$htmlreportexchnage = $htmlreportexchnage -replace "(?s)<style>.*?</style>", " "

if ((Test-Path "c:\DFI-Maintenance\HealthChecker\Test-ExchangeServerHealth.ps1") -eq $false){
Invoke-WebRequest -Uri "https://owncloud.dfinet.ch/index.php/s/8aISULiml9lR83S/download" -OutFile "c:\DFI-Maintenance\HealthChecker\Test-ExchangeServerHealth.ps1"
}
.\Test-ExchangeServerHealth.ps1 -reportmode

$htmlreportexchnageSoft = get-content "C:\DFI-Maintenance\HealthChecker\exchangeserverhealth.html" -raw
$htmlreportexchnageSoft = $htmlreportexchnageSoft -replace "(?s)<style>.*?</style>", " "


#Check de la taille des boites mails
$MailboxSize = Get-MailboxDatabase | Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | Select-Object DisplayName, TotalItemSize, ItemCount, @{Label="Database"; Expression={$_.Database.Name}}
$TopMailboxSize = $MailboxSize | select -first 10 | convertto-html -Property DisplayName, TotalItemSize, ItemCount, Database -PreContent "<h2>Mailbox size</h2>"


#Rapport complet des boites mails du client
$SMTPAdresses = Get-mailbox | where {$_.PrimarySmtpAddress -notmatch "Discovery"} | Select-Object PrimarySmtpAddress, @{Name="EmailAddresses"; Expression={$_.EmailAddresses | Where-Object {$_ -like "smtp:*" } | ForEach-Object {$_ -replace "^smtp:",""}}} | convertto-html -Property PrimarySmtpAddress,EmailAddresses -PreContent "<h2>Smtp adresses</h2>"
$CompleteMailboxSizeReport = $MailboxSize | convertto-html -Property DisplayName, TotalItemSize, ItemCount, Database -PreContent "<h2>Mailbox size</h2>"
$MailboxSizeTitle = "<h1>Check MailboxSize</h1>"
$SizeReport = ConvertTo-HTML -Body "$MailboxSizeTitle $CompleteMailboxSizeReport $SMTPAdresses" -Head $header
$SizeReport | Out-File "$($ExchangeSizeReport)\MailboxSize_$(get-date -format dd-MM-yy).html"

#Report Check Exchange
$ExchangeReport = ConvertTo-HTML -Body "$EXCHService $EXCHSSL $TopMailboxSize"
$ExchangeReport | Add-Content $outfile

write-host "Check Exchange OK"
}
#endregion check Exchange

#region WindowsUpdate
write-host "Check Windows Update"
write-host ".......Check 1/2 Windows Update : Update Available"
#MAJ en attente d'installation
$PackageNuget = Get-packageprovider -ListAvailable | where {$_.Name -eq "NuGet"}
$moduleWU = Get-Module -ListAvailable | where {$_.Name -eq "PSWindowsUpdate"}

If (($PackageNuget).version -lt "2.8.5.208" -or ($PackageNuget).count -eq "0"){
Install-PackageProvider -Name NuGet -Force -erroraction SilentlyContinue | out-null
}

if (($PackageNuget).count -gt "0") {
if (($moduleWU).count -eq "0") {
Install-Module PSWindowsUpdate -force -erroraction SilentlyContinue | out-null
}
}

import-module PSWindowsUpdate
update-module PSWindowsUpdate -force -ErrorAction silentlyContinue

if ((get-module).name -eq "PSWindowsUpdate"){
$UpdateToInstall= @()
$UpdateAvailable=Get-WindowsUpdate
$UpdateToInstall += $UpdateAvailable
$Update=$UpdateToInstall | Select Title,Size | ConvertTo-Html -Fragment -PreContent "<h2>Check Windows Update</h2><h3>Update Available</h3>"
}

if (($Update).count -eq "0") {
$Update = "<h2>Check Windows Update</h2><h3>Update Available</h3><p>No Available Update</p>"
}

write-host ".......Check 2/2 Windows Update : Update History"
$session = (New-Object -ComObject 'Microsoft.Update.Session')
$history = $session.QueryHistory("", 0, 50) | where {$_.Title -notmatch "Defender"}
$ReportWU=$history | Select Date, Title,
@{N="Result";E={$_.ResultCode -replace '1', 'Success' -replace '2', 'success' -replace '3', 'Success with Warning' -replace '4', 'Failed' -replace '5', 'Last failed install attempt' }} |
ConvertTo-Html -Fragment -PreContent "<h3>Latest Windows Updates Installed</h3>"
$ReportWU = $ReportWU -replace '<td>Success</td>','<td class="SuccessWU">Success</td>'
$ReportWU = $ReportWU -replace '<td>Failed</td>','<td class="FailedWU">Failed</td>'

$wuservice = get-service | where {$_.Name -eq "wuauserv"}
if (($wuservice).status -eq "Stopped"){
$ReportWU = "<h2>Latest Windows Updates Installed</h2><p>Windows Update Service is Stopped</p>"
}

#Report WU Check
$ReportWU = ConvertTo-HTML -Body "$Update $ReportWU" -PostContent "<p id='CreationDate'>Creation Date: $(Get-Date)</p>"
$ReportWU | Add-Content $outfile

write-host "Check Windows Update OK"
#endregion Windows Update

write-host "Report generated here : " $outfile

#region Report HTML other script
$htmlreportexchnageSoft | Add-Content $outfile
$htmlreportexchnage | Add-Content $outfile
#endregion Report HTML other script

####################################################################################

<# Note de version

Ajout du check de la version actuelle pour certaines applications

Création du rapport au fur et à mesure de l'avancement des checks

Activation du check AV defender

Activation du check des mises à jours en attente

Remplacement des texte FR -> EN

Prise en charge du check du repository cloud S3, ...

Modification de l'affichage de la licence VeeamOffice365 pour afficher les infos complètes de la licence

Ajout du check du dernier backup AD dans les Check AD

Ajout du check de l'event 2004 de log systeme

Suppression du check eventlog dans le test dcdiag

#>

####################################################################################
