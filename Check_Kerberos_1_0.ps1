
<#PSScriptInfo

.VERSION 1.0

.GUID 40505b8d-7fb8-4335-8087-fa0c7a20af14

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
 check kerberos 

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
$ScriptVersion = Test-ScriptFileInfo -Path ".\Check_Kerberos_1_0.ps1"
write-host "Version du script : $(($ScriptVersion).version.major).$(($ScriptVersion).version.minor)"
$reportPath = "c:\DFI-Maintenance\Kerberos"
$outfile = "$($reportPath)\Kerberos_$($env:computername)_$(get-date -format dd-MM-yy ).html"

#region test path
If ((Test-Path "C:\DFI-Maintenance") -eq $false) {
	New-Item "C:\DFI-Maintenance" -itemType Directory | out-null
}
If ((Test-Path $reportPath) -eq $false) {
	New-Item "C:\DFI-Maintenance\Kerberos" -itemType Directory | out-null
}
If (Test-Path $outfile) {
	Remove-Item $outfile
}
#endregion test path


$ReportTitle = "<h1>Check Kerberos</h1>"

Write-Host "Check compte Kerberos"
$KRBTGTAccount = Get-ADUser "krbtgt" -Property Created, PasswordLastSet
$KRBTGTPwdAge = $KRBTGTAccount | select Created,DistinguishedName,Enable,Name,PasswordLastSet | ConvertTo-Html -As List -Fragment -PreContent "<h2>Krbtgt account</h2>"

#check event
Write-Host "Check event 4769"
$events4769 = Get-WinEvent -FilterHashtable @{logname="Security"; id=4769; StartTime=(Get-Date).AddDays(-30)} -ErrorAction SilentlyContinue
$resultHashtable = @{}

foreach($event in $events4769){
    $eventProperties = [xml]$event.ToXml()
    $serviceName = $eventProperties.Event.EventData.Data | Where-Object {$_.Name -eq "ServiceName"} | Select-Object -ExpandProperty "InnerText"
    $ticketEncryptionType = $eventProperties.Event.EventData.Data | Where-Object {$_.Name -eq "TicketEncryptionType"} | Select-Object -ExpandProperty "InnerText"
        if(-not $resultHashtable.ContainsKey($serviceName)){
            $resultHashtable[$serviceName] = [pscustomobject]@{
                "Date" = $event.TimeCreated
                "Service name" = $serviceName
                "Ticket Encryption Type" = $ticketEncryptionType
        }
    }
}

$groupedResults4769 = $resultHashtable.Values | Group-Object "Ticket Encryption Type" | Select-Object Name, Count 

$events4769 = $groupedResults4769 | convertTo-html -Fragment -PreContent "<h2>Event 4769</h2>"
$Complete4769ReportHTML = $resultHashtable.Values | Sort-Object "Ticket Encryption Type" | convertTo-html -Fragment -PreContent "<h2>Event 4769 (all result)</h2>"
$0x17_4769ReportHTML = $resultHashtable.Values | where {$_."Ticket Encryption Type" -eq "0x17"}| convertTo-html -Fragment -PreContent "<h2>Event 4769 (0x17 only)</h2>"

#check AD Kerberos
Write-Host "Check encryption type user ad"
$UserKerberosEncryptionType = Get-ADUser -Filter * -Property * |
Select-Object Name,DistinguishedName,ObjectClass,@{Name='KerberosEncryptionType';Expression={$_.KerberosEncryptionType -join ','}},@{Name='msDS-SupportedEncryptionTypes';Expression={$_.{'msDS-SupportedEncryptionTypes'} -join ','}} |
Sort-Object msDS-SupportedEncryptionTypes -Descending |
ConvertTo-Html -Fragment -PreContent "<h2>Check KerberosEncryptionType (ADUser) </h2>"

#Creation du rapport
$Complete4769Report = ConvertTo-HTML -Body "$ReportTitle $KRBTGTPwdAge $events4769 $Complete4769ReportHTML $0x17_4769ReportHTML $UserKerberosEncryptionType" -Head $header
$Complete4769Report | Out-File $outfile

write-host "Rapport généré ici : " $outfile
read-host "appuyer sur enter pour fermer"