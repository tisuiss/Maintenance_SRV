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

$reportPath = "c:\temp\"
$outfileMailboxSize = "$($reportPath)\MailboxSize_$(get-date -format dd-MM-yy).html"

If ((Test-Path $reportPath) -eq $false) {
	New-Item "C:\temp" -itemType Directory | out-null
}

Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$MailboxSize = Get-MailboxDatabase | Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | Select-Object DisplayName, TotalItemSize, ItemCount, @{Label="Database"; Expression={$_.Database.Name}}
$SMTPAdresses = Get-mailbox | where {$_.PrimarySmtpAddress -notmatch "Discovery"} | Select-Object PrimarySmtpAddress, @{Name="EmailAddresses"; Expression={$_.EmailAddresses | Where-Object {$_ -like "smtp:*" } | ForEach-Object {$_ -replace "^smtp:",""}}} | convertto-html -Property PrimarySmtpAddress,EmailAddresses -PreContent "<h2>Smtp adresses</h2>"

$CompleteMailboxSizeReport = $MailboxSize | convertto-html -Property DisplayName, TotalItemSize, ItemCount, Database -PreContent "<h2>Mailbox size</h2>"
$MailboxSizeTitle = "<h1>Check MailboxSize</h1>"
$SizeReport = ConvertTo-HTML -Body "$MailboxSizeTitle $CompleteMailboxSizeReport $SMTPAdresses" -Head $header
$SizeReport | Out-File $outfileMailboxSize

write-host "Le fichier a été généré ici : $($outfileMailboxSize)"
Read-host "Appuyer sur enter pour fermer la session"

