    #############################################################################
    ###  Script pour exporter la liste des dossiers des boites mails sur EXO
    ###  
    ###
    ###
    ###
    ###
    #############################################################################


# Variables
$Listealias = (Get-mailbox | Where-Object { $_.Alias -notmatch "Discovery"}).alias
$TestPathFolder = Test-Path -Path "C:\Temp\$PathFolderName"


#region Boite de dialogue selection du dossier de sortie
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$title = "Sélection du dossier de sortie"
$msg = "Merci d'entrer le nom du dossier de sortie ou seront stockés les logs dans c:\temp :"
$FolderName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
if ($FolderName -eq [System.Windows.Forms.DialogResult]::OK) {
    write-Host $FolderName
 }elseif ($FolderName -eq [System.Windows.Forms.DialogResult]::Cancel) {
    write-host $FolderName
}
#endregion Boite de dialogue selection du dossier de sortie

if ($FolderName -eq "") {
    exit
}


#region Execution de l'export


# Test de l'existance du dossier, création si non existant
if ((Test-Path -Path "C:\Temp\$FolderName") -ne $True) {
    New-Item "C:\Temp\$FolderName" -itemType Directory | Out-Null
}

#Réalisation de l'export par mailbox et génération du fichier de sortie
ForEach ($User in $ListeAlias) {
    Write-host "Traitement de la mailbox " -NoNewline
    Write-Host $user -ForegroundColor Green
    $PathFolder = "C:\Temp\$FolderName\$($user)_FolderSize_EXCHANGE.txt" # ou $PathFolder = "C:\Temp\$FolderName\"+$user+"_FolderSize_EXCHANGE.txt"   
    $ResultFolder = Get-MailboxFolderStatistics -Identity $user | Format-Table Identity, FolderAndSubfolderSize -AutoSize
    $ResultFolder> $PathFolder
}
#endregion Execution de l'export