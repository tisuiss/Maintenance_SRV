#activation de la transcription
If ((Test-Path "C:\temp") -eq $false) {
	New-Item "C:\temp" -itemType Directory | out-null
}
Start-Transcript -Path "C:\temp\SSL_Install_Transcript.txt"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Check de la version d'Exchange
$ProductVersion = (Get-Command Exsetup.exe | ForEach {$_.FileVersionInfo}).ProductVersion
$VersionExchange = $ProductVersion.substring(0,5)

#creation de la forme
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SSL Installator'
$form.Size = New-Object System.Drawing.Size(300,400)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,320)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,320)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

#question 1
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Merci de renseigner le common name du certificat :'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$textBox.Text = ""
$form.Controls.Add($textBox)

$form.Topmost = $true

#question 2
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,65)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = 'Merci de renseigner le mot de passe du .pfx :'
$form.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,85)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
$textBox2.Text = ""
$form.Controls.Add($textBox2)

#question 3
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10,115)
$label3.Size = New-Object System.Drawing.Size(280,30)
$label3.Text = "Merci de renseigner le Nom Commun a donner au certificat SSL (information qui apparait dans l'ECP) :"
$form.Controls.Add($label3)

$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(10,145)
$textBox3.Size = New-Object System.Drawing.Size(260,20)
$textBox3.Text = ""
$form.Controls.Add($textBox3)

#question 4
$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(10,175)
$label4.Size = New-Object System.Drawing.Size(280,20)
$label4.Text = "Sélectionner un fichier .pfx"
$form.Controls.Add($label4)

#Création d'une zone texte (TextBox).
$textBox4 = New-Object Windows.Forms.TextBox
$textBox4.Location = New-Object Drawing.Point 10,195
$textBox4.Size = New-Object Drawing.Point 260,20
$textBox4.Text = ""
$form.Controls.Add($textBox4)

#Création d'un bouton parcourir (Button + OpenFileDialog).
$bouton1 = New-Object Windows.Forms.Button
$bouton1.Location = New-Object Drawing.Point(160,230)
$bouton1.Size = New-Object Drawing.Point(100,25)
$bouton1.text = "Parcourir"
$form.Controls.Add($bouton1)

$bouton1.add_click({
                     #Création d'un objet "ouverture de fichier".
                     $ouvrir1 = New-Object System.Windows.Forms.OpenFileDialog

                     #Initialisation du chemin par défaut.
                     $ouvrir1.initialDirectory = "C:\"

                     #Ici on va afficher que les fichiers en ".txt".
                     $ouvrir1.filter = "PFX Files (*.pfx)| *.pfx"

                     #Affiche la fenêtre d'ouverture de fichier.
                     $retour1 = $ouvrir1.ShowDialog()

                     #Traitement du retour.
                     #Si "OK" on affiche le fichier sélectionné dans la TextBox.
                     #Sinon on afficher un fichier par défaut.
                     if ($retour1 -eq "OK") { $textBox4.Text = $ouvrir1.filename }
                     else { $textBox4.Text = "" }
                  })




$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
if ([string]::IsNullOrWhitespace(($textBox.Text))) {
write-host "Comon Name manquant"
}
if ([string]::IsNullOrWhitespace(($textBox2.Text))) {
write-host "Mot de passe du .pfx manquant"
}
if ([string]::IsNullOrWhitespace(($textBox3.Text))) {
write-host "Friendly Name manquant"
}
if ([string]::IsNullOrWhitespace(($textBox4.Text))) {
write-host "Path du .pfx manquant"
}


#Transformation de l'URL en vue de l'installation du SSL
$Dollars = "$"
$Localhost = "\\localhost\"
$ReplacePath = ($textBox4.Text).Replace(":", $Dollars)
$PathSSL = $Localhost+$ReplacePath

#importation du SSL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

if(15.02 -eq $VersionExchange -or 15.01 -eq $VersionExchange) {
Import-ExchangeCertificate -FileData ([System.IO.File]::ReadAllBytes($PathSSL)) -Password (ConvertTo-SecureString -String $textBox2.Text -AsPlainText -Force) -ErrorAction Inquire  #pour exchange 2016 / 2019

}elseif (15.00 -eq $VersionExchange){
Import-ExchangeCertificate -FileName "$PathSSL" -Password $textBox2.Text -ErrorAction Inquire  # pour Exchange 2013
}

#Liste des SSL installes sur la machine
#Récupération des certificats
$lists = Get-ChildItem -Path Cert:LocalMachine\MY | Select FriendlyName, Subject, Services, NotAfter, thumbprint

#Création d'un objet ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.Columns.Add("Thumbprint", 100) | Out-Null
$listView.Columns.Add("Subject", 280) | Out-Null
$listView.Columns.Add("Not After", 80) | Out-Null
$listView.Size = New-Object System.Drawing.Size(465, 150)
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::None

# Calcul de la position pour centrer la ListBox
$positionX = ($form2.ClientSize.Width - $listView.Size.Width) / 2
$positionY = "20"
$listView.Location = New-Object System.Drawing.Point(20,20)

#Ajout des éléments de $Lists dans la ListView
foreach ($List in $Lists) {
    $item = New-Object System.Windows.Forms.ListViewItem(($List).Thumbprint)
    $item.SubItems.Add(($List).Subject)
    $item.SubItems.Add(($List).NotAfter.ToString("dd.MM.yyyy"))
    [void]$listView.Items.Add($item)
}

#Création de la boîte de dialogue
$form2 = New-Object System.Windows.Forms.Form
$form2.Text = "Sélectionnez le certificat à installer :"
$form2.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form2.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form2.ClientSize = New-Object System.Drawing.Size(600, 300)
$form2.Controls.Add($listView)

#Création des boutons "OK" et "Annuler"
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$okButton.Location = New-Object System.Drawing.Point(440, 220)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$form2.AcceptButton = $okButton
$form2.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Annuler"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cancelButton.Location = New-Object System.Drawing.Point(520, 220)
$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
$form2.CancelButton = $cancelButton
$form2.Controls.Add($cancelButton)

#Affichage de la boîte de dialogue
$result = $form2.ShowDialog()

#Gestion du résultat sélectionné de la liste
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $ThumbPrint=($listview.SelectedItems).text
}elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
    exit
}

#Assignation des services sur le SSL selectionne dans la liste
Enable-ExchangeCertificate -Thumbprint $ThumbPrint -Services IIS,SMTP -force -Confirm:$false

#renommer le friendly name du certificat
(Get-ChildItem -Path Cert:\LocalMachine\My\$($ThumbPrint)).FriendlyName = $textBox3.Text

Add-Type -AssemblyName System.Windows.Forms
$forme3 = New-Object System.Windows.Forms.Form
$forme3.Text = "OK"
$forme3.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$forme3.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$forme3.AutoSize = $false
$forme3.Size = New-Object System.Drawing.Size(220,120)

$label5 = New-Object System.Windows.Forms.Label
$label5.Text = "Le certificat a bien été installé"
$label5.AutoSize = $true
$label5.Location = New-Object System.Drawing.Point(20, 20)
$forme3.Controls.Add($label5)

$button3 = New-Object System.Windows.Forms.Button
$button3.Text = "OK"
$button3.Location = New-Object System.Drawing.Point(70,60)
$button3.DialogResult = [System.Windows.Forms.DialogResult]::OK
$forme3.AcceptButton = $button3
$forme3.Controls.Add($button3)

$forme3.ShowDialog() | Out-Null

#afficher les certificats installés sur le serveur pour validation visuelle
Get-ExchangeCertificate | where {$_.Status -eq "Valid"} | ft FriendlyName,Subject,Services,NotAfter,Status -autosize
}
Stop-Transcript

Read-Host -Prompt "Press Enter to exit"