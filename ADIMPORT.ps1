# Import-ADUsersGUI.ps1
# Script PowerShell pour importer des utilisateurs dans Active Directory à partir d'un fichier CSV en utilisant une interface graphique.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Fonction pour afficher une boîte de message
function Show-Message {
    param (
        [string]$message,
        [string]$title = "Information"
    )
    [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# Fonction pour créer et afficher la fenêtre principale
function Create-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Importation d'utilisateurs Active Directory"
    $form.Size = New-Object System.Drawing.Size(400,250)
    $form.StartPosition = "CenterScreen"

    $labelDomain = New-Object System.Windows.Forms.Label
    $labelDomain.Text = "Entrez le domaine :"
    $labelDomain.AutoSize = $true
    $labelDomain.Location = New-Object System.Drawing.Point(10,20)
    $form.Controls.Add($labelDomain)

    $textBoxDomain = New-Object System.Windows.Forms.TextBox
    $textBoxDomain.Size = New-Object System.Drawing.Size(260,20)
    $textBoxDomain.Location = New-Object System.Drawing.Point(10,50)
    $form.Controls.Add($textBoxDomain)

    $labelCSV = New-Object System.Windows.Forms.Label
    $labelCSV.Text = "Sélectionnez le fichier CSV à importer :"
    $labelCSV.AutoSize = $true
    $labelCSV.Location = New-Object System.Drawing.Point(10,80)
    $form.Controls.Add($labelCSV)

    $textBoxCSV = New-Object System.Windows.Forms.TextBox
    $textBoxCSV.Size = New-Object System.Drawing.Size(260,20)
    $textBoxCSV.Location = New-Object System.Drawing.Point(10,110)
    $form.Controls.Add($textBoxCSV)

    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Text = "Parcourir..."
    $buttonBrowse.Location = New-Object System.Drawing.Point(280,110)
    $buttonBrowse.Add_Click({
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $OpenFileDialog.Title = "Sélectionnez le fichier CSV à importer"
        if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxCSV.Text = $OpenFileDialog.FileName
        }
    })
    $form.Controls.Add($buttonBrowse)

    $buttonImport = New-Object System.Windows.Forms.Button
    $buttonImport.Text = "Importer"
    $buttonImport.Location = New-Object System.Drawing.Point(150,150)
    $buttonImport.Add_Click({
        $csvFilePath = $textBoxCSV.Text
        $domain = $textBoxDomain.Text
        if (-not [string]::IsNullOrEmpty($csvFilePath) -and (Test-Path $csvFilePath) -and (-not [string]::IsNullOrEmpty($domain))) {
            try {
                # Importer le module Active Directory
                Import-Module ActiveDirectory

                # Lire le fichier CSV
                $userList = Import-Csv -Path $csvFilePath

                # Parcourir chaque ligne du fichier CSV
                foreach ($user in $userList) {
                    # Définir les propriétés de l'utilisateur
                    $firstName = $user.FirstName
                    $lastName = $user.LastName
                    $userName = $user.UserName
                    $password = $user.Password
                    $ou = $user.OU

                    # Créer le nom complet de l'utilisateur
                    $fullName = "$firstName $lastName"

                    # Définir les paramètres pour le compte utilisateur
                    $userParams = @{
                        SamAccountName = $userName
                        UserPrincipalName = "$userName@$domain"
                        Name = $fullName
                        GivenName = $firstName
                        Surname = $lastName
                        DisplayName = $fullName
                        Path = $ou
                        AccountPassword = (ConvertTo-SecureString $password -AsPlainText -Force)
                        Enabled = $true
                        ChangePasswordAtLogon = $false
                        PasswordNeverExpires = $true
                    }

                    # Créer l'utilisateur dans Active Directory
                    try {
                        New-ADUser @userParams
                        Show-Message -message "Utilisateur $fullName ($userName) créé avec succès."
                    } catch {
                        Show-Message -message "Erreur lors de la création de l'utilisateur $fullName ($userName) : $_" -title "Erreur"
                    }
                }

                Show-Message -message "Importation terminée."

            } catch {
                Show-Message -message "Une erreur s'est produite lors de l'importation : $_" -title "Erreur"
            }
        } else {
            Show-Message -message "Veuillez sélectionner un fichier CSV valide et entrer un domaine." -title "Erreur"
        }
    })
    $form.Controls.Add($buttonImport)

    $form.ShowDialog()
}

# Créer et afficher le formulaire principal
Create-MainForm