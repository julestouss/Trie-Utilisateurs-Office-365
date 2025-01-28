if (-not (Get-Module -ListAvailable -Name AzureAD)){
    Install-Module -Name AzureAD -RequiredVersion 2.0.2.140 -Force -Scope CurrentUser
}
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)){
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
}
if (-not (Get-Module -ListAvailable -Name MSOnline)){
    Install-Module -Name MSOnline -RequiredVersion 1.1.183.66 -Force -Scope CurrentUser
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Beta)){
    Install-Module Microsoft.Graph.Beta -Force -Scope CurrentUser -AllowClobber
}

Import-Module ExchangeOnlineManagement

# récup info de connection 
$credentials = Get-Credential -Message "Entrez vos informations d'identification pour le tenant O365 voulu"


# Se connecter à Office 365
Connect-AzureAD -Credential $credentials # Pour le module AzureAD

# Se connecter a exchange Online 
Connect-ExchangeOnline -Credential $credentials -ShowProgress $true

Connect-MsolService -Credential $credentials



# Connect-MgGraph

#pour AzureAD
$users = Get-AzureADUser -All $true | Select-Object DisplayName, UserPrincipalName, AssignedLicence, @{Name='ProxyAddresses'; Expression={($_.ProxyAddresses -join ", ")}}, AccountEnabled, UserType

# Créer une liste pour stocker les résultats avec le type de boîte aux lettres
$result = @()


Write-Host "Récupération des informations des utilisateurs ... `n"
# Parcourir chaque utilisateur pour obtenir le type de boîte aux lettres
foreach ($user in $users) {
    # Récupérer la boîte aux lettres associée
    $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
    
    # Déterminer le type de boîte aux lettres
    if ($mailbox) {
        $mailboxType = if ($mailbox.RecipientTypeDetails -eq 'UserMailbox') { 'BAL Utilisateur' } 
                       elseif ($mailbox.RecipientTypeDetails -eq 'SharedMailbox') { 'BAL Partage' }
                       else { 'Autre' }
    } else {
        $mailboxType = 'Aucune boite aux lettres'
    }

    $Licence=Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Select-Object DisplayName, Licenses  
    $licenceList = $Licence.Licences | ForEach-Object {"$($Licence.Licenses.AccountSkuID)" }             
    
    # Ajouter les informations à la liste des résultats
    $result += [PSCustomObject]@{
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        ProxyAddresses = $user.ProxyAddresses
        # AccountEnabled = $user.AccountEnabled
        # UserType = $user.UserType
        MailboxType = $mailboxType
        licence=$licenceList -join "; "
    }
}
# Exporter les utilisateurs vers un fichier CSV
$result | Export-Csv -Path ".\csv\Office365_Users.csv" -NoTypeInformation -Encoding UTF8

# Message de confirmation
Write-Host "Les utilisateurs ont été exportés vers .\csv\Office365_Users.csv"

Write-Host "Procédé à l'installation de python si cela n'est pas déjà fait (! cochez la case pour l'intégrer dans le path) `n"

python -m ensurepip --upgrade

pip install pandas 
pip install openpyxl

python 'py/main.py'

