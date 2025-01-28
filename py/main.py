import pandas as pd
import os

# important ! pour le moment tout les path doivent être changé a la main pour chaque utilisateur --> les chemins relatifs ne marche pas
# il est prévu de pallier à ce problème au plus vite

def load_license_dictionary(file_path):
    """
    Charge un fichier CSV contenant les noms de licences dans un dictionnaire.
    :param file_path: chemin du fichier CSV
    :return: dictionnaire {SkuPartNumber: ChildServicePlanName}
    """
    df = pd.read_csv(file_path)
    
    # Créer un dictionnaire à partir du CSV : clé = SkuPartNumber, valeur = ChildServicePlanName
    license_dict = dict(zip(df['SkuPartNumber'], df['SkuName']))
    return license_dict


def replace_licenses_with_friendly(licenses_str, license_dict, delimiter=' '):
    """
    Remplace les licences par leurs noms amis à partir d'un dictionnaire.
    """
    # Vérifier si la chaîne est NaN
    if pd.isna(licenses_str):
        return ''  # Ou une autre valeur par défaut, si vous le souhaitez

    # Séparer les licences par le délimiteur (espace par défaut)
    licenses = licenses_str.split(delimiter)
    
    # Remplacer chaque licence par son nom ami si trouvé, sinon laisser la licence inchangée
    friendly_licenses = [license_dict.get(lic, lic) for lic in licenses]
    
    # Rejoindre les licences remplacées par le délimiteur
    return delimiter.join(friendly_licenses)


def process_csv(file_path, license_dict):
    """
    Charge et traite le fichier CSV des utilisateurs.
    Remplace les licences par leurs noms amis et génère un fichier Excel avec les utilisateurs triés par domaine.
    """
    # Charger le fichier CSV des utilisateurs
    df = pd.read_csv(file_path)

    # Vérification du contenu de la colonne 'licence' avant transformation
    print("Contenu de la colonne 'licence' avant transformation :")
    print(df['licence'].head())

    # Supprimer les caractères inutiles pour l'utilisateur
    df['licence'] = df['licence'].str.replace("reseller-account:", "", regex=False)
    df['licence'] = df['licence'].str.replace("License: ", "", regex=False)
    df['ProxyAddresses'] = df['ProxyAddresses'].str.replace("SMTP:", "", regex=False)
    df['ProxyAddresses'] = df['ProxyAddresses'].str.replace("smtp:", "", regex=False) 
    

    # Remplacer les licences par les noms amis
    df['licence'] = df['licence'].apply(lambda x: replace_licenses_with_friendly(x, license_dict))


    # Extraire le nom de domaine de 'UserPrincipalName'
    df['Domaine'] = df['UserPrincipalName'].str.split('@').str[1]

    # Trier le DataFrame par Domaine et par DisplayName
    df_sorted = df.sort_values(by=['Domaine', 'DisplayName'])

    #donner des noms facile au colonnes 
    df_sorted.rename(columns={
        'UserPrincipalName': 'Email',
        'licence': 'Licences utilisateur',
        'DisplayName': "Nom d'affichage",
        'ProxyAddresses': "Alias du mail",
        'MailboxType':  'Type de boite au lettres'
    }, inplace=True)

    # Créer un dictionnaire pour stocker les DataFrames groupés par domaine
    grouped = {domaine: group for domaine, group in df_sorted.groupby('Domaine')}

    # Exporter chaque groupe dans une feuille séparée dans un fichier Excel
    excel_file = r"\trie_user_office\excel\utilisateurs_o365.xlsx"
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        for domaine, group in grouped.items():
            # Utiliser le domaine comme nom de feuille, limité à 31 caractères
            group.to_excel(writer, sheet_name=domaine[:31], index=False)

    print(f"Les utilisateurs ont été classés par domaine et enregistrés dans '{excel_file}'")


# Chemin vers le fichier CSV des noms amis de licences généré par le script PowerShell cf https://github.com/junecastillote/Microsoft-365-License-Friendly-Names/blob/master/Get-m365ProductIDTable.ps1
license_dict_path = r"\trie_user_office\csv\m365ProductIDTable.csv"

# Charger les noms amis de licences depuis le CSV
license_dict = load_license_dictionary(license_dict_path)

# Chemin vers le fichier CSV des utilisateurs
csv_file_path = r"\trie_user_office\csv\Office365_Users.csv"

# Traiter le fichier CSV des utilisateurs avec le dictionnaire de licences
process_csv(csv_file_path, license_dict)