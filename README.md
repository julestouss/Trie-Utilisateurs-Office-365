### **Tri des utilisateurs Office 365**

Ce projet utilise le code provenant de [Microsoft-365-License-Friendly-Names, @junecastillote](https://github.com/junecastillote/Microsoft-365-License-Friendly-Names/tree/master), disponible sous la licence MIT.

Ces scripts permettent de trier les utilisateurs Office 365 par nom de domaine de leur adresse e-mail principale.  
Ils permettent également d'obtenir les licences des utilisateurs et de convertir leur ID en "friendly name" pour une meilleure compréhension du tableau.

Le fichier Excel final se trouve dans `./excel/utilisateurs_o365.xlsx`.

### **Pré-requis**

Avant de commencer, assurez-vous d'avoir autorisé l'exécution de scripts :  

```powershell
Set-ExecutionPolicy RemoteSigned
# à exécuter en tant qu'administrateur
```

Assurez-vous également d'avoir installé Python et de l'avoir ajouté à la variable PATH de votre système.

Il faudra modifier les chemins d'accès dans les lignes **78, 88 et 94** du fichier `./py/main.py`.  
⚠️ **Les chemins relatifs ne fonctionnent pas, il est nécessaire d'utiliser des chemins complets.**

### **Utilisation**

Exécutez `./main.ps1` et renseignez les informations de connexion du compte administrateur du tenant souhaité.

### **TODO**

- Mettre en place les chemins relatifs dans le fichier Python.  
- Trouver un moyen d'accélérer le traitement des tenants avec un grand nombre d'utilisateurs.

