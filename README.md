### **Trie des utilisateurs office 365**

Ce projet utilise le code provenant de [Microsoft-365-License-Friendly-Names, @junecastillote](https://github.com/junecastillote/Microsoft-365-License-Friendly-Names/tree/master) disponible sous la licence MIT.


Ces scripts permettent de trier les utilisateurs office par nom de domaine de leur adress mail principale,
Il permet aussi d'obtenir les licences des utilisateurs et convertie leur ID en "friendly name" pour une meilleur compréhension du tableau.

Le fichier excel finale se trouve dans "./excel/utilisateurs_o365.xlsx"


Avant de commencer l'utilisation soyez sure d'avoir autorisé l'execution de script :
```powershell
Set-ExecutionPolicy RemoteSigned
# en adminstrateur
```
Soyez aussi sur d'avoir installé python et de l'avoir ajouté dans la variable PATH de votre systeme.

Il faudra aussi changer les chemin d'accès dans les lignes 78, 88 et 94 du fichier "./py/main.py"
! les chemins relatif ne marchent pas il faut mettre les chemins complets 

Il vous suffira ensuite de lancer ./main.ps1 et de rentrer les informations de connection au compte admin du tenant voulue

# **TODO**
TODO : Mettre en place les chemins relatifs dans le fichier python
TODO : trouver un moyen de traiter plus rapidement les tenants avec un grand nombre d'utilisateurs