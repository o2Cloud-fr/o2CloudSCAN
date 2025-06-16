# Scanner Pro 📱

Une application PowerShell moderne avec interface graphique pour la numérisation de documents, offrant une expérience utilisateur intuitive et des fonctionnalités avancées.

## ✨ Fonctionnalités

- **Interface graphique moderne** : Design sombre et élégant avec une interface utilisateur fluide
- **Numérisation multi-format** : Support des formats PNG, JPG, PDF, TIFF et BMP
- **Gestion automatique des scanners** : Détection et sélection automatique des scanners disponibles
- **Historique des scans** : Suivi complet de tous vos documents numérisés
- **Organisation intelligente** : Génération automatique de noms de fichiers avec horodatage
- **Gestion des dossiers** : Création automatique des dossiers de destination
- **Conversion PDF** : Conversion automatique d'images vers PDF via Microsoft Word

## 🚀 Installation

### Prérequis
- Windows 10/11
- PowerShell 5.1 ou supérieur
- Un scanner compatible WIA (Windows Image Acquisition)
- Microsoft Word (optionnel, pour la conversion PDF)

### Installation rapide

1. **Cloner le repository**
```bash
git clone https://github.com/o2Cloud-fr/scanner-pro.git
cd scanner-pro
```

2. **Exécuter l'application**
```powershell
# Méthode 1 : Exécution directe
PowerShell -ExecutionPolicy Bypass -File "ScannerPro.ps1"

# Méthode 2 : Via PowerShell ISE
# Ouvrir PowerShell ISE et charger le fichier
```

3. **Première utilisation**
   - L'application créera automatiquement les dossiers nécessaires
   - Dossier de configuration : `%USERPROFILE%\Documents\ScannerApp`
   - Dossier de scans par défaut : `%USERPROFILE%\Documents\Scans`

## 📝 Licence

[MIT License](https://opensource.org/licenses/MIT)

## 📋 Utilisation

### Interface principale

L'application se compose de plusieurs sections :

#### 🔧 Configuration du Scan
- **Nom du fichier** : Nom personnalisé (génération automatique avec horodatage)
- **Format** : Sélection du format de sortie (PNG, JPG, PDF, TIFF, BMP)
- **Destination** : Chemin de sauvegarde personnalisable
- **Scanner** : Sélection du scanner à utiliser

#### 📚 Historique des Scans
- Liste chronologique de tous les scans effectués
- Informations détaillées (nom, date, taille, scanner utilisé)
- Actions rapides : ouvrir, actualiser, supprimer

#### ℹ️ Zone d'informations
- Statut en temps réel des opérations
- Messages d'aide et conseils d'utilisation
- Rapport détaillé des scans effectués

### Workflow de numérisation

1. **Préparer le document** sur votre scanner
2. **Configurer** le nom de fichier et le format souhaité
3. **Sélectionner** le scanner (si plusieurs disponibles)
4. **Cliquer** sur "DÉMARRER LA NUMÉRISATION"
5. **Suivre** les instructions de votre interface de scanner
6. **Vérifier** le résultat dans l'historique

## 🛠️ Fonctionnalités techniques

### Gestion des scanners
- Détection automatique via WIA (Windows Image Acquisition)
- Support des scanners réseau et USB
- Fallback vers l'application Windows Scanner

### Formats supportés
- **PNG** : Qualité optimale pour documents
- **JPG** : Compression pour photos
- **PDF** : Format universel (conversion via Word)
- **TIFF** : Format professionnel
- **BMP** : Format non compressé

### Historique persistant
- Sauvegarde JSON des métadonnées
- Limitation à 100 entrées récentes
- Informations stockées : nom, chemin, date, scanner, taille

## 🎨 Personnalisation

### Thème sombre
L'application utilise un thème sombre moderne avec :
- Couleurs principales : Noir (#0F0F0F), Gris foncé (#191919)
- Accents : Bleu (#0096FF), Vert (#00C864)
- Police : Segoe UI pour une lisibilité optimale

### Configuration
Les paramètres sont automatiquement sauvegardés dans :
```
%USERPROFILE%\Documents\ScannerApp\scan_history.json
```

## 🔧 Dépannage

### Problèmes courants

#### Scanner non détecté
```powershell
# Vérifier les pilotes WIA
Get-WmiObject -Class Win32_PnPSignedDriver | Where-Object {$_.DeviceName -like "*scan*"}
```

#### Erreur de permissions
```powershell
# Exécuter en tant qu'administrateur
Start-Process PowerShell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File ScannerPro.ps1"
```

#### Conversion PDF échoue
- Vérifier l'installation de Microsoft Word
- S'assurer que Word peut s'exécuter en mode automatisé

### Logs et diagnostic
Les erreurs sont affichées dans la zone d'informations de l'application.

### Standards de code
- Respecter la syntaxe PowerShell standard
- Commenter les fonctions complexes
- Tester sur Windows 10 et 11

## 📋 TODO / Roadmap

- [ ] Support des scans batch (multiple pages)
- [ ] OCR intégré pour reconnaissance de texte
- [ ] Export vers cloud (OneDrive, Google Drive)
- [ ] Paramètres de qualité avancés
- [ ] Mode ligne de commande
- [ ] Notifications système
- [ ] Prévisualisation avant sauvegarde
- [ ] Templates de noms de fichiers personnalisés

## 👤 Auteur

**Votre Nom**
- GitHub: [@o2Cloud-fr](https://github.com/o2Cloud-fr)
- Email: github@o2cloud.fr

## 🔗 Liens

[![linkedin](https://img.shields.io/badge/linkedin-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/remi-simier-2b30142a1/)
[![github](https://img.shields.io/badge/github-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/o2Cloud-fr/)


## 🙏 Remerciements

- Microsoft pour les APIs WIA et Windows Forms
- Communauté PowerShell pour les exemples et bonnes pratiques
- Utilisateurs testeurs pour leurs retours précieux

---

⭐ **N'hésitez pas à donner une étoile si ce projet vous aide !**