# 📚 Documentation - Excel Validator Pro

## 1. Introduction
**Excel Validator Pro** est un outil interactif permettant de valider des fichiers Excel selon des règles métier personnalisées. 
Grâce à son interface intuitive propulsée par **Gradio**, il permet de créer, gérer et exécuter différents types de validations 
sur une ou plusieurs colonnes simultanément.

---

## 2. Architecture du projet

### Structure des fichiers :
```
excel_validator_pro/
├── rules_manager.py          # Gestionnaire centralisé des règles
├── excel_validator_core.py   # Moteur principal de validation des fichiers Excel
├── gradio_interface.py       # Interface graphique utilisateur
├── app.py                    # Script principal de lancement
├── requirements.txt          # Dépendances Python
├── README.md                 # Guide rapide
├── lancer_app.bat            # Script Windows
├── lancer_app.sh             # Script Linux/Mac
```

### Composants principaux :
- **rules_manager.py** : Gestion et stockage des règles (simples, multicolonnes, conditionnelles)
- **excel_validator_core.py** : Logique de validation des fichiers Excel
- **gradio_interface.py** : Interface web pour créer et exécuter les règles
- **app.py** : Point d'entrée de l'application

---

## 3. Types de règles supportées

### 3.1. Règles simples (1 colonne)
- NotBlank
- Length
- Type
- Regex
- Email
- Choice
- Country
- Date
- Comparison
- Duplicate

### 3.2. Règles simples multicolonnes
Appliquer un même type de validation simple à plusieurs colonnes.

### 3.3. Règles conditionnelles
Validation basée sur des conditions logiques **ET / OU** appliquées à des colonnes, 
avec des actions à exécuter si les conditions sont remplies.

### 3.4. Règles multicolonnes
- Somme égale : `A + B = C`
- Somme dans une plage : min/max
- Comparaison de dates
- Écart de dates
- Pourcentage de
- Tout ou rien
- Combinaison unique
- Somme conditionnelle
- Max/Min

---

## 4. Fonctionnement interne

### 4.1. Chargement des règles
Les règles sont stockées dans `rules.json` et gérées par la classe **RulesManager**.

### 4.2. Lecture des fichiers Excel
- Utilisation de **openpyxl** pour lire les données
- Conversion des données en dictionnaire `{ligne: {colonne: valeur}}`
- Support des colonnes par lettre ou par nom d'en-tête

### 4.3. Application des règles
- Chaque type de règle possède une fonction de validation dédiée
- Les erreurs détectées sont stockées sous forme d'objets **ValidationError**
- Génération d'un fichier Excel annoté en sortie si des erreurs sont trouvées

---

## 5. Interface utilisateur

L'application utilise **Gradio** pour proposer une interface web simple avec plusieurs onglets :
1. **Accueil** : Présentation des fonctionnalités
2. **Créer des règles** : Interface pour ajouter des règles simples, multicolonnes et conditionnelles
3. **Gérer les règles** : Activer/désactiver ou supprimer des règles
4. **Validation Excel** : Importer un fichier Excel et lancer la validation

---

## 6. Installation

```bash
git clone <repo_url>
cd excel_validator_pro
pip install -r requirements.txt
python app.py
```

L'application est accessible sur [http://localhost:7860](http://localhost:7860).

---

## 7. Bonnes pratiques
- Sauvegarder les règles après chaque création
- Utiliser des noms d'en-têtes cohérents dans les fichiers Excel
- Tester les règles sur un petit échantillon avant d'appliquer à un fichier complet

---

## 8. Licence
Ce projet est distribué sous licence MIT. Vous êtes libre de l'utiliser, le modifier et le redistribuer.

