⚡ Démarrage Rapide – Excel Validator Pro
🚀 Installation en 3 étapes
1️⃣ Créer les fichiers

Créez un dossier excel_validator_pro et copiez les fichiers fournis :

excel_validator_pro/
├── rules_manager.py          # Gestion des règles (simples, conditionnelles, multicolonnes)
├── excel_validator_core.py   # Moteur de validation Excel
├── gradio_interface.py       # Interface utilisateur Gradio
├── app.py                    # Script principal
├── requirements.txt          # Dépendances Python
├── README.md                 # Documentation
├── lancer_app.bat            # Script Windows
└── lancer_app.sh             # Script Linux/Mac

2️⃣ Installer les dépendances
cd excel_validator_pro
pip install -r requirements.txt

3️⃣ Lancer l'application
python app.py


🌐 L'application s'ouvre sur http://localhost:7860

📋 Test rapide
1. Créer une règle simple

Onglet "Créer des Règles" → "Règles Simples" → "Une colonne"

Colonne : A

Type : Ne pas être vide

Message : La colonne A est obligatoire

Créer la règle

2. Créer une règle conditionnelle

Onglet "Créer des Règles" → "Règles Conditionnelles"

Condition : Colonne B = VIP

Action : Colonne C ne doit pas être vide

Message : Les clients VIP doivent avoir un montant

Créer la règle conditionnelle

3. Créer une règle multicolonne

Onglet "Créer des Règles" → "Règles Multicolonnes"

Colonnes : A,B,C

Type : Somme égale (A + B = C)

Tolérance : 0.01

Message : La somme est incorrecte

Créer la règle multicolonne

4. Tester avec un fichier Excel

Fichier exemple :

A    | B    | C
-----+------+-----
John | VIP  | 1000
     | VIP  |     
Jane | STD  | 500


Validation :

Ligne 2, Colonne A → "La colonne A est obligatoire"

Ligne 2, Colonne C → "Les clients VIP doivent avoir un montant"

🎯 Fonctionnalités principales
📝 Règles simples (1 colonne)

NotBlank, Length, Type, Regex, Email, Choice, Country, Date, Comparison, Duplicate

📊 Règles simples multicolonnes

Appliquer un même type de règle simple à plusieurs colonnes à la fois.

🔗 Règles conditionnelles

Combiner plusieurs conditions (ET / OU) et exécuter des actions si elles sont remplies.

Compatible avec noms de colonnes ou lettres de colonnes (A, B…).

🔢 Règles multicolonnes

Somme égale : A + B = C

Somme dans une plage : min/max

Comparaison de dates : < ou >

Écart de dates : min/max jours

Pourcentage de : A = x% de B (± tolérance)

Tout ou rien : toutes vides ou toutes remplies

Combinaison unique : valeurs uniques sur un ensemble de colonnes

Somme conditionnelle : somme > / < / = si condition remplie

Max/Min : vérifie si une colonne contient le max/min des autres

📂 Gestion des règles

Activation/désactivation

Suppression

Sauvegarde dans rules.json

Rechargement automatique au démarrage

📑 Résultats

Rapport détaillé dans un fichier Excel annoté

Export CSV des erreurs

💡 Bonnes pratiques

Sauvegardez vos règles après création pour qu’elles soient appliquées lors de la validation.

Si vos conditions utilisent des noms de colonnes, assurez-vous que la première ligne de votre Excel est l’en-tête.
