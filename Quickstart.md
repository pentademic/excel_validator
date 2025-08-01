# ⚡ Démarrage Rapide - Excel Validator Pro

## 🚀 Installation en 3 étapes

### 1️⃣ Créer les fichiers
Créez un dossier `excel_validator_pro` et copiez les 8 fichiers fournis :

```
excel_validator_pro/
├── rules_manager.py          # ✅ Copier artifact 1
├── excel_validator_core.py   # ✅ Copier artifact 2  
├── gradio_interface.py       # ✅ Copier artifact 3
├── app.py                    # ✅ Copier artifact 4
├── requirements.txt          # ✅ Copier artifact 5
├── README.md                 # ✅ Copier artifact 6
├── lancer_app.bat            # ✅ Copier script Windows
└── lancer_app.sh             # ✅ Copier script Linux/Mac
```

### 2️⃣ Installer les dépendances
```bash
cd excel_validator_pro
pip install -r requirements.txt
```

### 3️⃣ Lancer l'application
```bash
python app.py
```

**🌐 L'application s'ouvre automatiquement sur http://localhost:7860**

---

## 📋 Test rapide (5 minutes)

### Étape 1 : Créer une règle simple
1. Allez sur **"Créer des Règles"** → **"Règles Simples"**
2. Colonne : `A`
3. Type : `Ne pas être vide`
4. Message : `La colonne A est obligatoire`
5. Cliquez **"Créer la règle"**

### Étape 2 : Créer une règle conditionnelle
1. Allez sur **"Créer des Règles"** → **"Règles Conditionnelles"**
2. **Condition 1** : Colonne `B` égal à `VIP`
3. **Action** : Colonne `C` ne doit pas être vide
4. **Message** : `Les clients VIP doivent avoir un montant`
5. Cliquez **"Créer la règle conditionnelle"**

### Étape 3 : Tester avec un fichier Excel
1. Créez un fichier Excel simple :
   ```
   A    | B    | C
   -----+------+-----
   John | VIP  | 1000
        | VIP  |     
   Jane | STD  | 500
   ```
2. Allez sur **"Validation Excel"**
3. Glissez votre fichier Excel
4. Cliquez **"Lancer la validation"**

**Résultat attendu : 2 erreurs détectées**
- Ligne 2, Colonne A : "La colonne A est obligatoire"
- Ligne 2, Colonne C : "Les clients VIP doivent avoir un montant"

---

## 🔧 Dépannage express

### ❌ "Module not found"
```bash
pip install -r requirements.txt --upgrade
```

### ❌ "Port already in use"
Modifiez le port dans `app.py` :
```python
server_port=7861  # Changez 7860 en 7861
```

### ❌ "Permission denied" (Linux/Mac)
```bash
chmod +x lancer_app.sh
./lancer_app.sh
```

---

## 🎯 Fonctionnalités principales

### ✅ Règles simples (10 types)
- **NotBlank** : Ne pas être vide
- **Length** : Longueur min/max
- **Type** : Entier, décimal, booléen
- **Regex** : Expression régulière
- **Email** : Adresse email valide
- **Choice** : Valeurs dans une liste
- **Country** : Nom de pays
- **Date** : Format de date
- **Comparison** : Comparaisons (=, ≠, >, <, etc.)
- **Duplicate** : Détection doublons

### 🔗 Règles conditionnelles
- **"Si... alors..."** avec conditions multiples
- **Opérateurs ET/OU** pour combiner conditions
- **12 opérateurs** de comparaison
- **5 types d'actions** conditionnelles

### 📊 Gestion et validation
- **Import/Export** règles en JSON
- **Activation/désactivation** des règles
- **Rapports détaillés** avec fichiers Excel annotés
- **Export CSV** des erreurs

---

## 💡 Exemples d'usage

### 📋 Validation RH
```
- Colonne "Nom" : Ne pas être vide
- Colonne "Email" : Format email valide
- Si "Statut" = "CDI" alors "Salaire" doit être > 1500
```

### 💰 Validation Finance
```
- Colonne "Montant" : Type décimal
- Colonne "Devise" : Dans la liste [EUR, USD, GBP]
- Si "Montant" > 10000 alors "Validation" ne doit pas être vide
```

### 📦 Validation Inventaire
```
- Colonne "SKU" : Pattern regex "^[A-Z]{3}-\d{4}$"
- Colonne "Stock" : Entier positif
- Détection doublons sur colonne "SKU"
```

---

**🎉 Vous êtes prêt ! L'application est maintenant fonctionnelle et prête à valider vos fichiers Excel.**
