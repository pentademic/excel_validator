import json
import os
from datetime import datetime
from typing import Dict, List, Any, Optional
import pandas as pd

class RulesManager:
    """Gestionnaire de règles de validation avec support multicolonne"""
    
    def __init__(self, rules_file: str = "rules.json"):
        self.rules_file = rules_file
        self.rules = {
            "simple_rules": [],
            "conditional_rules": [],
            "multicolumn_rules": [],  # Nouvelle section pour les règles multicolonnes
            "metadata": {
                "created_at": datetime.now().isoformat(),
                "version": "1.1",  # Version mise à jour
                "last_modified": datetime.now().isoformat()
            }
        }
        self.load_rules()
    
    def add_multicolumn_rule(self, columns: List[str], rule_type: str, params: Dict[str, Any], 
                           message: str = "") -> Dict[str, Any]:
        """
        Ajoute une règle de validation multicolonne
        
        Args:
            columns: Liste des colonnes concernées (ex: ["A", "B", "C"])
            rule_type: Type de règle multicolonne
            params: Paramètres spécifiques à la règle
            message: Message d'erreur personnalisé
        
        Returns:
            La règle créée
        """
        rule_id = f"multi_{len(self.rules['multicolumn_rules']) + 1}_{int(datetime.now().timestamp())}"
        
        rule = {
            "id": rule_id,
            "columns": columns,
            "rule_type": rule_type,
            "params": params,
            "message": message or f"Erreur de validation multicolonne {rule_type}",
            "active": True,
            "created_at": datetime.now().isoformat()
        }
        
        self.rules["multicolumn_rules"].append(rule)
        self._update_metadata()
        return rule
    
    def get_multicolumn_rule_types(self) -> Dict[str, Dict]:
        """Retourne les types de règles multicolonnes disponibles avec leurs descriptions"""
        return {
            "sum_equals": {
                "name": "Somme égale",
                "description": "La somme des colonnes A et B doit égaler la colonne C",
                "params": ["target_column"],
                "example": "A + B = C"
            },
            "sum_range": {
                "name": "Somme dans une plage",
                "description": "La somme des colonnes doit être dans une plage donnée",
                "params": ["min_value", "max_value"],
                "example": "A + B + C entre 100 et 1000"
            },
            "date_before": {
                "name": "Date antérieure",
                "description": "La date de la colonne A doit être antérieure à celle de la colonne B",
                "params": [],
                "example": "Date_début < Date_fin"
            },
            "date_after": {
                "name": "Date postérieure",
                "description": "La date de la colonne A doit être postérieure à celle de la colonne B",
                "params": [],
                "example": "Date_fin > Date_début"
            },
            "date_range": {
                "name": "Écart de dates",
                "description": "L'écart entre deux dates doit respecter une plage",
                "params": ["min_days", "max_days"],
                "example": "Date_fin - Date_début entre 1 et 30 jours"
            },
            "percentage_of": {
                "name": "Pourcentage de",
                "description": "La colonne A doit représenter un pourcentage de la colonne B",
                "params": ["percentage", "tolerance"],
                "example": "A = 20% de B (±2%)"
            },
            "all_or_none": {
                "name": "Tout ou rien",
                "description": "Toutes les colonnes doivent être remplies ou toutes vides",
                "params": [],
                "example": "A, B, C toutes remplies OU toutes vides"
            },
            "unique_combination": {
                "name": "Combinaison unique",
                "description": "La combinaison des valeurs des colonnes doit être unique",
                "params": ["case_sensitive"],
                "example": "Combinaison A+B+C unique dans tout le fichier"
            },
            "conditional_sum": {
                "name": "Somme conditionnelle",
                "description": "Si la colonne condition a une valeur donnée, alors la somme des autres colonnes doit respecter une règle",
                "params": ["condition_column", "condition_value", "operator", "target_value"],
                "example": "Si D='VIP' alors A+B+C > 1000"
            },
            "max_min_check": {
                "name": "Maximum/Minimum",
                "description": "Une colonne doit contenir la valeur max/min parmi plusieurs colonnes",
                "params": ["operation", "target_column"],
                "example": "C = MAX(A, B) ou C = MIN(A, B)"
            }
        }
    
    def add_multi_simple_rule(self, columns: List[str], rule_type: str, params: Dict[str, Any], 
                         message: str = "") -> Dict[str, Any]:
        """
        Ajoute une règle simple appliquée à plusieurs colonnes
        
        Args:
            columns: Liste des colonnes concernées (ex: ["A", "B", "C"])
            rule_type: Type de règle simple (NotBlank, Length, Type, etc.)
            params: Paramètres de la règle
            message: Message d'erreur personnalisé
        
        Returns:
            La règle créée
        """
        rule_id = f"multi_simple_{len(self.rules.get('multi_simple_rules', [])) + 1}_{int(datetime.now().timestamp())}"
        
        rule = {
            "id": rule_id,
            "columns": columns,
            "rule_type": rule_type,
            "params": params,
            "message": message or f"Erreur de validation {rule_type} sur plusieurs colonnes",
            "active": True,
            "created_at": datetime.now().isoformat()
        }
        
        # Initialiser la section si elle n'existe pas
        if "multi_simple_rules" not in self.rules:
            self.rules["multi_simple_rules"] = []
        
        self.rules["multi_simple_rules"].append(rule)
        self._update_metadata()
        return rule

    def get_multi_simple_rule_types(self) -> Dict[str, Dict]:
        """Retourne les types de règles simples applicables à plusieurs colonnes"""
        return {
            "NotBlank": {
                "name": "Non vide (multi)",
                "description": "Toutes les colonnes spécifiées doivent être non vides",
                "params": [],
                "example": "A, B, C toutes non vides"
            },
            "Length": {
                "name": "Longueur (multi)",
                "description": "Toutes les colonnes doivent respecter les contraintes de longueur",
                "params": ["min", "max"],
                "example": "A, B, C entre 3 et 50 caractères"
            },
            "Type": {
                "name": "Type de données (multi)",
                "description": "Toutes les colonnes doivent être du même type",
                "params": ["type"],
                "example": "A, B, C tous des nombres"
            },
            "Regex": {
                "name": "Expression régulière (multi)",
                "description": "Toutes les colonnes doivent correspondre au pattern",
                "params": ["pattern"],
                "example": "A, B, C respectent le format XXX-999"
            },
            "Choice": {
                "name": "Choix dans liste (multi)",
                "description": "Toutes les colonnes doivent contenir des valeurs de la liste",
                "params": ["choices", "caseSensitive"],
                "example": "A, B, C dans ['Oui', 'Non', 'Peut-être']"
            },
            "Email": {
                "name": "Email valide (multi)",
                "description": "Toutes les colonnes doivent contenir des emails valides",
                "params": [],
                "example": "A, B, C des adresses emails valides"
            },
            "Date": {
                "name": "Date valide (multi)",
                "description": "Toutes les colonnes doivent contenir des dates valides",
                "params": ["format"],
                "example": "A, B, C des dates au format JJ/MM/AAAA"
            },
            "Comparison": {
                "name": "Comparaison (multi)",
                "description": "Toutes les colonnes doivent respecter la même condition",
                "params": ["operator", "value"],
                "example": "A, B, C toutes > 100"
            }
        }

    def add_simple_rule(self, column: str, rule_type: str, params: Dict[str, Any], 
                       message: str = "") -> Dict[str, Any]:
        """Ajoute une règle de validation simple"""
        rule_id = f"rule_{len(self.rules['simple_rules']) + 1}_{int(datetime.now().timestamp())}"
        
        rule = {
            "id": rule_id,
            "column": column,
            "rule_type": rule_type,
            "params": params,
            "message": message or f"Erreur de validation {rule_type}",
            "active": True,
            "created_at": datetime.now().isoformat()
        }
        
        self.rules["simple_rules"].append(rule)
        self._update_metadata()
        return rule
    
    def add_conditional_rule(self, conditions: List[Dict], actions: List[Dict], 
                           message: str = "", logic: str = "AND") -> Dict[str, Any]:
        """Ajoute une règle conditionnelle"""
        rule_id = f"cond_{len(self.rules['conditional_rules']) + 1}_{int(datetime.now().timestamp())}"
        
        rule = {
            "id": rule_id,
            "conditions": conditions,
            "actions": actions,
            "logic": logic,
            "message": message or "Règle conditionnelle non respectée",
            "active": True,
            "created_at": datetime.now().isoformat()
        }
        
        self.rules["conditional_rules"].append(rule)
        self._update_metadata()
        return rule
    
    def load_rules(self):
        """Charge les règles depuis le fichier"""
        if os.path.exists(self.rules_file):
            try:
                with open(self.rules_file, 'r', encoding='utf-8') as f:
                    loaded_rules = json.load(f)
                    
                # Mise à jour de la structure pour inclure les nouvelles règles
                if "multicolumn_rules" not in loaded_rules:
                    loaded_rules["multicolumn_rules"] = []
                if "multi_simple_rules" not in loaded_rules:  # NOUVEAU
                    loaded_rules["multi_simple_rules"] = []
                
                self.rules = loaded_rules
                
                # Mise à jour de la version si nécessaire
                if self.rules["metadata"].get("version", "1.0") < "1.2":
                    self.rules["metadata"]["version"] = "1.2"
                    self.save_rules()
                    
            except Exception as e:
                print(f"Erreur lors du chargement des règles: {e}")
                self._create_default_rules()
        else:
            self._create_default_rules()
    
    def save_rules(self):
        """Sauvegarde les règles dans le fichier"""
        try:
            with open(self.rules_file, 'w', encoding='utf-8') as f:
                json.dump(self.rules, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Erreur lors de la sauvegarde des règles: {e}")
    
    def get_rules_summary(self) -> pd.DataFrame:
        """Retourne un résumé de toutes les règles sous forme de DataFrame"""
        summary_data = []
        
        # Règles simples
        for rule in self.rules["simple_rules"]:
            summary_data.append([
                rule["id"],
                "Simple",
                rule["column"],
                f"{rule['rule_type']}",
                rule["message"],
                "✅ Actif" if rule["active"] else "❌ Inactif"
            ])
        
        # NOUVEAU: Règles simples multicolonnes
        for rule in self.rules.get("multi_simple_rules", []):
            columns_str = ", ".join(rule["columns"])
            rule_desc = f"{rule['rule_type']} (multi)"
            summary_data.append([
                rule["id"],
                "Simple Multi",
                columns_str,
                rule_desc,
                rule["message"],
                "✅ Actif" if rule["active"] else "❌ Inactif"
            ])
        
        # Règles conditionnelles
        for rule in self.rules["conditional_rules"]:
            conditions_str = f"{len(rule['conditions'])} condition(s)"
            actions_str = f"{len(rule['actions'])} action(s)"
            summary_data.append([
                rule["id"],
                "Conditionnelle",
                f"Conditions: {conditions_str}",
                f"{rule['logic']} → {actions_str}",
                rule["message"],
                "✅ Actif" if rule["active"] else "❌ Inactif"
            ])
        
        # Règles multicolonnes
        for rule in self.rules["multicolumn_rules"]:
            columns_str = ", ".join(rule["columns"])
            rule_desc = self.get_multicolumn_rule_types().get(rule["rule_type"], {}).get("name", rule["rule_type"])
            summary_data.append([
                rule["id"],
                "Multicolonne",
                columns_str,
                f"{rule_desc}",
                rule["message"],
                "✅ Actif" if rule["active"] else "❌ Inactif"
            ])
        
        if not summary_data:
            return pd.DataFrame(columns=["ID", "Type", "Colonne(s)", "Règle", "Message", "Status"])
        
        return pd.DataFrame(summary_data, columns=["ID", "Type", "Colonne(s)", "Règle", "Message", "Status"])

    
    def toggle_rule(self, rule_id: str, rule_type: str = None) -> bool:
        """Active ou désactive une règle"""
        rule_collections = [
            self.rules["simple_rules"],
            self.rules["conditional_rules"],
            self.rules["multicolumn_rules"],
            self.rules.get("multi_simple_rules", [])  # NOUVEAU
        ]
        
        for collection in rule_collections:
            for rule in collection:
                if rule["id"] == rule_id:
                    rule["active"] = not rule["active"]
                    self._update_metadata()
                    return True
        
        return False

    def delete_rule(self, rule_id: str, rule_type: str = None) -> bool:
        """Supprime une règle"""
        rule_collections = [
            ("simple_rules", self.rules["simple_rules"]),
            ("conditional_rules", self.rules["conditional_rules"]),
            ("multicolumn_rules", self.rules["multicolumn_rules"]),
            ("multi_simple_rules", self.rules.get("multi_simple_rules", []))  # NOUVEAU
        ]
        
        for collection_name, collection in rule_collections:
            for i, rule in enumerate(collection):
                if rule["id"] == rule_id:
                    del collection[i]
                    self._update_metadata()
                    return True
        
        return False
    
    def convert_to_yaml_config(self) -> Dict[str, Any]:
        """Convertit les règles au format de configuration YAML pour le validateur"""
        config = {
            "validators": {
                "columns": {},
                "default": []
            },
            "excludes": [],
            "header": True,
            "conditional_rules": [],
            "multicolumn_rules": [],
            "multi_simple_rules": []  # NOUVEAU
        }
        
        # Conversion des règles simples
        for rule in self.rules["simple_rules"]:
            if not rule["active"]:
                continue
                
            column = rule["column"]
            if column not in config["validators"]["columns"]:
                config["validators"]["columns"][column] = []
            
            rule_config = {
                rule["rule_type"]: {
                    **rule["params"],
                    "message": rule["message"]
                }
            }
            config["validators"]["columns"][column].append(rule_config)
        
        # NOUVEAU: Conversion des règles simples multicolonnes
        for rule in self.rules.get("multi_simple_rules", []):
            if rule["active"]:
                config["multi_simple_rules"].append({
                    "id": rule["id"],
                    "columns": rule["columns"],
                    "rule_type": rule["rule_type"],
                    "params": rule["params"],
                    "message": rule["message"]
                })
        
        # Conversion des règles conditionnelles
        for rule in self.rules["conditional_rules"]:
            if rule["active"]:
                config["conditional_rules"].append({
                    "conditions": rule["conditions"],
                    "actions": rule["actions"],
                    "logic": rule["logic"],
                    "message": rule["message"],
                    "active": True
                })
        
        # Conversion des règles multicolonnes
        for rule in self.rules["multicolumn_rules"]:
            if rule["active"]:
                config["multicolumn_rules"].append({
                    "id": rule["id"],
                    "columns": rule["columns"],
                    "rule_type": rule["rule_type"],
                    "params": rule["params"],
                    "message": rule["message"]
                })
        
        return config
    
    def export_rules(self, filename: str) -> bool:
        """Exporte les règles vers un fichier JSON"""
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(self.rules, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erreur lors de l'export: {e}")
            return False
    
    def import_rules(self, filename: str) -> bool:
        """Importe des règles depuis un fichier JSON"""
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                imported_rules = json.load(f)
            
            # Validation et fusion des règles
            if "simple_rules" in imported_rules:
                self.rules["simple_rules"].extend(imported_rules["simple_rules"])
            if "conditional_rules" in imported_rules:
                self.rules["conditional_rules"].extend(imported_rules["conditional_rules"])
            if "multicolumn_rules" in imported_rules:
                self.rules["multicolumn_rules"].extend(imported_rules["multicolumn_rules"])
            
            self._update_metadata()
            self.save_rules()
            return True
        except Exception as e:
            print(f"Erreur lors de l'import: {e}")
            return False
    
    def _create_default_rules(self):
        """Crée un fichier de règles par défaut"""
        self.rules = {
            "simple_rules": [],
            "multi_simple_rules": [],  # NOUVEAU
            "conditional_rules": [],
            "multicolumn_rules": [],
            "metadata": {
                "created_at": datetime.now().isoformat(),
                "version": "1.2",  # Version mise à jour
                "last_modified": datetime.now().isoformat()
            }
        }
        self.save_rules()
    
    def _update_metadata(self):
        """Met à jour les métadonnées"""
        self.rules["metadata"]["last_modified"] = datetime.now().isoformat()
    
    def get_statistics(self) -> Dict[str, Any]:
        """Retourne des statistiques sur les règles"""
        multi_simple_count = len(self.rules.get("multi_simple_rules", []))
        
        return {
            "total_rules": len(self.rules["simple_rules"]) + len(self.rules["conditional_rules"]) + len(self.rules["multicolumn_rules"]) + multi_simple_count,
            "simple_rules": len(self.rules["simple_rules"]),
            "multi_simple_rules": multi_simple_count,  # NOUVEAU
            "conditional_rules": len(self.rules["conditional_rules"]),
            "multicolumn_rules": len(self.rules["multicolumn_rules"]),
            "active_rules": sum([
                len([r for r in self.rules["simple_rules"] if r["active"]]),
                len([r for r in self.rules.get("multi_simple_rules", []) if r["active"]]),  # NOUVEAU
                len([r for r in self.rules["conditional_rules"] if r["active"]]),
                len([r for r in self.rules["multicolumn_rules"] if r["active"]])
            ]),
            "created_at": self.rules["metadata"]["created_at"],
            "last_modified": self.rules["metadata"]["last_modified"],
            "version": self.rules["metadata"]["version"]
        }
