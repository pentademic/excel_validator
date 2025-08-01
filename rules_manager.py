import json
import os
from typing import Dict, List, Any, Optional
from datetime import datetime

class RulesManager:
    """Gestionnaire pour créer, sauvegarder et charger les règles de validation"""
    
    def __init__(self, rules_file: str = "validation_rules.json"):
        self.rules_file = rules_file
        self.rules = self._load_rules()
    
    def _load_rules(self) -> Dict[str, Any]:
        """Charge les règles depuis le fichier JSON"""
        if os.path.exists(self.rules_file):
            try:
                with open(self.rules_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        return {
            "simple_rules": [],
            "conditional_rules": [],
            "metadata": {
                "created_at": datetime.now().isoformat(),
                "version": "1.0"
            }
        }
    
    def save_rules(self) -> bool:
        """Sauvegarde les règles dans le fichier JSON"""
        try:
            self.rules["metadata"]["last_modified"] = datetime.now().isoformat()
            with open(self.rules_file, 'w', encoding='utf-8') as f:
                json.dump(self.rules, f, indent=2, ensure_ascii=False)
            return True
        except IOError:
            return False
    
    def add_simple_rule(self, column: str, rule_type: str, params: Dict[str, Any], 
                       message: str, active: bool = True) -> Dict[str, Any]:
        """Ajoute une règle simple"""
        rule = {
            "id": f"rule_{len(self.rules['simple_rules']) + 1}_{int(datetime.now().timestamp())}",
            "column": column,
            "rule_type": rule_type,
            "params": params,
            "message": message,
            "active": active,
            "created_at": datetime.now().isoformat()
        }
        self.rules["simple_rules"].append(rule)
        return rule
    
    def add_conditional_rule(self, conditions: List[Dict], actions: List[Dict], 
                           message: str, logic: str = "AND", active: bool = True) -> Dict[str, Any]:
        """Ajoute une règle conditionnelle"""
        rule = {
            "id": f"cond_rule_{len(self.rules['conditional_rules']) + 1}_{int(datetime.now().timestamp())}",
            "conditions": conditions,
            "actions": actions,
            "logic": logic,
            "message": message,
            "active": active,
            "created_at": datetime.now().isoformat()
        }
        self.rules["conditional_rules"].append(rule)
        return rule
    
    def delete_rule(self, rule_id: str, rule_type: str = "simple") -> bool:
        """Supprime une règle par son ID"""
        rules_list = self.rules["simple_rules"] if rule_type == "simple" else self.rules["conditional_rules"]
        for i, rule in enumerate(rules_list):
            if rule["id"] == rule_id:
                del rules_list[i]
                return True
        return False
    
    def toggle_rule(self, rule_id: str, rule_type: str = "simple") -> bool:
        """Active/désactive une règle"""
        rules_list = self.rules["simple_rules"] if rule_type == "simple" else self.rules["conditional_rules"]
        for rule in rules_list:
            if rule["id"] == rule_id:
                rule["active"] = not rule["active"]
                return True
        return False
    
    def get_rules_summary(self) -> List[List[str]]:
        """Retourne un résumé des règles pour l'affichage en tableau"""
        summary = []
        
        # Règles simples
        for rule in self.rules["simple_rules"]:
            summary.append([
                rule["id"],
                "Simple",
                rule["column"],
                rule["rule_type"],
                rule["message"][:50] + "..." if len(rule["message"]) > 50 else rule["message"],
                "✅" if rule["active"] else "❌"
            ])
        
        # Règles conditionnelles
        for rule in self.rules["conditional_rules"]:
            conditions_str = f"{len(rule['conditions'])} condition(s)"
            actions_str = f"{len(rule['actions'])} action(s)"
            summary.append([
                rule["id"],
                "Conditionnelle",
                f"{conditions_str} → {actions_str}",
                "Conditional",
                rule["message"][:50] + "..." if len(rule["message"]) > 50 else rule["message"],
                "✅" if rule["active"] else "❌"
            ])
        
        return summary
    
    def export_rules(self, filepath: str) -> bool:
        """Exporte les règles vers un fichier JSON"""
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.rules, f, indent=2, ensure_ascii=False)
            return True
        except IOError:
            return False
    
    def import_rules(self, filepath: str) -> bool:
        """Importe les règles depuis un fichier JSON"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                imported_rules = json.load(f)
                if "simple_rules" in imported_rules and "conditional_rules" in imported_rules:
                    self.rules = imported_rules
                    return True
        except (json.JSONDecodeError, IOError):
            pass
        return False
    
    def convert_to_yaml_config(self) -> Dict[str, Any]:
        """Convertit les règles JSON en format YAML compatible"""
        yaml_config = {
            "validators": {
                "columns": {},
                "default": []
            },
            "excludes": [],
            "header": True,
            "conditional_rules": [r for r in self.rules["conditional_rules"] if r["active"]]
        }
        
        # Traitement des règles simples
        for rule in self.rules["simple_rules"]:
            if not rule["active"]:
                continue
                
            column = rule["column"]
            if column not in yaml_config["validators"]["columns"]:
                yaml_config["validators"]["columns"][column] = []
            
            # Conversion selon le type de règle
            rule_config = {rule["rule_type"]: rule["params"].copy()}
            if rule["message"]:
                rule_config[rule["rule_type"]]["message"] = rule["message"]
            
            yaml_config["validators"]["columns"][column].append(rule_config)
        
        return yaml_config