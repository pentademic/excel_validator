import os
import tempfile
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from typing import Dict, List, Tuple, Any, Optional
import re
from datetime import datetime
from validate_email import validate_email
import pycountry

class ValidationError:
    """Classe pour représenter une erreur de validation"""
    def __init__(self, row: int, column: str, message: str, value: Any = None):
        self.row = row
        self.column = column
        self.message = message
        self.value = value
        self.coordinate = f"{column}{row}"

class ExcelValidatorCore:
    """Validateur Excel principal"""
    
    def __init__(self):
        self.errors: List[ValidationError] = []
        self.worksheet_data = {}  # Cache des données pour les règles conditionnelles
    
    def validate_file(self, file_path: str, rules_config: Dict[str, Any], 
                     sheet_name: str = None) -> Tuple[bool, List[ValidationError], str]:
        """Valide un fichier Excel selon les règles fournies"""
        self.errors = []
        self.worksheet_data = {}
        
        try:
            wb = load_workbook(file_path, data_only=True, read_only=True)
            
            if sheet_name and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.active
            
            # Charger toutes les données en mémoire pour les règles conditionnelles
            self._load_worksheet_data(ws)
            
            validators = rules_config.get("validators", {}).get("columns", {})
            default_validator = rules_config.get("validators", {}).get("default")
            excludes = rules_config.get("excludes", [])
            header_row = rules_config.get("header", True)
            conditional_rules = rules_config.get("conditional_rules", [])
            
            # Validation des règles simples
            self._validate_worksheet(ws, validators, default_validator, excludes, header_row)
            
            # Validation des règles conditionnelles
            self._validate_conditional_rules(conditional_rules)
            
            error_file_path = None
            if self.errors:
                error_file_path = self._generate_error_file(file_path, ws, self.errors)
            
            wb.close()
            return len(self.errors) == 0, self.errors, error_file_path
            
        except Exception as e:
            error = ValidationError(0, "A", f"Erreur lors de la lecture du fichier: {str(e)}")
            return False, [error], None
    
    def _load_worksheet_data(self, ws):
        """Charge toutes les données de la feuille en mémoire"""
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if all(cell is None or cell == "" for cell in row):
                continue
            self.worksheet_data[row_idx] = {}
            for col_idx, value in enumerate(row, 1):
                column_letter = get_column_letter(col_idx)
                self.worksheet_data[row_idx][column_letter] = value
    
    def _validate_conditional_rules(self, conditional_rules: List[Dict]):
        """Valide les règles conditionnelles"""
        for rule in conditional_rules:
            if not rule.get("active", True):
                continue
                
            conditions = rule.get("conditions", [])
            actions = rule.get("actions", [])
            logic = rule.get("logic", "AND")
            message = rule.get("message", "Règle conditionnelle non respectée")
            
            for row_idx, row_data in self.worksheet_data.items():
                if row_idx == 1:  # Skip header
                    continue
                    
                # Évaluer les conditions
                condition_results = []
                for condition in conditions:
                    column = condition["column"]
                    operator = condition["operator"]
                    value = condition.get("value", "")
                    
                    cell_value = row_data.get(column)
                    condition_result = self._evaluate_condition(cell_value, operator, value)
                    condition_results.append(condition_result)
                
                # Combiner les résultats selon la logique
                if logic == "AND":
                    conditions_met = all(condition_results)
                else:  # OR
                    conditions_met = any(condition_results)
                
                # Si conditions remplies, vérifier les actions
                if conditions_met:
                    for action in actions:
                        action_column = action["column"]
                        action_type = action["type"]
                        action_params = action.get("params", {})
                        
                        cell_value = row_data.get(action_column)
                        if not self._validate_action(cell_value, action_type, action_params):
                            error = ValidationError(row_idx, action_column, message, cell_value)
                            self.errors.append(error)
    
    def _evaluate_condition(self, cell_value: Any, operator: str, compare_value: str) -> bool:
        """Évalue une condition"""
        if cell_value is None:
            cell_value = ""
        
        cell_str = str(cell_value).strip()
        compare_str = str(compare_value).strip()
        
        try:
            if operator == "equals":
                return cell_str == compare_str
            elif operator == "not_equals":
                return cell_str != compare_str
            elif operator == "greater_than":
                return float(cell_value) > float(compare_value)
            elif operator == "less_than":
                return float(cell_value) < float(compare_value)
            elif operator == "greater_equal":
                return float(cell_value) >= float(compare_value)
            elif operator == "less_equal":
                return float(cell_value) <= float(compare_value)
            elif operator == "starts_with":
                return cell_str.startswith(compare_str)
            elif operator == "ends_with":
                return cell_str.endswith(compare_str)
            elif operator == "contains":
                return compare_str in cell_str
            elif operator == "not_contains":
                return compare_str not in cell_str
            elif operator == "is_empty":
                return cell_str == ""
            elif operator == "is_not_empty":
                return cell_str != ""
        except (ValueError, TypeError):
            return False
        
        return False
    
    def _validate_action(self, value: Any, action_type: str, params: Dict) -> bool:
        """Valide une action conditionnelle"""
        if action_type == "must_be_empty":
            return value is None or str(value).strip() == ""
        elif action_type == "must_not_be_empty":
            return value is not None and str(value).strip() != ""
        elif action_type == "must_be_between":
            try:
                val = float(value)
                min_val = params.get("min", 0)
                max_val = params.get("max", 100)
                return min_val <= val <= max_val
            except (ValueError, TypeError):
                return False
        elif action_type == "must_be_in_list":
            values_list = params.get("values", [])
            return str(value) in [str(v) for v in values_list]
        elif action_type == "must_match_pattern":
            pattern = params.get("pattern", "")
            try:
                return bool(re.match(pattern, str(value)))
            except re.error:
                return False
        
        return True
    
    def _validate_worksheet(self, ws, validators: Dict, default_validator: Optional[Dict], 
                          excludes: List[str], header_row: Any):
        """Valide une feuille de calcul avec les règles simples"""
        header_found = header_row is True
        
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if all(cell is None or cell == "" for cell in row):
                continue
            
            if not header_found and header_row != True:
                if header_row in row:
                    header_found = True
                continue
            
            for col_idx, value in enumerate(row, 1):
                column_letter = get_column_letter(col_idx)
                
                if column_letter in excludes:
                    continue
                
                if column_letter in validators:
                    for rule in validators[column_letter]:
                        self._apply_validation_rule(rule, value, row_idx, column_letter)
                elif default_validator:
                    for rule in default_validator:
                        self._apply_validation_rule(rule, value, row_idx, column_letter)
    
    def _apply_validation_rule(self, rule: Dict, value: Any, row: int, column: str):
        """Applique une règle de validation à une cellule"""
        rule_type = list(rule.keys())[0]
        rule_params = rule[rule_type]
        
        try:
            if rule_params.get("trim", False) and isinstance(value, str):
                value = value.strip()
            
            is_valid = True
            error_message = rule_params.get("message", f"Erreur de validation {rule_type}")
            
            if rule_type == "NotBlank":
                is_valid = self._validate_not_blank(value, rule_params)
            elif rule_type == "Length":
                is_valid = self._validate_length(value, rule_params)
            elif rule_type == "Type":
                is_valid = self._validate_type(value, rule_params)
            elif rule_type == "Regex":
                is_valid = self._validate_regex(value, rule_params)
            elif rule_type == "Email":
                is_valid = self._validate_email(value, rule_params)
            elif rule_type == "Choice":
                is_valid = self._validate_choice(value, rule_params)
            elif rule_type == "Country":
                is_valid = self._validate_country(value, rule_params)
            elif rule_type == "Date" or rule_type == "ExcelDate":
                is_valid = self._validate_date(value, rule_params)
            elif rule_type == "Comparison":
                is_valid = self._validate_comparison(value, rule_params)
            elif rule_type == "Duplicate":
                is_valid = self._validate_duplicate(value, column, row, rule_params)
            
            if not is_valid:
                error = ValidationError(row, column, error_message, value)
                self.errors.append(error)
                
        except Exception as e:
            error = ValidationError(row, column, f"Erreur lors de la validation: {str(e)}", value)
            self.errors.append(error)
    
    def _validate_not_blank(self, value: Any, params: Dict) -> bool:
        """Validation NotBlank"""
        return value is not None and str(value).strip() != ""
    
    def _validate_length(self, value: Any, params: Dict) -> bool:
        """Validation Length"""
        if value is None:
            return True
        str_value = str(value)
        min_length = params.get("min")
        max_length = params.get("max")
        
        if min_length is not None and len(str_value) < min_length:
            return False
        if max_length is not None and len(str_value) > max_length:
            return False
        return True
    
    def _validate_type(self, value: Any, params: Dict) -> bool:
        """Validation Type"""
        if value is None:
            return True
        expected_type = params.get("type", "").lower()
        
        try:
            if expected_type == "integer":
                int(value)
                return True
            elif expected_type == "float":
                float(value)
                return True
            elif expected_type == "bool":
                return str(value) in ["0", "1", "True", "False", "true", "false"]
        except (ValueError, TypeError):
            return False
        return True
    
    def _validate_regex(self, value: Any, params: Dict) -> bool:
        """Validation Regex"""
        if value is None:
            return True
        pattern = params.get("pattern")
        if not pattern:
            return True
        try:
            return bool(re.match(pattern, str(value)))
        except re.error:
            return False
    
    def _validate_email(self, value: Any, params: Dict) -> bool:
        """Validation Email"""
        if value is None:
            return True
        if not isinstance(value, str):
            return False
        try:
            return validate_email(value)
        except:
            return False
    
    def _validate_choice(self, value: Any, params: Dict) -> bool:
        """Validation Choice"""
        if value is None:
            return True
        choices = params.get("choices", [])
        case_sensitive = params.get("caseSensitive", True)
        
        str_value = str(value)
        if not case_sensitive:
            str_value = str_value.lower()
            choices = [str(choice).lower() for choice in choices]
        
        return str_value in choices
    
    def _validate_country(self, value: Any, params: Dict) -> bool:
        """Validation Country"""
        if value is None:
            return True
        try:
            pycountry.countries.get(name=str(value))
            return True
        except (KeyError, LookupError):
            return False
    
    def _validate_date(self, value: Any, params: Dict) -> bool:
        """Validation Date"""
        if value is None:
            return True
        if isinstance(value, datetime):
            return True
        if isinstance(value, str):
            date_format = params.get("format", "%Y-%m-%d")
            try:
                datetime.strptime(value, date_format)
                return True
            except ValueError:
                return False
        return False
    
    def _validate_comparison(self, value: Any, params: Dict) -> bool:
        """Validation Comparison"""
        if value is None:
            return True
        
        operator = params.get("operator", "equals")
        compare_value = params.get("value", "")
        
        return self._evaluate_condition(value, operator, compare_value)
    
    def _validate_duplicate(self, value: Any, column: str, current_row: int, params: Dict) -> bool:
        """Validation Duplicate"""
        if value is None or str(value).strip() == "":
            return True
        
        case_sensitive = params.get("caseSensitive", True)
        check_value = str(value)
        if not case_sensitive:
            check_value = check_value.lower()
        
        # Compter les occurrences dans la colonne
        count = 0
        for row_idx, row_data in self.worksheet_data.items():
            if row_idx == current_row:
                continue
            
            cell_value = row_data.get(column)
            if cell_value is not None:
                compare_value = str(cell_value)
                if not case_sensitive:
                    compare_value = compare_value.lower()
                
                if compare_value == check_value:
                    count += 1
        
        return count == 0
    
    def _generate_error_file(self, original_file: str, ws, errors: List[ValidationError]) -> str:
        """Génère un fichier Excel avec les erreurs marquées"""
        try:
            timestamp = int(time.time())
            base_name = os.path.splitext(os.path.basename(original_file))[0]
            error_file = os.path.join(tempfile.gettempdir(), 
                                    f"errors_{timestamp}_{base_name}.xlsx")
            
            wb = load_workbook(original_file, data_only=True)
            ws = wb.active if ws is None else wb[ws.title]
            
            red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
            
            for error in errors:
                try:
                    cell = ws[error.coordinate]
                    cell.fill = red_fill
                except:
                    continue
            
            wb.save(error_file)
            wb.close()
            return error_file
        except Exception as e:
            print(f"Erreur génération fichier d'erreurs: {e}")
            return None
    
    def get_errors_as_dataframe(self) -> pd.DataFrame:
        """Retourne les erreurs sous forme de DataFrame"""
        if not self.errors:
            return pd.DataFrame(columns=["Ligne", "Colonne", "Coordonnée", "Message", "Valeur"])
        
        data = []
        for error in self.errors:
            data.append({
                "Ligne": error.row,
                "Colonne": error.column,
                "Coordonnée": error.coordinate,
                "Message": error.message,
                "Valeur": str(error.value) if error.value is not None else ""
            })
        
        return pd.DataFrame(data)
    
    def get_validation_summary(self) -> Dict[str, Any]:
        """Retourne un résumé de la validation"""
        if not self.errors:
            return {
                "status": "success",
                "total_errors": 0,
                "message": "✅ Aucune erreur détectée ! Le fichier est conforme aux règles."
            }
        
        error_by_type = {}
        for error in self.errors:
            error_type = error.message.split(":")[0] if ":" in error.message else "Erreur générale"
            error_by_type[error_type] = error_by_type.get(error_type, 0) + 1
        
        return {
            "status": "error",
            "total_errors": len(self.errors),
            "errors_by_type": error_by_type,
            "message": f"❌ {len(self.errors)} erreur(s) détectée(s)"
        }