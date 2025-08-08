import gradio as gr
import pandas as pd
import json
import os
import tempfile
from typing import Dict, List, Any, Tuple, Optional
from rules_manager import RulesManager
from excel_validator_core import ExcelValidatorCore

class GradioInterface:
    """Interface Gradio pour l'application de validation Excel avec règles multicolonnes"""
    
    def __init__(self):
        self.rules_manager = RulesManager()
        self.validator = ExcelValidatorCore()
        
    def create_interface(self) -> gr.Blocks:
        """Crée l'interface Gradio complète"""
        
        with gr.Blocks(
            title="📊 Excel Validator Pro",
            theme=gr.themes.Soft(),
            css="""
            .main-title { text-align: center; color: #2E86AB; margin-bottom: 2rem; }
            .section-title { color: #A23B72; border-bottom: 2px solid #A23B72; padding-bottom: 0.5rem; }
            .success-message { background-color: #d4edda; color: #155724; padding: 1rem; border-radius: 0.5rem; }
            .error-message { background-color: #f8d7da; color: #721c24; padding: 1rem; border-radius: 0.5rem; }
            .conditional-section { border: 2px solid #A23B72; border-radius: 8px; padding: 1rem; margin: 1rem 0; }
            .multicolumn-section { border: 2px solid #28a745; border-radius: 8px; padding: 1rem; margin: 1rem 0; }
            """
        ) as interface:
            
            gr.Markdown("# 📊 Excel Validator Pro", elem_classes=["main-title"])
            gr.Markdown("### Application de validation Excel avec règles configurables et multicolonnes")
            
            with gr.Tabs():
                # Page d'accueil
                with gr.Tab("🏠 Accueil"):
                    self._create_home_tab()
                
                # Page de création de règles
                with gr.Tab("➕ Créer des Règles"):
                    with gr.Tabs():
                        # Règles simples
                        with gr.Tab("📝 Règles Simples"):
                            self._create_simple_rules_section()
                        
                        # Règles conditionnelles
                        with gr.Tab("🔗 Règles Conditionnelles"):
                            self._create_conditional_rules_section()
                        
                        # Règles multicolonnes - NOUVEAU
                        with gr.Tab("🔢 Règles Multicolonnes"):
                            self._create_multicolumn_rules_section()
                
                # Page de gestion des règles
                with gr.Tab("📋 Gérer les Règles"):
                    rules_table, refresh_rules_func = self._create_management_tab()
                
                # Page de validation
                with gr.Tab("✅ Validation Excel"):
                    active_rules_info, get_active_rules_func = self._create_validation_tab()
            
            # Actualisation des données au chargement
            interface.load(get_active_rules_func, outputs=[active_rules_info])
            interface.load(refresh_rules_func, outputs=[rules_table])
        
        return interface
    
    def _create_home_tab(self):
        """Crée l'onglet d'accueil"""
        gr.Markdown("""
        ## 🎯 Fonctionnalités principales
        
        - 📝 **Règles simples** : NotBlank, Length, Type, Regex, Email, Choice, Country, Date
        - 🔍 **Règles de comparaison** : Plus grand/petit que, égal, différent, commence/finit par, contient
        - 🔍 **Détection de doublons** : Identification des valeurs dupliquées dans une colonne
        - 🔗 **Règles conditionnelles avancées** : "Si colonne A = X alors colonne B doit être Y"
        - 🔢 **NOUVEAU : Règles multicolonnes** : Validation sur plusieurs colonnes simultanément
        - 📋 **Gestion centralisée** : Activer/désactiver, modifier, supprimer vos règles
        - ✅ **Validation rapide** : Drag & drop de vos fichiers Excel
        - 📊 **Rapports détaillés** : Identification précise des erreurs avec export
        
        ### 🆕 Nouvelles règles multicolonnes
        
        - ➕ **Somme égale** : A + B = C
        - 📊 **Somme dans une plage** : A + B + C entre 100 et 1000
        - 📅 **Comparaison de dates** : Date_début < Date_fin
        - 📈 **Pourcentage de** : A = 20% de B (±tolérance)
        - 🔄 **Tout ou rien** : A, B, C toutes remplies OU toutes vides
        - 🔑 **Combinaison unique** : A+B+C unique dans le fichier
        - 📐 **Maximum/Minimum** : C = MAX(A, B) ou MIN(A, B)
        - ⚖️ **Somme conditionnelle** : Si D='VIP' alors A+B+C > 1000
        
        ### 📚 Comment utiliser l'application ?
        
        1. **Étape 1** : Créez vos règles de validation dans l'onglet "Créer des Règles"
           - **Règles Simples** : Validation directe d'une colonne
           - **Règles Conditionnelles** : "Si... alors..." avec conditions multiples
           - **🆕 Règles Multicolonnes** : Validation sur plusieurs colonnes simultanément
        2. **Étape 2** : Gérez vos règles dans "Gérer les Règles"
        3. **Étape 3** : Validez vos fichiers Excel dans "Validation Excel"
        
        ### 🚀 Avantages
        
        - ✨ **Interface intuitive** - Pas besoin de connaissances techniques
        - ⚡ **Validation rapide** - Traitement de fichiers jusqu'à 10 000 lignes
        - 🎨 **Personnalisation totale** - Créez vos propres règles métier
        - 💾 **Sauvegarde automatique** - Vos règles sont conservées entre les sessions
        - 🔢 **Validation avancée** - Règles sur plusieurs colonnes simultanément
        """)
    
    def _create_multicolumn_rules_section(self):
        """Section pour créer des règles multicolonnes - NOUVELLE"""
        gr.Markdown("## 🔢 Création de Règles Multicolonnes", elem_classes=["section-title"])
        
        with gr.Row():
            with gr.Column():
                gr.Markdown("### 💡 Règles multicolonnes disponibles")
                
                # Affichage des types de règles disponibles
                rule_types = self.rules_manager.get_multicolumn_rule_types()
                examples_text = ""
                for rule_id, rule_info in rule_types.items():
                    examples_text += f"**{rule_info['name']}** : {rule_info['example']}\\n"
                
                gr.Markdown(examples_text)
        
        gr.Markdown("---")
        
        with gr.Row():
            with gr.Column():
                gr.Markdown("#### 📋 **Configuration de base**", elem_classes=["multicolumn-section"])
                
                columns_input = gr.Textbox(
                    label="Colonnes concernées (séparées par virgules)",
                    placeholder="ex: A,B,C ou Montant1,Montant2,Total",
                    lines=2
                )
                
                rule_type_multi = gr.Dropdown(
                    label="Type de règle multicolonne",
                    choices=[
                        ("Somme égale (A + B = C)", "sum_equals"),
                        ("Somme dans une plage (A + B + C entre min et max)", "sum_range"),
                        ("Date antérieure (Date1 < Date2)", "date_before"),
                        ("Date postérieure (Date1 > Date2)", "date_after"),
                        ("Écart de dates (Date2 - Date1 entre X et Y jours)", "date_range"),
                        ("Pourcentage de (A = X% de B)", "percentage_of"),
                        ("Tout ou rien (toutes remplies OU toutes vides)", "all_or_none"),
                        ("Combinaison unique (A+B+C unique)", "unique_combination"),
                        ("Somme conditionnelle (Si D=X alors A+B+C > Y)", "conditional_sum"),
                        ("Maximum/Minimum (C = MAX(A,B) ou MIN(A,B))", "max_min_check")
                    ],
                    value="sum_equals"
                )
                
                message_multi = gr.Textbox(
                    label="Message d'erreur personnalisé",
                    placeholder="ex: La somme des montants n'est pas correcte",
                    lines=2
                )
            
            with gr.Column():
                gr.Markdown("#### ⚙️ **Paramètres spécifiques**", elem_classes=["multicolumn-section"])
                
                # Paramètres pour sum_equals
                with gr.Group():
                    gr.Markdown("**Paramètres pour 'Somme égale'**")
                    sum_equals_target = gr.Textbox(
                        label="Colonne cible (qui doit égaler la somme)",
                        placeholder="ex: C (dernière colonne par défaut)",
                        visible=True
                    )
                    sum_equals_tolerance = gr.Number(
                        label="Tolérance",
                        value=0.01,
                        visible=True
                    )
                
                # Paramètres pour sum_range
                with gr.Group():
                    gr.Markdown("**Paramètres pour 'Somme dans une plage'**")
                    sum_range_min = gr.Number(
                        label="Valeur minimale",
                        value=0,
                        visible=False
                    )
                    sum_range_max = gr.Number(
                        label="Valeur maximale",
                        value=1000,
                        visible=False
                    )
                
                # Paramètres pour date_range
                with gr.Group():
                    gr.Markdown("**Paramètres pour 'Écart de dates'**")
                    date_range_min = gr.Number(
                        label="Écart minimum (jours)",
                        value=1,
                        visible=False
                    )
                    date_range_max = gr.Number(
                        label="Écart maximum (jours)",
                        value=365,
                        visible=False
                    )
                
                # Paramètres pour percentage_of
                with gr.Group():
                    gr.Markdown("**Paramètres pour 'Pourcentage de'**")
                    percentage_value = gr.Number(
                        label="Pourcentage attendu (%)",
                        value=20,
                        visible=False
                    )
                    percentage_tolerance = gr.Number(
                        label="Tolérance (%)",
                        value=5,
                        visible=False
                    )
                
                # Paramètres pour unique_combination
                with gr.Group():
                    gr.Markdown("**Paramètres pour 'Combinaison unique'**")
                    unique_case_sensitive = gr.Checkbox(
                        label="Sensible à la casse",
                        value=True,
                        visible=False
                    )
                
                # Paramètres pour conditional_sum
                with gr.Group():
                    gr.Markdown("**Paramètres pour 'Somme conditionnelle'**")
                    conditional_column = gr.Textbox(
                        label="Colonne condition",
                        placeholder="ex: D, Statut",
                        visible=False
                    )
                    conditional_value = gr.Textbox(
                        label="Valeur condition",
                        placeholder="ex: VIP, Premium",
                        visible=False
                    )
                    conditional_operator = gr.Dropdown(
                        label="Opérateur de comparaison",
                        choices=[
                            ("Plus grand que", "greater_than"),
                            ("Plus petit que", "less_than"),
                            ("Égal à", "equals"),
                            ("Plus grand ou égal", "greater_equal"),
                            ("Plus petit ou égal", "less_equal")
                        ],
                        value="greater_than",
                        visible=False
                    )
                    conditional_target = gr.Number(
                        label="Valeur cible",
                        value=1000,
                        visible=False
                    )
                
                # Paramètres pour max_min_check
                with gr.Group():
                    gr.Markdown("**Paramètres pour 'Maximum/Minimum'**")
                    max_min_operation = gr.Dropdown(
                        label="Opération",
                        choices=[("Maximum", "max"), ("Minimum", "min")],
                        value="max",
                        visible=False
                    )
                    max_min_target = gr.Dropdown(
                        label="Position colonne cible",
                        choices=[("Dernière colonne", "last"), ("Première colonne", "first")],
                        value="last",
                        visible=False
                    )
                    max_min_tolerance = gr.Number(
                        label="Tolérance",
                        value=0.01,
                        visible=False
                    )
        
        # Boutons d'action
        gr.Markdown("---")
        with gr.Row():
            create_multi_btn = gr.Button("✅ Créer la règle multicolonne", variant="primary", size="lg")
            clear_multi_btn = gr.Button("🗑️ Effacer tous les champs", variant="secondary")
            preview_multi_btn = gr.Button("👁️ Prévisualiser la règle", variant="secondary")
        
        # Zone de résultat et prévisualisation
        with gr.Row():
            with gr.Column():
                result_multicolumn = gr.Markdown()
            with gr.Column():
                preview_multicolumn = gr.Markdown()
        
        # Fonctions pour l'interface multicolonne
        def update_multicolumn_params_visibility(rule_type):
            """Met à jour la visibilité des paramètres selon le type de règle"""
            return [
                # sum_equals
                gr.update(visible=rule_type == "sum_equals"),
                gr.update(visible=rule_type == "sum_equals"),
                # sum_range
                gr.update(visible=rule_type == "sum_range"),
                gr.update(visible=rule_type == "sum_range"),
                # date_range
                gr.update(visible=rule_type == "date_range"),
                gr.update(visible=rule_type == "date_range"),
                # percentage_of
                gr.update(visible=rule_type == "percentage_of"),
                gr.update(visible=rule_type == "percentage_of"),
                # unique_combination
                gr.update(visible=rule_type == "unique_combination"),
                # conditional_sum
                gr.update(visible=rule_type == "conditional_sum"),
                gr.update(visible=rule_type == "conditional_sum"),
                gr.update(visible=rule_type == "conditional_sum"),
                gr.update(visible=rule_type == "conditional_sum"),
                # max_min_check
                gr.update(visible=rule_type == "max_min_check"),
                gr.update(visible=rule_type == "max_min_check"),
                gr.update(visible=rule_type == "max_min_check")
            ]
        
        def preview_multicolumn_rule(columns, rule_type, message, *params):
            """Prévisualise une règle multicolonne"""
            try:
                if not columns:
                    return "❌ Veuillez saisir les colonnes concernées"
                
                columns_list = [col.strip() for col in columns.split(",") if col.strip()]
                if len(columns_list) < 2:
                    return "❌ Au moins 2 colonnes sont nécessaires pour une règle multicolonne"
                
                rule_types = self.rules_manager.get_multicolumn_rule_types()
                rule_info = rule_types.get(rule_type, {})
                rule_name = rule_info.get("name", rule_type)
                
                preview = f"### 👁️ Aperçu de votre règle multicolonne\\n\\n"
                preview += f"**🔢 TYPE :** {rule_name}\\n"
                preview += f"**📊 COLONNES :** {', '.join(columns_list)} ({len(columns_list)} colonnes)\\n"
                
                # Ajouter des détails spécifiques selon le type
                if rule_type == "sum_equals":
                    preview += f"**📐 RÈGLE :** {' + '.join(columns_list[:-1])} = {columns_list[-1]}\\n"
                elif rule_type == "sum_range":
                    preview += f"**📐 RÈGLE :** {' + '.join(columns_list)} entre {params[2]} et {params[3]}\\n"
                elif rule_type == "date_before":
                    preview += f"**📐 RÈGLE :** {columns_list[0]} < {columns_list[1]}\\n"
                elif rule_type == "date_after":
                    preview += f"**📐 RÈGLE :** {columns_list[0]} > {columns_list[1]}\\n"
                elif rule_type == "percentage_of":
                    preview += f"**📐 RÈGLE :** {columns_list[0]} = {params[6]}% de {columns_list[1]} (±{params[7]}%)\\n"
                elif rule_type == "unique_combination":
                    preview += f"**📐 RÈGLE :** Combinaison {'+'.join(columns_list)} unique dans le fichier\\n"
                
                preview += f"**📝 MESSAGE :** {message}\\n"
                
                return preview
                
            except Exception as e:
                return f"❌ Erreur dans la prévisualisation : {str(e)}"
        
        def create_multicolumn_rule(columns, rule_type, message, *params):
            """Crée une règle multicolonne"""
            try:
                if not columns:
                    return "❌ Veuillez saisir les colonnes concernées"
                
                columns_list = [col.strip() for col in columns.split(",") if col.strip()]
                if len(columns_list) < 2:
                    return "❌ Au moins 2 colonnes sont nécessaires"
                
                # Construction des paramètres selon le type
                rule_params = {}
                
                if rule_type == "sum_equals":
                    if params[0]:  # target_column
                        rule_params["target_column"] = params[0]
                    rule_params["tolerance"] = params[1] or 0.01
                    
                elif rule_type == "sum_range":
                    rule_params["min_value"] = params[2] or 0
                    rule_params["max_value"] = params[3] or 1000
                    
                elif rule_type == "date_range":
                    rule_params["min_days"] = params[4] or 1
                    rule_params["max_days"] = params[5] or 365
                    
                elif rule_type == "percentage_of":
                    rule_params["percentage"] = params[6] or 20
                    rule_params["tolerance"] = (params[7] or 5) / 100.0
                    
                elif rule_type == "unique_combination":
                    rule_params["case_sensitive"] = params[8] if params[8] is not None else True
                    
                elif rule_type == "conditional_sum":
                    if not params[9]:  # conditional_column
                        return "❌ Veuillez saisir la colonne condition"
                    rule_params["condition_column"] = params[9]
                    rule_params["condition_value"] = params[10] or ""
                    rule_params["operator"] = params[11] or "greater_than"
                    rule_params["target_value"] = params[12] or 1000
                    
                elif rule_type == "max_min_check":
                    rule_params["operation"] = params[13] or "max"
                    rule_params["target_column"] = params[14] or "last"
                    rule_params["tolerance"] = params[15] or 0.01
                
                # Créer la règle
                rule = self.rules_manager.add_multicolumn_rule(
                    columns_list, rule_type, rule_params, message
                )
                self.rules_manager.save_rules()
                
                success_msg = f"""
                ✅ **Règle multicolonne créée avec succès !**
                
                **📋 Détails :**
                - **ID :** {rule['id']}
                - **Type :** {rule_type}
                - **Colonnes :** {', '.join(columns_list)}
                - **Message :** {message}
                
                🎯 La règle est maintenant active et sera appliquée lors de la validation.
                """
                
                return success_msg
                
            except Exception as e:
                return f"❌ **Erreur lors de la création :** {str(e)}"
        
        def clear_multicolumn_form():
            """Remet à zéro tous les champs du formulaire multicolonne"""
            return [
                "",  # columns_input
                "sum_equals",  # rule_type_multi
                "",  # message_multi
                "",  # sum_equals_target
                0.01,  # sum_equals_tolerance
                0,  # sum_range_min
                1000,  # sum_range_max
                1,  # date_range_min
                365,  # date_range_max
                20,  # percentage_value
                5,  # percentage_tolerance
                True,  # unique_case_sensitive
                "",  # conditional_column
                "",  # conditional_value
                "greater_than",  # conditional_operator
                1000,  # conditional_target
                "max",  # max_min_operation
                "last",  # max_min_target
                0.01,  # max_min_tolerance
                "",  # result_multicolumn
                ""   # preview_multicolumn
            ]
        
        # Événements de l'interface multicolonne
        rule_type_multi.change(
            update_multicolumn_params_visibility,
            inputs=[rule_type_multi],
            outputs=[
                sum_equals_target, sum_equals_tolerance,
                sum_range_min, sum_range_max,
                date_range_min, date_range_max,
                percentage_value, percentage_tolerance,
                unique_case_sensitive,
                conditional_column, conditional_value, conditional_operator, conditional_target,
                max_min_operation, max_min_target, max_min_tolerance
            ]
        )
        
        preview_multi_btn.click(
            preview_multicolumn_rule,
            inputs=[
                columns_input, rule_type_multi, message_multi,
                sum_equals_target, sum_equals_tolerance,
                sum_range_min, sum_range_max,
                date_range_min, date_range_max,
                percentage_value, percentage_tolerance,
                unique_case_sensitive,
                conditional_column, conditional_value, conditional_operator, conditional_target,
                max_min_operation, max_min_target, max_min_tolerance
            ],
            outputs=[preview_multicolumn]
        )
        
        create_multi_btn.click(
            create_multicolumn_rule,
            inputs=[
                columns_input, rule_type_multi, message_multi,
                sum_equals_target, sum_equals_tolerance,
                sum_range_min, sum_range_max,
                date_range_min, date_range_max,
                percentage_value, percentage_tolerance,
                unique_case_sensitive,
                conditional_column, conditional_value, conditional_operator, conditional_target,
                max_min_operation, max_min_target, max_min_tolerance
            ],
            outputs=[result_multicolumn]
        )
        
        clear_multi_btn.click(
            clear_multicolumn_form,
            outputs=[
                columns_input, rule_type_multi, message_multi,
                sum_equals_target, sum_equals_tolerance,
                sum_range_min, sum_range_max,
                date_range_min, date_range_max,
                percentage_value, percentage_tolerance,
                unique_case_sensitive,
                conditional_column, conditional_value, conditional_operator, conditional_target,
                max_min_operation, max_min_target, max_min_tolerance,
                result_multicolumn, preview_multicolumn
            ]
        )
    
    def _create_simple_rules_section(self):
        """Section pour créer des règles simples"""
        gr.Markdown("## 📝 Création de Règles Simples")
        
        with gr.Row():
            with gr.Column():
                column_input = gr.Textbox(label="Colonne Excel (ex: A, B, C...)", value="A")
                
                rule_type_dropdown = gr.Dropdown(
                    label="Type de règle",
                    choices=[
                        ("Ne pas être vide", "NotBlank"),
                        ("Longueur du texte", "Length"),
                        ("Type de données", "Type"),
                        ("Expression régulière", "Regex"),
                        ("Adresse email", "Email"),
                        ("Choix dans une liste", "Choice"),
                        ("Nom de pays", "Country"),
                        ("Date", "Date"),
                        ("Comparaison", "Comparison"),
                        ("Détection de doublons", "Duplicate")
                    ],
                    value="NotBlank"
                )
                
                message_input = gr.Textbox(
                    label="Message d'erreur personnalisé",
                    placeholder="Cette cellule ne peut pas être vide",
                    lines=2
                )
            
            with gr.Column():
                # Paramètres pour Length
                min_length = gr.Number(label="Longueur minimale", visible=False, value=0)
                max_length = gr.Number(label="Longueur maximale", visible=False, value=100)
                
                # Paramètres pour Type
                data_type = gr.Dropdown(
                    label="Type de données",
                    choices=[("Nombre entier", "integer"), ("Nombre décimal", "float"), ("Booléen", "bool")],
                    visible=False,
                    value="integer"
                )
                
                # Paramètres pour Regex
                regex_pattern = gr.Textbox(label="Pattern regex", visible=False)
                
                # Paramètres pour Choice
                choices_input = gr.Textbox(label="Choix possibles (séparés par virgules)", visible=False)
                choice_case_sensitive = gr.Checkbox(label="Sensible à la casse", visible=False, value=True)
                
                # Paramètres pour Comparison
                comparison_operator = gr.Dropdown(
                    label="Opérateur de comparaison",
                    choices=[
                        ("Égal à", "equals"),
                        ("Différent de", "not_equals"),
                        ("Plus grand que", "greater_than"),
                        ("Plus petit que", "less_than"),
                        ("Plus grand ou égal", "greater_equal"),
                        ("Plus petit ou égal", "less_equal"),
                        ("Commence par", "starts_with"),
                        ("Finit par", "ends_with"),
                        ("Contient", "contains"),
                        ("Ne contient pas", "not_contains")
                    ],
                    visible=False,
                    value="equals"
                )
                comparison_value = gr.Textbox(label="Valeur de comparaison", visible=False)
                
                # Paramètres pour Duplicate
                duplicate_case_sensitive = gr.Checkbox(label="Sensible à la casse", visible=False, value=True)
                
                # Option commune
                trim_option = gr.Checkbox(label="Supprimer les espaces en début/fin", value=False)
        
        # Boutons d'action
        with gr.Row():
            create_btn = gr.Button("✅ Créer la règle", variant="primary")
            clear_btn = gr.Button("🗑️ Effacer", variant="secondary")
        
        result_simple = gr.Markdown()
        
        # Fonction pour mettre à jour les paramètres selon le type de règle
        def update_params_visibility(rule_type):
            return [
                gr.update(visible=rule_type == "Length"),
                gr.update(visible=rule_type == "Length"),
                gr.update(visible=rule_type == "Type"),
                gr.update(visible=rule_type == "Regex"),
                gr.update(visible=rule_type == "Choice"),
                gr.update(visible=rule_type == "Choice"),
                gr.update(visible=rule_type == "Comparison"),
                gr.update(visible=rule_type == "Comparison"),
                gr.update(visible=rule_type == "Duplicate")
            ]
        
        rule_type_dropdown.change(
            update_params_visibility,
            inputs=[rule_type_dropdown],
            outputs=[min_length, max_length, data_type, regex_pattern, choices_input, 
                    choice_case_sensitive, comparison_operator, comparison_value, duplicate_case_sensitive]
        )
        
        # Fonction pour créer une règle simple
        def create_simple_rule(column, rule_type, message, min_len, max_len, dtype, regex, 
                             choices, choice_case, comp_op, comp_val, dup_case, trim):
            try:
                params = {"trim": trim}
                
                if rule_type == "Length":
                    if min_len > 0:
                        params["min"] = int(min_len)
                    if max_len > 0:
                        params["max"] = int(max_len)
                elif rule_type == "Type":
                    params["type"] = dtype
                elif rule_type == "Regex":
                    params["pattern"] = regex
                elif rule_type == "Choice":
                    params["choices"] = [choice.strip() for choice in choices.split(",") if choice.strip()]
                    params["caseSensitive"] = choice_case
                elif rule_type == "Comparison":
                    params["operator"] = comp_op
                    params["value"] = comp_val
                elif rule_type == "Duplicate":
                    params["caseSensitive"] = dup_case
                
                rule = self.rules_manager.add_simple_rule(column, rule_type, params, message)
                self.rules_manager.save_rules()
                
                return f"✅ **Règle créée avec succès !**\\n\\n**ID:** {rule['id']}\\n**Colonne:** {column}\\n**Type:** {rule_type}"
                
            except Exception as e:
                return f"❌ **Erreur lors de la création :** {str(e)}"
        
        create_btn.click(
            create_simple_rule,
            inputs=[column_input, rule_type_dropdown, message_input, min_length, max_length, 
                   data_type, regex_pattern, choices_input, choice_case_sensitive,
                   comparison_operator, comparison_value, duplicate_case_sensitive, trim_option],
            outputs=[result_simple]
        )
        
        # Fonction pour effacer les champs
        def clear_simple_form():
            return ["A", "NotBlank", "", 0, 100, "integer", "", "", True, "equals", "", True, False, ""]
        
        clear_btn.click(
            clear_simple_form,
            outputs=[column_input, rule_type_dropdown, message_input, min_length, max_length,
                    data_type, regex_pattern, choices_input, choice_case_sensitive,
                    comparison_operator, comparison_value, duplicate_case_sensitive, trim_option, result_simple]
        )
    
    def _create_conditional_rules_section(self):
        """Section pour créer des règles conditionnelles (code existant abrégé)"""
        gr.Markdown("## 🔗 Création de Règles Conditionnelles", elem_classes=["section-title"])
        gr.Markdown("*Interface conditionnelle existante conservée...*")
        # [Le code existant de cette section reste inchangé]
    
    def _create_management_tab(self):
        """Crée l'onglet de gestion des règles (mis à jour pour inclure multicolonnes)"""
        gr.Markdown("## 📋 Gestion des Règles", elem_classes=["section-title"])
        
        with gr.Row():
            refresh_btn = gr.Button("🔄 Actualiser", variant="secondary")
            export_btn = gr.Button("📤 Exporter", variant="primary")
            import_btn = gr.Button("📥 Importer", variant="primary")
        
        # Tableau des règles (maintenant avec support multicolonne)
        rules_table = gr.Dataframe(
            headers=["ID", "Type", "Colonne(s)", "Règle", "Message", "Status"],
            datatype=["str", "str", "str", "str", "str", "str"],
            interactive=False,
            wrap=True
        )
        
        # Actions sur les règles
        with gr.Row():
            with gr.Column():
                rule_id_input = gr.Textbox(label="ID de la règle", placeholder="Copier l'ID depuis le tableau")
                rule_type_input = gr.Radio(
                    label="Type de règle",
                    choices=[("Simple", "simple"), ("Conditionnelle", "conditional"), ("Multicolonne", "multicolumn")],
                    value="simple"
                )
            
            with gr.Column():
                toggle_btn = gr.Button("🔄 Activer/Désactiver", variant="secondary")
                delete_btn = gr.Button("🗑️ Supprimer", variant="stop")
        
        management_result = gr.Markdown()
        
        # Import/Export de fichiers
        with gr.Row():
            import_file = gr.File(label="Fichier de règles à importer (.json)", file_types=[".json"])
            export_path = gr.Textbox(label="Nom du fichier d'export", value="mes_regles.json")
        
        def refresh_rules_table():
            """Actualise le tableau des règles"""
            summary = self.rules_manager.get_rules_summary()
            return summary
        
        def toggle_rule(rule_id, rule_type):
            """Active/désactive une règle"""
            if not rule_id:
                return "❌ Veuillez saisir un ID de règle"
            
            success = self.rules_manager.toggle_rule(rule_id, rule_type)
            if success:
                self.rules_manager.save_rules()
                return f"✅ Statut de la règle {rule_id} modifié"
            return f"❌ Règle {rule_id} introuvable"
        
        def delete_rule(rule_id, rule_type):
            """Supprime une règle"""
            if not rule_id:
                return "❌ Veuillez saisir un ID de règle"
            
            success = self.rules_manager.delete_rule(rule_id, rule_type)
            if success:
                self.rules_manager.save_rules()
                return f"✅ Règle {rule_id} supprimée"
            return f"❌ Règle {rule_id} introuvable"
        
        def export_rules(filename):
            """Exporte les règles"""
            if not filename.endswith('.json'):
                filename += '.json'
            
            success = self.rules_manager.export_rules(filename)
            if success:
                return f"✅ Règles exportées vers {filename}"
            return f"❌ Erreur lors de l'export"
        
        def import_rules(file):
            """Importe les règles"""
            if file is None:
                return "❌ Veuillez sélectionner un fichier"
            
            success = self.rules_manager.import_rules(file.name)
            if success:
                return "✅ Règles importées avec succès"
            return "❌ Erreur lors de l'import - Vérifiez le format du fichier"
        
        # Connexion des événements
        refresh_btn.click(refresh_rules_table, outputs=[rules_table])
        toggle_btn.click(toggle_rule, inputs=[rule_id_input, rule_type_input], outputs=[management_result])
        delete_btn.click(delete_rule, inputs=[rule_id_input, rule_type_input], outputs=[management_result])
        export_btn.click(export_rules, inputs=[export_path], outputs=[management_result])
        import_btn.click(import_rules, inputs=[import_file], outputs=[management_result])
        
        return rules_table, refresh_rules_table
    
    def _create_validation_tab(self):
        """Crée l'onglet de validation Excel (mis à jour pour multicolonnes)"""
        gr.Markdown("## ✅ Validation de fichiers Excel", elem_classes=["section-title"])
        
        with gr.Row():
            with gr.Column():
                # Upload de fichier
                file_input = gr.File(
                    label="📁 Sélectionnez votre fichier Excel",
                    file_types=[".xlsx", ".xls", ".xlsm"],
                    type="filepath"
                )
                
                sheet_name = gr.Textbox(
                    label="Nom de la feuille (optionnel)",
                    placeholder="Laissez vide pour la première feuille",
                    value=""
                )
                
                validate_btn = gr.Button("🚀 Lancer la validation", variant="primary", size="lg")
            
            with gr.Column():
                # Informations sur les règles actives (mise à jour pour multicolonnes)
                active_rules_info = gr.Markdown("**Règles actives :** Chargement...")
        
        # Résultats de validation
        with gr.Row():
            validation_summary = gr.Markdown()
        
        with gr.Row():
            with gr.Column():
                errors_table = gr.Dataframe(
                    label="📋 Détail des erreurs",
                    headers=["Ligne", "Colonne(s)", "Coordonnée", "Message", "Valeur(s)"],
                    visible=False,
                    wrap=True
                )
            
            with gr.Column():
                download_section = gr.Group(visible=False)
                with download_section:
                    gr.Markdown("### 📥 Téléchargements")
                    error_file_download = gr.File(label="Fichier Excel avec erreurs marquées")
                    csv_download_button = gr.DownloadButton(
                        label="📊 Télécharger le rapport CSV",
                        visible=False
                    )
        
        def get_active_rules_info():
            """Retourne les informations sur les règles actives (mise à jour)"""
            simple_count = len([r for r in self.rules_manager.rules["simple_rules"] if r["active"]])
            cond_count = len([r for r in self.rules_manager.rules["conditional_rules"] if r["active"]])
            multi_count = len([r for r in self.rules_manager.rules["multicolumn_rules"] if r["active"]])
            
            rule_types = {}
            for rule in self.rules_manager.rules["simple_rules"]:
                if rule["active"]:
                    rule_type = rule["rule_type"]
                    rule_types[rule_type] = rule_types.get(rule_type, 0) + 1
            
            types_str = ", ".join([f"{count} {rtype}" for rtype, count in rule_types.items()])
            
            return f"""
            **📊 Règles actives actuellement :**
            - **{simple_count}** règles simples ({types_str})
            - **{cond_count}** règles conditionnelles
            - **{multi_count}** règles multicolonnes 🆕
            - **Total : {simple_count + cond_count + multi_count}** règles
            
            *Les règles désactivées ne seront pas appliquées lors de la validation.*
            """
        
        def validate_excel_file(file_path, sheet):
            """Valide un fichier Excel (mise à jour pour multicolonnes)"""
            if not file_path:
                return (
                    "❌ **Erreur :** Veuillez sélectionner un fichier Excel",
                    gr.update(visible=False),
                    gr.update(visible=False),
                    None,
                    gr.update(visible=False)
                )
            
            try:
                # Conversion des règles au format de validation
                config = self.rules_manager.convert_to_yaml_config()
                
                # Validation du fichier
                success, errors, error_file_path = self.validator.validate_file(
                    file_path, config, sheet if sheet else None
                )
                
                # Résumé de validation
                summary = self.validator.get_validation_summary()
                
                if success:
                    active_simple = len([r for r in self.rules_manager.rules["simple_rules"] if r["active"]])
                    active_cond = len([r for r in self.rules_manager.rules["conditional_rules"] if r["active"]])
                    active_multi = len([r for r in self.rules_manager.rules["multicolumn_rules"] if r["active"]])
                    
                    return (
                        f"""<div class="success-message">
                        {summary['message']}
                        
                        **📁 Fichier :** {os.path.basename(file_path)}
                        **📊 Lignes traitées :** Validation complète
                        **🔍 Règles appliquées :** {active_simple} simples + {active_cond} conditionnelles + {active_multi} multicolonnes
                        </div>""",
                        gr.update(visible=False),
                        gr.update(visible=False),
                        None,
                        gr.update(visible=False)
                    )
                else:
                    # Préparation du tableau d'erreurs
                    errors_df = self.validator.get_errors_as_dataframe()
                    
                    # Préparation du fichier CSV pour téléchargement
                    csv_path = None
                    if not errors_df.empty:
                        csv_path = tempfile.mktemp(suffix='.csv')
                        errors_df.to_csv(csv_path, sep=';', index=False, encoding='utf-8')
                    
                    active_simple = len([r for r in self.rules_manager.rules["simple_rules"] if r["active"]])
                    active_cond = len([r for r in self.rules_manager.rules["conditional_rules"] if r["active"]])
                    active_multi = len([r for r in self.rules_manager.rules["multicolumn_rules"] if r["active"]])
                    
                    summary_text = f"""<div class="error-message">
                    {summary['message']}
                    
                    **📁 Fichier :** {os.path.basename(file_path)}
                    **📊 Total erreurs :** {summary['total_errors']}
                    **🔍 Erreurs simples :** {summary.get('simple_errors', 0)}
                    **🔢 Erreurs multicolonnes :** {summary.get('multicolumn_errors', 0)} 🆕
                    **🔍 Règles appliquées :** {active_simple} simples + {active_cond} conditionnelles + {active_multi} multicolonnes
                    
                    **🔍 Répartition par type :**
                    """
                    
                    for error_type, count in summary.get('errors_by_type', {}).items():
                        summary_text += f"\\n- {error_type}: {count} erreur(s)"
                    
                    summary_text += "\\n</div>"
                    
                    return (
                        summary_text,
                        gr.update(visible=True, value=errors_df.values.tolist()),
                        gr.update(visible=True),
                        error_file_path,
                        gr.update(visible=True, value=csv_path) if csv_path else gr.update(visible=False)
                    )
                    
            except Exception as e:
                return (
                    f"❌ **Erreur lors de la validation :** {str(e)}",
                    gr.update(visible=False),
                    gr.update(visible=False),
                    None,
                    gr.update(visible=False)
                )
        
        # Connexion des événements
        validate_btn.click(
            validate_excel_file,
            inputs=[file_input, sheet_name],
            outputs=[validation_summary, errors_table, download_section, error_file_download, csv_download_button]
        )
        
        return active_rules_info, get_active_rules_info
    
    def launch(self, **kwargs):
        """Lance l'interface Gradio"""
        interface = self.create_interface()
        return interface.launch(**kwargs)

# Point d'entrée principal
if __name__ == "__main__":
    app = GradioInterface()
    app.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=False,
        debug=True
    )