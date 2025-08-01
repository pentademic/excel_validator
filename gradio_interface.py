import gradio as gr
import pandas as pd
import json
import os
import tempfile
from typing import Dict, List, Any, Tuple, Optional
from rules_manager import RulesManager
from excel_validator_core import ExcelValidatorCore

class GradioInterface:
    """Interface Gradio pour l'application de validation Excel"""
    
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
            """
        ) as interface:
            
            gr.Markdown("# 📊 Excel Validator Pro", elem_classes=["main-title"])
            gr.Markdown("### Application de validation Excel avec règles configurables")
            
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
                
                # Page de gestion des règles
                with gr.Tab("📋 Gérer les Règles"):
                    rules_table, refresh_rules_func = self._create_management_tab()
                
                # Page de validation
                with gr.Tab("✅ Validation Excel"):
                    active_rules_info, get_active_rules_func = self._create_validation_tab()
            
            # CORRECTION : Actualisation des données au chargement - DANS le contexte Blocks
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
        - 📋 **Gestion centralisée** : Activer/désactiver, modifier, supprimer vos règles
        - ✅ **Validation rapide** : Drag & drop de vos fichiers Excel
        - 📊 **Rapports détaillés** : Identification précise des erreurs avec export
        
        ### 📚 Comment utiliser l'application ?
        
        1. **Étape 1** : Créez vos règles de validation dans l'onglet "Créer des Règles"
           - **Règles Simples** : Validation directe d'une colonne
           - **Règles Conditionnelles** : "Si... alors..." avec conditions multiples
        2. **Étape 2** : Gérez vos règles dans "Gérer les Règles"
        3. **Étape 3** : Validez vos fichiers Excel dans "Validation Excel"
        
        ### 🚀 Avantages
        
        - ✨ **Interface intuitive** - Pas besoin de connaissances techniques
        - ⚡ **Validation rapide** - Traitement de fichiers jusqu'à 10 000 lignes
        - 🎨 **Personnalisation totale** - Créez vos propres règles métier
        - 💾 **Sauvegarde automatique** - Vos règles sont conservées entre les sessions
        """)
    
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
        """Section pour créer des règles conditionnelles"""
        gr.Markdown("## 🔗 Création de Règles Conditionnelles", elem_classes=["section-title"])
        
        with gr.Row():
            with gr.Column():
                gr.Markdown("### 💡 Exemple d'utilisation")
                gr.Markdown("""
                **Cas d'usage typique :**
                - Si colonne **Statut** = "VIP" **ET** colonne **Pays** = "France"
                - Alors colonne **Montant** doit être entre 1000 et 10000
                
                **Comment procéder :**
                1. Définissez vos conditions (jusqu'à 3)
                2. Choisissez l'opérateur logique (ET/OU)
                3. Définissez l'action à effectuer
                4. Personnalisez le message d'erreur
                """)
        
        gr.Markdown("---")
        
        with gr.Row():
            with gr.Column():
                gr.Markdown("#### 🔍 **ÉTAPE 1 : Définir les Conditions**", elem_classes=["conditional-section"])
                
                # Condition 1 (obligatoire)
                with gr.Group():
                    gr.Markdown("**🔸 Condition 1** (obligatoire)")
                    cond1_column = gr.Textbox(
                        label="Colonne à vérifier",
                        value="A",
                        placeholder="ex: A, B, Statut..."
                    )
                    cond1_operator = gr.Dropdown(
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
                            ("Ne contient pas", "not_contains"),
                            ("Est vide", "is_empty"),
                            ("N'est pas vide", "is_not_empty")
                        ],
                        value="equals"
                    )
                    cond1_value = gr.Textbox(
                        label="Valeur de comparaison",
                        placeholder="ex: VIP, 100, France...",
                        visible=True
                    )
                
                # Condition 2 (optionnelle)
                with gr.Group():
                    gr.Markdown("**🔸 Condition 2** (optionnelle)")
                    cond2_enabled = gr.Checkbox(
                        label="🔄 Activer la condition 2",
                        value=False
                    )
                    cond2_logic = gr.Radio(
                        label="Opérateur logique avec condition 1",
                        choices=[("ET (toutes les conditions)", "AND"), ("OU (au moins une condition)", "OR")],
                        value="AND",
                        visible=False
                    )
                    cond2_column = gr.Textbox(
                        label="Colonne à vérifier",
                        value="B",
                        placeholder="ex: B, C, Pays...",
                        visible=False
                    )
                    cond2_operator = gr.Dropdown(
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
                            ("Ne contient pas", "not_contains"),
                            ("Est vide", "is_empty"),
                            ("N'est pas vide", "is_not_empty")
                        ],
                        value="equals",
                        visible=False
                    )
                    cond2_value = gr.Textbox(
                        label="Valeur de comparaison",
                        placeholder="ex: France, 18, Premium...",
                        visible=False
                    )
                
                # Condition 3 (optionnelle)
                with gr.Group():
                    gr.Markdown("**🔸 Condition 3** (optionnelle)")
                    cond3_enabled = gr.Checkbox(
                        label="🔄 Activer la condition 3",
                        value=False
                    )
                    cond3_column = gr.Textbox(
                        label="Colonne à vérifier",
                        value="C",
                        placeholder="ex: C, D, Age...",
                        visible=False
                    )
                    cond3_operator = gr.Dropdown(
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
                            ("Ne contient pas", "not_contains"),
                            ("Est vide", "is_empty"),
                            ("N'est pas vide", "is_not_empty")
                        ],
                        value="equals",
                        visible=False
                    )
                    cond3_value = gr.Textbox(
                        label="Valeur de comparaison",
                        placeholder="ex: Actif, 2024, Premium...",
                        visible=False
                    )
            
            with gr.Column():
                gr.Markdown("#### ⚡ **ÉTAPE 2 : Définir l'Action**", elem_classes=["conditional-section"])
                
                # Action principale
                with gr.Group():
                    gr.Markdown("**🎯 Action à effectuer QUAND les conditions sont vraies**")
                    action_column = gr.Textbox(
                        label="Colonne cible (qui sera vérifiée)",
                        value="E",
                        placeholder="ex: E, F, Montant..."
                    )
                    action_type = gr.Dropdown(
                        label="Type de validation à appliquer",
                        choices=[
                            ("Doit être vide", "must_be_empty"),
                            ("Ne doit pas être vide", "must_not_be_empty"),
                            ("Doit être entre deux valeurs", "must_be_between"),
                            ("Doit être dans la liste", "must_be_in_list"),
                            ("Doit correspondre au pattern regex", "must_match_pattern")
                        ],
                        value="must_not_be_empty"
                    )
                    
                    # Paramètres d'action dynamiques
                    with gr.Group():
                        action_min = gr.Number(
                            label="Valeur minimale",
                            value=0,
                            visible=False
                        )
                        action_max = gr.Number(
                            label="Valeur maximale",
                            value=100,
                            visible=False
                        )
                        action_list = gr.Textbox(
                            label="Liste de valeurs autorisées (séparées par des virgules)",
                            placeholder="ex: Oui,Non,Peut-être",
                            visible=False
                        )
                        action_pattern = gr.Textbox(
                            label="Pattern regex à respecter",
                            placeholder="ex: \\\\d{2}-\\\\d{2}-\\\\d{4}",
                            visible=False
                        )
                
                # Message d'erreur et logique générale
                gr.Markdown("#### 📝 **ÉTAPE 3 : Configuration Finale**", elem_classes=["conditional-section"])
                
                with gr.Group():
                    main_logic = gr.Radio(
                        label="Si plusieurs conditions, logique générale",
                        choices=[
                            ("Toutes les conditions doivent être vraies (ET)", "AND"), 
                            ("Au moins une condition doit être vraie (OU)", "OR")
                        ],
                        value="AND"
                    )
                    
                    cond_message = gr.Textbox(
                        label="Message d'erreur personnalisé",
                        placeholder="ex: Les clients VIP doivent avoir un montant entre 1000 et 10000",
                        lines=3
                    )
        
        # Boutons d'action
        gr.Markdown("---")
        with gr.Row():
            create_cond_btn = gr.Button("✅ Créer la règle conditionnelle", variant="primary", size="lg")
            clear_cond_btn = gr.Button("🗑️ Effacer tous les champs", variant="secondary")
            preview_cond_btn = gr.Button("👁️ Prévisualiser la règle", variant="secondary")
        
        # Zone de résultat et prévisualisation
        with gr.Row():
            with gr.Column():
                result_conditional = gr.Markdown()
            with gr.Column():
                preview_conditional = gr.Markdown()
        
        # Fonctions pour l'interface conditionnelle
        def toggle_condition2(enabled):
            return [
                gr.update(visible=enabled),
                gr.update(visible=enabled),
                gr.update(visible=enabled),
                gr.update(visible=enabled)
            ]
        
        def toggle_condition3(enabled):
            return [
                gr.update(visible=enabled),
                gr.update(visible=enabled),
                gr.update(visible=enabled)
            ]
        
        def update_action_params(action_type):
            return [
                gr.update(visible=action_type == "must_be_between"),
                gr.update(visible=action_type == "must_be_between"),
                gr.update(visible=action_type == "must_be_in_list"),
                gr.update(visible=action_type == "must_match_pattern")
            ]
        
        def update_cond1_value_visibility(operator):
            return gr.update(visible=operator not in ["is_empty", "is_not_empty"])
        
        def update_cond2_value_visibility(operator):
            return gr.update(visible=operator not in ["is_empty", "is_not_empty"])
        
        def update_cond3_value_visibility(operator):
            return gr.update(visible=operator not in ["is_empty", "is_not_empty"])
        
        def preview_conditional_rule(c1_col, c1_op, c1_val, c2_enabled, c2_logic, c2_col, c2_op, c2_val,
                                   c3_enabled, c3_col, c3_op, c3_val, a_col, a_type, a_min, a_max, 
                                   a_list, a_pattern, message, logic):
            try:
                preview = "### 👁️ Aperçu de votre règle conditionnelle\\n\\n"
                preview += "**🔍 CONDITIONS :**\\n"
                preview += f"- Si colonne **{c1_col}** {c1_op.replace('_', ' ')} "
                
                if c1_op not in ["is_empty", "is_not_empty"]:
                    preview += f"**'{c1_val}'**"
                
                if c2_enabled and c2_col:
                    logic_word = "ET" if c2_logic == "AND" else "OU"
                    preview += f"\\n- {logic_word} colonne **{c2_col}** {c2_op.replace('_', ' ')} "
                    if c2_op not in ["is_empty", "is_not_empty"]:
                        preview += f"**'{c2_val}'**"
                
                if c3_enabled and c3_col:
                    logic_word = "ET" if logic == "AND" else "OU"
                    preview += f"\\n- {logic_word} colonne **{c3_col}** {c3_op.replace('_', ' ')} "
                    if c3_op not in ["is_empty", "is_not_empty"]:
                        preview += f"**'{c3_val}'**"
                
                preview += f"\\n\\n**⚡ ACTION :**\\n"
                preview += f"- Alors colonne **{a_col}** {a_type.replace('_', ' ').replace('must ', 'doit ')}"
                
                if a_type == "must_be_between":
                    preview += f" **{a_min}** et **{a_max}**"
                elif a_type == "must_be_in_list" and a_list:
                    preview += f" : **{a_list}**"
                elif a_type == "must_match_pattern" and a_pattern:
                    preview += f" : **{a_pattern}**"
                
                preview += f"\\n\\n**📝 MESSAGE :** {message}"
                
                return preview
                
            except Exception as e:
                return f"❌ Erreur dans la prévisualisation : {str(e)}"
        
        def create_conditional_rule(c1_col, c1_op, c1_val, c2_enabled, c2_logic, c2_col, c2_op, c2_val,
                                  c3_enabled, c3_col, c3_op, c3_val, a_col, a_type, a_min, a_max, 
                                  a_list, a_pattern, message, logic):
            try:
                # Construction des conditions
                conditions = [{
                    "column": c1_col,
                    "operator": c1_op,
                    "value": c1_val if c1_op not in ["is_empty", "is_not_empty"] else ""
                }]
                
                if c2_enabled and c2_col:
                    conditions.append({
                        "column": c2_col,
                        "operator": c2_op,
                        "value": c2_val if c2_op not in ["is_empty", "is_not_empty"] else ""
                    })
                
                if c3_enabled and c3_col:
                    conditions.append({
                        "column": c3_col,
                        "operator": c3_op,
                        "value": c3_val if c3_op not in ["is_empty", "is_not_empty"] else ""
                    })
                
                # Construction des actions
                action_params = {}
                if a_type == "must_be_between":
                    action_params = {"min": a_min, "max": a_max}
                elif a_type == "must_be_in_list":
                    action_params = {"values": [v.strip() for v in a_list.split(",") if v.strip()]}
                elif a_type == "must_match_pattern":
                    action_params = {"pattern": a_pattern}
                
                actions = [{
                    "column": a_col,
                    "type": a_type,
                    "params": action_params
                }]
                
                # Création de la règle
                rule = self.rules_manager.add_conditional_rule(conditions, actions, message, logic)
                self.rules_manager.save_rules()
                
                success_msg = f"""
                ✅ **Règle conditionnelle créée avec succès !**
                
                **📋 Détails :**
                - **ID :** {rule['id']}
                - **Conditions :** {len(conditions)} condition(s)
                - **Actions :** {len(actions)} action(s)
                - **Logique :** {logic}
                - **Message :** {message}
                
                🎯 La règle est maintenant active et sera appliquée lors de la validation.
                """
                
                return success_msg
                
            except Exception as e:
                return f"❌ **Erreur lors de la création :** {str(e)}"
        
        def clear_conditional_form():
            """Remet à zéro tous les champs du formulaire conditionnel"""
            return [
                "A",  # cond1_column
                "equals",  # cond1_operator
                "",  # cond1_value
                False,  # cond2_enabled
                "AND",  # cond2_logic
                "B",  # cond2_column
                "equals",  # cond2_operator
                "",  # cond2_value
                False,  # cond3_enabled
                "C",  # cond3_column
                "equals",  # cond3_operator
                "",  # cond3_value
                "E",  # action_column
                "must_not_be_empty",  # action_type
                0,  # action_min
                100,  # action_max
                "",  # action_list
                "",  # action_pattern
                "AND",  # main_logic
                "La condition n'est pas respectée",  # cond_message
                "",  # result_conditional
                ""   # preview_conditional
            ]
        
        # Événements de l'interface
        cond2_enabled.change(
            toggle_condition2, 
            inputs=[cond2_enabled], 
            outputs=[cond2_logic, cond2_column, cond2_operator, cond2_value]
        )
        
        cond3_enabled.change(
            toggle_condition3, 
            inputs=[cond3_enabled], 
            outputs=[cond3_column, cond3_operator, cond3_value]
        )
        
        action_type.change(
            update_action_params, 
            inputs=[action_type], 
            outputs=[action_min, action_max, action_list, action_pattern]
        )
        
        cond1_operator.change(update_cond1_value_visibility, inputs=[cond1_operator], outputs=[cond1_value])
        cond2_operator.change(update_cond2_value_visibility, inputs=[cond2_operator], outputs=[cond2_value])
        cond3_operator.change(update_cond3_value_visibility, inputs=[cond3_operator], outputs=[cond3_value])
        
        preview_cond_btn.click(
            preview_conditional_rule,
            inputs=[cond1_column, cond1_operator, cond1_value, cond2_enabled, cond2_logic,
                   cond2_column, cond2_operator, cond2_value, cond3_enabled, cond3_column,
                   cond3_operator, cond3_value, action_column, action_type, action_min,
                   action_max, action_list, action_pattern, cond_message, main_logic],
            outputs=[preview_conditional]
        )
        
        create_cond_btn.click(
            create_conditional_rule,
            inputs=[cond1_column, cond1_operator, cond1_value, cond2_enabled, cond2_logic,
                   cond2_column, cond2_operator, cond2_value, cond3_enabled, cond3_column,
                   cond3_operator, cond3_value, action_column, action_type, action_min,
                   action_max, action_list, action_pattern, cond_message, main_logic],
            outputs=[result_conditional]
        )
        
        clear_cond_btn.click(
            clear_conditional_form,
            outputs=[cond1_column, cond1_operator, cond1_value, cond2_enabled, cond2_logic,
                    cond2_column, cond2_operator, cond2_value, cond3_enabled, cond3_column,
                    cond3_operator, cond3_value, action_column, action_type, action_min,
                    action_max, action_list, action_pattern, main_logic, cond_message, 
                    result_conditional, preview_conditional]
        )
    
    def _create_management_tab(self):
        """Crée l'onglet de gestion des règles"""
        gr.Markdown("## 📋 Gestion des Règles", elem_classes=["section-title"])
        
        with gr.Row():
            refresh_btn = gr.Button("🔄 Actualiser", variant="secondary")
            export_btn = gr.Button("📤 Exporter", variant="primary")
            import_btn = gr.Button("📥 Importer", variant="primary")
        
        # Tableau des règles
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
                    choices=[("Simple", "simple"), ("Conditionnelle", "conditional")],
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
        
        # Retourner la fonction de refresh pour pouvoir l'utiliser dans create_interface
        return rules_table, refresh_rules_table
    
    def _create_validation_tab(self):
        """Crée l'onglet de validation Excel"""
        gr.Markdown("## ✅ Validation de fichiers Excel", elem_classes=["section-title"])
        
        with gr.Row():
            with gr.Column():
                # Upload de fichier
                file_input = gr.File(
                    label="📁 Sélectionnez votre fichier Excel",
                    file_types=[".xlsx", ".xls", ".xlsm"],
                    type="filepath"
                )
                
                # Options de validation
                sheet_name = gr.Textbox(
                    label="Nom de la feuille (optionnel)",
                    placeholder="Laissez vide pour la première feuille",
                    value=""
                )
                
                validate_btn = gr.Button("🚀 Lancer la validation", variant="primary", size="lg")
            
            with gr.Column():
                # Informations sur les règles actives
                active_rules_info = gr.Markdown("**Règles actives :** Chargement...")
        
        # Résultats de validation
        with gr.Row():
            validation_summary = gr.Markdown()
        
        with gr.Row():
            with gr.Column():
                errors_table = gr.Dataframe(
                    label="📋 Détail des erreurs",
                    headers=["Ligne", "Colonne", "Coordonnée", "Message", "Valeur"],
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
            """Retourne les informations sur les règles actives"""
            simple_count = len([r for r in self.rules_manager.rules["simple_rules"] if r["active"]])
            cond_count = len([r for r in self.rules_manager.rules["conditional_rules"] if r["active"]])
            
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
            - **Total : {simple_count + cond_count}** règles
            
            *Les règles désactivées ne seront pas appliquées lors de la validation.*
            """
        
        def validate_excel_file(file_path, sheet):
            """Valide un fichier Excel"""
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
                    return (
                        f"""<div class="success-message">
                        {summary['message']}
                        
                        **📁 Fichier :** {os.path.basename(file_path)}
                        **📊 Lignes traitées :** Validation complète
                        **🔍 Règles appliquées :** {len([r for r in self.rules_manager.rules["simple_rules"] if r["active"]])} simples + {len([r for r in self.rules_manager.rules["conditional_rules"] if r["active"]])} conditionnelles
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
                        errors_df.to_csv(csv_path,sep=';', index=False, encoding='utf-8')
                    
                    summary_text = f"""<div class="error-message">
                    {summary['message']}
                    
                    **📁 Fichier :** {os.path.basename(file_path)}
                    **📊 Total erreurs :** {summary['total_errors']}
                    **🔍 Règles appliquées :** {len([r for r in self.rules_manager.rules["simple_rules"] if r["active"]])} simples + {len([r for r in self.rules_manager.rules["conditional_rules"] if r["active"]])} conditionnelles
                    
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
        
        # Retourner les composants pour pouvoir les connecter dans create_interface
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