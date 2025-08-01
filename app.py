#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Validator Pro - Application principale
"""

import sys
import os
import traceback
from pathlib import Path

# Ajouter le répertoire courant au PYTHONPATH
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

def main():
    """Fonction principale de l'application"""
    try:
        from gradio_interface import GradioInterface
        
        print("🚀 Démarrage d'Excel Validator Pro...")
        print("📊 Chargement de l'interface utilisateur...")
        
        app = GradioInterface()
        
        print("✅ Interface chargée avec succès !")
        print("🌐 Lancement du serveur web...")
        print("📱 Accessible à : http://localhost:7860")
        print("🔄 Appuyez sur Ctrl+C pour arrêter")
        print("-" * 60)
        
        app.launch(
            server_name="127.0.0.1",
            server_port=7860,
            share=False,
            debug=False,
            show_error=True,
            inbrowser=True
        )
        
    except ImportError as e:
        print("❌ Erreur d'importation :")
        print(f"   {e}")
        print("💡 Vérifiez que toutes les dépendances sont installées :")
        print("   pip install -r requirements.txt")
        sys.exit(1)
        
    except Exception as e:
        print("❌ Erreur inattendue :")
        print(f"   {e}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()