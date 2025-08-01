@echo off
echo ========================================
echo   Excel Validator Pro - Installation
echo ========================================
echo.

:: Définir le dossier d'installation court
set INSTALL_DIR=C:\gvenv_validator
set VENV_DIR=%INSTALL_DIR%\venv

:: Créer le dossier si nécessaire
if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
)

:: Se déplacer dans le dossier d'installation
cd /d "%INSTALL_DIR%"

:: Vérification de Python
echo [1/5] Verification de Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installé. Installez-le depuis https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Création de l'environnement virtuel
if not exist "%VENV_DIR%" (
    echo [2/5] Creation de l'environnement virtuel...
    python -m venv venv
)

:: Activation de l'environnement virtuel
echo [3/5] Activation de l'environnement...
call "%VENV_DIR%\Scripts\activate.bat"

:: Installer les dépendances
echo [4/5] Installation des dépendances...
if exist "%~dp0requirements.txt" (
    pip install --no-cache-dir -r "%~dp0requirements.txt"
) else (
    echo Aucun fichier requirements.txt trouvé dans le dossier du script.
    pause
    exit /b 1
)

:: Lancer l'application
echo [5/5] Lancement de l'application...
if exist "%~dp0app.py" (
    python "%~dp0app.py"
) else (
    echo ERREUR : Le fichier app.py est introuvable dans le dossier du script.
    pause
    exit /b 1
)

pause
