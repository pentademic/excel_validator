@echo off
echo ========================================
echo   Excel Validator Pro - Installation
echo ========================================
echo.

:: Définir le dossier de l'application comme dossier courant
cd /d "%~dp0"

:: Définir le dossier de l'environnement virtuel dans le dossier de l'app
set VENV_DIR=venv

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
if errorlevel 1 (
    echo [ERREUR] Impossible d'activer l'environnement virtuel.
    pause
    exit /b 1
)

:: Installation des dépendances
echo [4/5] Installation des dependances...
if exist "requirements.txt" (
    pip install --no-cache-dir -r "requirements.txt"
    if errorlevel 1 (
        echo [ERREUR] Impossible d'installer les dependances.
        pause
        exit /b 1
    )
) else (
    echo [ERREUR] Le fichier requirements.txt est introuvable.
    pause
    exit /b 1
)

:: Lancement de l'application
echo [5/5] Lancement de l'application...
if exist "app.py" (
    python "app.py"
    if errorlevel 1 (
        echo [ERREUR] Une erreur s'est produite lors du lancement de l'application.
        pause
        exit /b 1
    )
) else (
    echo [ERREUR] Le fichier app.py est introuvable.
    pause
    exit /b 1
)

pause
