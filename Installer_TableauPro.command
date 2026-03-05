#!/bin/bash

# Script d'installation automatique pour TableauPro
echo "--- Installation de l'extension Tableau Pro ---"

# 1. Définition des chemins
TARGET_DIR="$HOME/Library/Application Support/Adobe/CEP/extensions"
SOURCE_DIR="$HOME/Desktop/TableauPro"

# 2. Création des dossiers nécessaires
mkdir -p "$TARGET_DIR"

# 3. Nettoyage et création du lien symbolique
if [ -d "$SOURCE_DIR" ]; then
    rm -rf "$TARGET_DIR/TableauPro"
    ln -s "$SOURCE_DIR" "$TARGET_DIR/TableauPro"
    echo "✅ Lien symbolique créé vers le Bureau."
else
    echo "❌ Erreur : Le dossier 'TableauPro' est introuvable sur le Bureau."
    exit 1
fi

# 4. Activation du mode Debug pour toutes les versions récentes
defaults write com.adobe.CSXS.10 PlayerDebugMode 1
defaults write com.adobe.CSXS.11 PlayerDebugMode 1
defaults write com.adobe.CSXS.12 PlayerDebugMode 1
echo "✅ Autorisations (Mode Debug) activées."

echo "---"
echo "Installation terminée ! Relancez Illustrator."
echo "Allez dans Fenêtre > Extensions > Tableau Pro CSV."