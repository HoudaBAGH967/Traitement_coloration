#!/bin/bash

# Heure de début
start_time=$(date +%s)
echo "Début du programme : $(date)"

#echo"Lancement du programme d'extractionet de coloration des information  a partir d'un fichier Excel"
fich_data="/home/ac2i/Traitement_Coloration/Input/Schemas_en_T_V20250604.xlsx"
fich_output="/home/ac2i/Traitement_Coloration/Output/Fichier_sortie.docx"
fich_dictionnaire="/home/ac2i/Traitement_Coloration/Input/dictionnaire.xlsx"
INPUT_DATA_FILE_PATH="/home/ac2i/Traitement_Coloration/Input/DEPO1-01_IRIS_TRI - Test.xlsx"
OUTPUT_FILE_PATH="/home/ac2i/Traitement_Coloration/Output/output.xlsx"
NOM_FICHIER_MOTS_CLES_PATH="/home/ac2i/Traitement_Coloration/Input/dictionnaire.xlsx"

python scenario_coloration.py "$fich_data" "$fich_output"

python traitement.py "$INPUT_DATA_FILE_PATH" "$OUTPUT_FILE_PATH" "$NOM_FICHIER_MOTS_CLES_PATH"

# Heure de fin
end_time=$(date +%s)
echo "Fin du programme : $(date)"

# Temps écoulé
elapsed=$((end_time - start_time))
minutes=$((elapsed / 60))
seconds=$((elapsed % 60))
echo "Temps écoulé : ${minutes} minutes et ${seconds} secondes"