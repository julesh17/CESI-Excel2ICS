# Convertisseur d’emplois du temps (Excel → ICS)

Cette application **Streamlit** permet de convertir un emploi du temps stocké dans un fichier Excel (`.xlsx`) au "bon format" en un fichier calendrier au format **ICS** compatible avec la plupart des logiciels (Google Calendar, Outlook, Apple Calendar, etc.).

Elle est conçue pour traiter des feuilles de type `EDT P1` ou `EDT P2` contenant des créneaux d’emplois du temps avec horaires, cours, enseignants et groupes.

---

## 🚀 Fonctionnalités

- Import d’un fichier **Excel** (`.xlsx`) contenant des emplois du temps.  
- Détection automatique des feuilles `EDT P1` et `EDT P2`.  
- Extraction des événements : matière, enseignants, groupes, description, créneaux horaires.  
- Conversion en événements **ICS** avec description enrichie.  
- Téléchargement direct des fichiers `.ics`.  
