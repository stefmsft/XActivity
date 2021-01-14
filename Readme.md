# Script d'extraction de données d'un calendrier Office 365 afin de generer un fichier synthetique d'activité

## Prerequis
 - Python 3
 - Jupyter lab
 - Package additionnel Python
      - jupyterlab
      - Pandas
      - openpyxl
   Vous pouvez installer ces package par un "pip install -r requirement.txt"
    
## Preparation

il faut creer un fichier Excel connecté sur une source de donnée O365 en selectionnant la section calendrier seulement car les autres informations (contact, email, etc ..) ne nous interressent pas. Il aussi penser à étendre la colome "Catégorie" pour qu'il affiche les valeur plutot que le type "liste" dans cette colonne.

Il s'appelle Raw-Data-Calendar.xslx dans le code

Avant chaque execution du script, si des changement sont intervenu dans vos rendez vous, il faut ouvrir le fichier "raw" et demander un "Actualiser Tout" dans le menu "Donnée" d'Excel

## Usage

 - Ouvrez le notebook GenSpreadSheetRepport 
 - Executer les 2 premier cellules pour importer les modules et definir les fonctions
 - Renseignez les variables
    - YEAR_OF_QUERY
    - MONTH_OF_QUERY
    - OUT_EXCEL_NAME
    - IMPUT_RAW
  - Executer un chargement du fichier Raw en Dataframe
    - df_raw = pd.read_excel(IMPUT_RAW)
 - LAncez la generation du rapport de synthe pour le mois choisi
    - GenReport(YEAR_OF_QUERY,m,df_raw,OUT_EXCEL_NAME)
    
## Customisation

TBD