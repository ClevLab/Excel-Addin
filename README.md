# Excel-Addin
Réalisation d'un projet de démonstration pour concevoir un add-in pour Excel. 
Ce projet est en relation avec l'article de notre blog : [Réaliser un add-in / plugin Excel](https://blog.clevlab.fr/2018/04/30/realiser-un-add-in-plugin-excel/)

A l'intérieur de ce projet, nous avons également partagé une classe ExcelServices avec diverses fonctionnalités pour simplifier vos développements :
* **GetAllWorksheetsName(Workbook document)** : pour lister tous les onglets de votre Excel
* **GetSelectedWorksheet(Workbook document, String worksheetName)** : pour sélectionner le worksheet demandé
* **DeleteWorksheetsWithName(Workbook document, String nameToDelete)** : pour supprimer le worksheet demandé
* **GetColumnIndex(Worksheet worksheet, int rowIndex, String textToSearch)** : pour obtenir l'index d'une colonne, contenant le texte recherché
* **GetRowIndex(Worksheet worksheet, int columnIndex, String textToSearch)** : pour obtenir l'index d'une ligne, contenant le texte recherché
* **GetExcelColumnName(int columnIndex)** : pour obtenir le nom d'une colonne à partir de son index
* **CountRowsWithColor(Worksheet worksheet, int columnIndex, int firstRowIndex, int lastRowIndex, Color selectedColor)** : pour compter les cellules ayant une couleur donnée en background
