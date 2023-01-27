Attribute VB_Name = "Module1"
Sub Macro_PLR()
Attribute Macro_PLR.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False 'Emp�cher le changement de fen�tre
Application.DisplayAlerts = False 'Enlever les alertes'
Application.Calculation = xlManual 'Rendre le mode de calcul des formules automatique


Set Classeur_SuiviPLR = ActiveWorkbook
Set Feuille_ListeProjetsPLR = Classeur_SuiviPLR.Sheets("Liste projets PLR")
Set Feuille_ListePLR = Classeur_SuiviPLR.Sheets("Liste PLR")

'D�limitation du tableau pr�sent dans la liste projets PLR

Tableau_ListeProjetsPLR_PremLigne = 5
Tableau_ListeProjetsPLR_DerLigne = Feuille_ListeProjetsPLR.Cells(Rows.Count, Tableau_ListeProjetsPLR_NumeroColonnePLR).End(xlUp).Row 'la s�lection remonte la colonne jusqu'� trouver une valeur non vide
Tableau_ListeProjetsPLR_NumeroColonneAffaire = Feuille_ListeProjetsPLR.Range("Affaire").Column 'voir Gestionnaire de noms'
Tableau_ListeProjetsPLR_NumeroColonneSelect = Feuille_ListeProjetsPLR.Range("Select_PLR").Column 'voir Gestionnaire de noms'
Tableau_ListeProjetsPLR_NumeroColonnePLR = Feuille_ListeProjetsPLR.Range("PLR").Column 'voir Gestionnaire de noms'


Tableau_ListePLR_PremLigne = 2
Tableau_ListePLR_indice = Tableau_ListePLR_PremLigne


Feuille_ListeProjetsPLR.Range("Template").Hyperlinks(1).Follow 'Ouvre le template du PLR
Set Classeur_SuiviAffaireTemplate = ActiveWorkbook
Set Feuille_PLRTemplate = Classeur_SuiviAffaireTemplate.Sheets("PLR")

'D�limitation du tableau pr�sent dans le template du PLR, le template est la r�f�rence

Tableau_TemplatePLR_PremLigne = Feuille_PLRTemplate.Range("En_tetes").Rows(1).Row
Tableau_TemplatePLR_PremColonne = Feuille_PLRTemplate.Range("En_tetes").Columns(1).Column
Tableau_TemplatePLR_DerColonne = Feuille_PLRTemplate.Range("En_tetes").Columns.Count + Tableau_TemplatePLR_PremColonne - 1
Tableau_TemplatePLR_Longueur = Feuille_PLRTemplate.Range("En_tetes").Columns.Count
Tableau_TemplatePLR_Largeur = Feuille_PLRTemplate.Range("En_tetes").Rows.Count
Tableau_TemplatePLR_NumeroColonneRisque = Feuille_PLRTemplate.Range("Colonne_risque").Column 'voir Gestionnaire de noms

Feuille_PLRTemplate.Range("En_tetes").Copy 'Copier l'en-t�te du template
Feuille_ListePLR.Cells(Tableau_ListePLR_PremLigne, Tableau_TemplatePLR_PremColonne).PasteSpecial (xlPasteAll) 'Coller l'en-t�te du template
Tableau_ListePLR_indice = Tableau_ListePLR_indice + Tableau_TemplatePLR_Largeur 'L'indice se d�place � la ligne apr�s l'en-t�te

Classeur_SuiviAffaireTemplate.Close False

Feuille_ListePLR.Rows(Tableau_ListePLR_PremLigne + Tableau_TemplatePLR_Largeur & ":" & Feuille_ListePLR.Rows.Count).Delete 'Supprimer l'ancien contenu de la feuille

For i = Tableau_ListeProjetsPLR_PremLigne To Tableau_ListeProjetsPLR_DerLigne
'Boucle sur tout le tableau Liste projets PLR

    If Not IsEmpty(Feuille_ListeProjetsPLR.Cells(i, Tableau_ListeProjetsPLR_NumeroColonneSelect)) Then
    'Si Select PLR est non vide sur la ligne'
    
        'Mise en forme du ruban vertical et horizontal de l'affaire
        Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne) = Feuille_ListeProjetsPLR.Cells(i, Tableau_ListeProjetsPLR_NumeroColonneAffaire).Value 'Ecrit le nom d'affaire sur le ruban horizontal
        Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne).HorizontalAlignment = xlCenter 'Alignement du ruban horizontal
        Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne).VerticalAlignment = xlCenter 'Alignement du ruban horizontal
        Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne).Font.Size = 20 'Taille du texte du ruban horizontal
        Feuille_ListePLR.Range(Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne), Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_DerColonne)).Merge 'Fusionne les cellules du ruban horizontal
        Feuille_ListePLR.Range(Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne), Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_DerColonne)).Borders.Color = RGB(0, 0, 0) 'Bords du ruban horizontal
        Feuille_ListePLR.Range(Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne), Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_DerColonne)).Borders.Weight = xlThick 'Bords �pais du ruban horizontal
        Feuille_ListePLR.Range(Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne - 1), Feuille_ListePLR.Cells(Feuille_ListePLR.Cells(Rows.Count, Tableau_TemplatePLR_PremColonne).End(xlUp).Row, Tableau_TemplatePLR_PremColonne - 1)) = Feuille_ListeProjetsPLR.Cells(i, Tableau_ListeProjetsPLR_NumeroColonneAffaire) 'Num�ros d'affaire du ruban vertical
        Feuille_ListePLR.Range(Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne - 1), Feuille_ListePLR.Cells(Feuille_ListePLR.Cells(Rows.Count, Tableau_TemplatePLR_PremColonne).End(xlUp).Row, Tableau_TemplatePLR_PremColonne - 1)).HorizontalAlignment = xlCenter 'Alignement des num�ros d'affaire du ruban vertical
        Feuille_ListePLR.Range(Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne - 1), Feuille_ListePLR.Cells(Feuille_ListePLR.Cells(Rows.Count, Tableau_TemplatePLR_PremColonne).End(xlUp).Row, Tableau_TemplatePLR_PremColonne - 1)).VerticalAlignment = xlCenter 'Alignement des num�ros d'affaire du ruban vertical
       
        Tableau_ListePLR_indice = Tableau_ListePLR_indice + 1 'L'indice se d�place � la ligne apr�s le ruban
        
        Feuille_ListeProjetsPLR.Cells(i, Tableau_ListeProjetsPLR_NumeroColonnePLR).Hyperlinks(1).Follow 'Le PLR est ouvert
        Set Classeur_SuiviAffaire = ActiveWorkbook
        Set Feuille_PLR = Classeur_SuiviAffaire.Sheets("PLR")
        
        Tableau_PLR_DerLigne = Feuille_PLR.Cells(Rows.Count, Tableau_TemplatePLR_NumeroColonneRisque).End(xlUp).Row 'la s�lection remonte la colonne risque du PLR ouvert jusqu'� trouver une valeur non vide
        Feuille_PLR.Range(Feuille_PLR.Cells(Tableau_TemplatePLR_PremLigne, Tableau_TemplatePLR_PremColonne), Feuille_PLR.Cells(Tableau_PLR_DerLigne, Tableau_TemplatePLR_DerColonne)).SpecialCells(xlCellTypeVisible).Copy 'Copier le PLR sauf en-t�te
        Feuille_ListePLR.Cells(Tableau_ListePLR_indice, Tableau_TemplatePLR_PremColonne).PasteSpecial (xlPasteAll) 'Coller le PLR � l'indice du tableau dans la feuille liste PLR
        Feuille_ListePLR.Rows(Tableau_ListePLR_indice & ":" & Tableau_ListePLR_indice + Tableau_TemplatePLR_Largeur - 1).Delete
        
        Tableau_ListePLR_indice = Feuille_ListePLR.Cells(Rows.Count, Tableau_TemplatePLR_NumeroColonneRisque).End(xlUp).Row + 1 'Calcul de la derni�re ligne du tableau dans la feuille liste PLR
        
        Classeur_SuiviAffaire.Close False
        
    End If
Next

'Suppression des lignes vides pr�sentes par d�faut dans les PLR

i = Tableau_ListePLR_PremLigne + Tableau_TemplatePLR_Largeur

Do While (i <> Feuille_ListePLR.Cells(Rows.Count, Tableau_TemplatePLR_NumeroColonneRisque).End(xlUp).Row)
'Boucle jusqu'� ce que i d�passe les limites du tableau dans la feuille liste PLR, la limite est recalcul�e � chaque it�ration
    
    If (IsEmpty(Feuille_ListePLR.Cells(i, Tableau_TemplatePLR_PremColonne)) Or Feuille_ListePLR.Cells(i, Tableau_TemplatePLR_PremColonne) = "") Then
    'Si la cellule de la colonne date est vide la ligne est supprim�e
        
        Feuille_ListePLR.Rows(i).Delete
        i = i - 1
    
    End If
    
    i = i + 1
Loop
       
Feuille_ListePLR.Rows.RowHeight = 30
    
Application.ScreenUpdating = True 'Emp�cher le changement de fen�tre
Application.DisplayAlerts = True 'Enlever les alertes'
Application.Calculation = xlAutomatic 'Rendre le mode de calcul des formules automatique

End Sub
