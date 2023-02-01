Attribute VB_Name = "Module2"

Sub Cr�er_barres(ConcernedRange As Range, WarningsAR_MaxProg As Double, Feuille_SuiviAR As Worksheet)
    'ConcernedRange est la plage de donn�es sur laquelle les barres vont �tre appliqu�es
    ConcernedRange.FormatConditions.AddDatabar
    ConcernedRange.FormatConditions(ConcernedRange.FormatConditions.Count).ShowValue = True
    ConcernedRange.FormatConditions(ConcernedRange.FormatConditions.Count).SetFirstPriority
    ConcernedRange.FormatConditions(1).MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    ConcernedRange.FormatConditions(1).MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=Abs(WarningsAR_MaxProg)
                        
    ConcernedRange.FormatConditions(1).BarColor.Color = 13012579
    ConcernedRange.FormatConditions(1).BarColor.TintAndShade = 0
                        
    ConcernedRange.FormatConditions(1).BarFillType = xlDataBarFillGradient
    ConcernedRange.FormatConditions(1).Direction = xlContext
    ConcernedRange.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    ConcernedRange.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    ConcernedRange.FormatConditions(1).NegativeBarFormat.BorderColorType = xlDataBarColor
    ConcernedRange.FormatConditions(1).BarBorder.Color.Color = 13012579
    ConcernedRange.FormatConditions(1).BarBorder.Color.TintAndShade = 0
                        
    ConcernedRange.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    ConcernedRange.FormatConditions(1).AxisColor.Color = 0
    ConcernedRange.FormatConditions(1).AxisColor.TintAndShade = 0
                        
    ConcernedRange.FormatConditions(1).NegativeBarFormat.Color.Color = 255
    ConcernedRange.FormatConditions(1).NegativeBarFormat.Color.TintAndShade = 0
                        
    ConcernedRange.FormatConditions(1).NegativeBarFormat.BorderColor.Color = 255
    ConcernedRange.FormatConditions(1).NegativeBarFormat.BorderColor.TintAndShade = 0
End Sub


Sub Suivi()

    Dim Classeur_GDP04 As Workbook
    Dim Feuille_SuiviAR As Worksheet
    Dim Feuille_ListeProjetsAR As Worksheet
    
    Dim Tableau_ListeProjetsAR_AffaireVoulue As String
    Dim Tableau_ListeProjetsAR_DateVoulue As Variant
    
    Dim Feuille_WarningsAR As Worksheet
    
    Dim Tableau_WarningsAR_nColonneRR As Long
    Dim Tableau_WarningsAR_nColonneRP As Long
    Dim Tableau_ListeProjetsAR_nColonneDV As Long
    Dim Tableau_ListeProjetsAR_nColonneAV As Long
    Dim Tableau_ListeProjetsAR_nColonneSelect As Long
    Dim Tableau_WarningsAR_nColonneAffaire As Long
    Dim Tableau_ListeProjetsAR_nColonneAutre1 As Long
    Dim Tableau_ListeProjetsAR_nColonneAutre2 As Long
    Dim Tableau_ListeProjetsAR_IndiceAV As Long
    Dim DateAjd As Date
    Dim PlageR�sultante As Variant
    Dim WarningsAR_MaxProg As Double
    
    Dim Feuille_Suivi As Worksheet
    
    Dim Tableau_Suivi_Indice As Long
    
    Dim Feuille_ExtractNomclProjets As Worksheet
    
    Dim Chemin_GDP06 As String
    Dim Classeur_GDP06 As Workbook
    Dim Feuille_Feuil1 As Worksheet
    Dim objConnection As WorkbookConnection
    
    Dim Tableau_GDP06 As Variant
    Dim Tableau_GDP06_Affaire As Variant
    Dim Tableau_GDP06_Texte As String
    Dim Tableau_GDP06_DateAR As Variant
    Dim Tableau_GDP06_nCommande As Variant
    Dim Tableau_GDP06_NomFournisseur As Variant
    Dim Tableau_GDP06_Commentaire As Variant
    Dim Tableau_GDP06_Ref As Variant
    Dim Tableau_GDP06_DateLiv As Variant
    Dim Tableau_GDP06_Rubrique As Variant
    Dim Tableau_GDP06_QteRestante As Variant
    Dim Tableau_GDP06_Qte As Long
    Dim Tableau_GDP06_DerLigne As Double
    Dim Tableau_GDP06_DerColonne As Long
    
    Dim Classeur_Nomenclature As Workbook
    Dim Feuille_Nomenclature As Worksheet
    
    Application.ScreenUpdating = False 'Emp�cher le changement de fen�tre
    Application.DisplayAlerts = False 'Enlever les alertes'
    Application.Calculation = xlAutomatic 'Rendre le mode de calcul des formules automatique
    
    Chemin_GDP06 = "T:\ZZ_Planning\CDP\GDP_006_A_Extract CMD EVERWIN (base donn�es).xlsx" 'Le chemin d'acc�s du classeur Extract_CMD
    
    Set Classeur_GDP04 = ActiveWorkbook
    
    Set Feuille_ListeProjetsAR = Classeur_GDP04.Sheets("Liste projets AR")
    Tableau_ListeProjetsAR_PremLigne = Feuille_ListeProjetsAR.Range("ListeProjetsAR_ET").Rows(1).Row
    Tableau_ListeProjetsAR_nColonneDV = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Date de besoin").Column
    Tableau_ListeProjetsAR_nColonneAV = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Num�ro affaire").Column
    Tableau_ListeProjetsAR_nColonneSelect = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Select Suivi").Column
    Tableau_ListeProjetsAR_nColonneMeca = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature M�ca").Column
    Tableau_ListeProjetsAR_DerLigne = Feuille_ListeProjetsAR.Cells(Rows.Count, Tableau_ListeProjetsAR_nColonneMeca).End(xlUp).Row 'La derni�re ligne � traiter de la feuille "Liste projets AR" est calcul�e en remontant la colonne "Nomenclature m�ca" jusqu'� trouver une cellule non vide
    Tableau_ListeProjetsAR_nColonneElec = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature Elec").Column
    Tableau_ListeProjetsAR_nColonneAutre1 = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature 3").Column
    Tableau_ListeProjetsAR_nColonneAutre2 = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature 4").Column
    Tableau_ListeProjetsAR_IndiceAV = Tableau_ListeProjetsAR_PremLigne + 1 'L'indice qui parcoure les affaires dans Liste projets AR
    
    Set Feuille_WarningsAR = Classeur_GDP04.Sheets("Warnings AR")
    Tableau_WarningsAR_PremLigne = Feuille_WarningsAR.Range("WarningsAR_ET").Rows(1).Row 'la ligne o� commence le tableau des warnings pour l'affaire voulue
    Tableau_WarningsAR_PremColonne = Feuille_WarningsAR.Range("WarningsAR_ET").Columns(1).Column
    Tableau_WarningsAR_nColonneRR = Feuille_WarningsAR.Rows(Tableau_WarningsAR_PremLigne).Find("Retard de r�ception Sym�trie (en jours)").Column
    Tableau_WarningsAR_nColonneRP = Feuille_WarningsAR.Rows(Tableau_WarningsAR_PremLigne).Find("Retard projet (en jours)").Column
    Tableau_WarningsAR_nColonneAffaire = Feuille_WarningsAR.Rows(Tableau_WarningsAR_PremLigne).Find("Affaire").Column
    Tableau_WarningsAR_DerColonne = Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Feuille_WarningsAR.Columns.Count).End(xlToLeft).Column
    WarningsAR_MaxProg = 0 'Le max des barres de progression
    DateAjd = Date
    
    Set Feuille_Suivi = Classeur_GDP04.Sheets("Suivi")
    Tableau_Suivi_Indice = 3 'L'indice qui parcoure la feuille Suivi
    Tableau_Suivi_PremColonne = 2 'La premi�re colonne o� commencent les diff�rents tableaux
    Feuille_Suivi.Rows(Tableau_Suivi_Indice & ":" & Feuille_Suivi.Rows.Count).Delete
     
    Set Feuille_ExtractNomclProjets = Classeur_GDP04.Sheets("Extract Nomcl projets")
    
    Set Feuille_Nomenclatures = Classeur_GDP04.Sheets("Nomenclatures")
    Tableau_Nomenclatures_PremLigne = Feuille_Nomenclatures.Range("Nomenclatures_ET").Rows(1).Row
    Tableau_Nomenclatures_PremColonne = Feuille_Nomenclatures.Range("Nomenclatures_ET").Columns(1).Column
    Tableau_Nomenclatures_DerColonne = Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_PremLigne, Feuille_Nomenclatures.Columns.Count).End(xlToLeft).Column

    
    '------------------------------------------------------------------------------------------
    'Mettre � jour l'extract CMD
    
    If MsgBox("Mettre � jour la BDD Everwin ?", vbYesNo) = vbYes Then
        
        Set Classeur_GDP06 = Workbooks.Open(Chemin_GDP06)
        Classeur_GDP06.RefreshAll
    
    Feuille_ListeProjetsAR.Cells(2, 6) = DateAjd & Chr(13) & Chr(10) & Time
    Classeur_GDP06.Close True
    
    End If
    
    '------------------------------------------------------------------------------------------
    
    Set Classeur_GDP06 = Workbooks.Open(Chemin_GDP06)
    Set Feuille_Feuil1 = Classeur_GDP06.Sheets("Feuil1")
    
    Tableau_GDP06_DerLigne = Feuille_Feuil1.Cells(Rows.Count, 1).End(xlUp).Row
    Tableau_GDP06_DerColonne = Feuille_Feuil1.Cells(1, Columns.Count).End(xlToLeft).Column
    Tableau_GDP06 = Feuille_Feuil1.Range(Feuille_Feuil1.Cells(1, 1), Feuille_Feuil1.Cells(Tableau_GDP06_DerLigne, Tableau_GDP06_DerColonne)).Value 'La lecture du tableau se fait en une seule fois par cette commande
    
    Do While (Not IsEmpty(Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneAV)))
    'La premi�re boucle s'arr�te quand une cellule de la colonne "Numero affaire" du tableau pr�sent dans la feuille "Liste projets AR" est vide
    
    
        Tableau_ListeProjetsAR_AffaireVoulue = Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneAV)
        Tableau_ListeProjetsAR_DateVoulue = Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneDV)
        
        If (Not IsEmpty(Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneSelect))) Then
        'La condition v�rifie que la s�lection est coch�e pour l'affaire voulue
           
            ExtractNomclProjets_LongueurPlageAffaire = Feuille_ExtractNomclProjets.Range("Affaire_" & Tableau_ListeProjetsAR_AffaireVoulue).Rows.Count 'Voir Gestionnaire de noms
            ExtractNomclProjets_LargeurPlageAffaire = Feuille_ExtractNomclProjets.Range("Affaire_" & Tableau_ListeProjetsAR_AffaireVoulue).Columns.Count 'Voir Gestionnaire de noms
                    
            Feuille_ExtractNomclProjets.Range("Affaire_" & Tableau_ListeProjetsAR_AffaireVoulue).Copy
            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Suivi_PremColonne).PasteSpecial (xlPasteAll)
            
            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Suivi_PremColonne + ExtractNomclProjets_LargeurPlageAffaire - 1), Feuille_Suivi.Cells(Tableau_Suivi_Indice + ExtractNomclProjets_LongueurPlageAffaire - 1, Tableau_Suivi_PremColonne + ExtractNomclProjets_LargeurPlageAffaire - 1)).Formula = Feuille_ExtractNomclProjets.Range("Affaire_" & Tableau_ListeProjetsAR_AffaireVoulue).Columns(ExtractNomclProjets_LargeurPlageAffaire).Formula
        
            Tableau_Suivi_Indice = Tableau_Suivi_Indice + ExtractNomclProjets_LongueurPlageAffaire + 1
               
            Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_PremLigne, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_PremLigne, Tableau_Nomenclatures_DerColonne)).Copy
            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Suivi_PremColonne).PasteSpecial (xlPasteAll)
                
            For j = Tableau_ListeProjetsAR_nColonneMeca To Tableau_ListeProjetsAR_nColonneAutre2
            'Cette boucle s'arr�te quand toutes les nomenclatures de la ligne ont �t� parcourues
    
                    
                If Not IsEmpty(Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, j)) And Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, j) <> "" Then
                'La condition v�rifie que les cellules des colonnes Nomenclatures ne sont pas vides
                    
                        
                    Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, j).Hyperlinks(1).Follow
                    Set Classeur_Nomenclature = ActiveWorkbook
                    Set Feuille_Nomenclature = Classeur_Nomenclature.Worksheets("Nomenclature")
                        
                    Tableau_Nomenclature_nColonneAffaireSource = Feuille_Nomenclature.Rows(2).Find("Affaire source").Column
                    Tableau_Nomenclature_nColonneQuantite = Feuille_Nomenclature.Rows(2).Find("Quantit�").Column
                    Tableau_Nomenclature_nColonneD�signation = Feuille_Nomenclature.Rows(2).Find("D�signation").Column
                    Tableau_Nomenclature_nColonneR�f�rence = Feuille_Nomenclature.Rows(2).Find("R�f�rence").Column
                    Tableau_Nomenclature_nColonneDistributeur = Feuille_Nomenclature.Rows(2).Find("Distributeur").Column
                    Tableau_Nomenclature_nColonneR�fDistributeur = Feuille_Nomenclature.Rows(2).Find("R�f. Distributeur").Column
                    Tableau_Nomenclature_nColonneRemarque = Feuille_Nomenclature.Rows(2).Find("Remarques").Column
                    Tableau_Nomenclature_nColonneEtat = Feuille_Nomenclature.Rows(2).Find("Etat").Column
                    Tableau_Nomenclature_nColonneLocalisation = Feuille_Nomenclature.Rows(2).Find("Localisation").Column
                            
                    'La colonne Rep�re n'est pas toujours pr�sente dans les nomenclatures, d'o� la condition
                    If Not Feuille_Nomenclature.Rows(2).Find("Rep�re") Is Nothing Then
                        Tableau_Nomenclature_nColonneRep�re = Feuille_Nomenclature.Rows(2).Find("Rep�re").Column
                    Else
                        Tableau_Nomenclature_nColonneRep�re = 0
                    End If
                            
                    'Selon la nomenclature, il est �crit Fabriquant ou Fournisseur, d'o� les conditions
                    If Not Feuille_Nomenclature.Rows(2).Find("Fabriquant") Is Nothing Then
                        Tableau_Nomenclature_nColonneFabriquant = Feuille_Nomenclature.Rows(2).Find("Fabriquant").Column
                        Tableau_Nomenclature_nColonneFournisseur = 0
                    ElseIf Not Feuille_Nomenclature.Rows(2).Find("Fournisseur") Is Nothing Then
                        Tableau_Nomenclature_nColonneFournisseur = Feuille_Nomenclature.Rows(2).Find("Fournisseur").Column
                        Tableau_Nomenclature_nColonneFabriquant = 0
                    End If
                            
                    'La derni�re ligne de la nomenclature est calcul�e en remontant la colonne D�signation par le bas jusqu'� trouver une cellule non vide
                    Tableau_Nomenclature_DerLigne = Feuille_Nomenclature.Cells(Rows.Count, Tableau_Nomenclature_nColonneD�signation).End(xlUp).Row
                    Tableau_Nomenclature = Feuille_Nomenclature.Range(Feuille_Nomenclature.Cells(2, 1), Feuille_Nomenclature.Cells(Tableau_Nomenclature_DerLigne, Tableau_Nomenclature_nColonneLocalisation)).Value
                    
                    Tableau_Suivi_Indice = Tableau_Suivi_Indice + 1
                            
                            
                    For i = 2 To Tableau_Nomenclature_DerLigne - 1
                            
                        Quantite = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneQuantite)
                            
                        If (Quantite <> 0 Or IsEmpty(Quantite)) And Feuille_Nomenclature.Cells(i + 1, Tableau_Nomenclature_nColonneQuantite).Font.Strikethrough = False Then
                        'La condition v�rifie si la Quantit� de la ligne est diff�rente de 0 ou vide et que la ligne est non barr�e
                                
                            Tableau_Nomenclature_AffSource = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneAffaireSource)
                            Tableau_Nomenclature_Reference = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneR�f�rence)
                            Tableau_Nomenclature_Distributeur = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneDistributeur)
                            Tableau_Nomenclature_RefDistrib = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneR�fDistributeur)
                            Tableau_Nomenclature_Remarques = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneRemarque)
                            Tableau_Nomenclature_Etat = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneEtat)
                            Tableau_Nomenclature_D�signation = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneD�signation)
                                
                            If Tableau_Nomenclature_nColonneRep�re <> 0 Then
                            'Si la colonne Rep�re n'existe pas, le Rep�re est Empty
                                Tableau_Nomenclature_Rep�re = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneRep�re)
                            Else
                                Tableau_Nomenclature_Rep�re = Empty
                            End If
                                
                            If Tableau_Nomenclature_nColonneFabriquant = 0 Then
                            'Si la colonne Fournisseur n'existe pas, on renseigne le fabriquant, et inversement
                                Fournisseur = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneFournisseur)
                            ElseIf Tableau_Nomenclature_nColonneFournisseur = 0 Then
                                Fabriquant = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneFabriquant)
                            End If
                                
                            If UCase(Tableau_Nomenclature_Etat) = UCase("BPC") Or UCase(Tableau_Nomenclature_Etat) = UCase("Consult�") Or UCase(Tableau_Nomenclature_Etat) = UCase("Etude") Or IsEmpty(Tableau_Nomenclature_Etat) And Not IsEmpty(Tableau_Nomenclature_D�signation) And Tableau_Nomenclature_D�signation <> "" Then
                            'Si la ligne est en BPC, Consult�, Etude et la D�signation est non vide
                                
                                PlageR�sultante = Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Value
                                    
                                PlageR�sultante(1, 1) = Tableau_ListeProjetsAR_AffaireVoulue
                                PlageR�sultante(1, 2) = Tableau_Nomenclature_AffSource
                                PlageR�sultante(1, 3) = Tableau_Nomenclature_Rep�re
                                PlageR�sultante(1, 4) = Tableau_Nomenclature_D�signation
                                    
                                If Tableau_Nomenclature_nColonneFabriquant = 0 Then
                                    PlageR�sultante(1, 5) = Fournisseur
                                    
                                ElseIf Tableau_Nomenclature_nColonneFournisseur = 0 Then
                                    PlageR�sultante(1, 5) = Fabriquant
                                
                                End If
                                    
                                PlageR�sultante(1, 6) = Tableau_Nomenclature_Reference
                                PlageR�sultante(1, 7) = Tableau_Nomenclature_Distributeur
                                PlageR�sultante(1, 8) = Tableau_Nomenclature_RefDistrib
                                PlageR�sultante(1, 9) = Tableau_Nomenclature_Remarques
                                PlageR�sultante(1, 10) = Tableau_Nomenclature_Etat
                                    
                                Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Value = PlageR�sultante
                                Feuille_Suivi.Rows.AutoFit
                                Feuille_Suivi.Columns.Font.Size = 28
                                Feuille_Suivi.Columns.AutoFit
                                    
                                If UCase(Tableau_Nomenclature_Etat) = UCase("Etude") Then
                                'Attribution de la couleur violette pour les lignes en �tude
                                    Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Interior.Color = RGB(204, 102, 255)
                                End If
                                
                                If UCase(Tableau_Nomenclature_Etat) = UCase("Consult�") Then
                                'Attribution de la couleur jaune pour les lignes en consult�
                                    Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Interior.Color = RGB(255, 192, 0)
                                End If
                                    
                                Tableau_Suivi_Indice = Tableau_Suivi_Indice + 1
                                Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                                Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                            End If
                        End If
                        Next
                            
                    
                    'Bordure basse entre chaque ligne des nomenclatures
                    Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).Weight = xlThick
                    Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                    Classeur_Nomenclature.Close False
                    End If
                Next
                
            Tableau_Suivi_Indice = Tableau_Suivi_Indice + 2
                
            'Mise en forme et �criture du ruban pour le tableau Warnings AR (voir feuille Warnings AR)
            Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_DerColonne)).Copy
            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_Suivi_PremColonne).PasteSpecial (xlPasteAll)
            
            Tableau_Suivi_Indice = Tableau_Suivi_Indice + 1
    
                
            For i = 2 To Tableau_GDP06_DerLigne
                
                Tableau_GDP06_Texte = Tableau_GDP06(i, 8)
                Tableau_GDP06_Affaire = Tableau_GDP06(i, 5)
                Tableau_GDP06_DateAR = Tableau_GDP06(i, 15)
                Tableau_GDP06_Commentaire = Tableau_GDP06(i, 16)
                Tableau_GDP06_nCommande = Tableau_GDP06(i, 3)
                Tableau_GDP06_NomFournisseur = Tableau_GDP06(i, 4)
                Tableau_GDP06_Ref = Tableau_GDP06(i, 7)
                Tableau_GDP06_DateLiv = Tableau_GDP06(i, 14)
                Tableau_GDP06_Rubrique = Tableau_GDP06(i, 6)
                Tableau_GDP06_QteRestante = Tableau_GDP06(i, 18)
                Tableau_GDP06_Qte = Tableau_GDP06(i, 9)
                
        '------------------------------------------------------------------------------------------
        '6 cas temporels sont possibles dont 5 l�vent des warnings, ici on s'assure que le cas non warning n'est pas relev� d'o� le Not
        'Pour bien comprendre quelle est la condition il faut aller voir la feuille "Sch�ma warnings" et lire l'encadr� � c�t� du sch�ma d'explication des barres de progression
                
                If (((IsEmpty(Tableau_GDP06_QteRestante) Or Tableau_GDP06_QteRestante = "") And (IsEmpty(Tableau_GDP06_Commentaire) Or Tableau_GDP06_Commentaire = "") Or (Tableau_GDP06_QteRestante <> "0" And Not (IsEmpty(Tableau_GDP06_QteRestante) Or Tableau_GDP06_QteRestante = ""))) And Not (DateAjd <= Tableau_ListeProjetsAR_DateVoulue And ((DateAjd <= CDate(Tableau_GDP06_DateAR) And CDate(Tableau_GDP06_DateAR) <= Tableau_ListeProjetsAR_DateVoulue) Or (DateAjd <= CDate(Tableau_GDP06_DateLiv) And CDate(Tableau_GDP06_DateLiv) <= Tableau_ListeProjetsAR_DateVoulue)))) And (Not (IsEmpty(Tableau_GDP06_DateAR)) Or Not IsEmpty(Tableau_GDP06_DateLiv)) And Not IsEmpty(Tableau_GDP06_Affaire) And InStr(1, Tableau_GDP06_Affaire, Tableau_ListeProjetsAR_AffaireVoulue) And Tableau_GDP06_Rubrique = "ACHA" Then
                  'If (((IsEmpty(Tableau_GDP06_QteRestante) Or Tableau_GDP06_QteRestante = "") Or (Tableau_GDP06_QteRestante <> "0" And Not (IsEmpty(Tableau_GDP06_QteRestante) Or Tableau_GDP06_QteRestante = ""))) And Not (DateAjd <= Tableau_ListeProjetsAR_DateVoulue And ((DateAjd <= CDate(Tableau_GDP06_DateAR) And CDate(Tableau_GDP06_DateAR) <= Tableau_ListeProjetsAR_DateVoulue) Or (DateAjd <= CDate(Tableau_GDP06_DateLiv) And CDate(Tableau_GDP06_DateLiv) <= Tableau_ListeProjetsAR_DateVoulue)))) And (Not (IsEmpty(Tableau_GDP06_DateAR)) Or Not IsEmpty(Tableau_GDP06_DateLiv)) And Not IsEmpty(Tableau_GDP06_Affaire) And InStr(1, Tableau_GDP06_Affaire, Tableau_ListeProjetsAR_AffaireVoulue) And Tableau_GDP06_Rubrique = "ACHA" Then
                    
                    PlageR�sultante = Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value
                    
                    PlageR�sultante(1, 1) = Tableau_GDP06_Affaire
                    PlageR�sultante(1, 2) = Tableau_GDP06_nCommande
                    PlageR�sultante(1, 3) = Tableau_GDP06_NomFournisseur
                    PlageR�sultante(1, 4) = Tableau_GDP06_Ref
                    PlageR�sultante(1, 5) = Tableau_GDP06_Texte
                    PlageR�sultante(1, 6) = Tableau_GDP06_DateAR
                    PlageR�sultante(1, 7) = Tableau_GDP06_DateLiv
                    PlageR�sultante(1, 8) = Tableau_GDP06_Commentaire
                    PlageR�sultante(1, 9) = Tableau_GDP06_QteRestante
                    PlageR�sultante(1, 10) = Tableau_ListeProjetsAR_DateVoulue
                    
                    Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                    
        '------------------------------------------------------------------------------------------
        'On attribue les couleurs/barres de progression � chaque cas en priorisant AR sur livraison
                    
                    If IsEmpty(Tableau_GDP06_DateAR) Then
                    
                        If Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd >= CDate(Tableau_GDP06_DateLiv) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                            
                            PlageR�sultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateLiv))
                            PlageR�sultante(1, 12) = DateAjd - Tableau_ListeProjetsAR_DateVoulue
                            
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            
                            If CDate(Tableau_GDP06_DateLiv) >= Tableau_ListeProjetsAR_DateVoulue Then
                                WarningsAR_MaxProg = 1
                            Else
                                WarningsAR_MaxProg = Tableau_ListeProjetsAR_DateVoulue - CDate(Tableau_GDP06_DateLiv)
                            End If
                            
                            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                            Cr�er_barres Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_Suivi
        
                        ElseIf Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd <= CDate(Tableau_GDP06_DateLiv) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                            
                            PlageR�sultante(1, 12) = CDate(Tableau_GDP06_DateLiv) - Tableau_ListeProjetsAR_DateVoulue
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                            
                        ElseIf Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd >= CDate(Tableau_GDP06_DateLiv) Then
                            
                            PlageR�sultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateLiv))
                            
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            
                            WarningsAR_MaxProg = CDate(Tableau_GDP06_DateLiv) - Tableau_ListeProjetsAR_DateVoulue
                            Cr�er_barres Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_Suivi
                            
                        ElseIf Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd <= CDate(Tableau_GDP06_DateLiv) And Tableau_ListeProjetsAR_DateVoulue <= CDate(Tableau_GDP06_DateLiv) Then
                            
                            PlageR�sultante(1, 12) = CDate(Tableau_GDP06_DateLiv) - Tableau_ListeProjetsAR_DateVoulue
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                        
                        End If
                    
                    Else
                        If Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd >= CDate(Tableau_GDP06_DateAR) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                            
                            PlageR�sultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateAR))
                            PlageR�sultante(1, 12) = DateAjd - Tableau_ListeProjetsAR_DateVoulue
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            
                            If CDate(Tableau_GDP06_DateAR) >= Tableau_ListeProjetsAR_DateVoulue Then
                                WarningsAR_MaxProg = 1
                            Else
                                WarningsAR_MaxProg = Tableau_ListeProjetsAR_DateVoulue - CDate(Tableau_GDP06_DateAR)
                            End If
                            
                            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                            Cr�er_barres Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_Suivi
                            
                        ElseIf Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd <= CDate(Tableau_GDP06_DateAR) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                            
                            PlageR�sultante(1, 12) = CDate(Tableau_GDP06_DateAR) - Tableau_ListeProjetsAR_DateVoulue
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                            
                        ElseIf Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd >= CDate(Tableau_GDP06_DateAR) Then
                            
                            PlageR�sultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateAR))
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            WarningsAR_MaxProg = CDate(Tableau_GDP06_DateAR) - Tableau_ListeProjetsAR_DateVoulue
                            Cr�er_barres Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_Suivi
                            
                        ElseIf Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd <= CDate(Tableau_GDP06_DateAR) And Tableau_ListeProjetsAR_DateVoulue <= CDate(Tableau_GDP06_DateAR) Then
                            
                            PlageR�sultante(1, 12) = CDate(Tableau_GDP06_DateAR) - Tableau_ListeProjetsAR_DateVoulue
                            Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_PremColonne), Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageR�sultante
                            Feuille_Suivi.Cells(Tableau_Suivi_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                        
                        End If
                    End If
                
                Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_Suivi.Cells(Tableau_Suivi_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_Suivi.Cells(Tableau_Suivi_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                
                Tableau_Suivi_Indice = Tableau_Suivi_Indice + 1
                
                End If
            
            Next
        
        '------------------------------------------------------------------------------------------
        'Bordure par affaire
        Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_Suivi.Cells(Tableau_Suivi_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_Suivi.Cells(Tableau_Suivi_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).Weight = xlThick
        Feuille_Suivi.Range(Feuille_Suivi.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_Suivi.Cells(Tableau_Suivi_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                   
        Tableau_Suivi_Indice = Tableau_Suivi_Indice + 10 'espace de 10 lignes entre 2 Suivis
                
        End If
        
        Tableau_ListeProjetsAR_IndiceAV = Tableau_ListeProjetsAR_IndiceAV + 1 'On passe � l'affaire voulue suivante
    Loop


    Feuille_Suivi.Rows(1 & ":" & Tableau_Suivi_Indice).RowHeight = 25
    Feuille_Suivi.Rows(1 & ":" & Tableau_Suivi_Indice).Font.Size = 10
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    
    Feuille_Suivi.Range("A2").FormulaArray = "=NO.SEMAINE.ISO(AUJOURDHUI())"
    Classeur_GDP06.Close False
    
End Sub







