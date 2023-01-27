Attribute VB_Name = "Module3"
Sub Créer_barres(ConcernedRange As Range, WarningsAR_MaxProg As Double, Feuille_WarningsAR As Worksheet)
    
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


Sub MàJ_Warnings_V2()
 
    'On Error GoTo err
    
    Dim Classeur_GDP06 As Workbook
    Dim Feuille_Feuil1 As Worksheet
    
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
    Dim Tableau_GDP06_DerLigne As Long
    Dim Tableau_GDP06_DerColonne As Long
    Dim objConnection As WorkbookConnection
    Dim Chemin_GDP06 As String

    Dim Classeur_GDP04 As Workbook
    Dim Feuille_WarningsAR As Worksheet
    
    Dim WarningsAR_MaxProg As Double
    Dim Tableau_WarningsAR_nColonneRR As Long
    Dim Tableau_WarningsAR_nColonneRP As Long
    Dim Tableau_WarningsAR_Indice As Long
    Dim Tableau_WarningsAR_nColonneAffaire As Long
    Dim PlageRésultante As Variant
    Dim DateAjd As Date
    
    Dim Feuille_ListeProjetsAR As Worksheet
    
    Dim Tableau_ListeProjetsAR_AffaireVoulue As String
    Dim Tableau_ListeProjetsAR_Sélection As String
    Dim Tableau_ListeProjetsAR_DateVoulue As Variant
    Dim Tableau_ListeProjetsAR_IndiceAV As Long
    Dim Tableau_ListeProjetsAR_nColonneDV As Long
    Dim Tableau_ListeProjetsAR_nColonneAV As Long
    Dim Tableau_ListeProjetsAR_nColonneSelect As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    Chemin_GDP06 = "T:\ZZ_Planning\CDP\GDP_006_A_Extract CMD EVERWIN (base données).xlsx"
    DateAjd = Date
    
    Set Classeur_GDP04 = ActiveWorkbook
    
    Set Feuille_WarningsAR = Classeur_GDP04.Sheets("Warnings AR")
    Tableau_WarningsAR_PremLigne = 2 'la ligne où commence le tableau des warnings pour l'affaire voulue
    Tableau_WarningsAR_PremColonne = 2
    Tableau_WarningsAR_DerColonne = Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Feuille_WarningsAR.Columns.Count).End(xlToLeft).Column
    Tableau_WarningsAR_Indice = Tableau_WarningsAR_PremLigne + 1
    Tableau_WarningsAR_nColonneRR = Feuille_WarningsAR.Rows(Tableau_WarningsAR_PremLigne).Find("Retard de réception Symétrie (en jours)").Column
    Tableau_WarningsAR_nColonneRP = Feuille_WarningsAR.Rows(Tableau_WarningsAR_PremLigne).Find("Retard projet (en jours)").Column
    Tableau_WarningsAR_nColonneAffaire = Feuille_WarningsAR.Rows(Tableau_WarningsAR_PremLigne).Find("Affaire").Column
    WarningsAR_MaxProg = 0 'valeur WarningsAR_MaxProgimale de la barre de progression
    
    Set Feuille_ListeProjetsAR = Classeur_GDP04.Sheets("Liste projets AR")
    Tableau_ListeProjetsAR_PremLigne = 4
    Tableau_ListeProjetsAR_IndiceAV = Tableau_ListeProjetsAR_PremLigne + 1  'la ligne actuelle dans la colonne Numero affaire de la feuille Liste projets AR
    Tableau_ListeProjetsAR_nColonneDV = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Date de besoin").Column
    Tableau_ListeProjetsAR_nColonneAV = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Numéro affaire").Column
    Tableau_ListeProjetsAR_nColonneSelect = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Select Warnings").Column
    Tableau_ListeProjetsAR_IndiceAV = Tableau_ListeProjetsAR_PremLigne + 1
    
    Feuille_WarningsAR.Rows(Tableau_WarningsAR_Indice & ":" & Feuille_WarningsAR.Rows.Count).Delete
    Feuille_WarningsAR.Rows(Tableau_WarningsAR_Indice & ":" & Feuille_WarningsAR.Rows.Count).Interior.Color = RGB(255, 255, 255)
    
    '------------------------------------------------------------------------------------------
    'Actualisation de GDP06
    
    If MsgBox("Mettre à jour la BDD Everwin ?", vbYesNo) = vbYes Then
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
    
    '------------------------------------------------------------------------------------------
    'En VBA pour économiser un grand temps de calcul en lecture on lit toutes les données d'un grand tableau (ici le tableau de la Feuille1 de l'extract CMD) une seule fois grâce à la commande Range.Value
    
    Tableau_GDP06 = Feuille_Feuil1.Range(Cells(1, 1), Cells(Tableau_GDP06_DerLigne, Tableau_GDP06_DerColonne)).Value
    
    '------------------------------------------------------------------------------------------
    'La première boucle s'arrête quand une cellule de la colonne "Numero affaire" du tableau présent dans la feuille "Liste projets AR" est vide
    
    Do While (Not IsEmpty(Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneAV)))
    
        Tableau_ListeProjetsAR_AffaireVoulue = Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneAV)
        Tableau_ListeProjetsAR_DateVoulue = Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneDV)
        Tableau_ListeProjetsAR_Sélection = Feuille_ListeProjetsAR.Cells(Tableau_ListeProjetsAR_IndiceAV, Tableau_ListeProjetsAR_nColonneSelect)
        
        If (IsEmpty(Tableau_ListeProjetsAR_Sélection) Or Tableau_ListeProjetsAR_Sélection = "") Then
            GoTo FinDoWhile
        End If
        
    '------------------------------------------------------------------------------------------
    'Du début à la fin de l'extract CMD
        
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
    '6 cas temporels sont possibles dont 5 lèvent des warnings, ici on s'assure que le cas non warning n'est pas relevé d'où le Not
    'Pour bien comprendre quelle est la condition il faut aller voir la feuille "Schéma warnings" et lire l'encadré à côté du schéma d'explication des barres de progression
            
            If (((IsEmpty(Tableau_GDP06_QteRestante) Or Tableau_GDP06_QteRestante = "") And (IsEmpty(Tableau_GDP06_Commentaire) Or Tableau_GDP06_Commentaire = "") Or (Tableau_GDP06_QteRestante <> "0" And Not (IsEmpty(Tableau_GDP06_QteRestante) Or Tableau_GDP06_QteRestante = ""))) And Not (DateAjd <= Tableau_ListeProjetsAR_DateVoulue And ((DateAjd <= CDate(Tableau_GDP06_DateAR) And CDate(Tableau_GDP06_DateAR) <= Tableau_ListeProjetsAR_DateVoulue) Or (DateAjd <= CDate(Tableau_GDP06_DateLiv) And CDate(Tableau_GDP06_DateLiv) <= Tableau_ListeProjetsAR_DateVoulue)))) And (Not (IsEmpty(Tableau_GDP06_DateAR)) Or Not IsEmpty(Tableau_GDP06_DateLiv)) And Not IsEmpty(Tableau_GDP06_Affaire) And InStr(1, Tableau_GDP06_Affaire, Tableau_ListeProjetsAR_AffaireVoulue) And Tableau_GDP06_Rubrique = "ACHA" Then
                
                PlageRésultante = Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value
                
                PlageRésultante(1, 1) = Tableau_GDP06_Affaire
                PlageRésultante(1, 2) = Tableau_GDP06_nCommande
                PlageRésultante(1, 3) = Tableau_GDP06_NomFournisseur
                PlageRésultante(1, 4) = Tableau_GDP06_Ref
                PlageRésultante(1, 5) = Tableau_GDP06_Texte
                PlageRésultante(1, 6) = Tableau_GDP06_DateAR
                PlageRésultante(1, 7) = Tableau_GDP06_DateLiv
                PlageRésultante(1, 8) = Tableau_GDP06_Commentaire
                PlageRésultante(1, 9) = Tableau_GDP06_QteRestante
                PlageRésultante(1, 10) = Tableau_ListeProjetsAR_DateVoulue
                
                Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                
    '------------------------------------------------------------------------------------------
    'On attribue les couleurs/barres de progression à chaque cas en priorisant AR sur livraison
                
                If IsEmpty(Tableau_GDP06_DateAR) Then
                
                    If Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd >= CDate(Tableau_GDP06_DateLiv) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                        
                        PlageRésultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateLiv))
                        PlageRésultante(1, 12) = DateAjd - Tableau_ListeProjetsAR_DateVoulue
                        
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        
                        If CDate(Tableau_GDP06_DateLiv) >= Tableau_ListeProjetsAR_DateVoulue Then
                            WarningsAR_MaxProg = 1
                        Else
                            WarningsAR_MaxProg = Tableau_ListeProjetsAR_DateVoulue - CDate(Tableau_GDP06_DateLiv)
                        End If
                        
                        Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                        Créer_barres Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_WarningsAR
    
                    ElseIf Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd <= CDate(Tableau_GDP06_DateLiv) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                        
                        PlageRésultante(1, 12) = CDate(Tableau_GDP06_DateLiv) - Tableau_ListeProjetsAR_DateVoulue
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                        
                    ElseIf Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd >= CDate(Tableau_GDP06_DateLiv) Then
                        
                        PlageRésultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateLiv))
                        
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        
                        WarningsAR_MaxProg = CDate(Tableau_GDP06_DateLiv) - Tableau_ListeProjetsAR_DateVoulue
                        Créer_barres Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_WarningsAR
                        
                    ElseIf Not IsEmpty(Tableau_GDP06_DateLiv) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd <= CDate(Tableau_GDP06_DateLiv) And Tableau_ListeProjetsAR_DateVoulue <= CDate(Tableau_GDP06_DateLiv) Then
                        
                        PlageRésultante(1, 12) = CDate(Tableau_GDP06_DateLiv) - Tableau_ListeProjetsAR_DateVoulue
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                    
                    End If
                
                Else
                    If Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd >= CDate(Tableau_GDP06_DateAR) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                        
                        PlageRésultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateAR))
                        PlageRésultante(1, 12) = DateAjd - Tableau_ListeProjetsAR_DateVoulue
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        
                        If CDate(Tableau_GDP06_DateAR) >= Tableau_ListeProjetsAR_DateVoulue Then
                            WarningsAR_MaxProg = 1
                        Else
                            WarningsAR_MaxProg = Tableau_ListeProjetsAR_DateVoulue - CDate(Tableau_GDP06_DateAR)
                        End If
                        
                        Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                        Créer_barres Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_WarningsAR
                        
                    ElseIf Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd <= CDate(Tableau_GDP06_DateAR) And DateAjd >= Tableau_ListeProjetsAR_DateVoulue Then
                        
                        PlageRésultante(1, 12) = CDate(Tableau_GDP06_DateAR) - Tableau_ListeProjetsAR_DateVoulue
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                        
                    ElseIf Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd >= CDate(Tableau_GDP06_DateAR) Then
                        
                        PlageRésultante(1, 11) = (DateAjd - CDate(Tableau_GDP06_DateAR))
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        WarningsAR_MaxProg = CDate(Tableau_GDP06_DateAR) - Tableau_ListeProjetsAR_DateVoulue
                        Créer_barres Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRR), WarningsAR_MaxProg, Feuille_WarningsAR
                        
                    ElseIf Not IsEmpty(Tableau_GDP06_DateAR) And DateAjd <= Tableau_ListeProjetsAR_DateVoulue And DateAjd <= CDate(Tableau_GDP06_DateAR) And Tableau_ListeProjetsAR_DateVoulue <= CDate(Tableau_GDP06_DateAR) Then
                        
                        PlageRésultante(1, 12) = CDate(Tableau_GDP06_DateAR) - Tableau_ListeProjetsAR_DateVoulue
                        Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_PremColonne), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_DerColonne)).Value = PlageRésultante
                        Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice, Tableau_WarningsAR_nColonneRP).Interior.Color = RGB(255, 242, 204)
                    
                    End If
                End If
            
            Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
            
            Tableau_WarningsAR_Indice = Tableau_WarningsAR_Indice + 1
            
            End If
        
        Next

    '------------------------------------------------------------------------------------------
    'Bordure par affaire
    
    Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).Weight = xlThick
    Feuille_WarningsAR.Range(Feuille_WarningsAR.Cells(Tableau_WarningsAR_PremLigne, Tableau_WarningsAR_nColonneAffaire), Feuille_WarningsAR.Cells(Tableau_WarningsAR_Indice - 1, Tableau_WarningsAR_nColonneRP)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
       
FinDoWhile:
    
    '------------------------------------------------------------------------------------------
    'Ici on passe à la ligne suivante de la colonne "Affaire voulue"
    
    Tableau_ListeProjetsAR_IndiceAV = Tableau_ListeProjetsAR_IndiceAV + 1
    Loop
    
    Classeur_GDP06.Close False
    
    'fin:
        'Exit Sub
    
    'err:
        'Select Case err.Number
            'Case 13: MsgBox "Pas de MàJ effectuée"
            'Case 1004: MsgBox "La macro n'arrive pas à accéder au fichier extrait des CMD extraites d'Everwin"
            'Case Else: MsgBox "Erreur inconnue"
        'End Select
    
    'Resume fin
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
End Sub








