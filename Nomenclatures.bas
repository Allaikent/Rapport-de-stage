Attribute VB_Name = "Module1"
Sub Nomenclatures()
Attribute Nomenclatures.VB_ProcData.VB_Invoke_Func = " \n14"

'On Error GoTo err
    
    Dim Classeur_GDP04 As Workbook
    Dim Feuille_ListeProjetsAR As Worksheet
    Dim Feuille_Nomenclatures As Worksheet
    Dim Classeur_Nomenclature As Workbook
    Dim Feuille_Nomenclature As Worksheet
    
    Application.ScreenUpdating = False 'Empêcher le changement d'affichage
    Application.DisplayAlerts = False 'Empêcher les pop-ups
    Application.Calculation = xlManual
    
    Set Classeur_GDP04 = ActiveWorkbook
    
    Set Feuille_ListeProjetsAR = Classeur_GDP04.Sheets("Liste projets AR")
    Tableau_ListeProjetsAR_PremLigne = Feuille_ListeProjetsAR.Range("ListeProjetsAR_ET").Rows(1).Row
    Tableau_ListeProjetsAR_nColonneMeca = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature Méca").Column
    Tableau_ListeProjetsAR_DerLigne = Feuille_ListeProjetsAR.Cells(Rows.Count, Tableau_ListeProjetsAR_nColonneMeca).End(xlUp).Row 'La dernière ligne à traiter de la feuille "Liste projets AR" est calculée en remontant la colonne "Nomenclature méca" jusqu'à trouver une cellule non vide
    Tableau_ListeProjetsAR_nColonneElec = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature Elec").Column
    Tableau_ListeProjetsAR_nColonneAutre1 = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature 3").Column
    Tableau_ListeProjetsAR_nColonneAutre2 = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Nomenclature 4").Column
    Tableau_ListeProjetsAR_nColonneSelectionNom = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Select Nom").Column
    Tableau_ListeProjetsAR_nColonneDV = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Date de besoin").Column
    Tableau_ListeProjetsAR_nColonneAV = Feuille_ListeProjetsAR.Rows(Tableau_ListeProjetsAR_PremLigne).Find("Numéro affaire").Column
    
    Set Feuille_Nomenclatures = Classeur_GDP04.Sheets("Nomenclatures")
    Tableau_Nomenclatures_PremLigne = Feuille_Nomenclatures.Range("Nomenclatures_ET").Rows(1).Row
    Tableau_Nomenclatures_Indice = Tableau_Nomenclatures_PremLigne + 1
    Tableau_Nomenclatures_PremColonne = Feuille_Nomenclatures.Range("Nomenclatures_ET").Columns(1).Column
    Tableau_Nomenclatures_DerColonne = Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_PremLigne, Feuille_Nomenclatures.Columns.Count).End(xlToLeft).Column
    Feuille_Nomenclatures.Rows(Tableau_Nomenclatures_Indice & ":" & Feuille_Nomenclatures.Rows.Count).Delete
    
    
    For k = Tableau_ListeProjetsAR_PremLigne + 1 To Tableau_ListeProjetsAR_DerLigne
    'La boucle parcoure tout le tableau ListeProjetsAR

        If Not IsEmpty(Feuille_ListeProjetsAR.Cells(k, Tableau_ListeProjetsAR_nColonneSelectionNom)) And Feuille_ListeProjetsAR.Cells(k, Tableau_ListeProjetsAR_nColonneSelectionNom) <> "" Then
        'La condition vérifie si la cellule correspondante au numéro d'affaire dans la colonne Sélection n'est pas vide
            
            ListeProjetsAR_AffaireVoulue = Feuille_ListeProjetsAR.Cells(k, Tableau_ListeProjetsAR_nColonneAV)
            ListeProjetsAR_DateVoulue = Feuille_ListeProjetsAR.Cells(k, Tableau_ListeProjetsAR_nColonneDV)
            
            For j = Tableau_ListeProjetsAR_nColonneMeca To Tableau_ListeProjetsAR_nColonneAutre2
            'La boucle parcoure les colonnes Nomenclatures X (meca, elec, autre1, autre 2)
                
                If Not IsEmpty(Feuille_ListeProjetsAR.Cells(k, j)) And Feuille_ListeProjetsAR.Cells(k, j) <> "" Then
                        'Si la cellule Nomenclature X est non vide, la nomenclature est ouverte et fouillée
                        
                        Feuille_ListeProjetsAR.Cells(k, j).Hyperlinks(1).Follow
                        Set Classeur_Nomenclature = ActiveWorkbook
                        Set Feuille_Nomenclature = Classeur_Nomenclature.Worksheets("Nomenclature")
    
                        Tableau_Nomenclature_nColonneAffaireSource = Feuille_Nomenclature.Rows(2).Find("Affaire source").Column
                        Tableau_Nomenclature_nColonneQuantite = Feuille_Nomenclature.Rows(2).Find("Quantité").Column
                        Tableau_Nomenclature_nColonneDésignation = Feuille_Nomenclature.Rows(2).Find("Désignation").Column
                        Tableau_Nomenclature_nColonneReference = Feuille_Nomenclature.Rows(2).Find("Référence").Column
                        Tableau_Nomenclature_nColonneDistributeur = Feuille_Nomenclature.Rows(2).Find("Distributeur").Column
                        Tableau_Nomenclature_nColonneReferenceDistributeur = Feuille_Nomenclature.Rows(2).Find("Réf. Distributeur").Column
                        Tableau_Nomenclature_nColonneRemarque = Feuille_Nomenclature.Rows(2).Find("Remarques").Column
                        Tableau_Nomenclature_nColonneEtat = Feuille_Nomenclature.Rows(2).Find("Etat").Column
                        Tableau_Nomenclature_nColonneLocalisation = Feuille_Nomenclature.Rows(2).Find("Localisation").Column
                        
                        'La colonne Repère n'est pas toujours présente dans les nomenclatures, d'où la condition
                        If Not Feuille_Nomenclature.Rows(2).Find("Repère") Is Nothing Then
                            Tableau_Nomenclature_nColonneRepère = Feuille_Nomenclature.Rows(2).Find("Repère").Column
                        Else
                            Tableau_Nomenclature_nColonneRepère = 0
                        End If
                        
                        'Selon la nomenclature, il est écrit Fabriquant ou Fournisseur, d'où les conditions
                        If Not Feuille_Nomenclature.Rows(2).Find("Fabriquant") Is Nothing Then
                            Tableau_Nomenclature_nColonneFabriquant = Feuille_Nomenclature.Rows(2).Find("Fabriquant").Column
                            Tableau_Nomenclature_nColonneFournisseur = 0
                        ElseIf Not Feuille_Nomenclature.Rows(2).Find("Fournisseur") Is Nothing Then
                            Tableau_Nomenclature_nColonneFournisseur = Feuille_Nomenclature.Rows(2).Find("Fournisseur").Column
                            Tableau_Nomenclature_nColonneFabriquant = 0
                        End If
                        
                        'La dernière ligne de la nomenclature est calculée en remontant la colonne Désignation par le bas jusqu'à trouver une cellule non vide
                        Tableau_Nomenclature_DerLigne = Feuille_Nomenclature.Cells(Rows.Count, Tableau_Nomenclature_nColonneDésignation).End(xlUp).Row
                        Tableau_Nomenclature = Feuille_Nomenclature.Range(Feuille_Nomenclature.Cells(2, 1), Feuille_Nomenclature.Cells(Tableau_Nomenclature_DerLigne, Tableau_Nomenclature_nColonneLocalisation)).Value
                        
                        
                        For i = 2 To Tableau_Nomenclature_DerLigne - 1
                        
                        Quantite = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneQuantite)
                        
                        If (Quantite <> 0 Or IsEmpty(Quantite)) And Feuille_Nomenclature.Cells(i + 1, Tableau_Nomenclature_nColonneQuantite).Font.Strikethrough = False Then
                        'La condition vérifie si la Quantité de la ligne est différente de 0 ou vide et que la ligne est non barrée
                            
                            AffaireSource = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneAffaireSource)
                            Reference = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneReference)
                            Distributeur = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneDistributeur)
                            RéférenceDistributeur = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneReferenceDistributeur)
                            Remarques = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneRemarque)
                            Etat = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneEtat)
                            Désignation = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneDésignation)
                            
                            If Tableau_Nomenclature_nColonneRepère <> 0 Then
                            'Si la colonne Repère n'existe pas, le Repère est Empty
                                Repère = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneRepère)
                            Else
                                Repère = Empty
                            End If
                            
                            If Tableau_Nomenclature_nColonneFabriquant = 0 Then
                            'Si la colonne Fournisseur n'existe pas, on renseigne le fabriquant, et inversement
                                Fournisseur = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneFournisseur)
                            ElseIf Tableau_Nomenclature_nColonneFournisseur = 0 Then
                                Fabriquant = Tableau_Nomenclature(i, Tableau_Nomenclature_nColonneFabriquant)
                            End If
                            
                            If UCase(Etat) = UCase("BPC") Or UCase(Etat) = UCase("Consulté") Or UCase(Etat) = UCase("Etude") Or IsEmpty(Etat) And Not IsEmpty(Désignation) And Désignation <> "" Then
                            'Si la ligne est en BPC, Consulté, Etude et la Désignation est non vide
                            
                                PlageRésultante = Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Value
                                
                                PlageRésultante(1, 1) = ListeProjetsAR_AffaireVoulue
                                PlageRésultante(1, 2) = AffaireSource
                                PlageRésultante(1, 3) = Repère
                                PlageRésultante(1, 4) = Désignation
                                
                                If Tableau_Nomenclature_nColonneFabriquant = 0 Then
                                PlageRésultante(1, 5) = Fournisseur
                                
                                ElseIf Tableau_Nomenclature_nColonneFournisseur = 0 Then
                                PlageRésultante(1, 5) = Fabriquant
                                End If
                                
                                PlageRésultante(1, 6) = Reference
                                PlageRésultante(1, 7) = Distributeur
                                PlageRésultante(1, 8) = RéférenceDistributeur
                                PlageRésultante(1, 9) = Remarques
                                PlageRésultante(1, 10) = Etat
                                
                                Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Value = PlageRésultante
                                Feuille_Nomenclatures.Rows.AutoFit
                                Feuille_Nomenclatures.Columns.Font.Size = 28
                                Feuille_Nomenclatures.Columns.AutoFit
                                
                                If UCase(Etat) = UCase("Etude") Then
                                'Attribution de la couleur violette pour les lignes en étude
                                    Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Interior.Color = RGB(204, 102, 255)
                                End If
                                If UCase(Etat) = UCase("Consulté") Then
                                'Attribution de la couleur jaune pour les lignes en consulté
                                    Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Interior.Color = RGB(255, 192, 0)
                                End If
                                
                                Tableau_Nomenclatures_Indice = Tableau_Nomenclatures_Indice + 1
                                Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                                Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                            End If
                        End If
                        Next
                        
                        '------------------------------------------------------------------------------------------
                        'Bordure basse entre chaque ligne des nomenclatures
                        
                        Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).Weight = xlThick
                        Feuille_Nomenclatures.Range(Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_PremColonne), Feuille_Nomenclatures.Cells(Tableau_Nomenclatures_Indice, Tableau_Nomenclatures_DerColonne)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                        Classeur_Nomenclature.Close False
                        End If
                    Next
        End If
    Next

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
