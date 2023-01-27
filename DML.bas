Attribute VB_Name = "Module1"
Sub DML_DepuisUneNomenclature()
    
    On Error GoTo err
    
    Dim Affaire As Variant
    Dim Masse As Variant
    Dim Indice As Integer
    Dim Désignation As Variant
    Dim Matériau As Variant
    Dim Traitement As Variant
    Dim CopieTableau As Variant
    Dim PlageRésultante As Variant
    Dim Quantité As Variant
    Dim IndiceMasse As Long
    Dim IndicePourcentageMasse As Long
    Dim DerLigne As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'Calcul de la première ligne du CopieTableau à traiter
    
    For Each ligne In Cells.Rows
        If Cells(ligne.Row, Columns.Count).End(xlToLeft).Column <> 1 Then
            PremLigne = ligne.Row
            Exit For
        End If
    Next
           
    'Calcul de la première colonne du CopieTableau à traiter
    
    For Each colonne In Cells.Columns
        If Cells(Rows.Count, colonne.Column).End(xlUp).Row <> 1 Then
            PremColonne = colonne.Column
            Exit For
        End If
    Next
        
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    Indice = PremLigne + 1 'L'indice du CopieTableau commence après l'en-tête
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'Chaque numéro de colonne est recherché par son nom en y ajoutant l'offset
    
    Offset = ActiveSheet.Cells(PremLigne, PremColonne).Column
    
    IndiceAffaire = ActiveSheet.Rows(PremLigne).Find("Affaire").Column - Offset + 1
    IndiceRepère = ActiveSheet.Rows(PremLigne).Find("Repère").Column - Offset + 1
    IndiceDésignation = ActiveSheet.Rows(PremLigne).Find("Désignation").Column - Offset + 1
    IndiceMatériau = ActiveSheet.Rows(PremLigne).Find("Matériau").Column - Offset + 1
    IndiceTraitement = ActiveSheet.Rows(PremLigne).Find("Traitement").Column - Offset + 1
    IndiceMasse = ActiveSheet.Rows(PremLigne).Find("Masse").Column - Offset + 1
    IndiceRévision = ActiveSheet.Rows(PremLigne).Find("Révision").Column - Offset + 1
    IndicePourcentageMasse = ActiveSheet.Rows(PremLigne).Find("Configuration").Column - Offset + 1
    IndiceQuantité = ActiveSheet.Rows(PremLigne).Find("Compte de référence").Column - Offset + 1
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'Délimitation du Tableau à traiter
    
    DerLigne = Cells(Rows.Count, Cells(PremLigne, IndiceQuantité + Offset - 1).Column).End(xlUp).Row 'La dernière ligne du CopieTableau à traiter est calculée en remontant la colonne Compte de référence jusqu'à trouver une valeur non vide
    DerColonne = Cells(Cells(PremLigne, IndiceQuantité + Offset - 1).Row, Columns.Count).End(xlToLeft).Column 'La dernière colonne du CopieTableau à traiter est calculée en allant vers la gauche de l'en-tête jusqu'à trouver une valeur non vide
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'Définition de la plage de travail
    
    CopieTableau = Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Value
    Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Offset(1).ClearContents 'Tout l'ancien CopieTableau contenu dans la feuille est supprimé sauf l'en-tête
    
    '---------------------------------------------------------------------------------------------------------------------------------'

    For i = 2 To DerLigne - PremLigne + 1
    'Boucle de la ligne après l'en-tête jusqu'à la fin du CopieTableau
        
        Masse = 0#
        Désignation = Null
        Traitement = Null
        Matériau = CopieTableau(i, IndiceMatériau)
        Affaire = CopieTableau(i, IndiceAffaire)
        Traitement = CopieTableau(i, IndiceTraitement)
        
        If (Not IsEmpty(Affaire) Or Affaire <> "") Or (Not IsEmpty(Repère) Or Repère <> "") Then
        'Si l'affaire et/ou le repère de la pièce est vide
            end_char = 0
            For k = 2 To DerLigne - PremLigne + 1
            'Boucle de la ligne après l'en-tête jusqu'à la fin du CopieTableau
                
                If CopieTableau(k, IndiceMatériau) = Matériau And CopieTableau(k, IndiceTraitement) = Traitement And (Not IsEmpty(CopieTableau(k, IndiceDésignation)) Or CopieTableau(k, IndiceDésignation) <> "") Then
                'Si un autre composant possède le même couple (Matériau, Traitement), les désignations sont concaténées et les masses ajoutées
                    
                    end_char = 1
                    Masse = Masse + CopieTableau(k, IndiceQuantité) * CopieTableau(k, IndiceMasse)
                    If CopieTableau(k, IndiceQuantité) = 1 Then
                        Désignation = Désignation & CopieTableau(k, IndiceDésignation) & "," & Chr(10)
                    Else
                        Désignation = Désignation & CopieTableau(k, IndiceQuantité) & "x " & CopieTableau(k, IndiceDésignation) & "," & Chr(10)
                    End If
                End If
            Next
            
            If end_char = 1 Then
            'Si il y a eu au moins une correspondance entre deux composants, l'espace et la virgule à la fin de la nouvelle désignation sont supprimées
                Désignation = Left(Désignation, Len(Désignation) - 2)
            End If
            
            PlageRésultante = Range(Cells(Indice, PremColonne), Cells(Indice, DerColonne)).Value
                
            PlageRésultante(1, IndiceAffaire) = "XXX"
            PlageRésultante(1, IndiceRepère) = "XXX"
            PlageRésultante(1, IndiceDésignation) = Désignation
            PlageRésultante(1, IndiceMatériau) = Matériau
            PlageRésultante(1, IndiceTraitement) = Traitement
            PlageRésultante(1, IndiceMasse) = Masse
            PlageRésultante(1, IndiceRévision) = "XXX"
            PlageRésultante(1, IndiceQuantité) = 1
            
            Range(Cells(Indice, PremColonne), Cells(Indice, DerColonne)).Value = PlageRésultante
            
            Indice = Indice + 1
            
        End If
    Next
    
    DerLigne = Cells(Rows.Count, Cells(PremLigne, IndiceQuantité + Offset - 1).Column).End(xlUp).Row
    
    ActiveSheet.Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).RemoveDuplicates Columns:=Array(IndiceAffaire, IndiceDésignation, IndiceMatériau, IndiceTraitement, IndiceMasse), Header _
        :=xlNo 'L'algorithme précédent génère des doublons pour chaque couple (Matériau, Traitement) qui sont supprimés
        
    Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Rows.AutoFit
    Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Columns.AutoFit
    
    DerLigne = Cells(Rows.Count, Cells(PremLigne, IndiceQuantité + Offset - 1).Column).End(xlUp).Row
    
    For l = PremLigne + 1 To DerLigne
        
        Masse_matériau = Cells(l, IndiceMasse)
        Masse_totale = Application.WorksheetFunction.Sum(Range(Cells(PremLigne + 1, IndiceMasse), Cells(DerLigne, IndiceMasse)))
        If Masse_totale <> 0 Then
            Cells(l, IndicePourcentageMasse) = Round(CDbl(Masse_matériau) / Masse_totale * 100, 2)
        End If
        
    Next
    
    Rows(PremLigne & ":" & DerLigne).Sort Key1:=Range(Cells(PremLigne, IndiceMasse), Cells(DerLigne, IndiceMasse)), order1:=xlDescending
    Cells(DerLigne + 1, DerColonne + 1) = "Masse totale :" & " " & Masse_totale

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    
fin:
        Exit Sub
    
err:
        Select Case err.Number
            Case 13: MsgBox "Erreur13. Impossible de réaliser les calculs. Vérifier qu'il n'y a pas de texte dans les colonnes ''Masse'' et ''Compte de référence''"
            Case Else: MsgBox "Erreur inconnue"
        End Select
        
    Resume fin

End Sub


