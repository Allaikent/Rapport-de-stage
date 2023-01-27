Attribute VB_Name = "Module1"
Sub DML_DepuisUneNomenclature()
    
    On Error GoTo err
    
    Dim Affaire As Variant
    Dim Masse As Variant
    Dim Indice As Integer
    Dim D�signation As Variant
    Dim Mat�riau As Variant
    Dim Traitement As Variant
    Dim CopieTableau As Variant
    Dim PlageR�sultante As Variant
    Dim Quantit� As Variant
    Dim IndiceMasse As Long
    Dim IndicePourcentageMasse As Long
    Dim DerLigne As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'Calcul de la premi�re ligne du CopieTableau � traiter
    
    For Each ligne In Cells.Rows
        If Cells(ligne.Row, Columns.Count).End(xlToLeft).Column <> 1 Then
            PremLigne = ligne.Row
            Exit For
        End If
    Next
           
    'Calcul de la premi�re colonne du CopieTableau � traiter
    
    For Each colonne In Cells.Columns
        If Cells(Rows.Count, colonne.Column).End(xlUp).Row <> 1 Then
            PremColonne = colonne.Column
            Exit For
        End If
    Next
        
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    Indice = PremLigne + 1 'L'indice du CopieTableau commence apr�s l'en-t�te
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'Chaque num�ro de colonne est recherch� par son nom en y ajoutant l'offset
    
    Offset = ActiveSheet.Cells(PremLigne, PremColonne).Column
    
    IndiceAffaire = ActiveSheet.Rows(PremLigne).Find("Affaire").Column - Offset + 1
    IndiceRep�re = ActiveSheet.Rows(PremLigne).Find("Rep�re").Column - Offset + 1
    IndiceD�signation = ActiveSheet.Rows(PremLigne).Find("D�signation").Column - Offset + 1
    IndiceMat�riau = ActiveSheet.Rows(PremLigne).Find("Mat�riau").Column - Offset + 1
    IndiceTraitement = ActiveSheet.Rows(PremLigne).Find("Traitement").Column - Offset + 1
    IndiceMasse = ActiveSheet.Rows(PremLigne).Find("Masse").Column - Offset + 1
    IndiceR�vision = ActiveSheet.Rows(PremLigne).Find("R�vision").Column - Offset + 1
    IndicePourcentageMasse = ActiveSheet.Rows(PremLigne).Find("Configuration").Column - Offset + 1
    IndiceQuantit� = ActiveSheet.Rows(PremLigne).Find("Compte de r�f�rence").Column - Offset + 1
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'D�limitation du Tableau � traiter
    
    DerLigne = Cells(Rows.Count, Cells(PremLigne, IndiceQuantit� + Offset - 1).Column).End(xlUp).Row 'La derni�re ligne du CopieTableau � traiter est calcul�e en remontant la colonne Compte de r�f�rence jusqu'� trouver une valeur non vide
    DerColonne = Cells(Cells(PremLigne, IndiceQuantit� + Offset - 1).Row, Columns.Count).End(xlToLeft).Column 'La derni�re colonne du CopieTableau � traiter est calcul�e en allant vers la gauche de l'en-t�te jusqu'� trouver une valeur non vide
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    'D�finition de la plage de travail
    
    CopieTableau = Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Value
    Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Offset(1).ClearContents 'Tout l'ancien CopieTableau contenu dans la feuille est supprim� sauf l'en-t�te
    
    '---------------------------------------------------------------------------------------------------------------------------------'

    For i = 2 To DerLigne - PremLigne + 1
    'Boucle de la ligne apr�s l'en-t�te jusqu'� la fin du CopieTableau
        
        Masse = 0#
        D�signation = Null
        Traitement = Null
        Mat�riau = CopieTableau(i, IndiceMat�riau)
        Affaire = CopieTableau(i, IndiceAffaire)
        Traitement = CopieTableau(i, IndiceTraitement)
        
        If (Not IsEmpty(Affaire) Or Affaire <> "") Or (Not IsEmpty(Rep�re) Or Rep�re <> "") Then
        'Si l'affaire et/ou le rep�re de la pi�ce est vide
            end_char = 0
            For k = 2 To DerLigne - PremLigne + 1
            'Boucle de la ligne apr�s l'en-t�te jusqu'� la fin du CopieTableau
                
                If CopieTableau(k, IndiceMat�riau) = Mat�riau And CopieTableau(k, IndiceTraitement) = Traitement And (Not IsEmpty(CopieTableau(k, IndiceD�signation)) Or CopieTableau(k, IndiceD�signation) <> "") Then
                'Si un autre composant poss�de le m�me couple (Mat�riau, Traitement), les d�signations sont concat�n�es et les masses ajout�es
                    
                    end_char = 1
                    Masse = Masse + CopieTableau(k, IndiceQuantit�) * CopieTableau(k, IndiceMasse)
                    If CopieTableau(k, IndiceQuantit�) = 1 Then
                        D�signation = D�signation & CopieTableau(k, IndiceD�signation) & "," & Chr(10)
                    Else
                        D�signation = D�signation & CopieTableau(k, IndiceQuantit�) & "x " & CopieTableau(k, IndiceD�signation) & "," & Chr(10)
                    End If
                End If
            Next
            
            If end_char = 1 Then
            'Si il y a eu au moins une correspondance entre deux composants, l'espace et la virgule � la fin de la nouvelle d�signation sont supprim�es
                D�signation = Left(D�signation, Len(D�signation) - 2)
            End If
            
            PlageR�sultante = Range(Cells(Indice, PremColonne), Cells(Indice, DerColonne)).Value
                
            PlageR�sultante(1, IndiceAffaire) = "XXX"
            PlageR�sultante(1, IndiceRep�re) = "XXX"
            PlageR�sultante(1, IndiceD�signation) = D�signation
            PlageR�sultante(1, IndiceMat�riau) = Mat�riau
            PlageR�sultante(1, IndiceTraitement) = Traitement
            PlageR�sultante(1, IndiceMasse) = Masse
            PlageR�sultante(1, IndiceR�vision) = "XXX"
            PlageR�sultante(1, IndiceQuantit�) = 1
            
            Range(Cells(Indice, PremColonne), Cells(Indice, DerColonne)).Value = PlageR�sultante
            
            Indice = Indice + 1
            
        End If
    Next
    
    DerLigne = Cells(Rows.Count, Cells(PremLigne, IndiceQuantit� + Offset - 1).Column).End(xlUp).Row
    
    ActiveSheet.Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).RemoveDuplicates Columns:=Array(IndiceAffaire, IndiceD�signation, IndiceMat�riau, IndiceTraitement, IndiceMasse), Header _
        :=xlNo 'L'algorithme pr�c�dent g�n�re des doublons pour chaque couple (Mat�riau, Traitement) qui sont supprim�s
        
    Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Rows.AutoFit
    Range(Cells(PremLigne, PremColonne), Cells(DerLigne, DerColonne)).Columns.AutoFit
    
    DerLigne = Cells(Rows.Count, Cells(PremLigne, IndiceQuantit� + Offset - 1).Column).End(xlUp).Row
    
    For l = PremLigne + 1 To DerLigne
        
        Masse_mat�riau = Cells(l, IndiceMasse)
        Masse_totale = Application.WorksheetFunction.Sum(Range(Cells(PremLigne + 1, IndiceMasse), Cells(DerLigne, IndiceMasse)))
        If Masse_totale <> 0 Then
            Cells(l, IndicePourcentageMasse) = Round(CDbl(Masse_mat�riau) / Masse_totale * 100, 2)
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
            Case 13: MsgBox "Erreur13. Impossible de r�aliser les calculs. V�rifier qu'il n'y a pas de texte dans les colonnes ''Masse'' et ''Compte de r�f�rence''"
            Case Else: MsgBox "Erreur inconnue"
        End Select
        
    Resume fin

End Sub


