Attribute VB_Name = "Module1"
Function ConvertToLetter(iCol As Long) As String
   Dim a As Long
   Dim b As Long
   Dim c As Long
   a = iCol
   c = iCol
   ConvertToLetter = ""
   Do While c > 0
      a = Int((iCol - 1) / 26)
      b = (iCol - 1) Mod 26
      ConvertToLetter = Chr(b + 65) & ConvertToLetter
      c = a
   Loop
End Function

Sub DML_DepuisUneNomenclature()
    
    On Error GoTo err
    Dim Affaire As Variant
    Dim Masse As Variant
    Dim indice_tableau As Integer
    Dim Désignation As Variant
    Dim Matériau As Variant
    Dim Traitement As Variant
    Dim Tab_Range As Variant
    Dim Result_Range As Variant
    Dim Quantité As Variant
    Dim indice_masse As Long
    Dim indice_pourcentage_masse As Long
    Dim Der_Ligne As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    For Each rw In Cells.Rows
        If Cells(rw.Row, Columns.Count).End(xlToLeft).Column <> 1 Then
            Prem_Ligne = rw.Row
            Exit For
        End If
    Next
        
    For Each clmn In Cells.Columns
        If Cells(Rows.Count, clmn.Column).End(xlUp).Row <> 1 Then
            Lettre_début = ConvertToLetter(clmn.Column)
            Exit For
        End If
    Next
        
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    indice_tableau = Prem_Ligne + 1
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    'Ici c'est l'ordre des colonnes, la première commence toujours à 1.
    
    Offset = ActiveSheet.Range(Lettre_début & Prem_Ligne).Column
    
    indice_affaire = ActiveSheet.Rows(Prem_Ligne).Find("Affaire").Column - Offset + 1
    indice_repère = ActiveSheet.Rows(Prem_Ligne).Find("Repère").Column - Offset + 1
    indice_désignation = ActiveSheet.Rows(Prem_Ligne).Find("Désignation").Column - Offset + 1
    indice_matériau = ActiveSheet.Rows(Prem_Ligne).Find("Matériau").Column - Offset + 1
    indice_traitement = ActiveSheet.Rows(Prem_Ligne).Find("Traitement").Column - Offset + 1
    indice_masse = ActiveSheet.Rows(Prem_Ligne).Find("Masse").Column - Offset + 1
    indice_révision = ActiveSheet.Rows(Prem_Ligne).Find("Révision").Column - Offset + 1
    indice_pourcentage_masse = ActiveSheet.Rows(Prem_Ligne).Find("Configuration").Column - Offset + 1
    indice_quantité = ActiveSheet.Rows(Prem_Ligne).Find("Compte de référence").Column - Offset + 1
    
    Der_Ligne = Cells(Rows.Count, Cells(Prem_Ligne, indice_quantité + Offset - 1).Column).End(xlUp).Row
    Lettre_fin = ConvertToLetter(Cells(Cells(Prem_Ligne, indice_quantité + Offset - 1).Row, Columns.Count).End(xlToLeft).Column)
    Der_Colonne = Cells(Cells(Prem_Ligne, indice_quantité + Offset - 1).Row, Columns.Count).End(xlToLeft).Column
    
    '---------------------------------------------------------------------------------------------------------------------------------'
    
    Plage_tableau = Lettre_début & Prem_Ligne & ":" & Lettre_fin & Der_Ligne
    Tab_Range = Range(Plage_tableau).Value
    Range(Plage_tableau).Offset(1).ClearContents

    For i = 2 To Der_Ligne - Prem_Ligne + 1
        Masse = 0#
        Désignation = Null
        Traitement = Null
        Matériau = Tab_Range(i, indice_matériau)
        Affaire = Tab_Range(i, indice_affaire)
        Traitement = Tab_Range(i, indice_traitement)
        Plage_ligne_tableau = Lettre_début & indice_tableau & ":" & Lettre_fin & indice_tableau
        
        If (Not IsEmpty(Affaire) Or Affaire <> "") Or (Not IsEmpty(Repère) Or Repère <> "") Then
            end_char = 0
            For k = 2 To Der_Ligne - Prem_Ligne + 1
                If Tab_Range(k, indice_matériau) = Matériau And Tab_Range(k, indice_traitement) = Traitement And (Not IsEmpty(Tab_Range(k, indice_désignation)) Or Tab_Range(k, indice_désignation) <> "") Then
                    end_char = 1
                    Masse = Masse + Tab_Range(k, indice_quantité) * Tab_Range(k, indice_masse)
                    If Tab_Range(k, indice_quantité) = 1 Then
                        Désignation = Désignation & Tab_Range(k, indice_désignation) & "," & Chr(10)
                    Else
                        Désignation = Désignation & Tab_Range(k, indice_quantité) & "x " & Tab_Range(k, indice_désignation) & "," & Chr(10)
                    End If
                End If
            Next
            
            If end_char = 1 Then
                Désignation = Left(Désignation, Len(Désignation) - 2)
            End If
            
            Result_Range = Range(Plage_ligne_tableau).Value
                
            Result_Range(1, indice_affaire) = "XXX"
            Result_Range(1, indice_repère) = "XXX"
            Result_Range(1, indice_désignation) = Désignation
            Result_Range(1, indice_matériau) = Matériau
            Result_Range(1, indice_traitement) = Traitement
            Result_Range(1, indice_masse) = Masse
            Result_Range(1, indice_révision) = "XXX"
            Result_Range(1, indice_quantité) = 1
            
            Range(Plage_ligne_tableau).Value = Result_Range
            
            indice_tableau = indice_tableau + 1
        End If
    Next
    
    Der_Ligne = Cells(Rows.Count, Cells(Prem_Ligne, indice_quantité + Offset - 1).Column).End(xlUp).Row
    ActiveSheet.Range("$" & Lettre_début & "$" & Prem_Ligne & ":" & "$" & Lettre_fin & "$" & Der_Ligne).RemoveDuplicates Columns:=Array(indice_affaire, indice_désignation, indice_matériau, indice_traitement, indice_masse), Header _
        :=xlNo
    Range(Plage_tableau).Rows.AutoFit
    Range(Plage_tableau).Columns.AutoFit
    Der_Ligne = Cells(Rows.Count, Cells(Prem_Ligne, indice_quantité + Offset - 1).Column).End(xlUp).Row
    
    
    For l = Prem_Ligne + 1 To Der_Ligne
        Letter_masse = ConvertToLetter(indice_masse + Offset - 1)
        Letter_pourcentage_masse = ConvertToLetter(indice_pourcentage_masse + Offset - 1)
        Masse_matériau = Range(Letter_masse & l)
        Masse_totale = Application.WorksheetFunction.Sum(Range(Letter_masse & Prem_Ligne + 1 & ":" & Letter_masse & Der_Ligne))
        If Masse_totale <> 0 Then
            Range(Letter_pourcentage_masse & l) = Round(CDbl(Masse_matériau) / Masse_totale * 100, 2)
        End If
    Next
    
    Rows(Prem_Ligne & ":" & Der_Ligne).Sort Key1:=Range(Cells(Prem_Ligne, indice_masse), Cells(Der_Ligne, indice_masse)), order1:=xlDescending
    Cells(Der_Ligne + 1, Der_Colonne + 1) = "Masse totale :" & " " & Masse_totale

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


