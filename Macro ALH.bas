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

Sub Group_by()

'On Error GoTo err
Dim Affaire As Variant
Dim Masse As Variant
Dim indice_tableau As Integer
Dim D�signation As Variant
Dim Mat�riau As Variant
Dim Traitement As Variant
Dim Tab_Range As Variant
Dim Result_Range As Variant
Dim Quantit� As Variant
Dim indice_masse As Long
Dim indice_pourcentage_masse As Long
Dim Der_Ligne As Long

Application.ScreenUpdating = False

indice_tableau = 2
indice_affaire = 1
indice_rep�re = 2
indice_d�signation = 3
indice_mat�riau = 4
indice_traitement = 5
indice_masse = 6
indice_r�vision = 7
indice_pourcentage_masse = 8
indice_quantit� = 9

Prem_Ligne = 1
Der_Ligne = Cells(Rows.Count, 1).End(xlUp).Row
Lettre_d�but = "A"
Lettre_fin = ConvertToLetter(Cells(1, Columns.Count).End(xlToLeft).Column)

Plage_tableau = Lettre_d�but & Prem_Ligne & ":" & Lettre_fin & Der_Ligne
Tab_Range = Range(Plage_tableau).Value
Range(Plage_tableau).Offset(1).Delete

For i = 2 To Der_Ligne
    Masse = 0#
    D�signation = Null
    Traitement = Null
    Mat�riau = Tab_Range(i, indice_mat�riau)
    Affaire = Tab_Range(i, indice_affaire)
    Traitement = Tab_Range(i, indice_traitement)
    Plage_ligne_tableau = Lettre_d�but & indice_tableau & ":" & Lettre_fin & indice_tableau
    
    If (Not IsEmpty(Mat�riau) Or Mat�riau <> "") Or (Not IsEmpty(Affaire) Or Affaire <> "") Then
        end_char = 0
        For k = 2 To Der_Ligne
            If Tab_Range(k, indice_mat�riau) = Mat�riau And Tab_Range(k, indice_traitement) = Traitement And (Not IsEmpty(Tab_Range(k, indice_affaire)) Or Tab_Range(k, indice_affaire) <> "") Then
                end_char = 1
                Masse = Masse + Tab_Range(k, indice_quantit�) * Tab_Range(k, indice_masse)
                If Tab_Range(k, indice_quantit�) = 1 Then
                    D�signation = D�signation & Tab_Range(k, indice_d�signation) & "," & Chr(10)
                Else
                    D�signation = D�signation & Tab_Range(k, indice_quantit�) & "x " & Tab_Range(k, indice_d�signation) & "," & Chr(10)
                End If
            End If
        Next
        
        If end_char = 1 Then
            D�signation = Left(D�signation, Len(D�signation) - 2)
        End If
        
        
        Result_Range = Range(Plage_ligne_tableau).Value
            
        Result_Range(1, indice_affaire) = "XXX"
        Result_Range(1, indice_rep�re) = "XXX"
        Result_Range(1, indice_d�signation) = D�signation
        Result_Range(1, indice_mat�riau) = Mat�riau
        Result_Range(1, indice_traitement) = Traitement
        Result_Range(1, indice_masse) = Masse
        Result_Range(1, indice_r�vision) = "XXX"
        Result_Range(1, indice_quantit�) = 1
        
        Range(Plage_ligne_tableau).Value = Result_Range
        
        indice_tableau = indice_tableau + 1
    End If
Next

Der_Ligne = [I6000].End(xlUp).Row
ActiveSheet.Range("$" & Lettre_d�but & "$" & Prem_Ligne & ":" & "$" & Lettre_fin & "$" & Der_Ligne).RemoveDuplicates Columns:=Array(indice_mat�riau, indice_traitement, indice_d�signation, indice_affaire, indice_rep�re, indice_masse), Header _
    :=xlNo
Range(Plage_tableau).Rows.AutoFit
Range(Plage_tableau).Columns.AutoFit
Der_Ligne = [I6000].End(xlUp).Row


For l = Prem_Ligne + 1 To Der_Ligne
    Letter_masse = ConvertToLetter(indice_masse)
    Letter_pourcentage_masse = ConvertToLetter(indice_pourcentage_masse)
    Masse_mat�riau = Range(Letter_masse & l)
    Masse_totale = Application.WorksheetFunction.Sum(Range(Letter_masse & "2:" & Letter_masse & Der_Ligne))
    If Masse_totale <> 0 Then
        Range(Letter_pourcentage_masse & l) = Round(CDbl(Masse_mat�riau) / Masse_totale, 2)
    End If
Next
    
Application.ScreenUpdating = True

fin:
    Exit Sub

err:
    Select Case err.Number
        Case Else: MsgBox "Erreur inconnue"
    End Select
    
Resume fin

End Sub


