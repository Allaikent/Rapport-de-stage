Attribute VB_Name = "Module3"
Sub Cr�er_barres(ConcernedRange As Range, Max As Double, wsWarningsAR As Worksheet)
    ConcernedRange.FormatConditions.AddDatabar
    ConcernedRange.FormatConditions(ConcernedRange.FormatConditions.Count).ShowValue = True
    ConcernedRange.FormatConditions(ConcernedRange.FormatConditions.Count).SetFirstPriority
    ConcernedRange.FormatConditions(1).MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    ConcernedRange.FormatConditions(1).MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=Abs(Max)
                        
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

Sub M�J_Warnings_V2()
 
    'On Error GoTo err
    
    '------------------------------------------------------------------------------------------
    'D�claration des variables
    
    Dim Affaire As Variant
    Dim AffaireVoulue As String
    Dim S�lection As String
    Dim DateVoulue As Variant
    Dim Texte As String
    Dim DateAr As Variant
    Dim NumeroCommande As Variant
    Dim NomFournisseur As Variant
    Dim Commentaire As Variant
    Dim Ref As Variant
    Dim DateLiv As Variant
    Dim Rubrique As Variant
    Dim qte_a_livrer As Variant
    Dim qte As Long
    
    Dim CMDRange As Variant
    Dim NomRange As Variant
    
    Dim Der_Ligne As Long
    Dim Der_Colonne As Long
    Dim indice_tableau_warnings As Long
    Dim indice_affaire_voulue As Long
    Dim DateToday As Date

    Dim wkNom As Workbook
    Dim wkCMD As Workbook
    Dim Path_CMD As String
    Dim wsWarningsAR As Worksheet
    Dim wsFeuil1 As Worksheet
    Dim wsListeProjetsAR As Worksheet
    Dim objConnection As WorkbookConnection
    
    Dim Max As Double
    Dim numero_colonne_rr As Long
    Dim numero_colonne_rp As Long
    Dim numero_colonne_dv As Long
    Dim numero_colonne_av As Long
    Dim numero_colonne_select As Long
    Dim numero_colonne_affaire As Long
    
    Dim letter_rr As String
    Dim letter_rp As String
    
    
    '------------------------------------------------------------------------------------------
    'Emp�cher les pop-ups et le changement d'affichage du classeur actuel
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    '------------------------------------------------------------------------------------------
    'indice_affaire_voulue : g�re la ligne actuelle dans la colonne Numero affaire de la feuille Liste projets AR,
    'indice_tableau_warnings : g�re la ligne actuelle dans la feuille Suivi sur laquelle est �crit les donn�es
    
    'ligne_debut_tableau_warnings : est la ligne o� commence le tableau des warnings pour l'affaire voulue
    'lettre_debut_tableau_warnings : est la colonne o� commence le tableau des warnings pour l'affaire voulue
    
    'Max est la valeur maximale de la barre de progression
    'DateToday est la date d'aujourd'hui
    
    Path_CMD = "T:\ZZ_Planning\CDP\GDP_006_A_Extract CMD EVERWIN (base donn�es).xlsx"
    
    ligne_debut_tableau_warnings = 3
    indice_tableau_warnings = 3
    indice_affaire_voulue = 5
    
    Max = 0
    DateToday = Date
    
    '------------------------------------------------------------------------------------------
    'Dans le classeur "Suivi approvisionnements projets", certaines plages utilis�es dans l'algorithme qui suit sont nomm�es gr�ce au Gestionnaire des noms
    
    numero_colonne_rr = Range("Retard_livraisonAR").Column
    numero_colonne_rp = Range("Retard_projet").Column
    numero_colonne_dv = Range("Date_voulue").Column
    numero_colonne_av = Range("Affaire_voulue").Column
    numero_colonne_select = Range("S�lection").Column
    numero_colonne_affaire = Range("Affaire").Column
    
    letter_affaire = ConvertToLetter(numero_colonne_affaire)
    letter_rr = ConvertToLetter(numero_colonne_rr)
    letter_rp = ConvertToLetter(numero_colonne_rp)
    
    '------------------------------------------------------------------------------------------
    'On attribue le classeur ouvert et chaque feuille � une variable
    
    Set wkNom = ActiveWorkbook
    Set wsWarningsAR = wkNom.Sheets("Warnings AR")
    Set wsListeProjetsAR = wkNom.Sheets("Liste projets AR")
    
    '------------------------------------------------------------------------------------------
    'On supprime le texte et les mises en forme de toute la feuille Warnings AR
    
    For Each rw In wsWarningsAR.Rows(ligne_debut_tableau_warnings & ":" & wsWarningsAR.Cells(Rows.Count, 2).End(xlUp).Row + 1)
        rw.Interior.Pattern = xlNone
        rw.Interior.TintAndShade = 1
        rw.Interior.PatternTintAndShade = 1
        rw.Interior.Color = RGB(255, 255, 255)
    Next
    wsWarningsAR.Rows(ligne_debut_tableau_warnings & ":" & wsWarningsAR.Rows.Count).ClearContents
    wsWarningsAR.Rows(ligne_debut_tableau_warnings & ":" & wsWarningsAR.Rows.Count).FormatConditions.Delete
    wsWarningsAR.Rows(ligne_debut_tableau_warnings & ":" & wsWarningsAR.Rows.Count).Borders.LineStyle = xlLineStyleNone
    
    '------------------------------------------------------------------------------------------
    'On actualise Extract CMD Everwin
    
    If MsgBox("Mettre � jour la BDD Everwin ?", vbYesNo) = vbYes Then
        Set wkCMD = Workbooks.Open(Path_CMD)
        
        For Each objConnection In wkCMD.Connections
            bBackground = objConnection.OLEDBConnection.BackgroundQuery
            objConnection.OLEDBConnection.BackgroundQuery = False
            objConnection.Refresh
            objConnection.OLEDBConnection.BackgroundQuery = bBackground
        Next
    
    wsListeProjetsAR.Cells(2, 6) = DateToday & Chr(13) & Chr(10) & Time
    wkCMD.Close True
    
    End If
    
    '------------------------------------------------------------------------------------------
    'On ouvre l'extract CMD et on calcule sa derni�re ligne et sa derni�re colonne
    
    Set wkCMD = Workbooks.Open(Path_CMD)
    Set wsFeuil1 = wkCMD.Sheets("Feuil1")
    
    Der_Ligne = wsFeuil1.Cells(Rows.Count, 1).End(xlUp).Row
    Der_Colonne = wsFeuil1.Cells(1, Columns.Count).End(xlToLeft).Column
    
    '------------------------------------------------------------------------------------------
    'En VBA pour �conomiser un grand temps de calcul en lecture on lit toutes les donn�es d'un grand tableau (ici le tableau de la Feuille1 de l'extract CMD) une seule fois gr�ce � la commande Range.Value
    
    CMDRange = wsFeuil1.Range("A1:" & ConvertToLetter(CLng(Der_Colonne)) & Der_Ligne).Value
    
    '------------------------------------------------------------------------------------------
    'La premi�re boucle s'arr�te quand une cellule de la colonne "Numero affaire" du tableau pr�sent dans la feuille "Liste projets AR" est vide
    
    Do While (Not IsEmpty(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)))
    
        AffaireVoulue = wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)
        DateVoulue = wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_dv)
        S�lection = wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_select)
        
        If (IsEmpty(S�lection) Or S�lection = "") Then
            GoTo FinDoWhile
        End If
        
    '------------------------------------------------------------------------------------------
    'Du d�but � la fin de l'extract CMD
        
        For i = 2 To Der_Ligne
            
            Texte = CMDRange(i, 8)
            Affaire = CMDRange(i, 5)
            DateAr = CMDRange(i, 15)
            Commentaire = CMDRange(i, 16)
            NumeroCommande = CMDRange(i, 3)
            NomFournisseur = CMDRange(i, 4)
            Ref = CMDRange(i, 7)
            DateLiv = CMDRange(i, 14)
            Rubrique = CMDRange(i, 6)
            qte_a_livrer = CMDRange(i, 18)
            qte = CMDRange(i, 9)
            
    '------------------------------------------------------------------------------------------
    '6 cas temporels sont possibles dont 5 l�vent des warnings, ici on s'assure que le cas non warning n'est pas relev� d'o� le Not
    'Pour bien comprendre quelle est la condition il faut aller voir la feuille "Sch�ma warnings" et lire l'encadr� � c�t� du sch�ma d'explication des barres de progression
            
            If (((IsEmpty(qte_a_livrer) Or qte_a_livrer = "") And (IsEmpty(Commentaire) Or Commentaire = "") Or (qte_a_livrer <> "0" And Not (IsEmpty(qte_a_livrer) Or qte_a_livrer = ""))) And Not (DateToday <= DateVoulue And ((DateToday <= CDate(DateAr) And CDate(DateAr) <= DateVoulue) Or (DateToday <= CDate(DateLiv) And CDate(DateLiv) <= DateVoulue)))) And (Not (IsEmpty(DateAr)) Or Not IsEmpty(DateLiv)) And Not IsEmpty(Affaire) And InStr(1, Affaire, AffaireVoulue) And Rubrique = "ACHA" Then
                
                NomRange = wsWarningsAR.Range("B" & indice_tableau_warnings & ":K" & indice_tableau_warnings).Value
                
                NomRange(1, 1) = Affaire
                NomRange(1, 2) = NumeroCommande
                NomRange(1, 3) = NomFournisseur
                NomRange(1, 4) = Ref
                NomRange(1, 5) = Texte
                NomRange(1, 6) = DateAr
                NomRange(1, 7) = DateLiv
                NomRange(1, 8) = Commentaire
                NomRange(1, 9) = qte_a_livrer
                NomRange(1, 10) = DateVoulue
                
                wsWarningsAR.Range("B" & indice_tableau_warnings & ":K" & indice_tableau_warnings).Value = NomRange
                
    '------------------------------------------------------------------------------------------
    'On attribue les couleurs/barres de progression � chaque cas en priorisant AR sur livraison
                
                If IsEmpty(DateAr) Then
                
                    If Not IsEmpty(DateLiv) And DateToday >= CDate(DateLiv) And DateToday >= DateVoulue Then
                        
                        NomRange = wsWarningsAR.Range(letter_rr & indice_tableau_warnings & ":" & letter_rp & indice_tableau_warnings).Value
                        NomRange(1, 1) = (DateToday - CDate(DateLiv))
                        NomRange(1, 2) = DateToday - DateVoulue
                        wsWarningsAR.Range(letter_rr & indice_tableau_warnings & ":" & letter_rp & indice_tableau_warnings).Value = NomRange
                        wsWarningsAR.Cells(indice_tableau_warnings, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                        If CDate(DateLiv) >= DateVoulue Then
                            Max = 1
                        Else
                            Max = DateVoulue - CDate(DateLiv)
                        End If
                        Cr�er_barres wsWarningsAR.Range(ConvertToLetter(numero_colonne_rr) & indice_tableau_warnings), Max, wsWarningsAR
    
                    ElseIf Not IsEmpty(DateLiv) And DateToday <= CDate(DateLiv) And DateToday >= DateVoulue Then
                        
                        NomRange = wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value
                        NomRange = CDate(DateLiv) - DateVoulue
                        wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value = NomRange
                        wsWarningsAR.Cells(indice_tableau_warnings, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                        
                    ElseIf Not IsEmpty(DateLiv) And DateToday <= DateVoulue And DateToday >= CDate(DateLiv) Then
                        
                        NomRange = wsWarningsAR.Range(letter_rr & indice_tableau_warnings).Value
                        NomRange = (DateToday - CDate(DateLiv))
                        wsWarningsAR.Range(letter_rr & indice_tableau_warnings).Value = NomRange
                        Max = CDate(DateLiv) - DateVoulue
                        Cr�er_barres wsWarningsAR.Range(ConvertToLetter(numero_colonne_rr) & indice_tableau_warnings), Max, wsWarningsAR
                        
                    ElseIf Not IsEmpty(DateLiv) And DateToday <= DateVoulue And DateToday <= CDate(DateLiv) And DateVoulue <= CDate(DateLiv) Then
                        
                        NomRange = wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value
                        NomRange = CDate(DateLiv) - DateVoulue
                        wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value = NomRange
                        wsWarningsAR.Cells(indice_tableau_warnings, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                    
                    End If
                
                Else
                    If Not IsEmpty(DateAr) And DateToday >= CDate(DateAr) And DateToday >= DateVoulue Then
                        
                        NomRange = wsWarningsAR.Range(letter_rr & indice_tableau_warnings & ":" & letter_rp & indice_tableau_warnings).Value
                        NomRange(1, 1) = (DateToday - CDate(DateAr))
                        NomRange(1, 2) = DateToday - DateVoulue
                        wsWarningsAR.Range(letter_rr & indice_tableau_warnings & ":" & letter_rp & indice_tableau_warnings).Value = NomRange
                        wsWarningsAR.Cells(indice_tableau_warnings, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                        
                        If CDate(DateAr) >= DateVoulue Then
                            Max = 1
                        Else
                            Max = DateVoulue - CDate(DateAr)
                        End If
                        Cr�er_barres wsWarningsAR.Range(letter_rr & indice_tableau_warnings), Max, wsWarningsAR
                        
                    ElseIf Not IsEmpty(DateAr) And DateToday <= CDate(DateAr) And DateToday >= DateVoulue Then
                        
                        NomRange = wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value
                        NomRange = CDate(DateAr) - DateVoulue
                        wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value = NomRange
                        wsWarningsAR.Cells(indice_tableau_warnings, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                        
                    ElseIf Not IsEmpty(DateAr) And DateToday <= DateVoulue And DateToday >= CDate(DateAr) Then
                        
                        NomRange = wsWarningsAR.Range(letter_rr & indice_tableau_warnings).Value
                        NomRange = (DateToday - CDate(DateAr))
                        wsWarningsAR.Range(letter_rr & indice_tableau_warnings).Value = NomRange
                        Max = CDate(DateAr) - DateVoulue
                        Cr�er_barres wsWarningsAR.Range(letter_rr & indice_tableau_warnings), Max, wsWarningsAR
                        
                    ElseIf Not IsEmpty(DateAr) And DateToday <= DateVoulue And DateToday <= CDate(DateAr) And DateVoulue <= CDate(DateAr) Then
                        
                        NomRange = wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value
                        NomRange = CDate(DateAr) - DateVoulue
                        wsWarningsAR.Range(letter_rp & indice_tableau_warnings).Value = NomRange
                        wsWarningsAR.Cells(indice_tableau_warnings, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                    
                    End If
                End If
                
    '------------------------------------------------------------------------------------------
    'Bordure entre chaque ligne des warnings
            
            wsWarningsAR.Range(letter_affaire & ligne_debut_tableau_warnings & ":" & letter_rp & indice_tableau_warnings - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
            wsWarningsAR.Range(letter_affaire & ligne_debut_tableau_warnings & ":" & letter_rp & indice_tableau_warnings - 1).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
            
            indice_tableau_warnings = indice_tableau_warnings + 1
            
            End If
        
        Next

    '------------------------------------------------------------------------------------------
    'Bordure par affaire
    
    wsWarningsAR.Range(letter_affaire & ligne_debut_tableau_warnings & ":" & letter_rp & indice_tableau_warnings - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    wsWarningsAR.Range(letter_affaire & ligne_debut_tableau_warnings & ":" & letter_rp & indice_tableau_warnings - 1).Borders(xlEdgeBottom).Weight = xlThick
    wsWarningsAR.Range(letter_affaire & ligne_debut_tableau_warnings & ":" & letter_rp & indice_tableau_warnings - 1).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
       
FinDoWhile:
    
    '------------------------------------------------------------------------------------------
    'Ici on passe � la ligne suivante de la colonne "Affaire voulue" dans Extract Nomenclature
    
    indice_affaire_voulue = indice_affaire_voulue + 1
    Loop
    
    '------------------------------------------------------------------------------------------
    'Bordure retards
    wsWarningsAR.Range(letter_rr & "3:" & letter_rr & indice_tableau_warnings - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    
    wkCMD.Close False
    
    'fin:
        'Exit Sub
    
    'err:
        'Select Case err.Number
            'Case 13: MsgBox "Pas de M�J effectu�e"
            'Case 1004: MsgBox "La macro n'arrive pas � acc�der au fichier extrait des CMD extraites d'Everwin"
            'Case Else: MsgBox "Erreur inconnue"
        'End Select
    
    'Resume fin
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
End Sub








