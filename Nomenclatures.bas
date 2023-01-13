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

Sub Nomenclatures()
Attribute Nomenclatures.VB_ProcData.VB_Invoke_Func = " \n14"

'On Error GoTo err

    '------------------------------------------------------------------------------------------
    'D�claration des variables
    
    Dim wkListeProjetsAR As Workbook
    Dim wsListeProjetsAR As Worksheet
    Dim wsNomenclatures As Worksheet
    Dim wkNom As Workbook
    Dim wsNom As Worksheet
    
    '------------------------------------------------------------------------------------------
    'Emp�cher les pop-ups et le changement d'affichage du classeur actuel
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    '------------------------------------------------------------------------------------------
    'Dans le classeur "Suivi approvisionnements projets", certaines plages utilis�es dans l'algorithme qui suit sont nomm�es gr�ce au Gestionnaire des noms
    
    numero_colonne_meca = Range("Nomenclature_m�ca").Column
    numero_colonne_elec = Range("Nomenclature_�lec").Column
    numero_colonne_autre1 = Range("Nomenclature_autre1").Column
    numero_colonne_autre2 = Range("Nomenclature_autre2").Column
    numero_colonne_selection = Range("S�lection2").Column
    numero_colonne_dv = Range("Date_voulue").Column
    numero_colonne_av = Range("Affaire_voulue").Column
    
    indice_tableau_nomenclatures = 3
    Prem_Ligne_Projets = 5
    
    lettre_debut_tableau_nomenclatures = "B"
    lettre_fin_tableau_nomenclatures = "K"
    
    '------------------------------------------------------------------------------------------
    'On attribue le classeur ouvert et les feuilles "Liste projets AR" et "Nomenclatures" � une variable
    
    Set wkListeProjetsAR = ActiveWorkbook
    Set wsListeProjetsAR = wkListeProjetsAR.Sheets("Liste projets AR")
    Set wsNomenclatures = wkListeProjetsAR.Sheets("Nomenclatures")
    
    '------------------------------------------------------------------------------------------
    'On supprime le texte et les mises en forme de toute la feuille Nomenclatures
    
    wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).FormatConditions.Delete
    wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).Font.Bold = False
    wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).Font.Color = RGB(0, 0, 0)
    wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).ClearContents
    wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).Borders.LineStyle = xlLineStyleNone
    wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).HorizontalAlignment = xlCenter
    wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).VerticalAlignment = xlCenter
    With wsNomenclatures.Rows(indice_tableau_nomenclatures & ":" & wsNomenclatures.Rows.Count).Interior
            .Pattern = xlNone
            .TintAndShade = 1
            .PatternTintAndShade = 1
            .Color = RGB(255, 255, 255)
    End With
    
    '------------------------------------------------------------------------------------------
    'La derni�re ligne � traiter de la feuille "Liste projets AR" est calcul�e en remontant la colonne "Nomenclature m�ca" jusqu'� trouver une cellule non vide
    
    Der_Ligne_Projets = wsListeProjetsAR.Cells(Rows.Count, numero_colonne_meca).End(xlUp).Row
    
    For k = Prem_Ligne_Projets To Der_Ligne_Projets
    
    '------------------------------------------------------------------------------------------
    'La condition v�rifie si la cellule correspondante au num�ro d'affaire dans la colonne S�lection n'est pas vide

        If Not IsEmpty(wsListeProjetsAR.Cells(k, numero_colonne_selection)) And wsListeProjetsAR.Cells(k, numero_colonne_selection) <> "" Then
            AffaireVoulue = wsListeProjetsAR.Cells(k, numero_colonne_av)
            DateVoulue = wsListeProjetsAR.Cells(k, numero_colonne_dv)
            
            For j = numero_colonne_meca To numero_colonne_autre2
                
                If Not IsEmpty(wsListeProjetsAR.Cells(k, j)) And wsListeProjetsAR.Cells(k, j) <> "" Then
                        
    '------------------------------------------------------------------------------------------
    'Si la condition est v�rifi�e, la nomenclature est ouverte et les colonnes voulues en sortie sont recherch�es et leur num�ro de colonne affect�e � des variables
                        
                        Adresse_nom = wsListeProjetsAR.Cells(k, j).Hyperlinks(1).Address
                        Set wkNom = Workbooks.Open(Adresse_nom, UpdateLinks:=0)
                        Set wsNom = wkNom.Worksheets("Nomenclature")
                        
                        numero_colonne_affaire_source = wsNom.Rows(2).Find("Affaire source").Column
                        numero_colonne_quantite = wsNom.Rows(2).Find("Quantit�").Column
                        
    '------------------------------------------------------------------------------------------
    'La colonne Rep�re n'est pas toujours pr�sente dans les nomenclatures, d'o� la condition
                        
                        If Not wsNom.Rows(2).Find("Rep�re") Is Nothing Then
                            numero_colonne_repere = wsNom.Rows(2).Find("Rep�re").Column
                        Else
                            numero_colonne_repere = 0
                        End If
                        
                        numero_colonne_designation = wsNom.Rows(2).Find("D�signation").Column
                        
    '------------------------------------------------------------------------------------------
    'Selon la nomenclature, il est �crit Fabriquant ou Fournisseur, d'o� les conditions
                        
                        If Not wsNom.Rows(2).Find("Fabriquant") Is Nothing Then
                            numero_colonne_fabriquant = wsNom.Rows(2).Find("Fabriquant").Column
                            numero_colonne_fournisseur = 0
                        ElseIf Not wsNom.Rows(2).Find("Fournisseur") Is Nothing Then
                            numero_colonne_fournisseur = wsNom.Rows(2).Find("Fournisseur").Column
                            numero_colonne_fabriquant = 0
                        End If
                        
                        numero_colonne_reference = wsNom.Rows(2).Find("R�f�rence").Column
                        numero_colonne_Distributeur = wsNom.Rows(2).Find("Distributeur").Column
                        numero_colonne_reference_distributeur = wsNom.Rows(2).Find("R�f. Distributeur").Column
                        numero_colonne_remarques = wsNom.Rows(2).Find("Remarques").Column
                        numero_colonne_etat = wsNom.Rows(2).Find("Etat").Column
                        numero_colonne_localisation = wsNom.Rows(2).Find("Localisation").Column
                        
    '------------------------------------------------------------------------------------------
    'La derni�re ligne de la nomenclature est calcul�e en remontant la colonne D�signation par le bas jusqu'� trouver une cellule non vide
                        
                        Der_Ligne_Nomenclature = wsNom.Cells(Rows.Count, numero_colonne_designation).End(xlUp).Row
                        NomRange = wsNom.Range("A2:" & ConvertToLetter(CLng(numero_colonne_localisation)) & Der_Ligne_Nomenclature).Value
                        
                        For i = 2 To Der_Ligne_Nomenclature - 1
                        Quantite = NomRange(i, numero_colonne_quantite)
                        
    '------------------------------------------------------------------------------------------
    'La condition v�rifie si la Quantit� de la ligne est diff�rente de 0 ou vide et que la ligne est non barr�e
                        
                        If (Quantite <> 0 Or IsEmpty(Quantite)) And wsNom.Cells(i + 1, numero_colonne_quantite).Font.Strikethrough = False Then
                            
                            Affaire_source = NomRange(i, numero_colonne_affaire_source)
                            
                            If numero_colonne_repere <> 0 Then
                                Repere = NomRange(i, numero_colonne_repere)
                            Else
                                Repere = Empty
                            End If
                            
                            Designation = NomRange(i, numero_colonne_designation)
                            
                            If numero_colonne_fabriquant = 0 Then
                                Fournisseur = NomRange(i, numero_colonne_fournisseur)
                            ElseIf numero_colonne_fournisseur = 0 Then
                                Fabriquant = NomRange(i, numero_colonne_fabriquant)
                            End If
                            
                            Reference = NomRange(i, numero_colonne_reference)
                            Distributeur = NomRange(i, numero_colonne_Distributeur)
                            Reference_distributeur = NomRange(i, numero_colonne_reference_distributeur)
                            Remarques = NomRange(i, numero_colonne_remarques)
                            Etat = NomRange(i, numero_colonne_etat)
                            
                            If UCase(Etat) = UCase("BPC") Or UCase(Etat) = UCase("Consult�") Or UCase(Etat) = UCase("Etude") Or IsEmpty(Etat) And Not IsEmpty(Designation) And Designation <> "" Then
                                
                                ResultRange = wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & indice_tableau_nomenclatures & ":" & lettre_fin_tableau_nomenclatures & indice_tableau_nomenclatures).Value
                                
                                ResultRange(1, 1) = AffaireVoulue
                                ResultRange(1, 2) = Affaire_source
                                ResultRange(1, 3) = Repere
                                ResultRange(1, 4) = Designation
                                If numero_colonne_fabriquant = 0 Then
                                ResultRange(1, 5) = Fournisseur
                                ElseIf numero_colonne_fournisseur = 0 Then
                                ResultRange(1, 5) = Fabriquant
                                End If
                                ResultRange(1, 6) = Reference
                                ResultRange(1, 7) = Distributeur
                                ResultRange(1, 8) = Reference_distributeur
                                ResultRange(1, 9) = Remarques
                                ResultRange(1, 10) = Etat
                                
                                wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & indice_tableau_nomenclatures & ":" & lettre_fin_tableau_nomenclatures & indice_tableau_nomenclatures).Value = ResultRange
                                wsNomenclatures.Rows.AutoFit
                                wsNomenclatures.Columns.Font.Size = 28
                                wsNomenclatures.Columns.AutoFit
                                
                                If UCase(Etat) = UCase("Etude") Then
                                    wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & indice_tableau_nomenclatures & ":" & lettre_fin_tableau_nomenclatures & indice_tableau_nomenclatures).Interior.Color = RGB(192, 0, 0)
                                End If
                                If UCase(Etat) = UCase("Consult�") Then
                                    wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & indice_tableau_nomenclatures & ":" & lettre_fin_tableau_nomenclatures & indice_tableau_nomenclatures).Interior.Color = RGB(255, 192, 0)
                                End If
                                
                                indice_tableau_nomenclatures = indice_tableau_nomenclatures + 1
                                wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & (indice_tableau_nomenclatures - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_tableau_nomenclatures - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                                wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & (indice_tableau_nomenclatures - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_tableau_nomenclatures - 1)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                            End If
                        End If
                        Next
                        
    '------------------------------------------------------------------------------------------
    'Bordure basse entre chaque ligne des nomenclatures
                        
                        wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & (indice_tableau_nomenclatures - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_tableau_nomenclatures - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & (indice_tableau_nomenclatures - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_tableau_nomenclatures - 1)).Borders(xlEdgeBottom).Weight = xlThick
                        wsNomenclatures.Range(lettre_debut_tableau_nomenclatures & (indice_tableau_nomenclatures - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_tableau_nomenclatures - 1)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                        wkNom.Close False
                        End If
                    Next
        End If
    Next

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
