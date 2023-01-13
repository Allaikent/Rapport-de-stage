Attribute VB_Name = "Module2"
Sub Cr�er_barres(ConcernedRange As Range, Max As Double, wsSuiviAR As Worksheet)
    'ConcernedRange est la plage de donn�es sur laquelle les barres vont �tre appliqu�es
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

Sub Suivi()
    
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
    
    Dim Der_Ligne As Double
    Dim Der_Colonne As Long
    Dim indice_feuille_suivi As Long
    Dim indice_affaire_voulue As Long
    Dim DateToday As Date

    Dim wkSuivi_appros As Workbook
    Dim wsSuiviAR As Worksheet
    Dim wsListeProjetsAR As Worksheet
    Dim wsSuivi As Worksheet
    Dim wsExtractNomcl As Worksheet
    
    Dim Path_CMD As String
    Dim wkCMD As Workbook
    Dim wsFeuil1 As Worksheet
    Dim objConnection As WorkbookConnection
    
    Dim wkNom As Workbook
    Dim wsNom As Worksheet
    
    Dim Max As Double
    Dim numero_colonne_rr As Long
    Dim numero_colonne_rp As Long
    Dim numero_colonne_dv As Long
    Dim numero_colonne_av As Long
    Dim numero_colonne_select As Long
    Dim numero_colonne_affaire As Long
    Dim numero_colonne_autre1 As Long
    Dim numero_colonne_autre2 As Long
    
    Dim numero_colonne_debut_extract_nomcl_audric As Long
    Dim numero_colonne_fin_extract_nomcl_audric As Long
    Dim numero_colonne_debut_extract_nomcl_vincent As Long
    Dim numero_colonne_fin_extract_nomcl_vincent As Long
    
    Dim lettre_rr As String
    Dim lettre_rp As String
    
    
    '------------------------------------------------------------------------------------------
    'Emp�cher les pop-ups et le changement d'affichage du classeur actuel
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlAutomatic
    
    '------------------------------------------------------------------------------------------
    'indice_affaire_voulue : g�re la ligne actuelle dans la colonne Numero affaire de la feuille Liste projets AR,
    'indice_feuille_suivi : g�re la ligne actuelle dans la feuille Suivi sur laquelle est �crit les donn�es
    
    'ligne_debut_tableau_warnings : est la ligne o� commence le tableau des warnings pour l'affaire voulue
    'lettre_debut_tableau_warnings : est la colonne o� commence le tableau des warnings pour l'affaire voulue
    
    'ligne_debut_tableau_nomenclatures : est la ligne o� commence le tableau des nomenclatures pour l'affaire voulue
    'lettre_debut_tableau_nomenclatures : est la colonne o� commence le tableau des nomenclatures pour l'affaire voulue
    'lettre_fin_tableau_nomenclatures : est la colonne o� termine le tableau des nomenclatures pour l'affaire voulue
    
    'Max est la valeur maximale de la barre de progression
    'DateToday est la date d'aujourd'hui
    
    Path_CMD = "T:\ZZ_Planning\CDP\GDP_006_A_Extract CMD EVERWIN (base donn�es).xlsx"
    
    indice_affaire_voulue = 5
    indice_feuille_suivi = 3
    
    ligne_debut_tableau_warnings = 3
    lettre_debut_tableau_warnings = "B"
    
    ligne_debut_tableau_nomenclatures = 3
    lettre_debut_tableau_nomenclatures = "B"
    lettre_fin_tableau_nomenclatures = "K"
    
    Max = 0
    DateToday = Date
    
    '------------------------------------------------------------------------------------------
    'On attribue le classeur ouvert et chaque feuille � une variable
    
    Set wkSuivi_appros = ActiveWorkbook
    Set wsListeProjetsAR = wkSuivi_appros.Sheets("Liste projets AR")
    Set wsSuivi = wkSuivi_appros.Sheets("Suivi")
    Set wsExtractNomcl = wkSuivi_appros.Sheets("Extract Nomcl projets")
    
    '------------------------------------------------------------------------------------------
    'Dans le classeur "Suivi approvisionnements projets", certaines plages utilis�es dans l'algorithme qui suit sont nomm�es gr�ce au Gestionnaire des noms
    
    numero_colonne_rr = Range("Retard_livraisonAR").Column
    numero_colonne_rp = Range("Retard_projet").Column
    numero_colonne_dv = Range("Date_voulue").Column
    numero_colonne_av = Range("Affaire_voulue").Column
    numero_colonne_select = wsListeProjetsAR.Range("S�lection3").Column
    numero_colonne_affaire = Range("Affaire").Column
    
    numero_colonne_meca = wsListeProjetsAR.Range("Nomenclature_m�ca").Column
    numero_colonne_elec = wsListeProjetsAR.Range("Nomenclature_�lec").Column
    numero_colonne_autre1 = wsListeProjetsAR.Range("Nomenclature_autre1").Column
    numero_colonne_autre2 = wsListeProjetsAR.Range("Nomenclature_autre2").Column
    
    numero_colonne_debut_extract_nomcl_audric = wsExtractNomcl.Range("Audric").Columns.Column
    numero_colonne_fin_extract_nomcl_audric = wsExtractNomcl.Range("Audric").Columns.Count + numero_colonne_debut_extract_nomcl_audric - 1
    numero_colonne_debut_extract_nomcl_vincent = wsExtractNomcl.Range("Vincent").Columns.Column
    numero_colonne_fin_extract_nomcl_vincent = wsExtractNomcl.Range("Vincent").Columns.Count + numero_colonne_debut_extract_nomcl_vincent - 1
    
    lettre_affaire = ConvertToLetter(numero_colonne_affaire)
    lettre_rr = ConvertToLetter(numero_colonne_rr)
    lettre_rp = ConvertToLetter(numero_colonne_rp)
    lettre_debut_audric = ConvertToLetter(CLng(numero_colonne_debut_extract_nomcl_audric))
    lettre_fin_audric = ConvertToLetter(CLng(numero_colonne_fin_extract_nomcl_audric))
    lettre_debut_vincent = ConvertToLetter(CLng(numero_colonne_debut_extract_nomcl_vincent))
    lettre_fin_vincent = ConvertToLetter(CLng(numero_colonne_fin_extract_nomcl_vincent))
    
    '------------------------------------------------------------------------------------------
    'On supprime le texte et les mises en forme de toute la feuille Suivi
    
    For Each rw In wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Cells(Rows.Count, 2).End(xlUp).Row + 1)
        rw.Interior.Pattern = xlNone
        rw.Interior.TintAndShade = 1
        rw.Interior.PatternTintAndShade = 1
        rw.Interior.Color = RGB(255, 255, 255)
    Next
    wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Rows.Count).ClearContents
    wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Rows.Count).FormatConditions.Delete
    wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Rows.Count).Borders.LineStyle = xlLineStyleNone
    wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Rows.Count).Font.Bold = False
    wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Rows.Count).Font.Color = RGB(0, 0, 0)
    wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Rows.Count).HorizontalAlignment = xlCenter
    wsSuivi.Rows(ligne_debut_tableau_warnings & ":" & wsSuivi.Rows.Count).VerticalAlignment = xlCenter
    
    '------------------------------------------------------------------------------------------
    'Mettre � jour l'extract CMD
    
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
    'En VBA pour �conomiser un grand temps de calcul en lecture on lit toutes les donn�es d'un grand tableau (ici le tableau de la Feuil1 de l'extract CMD) une seule fois gr�ce � la commande Range.Value
    
    CMDRange = wsFeuil1.Range("A1:" & ConvertToLetter(CLng(Der_Colonne)) & Der_Ligne).Value
    
    '------------------------------------------------------------------------------------------
    'La premi�re boucle s'arr�te quand une cellule de la colonne "Numero affaire" du tableau pr�sent dans la feuille "Liste projets AR" est vide
    
    Do While (Not IsEmpty(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)))
    
    AffaireVoulue = wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)
    DateVoulue = wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_dv)
    
    If (Not IsEmpty(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_select))) Then
        
    '------------------------------------------------------------------------------------------
    'La condition v�rifie que l'Affaire voulue est pr�sente dans une des cases de la feuille "Extract nomcl projets" sinon on passe � l'affaire voulue suivante
       
        If (Not wsExtractNomcl.Columns(numero_colonne_debut_extract_nomcl_audric).Find(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)) Is Nothing) Or (Not wsExtractNomcl.Columns(numero_colonne_debut_extract_nomcl_vincent).Find(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)) Is Nothing) Then
               
    '------------------------------------------------------------------------------------------
    'La condition copie la plage de donn�es de la feuille Extract nomcl projets correspondant � l'Affaire voulue en regardant si le projet appartient � Audric ou Vincent
            
            If Not wsExtractNomcl.Columns(numero_colonne_debut_extract_nomcl_audric).Find(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)) Is Nothing Then
                
                Ligne_Affaire_Extract_Nomcl = wsExtractNomcl.Columns(numero_colonne_debut_extract_nomcl_audric).Find(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)).Row
                wsExtractNomcl.Range(lettre_debut_audric & Ligne_Affaire_Extract_Nomcl & ":" & lettre_fin_audric & Ligne_Affaire_Extract_Nomcl + 16).Copy
                wsSuivi.Range(lettre_debut_audric & indice_feuille_suivi).PasteSpecial (xlPasteAll)
                wsSuivi.Range(lettre_fin_audric & indice_feuille_suivi & ":" & lettre_fin_audric & indice_feuille_suivi + 16).Formula = wsExtractNomcl.Range(lettre_fin_audric & Ligne_Affaire_Extract_Nomcl & ":" & lettre_fin_audric & Ligne_Affaire_Extract_Nomcl + 16).Formula
                
            ElseIf Not wsExtractNomcl.Columns(numero_colonne_debut_extract_nomcl_vincent).Find(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)) Is Nothing Then
                
                Ligne_Affaire_Extract_Nomcl = wsExtractNomcl.Columns(numero_colonne_debut_extract_nomcl_vincent).Find(wsListeProjetsAR.Cells(indice_affaire_voulue, numero_colonne_av)).Row
                wsExtractNomcl.Range(lettre_debut_vincent & Ligne_Affaire_Extract_Nomcl & ":" & lettre_fin_vincent & Ligne_Affaire_Extract_Nomcl + 16).Copy
                wsSuivi.Range(lettre_debut_audric & indice_feuille_suivi).PasteSpecial (xlPasteAll)
                wsSuivi.Range(lettre_fin_audric & indice_feuille_suivi & ":" & lettre_fin_audric & indice_feuille_suivi + 16).Formula = wsExtractNomcl.Range(lettre_fin_vincent & Ligne_Affaire_Extract_Nomcl & ":" & lettre_fin_vincent & Ligne_Affaire_Extract_Nomcl + 16).Formula
            
            End If
    
    '------------------------------------------------------------------------------------------
    'Il y a 16 lignes par tableau correspondant � l'affaire voulue dans la feuille Extract nomcl projets
    
            indice_feuille_suivi = indice_feuille_suivi + 19
            ligne_debut_tableau_warnings = indice_feuille_suivi
            
    '------------------------------------------------------------------------------------------
    'En VBA pour �conomiser un grand temps de calcul en �criture on �crit toutes les donn�es en bloc en copiant une plage apr�s l'avoir lue, en la remplissant, puis en affectant la plage remplie � la plage d'origine
           
            ResultRange = wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi - 1 & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi - 1).Value
            
            ResultRange(1, 1) = "Affaire voulue"
            ResultRange(1, 2) = "Affaire source"
            ResultRange(1, 3) = "Rep�re"
            ResultRange(1, 4) = "D�signation"
            ResultRange(1, 5) = "Fournisseur / Fabriquant"
            ResultRange(1, 6) = "R�f�rence"
            ResultRange(1, 7) = "Distributeur"
            ResultRange(1, 8) = "R�f�rence distributeur"
            ResultRange(1, 9) = "Remarques"
            ResultRange(1, 10) = "Etat"
                            
            wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi - 1 & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi - 1).Value = ResultRange
            wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi - 1 & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi - 1).Font.Color = RGB(255, 255, 255)
            wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi - 1 & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi - 1).Interior.Color = RGB(0, 51, 153)
            wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi - 1 & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi - 1).Font.Bold = True
               
    '------------------------------------------------------------------------------------------
    'Cette boucle s'arr�te quand toutes les nomenclatures de la ligne ont �t� parcourues
            
            For j = numero_colonne_meca To numero_colonne_autre2
            
    '------------------------------------------------------------------------------------------
    'La condition v�rifie que les cellules des colonnes Nomenclatures ne sont pas vides
                
                If Not IsEmpty(wsListeProjetsAR.Cells(indice_affaire_voulue, j)) And wsListeProjetsAR.Cells(indice_affaire_voulue, j) <> "" Then
                    
    '------------------------------------------------------------------------------------------
    'Si la condition est v�rifi�e, la nomenclature est ouverte et les colonnes voulues en sortie sont recherch�es et leur num�ro de colonne affect�e � des variables
                    
                    Adresse_nom = wsListeProjetsAR.Cells(indice_affaire_voulue, j).Hyperlinks(1).Address
                    Set wkNom = Workbooks.Open(Adresse_nom, UpdateLinks:=1)
                    Set wsNom = wkNom.Worksheets("Nomenclature")
                    
                    numero_colonne_affaire_source = wsNom.Rows(2).Find("Affaire source").Column
                    numero_colonne_quantite = wsNom.Rows(2).Find("Quantit�").Column
                    
    '------------------------------------------------------------------------------------------
    'La colonne rep�re n'est pas toujours pr�sente dans les nomenclatures, d'o� la condition
                    
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
                                
                                ResultRange = wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi).Value
                                
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
                                
                                wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi).Value = ResultRange
                                
                                If UCase(Etat) = UCase("Etude") Then
                                    wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi).Interior.Color = RGB(192, 0, 0)
                                End If
                                If UCase(Etat) = UCase("Consult�") Then
                                    wsSuivi.Range(lettre_debut_tableau_nomenclatures & indice_feuille_suivi & ":" & lettre_fin_tableau_nomenclatures & indice_feuille_suivi).Interior.Color = RGB(255, 192, 0)
                                End If
                                
                                indice_feuille_suivi = indice_feuille_suivi + 1
                                wsSuivi.Range(lettre_debut_tableau_nomenclatures & (indice_feuille_suivi - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_feuille_suivi - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                                wsSuivi.Range(lettre_debut_tableau_nomenclatures & (indice_feuille_suivi - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_feuille_suivi - 1)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                            End If
                        End If
                    Next
                    
    '------------------------------------------------------------------------------------------
    'Bordure basse entre chaque ligne des nomenclatures
                    
                    wsSuivi.Range(lettre_debut_tableau_nomenclatures & (indice_feuille_suivi - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_feuille_suivi - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    wsSuivi.Range(lettre_debut_tableau_nomenclatures & (indice_feuille_suivi - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_feuille_suivi - 1)).Borders(xlEdgeBottom).Weight = xlThick
                    wsSuivi.Range(lettre_debut_tableau_nomenclatures & (indice_feuille_suivi - 1) & ":" & lettre_fin_tableau_nomenclatures & (indice_feuille_suivi - 1)).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                    wkNom.Close False
                    End If
                Next
            
            indice_feuille_suivi = indice_feuille_suivi + 2
            
            NomRange = wsSuivi.Range(lettre_debut_tableau_warnings & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Value
            
            NomRange(1, 1) = "Affaire"
            NomRange(1, 2) = "N� commande"
            NomRange(1, 3) = "Fournisseur"
            NomRange(1, 4) = "R�f�rence"
            NomRange(1, 5) = "Texte"
            NomRange(1, 6) = "Date AR"
            NomRange(1, 7) = "Date livraison"
            NomRange(1, 8) = "Commentaire"
            NomRange(1, 9) = "Quantit� restante � livrer"
            NomRange(1, 10) = "Date voulue"
            
            wsSuivi.Range(lettre_debut_tableau_warnings & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Value = NomRange
            wsSuivi.Range(lettre_debut_tableau_warnings & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Font.Color = RGB(255, 255, 255)
            wsSuivi.Range(lettre_debut_tableau_warnings & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Interior.Color = RGB(0, 51, 153)
            wsSuivi.Range(lettre_debut_tableau_warnings & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Font.Bold = True
            
            NomRange = wsSuivi.Range(lettre_rr & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi).Value
            
            NomRange(1, 1) = "Retard de r�ception Sym�trie (en jours)"
            NomRange(1, 2) = "Retard projet (en jours)"
            
            wsSuivi.Range(lettre_rr & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Value = NomRange
            wsSuivi.Range(lettre_rr & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Font.Color = RGB(255, 255, 255)
            wsSuivi.Range(lettre_rr & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Interior.Color = RGB(0, 51, 153)
            wsSuivi.Range(lettre_rr & indice_feuille_suivi - 1 & ":" & lettre_rp & indice_feuille_suivi - 1).Font.Bold = True
            
    '------------------------------------------------------------------------------------------
    'Du d�but � la fin de l'extract CMD
            
            ligne_debut_tableau_warnings = indice_feuille_suivi
            
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
                    
                    NomRange = wsSuivi.Range(lettre_debut_tableau_warnings & indice_feuille_suivi & ":" & lettre_rp & indice_feuille_suivi).Value
                    
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
                    
                    wsSuivi.Range(lettre_debut_tableau_warnings & indice_feuille_suivi & ":" & lettre_rp & indice_feuille_suivi).Value = NomRange
                           
    '------------------------------------------------------------------------------------------
    'On attribue les couleurs/barres de progression � chaque cas en priorisant AR sur livraison
                    If IsEmpty(DateAr) Then
                    
                        If Not IsEmpty(DateLiv) And DateToday >= CDate(DateLiv) And DateToday >= DateVoulue Then
                            
                            NomRange = wsSuivi.Range(lettre_rr & indice_feuille_suivi & ":" & lettre_rp & indice_feuille_suivi).Value
                            NomRange(1, 1) = (DateToday - CDate(DateLiv))
                            NomRange(1, 2) = DateToday - DateVoulue
                            wsSuivi.Range(lettre_rr & indice_feuille_suivi & ":" & lettre_rp & indice_feuille_suivi).Value = NomRange
                            wsSuivi.Cells(indice_feuille_suivi, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                            
                            If CDate(DateLiv) >= DateVoulue Then
                                Max = 1
                            Else
                                Max = DateVoulue - CDate(DateLiv)
                            End If
                            Cr�er_barres wsSuivi.Range(ConvertToLetter(numero_colonne_rr) & indice_feuille_suivi), Max, wsSuivi
        
                        ElseIf Not IsEmpty(DateLiv) And DateToday <= CDate(DateLiv) And DateToday >= DateVoulue Then
                            
                            NomRange = wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value
                            NomRange = CDate(DateLiv) - DateVoulue
                            wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value = NomRange
                            wsSuivi.Cells(indice_feuille_suivi, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                            
                        ElseIf Not IsEmpty(DateLiv) And DateToday <= DateVoulue And DateToday >= CDate(DateLiv) Then
                            
                            NomRange = wsSuivi.Range(lettre_rr & indice_feuille_suivi).Value
                            NomRange = (DateToday - CDate(DateLiv))
                            wsSuivi.Range(lettre_rr & indice_feuille_suivi).Value = NomRange
                            Max = CDate(DateLiv) - DateVoulue
                            Cr�er_barres wsSuivi.Range(ConvertToLetter(numero_colonne_rr) & indice_feuille_suivi), Max, wsSuivi
                            
                        ElseIf Not IsEmpty(DateLiv) And DateToday <= DateVoulue And DateToday <= CDate(DateLiv) And DateVoulue <= CDate(DateLiv) Then
                            
                            NomRange = wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value
                            NomRange = CDate(DateLiv) - DateVoulue
                            wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value = NomRange
                            wsSuivi.Cells(indice_feuille_suivi, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                        
                        End If
                    
                    Else
                        If Not IsEmpty(DateAr) And DateToday >= CDate(DateAr) And DateToday >= DateVoulue Then
                            
                            NomRange = wsSuivi.Range(lettre_rr & indice_feuille_suivi & ":" & lettre_rp & indice_feuille_suivi).Value
                            NomRange(1, 1) = (DateToday - CDate(DateAr))
                            NomRange(1, 2) = DateToday - DateVoulue
                            wsSuivi.Range(lettre_rr & indice_feuille_suivi & ":" & lettre_rp & indice_feuille_suivi).Value = NomRange
                            wsSuivi.Cells(indice_feuille_suivi, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                            
                            If CDate(DateAr) >= DateVoulue Then
                                Max = 1
                            Else
                                Max = DateVoulue - CDate(DateAr)
                            End If
                            Cr�er_barres wsSuivi.Range(lettre_rr & indice_feuille_suivi), Max, wsSuivi
                            
                        ElseIf Not IsEmpty(DateAr) And DateToday <= CDate(DateAr) And DateToday >= DateVoulue Then
                            
                            NomRange = wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value
                            NomRange = CDate(DateAr) - DateVoulue
                            wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value = NomRange
                            wsSuivi.Cells(indice_feuille_suivi, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                            
                        ElseIf Not IsEmpty(DateAr) And DateToday <= DateVoulue And DateToday >= CDate(DateAr) Then
                            
                            NomRange = wsSuivi.Range(lettre_rr & indice_feuille_suivi).Value
                            NomRange = (DateToday - CDate(DateAr))
                            wsSuivi.Range(lettre_rr & indice_feuille_suivi).Value = NomRange
                            Max = CDate(DateAr) - DateVoulue
                            Cr�er_barres wsSuivi.Range(lettre_rr & indice_feuille_suivi), Max, wsSuivi
                            
                        ElseIf Not IsEmpty(DateAr) And DateToday <= DateVoulue And DateToday <= CDate(DateAr) And DateVoulue <= CDate(DateAr) Then
                            
                            NomRange = wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value
                            NomRange = CDate(DateAr) - DateVoulue
                            wsSuivi.Range(lettre_rp & indice_feuille_suivi).Value = NomRange
                            wsSuivi.Cells(indice_feuille_suivi, numero_colonne_rp).Interior.Color = RGB(255, 242, 204)
                        
                        End If
                    End If
                    
    '------------------------------------------------------------------------------------------
    'Bordure entre chaque ligne des warnings
        
                wsSuivi.Range(lettre_affaire & ligne_debut_tableau_warnings & ":" & lettre_rp & indice_feuille_suivi - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                wsSuivi.Range(lettre_affaire & ligne_debut_tableau_warnings & ":" & lettre_rp & indice_feuille_suivi - 1).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
                
                indice_feuille_suivi = indice_feuille_suivi + 1
                
                End If
            
    '------------------------------------------------------------------------------------------
    'Bordure retards
            wsSuivi.Range(lettre_rr & ligne_debut_tableau_warnings & ":" & lettre_rr & indice_feuille_suivi - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Next
    
    '------------------------------------------------------------------------------------------
    'Bordure par affaire
            wsSuivi.Range(lettre_affaire & ligne_debut_tableau_warnings & ":" & lettre_rp & indice_feuille_suivi - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
            wsSuivi.Range(lettre_affaire & ligne_debut_tableau_warnings & ":" & lettre_rp & indice_feuille_suivi - 1).Borders(xlEdgeBottom).Weight = xlThick
            wsSuivi.Range(lettre_affaire & ligne_debut_tableau_warnings & ":" & lettre_rp & indice_feuille_suivi - 1).Borders(xlEdgeBottom).Color = RGB(0, 51, 153)
               
            indice_feuille_suivi = indice_feuille_suivi + 10
            ligne_debut_tableau_nomenclatures = indice_feuille_suivi
            
        End If
    End If
    indice_affaire_voulue = indice_affaire_voulue + 1
Loop


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

wsSuivi.Range("A2").FormulaArray = "=NO.SEMAINE.ISO(AUJOURDHUI())"
wkCMD.Close False

End Sub






