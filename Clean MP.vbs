Sub MaxiMP()
Selection.HomeKey Unit:=wdStory

'Je sais que c'est mal codé, c'est vraiment bourrin et "long" mais je prévoit déjà JeanMichel je sais pas l'utiliser, plus c'est bourrin plus ce sera simple de corriger au téléphone sans se prendre le crâne. 
   'Et accessoirement je vois pas trop comment l'optimiser, c'est à dire qu'à la base je fais du fiscal et 3 tableurs, Word c'est un copain mais pas une passion dévorante. 

'Liste des transformations : 
'Abréviation de page, standardisation du format p.1 ou p. 1 
'Suppression des doubles Espaces 
'Espace avant un point = on supprime l'espace 
'Espace avant une virgule = on supprime l'espace 
'Espace avant deux points = insécable devant deux points 
'Espace avant un point virgule = insécable avant un point virgule 
'Pas d'espace après un guillemet ouvrant = insécable après un guillement ouvrant
'Pas d'espace avant un guillemet fermant = insécable avant un guillement fermant
'Espace avant une parenthèse fermante : pas d'espace avant une parenthèse fermante
'Espace après une parenthèse ouvrante : pas d'espace après une parenthèse ouvrante
'Espace avant un crochet fermant : pas d'espace avant un crochet fermant
'Espace après un crochet ouvrant  : pas d'espace un crochet ouvrant
'Remplacement des espaces avant les chiffres
   'Espace avant chiffre 1 = Insécable devant chiffre 1
   'Espace avant chiffre 2 = Insécable devant chiffre 2 
   'Espace avant chiffre 3 = Insécable devant chiffre 3
   'Espace avant chiffre 4 = Insécable devant chiffre 4
   'Espace avant chiffre 5 = Insécable devant chiffre 5
   'Espace avant chiffre 6 = Insécable devant chiffre 6
   'Espace avant chiffre 7 = Insécable devant chiffre 7
   'Espace avant chiffre 8 = Insécable devant chiffre 8
   'Espace avant chiffre 9 = Insécable devant chiffre 9 
   'Espace avant chiffre 0 = Insécable devant chiffre 0

'Espace avant un point d'interrogation = insécable devant un point d'interrogation
'Espace avant un insécable = insécable simple
'Double insécable = Insécable simple
'Un autre petit coup de double espace
'Passage des Ibid. en italique
'Passage des op. cit. en italique. 
'Passage de "via" en italique
'Passage de a priori en italique

'----------
Dim n 
n=0

'Choix du domaine d'action
Dim c As Integer
c = InputBox("Vous pouvez traiter : le texte seulement (Tapez 1), les notes de bas de page seulemen (tapez 2) ou les deux (Tapez 3)", "Champ d'action", 1)
intVar = CInt(c)


'Abréviation de page, standardisation du format p.1 ou p. 1
Dim p As Integer
p = InputBox("Ajouter un espace après le point de ''p.1'', tapez 1, sinon tapez 2 ", "Format Page", 1)
intVar = CInt(p)




If c = 1 Or c=3 Then
Msgbox ("Traitement du corps du texte, cliquez sur ''Ok'' pour lancer le processus")
'Traitement du corps du texte
ActiveDocument.Range.Select
If p = 2 Then
 'Page1
ActiveDocument.Range.Select
With Selection.Find
    .Text = " p."
    .Replacement.Text = " p. "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With
End If
If p = 1 Then
 'Page2
ActiveDocument.Range.Select
With Selection.Find
    .Text = " p. "
    .Replacement.Text = " p."
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With
End If

'--

 'Suppression des doubles Espaces
ActiveDocument.Range.Select
With Selection.Find
    .Text = "  "
    .Replacement.Text = " "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

'Espace avant un point = on supprime l'espace
ActiveDocument.Range.Select
With Selection.Find
    .Text = " ."
    .Replacement.Text = "."
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant une virgule = on supprime l'espace
ActiveDocument.Range.Select
With Selection.Find
    .Text = " ,"
    .Replacement.Text = ","
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant deux points = insécable devant deux points
ActiveDocument.Range.Select
With Selection.Find
    .Text = " :"
    .Replacement.Text = "^s:"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un point virgule = insécable avant un point virgule
ActiveDocument.Range.Select
With Selection.Find
    .Text = " ;"
    .Replacement.Text = "^s;"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Pas d'espace après un guillemet ouvrant = insécable après un guillement ouvrant
ActiveDocument.Range.Select
With Selection.Find
    .Text = "«"
    .Replacement.Text = " «^s"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Pas d'espace avant un guillemet fermant = insécable avant un guillement fermant
ActiveDocument.Range.Select
With Selection.Find
    .Text = "»"
    .Replacement.Text = "^s» "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant une parenthèse fermante : pas d'espace avant une parenthèse fermante
ActiveDocument.Range.Select
With Selection.Find
    .Text = " )"
    .Replacement.Text = ") "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

  'Espace après une parenthèse ouvrante : pas d'espace après une parenthèse ouvrante
ActiveDocument.Range.Select
With Selection.Find
    .Text = "( "
    .Replacement.Text = " ("
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un crochet fermant : pas d'espace avant un crochet fermant 
ActiveDocument.Range.Select
With Selection.Find
    .Text = " ]"
    .Replacement.Text = "] "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

  'Espace après un crochet ouvrant  : pas d'espace un crochet ouvrant
ActiveDocument.Range.Select
With Selection.Find
    .Text = "[ "
    .Replacement.Text = " ["
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

'Remplacement des espaces avant les chiffres
 'Espace avant chiffre 1 = Insécable devant chiffre 1
ActiveDocument.Range.Select
With Selection.Find
    .Text = " :"
    .Replacement.Text = "^s:"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 2 = Insécable devant chiffre 2
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 2"
    .Replacement.Text = "^s2"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 3 = Insécable devant chiffre 3
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 3"
    .Replacement.Text = "^s3"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 4 = Insécable devant chiffre 4
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 4"
    .Replacement.Text = "^s4"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 5 = Insécable devant chiffre 5
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 5"
    .Replacement.Text = "^s5"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 6 = Insécable devant chiffre 6
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 6"
    .Replacement.Text = "^s6"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 7= Insécable devant chiffre 7
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 7"
    .Replacement.Text = "^s7"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 8 = Insécable devant chiffre 8
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 8"
    .Replacement.Text = "^s8"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 9 = Insécable devant chiffre 9
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 9"
    .Replacement.Text = "^s9"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 0= Insécable devant chiffre 0
ActiveDocument.Range.Select
With Selection.Find
    .Text = " 0"
    .Replacement.Text = "^s0"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un point d'interrogation = insécable devant un point d'interrogation
ActiveDocument.Range.Select
With Selection.Find
    .Text = " ?"
    .Replacement.Text = "^s?"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un insécable = insécable simple
ActiveDocument.Range.Select
With Selection.Find
    .Text = " ^s"
    .Replacement.Text = "^s"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Double insécable = Insécable simple
ActiveDocument.Range.Select
With Selection.Find
    .Text = "^s^s"
    .Replacement.Text = "^s"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Un autre petit coup de double espace, plusieurs fois pour éviter les problèmes
 While n<10
ActiveDocument.Range.Select
With Selection.Find
    .Text = "  "
    .Replacement.Text = " "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With
n=n+1
Wend

'--

'On passe tous les Ibid en italique et on leur rajoute un point
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .Text = "ibid"
        .Replacement.Text = "Ibid."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '--
    
    'On passe tous les op. cit. en italique
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .Text = "op. cit."
        .Replacement.Text = "op. cit."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '--

    'On passe tous les via en italique
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .Text = " via "
        .Replacement.Text = " via "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '--
    
    'On passe tous les a priori en italique
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = False
    With Selection.Find
        .Text = " "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

   
   
   Msgbox ("Fin du traitement du corps du texte, cliquez sur ''Ok'' pour terminer ou lancer le traitement des notes de bas de page selon votre choix initial")
   
   
   End if 
    
  
    
    
    
    
    If c = 2 Or c=3 Then
    Msgbox ("Traitement des notes de bas de page, cliquez sur ''Ok'' pour lancer le processus")
    'Traitement des notes de bas de page
    Dim xDoc As Document
    Dim xRange As Range
    Set xDoc = ActiveDocument
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
    
    
If p = 1 Then
 'Page1
     If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    Selection.Find.Replacement.Font.Italic = False
    .Text = " p."
    .Replacement.Text = " p. "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With
End If
If p = 2 Then
 'Page2
     If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " p. "
    .Replacement.Text = " p."
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With
End If

'--

 'Suppression des doubles Espaces
     If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = "  "
    .Replacement.Text = " "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

'Espace avant un point = on supprime l'espace
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " ."
    .Replacement.Text = "."
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant une virgule = on supprime l'espace
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " ,"
    .Replacement.Text = ","
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant deux points = insécable devant deux points
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " :"
    .Replacement.Text = "^s:"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un point virgule = insécable avant un point virgule
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " ;"
    .Replacement.Text = "^s;"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Pas d'espace après un guillemet ouvrant = insécable après un guillement ouvrant
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = "«"
    .Replacement.Text = " «^s"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Pas d'espace avant un guillemet fermant = insécable avant un guillement fermant
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = "»"
    .Replacement.Text = "^s» "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant une parenthèse fermante : pas d'espace avant une parenthèse fermante
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " )"
    .Replacement.Text = ") "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

  'Espace après une parenthèse ouvrante : pas d'espace après une parenthèse ouvrante
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = "( "
    .Replacement.Text = " ("
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un crochet fermant : pas d'espace avant un crochet fermant 
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " ]"
    .Replacement.Text = "] "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

  'Espace après un crochet ouvrant  : pas d'espace un crochet ouvrant
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = "[ "
    .Replacement.Text = " ["
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

'Remplacement des espaces avant les chiffres
 'Espace avant chiffre 1 = Insécable devant chiffre 1
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " :"
    .Replacement.Text = "^s:"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 2 = Insécable devant chiffre 2
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 2"
    .Replacement.Text = "^s2"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 3 = Insécable devant chiffre 3
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 3"
    .Replacement.Text = "^s3"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 4 = Insécable devant chiffre 4
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 4"
    .Replacement.Text = "^s4"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 5 = Insécable devant chiffre 5
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 5"
    .Replacement.Text = "^s5"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 6 = Insécable devant chiffre 6
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 6"
    .Replacement.Text = "^s6"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 7= Insécable devant chiffre 7
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 7"
    .Replacement.Text = "^s7"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 8 = Insécable devant chiffre 8
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 8"
    .Replacement.Text = "^s8"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 9 = Insécable devant chiffre 9
 If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 9"
    .Replacement.Text = "^s9"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant chiffre 0= Insécable devant chiffre 0
  If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " 0"
    .Replacement.Text = "^s0"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un point d'interrogation = insécable devant un point d'interrogation
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " ?"
    .Replacement.Text = "^s?"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Espace avant un insécable = insécable simple
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = " ^s"
    .Replacement.Text = "^s"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Double insécable = Insécable simple
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = "^s^s"
    .Replacement.Text = "^s"
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

'--

 'Un autre petit coup de double espace, plusieurs fois pour éviter les problèmes
 While n<10
    If xDoc.Footnotes.Count > 0 Then
        Set xRange = xDoc.Footnotes(1).Range
        xRange.WholeStory
        xRange.Select
    End If
With Selection.Find
    .Text = "  "
    .Replacement.Text = " "
    .Forward = True
    .ClearFormatting
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With
n=n+1
Wend

'--

'On passe tous les Ibid en italique et on leur rajoute un point
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .Text = "ibid"
        .Replacement.Text = "Ibid."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '--
    
    'On passe tous les op. cit. en italique
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .Text = "op. cit."
        .Replacement.Text = "op. cit."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '--

    'On passe tous les via en italique
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .Text = " via "
        .Replacement.Text = " via "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '--
    
    'On passe tous les a priori en italique
   Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = True
    With Selection.Find
        .Text = " a priori "
        .Replacement.Text = " a priori "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Msgbox ("Fin du traitement des notes de bas de page, cliquez sur ''Ok'' pour terminer.")
    End If
    
End Sub
