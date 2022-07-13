Attribute VB_Name = "Module1"
Sub AcronymLister()
Application.ScreenUpdating = False

Dim oDoc_Source As Document
Dim oDoc_Target As Document
Dim strListSep As String
Dim strAcronym As String
Dim oTable As Table
Dim oRange As Range
Dim n As Long
Dim strAllFound As String
Dim Title As String
Dim Msg As String
Dim oCC As ContentControl

Dim StrTmp As String, StrAcronyms As String, i As Long, j As Long, k As Long, Rng As Range, Tbl As Table
StrAcronyms = "Acronym" & vbTab & "Definition" & vbTab & "Page" & vbCr
strListSep = Application.International(wdListSeparator)

strAllFound = "#"
With ActiveDocument
  With .Range
    With .Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .MatchWildcards = True
      .Wrap = wdFindStop
      .Text = "\([A-Z0-9][A-Zs0-9&]{1" & Application.International(wdListSeparator) & "}\)"
      '.Text = "<[A-Z]{2" & strListSep & "}>"
      .Replacement.Text = ""
      .Execute
    End With
    Do While .Find.Found = True
      StrTmp = Replace(Replace(.Text, "(", ""), ")", "")
      If (InStr(1, StrAcronyms, .Text, vbBinaryCompare) = 0) And (Not IsNumeric(StrTmp)) And InStr(1, strAllFound, "#" & StrTmp & "#") = 0 Then
        strAllFound = strAllFound & StrTmp & "#"
       ' And InStr(1, strAllFound, "#" & StrTmp & "#") = 0
       
        If (.Words.First.Previous.Previous.Words(1).Characters.First) = (Right(StrTmp, 1)) Then
          For i = Len(StrTmp) To 1 Step -1
            .MoveStartUntil Mid(StrTmp, i, 1), wdBackward
            .Start = .Start - 1
            If InStr(.Text, vbCr) > 0 Then
              .MoveStartUntil vbCr, wdForward
              .Start = .Start + 1
            End If
            If .Sentences.Count > 1 Then .Start = .Sentences.Last.Start
            If .Characters.Last.Information(wdWithInTable) = False Then
              If .Characters.First.Information(wdWithInTable) = True Then
                .Start = .Cells(.Cells.Count).Range.End + 1
              End If
            ElseIf .Cells.Count > 1 Then
              .Start = .Cells(.Cells.Count).Range.Start
            End If
          Next
        End If
        StrTmp = Replace(Replace(Replace(.Text, " (", "("), "(", "|"), ")", "")
        StrAcronyms = StrAcronyms & Split(StrTmp, "|")(1) & vbTab & Split(StrTmp, "|")(0) & vbTab & .Information(wdActiveEndAdjustedPageNumber) & vbCr
      End If
      .Collapse wdCollapseEnd
      .Find.Execute
    Loop
  End With
  StrAcronyms = Replace(Replace(Replace(StrAcronyms, " (", "("), "(", vbTab), ")", "")
  Set Rng = ActiveDocument.Range.Characters.Last
  With Rng
    If .Characters.First.Previous <> vbCr Then .InsertAfter vbCr
    .InsertAfter Chr(12)
    .Collapse wdCollapseEnd
    .Style = "Normal"
    .Text = StrAcronyms
     Set Tbl = .ConvertToTable(Separator:=vbTab, numrows:=.Paragraphs.Count, NumColumns:=3)
    ' Set Tbl = .Tables.Add(Separator:=vbTab, NumRows:=.Paragraphs.Count, NumColumns:=3)
    With Tbl
      .Columns.AutoFit
      .Rows(1).HeadingFormat = True
      .Rows(1).Range.Style = "Strong"
      .Rows.Alignment = wdAlignRowCenter
      
    End With
    .Collapse wdCollapseStart
  End With
  For i = 2 To Tbl.Rows.Count
    StrTmp = "": j = 0: k = 0
    With .Range
      With .Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = False
        .Forward = True
        .Text = Split(Tbl.Cell(i, 1).Range.Text, vbCr)(0)
        .MatchWildcards = True
        .Execute
      End With
      Do While .Find.Found = True
        If .InRange(Tbl.Range) Then Exit Do
        j = j + 1
        If j > 0 Then
         ' If k <> .Duplicate.Information(wdActiveEndAdjustedPageNumber) Then
          '  k = .Duplicate.Information(wdActiveEndAdjustedPageNumber)
          '  StrTmp = StrTmp & k & " "
            StrTmp = .Duplicate.Information(wdActiveEndAdjustedPageNumber)
            Exit Do
         ' End If
        End If
        .Collapse wdCollapseEnd
        .Find.Execute
      Loop
    End With
  Next
End With

With Selection
    ' .Sort ExcludeHeader:=True, FieldNumber:="Column 1", SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending
    Tbl.SortAscending
    
            
    'Go to start of document
    .HomeKey (wdStory)
End With

Application.ScreenUpdating = True

Set Rng = Nothing: Set Tbl = Nothing
End Sub
