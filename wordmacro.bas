Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.MoveDown Unit:=wdLine, Count:=6, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=48, Extend:=wdExtend
    Selection.Copy
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=5
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=8
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=67
    Selection.TypeText Text:="(a) "
    Selection.EndKey Unit:=wdLine
    Selection.TypeText Text:=" (b)"
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=" (c)"
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine
    Selection.TypeText Text:=" (d)"
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.EndKey Unit:=wdLine
    Selection.TypeText Text:=" (e)"
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:=" (f)"
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.InsertCrossReference ReferenceType:="Numbered item", _
        ReferenceKind:=wdNumberRelativeContext, ReferenceItem:="14", _
        InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, _
        SeparatorString:=" "
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.Copy
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=4
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" (b)"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="(c)"
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.MoveLeft Unit:=wdWord, Count:=2
    CommandBars("Research").Visible = False
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
    Selection.InlineShapes.AddPicture FileName:= _
        "C:\Users\Jalal\Documents\Mphil\DME structures\NaF\3DME\NaF(3DME) TS-f jn2287.png" _
        , LinkToFile:=True, SaveWithDocument:=True
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' Macro3 Macro
'
'
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.MoveUp Unit:=wdLine, Count:=2
    Selection.MoveRight Unit:=wdCharacter, Count:=2
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' Macro4 Macro
'
For Count = 1 To 10
 Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.EndKey Unit:=wdLine
    Selection.TypeText Text:=" "
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.EndKey Unit:=wdLine
    
Next
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro5"
'
' Macro5 Macro
'
'
    Selection.Font.Grow
    Selection.Font.Grow
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Font.Size = 10
    Selection.TypeText Text:="d"
    Selection.Font.Size = 12
    Selection.TypeBackspace
    ActiveDocument.Save
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro6"
'
' Macro6 Macro
'
'
    Selection.Font.Superscript = wdToggle
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro7"
'
' Macro7 Macro
'
'
End Sub
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro8"
'
' Macro8 Macro
'
'
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=" "
End Sub
Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro9"
'
' Macro9 Macro
'
'
    Selection.PasteExcelTable False, False, False
End Sub
Sub Macro10()
Attribute Macro10.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro10"
'
' Macro10 Macro
'
'
End Sub
Sub Macro11()
Attribute Macro11.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro11"
'
' Macro11 Macro
'
'
    Selection.PasteAndFormat (wdTableOriginalFormatting)
End Sub

Sub Macro12()
Attribute Macro12.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro12"
'
' Macro12 Macro
'
'
    
    For Count = 1 To 100
    If Count = 1 Then
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Superscript = True
        .Subscript = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    End If
    
    If Count = 1 Then
    Selection.Find.Execute
    Selection.MoveRight
    
    startloc = Selection.Range.Start
    Do
    Selection.Find.Execute
    Loop Until IsNumeric(Selection.Text) = True
    
    Selection.MoveLeft
    endloc = Selection.Range.Start
    
    Set rng = ActiveDocument.Range(startloc, endloc)
    
    rng.Select
    
    rng.Delete
    Else
    Selection.Find.Execute
    Selection.MoveRight
    Selection.TypeText " "
    
      startloc = Selection.Start
         Selection.Find.Execute
'         Selection.Find.Execute
    Selection.MoveLeft
    endloc = Selection.Range.Start
    
    Set rng = ActiveDocument.Range(startloc, endloc)
    
    rng.Select
    rng.Delete
    
    End If
    
    Next
  
End Sub

Sub macrobracket()

 With Selection.Find
        .Text = "(\[{2})(*)(\]{2})"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
   
    Selection.Find.Execute
    Selection.MoveRight
    
    startloc = Selection.Range.Start
    Do
    Selection.Find.Execute
    Loop Until IsNumeric(Selection.Text) = True
    
    Selection.MoveLeft
    endloc = Selection.Range.Start
    
    Set rng = ActiveDocument.Range(startloc, endloc)
    
    rng.Select
    
    rng.Delete
End Sub
