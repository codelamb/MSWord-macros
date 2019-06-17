Attribute VB_Name = "DateChangeISOtoRussian"
Sub GetDateAndReplace()
    Dim FoundOne As Boolean

    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    FoundOne = True ' loop at least once

    Do While FoundOne ' loop until no date is found
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "([0-9]{4})[-]([0-9]{1;2})[-]([0-9]{1;2})"
            .Format = True
            .Forward = True
            .MatchWildcards = True
        End With

        Selection.Find.Execute Replace:=wdReplaceNone

        ' check the find to be sure it's a date
        If IsDate(Selection.Text) Then
            Selection.Text = Format(Selection.Text, "dd.mm.yyyy")
            Selection.Collapse wdCollapseEnd
        Else ' not a date - end loop
            FoundOne = False
        End If
    Loop
End Sub
