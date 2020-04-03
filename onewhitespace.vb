Sub OneWhiteSpace() '所有英文单词有且只有一个空格
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "(>) {2,}"
            .Replacement.Text = "\1 "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchByte = True
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .MatchFuzzy = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
End Sub
