Sub ToggleInterpunction() '中英文标点互换
    Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
    '定义一个中文标点的数组对象
    ChineseInterpunction = Array("．", "，", "；", "：", "？", "！", "……", "〜", "（", "）")
    '定义一个英文标点的数组对象
    EnglishInterpunction = Array(".", ",", ";", ":", "?", "!", "…", "~", "(", ")")
    Application.ScreenUpdating = False '关闭屏幕更新
    For N = 0 To UBound(ChineseInterpunction) '从数组的下标到上标间作一个循环
    With ActiveDocument.Content.Find
        .ClearFormatting '不限定查找格式
        .MatchWildcards = False '不使用通配符
        '查找相应的英文标点,替换为对应的中文标点
        .Execute findtext:=ChineseInterpunction(N), replacewith:=EnglishInterpunction(N), Replace:=wdReplaceAll
    End With
        Next
End Sub

Sub OneWhiteSpace() '所有英文标点和后面的单词有且只有一个空格
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

Sub OneWhiteSpaceAfterSymbol() '保证在标点后面的有且仅有一个空格
    Dim EnglishInterpunction() As Variant '(".", ",", ";", ":", "?", "!",")", "<", ">")
    EnglishInterpunction = Array("(\.)<([A-z'0-9]@)>", "(\,)<([A-z'0-9]@)>", "(\;)<([A-z'0-9]@)>", "(\:)<([A-z'0-9]@)>", "(\?)<([A-z'0-9]@)>", "(\!)<([A-z'0-9]@)>", "(\))<([A-z'0-9]@)>", "(\<)<([A-z'0-9]@)>", "(\>)<([A-z'0-9]@)>")
    '定义一个英文标点的数组对象
    For N = 0 To UBound(EnglishInterpunction) '从数组的下标到上标间作一个循环
        With ActiveDocument.Content.Find
            .ClearFormatting '不限定查找格式
            .MatchWildcards = True '使用通配符
            .Execute findtext:=EnglishInterpunction(N), replacewith:="\1 \2", Replace:=wdReplaceAll
        End With
    Next
End Sub

Sub LeftQuotaWhite() '保证在标点后面的有且仅有一个空格
    With ActiveDocument.Content.Find
            .ClearFormatting '不限定查找格式
            .MatchWildcards = True '使用通配符
            .Execute findtext:=" ”", replacewith:="”", Replace:=wdReplaceAll
    End With
End Sub

Sub ClearLeftBraceWhite() '左括号左边的空格
    With ActiveDocument.Content.Find
            .ClearFormatting '不限定查找格式
            .MatchWildcards = True '使用通配符
            .Execute findtext:=" \(", replacewith:="(", Replace:=wdReplaceAll
    End With
End Sub

Sub ReplaceTable() 'replace table with whitespace
    With ActiveDocument.Content.Find
    .ClearFormatting
    .MatchWildcards = True
    .Execute findtext:="^t", replacewith:=" ", Replace:=wdReplaceAll
    End With
End Sub

Sub PruneWithSpaceWithShortLine() '有一个空格
     PatternList = Array("(_{3,})", " (_{3,})", "(_{3,}) ", " (_{3,}) ")
    For N = 0 To UBound(PatternList) '从数组的下标到上标间作一个循环
        With ActiveDocument.Content.Find
            .ClearFormatting '不限定查找格式
            .MatchWildcards = True '使用通配符
            .Execute findtext:=PatternList(N), replacewith:="____", Replace:=wdReplaceAll
        End With
    Next
End Sub
Sub RemoveSpaceBeforeSymbol()
    With ActiveDocument.Content.Find
    .ClearFormatting
    .MatchWildcards = True
    .Execute findtext:="> ([\,\.\?\!])", replacewith:="\1", Replace:=wdReplaceAll
    End With
End Sub

Sub RemoveBCDSpace()
    With ActiveDocument.Content.Find
    .ClearFormatting
    .MatchWildcards = True
    .Execute findtext:=" @([B-D].)", replacewith:=" \1", Replace:=wdReplaceAll
    End With
End Sub
Sub Main()
    ToggleInterpunction
    OneWhiteSpaceAfterSymbol
    OneWhiteSpace
    LeftQuotaWhite
    ClearLeftBraceWhite
    ReplaceTable
    PruneWithSpaceWithShortLine
    RemoveSpaceBeforeSymbol
    RemoveBCDSpace
End Sub
