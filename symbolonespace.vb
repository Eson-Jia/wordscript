Sub OneWhiteSpaceAfterSymbol() '保证在标点后面的有且仅有一个空格
    Dim Sta As String, TheSymbol As String
    Dim EnglishInterpunction() As Variant '(".", ",", ";", ":", "?", "!",")", "<", ">")
    EnglishInterpunction = Array("(\.)<([A-z]@)>", "(\,)<([A-z]@)>", "(\;)<([A-z]@)>", "(\:)<([A-z]@)>", "(\?)<([A-z]@)>", "(\!)<([A-z]@)>", "(\))<([A-z]@)>", "(\<)<([A-z]@)>", "(\>)<([A-z]@)>"
    '定义一个英文标点的数组对象
    Sta = "<([A-z]@)>"
    For N = 0 To UBound(EnglishInterpunction) '从数组的下标到上标间作一个循环
        With ActiveDocument.Content.Find
            .ClearFormatting '不限定查找格式
            .MatchWildcards = True '使用通配符
            .Execute findtext:=EnglishInterpunction(N), replacewith:="\1 \2", Replace:=wdReplaceAll
        End With
    Next
End Sub