Sub ToggleInterpunction() '中英文标点互换
    Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant, strFind As String, strRep As String
    '定义一个中文标点的数组对象
    ChineseInterpunction = Array("。", "，", "；", "：", "？", "！", "……", "—", "～", "（", "）", "《", "》", "“")
    '定义一个英文标点的数组对象
    EnglishInterpunction = Array(".", ",", ";", ":", "?", "!", "…", "-", "~", "(", ")", "<", ">", "")
    strFind = "“(*)”"
    strRep = """\1"""
    Application.ScreenUpdating = False '关闭屏幕更新
    For N = 0 To UBound(ChineseInterpunction) '从数组的下标到上标间作一个循环
    With ActiveDocument.Content.Find
        .ClearFormatting '不限定查找格式
        .MatchWildcards = False '不使用通配符
        '查找相应的英文标点,替换为对应的中文标点
        .Execute findtext:=ChineseInterpunction(N), replacewith:=EnglishInterpunction(N), Replace:=wdReplaceAll
    End With
        Next
    With ActiveDocument.Content.Find
        .ClearFormatting '不限定查找格式
        .MatchWildcards = True '使用通配符
        .Execute findtext:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
    End With
End Sub