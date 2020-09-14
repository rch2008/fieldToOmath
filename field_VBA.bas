Attribute VB_Name = "field_VBA"
Private mMatchs As Object

Sub fieldToOmath()
    Dim str As String
    Dim re As Object
    'Dim mMatch As Object, mmatchs As Object
    Dim field As field
    Dim finalCMD As String
    Dim strTemp As String
    Dim index As Long
    Dim strCmd As String, strCmdB As String, strText As String, strXL As String, strBrace As String, strFH As String
    
    strCmd = "\\([A-Za-z0-9])+"                     '命令
    strCmdB = "\\[\(\)\{\}\[\]\|\*\,\|]"            '括号命令
    strText = "[^\s\\ \(\)\{\}\[\]\|,]+"            '纯文本
    strXL = "[φπ]"                                '希腊
    strBrace = "\(|\)|\{|\}|\[|\]|\|"                '括号
    strFH = ","                                     '逗号
    
    Set re = New RegExp
    re.Global = True
    cmdFlag = 3        '0 主命令 1 辅命令  2 开组 3文本 4 ","分割 5 关组
    For Each field In ActiveDocument.Fields
        finalCMD = ""
        str = Trim(field.Code)
        If UCase(Left(str, 2)) = "EQ" Then
            str = Mid(str, 4)
            ''''''''''''''命令'''''|'''括号命令''''|''''纯文本'''''|''''希腊'''''|''''''括号''''''|''''逗号
            re.Pattern = strCmd + "|" + strCmdB + "|" + strText + "|" + strXL + "|" + strBrace + "|" + strFH
            Set mMatchs = re.Execute(str)
            For index = 0 To mMatchs.count - 1
                If exeCMD(strTemp, index) Then
                    finalCMD = finalCMD + strTemp
                Else
                    GoTo nextField
                End If
            Next
            field.Select
            field.Delete
            typeCMD finalCMD
        End If
nextField:
    Next
    allChangeToVec
End Sub

Function getCMD(ByVal index As Long, ByRef mFlag As Boolean) As String
    getCMD = mMatchs(index)
    mFlag = isMcmd(getCMD)
End Function

Function isMcmd(ByVal str As String) As Boolean
    str = UCase(str)
    If str = "\A" Or str = "\B" Or str = "\D" Or str = "\F" Or str = "\L" Or str = "\O" Or str = "\R" Or str = "\S" Or str = "\X" Or str = "\I" Then
        isMcmd = True
    Else
        isMcmd = False
    End If
End Function

Function typeCMD(str As String)
    Dim cmd() As String
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText Text:=str 'cmd(i)
    Selection.OMaths.BuildUp
    Selection.MoveRight Unit:=wdCharacter, count:=1
End Function


Function exeCMD(ByRef str As String, ByRef index As Long) As Boolean
    Dim cmd As String
    exeCMD = False
    
    cmd = CStr(mMatchs(index).Value)
    If isMcmd(cmd) Then
        If UCase(cmd) = "\A" Then
            cmdFlag = 0
            If cmdA(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\B" Then
            cmdFlag = 0
            If cmdB(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\D" Then
            cmdFlag = 0
            If cmdD(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\F" Then
            cmdFlag = 0
            If cmdF(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\I" Then
            cmdFlag = 0
            If cmdI(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\L" Then
            cmdFlag = 0
            If cmdL(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\O" Then
            cmdFlag = 0
            If cmdO(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\R" Then
            cmdFlag = 0
            If cmdR(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\S" Then
            cmdFlag = 0
            If cmdS(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        ElseIf UCase(cmd) = "\X" Then
            cmdFlag = 0
            If cmdF(cmd, index) Then
                str = cmd
                exeCMD = True
            End If
        End If
    Else            '文本，（括号，纯文本，逗号）
        str = cmd
        exeCMD = True
    End If
End Function

Function cmdA(ByRef cmd As String, ByRef index As Long) As Boolean
    Dim str As String
    Dim mFlag As Boolean
    Dim cmdFlag As Integer
    Dim scr As String
    Dim mat() As String
    Dim count As Integer
    Dim co As Integer
    
    cmdFlag = 0
    count = 0
    cmd = ""
    Do While cmdFlag <> 5
        index = index + 1
        str = getCMD(index, mFlag)
        If mFlag = True Then
        '''''''''参数为主命令，先执行
            If exeCMD(str, index) = False Then
                cmdA = False
                Exit Function
            End If
        End If
        
        If cmdFlag = 0 Then
            If str = "(" Then
                If co = 0 Then co = 1
                cmdFlag = 2
            Else
                If Left(UCase(str), 3) = "\CO" Then
                    co = CInt(Mid(str, 4))
                End If
            End If
        ElseIf cmdFlag = 2 Then
            If str = "(" Then
                count = count + 1
                scr = scr + str
            ElseIf str = ")" Then
                If count > 0 Then
                    scr = scr + str
                    count = count - 1
                Else
                    cmdFlag = 5
                End If
            ElseIf str = "\(" Then
                scr = scr + "("
            ElseIf str = "\)" Then
                scr = scr + ")"
            ElseIf str = "\," Then
                scr = scr + ","
            ElseIf str = "," Then
                scr = scr + Chr(0)
            Else
                scr = scr + str
            End If
        End If
    Loop
    mat = Split(scr, Chr(0))
    For i = 0 To UBound(mat)
        cmd = cmd + mat(i)
        If (i + 1) Mod co = 0 Then
            cmd = cmd + "@"
        Else
            cmd = cmd + "&"
        End If
    Next
    j = i Mod co
    If j = 0 Then
        cmd = Left(cmd, Len(cmd) - 1)
    Else
        i = 0
        j = co - j - 1
        Do While i < j
            cmd = cmd + "&"
            i = i + 1
        Loop
    End If
    cmd = "■(" + cmd + ")"
    cmdA = True
End Function

Function cmdB(ByRef cmd As String, ByRef index As Long) As Boolean
    Dim str As String
    Dim mFlag As Boolean
    Dim cmdFlag As Integer
    Dim scr As String
    Dim count As Integer
    Dim i As Integer
    Dim lr(1) As String
    
    cmdFlag = 0
    count = 0
    cmd = ""
    Do While cmdFlag <> 5
        index = index + 1
        str = getCMD(index, mFlag)
        If cmdFlag <> 0 And mFlag = True Then
        '''''''''参数为主命令，先执行
            If exeCMD(str, index) = False Then
                cmdB = False
                Exit Function
            End If
        End If
        
        If cmdFlag = 0 Then
            If str = "(" Then
                If lr(0) = "" And lr(1) = "" Then
                    lr(0) = "("
                    lr(1) = ")"
                Else
                    If lr(0) = "" And lr(1) <> "" Then
                        lr(0) = "├"
                    ElseIf lr(0) <> "" And lr(1) = "" Then
                        lr(1) = "┤"
                    End If
                End If
                cmdFlag = 2
            Else
                If Left(UCase(str), 3) = "\LC" Then
                    i = 0
                ElseIf Left(UCase(str), 3) = "\RC" Then
                    i = 1
                ElseIf Left(UCase(str), 3) = "\BC" Then
                    i = 2
                Else
                    If i = 2 Then
                        If Mid(str, 2) = "(" Then
                            lr(0) = "("
                            lr(1) = ")"
                        ElseIf Mid(str, 2) = "[" Then
                            lr(0) = "["
                            lr(1) = "]"
                        ElseIf Mid(str, 2) = "{" Then
                            lr(0) = "{"
                            lr(1) = "}"
                        Else
                            lr(0) = Mid(str, 2)
                            lr(1) = Mid(str, 2)
                        End If
                    Else
                        lr(i) = Mid(str, 2)
                    End If
                End If
            End If
        ElseIf cmdFlag = 2 Then
            If str = "(" Then
                count = count + 1
                scr = scr + str
            ElseIf str = ")" Then
                If count > 0 Then
                    scr = scr + str
                    count = count - 1
                Else
                    cmdFlag = 5
                End If
            ElseIf str = "\(" Then
                scr = scr + "("
            ElseIf str = "\)" Then
                scr = scr + ")"
            Else
                scr = scr + str
            End If
        End If
    Loop
    cmd = "" + lr(0) + scr + lr(1) + ""
    cmdB = True
End Function

Function cmdD(ByRef cmd As String, ByRef index As Long) As Boolean
    cmd = "none"
End Function

Function cmdF(ByRef cmd As String, ByRef index As Long) As Boolean
    'Dim index As Long
    Dim str As String
    Dim mFlag As Boolean
    Dim cmdFlag As Integer
    Dim num As String
    Dim den As String
    cmdFlag = 0
    'i = j
    cmd = ""
    Do While cmdFlag <> 5
        index = index + 1
        str = getCMD(index, mFlag)
        If mFlag = True Then
        '''''''''参数为主命令，先执行
            If exeCMD(str, index) = False Then
                cmdF = False
                Exit Function
            End If
        End If
        
        If cmdFlag = 0 Then
            If str = "(" Then
                cmdFlag = 2
            Else
                cmdFlag = 1
            End If
        ElseIf cmdFlag = 1 Then
        ElseIf cmdFlag = 2 Then
            If str <> "," Then
                num = num + str
            ElseIf str = "," Then
                cmdFlag = 4
            End If
        ElseIf cmdFlag = 4 Then
            If str <> ")" Then
                den = den + str
            ElseIf str = ")" Then
                cmdFlag = 5
            End If
        End If
    Loop
    cmd = cmd + "(" + num + ")/(" + den + ")"
    cmdF = True
End Function

Function cmdI(ByRef cmd As String, ByRef index As Long) As Boolean
    
End Function

Function cmdL(ByRef cmd As String, ByRef index As Long) As Boolean
    Dim str As String
    Dim mFlag As Boolean
    Dim cmdFlag As Integer
    Dim scr As String
    Dim count As Integer
    
    cmdFlag = 0
    count = 0
    cmd = ""
    Do While cmdFlag <> 5
        index = index + 1
        str = getCMD(index, mFlag)
        If mFlag = True Then
        '''''''''参数为主命令，先执行
            If exeCMD(str, index) = False Then
                cmdL = False
                Exit Function
            End If
        End If
        
        If cmdFlag = 0 Then
            If str = "(" Then
                cmdFlag = 2
            End If
        ElseIf cmdFlag = 2 Then
            If str = "(" Then
                count = count + 1
                scr = scr + str
            ElseIf str = ")" Then
                If count > 0 Then
                    scr = scr + str
                    count = count - 1
                Else
                    cmdFlag = 5
                End If
            ElseIf str = "\(" Then
                scr = scr + "("
            ElseIf str = "\)" Then
                scr = scr + ")"
            ElseIf str = "\," Then
                scr = scr + ","
            Else
                scr = scr + str
            End If
        End If
    Loop
    cmd = scr '"〖" + scr + "〗"
    cmdL = True
    
End Function

Function cmdO(ByRef cmd As String, ByRef index As Long) As Boolean
    Dim str As String
    Dim mFlag As Boolean
    Dim cmdFlag As Integer
    Dim scr As String
    Dim mat() As String
    Dim count As Integer
    
    cmdFlag = 0
    count = 0
    cmd = ""
    Do While cmdFlag <> 5
        index = index + 1
        str = getCMD(index, mFlag)
        If mFlag = True Then
        '''''''''参数为主命令，先执行
            If exeCMD(str, index) = False Then
                cmdO = False
                Exit Function
            End If
        End If
        
        If cmdFlag = 0 Then
            If str = "(" Then
                cmdFlag = 2
            End If
        ElseIf cmdFlag = 2 Then
            If str = "(" Then
                count = count + 1
                scr = scr + str
            ElseIf str = ")" Then
                If count > 0 Then
                    scr = scr + str
                    count = count - 1
                Else
                    cmdFlag = 5
                End If
            ElseIf str = "\(" Then
                scr = scr + "("
            ElseIf str = "\)" Then
                scr = scr + ")"
            ElseIf str = "\," Then
                scr = scr + ","
            ElseIf str = "," Then
                scr = scr + Chr(0)
            Else
                scr = scr + str
            End If
        End If
    Loop
    mat = Split(scr, Chr(0))
    If UBound(mat) = 1 Then
        cmdO = True
        If mat(0) = "→" Then
            cmd = "╠" + mat(1) + "╣"
        ElseIf mat(1) = "→" Then
            cmd = "╠" + mat(0) + "╣"
        ElseIf (Mid(mat(0), 2, 1) = "_" And Mid(mat(1), 2, 1) = "^") Or (Mid(mat(0), 2, 1) = "^" And Mid(mat(1), 2, 1) = "_") Then
            cmd = "〖 " + Mid(mat(0), 2, Len(mat(0)) - 2) + Mid(mat(1), 2, Len(mat(1)) - 2) + "〗"
        'ElseIf Mid(mat(0), 2, 1) = "^" And Mid(mat(1), 2, 1) = "_" Then
        '    cmd = "〖" + Mid(mat(0), 2, Len(mat(0)) - 2) + Mid(mat(1), 2, Len(mat(1)) - 2) + "〗"
        Else
            cmdO = False
        End If
    End If
End Function

Function cmdR(ByRef cmd As String, ByRef index As Long) As Boolean
    'Dim index As Long
    Dim str As String
    Dim mFlag As Boolean
    Dim cmdFlag As Integer
    Dim ind As String
    Dim rad As String
    cmdFlag = 0
    
    cmd = ""
    Do While cmdFlag <> 5
        index = index + 1
        str = getCMD(index, mFlag)
        If mFlag = True Then
        '''''''''参数为主命令，先执行
            If exeCMD(str, index) = False Then
                cmdR = False
                Exit Function
            End If
        End If
        
        If cmdFlag = 0 Then
            If str = "(" Then
                cmdFlag = 2
            Else
                cmdFlag = 1
            End If
        ElseIf cmdFlag = 1 Then
        ElseIf cmdFlag = 2 Then
            If str <> "," Then
                ind = ind + str
            ElseIf str = "," Then
                cmdFlag = 4
            End If
        ElseIf cmdFlag = 4 Then
            If str <> ")" Then
                rad = rad + str
            ElseIf str = ")" Then
                cmdFlag = 5
            End If
        End If
    Loop
    cmd = cmd + "〖√(" + ind + "&" + rad + ")〗"
    cmdR = True
End Function

Function cmdS(ByRef cmd As String, ByRef index As Long) As Boolean
    '\F(a\S\UP6(2)+1-3,2a),
    Dim str As String
    Dim mFlag As Boolean
    Dim cmdFlag As Integer
    Dim scr As String
    Dim count As Integer
    
    cmdFlag = 0
    count = 0
    cmd = ""
    Do While cmdFlag <> 5
        index = index + 1
        str = getCMD(index, mFlag)
        If mFlag = True Then
        '''''''''参数为主命令，先执行
            If exeCMD(str, index) = False Then
                cmdS = False
                Exit Function
            End If
        End If
        
        If cmdFlag = 0 Then
            If str = "(" Then
                cmdFlag = 2
            Else
                str = UCase(str)
                If Left(str, 3) = "\UP" Then
                    cmd = "^"
                ElseIf Left(str, 3) = "\DO" Then
                    cmd = "_"
                End If
                'cmdFlag = 1
            End If
        'ElseIf cmdFlag = 1 Then
        '    If str = "(" Then
        '        cmdFlag = 2
        '    Else
        '    End If
        ElseIf cmdFlag = 2 Then
            If str = "(" Then
                count = count + 1
                scr = scr + str
            ElseIf str = ")" Then
                If count > 0 Then
                    scr = scr + str
                    count = count - 1
                Else
                    cmdFlag = 5
                End If
            ElseIf str = "\(" Then
                scr = scr + "("
            ElseIf str = "\)" Then
                scr = scr + ")"
            Else
                scr = scr + str
            End If
        End If
    Loop
    If scr = "→" Then
        cmd = scr
    Else
        cmd = "〖" + cmd + "〖" + scr + "〗〗"
    End If
    cmdS = True
End Function

Function cmdX(ByRef cmd As String, ByRef index As Long) As Boolean
    
End Function

Function changeVec()
    Dim objRange As Range
    Dim objEq As OMath
    Dim objOMathFunction As OMathFunction
     
    Set objRange = Selection.Range
    Set objEq = objRange.OMaths(1)
    Set objOMathFunction = objEq.Functions.Add(objRange, wdOMathFunctionAcc)
    objOMathFunction.Acc.Char = 8407
End Function
Function allChangeToVec()
'
    Dim notend As Boolean
    notend = True
    Selection.HomeKey Unit:=wdStory
    '移动至首行

    Do
        If FindOrReplace("╠*╣", , True) Then
            changeVec
        Else
            notend = False
        End If
        DoEvents
    Loop While notend
    Selection.HomeKey Unit:=wdStory
    FindOrReplace "╠", "", False, , 2
    Selection.HomeKey Unit:=wdStory
    FindOrReplace "╣", "", False, , 2
End Function

Function FindOrReplace(fs As String, Optional rs As String = "", Optional TongPeiFu As Boolean = False, Optional FanWei As Integer = wdFindStop, Optional TiHuanShu As Integer = wdReplaceNone) As Boolean
'全部替换命令，无通配符
'wdFindAsk   2   'wdFindContinue  1  'wdFindStop  0
'wdReplaceNone  0   'wdReplaceOne   1   'wdReplaceAll   2
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    FindOrReplace = Selection.Find.Execute(fs, False, False, TongPeiFu, False, False, True, FanWei, False, rs, TiHuanShu, False, False, False, False)
End Function


Function omathSum()
'
' 宏1 插入 SUM
'
'
    Selection.OMaths.Add Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionNary)
        .Nary.Char = 8721
        .Nary.Grow = False
        .Nary.SubSupLim = False
        .Nary.HideSub = False
        .Nary.HideSup = False
        .Nary.Grow = True
        .Nary.E.Range.Text = "test"
        .Nary.Sub.Range.Text = "do"
        .Nary.Sup.Range.Text = "up"
    End With
End Function

Function omathMatrix()
'
' 宏1 宏
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionMat, 6, 2).mat
        .PlcHoldHidden = False
        .Cell(1, 1).Range.Text = "a"
        .Cell(1, 2).Range.Text = "b"
    End With
End Function

