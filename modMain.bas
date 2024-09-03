Attribute VB_Name = "modMain"
Option Explicit
'xrename replace -dir "c:\movie a\" -string /wma$/ig -newstring "rmvb" -type file:/.*\.wma/ -ignorecase yes -log yes -output "c:\list.txt"
'xrename replace -dir "C:\Documents and Settings\sysdzw\桌面\XRename\inetfilename" -string "[1]" -newstring "" -log yes
'xrename delete -dir "C:\Documents and Settings\sysdzw\桌面\XRename\inetfilename" -string "[1]"
'-ignoreExt 忽略处理后缀名
'直接从命令行参数获得的数据
Dim strCmdSub           As String   '二级命令
Dim strDirectory        As String   '工作目录
Dim strString           As String   '要替换的字符(可能为正则表达式全体)
Dim strNewString        As String   '替换后的字符
Dim strType             As String   '要替换的对象限定范围的参数，包含对象类型(file|dir|all)和过滤名称的正则表达式
Dim isDealSubDir        As Boolean  '是否递归子目录 默认值：false
Dim isIgnoreCase        As Boolean  '是否忽略字母大小写 默认值：true
Dim isIgnoreExt        As Boolean  '是否忽略处理后缀名 默认值：true
Dim isPutLog            As Boolean  '是否输出处理的log  默认值：false
Dim strOutputFile       As String   '输出文件列表的路径(仅用于XRename listfile命令)

Dim strStringPattern    As String   '从strString分离出来，要替换的内容的正则表达式，不包含//等
Dim strStringPatternP   As String   '从strString分离出来，要替换的内容的正则表达式的属性，为(i|g|ig)，默认为ig，普通字符串处理会转换成正则表达式处理，所以i会受isIgnoreCase影响

Dim strGrepTypePre          As String   '从strType分离出来，是操作对象的类型(file|dir|all)
Dim strTypePattern      As String   '从strType分离出来，是用于根据操作对象的名称进行过滤的正则表达式，不包含//等
Dim strTypePatternP     As String   '从strType分离出来，是用于根据操作对象的名称进行过滤的正则表达式的属性，为(i|g|ig)，一般为ig

Dim strCmd              As String   '程序完整命令行参数
Dim reg As Object
Dim matchs As Object, match As Object

Dim regForReplace As Object '专门用来替换用的
Dim regForTestType As Object '专门用来测试范围是否匹配用的
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
     
Sub Main()
    Set reg = CreateObject("vbscript.regexp")
    reg.Global = True
    reg.IgnoreCase = True
    
    Set regForReplace = CreateObject("vbscript.regexp")
    Set regForTestType = CreateObject("vbscript.regexp")
    
    strCmd = Trim(Command)
    regForReplace.Pattern = "^""(.+)""$" '删除掉最外围的双引号
    strCmd = regForReplace.Replace(strCmd, "$1")
    strCmd = Trim(strCmd)
    
    If strCmd = "" Then
        MsgBox "参数不能为空！" & vbCrLf & vbCrLf & _
                "语法如下:" & vbCrLf & _
                "(1) replace -dir directory -string string1 -new string2 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-ignoreExt {yes|no}] [-log {yes|no}]" & vbCrLf & _
                "(2) delete -dir directory -string string1 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]" & vbCrLf & _
                "(3) listfile -dir directory -string string1 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-output path]" & vbCrLf & _
                "(4) delfile -dir directory -string string1 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]" & vbCrLf & _
                "(5) utf8decode -dir directory [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]" & vbCrLf & _
                "(6) cn2number -dir directory [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]", vbExclamation
        Exit Sub
    End If
    
    Call SetParameter
    Call DoCommand
End Sub
'设置参数到各个变量
Private Sub SetParameter()
    Dim strCmdTmp As String
    strCmdTmp = strCmd & " "
    strCmdSub = regGetStrSub1(strCmdTmp, "^(.+?)\s+?")
    strDirectory = regGetStrSub2(strCmdTmp, "-(?:dir|path)\s+?(""?)(.+?)\1\s+?")
    
    strString = regGetStrSub1(strCmdTmp, "-string\s+?(/.*?/[^\s]*)") '先尝试//正则方式获取
    If strString = "" Then strString = regGetStrSub2(strCmdTmp, "-string\s+?(""?)(.+?)\1\s+?")
    
    strNewString = regGetStrSub2(strCmdTmp, "-(?:new|newstring|replacewith)\s+?(""?)(.*?)\1\s+?")
    
    strType = regGetStrSub2(strCmdTmp, "-type\s+?(""?)(.+?)\1\s+?")
    
    isIgnoreCase = IIf(LCase(regGetStrSub2(strCmdTmp, "-ignorecase\s+?(""?)(.+?)\1\s+?")) = "yes", True, False)
    isIgnoreExt = IIf(LCase(regGetStrSub2(strCmdTmp, "-ignoreext\s+?(""?)(.+?)\1\s+?")) = "yes", True, False)
    isPutLog = IIf(LCase(regGetStrSub2(strCmdTmp, "-log\s+?(""?)(.+?)\1\s+?")) = "yes", True, False)
    strOutputFile = regGetStrSub2(strCmdTmp, "-output\s+?(""?)(.+?)\1\s+?")
    
    strDirectory = Replace(strDirectory, "/", "\")
    If strDirectory = "" Then strDirectory = "."
    If Right(strDirectory, 1) <> "\" Then strDirectory = strDirectory & "\"
    
    If strOutputFile = "" Then strOutputFile = strDirectory & "XRename_list.txt"
    
    Dim v
    If strString <> "" Then '用户设置了-string参数
        If Left(strString, 1) = "/" Then '表示正则模式
            v = regGetStrSubs(strString, "/(.+?)/(.*)") '分离出正则表达式的值和类型。处理数据例如“/.*\.wma/ig
            strStringPattern = v(0) '要处理的对象过滤名称的正则表达式
            strStringPatternP = LCase(v(1)) '要处理的对象过滤名称的正则表达式的类型
        End If
    End If
    
    If strType <> "" Then '用户设置了-type参数
        Dim strTypeEx$
        v = regGetStrSubs(strType & " ", "(file|dir|all)(?:\:(""?)(.+?)\2)?\s+?") 'strType加个空格是为了方便处理，结尾\s区分。处理数据例如“file:*.wma”
        If v(0) <> "*NULL*" Then '表示这个参数有数据
            strGrepTypePre = LCase(v(0)) '要处理的对象的类型(file|dir|all)
            strTypeEx = v(2)
            If strTypeEx <> "" Then '这里可能是普通也可能是正则表达式
                v = regGetStrSubs(strTypeEx, "/(.+?)/(.*)") '分离出正则表达式的值和类型。处理数据例如“/.*\.wma/ig”
                If v(0) <> "*NULL*" Then
                    strTypePattern = v(0) '要处理的对象过滤名称的正则表达式
                    strTypePatternP = LCase(v(1)) '要处理的对象过滤名称的正则表达式的类型
                Else '匹配为空说明是普通字符串，下面执行转换为正则表达式，需要遵循两个规则：1.遇到?替换成. 2.遇到*替换成.*?  但是如果有*或者问号需要用正则处理。 *.txt -> .*\.txt 再例如： a?b 变成a.b
                    reg.Pattern = "(\[\]\(\)\{\}\.\+\-\/\|\^\$\=\,)"
                    reg.Global = True
                    strTypePattern = reg.Replace(strTypeEx, "\$1")
                    strTypePattern = Replace(strTypePattern, "?", ".")
                    If Left(strTypePattern, 1) <> "*" And InStr(strTypePattern, "*") > 0 Then strTypePattern = "^" & strTypePattern
                    If Right(strTypePattern, 1) <> "*" And InStr(strTypePattern, "*") Then strTypePattern = strTypePattern & "$"
                    strTypePattern = Replace(strTypePattern, "*", ".*?")

                    strTypePatternP = "ig"
                End If
            End If
        Else
            strGrepTypePre = "file"
        End If
    Else
        strGrepTypePre = "file"
        If strCmdSub = "deldir" Or strCmdSub = "deletedir" Then '如果是要删除目录的那么就是设置属性为目录了。
            strGrepTypePre = "dir"
        End If
    End If
End Sub
'开始处理
Private Sub DoCommand()
    If Not isNameMatch(strCmdSub, "^(replace|rep|del|delete|listfile|delfile|deletefile|deldir|deletedir|utf8decode|cn2number)$") Then
        MsgBox "二级命令错误，找不到""" & strCmdSub & """，只能为(replace,delete,listfile,delfile,deldir,utf8decode,cn2number)中的一种。" & vbCrLf & vbCrLf & "请仔细检查传入的参数:" & vbCrLf & strCmd, vbExclamation
        Exit Sub
    End If
    
    If strDirectory = "" Then '如果这个参数为空那么表示默认处理当前所在目录，在cmd中直接敲入命令的话不妥，建议在批处理bat中使用
        strDirectory = ".\"
    End If
    
    If Dir(strDirectory, vbDirectory) = "" Then
        MsgBox "指定要处理的文件夹""" & strDirectory & """不存在！" & vbCrLf & vbCrLf & "请仔细检查传入的参数:" & vbCrLf & strCmd, vbExclamation
        End
    End If
    
    If strString = "" And LCase(strCmdSub) <> "utf8decode" And LCase(strCmdSub) <> "cn2number" And LCase(strCmdSub) <> "deldir" And LCase(strCmdSub) <> "deletedir" Then
        MsgBox "缺少必选参数string。设置方法:-string 要替换的字符(可以为正则表达式)。" & vbCrLf & vbCrLf & "请仔细检查传入的参数:" & vbCrLf & strCmd, vbExclamation
        Exit Sub
    End If
    

    Dim strFileNameAll$, vFileName, i&
    Dim strFileName$, strFileNameFull$, strFileNamePre$, strFileNameExt$, v
    Dim strFileNameNew$, strFileNameNewFull$
    Dim strRenameStatus$
    Dim strDeleteFileStatus$
    Dim isDone As Boolean

    '得到文件或文件夹的集合
    strFileName = Dir(strDirectory, vbDirectory)
    Do While strFileName <> ""
        If strFileName <> "." And strFileName <> ".." Then
            If strGrepTypePre = "dir" Then
                If (GetAttr(strDirectory & strFileName) And vbDirectory) = vbDirectory Then strFileNameAll = strFileNameAll & strFileName & vbCrLf
            ElseIf strGrepTypePre = "file" Then
                If (GetAttr(strDirectory & strFileName) And vbDirectory) <> vbDirectory Then strFileNameAll = strFileNameAll & strFileName & vbCrLf
            ElseIf strGrepTypePre = "all" Then
                strFileNameAll = strFileNameAll & strFileName & vbCrLf
            End If
         End If
         
        strFileName = Dir '再次调用dir函数,此时可以不带参数
    Loop
    
    If strFileNameAll <> "" Then  '至少有一个文件才开始处理
        strFileNameAll = Left(strFileNameAll, Len(strFileNameAll) - 2)
        vFileName = Split(strFileNameAll, vbCrLf)
        
        regForReplace.Pattern = strStringPattern
        regForReplace.IgnoreCase = (InStr(strStringPatternP, "i") > 0)
        regForReplace.MultiLine = (InStr(strStringPatternP, "m") > 0)
        regForReplace.Global = (InStr(strStringPatternP, "g") > 0)
        
        regForTestType.Pattern = strTypePattern
        regForTestType.IgnoreCase = (InStr(strTypePatternP, "i") > 0)
        regForTestType.MultiLine = (InStr(strTypePatternP, "m") > 0)
        regForTestType.Global = (InStr(strTypePatternP, "g") > 0)
        
        Select Case LCase(strCmdSub)
            Case "rep", "replace" 'XRename replace -dir "c:\movie a\" -string "wma$" -replacewith "rmvb" -type file:".*\.wma" -ignorecase yes -log yes
                For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '如果处理范围的参数是空那么处理所有文件
                        isDone = True
                    Else '如果正则表达式存在那么去判断是否匹配来进行过滤
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameFull = strDirectory & vFileName(i) '当前文件的全路径

                        If isIgnoreExt And InStr(vFileName(i), ".") > 0 Then   '忽略后缀名。也就是不处理后缀名，当然如果没有后缀名的话直接走下面的分支替换
                            v = Split(vFileName(i), ".")
                            strFileNamePre = Left(vFileName(i), InStrRev(vFileName(i), ".") - 1) '后缀之前的内容
                            strFileNameExt = v(UBound(v)) '后缀
                            
                            If Left(strString, 1) = "/" Then '表示正则模式
                                strFileNameNew = regForReplace.Replace(strFileNamePre, strNewString) & "." & strFileNameExt '用正则替换
                            Else
                                strFileNameNew = Replace(strFileNamePre, strString, strNewString) & "." & strFileNameExt
                            End If
                        Else
                            If Left(strString, 1) = "/" Then '表示正则模式
                                strFileNameNew = regForReplace.Replace(vFileName(i), strNewString) '用正则替换
                            Else
                                strFileNameNew = Replace(vFileName(i), strString, strNewString) '正常替换
                            End If
                        End If
                        
                        strFileNameNewFull = strDirectory & strFileNameNew '即将替换成的文件的全路径
                        
                        If strFileNameFull <> strFileNameNewFull Then
                            strRenameStatus = DoRename(strFileNameFull, strFileNameNewFull)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strRenameStatus, False
                            If InStr(strRenameStatus, "状态:失败") > 0 Then writeToFile strDirectory & "err.log", strRenameStatus, False
                        End If
                    End If
                Next
            Case "del", "delete"
                For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '如果处理范围的参数是空那么处理所有文件
                        isDone = True
                    Else '如果正则表达式存在那么去判断是否匹配来进行过滤
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameFull = strDirectory & vFileName(i) '当前文件的全路径
                        
                        If isIgnoreExt And InStr(vFileName(i), ".") > 0 Then   '忽略后缀名。也就是不处理后缀名，当然如果没有后缀名的话直接走下面的分支替换
                            v = Split(vFileName(i), ".")
                            strFileNamePre = Left(vFileName(i), InStrRev(vFileName(i), ".") - 1) '后缀之前的内容
                            strFileNameExt = v(UBound(v)) '后缀
                            
                            If Left(strString, 1) = "/" Then '表示正则模式
                                strFileNameNew = regForReplace.Replace(strFileNamePre, "") & "." & strFileNameExt '用正则替换
                            Else
                                strFileNameNew = Replace(strFileNamePre, strString, "") & "." & strFileNameExt
                            End If
                        Else
                            If Left(strString, 1) = "/" Then '表示正则模式
                                strFileNameNew = regForReplace.Replace(vFileName(i), "") '用正则替换
                            Else
                                strFileNameNew = Replace(vFileName(i), strString, "") '正常替换
                            End If
                        End If
                        
                        strFileNameNewFull = strDirectory & strFileNameNew '即将替换成的文件的全路径
                        
                        If strFileNameFull <> strFileNameNewFull Then
                            strRenameStatus = DoRename(strFileNameFull, strFileNameNewFull)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strRenameStatus, False
                            If InStr(strRenameStatus, "状态:重命名失败") > 0 Then writeToFile strDirectory & "err.log", strRenameStatus, False
                        End If
                    End If
                Next
            Case "listfile"
                 For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '如果处理范围的参数是空那么处理所有文件
                        isDone = True
                    Else '如果正则表达式存在那么去判断是否匹配来进行过滤
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameFull = strDirectory & vFileName(i) '当前文件的全路径
                    
                        If regForReplace.test(vFileName(i)) Then
                            writeToFile strOutputFile, strDeleteFileStatus, False
                        End If
                    End If
                Next
            Case "delfile", "deletefile"
                 For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '如果处理范围的参数是空那么处理所有文件
                        isDone = True
                    Else '如果正则表达式存在那么去判断是否匹配来进行过滤
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameFull = strDirectory & vFileName(i) '当前文件的全路径
                    
                        If regForReplace.test(vFileName(i)) Then
                            strDeleteFileStatus = DoDelete(strFileNameFull)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strDeleteFileStatus, False
                            If InStr(strRenameStatus, "状态:删除名失败") > 0 Then writeToFile strDirectory & "err.log", strDeleteFileStatus, False
                        End If
                    End If
                Next
            Case "deldir", "deletedir" '未处理好20200924 deleteFolder
                 For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '如果处理范围的参数是空那么处理所有文件
                        isDone = True
                    Else '如果正则表达式存在那么去判断是否匹配来进行过滤
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameFull = strDirectory & vFileName(i) '当前文件的全路径
                    
                        If regForReplace.test(vFileName(i)) Then
                            strDeleteFileStatus = DoDelete(strFileNameFull)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strDeleteFileStatus, False
                            If InStr(strRenameStatus, "状态:删除名失败") > 0 Then writeToFile strDirectory & "err.log", strDeleteFileStatus, False
                        End If
                    End If
                Next
            Case "utf8decode"
                For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '如果处理范围的参数是空那么处理所有文件
                        isDone = True
                    Else '如果正则表达式存在那么去判断是否匹配来进行过滤
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameFull = strDirectory & vFileName(i) '当前文件的全路径
                
                        strFileNameNew = UTF8Decode(vFileName(i)) '短文件名进行UTF8编码转换
                        strFileNameNewFull = strDirectory & strFileNameNew '即将替换成的文件的全路径
                        
                        If strFileNameFull <> strFileNameNewFull Then
                            strRenameStatus = DoRename(strFileNameFull, strFileNameNewFull)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strRenameStatus, False
                            If InStr(strRenameStatus, "状态:失败") > 0 Then writeToFile strDirectory & "err.log", strRenameStatus, False
                        End If
                    End If
                Next
            Case "cn2number"
                For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '如果处理范围的参数是空那么处理所有文件
                        isDone = True
                    Else '如果正则表达式存在那么去判断是否匹配来进行过滤
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameFull = strDirectory & vFileName(i) '当前文件的全路径
                        strFileNameNew = vFileName(i)
                        
                        '替换中文数字
                        Dim vCnNumber, intCnNumberIndex As Integer, strNumberChanged As String
                        vCnNumber = regGetStrSubs(strFileNameNew, "([零一二三四五六七八九十百千万亿]+)")
                        If vCnNumber(0) <> "*NULL*" Then
                            For intCnNumberIndex = 0 To UBound(vCnNumber)
                                strNumberChanged = ChineseNumberToArabic(vCnNumber(intCnNumberIndex))
                                strFileNameNew = Replace(strFileNameNew, vCnNumber(intCnNumberIndex), strNumberChanged) '每个数字片段替换
                            Next
                        End If

                        strFileNameNewFull = strDirectory & strFileNameNew '即将替换成的文件的全路径
                        
                        If strFileNameFull <> strFileNameNewFull Then
                            strRenameStatus = DoRename(strFileNameFull, strFileNameNewFull)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strRenameStatus, False
                            If InStr(strRenameStatus, "状态:失败") > 0 Then writeToFile strDirectory & "err.log", strRenameStatus, False
                        End If
                    End If
                Next
        End Select
    End If
End Sub
'重命名文件名
Private Function DoRename(ByVal strFileName$, ByVal strFileNew$) As String
    Dim i%
    
    If LCase(strFileName) <> LCase(strFileNew) Then '如果是大小写造成的文件已经存在是允许修改的
        On Error Resume Next
        i = GetAttr(strFileNew) '判断文件或文件夹是否存在。如果是一个已存在的对象那么用GetAttr去获得属性时不会报错
        If Err.Number = 0 Then
            DoRename = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & " ==> " & strFileNew$ & vbCrLf & "状态:重命名失败。错误信息:已经存在相同名称的文件或者文件夹！" & vbCrLf
            Exit Function
        End If
    End If
    
    On Error GoTo Err1
    Name strFileName As strFileNew
    DoRename = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & " ==> " & strFileNew$ & vbCrLf & "状态:重命名成功。" & vbCrLf
    
    Exit Function
Err1:
    DoRename = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & " ==> " & strFileNew$ & vbCrLf & "状态:重命名失败。错误信息:" & Err.Description & " 错误号:" & Err.Number & vbCrLf
End Function
'删除指定文件或者文件夹
Private Function DoDelete(ByVal strFileName$) As String
    Dim i%
    
    On Error Resume Next
    i = GetAttr(strFileName) '判断文件或文件夹是否存在。如果是一个已存在的对象那么用GetAttr去获得属性时不会报错

    On Error GoTo Err1
    If i = 16 Then '删除文件
        Kill strFileName
    Else '删除文件夹
        deleteFolder strFileName
    End If
    DoDelete = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & vbCrLf & "状态:删除成功。" & vbCrLf
    
    Exit Function
Err1:
    DoDelete = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName$ & vbCrLf & "状态:删除失败。错误信息:" & Err.Description & " 错误号:" & Err.Number & vbCrLf
End Function
'删除指定文件夹  20200924做删除区分
Private Function DoDeleteDir(ByVal strPath$) As String
    Dim i%
    
    On Error Resume Next
    i = GetAttr(strPath) '判断文件或文件夹是否存在。如果是一个已存在的对象那么用GetAttr去获得属性时不会报错

    On Error GoTo Err1
    If i = 16 Then '是文件夹才删除，跳过文件
        deleteFolder strPath
        DoDeleteDir = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strPath & vbCrLf & "状态:删除文件夹成功。" & vbCrLf
    End If
    
    Exit Function
Err1:
    DoDeleteDir = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strPath & vbCrLf & "状态:删除文件夹失败。错误信息:" & Err.Description & " 错误号:" & Err.Number & vbCrLf
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：根据所给文件名和内容直接写文件
'函数名：writeToFile
'入口参数(如下)：
'  strFileName 所给的文件名；
'  strContent 要输入到上述文件的字符串
'  isCover 是否覆盖该文件，默认为覆盖
'返回值：True或False，成功则返回前者，否则返回后者
'备注：sysdzw 于 2007-5-2 提供
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function writeToFile(ByVal strFileName$, ByVal strContent$, Optional isCover As Boolean = True) As Boolean
    On Error GoTo Err1
    Dim fileHandl%
    fileHandl = FreeFile
    If isCover Then
        Open strFileName For Output As #fileHandl
    Else
        Open strFileName For Append As #fileHandl
    End If
    Print #fileHandl, strContent
    Close #fileHandl
    writeToFile = True
    Exit Function
Err1:
    writeToFile = False
End Function
'得到正则括号的第1个匹配项
Private Function regGetStrSub1(strData$, strPattern$) As String
    reg.Pattern = strPattern
    Set matchs = reg.Execute(strData$)
    If matchs.Count >= 1 Then
        regGetStrSub1 = matchs(0).SubMatches(0)
    End If
End Function
'得到正则括号的第2个匹配项
Private Function regGetStrSub2(strData$, strPattern$) As String
    reg.Pattern = strPattern
    Set matchs = reg.Execute(strData$)
    If matchs.Count >= 1 Then
        regGetStrSub2 = matchs(0).SubMatches(1)
    End If
End Function
'得到正则字匹配的所用内容，存放到一个数组中
Private Function regGetStrSubs(strData$, strPattern$)
    Dim s$, v, i%
    reg.Pattern = strPattern
    Set matchs = reg.Execute(strData$)
    If matchs.Count >= 1 Then
        For i = 0 To matchs(0).SubMatches.Count - 1
            s = s & matchs(0).SubMatches(i) & vbCrLf
        Next
    End If
    If s <> "" Then
        s = Left(s, Len(s) - 2)
    Else
        s = "*NULL*"
    End If
    
    regGetStrSubs = Split(s, vbCrLf)
End Function

'主要是用来测试文件或文件夹名是否匹配
Private Function isNameMatch(ByVal strData$, ByVal strPattern$) As Boolean
    regForTestType.Pattern = strPattern
    isNameMatch = regForTestType.test(strData$)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         UTF8 decode model                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UTF8Decode(ByVal code As String) As String
    If code = "" Then
        UTF8Decode = ""
        Exit Function
    End If
    
    Dim tmp As String
    Dim decodeStr As String
    Dim codelen As Long
    Dim result As String
    Dim leftStr As String
     
    leftStr = Left(code, 1)
     
    While (code <> "")
        codelen = Len(code)
        leftStr = Left(code, 1)
        If leftStr = "%" Then
                If (Mid(code, 2, 1) = "C" Or Mid(code, 2, 1) = "B") Then
                    decodeStr = Replace(Mid(code, 1, 6), "%", "")
                    tmp = c10ton(Val("&H" & Hex(Val("&H" & decodeStr) And &H1F3F)))
                    tmp = String(16 - Len(tmp), "0") & tmp
                    UTF8Decode = UTF8Decode & UTF8Decode & ChrW(Val("&H" & c2to16(Mid(tmp, 3, 4)) & c2to16(Mid(tmp, 7, 2) & Mid(tmp, 11, 2)) & Right(decodeStr, 1)))
                    code = Right(code, codelen - 6)
                ElseIf (Mid(code, 2, 1) = "E") Then
                    decodeStr = Replace(Mid(code, 1, 9), "%", "")
                    tmp = c10ton((Val("&H" & Mid(Hex(Val("&H" & decodeStr) And &HF3F3F), 2, 3))))
                    tmp = String(10 - Len(tmp), "0") & tmp
                    UTF8Decode = UTF8Decode & ChrW(Val("&H" & (Mid(decodeStr, 2, 1) & c2to16(Mid(tmp, 1, 4)) & c2to16(Mid(tmp, 5, 2) & Right(tmp, 2)) & Right(decodeStr, 1))))
                    code = Right(code, codelen - 9)
                End If
        Else
            UTF8Decode = UTF8Decode & leftStr
            code = Right(code, codelen - 1)
        End If
    Wend
End Function
'10进制转n进制(默认2)
Public Function c10ton(ByVal x As Integer, Optional ByVal n As Integer = 2) As String
    Dim i As Integer
    i = x \ n
    If i > 0 Then
        If x Mod n > 10 Then
            c10ton = c10ton(i, n) + Chr(x Mod n + 55)
        Else
            c10ton = c10ton(i, n) + CStr(x Mod n)
        End If
    Else
        If x > 10 Then
            c10ton = Chr(x + 55)
        Else
            c10ton = CStr(x)
        End If
    End If
End Function
'二进制代码转换为十六进制代码
Public Function c2to16(ByVal x As String) As String
   Dim i As Long
   i = 1
   For i = 1 To Len(x) Step 4
      c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
   Next
End Function
'二进制代码转换为十进制代码
Public Function c2to10(ByVal x As String) As String
   c2to10 = 0
   If x = "0" Then Exit Function
   Dim i As Long
   i = 0
   For i = 0 To Len(x) - 1
      If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
   Next
End Function
Private Sub deleteFolder(ByVal strPath$)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.deleteFolder strPath
    Set FSO = Nothing
End Sub
'将中文数字转换为阿拉伯数字
Function ChineseNumberToArabic(ByVal cnNumber As String) As Long
    Dim cnDigits As String
    Dim cnValues()
    Dim i As Integer, unit As Long
    Dim result As Long, current As Long, num As Long
    
    ' 中文数字字符
    cnDigits = "零一二三四五六七八九十百千万亿"
    
    ' 对应的阿拉伯数字值
    cnValues = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 100, 1000, 10000, 100000000)
    
    result = 0
    current = 0
    num = 0
    
    For i = 1 To Len(cnNumber)
        unit = InStr(cnDigits, Mid(cnNumber, i, 1))
        
        If unit > 0 Then
            unit = cnValues(unit - 1)
            
            If unit = 10 Or unit = 100 Or unit = 1000 Or unit = 10000 Or unit = 100000000 Then
                If current = 0 Then
                    current = 1
                End If
                current = current * unit
                result = result + current
                current = 0
            Else
                current = current + unit
            End If
        End If
    Next i
    
    If current > 0 Then
        result = result + current
    End If
    
    ChineseNumberToArabic = result
End Function

