Attribute VB_Name = "modMain"
Option Explicit
'xrename replace -dir "c:\movie a\" -string /wma$/ig -newstring "rmvb" -type file:/.*\.wma/ -ignorecase yes -log yes -output "c:\list.txt"
'xrename replace -dir "C:\Documents and Settings\sysdzw\����\XRename\inetfilename" -string "[1]" -newstring "" -log yes
'xrename delete -dir "C:\Documents and Settings\sysdzw\����\XRename\inetfilename" -string "[1]"
'ֱ�Ӵ������в�����õ�����
Dim strCmdSub           As String   '��������
Dim strDirectory        As String   '����Ŀ¼
Dim strString           As String   'Ҫ�滻���ַ�(����Ϊ�������ʽȫ��)
Dim strNewString        As String   '�滻����ַ�
Dim strType             As String   'Ҫ�滻�Ķ����޶���Χ�Ĳ�����������������(file|dir|all)�͹������Ƶ��������ʽ
Dim isDealSubDir        As Boolean  '�Ƿ�ݹ���Ŀ¼ Ĭ��ֵ��false
Dim isIgnoreCase        As Boolean  '�Ƿ������ĸ��Сд Ĭ��ֵ��true
Dim isPutLog            As Boolean  '�Ƿ����������log  Ĭ��ֵ��false
Dim strOutputFile       As String   '����ļ��б���·��(������XRename listfile����)

Dim strStringPattern    As String   '��strString���������Ҫ�滻�����ݵ��������ʽ��������//��
Dim strStringPatternP   As String   '��strString���������Ҫ�滻�����ݵ��������ʽ�����ԣ�Ϊ(i|g|ig)��Ĭ��Ϊig����ͨ�ַ���������ת�����������ʽ����������i����isIgnoreCaseӰ��

Dim strGrepTypePre          As String   '��strType����������ǲ������������(file|dir|all)
Dim strTypePattern      As String   '��strType��������������ڸ��ݲ�����������ƽ��й��˵��������ʽ��������//��
Dim strTypePatternP     As String   '��strType��������������ڸ��ݲ�����������ƽ��й��˵��������ʽ�����ԣ�Ϊ(i|g|ig)��һ��Ϊig

Dim strCmd              As String   '�������������в���
Dim reg As Object
Dim matchs As Object, match As Object

Dim regForReplace As Object 'ר�������滻�õ�
Dim regForTestType As Object 'ר���������Է�Χ�Ƿ�ƥ���õ�
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
     
Sub Main()
    Set reg = CreateObject("vbscript.regexp")
    reg.Global = True
    reg.IgnoreCase = True
    
    Set regForReplace = CreateObject("vbscript.regexp")
    Set regForTestType = CreateObject("vbscript.regexp")
    
    strCmd = Trim(Command)
    regForReplace.Pattern = "^""(.+)""$" 'ɾ��������Χ��˫����
    strCmd = regForReplace.Replace(strCmd, "$1")
    strCmd = Trim(strCmd)
    
    If strCmd = "" Then
        MsgBox "��������Ϊ�գ�" & vbCrLf & vbCrLf & _
                "�﷨����:" & vbCrLf & _
                "(1) replace -dir directory -string string1 -new string2 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]" & vbCrLf & _
                "(2) delete -dir directory -string string1 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]" & vbCrLf & _
                "(3) listfile -dir directory -string string1 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-output path]" & vbCrLf & _
                "(4) delfile -dir directory -string string1 [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]" & vbCrLf & _
                "(5) utf8rename -dir directory [-type (file|dir|all)[:string3]] [-ignorecase {yes|no}] [-log {yes|no}]", vbExclamation
        Exit Sub
    End If
    
    Call SetParameter
    Call DoCommand
End Sub
'���ò�������������
Private Sub SetParameter()
    Dim strCmdTmp As String
    strCmdTmp = strCmd & " "
    strCmdSub = regGetStrSub1(strCmdTmp, "^(.+?)\s+?")
    strDirectory = regGetStrSub2(strCmdTmp, "-(?:dir|path)\s+?(""?)(.+?)\1\s+?")
    
    strString = regGetStrSub2(strCmdTmp, "-string\s+?(""?)(.+?)\1\s+?")
    strNewString = regGetStrSub2(strCmdTmp, "-(?:new|newstring|replacewith)\s+?(""?)(.*?)\1\s+?")
    
    strType = regGetStrSub2(strCmdTmp, "-type\s*?(""?)(.+?)\1\s+?")
    isIgnoreCase = IIf(LCase(regGetStrSub2(strCmdTmp, "-ignorecase\s+?(""?)(.+?)\1\s+?")) = "yes", True, False)
    isPutLog = IIf(LCase(regGetStrSub2(strCmdTmp, "-log\s+?(""?)(.+?)\1\s+?")) = "yes", True, False)
    strOutputFile = regGetStrSub2(strCmdTmp, "-output\s+?(""?)(.+?)\1\s+?")
    
    strDirectory = Replace(strDirectory, "/", "\")
    If strDirectory = "" Then strDirectory = "."
    If Right(strDirectory, 1) <> "\" Then strDirectory = strDirectory & "\"
    
    If strOutputFile = "" Then strOutputFile = strDirectory & "XRename_list.txt"
    
    Dim v
    If strString <> "" Then '�û�������-string����
        v = regGetStrSubs(strString, "/(.+?)/(.*)") '������������ʽ��ֵ�����͡������������硰/.*\.wma/ig��
        If v(0) <> "*NULL*" Then '���ƥ��ɹ���ô��ʾ���������ʽ
            strStringPattern = v(0) 'Ҫ�����Ķ���������Ƶ��������ʽ
            strStringPatternP = LCase(v(1)) 'Ҫ�����Ķ���������Ƶ��������ʽ������
        Else 'ƥ��Ϊ��˵������ͨ�ַ���������ִ��ת��Ϊ�������ʽ
            reg.Pattern = "([\[\]\(\)\{\}\.\+\-\/\|\^\$\=\,\?\:])"
            reg.Global = True
            strStringPattern = reg.Replace(strString, "\$1")
            strStringPatternP = "ig" 'g��ʾglobal��ʾȫ��ƥ�䴦����Ĭ����Ҫ����g����Ҫ�ռ���֪ʶ���������Ĭ�ϵ�global��ignorecase��multiline���Զ���false
            If isIgnoreCase Then strStringPatternP = "i" & strStringPatternP '��ʾ�����ʱָ������Ҫ���Դ�Сд����ô��Ҫ����������������i����ʾignorecase
            strNewString = Replace(strNewString, "$", "\$") '�������ͨ�ַ����Ļ�����ô��ʾ������$��Ӧ��ת��
        End If
    End If
    
    If strType <> "" Then '�û�������-type����
        Dim strTypeEx$
        v = regGetStrSubs(strType & " ", "(file|dir|all)(?:\:(""?)(.+?)\2)?\s+?") 'strType�Ӹ��ո���Ϊ�˷��㴦������β\s���֡������������硰file:*.wma��
        If v(0) <> "*NULL*" Then '��ʾ�������������
            strGrepTypePre = LCase(v(0)) 'Ҫ�����Ķ��������(file|dir|all)
            strTypeEx = v(2)
            If strTypeEx <> "" Then '�����������ͨҲ�������������ʽ
                v = regGetStrSubs(strTypeEx, "/(.+?)/(.*)") '������������ʽ��ֵ�����͡������������硰/.*\.wma/ig��
                If v(0) <> "*NULL*" Then
                    strTypePattern = v(0) 'Ҫ�����Ķ���������Ƶ��������ʽ
                    strTypePatternP = LCase(v(1)) 'Ҫ�����Ķ���������Ƶ��������ʽ������
                Else 'ƥ��Ϊ��˵������ͨ�ַ���������ִ��ת��Ϊ�������ʽ����Ҫ��ѭ��������1.����?�滻��. 2.����*�滻��.*?
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
        If strCmdSub = "deldir" Or strCmdSub = "deletedir" Then '�����Ҫɾ��Ŀ¼����ô������������ΪĿ¼�ˡ�
            strGrepTypePre = "dir"
        End If
    End If
End Sub
'��ʼ����
Private Sub DoCommand()
    If Not isNameMatch(strCmdSub, "^(replace|rep|del|delete|listfile|delfile|deletefile|deldir|deletedir|utf8decode)$") Then
        MsgBox "������������Ҳ���""" & strCmdSub & """��ֻ��Ϊ(replace,delete,listfile,delfile,deldir,utf8decode)�е�һ�֡�" & vbCrLf & vbCrLf & "����ϸ��鴫��Ĳ���:" & vbCrLf & strCmd, vbExclamation
        Exit Sub
    End If
    
    If strDirectory = "" Then '����������Ϊ����ô��ʾĬ�ϴ�����ǰ����Ŀ¼����cmd��ֱ����������Ļ����ף�������������bat��ʹ��
        strDirectory = ".\"
    End If
    
    If Dir(strDirectory, vbDirectory) = "" Then
        MsgBox "ָ��Ҫ�������ļ���""" & strDirectory & """�����ڣ�" & vbCrLf & vbCrLf & "����ϸ��鴫��Ĳ���:" & vbCrLf & strCmd, vbExclamation
        End
    End If
    
    If strString = "" And LCase(strCmdSub) <> "utf8decode" And LCase(strCmdSub) <> "deldir" And LCase(strCmdSub) <> "deletedir" Then
        MsgBox "ȱ�ٱ�ѡ����string�����÷���:-string Ҫ�滻���ַ�(����Ϊ�������ʽ)��" & vbCrLf & vbCrLf & "����ϸ��鴫��Ĳ���:" & vbCrLf & strCmd, vbExclamation
        Exit Sub
    End If
    

    Dim strFileNameAll$, vFileName, i&
    Dim strFileName$, strFileNameEx$
    Dim strFileNameNew$, strFileNameNewEx$
    Dim strRenameStatus$
    Dim strDeleteFileStatus$
    Dim isDone As Boolean

    '�õ��ļ����ļ��еļ���
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
         
        strFileName = Dir '�ٴε���dir����,��ʱ���Բ�������
    Loop
    
    If strFileNameAll <> "" Then  '������һ���ļ��ſ�ʼ����
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
                    If strTypePattern = "" Then '���������Χ�Ĳ����ǿ���ô���������ļ�
                        isDone = True
                    Else '����������ʽ������ôȥ�ж��Ƿ�ƥ�������й���
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameEx = strDirectory & vFileName(i) '��ǰ�ļ���ȫ·��
                
                        strFileNameNew = regForReplace.Replace(vFileName(i), strNewString) '���ļ��������滻
                        strFileNameNewEx = strDirectory & strFileNameNew '�����滻�ɵ��ļ���ȫ·��
                        
                        If strFileNameEx <> strFileNameNewEx Then
                            strRenameStatus = DoRename(strFileNameEx, strFileNameNewEx)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strRenameStatus, False
                            If InStr(strRenameStatus, "״̬:ʧ��") > 0 Then writeToFile strDirectory & "err.log", strRenameStatus, False
                        End If
                    End If
                Next
            Case "del", "delete"
                For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '���������Χ�Ĳ����ǿ���ô���������ļ�
                        isDone = True
                    Else '����������ʽ������ôȥ�ж��Ƿ�ƥ�������й���
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameEx = strDirectory & vFileName(i) '��ǰ�ļ���ȫ·��
                        strFileNameNew = regForReplace.Replace(vFileName(i), "") '���ļ��������滻
                        strFileNameNewEx = strDirectory & strFileNameNew '�����滻�ɵ��ļ���ȫ·��
                        
                        If strFileNameEx <> strFileNameNewEx Then
                            strRenameStatus = DoRename(strFileNameEx, strFileNameNewEx)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strRenameStatus, False
                            If InStr(strRenameStatus, "״̬:������ʧ��") > 0 Then writeToFile strDirectory & "err.log", strRenameStatus, False
                        End If
                    End If
                Next
            Case "listfile"
                 For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '���������Χ�Ĳ����ǿ���ô���������ļ�
                        isDone = True
                    Else '����������ʽ������ôȥ�ж��Ƿ�ƥ�������й���
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameEx = strDirectory & vFileName(i) '��ǰ�ļ���ȫ·��
                    
                        If regForReplace.test(vFileName(i)) Then
                            writeToFile strOutputFile, strDeleteFileStatus, False
                        End If
                    End If
                Next
            Case "delfile", "deletefile"
                 For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '���������Χ�Ĳ����ǿ���ô���������ļ�
                        isDone = True
                    Else '����������ʽ������ôȥ�ж��Ƿ�ƥ�������й���
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameEx = strDirectory & vFileName(i) '��ǰ�ļ���ȫ·��
                    
                        If regForReplace.test(vFileName(i)) Then
                            strDeleteFileStatus = DoDelete(strFileNameEx)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strDeleteFileStatus, False
                            If InStr(strRenameStatus, "״̬:ɾ����ʧ��") > 0 Then writeToFile strDirectory & "err.log", strDeleteFileStatus, False
                        End If
                    End If
                Next
            Case "deldir", "deletedir" 'δ������20200924 deleteFonder
                 For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '���������Χ�Ĳ����ǿ���ô���������ļ�
                        isDone = True
                    Else '����������ʽ������ôȥ�ж��Ƿ�ƥ�������й���
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameEx = strDirectory & vFileName(i) '��ǰ�ļ���ȫ·��
                    
                        If regForReplace.test(vFileName(i)) Then
                            strDeleteFileStatus = DoDelete(strFileNameEx)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strDeleteFileStatus, False
                            If InStr(strRenameStatus, "״̬:ɾ����ʧ��") > 0 Then writeToFile strDirectory & "err.log", strDeleteFileStatus, False
                        End If
                    End If
                Next
            Case "utf8decode"
                For i = 0 To UBound(vFileName)
                    If strTypePattern = "" Then '���������Χ�Ĳ����ǿ���ô���������ļ�
                        isDone = True
                    Else '����������ʽ������ôȥ�ж��Ƿ�ƥ�������й���
                        isDone = isNameMatch(vFileName(i), strTypePattern)
                    End If
                    
                    If isDone Then
                        strFileNameEx = strDirectory & vFileName(i) '��ǰ�ļ���ȫ·��
                
                        strFileNameNew = UTF8Decode(vFileName(i)) '���ļ�������UTF8����ת��
                        strFileNameNewEx = strDirectory & strFileNameNew '�����滻�ɵ��ļ���ȫ·��
                        
                        If strFileNameEx <> strFileNameNewEx Then
                            strRenameStatus = DoRename(strFileNameEx, strFileNameNewEx)
                            If isPutLog Then writeToFile strDirectory & "XRename.log", strRenameStatus, False
                            If InStr(strRenameStatus, "״̬:ʧ��") > 0 Then writeToFile strDirectory & "err.log", strRenameStatus, False
                        End If
                    End If
                Next
        End Select
    End If
End Sub
'�������ļ���
Private Function DoRename(ByVal strFileName$, ByVal strFileNew$) As String
    Dim i%
    
    If LCase(strFileName) <> LCase(strFileNew) Then '����Ǵ�Сд��ɵ��ļ��Ѿ������������޸ĵ�
        On Error Resume Next
        i = GetAttr(strFileNew) '�ж��ļ����ļ����Ƿ���ڡ������һ���Ѵ��ڵĶ�����ô��GetAttrȥ�������ʱ���ᱨ��
        If Err.Number = 0 Then
            DoRename = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & " ==> " & strFileNew$ & vbCrLf & "״̬:������ʧ�ܡ�������Ϣ:�Ѿ�������ͬ���Ƶ��ļ������ļ��У�" & vbCrLf
            Exit Function
        End If
    End If
    
    On Error GoTo Err1
    Name strFileName As strFileNew
    DoRename = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & " ==> " & strFileNew$ & vbCrLf & "״̬:�������ɹ���" & vbCrLf
    
    Exit Function
Err1:
    DoRename = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & " ==> " & strFileNew$ & vbCrLf & "״̬:������ʧ�ܡ�������Ϣ:" & Err.Description & " �����:" & Err.Number & vbCrLf
End Function
'ɾ��ָ���ļ������ļ���
Private Function DoDelete(ByVal strFileName$) As String
    Dim i%
    
    On Error Resume Next
    i = GetAttr(strFileName) '�ж��ļ����ļ����Ƿ���ڡ������һ���Ѵ��ڵĶ�����ô��GetAttrȥ�������ʱ���ᱨ��

    On Error GoTo Err1
    If i = 16 Then 'ɾ���ļ�
        Kill strFileName
    Else 'ɾ���ļ���
        deleteFonder strFileName
    End If
    DoDelete = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName & vbCrLf & "״̬:ɾ���ɹ���" & vbCrLf
    
    Exit Function
Err1:
    DoDelete = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strFileName$ & vbCrLf & "״̬:ɾ��ʧ�ܡ�������Ϣ:" & Err.Description & " �����:" & Err.Number & vbCrLf
End Function
'ɾ��ָ���ļ���  20200924��ɾ������
Private Function DoDeleteDir(ByVal strPath$) As String
    Dim i%
    
    On Error Resume Next
    i = GetAttr(strPath) '�ж��ļ����ļ����Ƿ���ڡ������һ���Ѵ��ڵĶ�����ô��GetAttrȥ�������ʱ���ᱨ��

    On Error GoTo Err1
    If i = 16 Then '���ļ��в�ɾ���������ļ�
        deleteFonder strPath
        DoDeleteDir = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strPath & vbCrLf & "״̬:ɾ���ļ��гɹ���" & vbCrLf
    End If
    
    Exit Function
Err1:
    DoDeleteDir = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & strPath & vbCrLf & "״̬:ɾ���ļ���ʧ�ܡ�������Ϣ:" & Err.Description & " �����:" & Err.Number & vbCrLf
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ����������ļ���������ֱ��д�ļ�
'��������writeToFile
'��ڲ���(����)��
'  strFileName �������ļ�����
'  strContent Ҫ���뵽�����ļ����ַ���
'  isCover �Ƿ񸲸Ǹ��ļ���Ĭ��Ϊ����
'����ֵ��True��False���ɹ��򷵻�ǰ�ߣ����򷵻غ���
'��ע��sysdzw �� 2007-5-2 �ṩ
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
'�õ��������ŵĵ�1��ƥ����
Private Function regGetStrSub1(strData$, strPattern$) As String
    reg.Pattern = strPattern
    Set matchs = reg.Execute(strData$)
    If matchs.Count >= 1 Then
        regGetStrSub1 = matchs(0).SubMatches(0)
    End If
End Function
'�õ��������ŵĵ�2��ƥ����
Private Function regGetStrSub2(strData$, strPattern$) As String
    reg.Pattern = strPattern
    Set matchs = reg.Execute(strData$)
    If matchs.Count >= 1 Then
        regGetStrSub2 = matchs(0).SubMatches(1)
    End If
End Function

'�õ�������ƥ����������ݣ���ŵ�һ��������
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

'��Ҫ�����������ļ����ļ������Ƿ�ƥ��
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
'10����תn����(Ĭ��2)
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
'�����ƴ���ת��Ϊʮ�����ƴ���
Public Function c2to16(ByVal x As String) As String
   Dim i As Long
   i = 1
   For i = 1 To Len(x) Step 4
      c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
   Next
End Function
'�����ƴ���ת��Ϊʮ���ƴ���
Public Function c2to10(ByVal x As String) As String
   c2to10 = 0
   If x = "0" Then Exit Function
   Dim i As Long
   i = 0
   For i = 0 To Len(x) - 1
      If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
   Next
End Function
Private Sub deleteFonder(ByVal strPath$)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFolder strPath
    Set FSO = Nothing
End Sub