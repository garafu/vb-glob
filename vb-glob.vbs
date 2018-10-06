'  VB-Glob
'  
'  Copyright (c) 2018 Akinari Tsugo
'  
'  This software is released under the MIT License.
'  http://opensource.org/licenses/mit-license.php

Option Explicit

Public m_strPathDelimiter As String             ' Path control string for Windows.

Public m_fso As Object                          ' FileSystemObject

Public m_rgxIsIncludeControlChar As Object      ' Regular expression of Glob control string.

'
' Constructor.
'
Private Sub Class_Initialize()
    m_strPathDelimiter = "\"
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    Set m_rgxIsIncludeControlChar = RegExp("[\?\*\[]+")
End Sub

'
' Destructor.
'
Private Sub Class_Terminate()
    Set m_rgxIsIncludeControlChar = Nothing
    Set m_fso = Nothing
End Sub

'
' Get current directory acsolute path.
'
Private Function GetCurrentDirectory()
  If ThisWorkbook Is Nothing Then
    Dim objShell: Set objShell = CreateObject("WScript.Shell")
    GetCurrentDirectory = objShell.CurrentDirectory
  Else
    GetCurrentDirectory = ThisWorkbook.Path
  End If
End Function

'
' Create regular expression object.
'
Private Function RegExp(Pattern As String, Optional IsIgnoreCase As Boolean = True, Optional IsGlobal As Boolean = True)
    Dim objRegExp: Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = Pattern
    objRegExp.IgnoreCase = IsIgnoreCase
    objRegExp.Global = IsGlobal
    Set RegExp = objRegExp
End Function

'
' Convert Glob string to Regular expression string.
'
Private Function ConvertGlob2RegExp(Url)
    Dim text: text = Url
    text = Replace(text, "**", "*")
    text = Replace(text, ".", "\.")
    text = Replace(text, "*", ".*")
    ConvertGlob2RegExp = text
End Function

'
' Conbine specified strings as a path string.
'
Private Function PathJoin(a, b)
    If IsEmpty(a) Then
        PathJoin = b
    Else
        PathJoin = Join(Array(a, b), m_strPathDelimiter)
    End If
End Function

'
' Slice specified array.
'
Private Function Slice(List, StartIndex, Length)
    Dim aList(), i, n: n = 0
    ReDim aList(Length - 1)
    
    For i = StartIndex To StartIndex + Length - 1
        aList(n) = List(i)
        n = n + 1
    Next
    
    Slice = aList
End Function


'
' Search subdirectories.
'
Private Function SearchDirs(BaseDir, Match, RestPathList, Result)
    Dim folderCurrent, folderChild, blnNextMatchIsFile, strCrrtMatch, strNextMatch, rgxCrrtMatch, rgxNextMatch
    Dim strBaseDir2, strMatch2, astrRestPathList2
    
    If Not Me.m_fso.FolderExists(BaseDir) Then Exit Function
    
    Set folderCurrent = Me.m_fso.GetFolder(BaseDir)             ' Current folder object.
    strCrrtMatch = Match                                        ' Current matching string.
    strNextMatch = RestPathList(0)                              ' Next matching string
    blnNextMatchIsFile = Not CBool(UBound(RestPathList))        ' Whether the next matching is file or not.
    
    If Match = "**" Then        ' ****** 0 or more directories ******
        If Not blnNextMatchIsFile Then
            For Each folderChild In folderCurrent.SubFolders
                Set rgxNextMatch = RegExp(ConvertGlob2RegExp(strNextMatch))
                strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                If rgxNextMatch.Test(folderChild.name) Then
                    strMatch2 = RestPathList(1)
                    If 1 < UBound(RestPathList) Then
                        astrRestPathList2 = Slice(RestPathList, 2, UBound(RestPathList) - 1)
                        SearchDirs strBaseDir2, strMatch2, astrRestPathList2, Result
                    Else
                        SearchFiles strBaseDir2, strMatch2, Result
                    End If
                Else
                    SearchDirs strBaseDir2, Match, RestPathList, Result
                End If
            Next
        Else
            For Each folderChild In folderCurrent.SubFolders
                strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                SearchDirs strBaseDir2, strCrrtMatch, RestPathList, Result
            Next
            
            SearchFiles BaseDir, strNextMatch, Result
        End If
    ElseIf Match = "*" Then     ' ****** 0 or 1 directory ******
        If Not blnNextMatchIsFile Then
            Set rgxNextMatch = RegExp(ConvertGlob2RegExp(strNextMatch))
            For Each folderChild In folderCurrent.SubFolders
                If rgxNextMatch.Test(folderChild.name) Then
                    strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                    strMatch2 = RestPathList(1)
                    If 1 < UBound(RestPathList) Then
                        astrRestPathList2 = Slice(RestPathList, 2, UBound(RestPathList) - 1)
                        SearchDirs strBaseDir2, strMatch2, astrRestPathList2, Result
                    Else
                        SearchFiles strBaseDir2, strMatch2, Result
                    End If
                End If
            Next
        Else
            For Each folderChild In folderCurrent.SubFolders
                strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                SearchFiles strBaseDir2, strNextMatch, Result
            Next
            
            SearchFiles BaseDir, strNextMatch, Result
        End If
        
    Else                        ' ****** Specified directory ******
        ' Create regular expression.
        Set rgxCrrtMatch = RegExp(ConvertGlob2RegExp(Match))
        
        For Each folderChild In folderCurrent.SubFolders
            If rgxCrrtMatch.Test(folderChild.name) Then
                strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                If Not blnNextMatchIsFile Then
                    SearchDirs strBaseDir2, strNextMatch, Slice(RestPathList, 1, UBound(RestPathList)), Result
                Else
                    SearchFiles strBaseDir2, strNextMatch, Result
                End If
            End If
        Next
    End If
End Function


'
' Search files.
'
Private Function SearchFiles(BaseDir, Match, Result)
    Dim rgxMatch, folder, file
    
    ' Get folder object.
    Set folder = Me.m_fso.GetFolder(BaseDir)
    
    ' Create search condition for sub folder.
    Set rgxMatch = RegExp(ConvertGlob2RegExp(Match))
    
    For Each file In folder.Files
        If rgxMatch.Test(file) Then
            Result.Add file.Path
        End If
    Next
End Function

'
' Get absolute path of files which specified by Glob string.
'
Public Function Search(Path As String)
    Dim i, strBaseDir, astrResult, astrPath, strPath, rgpFile
    
    ' Array for return value.
    Set astrResult = CreateObject("System.Collections.ArrayList")

    ' Convert absolute path.
    If "." = Left(Path, 1) Then
        strPath = PathJoin(GetCurrentDirectory, Path)
    Else
        strPath = Path
    End If
    strPath = Me.m_fso.GetAbsolutePathName(strPath)
    
    ' Split path string by "\".
    astrPath = Split(strPath, m_strPathDelimiter)

    ' Move to the path which is not includes Glob expression.
    For i = LBound(astrPath) To UBound(astrPath) - 1
        If m_rgxIsIncludeControlChar.Test(astrPath(i)) Then
            Exit For
        Else
            strBaseDir = PathJoin(strBaseDir, astrPath(i))
        End If
    Next
    
    ' Whether the next search string is for directory or file.
    If i < UBound(astrPath) Then
        SearchDirs strBaseDir, astrPath(i), Slice(astrPath, i + 1, UBound(astrPath) - i), astrResult
    Else
        SearchFiles strBaseDir, astrPath(UBound(astrPath)), astrResult
    End If
    
    Set Search = astrResult
End Function

