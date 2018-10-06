'  VB-Glob
'  
'  Copyright (c) 2018 Akinari Tsugo
'  
'  This software is released under the MIT License.
'  http://opensource.org/licenses/mit-license.php

Option Explicit

Public m_strPathDelimiter As String             ' Windows向けパス制御文字の定義

Public m_fso As Object                          ' FileSystemObject

Public m_rgxIsIncludeControlChar As Object      ' Glob制御文字を含んでいるかどうかを判定する正規表現

'
' コンストラクタ
'
Private Sub Class_Initialize()
    m_strPathDelimiter = "\"
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    Set m_rgxIsIncludeControlChar = RegExp("[\?\*\[]+")
End Sub

'
' デストラクタ
'
Private Sub Class_Terminate()
    Set m_rgxIsIncludeControlChar = Nothing
    Set m_fso = Nothing
End Sub

'
' カレントディレクトリの絶対パスを取得する
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
' 正規表現を生成する
'
Private Function RegExp(Pattern As String, Optional IsIgnoreCase As Boolean = True, Optional IsGlobal As Boolean = True)
    Dim objRegExp: Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = Pattern
    objRegExp.IgnoreCase = IsIgnoreCase
    objRegExp.Global = IsGlobal
    Set RegExp = objRegExp
End Function

'
' Globを正規表現に変換する
'
Private Function ConvertGlob2RegExp(Url)
    Dim text: text = Url
    text = Replace(text, "**", "*")
    text = Replace(text, ".", "\.")
    text = Replace(text, "*", ".*")
    ConvertGlob2RegExp = text
End Function

'
' 引数をパスとして結合する
'
Private Function PathJoin(a, b)
    If IsEmpty(a) Then
        PathJoin = b
    Else
        PathJoin = Join(Array(a, b), m_strPathDelimiter)
    End If
End Function

'
' 配列の一部を切り取る
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
' フォルダを探索する
'
Private Function SearchDirs(BaseDir, Match, RestPathList, Result)
    Dim folderCurrent, folderChild, blnNextMatchIsFile, strCrrtMatch, strNextMatch, rgxCrrtMatch, rgxNextMatch
    Dim strBaseDir2, strMatch2, astrRestPathList2
    
    If Not Me.m_fso.FolderExists(BaseDir) Then Exit Function
    
    Set folderCurrent = Me.m_fso.GetFolder(BaseDir)             ' フォルダを取得
    strCrrtMatch = Match                                        ' 現在のマッチング文字列
    strNextMatch = RestPathList(0)                              ' 次のマッチング文字列
    blnNextMatchIsFile = Not CBool(UBound(RestPathList))        ' 次のマッチングがファイルかどうか
    
    If Match = "**" Then        ' ▼▼▼ 0個以上のディレクトリ ▼▼▼
        If Not blnNextMatchIsFile Then  ' 次がディレクトリ検索
            For Each folderChild In folderCurrent.SubFolders
                Set rgxNextMatch = RegExp(ConvertGlob2RegExp(strNextMatch))
                strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                If rgxNextMatch.Test(folderChild.name) Then ' 次の検索条件と一致する場合は検索を進める
                    strMatch2 = RestPathList(1)
                    If 1 < UBound(RestPathList) Then
                        astrRestPathList2 = Slice(RestPathList, 2, UBound(RestPathList) - 1)
                        SearchDirs strBaseDir2, strMatch2, astrRestPathList2, Result
                    Else
                        SearchFiles strBaseDir2, strMatch2, Result
                    End If
                Else                                        ' 次の検索条件と一致しない場合は単純に掘り下げる
                    SearchDirs strBaseDir2, Match, RestPathList, Result
                End If
            Next
        Else                            ' 次がファイル検索
            For Each folderChild In folderCurrent.SubFolders
                strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                SearchDirs strBaseDir2, strCrrtMatch, RestPathList, Result
            Next
            
            SearchFiles BaseDir, strNextMatch, Result
        End If
    ElseIf Match = "*" Then     ' ▼▼▼ 0 または 1個のディレクトリ ▼▼▼
        If Not blnNextMatchIsFile Then  ' 次がディレクトリ検索
            Set rgxNextMatch = RegExp(ConvertGlob2RegExp(strNextMatch))
            For Each folderChild In folderCurrent.SubFolders
                If rgxNextMatch.Test(folderChild.name) Then
                    strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                    strMatch2 = RestPathList(1)
                    If 1 < UBound(RestPathList) Then
                        ' 孫がフォルダ
                        astrRestPathList2 = Slice(RestPathList, 2, UBound(RestPathList) - 1)
                        SearchDirs strBaseDir2, strMatch2, astrRestPathList2, Result
                    Else
                        ' 孫がファイル
                        SearchFiles strBaseDir2, strMatch2, Result
                    End If
                End If
            Next
        Else                            ' 次がファイル検索
            For Each folderChild In folderCurrent.SubFolders
                strBaseDir2 = PathJoin(BaseDir, folderChild.name)
                SearchFiles strBaseDir2, strNextMatch, Result
            Next
            
            SearchFiles BaseDir, strNextMatch, Result
        End If
        
    Else                        ' ▼▼▼ 指定されたディレクトリ ▼▼▼
        ' 検索条件を作成
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
' ファイルを探す
'
Private Function SearchFiles(BaseDir, Match, Result)
    Dim rgxMatch, folder, file
    
    ' フォルダを取得
    Set folder = Me.m_fso.GetFolder(BaseDir)
    
    ' サブフォルダの検索条件を作成
    Set rgxMatch = RegExp(ConvertGlob2RegExp(Match))
    
    For Each file In folder.Files
        If rgxMatch.Test(file) Then
            Result.Add file.Path
        End If
    Next
End Function

'
' 指定された Glob パスに合致するファイルの絶対パス一覧を取得します。
'
Public Function Search(Path As String)
    Dim i, strBaseDir, astrResult, astrPath, strPath, rgpFile
    
    ' 戻り値用の配列
    Set astrResult = CreateObject("System.Collections.ArrayList")

    ' 相対パスと絶対パスを判別し、相対パスは絶対パスへ変換
    If "." = Left(Path, 1) Then
        strPath = PathJoin(GetCurrentDirectory, Path)
    End If
    strPath = Me.m_fso.GetAbsolutePathName(strPath)
    
    ' パスを分解
    astrPath = Split(strPath, m_strPathDelimiter)

    ' Glob表現を含まないパスまで移動
    For i = LBound(astrPath) To UBound(astrPath) - 1
        If m_rgxIsIncludeControlChar.Test(astrPath(i)) Then
            Exit For
        Else
            strBaseDir = PathJoin(strBaseDir, astrPath(i))
        End If
    Next
    
    ' 次の検索条件がディレクトリかどうか
    If i < UBound(astrPath) Then
        SearchDirs strBaseDir, astrPath(i), Slice(astrPath, i + 1, UBound(astrPath) - i), astrResult
    Else
        SearchFiles strBaseDir, astrPath(UBound(astrPath)), astrResult
    End If
    
    Set Search = astrResult
End Function

