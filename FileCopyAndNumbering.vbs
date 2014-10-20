Option Explicit

Dim args                ' 引数(ドラッグアンドドロップ用)
Dim fullPathfileName    ' フルパスファイル名
Dim fullPathLength      ' フルパスファイル名の長さ
Dim folderName          ' フォルダ名(最後の\無し)
Dim fileName            ' ファイル名(パス無し、拡張子無し)
Dim extName             ' 拡張子
Dim fileNamePos         ' ファイル名が始まる場所(左から数えて)
Dim extNamePos          ' 拡張子が始まる場所(左から数えて)
Dim inputNum            ' コピー回数
Dim loopCntFileNum      ' ループカウンタ(ファイル数)
Dim fileSystemObject    ' FileSystemObject
Dim fileNameCopyFrom    ' コピー元ファイル名(フルパス)
Dim fileNameCopyTo      ' コピー先ファイル名(フルパス)
Dim message             ' 表示用メッセージ
Dim title               ' タイトル
Dim scriptPath          ' 本スクリプトのパス
Dim scriptPath_tmp      ' 本スクリプトのパス取得用一時変数

title = "FileCopyAndNumbering"

Set args  = WScript.Arguments
Set fileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count <> 1 Then ' ドラッグ＆ドロップせずに起動した場合
    WScript.Echo "コピーしたいファイルをドラッグ＆ドロップしてください。"
    WScript.Quit
Else
    ' Do Nothing
End If

fullPathfileName = args(0)
fullPathLength   = Len(fullPathfileName)

fileNamePos = InStrRev(fullPathfileName, "\")
extNamePos  = InStrRev(fullPathfileName, ".")

' 拡張子切り出し
extName = Right(fullPathfileName, fullPathLength - extNamePos)

' フォルダ名切り出し
folderName = Left(fullPathfileName, fileNamePos - 1)

' ファイル名切り出し
fileName = Mid(fullPathfileName, fileNamePos + 1, fullPathLength - (Len(folderName) + Len(extName) + 2))

message = fileName & "." & extName & "をいくつコピーしますか？"
inputNum = InputBox(message, title)

If Len(inputNum) = 0 Then ' キャンセルが押された場合
    ' Do Nothing
ElseIf IsNumeric(inputNum) Then ' OKが押されて、数値が入力された場合
    fileNameCopyFrom = fullPathfileName

    scriptPath_tmp = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    scriptPath = Left(scriptPath_tmp, Len(scriptPath_tmp) - 1 )

    For loopCntFileNum = 1 TO inputNum
        fileNameCopyTo  = scriptPath & "\" & fileName & "_" & loopCntFileNum & "." & extName
        Call fileSystemObject.CopyFile(fileNameCopyFrom,fileNameCopyTo,False) ' コピー実施
    Next

    WScript.Echo inputNum & "個コピーしました"
Else ' OKが押されて、数値以外が入力された場合
    WScript.Echo "数値を入力して下さい"
End If

Set fileSystemObject = Nothing
