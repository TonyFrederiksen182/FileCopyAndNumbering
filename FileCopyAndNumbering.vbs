Option Explicit

Dim args                ' ����(�h���b�O�A���h�h���b�v�p)
Dim fullPathfileName    ' �t���p�X�t�@�C����
Dim fullPathLength      ' �t���p�X�t�@�C�����̒���
Dim folderName          ' �t�H���_��(�Ō��\����)
Dim fileName            ' �t�@�C����(�p�X�����A�g���q����)
Dim extName             ' �g���q
Dim fileNamePos         ' �t�@�C�������n�܂�ꏊ(�����琔����)
Dim extNamePos          ' �g���q���n�܂�ꏊ(�����琔����)
Dim inputNum            ' �R�s�[��
Dim loopCntFileNum      ' ���[�v�J�E���^(�t�@�C����)
Dim fileSystemObject    ' FileSystemObject
Dim fileNameCopyFrom    ' �R�s�[���t�@�C����(�t���p�X)
Dim fileNameCopyTo      ' �R�s�[��t�@�C����(�t���p�X)
Dim message             ' �\���p���b�Z�[�W
Dim title               ' �^�C�g��
Dim scriptPath          ' �{�X�N���v�g�̃p�X
Dim scriptPath_tmp      ' �{�X�N���v�g�̃p�X�擾�p�ꎞ�ϐ�

title = "FileCopyAndNumbering"

Set args  = WScript.Arguments
Set fileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count <> 1 Then ' �h���b�O���h���b�v�����ɋN�������ꍇ
    WScript.Echo "�R�s�[�������t�@�C�����h���b�O���h���b�v���Ă��������B"
    WScript.Quit
Else
    ' Do Nothing
End If

fullPathfileName = args(0)
fullPathLength   = Len(fullPathfileName)

fileNamePos = InStrRev(fullPathfileName, "\")
extNamePos  = InStrRev(fullPathfileName, ".")

' �g���q�؂�o��
extName = Right(fullPathfileName, fullPathLength - extNamePos)

' �t�H���_���؂�o��
folderName = Left(fullPathfileName, fileNamePos - 1)

' �t�@�C�����؂�o��
fileName = Mid(fullPathfileName, fileNamePos + 1, fullPathLength - (Len(folderName) + Len(extName) + 2))

message = fileName & "." & extName & "�������R�s�[���܂����H"
inputNum = InputBox(message, title)

If Len(inputNum) = 0 Then ' �L�����Z���������ꂽ�ꍇ
    ' Do Nothing
ElseIf IsNumeric(inputNum) Then ' OK��������āA���l�����͂��ꂽ�ꍇ
    fileNameCopyFrom = fullPathfileName

    scriptPath_tmp = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    scriptPath = Left(scriptPath_tmp, Len(scriptPath_tmp) - 1 )

    For loopCntFileNum = 1 TO inputNum
        fileNameCopyTo  = scriptPath & "\" & fileName & "_" & loopCntFileNum & "." & extName
        Call fileSystemObject.CopyFile(fileNameCopyFrom,fileNameCopyTo,False) ' �R�s�[���{
    Next

    WScript.Echo inputNum & "�R�s�[���܂���"
Else ' OK��������āA���l�ȊO�����͂��ꂽ�ꍇ
    WScript.Echo "���l����͂��ĉ�����"
End If

Set fileSystemObject = Nothing
