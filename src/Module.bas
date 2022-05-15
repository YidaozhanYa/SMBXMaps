Attribute VB_Name = "Module"
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'end api header

Public Const MapServer As String = "http://azure2.yidaozhan.ga:5244"
Public Const DummyPathURL As String = "路径 ..."


Public Function ChooseFile(ByVal frmTitle As String, ByVal fileDescription As String, ByVal fileFilter As String, ByVal onForm As Object) As String
'oleexp 选择文件
    On Error Resume Next
    Dim pChoose As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    Dim tFilt() As COMDLG_FILTERSPEC
    ReDim tFilt(0)
    tFilt(0).pszName = fileDescription
    tFilt(0).pszSpec = fileFilter
    With pChoose
        .SetFileTypes UBound(tFilt) + 1, VarPtr(tFilt(0))
        .SetTitle frmTitle
        .SetOptions FOS_FILEMUSTEXIST + FOS_DONTADDTORECENT
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseFile = BStrFromLPWStr(lpPath)
End Function


Public Function BStrFromLPWStr(lpWStr As Long) As String
    SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
End Function

Public Function GetRepoFolder(MapName As String) As String
    If IsNumeric(Left(MapName, 1)) Then
        GetRepoFolder = "0-9"
    Else
        Select Case Asc(UCase(Left(MapName, 1)))
        Case 60 To 90
            GetRepoFolder = UCase(Left(MapName, 1))
        Case Else
            GetRepoFolder = "Others"
        End Select
    End If
End Function

Public Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 0, True
    End With
End Sub

Public Function CheckFileExists(FilePath As String) As Boolean
'检查文件是否存在
    On Error GoTo Err
    If Len(FilePath) < 2 Then CheckFileExists = False: Exit Function
    If Dir$(FilePath, vbAllFileAttrib) <> vbNullString Then CheckFileExists = True
    Exit Function
Err:
    CheckFileExists = False
End Function

Public Function ReadTextFile(sFilePath As String) As String
    On Error Resume Next
    Dim handle As Integer
    If LenB(Dir$(sFilePath)) > 0 Then
        handle = FreeFile
        Open sFilePath For Binary As #handle
        ReadTextFile = Space$(LOF(handle))
        Get #handle, , ReadTextFile
        Close #handle
    End If
End Function


Public Function GetExt(FilePath As String) As String
GetExt = Mid(FilePath, InStrRev(FilePath, ".") + 1, Len(FilePath))
End Function
