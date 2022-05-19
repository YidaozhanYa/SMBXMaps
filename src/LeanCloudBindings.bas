Attribute VB_Name = "Lncld"
'LeanCloud database VB6 bindings
Public Const LncldBaseUrl As String = "https://wgcr3xhd.api.lncldglobal.com" ' LeanCloud REST API base url
Public Const LncldAppID As String = "wgcr3xHDSmfiaOJReHtlqD9z-MdYXbMMI" ' LeanCloud App ID
Public Const LncldAppKey As String = "7vDim8MYqChNNgt2D8NkFjtP" ' LeanCloud App Public Key


Function CreateLeanObject(jsonData As String, ClassName As String) As String
    With New MSXML2.ServerXMLHTTP30
        .Open "POST", LncldBaseUrl & "/1.1/classes/" & ClassName, True
        .setRequestHeader "X-LC-Id", LncldAppID
        .setRequestHeader "X-LC-Key", LncldAppKey
        .setRequestHeader "Content-Type", "application/json"
        .send jsonData
        Do While .readyState = 1
            Sleep 20
            DoEvents
        Loop
        Sleep 20
        If .readyState = 4 Then
            CreateLeanObject = .responseText
        Else
            MsgBox "发生错误！" & vbCrLf & "返回值：" & .readyState & vbCrLf & "错误：" & .Status & " " & .statusText, vbCritical
            MsgBox .responseText
            End
        End If
    End With
End Function

Function Query(jsonQuery As String, ClassName As String, Optional Limit As Integer = 0, Optional Skip As Integer = 0) As String
    With New MSXML2.ServerXMLHTTP30
    Dim OptionalValues As String
    If Limit <> 0 Then OptionalValues = OptionalValues & "&limit=" & CStr(Limit)
    If Skip <> 0 Then OptionalValues = OptionalValues & "&skip=" & CStr(Skip)
        .Open "GET", LncldBaseUrl & "/1.1/classes/" & ClassName & "?where=" & jsonQuery & OptionalValues, True
        .setRequestHeader "X-LC-Id", LncldAppID
        .setRequestHeader "X-LC-Key", LncldAppKey
        .setRequestHeader "Content-Type", "application/json"
        .send
        Do While .readyState <> 4
            Sleep 20
            DoEvents
        Loop
        Sleep 40
        If .readyState = 4 Then
            Query = .responseText
        Else
            MsgBox "发生错误！" & vbCrLf & "错误：" & .Status & " " & .statusText, vbCritical
            MsgBox .readyState
            MsgBox .responseText
            End
        End If
    End With
End Function

Function Count(jsonQuery As String, ClassName As String) As Integer
    With New MSXML2.ServerXMLHTTP30
        .Open "GET", LncldBaseUrl & "/1.1/classes/" & ClassName & "?count=1&where=" & jsonQuery, True
        .setRequestHeader "X-LC-Id", LncldAppID
        .setRequestHeader "X-LC-Key", LncldAppKey
        .setRequestHeader "Content-Type", "application/json"
        .send
        Do While .readyState <> 4
            Sleep 20
            DoEvents
        Loop
        Sleep 40
        If .readyState = 4 Then
            Count = JSON.parse(.responseText)("count")
        Else
            MsgBox "发生错误！" & vbCrLf & "错误：" & .Status & " " & .statusText, vbCritical
            MsgBox .readyState
            MsgBox .responseText
            End
        End If
    End With
End Function

