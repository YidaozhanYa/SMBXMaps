Attribute VB_Name = "modCommonControlsLoader"
Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" ( _
    iccex As tagInitCommonControlsEx) As Boolean

Private Sub InitCommonControls()
    Dim iccex As tagInitCommonControlsEx
    
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    On Error Resume Next
    InitCommonControlsEx iccex
End Sub

Private Sub Main() '´Ó´ËÆô¶¯
    InitCommonControls
    frmBrowse.Show
End Sub


