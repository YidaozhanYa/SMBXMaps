VERSION 5.00
Begin VB.Form frmDummy 
   BackColor       =   &H80000005&
   Caption         =   "请稍后 ..."
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   5370
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "正在加载中 ..."
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowDummy(DummyText As String, Optional DummyText2 As String = "")
Label1.Caption = DummyText & vbCrLf & DummyText2
Me.Show
End Sub

