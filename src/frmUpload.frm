VERSION 5.00
Begin VB.Form frmUpload 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMBX 地图仓库 - 申请投稿"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "微软雅黑 Light"
      Size            =   10.5
      Charset         =   134
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4485
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdPost 
      Caption         =   "投稿!"
      Height          =   420
      Left            =   3120
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtPublishUrl 
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Text            =   "可选填"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.CommandButton cmdDesc 
      Caption         =   "填写"
      Height          =   420
      Left            =   1200
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtMaker 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtURL 
      Height          =   420
      Left            =   1200
      TabIndex        =   8
      Text            =   "路径 / URL ..."
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "浏览"
      Height          =   420
      Left            =   3120
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtGameVersion 
      Height          =   420
      Left            =   3120
      TabIndex        =   5
      Text            =   "1.4.0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cbGameVersion 
      Height          =   420
      Left            =   1200
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtMapName 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "发布网址："
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2535
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "地图简介："
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "地图作者："
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2055
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "地图文件："
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1605
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "游戏版本："
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1125
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "地图名："
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "向地图仓库投稿地图"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MapDesc As String

Private Sub cmdBrowse_Click()
    txtURL.Text = ChooseMap
End Sub

Private Sub cmdPost_Click()
    If txtMapName.Text = "" Or txtMaker.Text = "" Or txtURL.Text = "" Then
        MsgBox "请填写完信息！", vbCritical
        Exit Sub
    End If
    If txtURL.Text = dummytexturl Then
        MsgBox "请选择文件！", vbCritical
        Exit Sub
    End If
    If Not CheckFileExists(txtURL.Text) Then
        MsgBox "发生错误：文件不存在!"
        Exit Sub
    End If
    frmDummy.ShowDummy "正在投稿中 ...", "正在处理数据"
    Dim Postfields As New Dictionary, RetVal As Object, FilePath As String
    With Postfields
        .Add "name", txtMapName.Text
        .Add "maker", txtMaker.Text
        Select Case cbGameVersion.Text
        Case "[38A] 1.4.5": .Add "version", "1.4.5"
        Case "[38A] 1.4.4": .Add "version", "1.4.4"
        Case "[38A] 其它"
            If txtGameVersion.Text <> "" Then
                .Add "version", txtGameVersion.Text
            Else
                MsgBox "版本信息错误", vbCritical
                Exit Sub
            End If
        Case "其它"
            If txtGameVersion.Text <> "" Then
                .Add "version", txtGameVersion.Text
            Else
                MsgBox "版本信息错误", vbCritical
                Exit Sub
            End If
        Case "[原版] 1.3": .Add "version", "1.3"
        Case "TheXTech / LunaDLL": .Add "version", "TheXTech"
        End Select
        If txtPublishUrl.Text <> "可选填" Then
            .Add "puburl", txtPublishUrl.Text
        Else
            .Add "puburl", ""
        End If
        If MapDesc <> "" Then
            .Add "desc", Base64Encode(MapDesc)
        Else
            .Add "desc", ""
        End If
        .Add "status", "pending"
        .Add "rel", "Initial"
        .Add "ext", GetExt(txtURL.Text)
        .Add "repofolder", GetRepoFolder(txtMapName.Text)
    End With
    frmDummy.ShowDummy "正在投稿中 ...", "正在复制临时文件"
        FilePath = Environ("Temp") & "\[" & Postfields.Item("version") & "] " & txtMapName.Text & "." & GetExt(txtURL.Text)
        FileCopy txtURL.Text, FilePath
    frmDummy.ShowDummy "正在投稿中 ...", "正在上传文件"
    ShellAndWait "cmd /c """ & App.Path & "\curl.exe " & """" & MapUploadServer & "/api/public/upload"" -X POST -H ""authorization:" & MapServerToken & """ -F ""path=/SMBX/" & GetRepoFolder(txtMapName.Text) & """ -F ""files=@" & FilePath & """ > """ & App.Path & "\Temp.txt" & """"
    Sleep 20
    If JSON.parse(ReadTextFile(App.Path & "\Temp.txt"))("code") <> 200 Then
        MsgBox "发生错误 " & ReadTextFile(App.Path & "\Temp.txt")
        frmDummy.Hide
        Exit Sub
    End If
    'MapServerToken 变量不包含在公开的源代码中
    frmDummy.ShowDummy "正在投稿中 ...", "正在写数据库"
    Set RetVal = JSON.parse(Lncld.CreateLeanObject(JSON.toString(Postfields), "DB"))
    frmDummy.Hide
    MsgBox "投稿成功！"
End Sub

Private Sub Form_Load() ' 初始化
    With cbGameVersion
        .Clear
        .AddItem "[38A] 1.4.5"
        .AddItem "[38A] 1.4.4"
        .AddItem "[38A] 其它"
        .AddItem "[原版] 1.3"
        .AddItem "TheXTech / LunaDLL"
        .AddItem "[LunaLua] 2.0"
        .AddItem "其它"
        .Text = "[38A] 1.4.5"
    End With
    txtGameVersion.Visible = False
    txtURL.Text = DummyPathURL
End Sub

Private Sub cbGameVersion_Click() ' 判断是否为其它
    If InStr(cbGameVersion.Text, "其它") Then
    txtGameVersion.Visible = True
    Else
    txtGameVersion.Visible = False
    End If
End Sub


Private Sub cmdDesc_Click()
MapDesc = InputBox("填写地图简介 ...", "投稿地图", MapDesc)
End Sub


'URLUtility.URLEncode(txtToEncode.Text, vbUnchecked)

Private Sub cbGameVersion_KeyPress(KeyAscii As Integer) ' 【禁止接触！】
KeyAscii = 0
End Sub


Public Function ChooseMap() As String
    On Error Resume Next
    Dim pChoose As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    Dim tFilt() As COMDLG_FILTERSPEC
    ReDim tFilt(0 To 5)
    tFilt(0).pszName = "ZIP 压缩包"
    tFilt(0).pszSpec = "*.zip"
    tFilt(1).pszName = "RAR 压缩包"
    tFilt(1).pszSpec = "*.rar"
    tFilt(2).pszName = "7Z 压缩包"
    tFilt(2).pszSpec = "*.7z"
    tFilt(3).pszName = "GZIP 压缩包"
    tFilt(3).pszSpec = "*.gz"
    tFilt(4).pszName = "ZSTD 压缩包"
    tFilt(4).pszSpec = "*.zst"
    tFilt(5).pszName = "TAR 文件归档"
    tFilt(5).pszSpec = "*.tar"
    With pChoose
        .SetFileTypes UBound(tFilt) + 1, VarPtr(tFilt(0))
        .SetTitle "选择地图压缩包 ..."
        .SetOptions FOS_FILEMUSTEXIST + FOS_DONTADDTORECENT
        .Show frmUpload.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseMap = BStrFromLPWStr(lpPath)
End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
