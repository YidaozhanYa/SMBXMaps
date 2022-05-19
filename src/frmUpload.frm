VERSION 5.00
Begin VB.Form frmUpload 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMBX ��ͼ�ֿ� - ����Ͷ��"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "΢���ź� Light"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdPost 
      Caption         =   "Ͷ��!"
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
      Text            =   "��ѡ��"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.CommandButton cmdDesc 
      Caption         =   "��д"
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
      Text            =   "·�� / URL ..."
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "���"
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
      Caption         =   "������ַ��"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2535
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͼ��飺"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͼ���ߣ�"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2055
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͼ�ļ���"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1605
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ϸ�汾��"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1125
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͼ����"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "���ͼ�ֿ�Ͷ���ͼ"
      BeginProperty Font 
         Name            =   "΢���ź� Light"
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
        MsgBox "����д����Ϣ��", vbCritical
        Exit Sub
    End If
    If txtURL.Text = dummytexturl Then
        MsgBox "��ѡ���ļ���", vbCritical
        Exit Sub
    End If
    If Not CheckFileExists(txtURL.Text) Then
        MsgBox "���������ļ�������!"
        Exit Sub
    End If
    frmDummy.ShowDummy "����Ͷ���� ...", "���ڴ�������"
    Dim Postfields As New Dictionary, RetVal As Object, FilePath As String
    With Postfields
        .Add "name", txtMapName.Text
        .Add "maker", txtMaker.Text
        Select Case cbGameVersion.Text
        Case "[38A] 1.4.5": .Add "version", "1.4.5"
        Case "[38A] 1.4.4": .Add "version", "1.4.4"
        Case "[38A] ����"
            If txtGameVersion.Text <> "" Then
                .Add "version", txtGameVersion.Text
            Else
                MsgBox "�汾��Ϣ����", vbCritical
                Exit Sub
            End If
        Case "����"
            If txtGameVersion.Text <> "" Then
                .Add "version", txtGameVersion.Text
            Else
                MsgBox "�汾��Ϣ����", vbCritical
                Exit Sub
            End If
        Case "[ԭ��] 1.3": .Add "version", "1.3"
        Case "TheXTech / LunaDLL": .Add "version", "TheXTech"
        End Select
        If txtPublishUrl.Text <> "��ѡ��" Then
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
    frmDummy.ShowDummy "����Ͷ���� ...", "���ڸ�����ʱ�ļ�"
        FilePath = Environ("Temp") & "\[" & Postfields.Item("version") & "] " & txtMapName.Text & "." & GetExt(txtURL.Text)
        FileCopy txtURL.Text, FilePath
    frmDummy.ShowDummy "����Ͷ���� ...", "�����ϴ��ļ�"
    ShellAndWait "cmd /c """ & App.Path & "\curl.exe " & """" & MapUploadServer & "/api/public/upload"" -X POST -H ""authorization:" & MapServerToken & """ -F ""path=/SMBX/" & GetRepoFolder(txtMapName.Text) & """ -F ""files=@" & FilePath & """ > """ & App.Path & "\Temp.txt" & """"
    Sleep 20
    If JSON.parse(ReadTextFile(App.Path & "\Temp.txt"))("code") <> 200 Then
        MsgBox "�������� " & ReadTextFile(App.Path & "\Temp.txt")
        frmDummy.Hide
        Exit Sub
    End If
    'MapServerToken �����������ڹ�����Դ������
    frmDummy.ShowDummy "����Ͷ���� ...", "����д���ݿ�"
    Set RetVal = JSON.parse(Lncld.CreateLeanObject(JSON.toString(Postfields), "DB"))
    frmDummy.Hide
    MsgBox "Ͷ��ɹ���"
End Sub

Private Sub Form_Load() ' ��ʼ��
    With cbGameVersion
        .Clear
        .AddItem "[38A] 1.4.5"
        .AddItem "[38A] 1.4.4"
        .AddItem "[38A] ����"
        .AddItem "[ԭ��] 1.3"
        .AddItem "TheXTech / LunaDLL"
        .AddItem "[LunaLua] 2.0"
        .AddItem "����"
        .Text = "[38A] 1.4.5"
    End With
    txtGameVersion.Visible = False
    txtURL.Text = DummyPathURL
End Sub

Private Sub cbGameVersion_Click() ' �ж��Ƿ�Ϊ����
    If InStr(cbGameVersion.Text, "����") Then
    txtGameVersion.Visible = True
    Else
    txtGameVersion.Visible = False
    End If
End Sub


Private Sub cmdDesc_Click()
MapDesc = InputBox("��д��ͼ��� ...", "Ͷ���ͼ", MapDesc)
End Sub


'URLUtility.URLEncode(txtToEncode.Text, vbUnchecked)

Private Sub cbGameVersion_KeyPress(KeyAscii As Integer) ' ����ֹ�Ӵ�����
KeyAscii = 0
End Sub


Public Function ChooseMap() As String
    On Error Resume Next
    Dim pChoose As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    Dim tFilt() As COMDLG_FILTERSPEC
    ReDim tFilt(0 To 5)
    tFilt(0).pszName = "ZIP ѹ����"
    tFilt(0).pszSpec = "*.zip"
    tFilt(1).pszName = "RAR ѹ����"
    tFilt(1).pszSpec = "*.rar"
    tFilt(2).pszName = "7Z ѹ����"
    tFilt(2).pszSpec = "*.7z"
    tFilt(3).pszName = "GZIP ѹ����"
    tFilt(3).pszSpec = "*.gz"
    tFilt(4).pszName = "ZSTD ѹ����"
    tFilt(4).pszSpec = "*.zst"
    tFilt(5).pszName = "TAR �ļ��鵵"
    tFilt(5).pszSpec = "*.tar"
    With pChoose
        .SetFileTypes UBound(tFilt) + 1, VarPtr(tFilt(0))
        .SetTitle "ѡ���ͼѹ���� ..."
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
