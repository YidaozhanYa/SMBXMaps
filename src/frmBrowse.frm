VERSION 5.00
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Begin VB.Form frmBrowse 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "΢���ź� Light"
      Size            =   10.5
      Charset         =   134
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   12120
   StartUpPosition =   3  '����ȱʡ
   Tag             =   "TL"
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "��ͼ��Ϣ"
      Height          =   6255
      Left            =   9120
      TabIndex        =   6
      Tag             =   "LH"
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdDownload 
         Caption         =   "����"
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Tag             =   "T"
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ�е�ͼ�Բ鿴��Ϣ"
         Height          =   5175
         Left            =   120
         TabIndex        =   7
         Tag             =   "H"
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "����"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Tag             =   "TL"
      Top             =   5760
      Width           =   1095
   End
   Begin VBCCR17.ImageList ImageList1 
      Left            =   9600
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      InitListImages  =   "frmBrowse.frx":54AA
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һҳ"
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Tag             =   "TL"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "��һҳ"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Tag             =   "TL"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Ͷ��"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Tag             =   "T"
      Top             =   5760
      Width           =   1215
   End
   Begin VBCCR17.ListView lst 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Tag             =   "HW"
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9763
      VisualTheme     =   1
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColumnHeaderIcons=   "ImageList1"
      GroupIcons      =   "ImageList1"
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   1
   End
   Begin VB6ResizerLib2.VB6Resizer VB6Resizer1 
      Left            =   8280
      Top             =   5280
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "���ڼ����� ..."
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5850
      Width           =   3615
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurrentPage As Integer, MaxPage As Integer, MaxCount As Integer, CurrentMaps As New Dictionary
Attribute MaxPage.VB_VarUserMemId = 1073938432
Attribute MaxCount.VB_VarUserMemId = 1073938432
Attribute CurrentMaps.VB_VarUserMemId = 1073938432



Private Sub cmdSearch_Click()
    Dim Search As String
    Search = InputBox("�����ͼ���ؼ��� ...", "����")
    If Search <> "" Then SearchMaps (Search)
End Sub

Private Sub cmdUpload_Click()
    frmUpload.Show
End Sub

Private Sub Form_Load()
    Me.Caption = "SMBX ��ͼ�ֿ� v" & App.Major & "." & App.Minor & "." & App.Revision

    '��ʼ�����б�
    lst.ColumnHeaders.Clear
    lst.ListItems.Clear
    lst.ColumnHeaders.Add 1, "version", "�汾", 1200
    lst.ColumnHeaders.Add 2, "name", "��ͼ", 6850
    lst.ColumnHeaders.Add 3, "maker", "����", 1700

    Me.Show
    DoEvents    '���ô������

    lst.ListItems.Add 1, "loading", "", , "unknown"
    lst.ListItems.Add 2, "loading2", ""
    lst.ListItems(1).SubItems(1) = "���ڼ��� ..."
    lst.ListItems(2).SubItems(1) = "���Ժ�"
    DoEvents


    cmdSearch.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    MaxCount = Lncld.Count("", "DB")
    MaxPage = (MaxCount \ 50) + 1
    lbl.Caption = "���� " & MaxCount & " �ŵ�ͼ��1/" & MaxPage

    CurrentPage = 1
    LoadListPage CurrentPage
End Sub

Private Sub LoadListPage(Page As Integer)
    cmdSearch.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    lbl.Caption = "���� " & MaxCount & " �ŵ�ͼ��" & Page & "/" & MaxPage
    DoEvents
    Dim SingleItem As Variant, ListIndex As Long
    ListIndex = 1
    CurrentMaps.RemoveAll
    lst.ListItems.Clear
    For Each SingleItem In JSON.parse(Lncld.Query("", "DB", 50, 50 * (Page - 1)))("results")
        lst.ListItems.Add ListIndex, SingleItem("objectId")
        If Left(SingleItem("version"), 3) = "1.4" Then    '38A
            lst.ListItems(ListIndex).SmallIcon = "38a"
        ElseIf SingleItem("version") = "1.3" Then    'legacy
            lst.ListItems(ListIndex).SmallIcon = "legacy"
        ElseIf Left(SingleItem("version"), 1) = "2" Then    'smbx2
            lst.ListItems(ListIndex).SmallIcon = "smbx2"
        ElseIf SingleItem("version") = "TheXTech" Or SingleItem("version") = "thextech" Then    'thextech
            lst.ListItems(ListIndex).SmallIcon = "thextech"
        Else
            lst.ListItems(ListIndex).SmallIcon = "unknown"
        End If
        If SingleItem("maker") = "Unknown" Then
            lst.ListItems(ListIndex).Text = "δ֪"
        Else
            lst.ListItems(ListIndex).Text = SingleItem("version")
        End If
        lst.ListItems(ListIndex).SubItems(1) = SingleItem("name")
        If SingleItem("maker") <> "Unknown" Then
            lst.ListItems(ListIndex).SubItems(2) = SingleItem("maker")
        Else
            lst.ListItems(ListIndex).SubItems(2) = "����"
        End If

        '������
        CurrentMaps.Add SingleItem("objectId"), SingleItem

        ListIndex = ListIndex + 1
    Next
    cmdSearch.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End Sub

Private Sub SearchMaps(SearchText As String)
    cmdSearch.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    lbl.Caption = "�������� ..."
    DoEvents
    Dim SingleItem As Variant, ListIndex As Long, SearchResults As Object
    ListIndex = 1
    CurrentMaps.RemoveAll
    lst.ListItems.Clear
    Set SearchResults = JSON.parse(Lncld.Query("{""name"":{""$regex"":""(?i)" & SearchText & """}}", "DB"))
    For Each SingleItem In SearchResults("results")
        lst.ListItems.Add ListIndex, SingleItem("objectId")
        If Left(SingleItem("version"), 3) = "1.4" Then    '38A
            lst.ListItems(ListIndex).SmallIcon = "38a"
        ElseIf SingleItem("version") = "1.3" Then    'legacy
            lst.ListItems(ListIndex).SmallIcon = "legacy"
        ElseIf Left(SingleItem("version"), 1) = "2" Then    'smbx2
            lst.ListItems(ListIndex).SmallIcon = "smbx2"
        ElseIf SingleItem("version") = "TheXTech" Or SingleItem("version") = "thextech" Then    'thextech
            lst.ListItems(ListIndex).SmallIcon = "thextech"
        Else
            lst.ListItems(ListIndex).SmallIcon = "unknown"
        End If
        If SingleItem("maker") = "Unknown" Then
            lst.ListItems(ListIndex).Text = "δ֪"
        Else
            lst.ListItems(ListIndex).Text = SingleItem("version")
        End If
        lst.ListItems(ListIndex).SubItems(1) = SingleItem("name")
        If SingleItem("maker") <> "Unknown" Then
            lst.ListItems(ListIndex).SubItems(2) = SingleItem("maker")
        Else
            lst.ListItems(ListIndex).SubItems(2) = "����"
        End If

        '������
        CurrentMaps.Add SingleItem("objectId"), SingleItem

        ListIndex = ListIndex + 1
    Next
    lbl.Caption = "������ " & CurrentMaps.Count & " �Ű��� " & SearchText & " �ĵ�ͼ"
    cmdSearch.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End Sub

Private Sub cmdNext_Click()
    If CurrentPage <> MaxPage Then
        CurrentPage = CurrentPage + 1
        LoadListPage CurrentPage
    End If
End Sub
Private Sub cmdPrev_Click()
    If CurrentPage <> 1 Then
        CurrentPage = CurrentPage - 1
        LoadListPage CurrentPage
    End If
End Sub

Private Sub lblInfo_Click()
    If lblInfo.ToolTipText <> "" Then Shell "cmd /c start """" """ & lblInfo.ToolTipText & """"
End Sub

Private Sub lst_Click()
If cmdNext.Enabled = False Then Exit Sub
    Dim SelectedMap As Object
    Set SelectedMap = CurrentMaps(lst.SelectedItem.key)
    lblInfo.ToolTipText = ""
    lblInfo.Caption = SelectedMap("name") & vbCrLf & _
    "����: " & SelectedMap("maker") & vbCrLf & _
    "�汾: " & SelectedMap("version")
    If SelectedMap("puburl") <> "" Then
    lblInfo.Caption = lblInfo.Caption & vbCrLf & SelectedMap("puburl")
    lblInfo.ToolTipText = SelectedMap("puburl")
    End If
    If SelectedMap("desc") <> "" Then lblInfo.Caption = lblInfo.Caption & vbCrLf & vbCrLf & Base64Decode(SelectedMap("desc"))
    lblInfo.Caption = lblInfo.Caption & vbCrLf & vbCrLf & "�ϴ��� " & Split(SelectedMap("createdAt"), "T")(0)
    Select Case SelectedMap("status")
        Case "pending":  lblInfo.Caption = lblInfo.Caption & vbCrLf & "״̬: ����ת�浫������"
        Case "hidden":  lblInfo.Caption = lblInfo.Caption & vbCrLf & "״̬: ����Դ���Ȩ����"
    End Select
End Sub

Private Sub cmdDownload_Click()
If cmdNext.Enabled = False Then Exit Sub
    Dim SelectedMap As Object
    Set SelectedMap = CurrentMaps(lst.SelectedItem.key)
    If SelectedMap("status") = "hidden" Then Exit Sub
    Shell "cmd /c start """" """ & MapDownloadServer & "/SMBX/" & SelectedMap("repofolder") & "/[" & SelectedMap("version") & "] " & SelectedMap("name") & "." & SelectedMap("ext") & """"
End Sub

Private Sub VB6Resizer1_AfterResize()
On Error Resume Next
    lst.ColumnHeaders(2).Width = Me.Width - 6580
End Sub

