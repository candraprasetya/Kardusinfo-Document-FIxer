VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardusinfo Doc Fixer"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "SF Pro Display"
      Size            =   9.75
      Charset         =   0
      Weight          =   200
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   0
      Left            =   360
      ScaleHeight     =   4815
      ScaleWidth      =   9495
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   9495
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2640
         Top             =   3480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "formMain.frx":0000
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         Picture         =   "formMain.frx":23A2
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "We are Scanning your Active Diskdrive and Repair your Files."
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Location"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1200
         TabIndex        =   12
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Report:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1200
         TabIndex        =   11
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Files:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   1200
         TabIndex        =   8
         Top             =   2280
         Width           =   390
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   7
         Top             =   2280
         Width           =   90
      End
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Threat Detected"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Report"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox PicMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   2
      Left            =   360
      ScaleHeight     =   4815
      ScaleWidth      =   9495
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Muh.Isfahani Ghiyath.YM"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   15
         Left            =   1560
         TabIndex        =   26
         Top             =   2280
         Width           =   1875
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks to:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   25
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://kardusinfo.com"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   1560
         TabIndex        =   24
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   23
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programmer:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   22
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Candra Ramadhan Prasetya"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   10
         Left            =   1560
         TabIndex        =   21
         Top             =   1440
         Width           =   2085
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reconstruct by Kardusinfo."
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   2130
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kardusinfo Doc Fixer."
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   2490
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unhidden your Files."
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Index           =   13
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   1620
      End
   End
   Begin VB.PictureBox PicMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   1
      Left            =   360
      ScaleHeight     =   4815
      ScaleWidth      =   9495
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   9495
      Begin MSComctlLib.ListView lvFile 
         Height          =   3375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CommandButton cmdFix 
         Caption         =   "Fix All Items.."
         Height          =   375
         Left            =   7320
         TabIndex        =   15
         Top             =   4080
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Berhenti As Boolean
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Const MAX_PATH = 260, MAXDWORD = &HFFFF, INVALID_HANDLE_VALUE = -1, FILE_ATTRIBUTE_ARCHIVE = &H20, FILE_ATTRIBUTE_DIRECTORY = &H10, FILE_ATTRIBUTE_HIDDEN = &H2, FILE_ATTRIBUTE_NORMAL = &H80, FILE_ATTRIBUTE_READONLY = &H1, FILE_ATTRIBUTE_SYSTEM = &H4, FILE_ATTRIBUTE_TEMPORARY = &H100, ALTERNATE As Long = 14
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Event AnalisisGetHidden(path As String, file As String, filekah As Boolean)

Public Function IsHidden(path As String) As Boolean
    If GetFileAttributes(StrPtr(path)) And FILE_ATTRIBUTE_HIDDEN Then IsHidden = True Else IsHidden = False
End Function
Public Function NormalHidden(path As String) As Boolean
    NormalHidden = SetFileAttributes(StrPtr(path), FILE_ATTRIBUTE_NORMAL)
End Function
Sub ScanFile(ByVal Fol As Scripting.Folder)
    If Berhenti = True Then Exit Sub
  
    On Error Resume Next
    Dim FI As Scripting.file
    Dim fo As Scripting.Folder
    Dim ada As Long
    Dim lv As ListItem

    For Each FI In Fol.Files
    DoEvents
    If Berhenti = True Then Exit Sub
    lblinfo(1).Caption = FI.Name

         If IsHidden(FI.path) = True Then
         Set lv = lvFile.ListItems.Add(, , FI.Name) ', , ImageList1.ListImages(2).Index)
         lv.SubItems(1) = "Hidden File"
         lv.SubItems(2) = FI.path
         lblinfo(6).Caption = lvFile.ListItems.Count
         End If
    Next

    For Each fo In Fol.SubFolders
        ScanFile fo
    Next
End Sub

Private Sub cmdFix_Click()
    Dim i As Long
    Dim atribut As Integer
    For i = 1 To lvFile.ListItems.Count
        atribut = NormalHidden(lvFile.ListItems(i).SubItems(2))
    Next
    If atribut = 1 Then
        MsgBox "Complete", vbInformation, "Dokper"
        lvFile.ListItems.Clear
    End If
End Sub

Private Sub cmdScan_Click()
    Dim GetFolder As String
    GetFolder = BrowseForFolder(Me.hwnd, "Pilih Folder yang akan di pindai.")
    txtPath.Text = GetFolder
    If txtPath.Text <> "" Then
        Berhenti = False
        lvFile.ListItems.Clear
        LookForHiddenFolder txtPath.Text
        lblinfo(6).Caption = "0"
        cmdScan.Enabled = False
        cmdStop.Enabled = True
        Dim fso As New FileSystemObject
        ScanFile fso.GetFolder(txtPath.Text)
        Berhenti = True
        cmdScan.Enabled = True
        cmdStop.Enabled = False
    Else
        MsgBox "Path not found!", vbCritical, "Dokper"
    End If
End Sub

Private Sub cmdStop_Click()
    cmdScan.Enabled = True
    cmdStop.Enabled = False
    Berhenti = True
End Sub


Private Sub Command1_Click()
PicMenu(0).Visible = True: PicMenu(1).Visible = False: PicMenu(2).Visible = False
End Sub

Private Sub Command2_Click()
PicMenu(1).Visible = True: PicMenu(0).Visible = False: PicMenu(2).Visible = False
End Sub

Private Sub Command3_Click()
PicMenu(2).Visible = True: PicMenu(1).Visible = False: PicMenu(0).Visible = False
End Sub

Private Sub Form_Load()
    lvwStyle lvFile
    SetFlatHeaders lvFile.hwnd
End Sub

Sub LookForHiddenFolder(PathFile As String)
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Dim fld2 As Folder
    Dim a As String
    Dim ls As ListItem
    Set fld = fso.GetFolder(PathFile)
    For Each fld2 In fld.SubFolders
        If fld2.Attributes = Hidden + System + ReadOnly Or fld2.Attributes = 18 Or fld2.Attributes = 22 Or fld2.Attributes = Hidden Or fld2.Attributes = 23 Then
        Set ls = lvFile.ListItems.Add(, , fld2.Name) ', , ImageList1.ListImages(1).Index)
           ls.SubItems(1) = "Hidden Folder"
           ls.SubItems(2) = PathFile & "\" & fld2.Name
        End If
    Next
End Sub

