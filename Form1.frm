VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Programming Utilities"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   9495
      TabIndex        =   0
      Top             =   480
      Width           =   9495
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         Height          =   375
         Left            =   7800
         TabIndex        =   18
         Top             =   3720
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   5040
         List            =   "Form1.frx":000A
         TabIndex        =   17
         Text            =   "*.FRM"
         Top             =   0
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   7680
         TabIndex        =   14
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find FRM files"
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   3720
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid Grille1 
         Height          =   2895
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
      End
      Begin VB.Label Label8 
         Caption         =   "File Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Drive :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   15
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Files"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   3240
         Width           =   9375
      End
      Begin VB.Label Label2 
         Caption         =   "Files Found"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   3840
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   9495
      TabIndex        =   5
      Top             =   480
      Width           =   9495
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   7800
         TabIndex        =   19
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search Function and Sub"
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include ''_"" character"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   3840
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid Grille2 
         Height          =   3015
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5318
         _Version        =   393216
         FixedCols       =   0
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   3240
         Width           =   9375
      End
      Begin VB.Label Label4 
         Caption         =   "Functions/Subs Found"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "VB Functions and Subs"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search File"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Function"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContent 
         Caption         =   "Help Content"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuSPMenu 
      Caption         =   "SPMenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rightmouse As Boolean
Dim okMNU As Boolean
Dim posx As Single
Dim posy As Single
Private Sub Command1_Click()
Grille2.Clear
Form2.Show 1
okMNU = True
End Sub
Public Sub AddFunctionSub(ByVal item As String, ByVal fpath As String)
With Grille2
     .AddItem item
     .Row = .Rows - 1
     .Col = 1
     .Text = fpath
End With
End Sub
Private Sub Command2_Click()
ProgressCancel = False
Grille1.Clear
Grille2.Clear
Grille1.Rows = 1
Grille2.Rows = 1
InitGrille1
InitGrille2
NbFile = 0
okMNU = False
Command1.Enabled = False
Form4.Show 1
Command1.Enabled = True
okMNU = True
End Sub


Private Sub Command3_Click()
End
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()
Command1.Enabled = False
ProgressCancel = False
Picture1.ZOrder
Drive1.Drive = "C:\"
InitGrille1
InitGrille2
End Sub
Private Sub InitGrille2()
With Grille2
.Cols = 2
.Rows = 1
.Row = 0
.Col = 0
.Text = "Function/Sub"
.Col = 1
.Text = "File"
.ColWidth(0) = 4550
.ColWidth(1) = 4550
.Width = 4550 + 4550 + 350
End With
End Sub
Private Sub InitGrille1()
With Grille1
.Cols = 3
.Rows = 1
.Row = 0
.Col = 0
.Text = "File Name"
.Col = 1
.Text = "Path"
.Col = 2
.Text = "Size"
.ColWidth(0) = 2550
.ColWidth(1) = 4950
.ColWidth(2) = 1600
.Width = 1600 + 4950 + 2550 + 350
End With
End Sub
'**********************************************
'* Function FindFile is From Planet-Source-Code
'* Strongly modified by Carlos 09-10-99
'***********************************************
Public Sub FindFile(ByVal path As String, ByVal ftype As String)
       Dim hFile As Long, ts As String, WFD As WIN32_FIND_DATA
       Dim result As Long, sAttempt As String, szPath As String
       Dim Strtemp
       If ProgressCancel Then Exit Sub
       Form4.ProgressBar1.Value = 1
       szPath = path & "*.*" & Chr$(0)
       'Start asking windows for files.
       putfileinpath path, ftype
       hFile = FindFirstFile(szPath, WFD)
       Do
         If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
          'Hey look, we've got a directory!
             ts = StripNull(WFD.cFileName)
             If Not (ts = "." Or ts = "..") Then
                 'Don't look for hidden or system directories
                 If Not (WFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
                     'Search directory recursively
                     FindFile path & ts & "\", ftype
                 End If
             End If
           End If
           WFD.cFileName = ""
           result = FindNextFile(hFile, WFD)
           Label1.Caption = "Searching in: " & path
           DoEvents
          If Form4.ProgressBar1.Value = Form4.ProgressBar1.Max Then Form4.ProgressBar1.Value = 1
          Form4.ProgressBar1.Value = Form4.ProgressBar1.Value + 1
        Loop Until result = 0
       FindClose hFile
End Sub
'**********************************************
'* Function putfileinpath is From Planet-Source-Code
'* Modified by Carlos 09-10-99
'***********************************************
Private Sub putfileinpath(ByVal zpath As String, ByVal FileType As String)
       Dim hFile As Long, result As Long, szPath As String
       Dim WFD As WIN32_FIND_DATA
       szPath = zpath & FileType & Chr$(0)
       'Start asking windows for files.
       hFile = FindFirstFile(szPath, WFD)
       Dim pos1
       Do
           pos1 = InStr(1, WFD.cFileName, Chr$(0), vbBinaryCompare)
           If Trim(Mid(WFD.cFileName, 1, pos1 - 1)) <> "" Then
              AddAfile WFD, zpath
           End If
           WFD.cFileName = ""
           result = FindNextFile(hFile, WFD)
          ' DoEvents
       Loop Until result = 0
       FindClose hFile
End Sub

Private Sub AddAfile(WFDP As WIN32_FIND_DATA, ByVal path As String)
          NbFile = NbFile + 1
          With Grille1
             .AddItem Trim(WFDP.cFileName)
             .Row = NbFile
             .Col = 1
             .Text = path
             .Col = 2
             .Text = WFDP.nFileSizeLow / 1000 & " Kb   "
          End With
          Label2.Caption = NbFile & " Files Found."
End Sub


Private Sub Grille2_Click()
If okMNU Then
  If Not rightmouse Then
    'MsgBox Grille1.Row
    PopupMenu mnuSPMenu
    
    rightmouse = False
  End If
End If
End Sub

Private Sub Grille2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
 rightmouse = True
End If
End Sub

Private Sub mnuAbout_Click()
Form3.Show 1
End Sub

Private Sub mnuCopy_Click()
MsgBox "Put copy to clipboard code HERE!"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuView_Click()
If Grille2.Col = 1 Then
  TypeView = 1
Else
  TypeView = 0
End If
StringToFind = Grille2.Text
Grille2.Col = 1
FileFSToOpen = Grille2.Text
Form6.Show 1
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem = "Search File" Then
 Picture1.ZOrder
Else
Picture2.ZOrder
End If
End Sub
