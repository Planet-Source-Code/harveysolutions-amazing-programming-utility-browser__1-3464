VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form4"
   ScaleHeight     =   1950
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3960
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "File Search in Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1920
      Left            =   15
      Top             =   15
      Width           =   4680
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim firstt As Boolean

Private Sub Command1_Click()
ProgressCancel = True
End Sub

Private Sub Form_Load()
firstt = True
Top = Form1.Top + ((Form1.Height / 2)) - (Height / 2)
Left = (Form1.Left + (Form1.Width / 2)) - (Width / 2)
End Sub

Private Sub Form_Paint()
If firstt Then Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
firstt = fasle
Form1.FindFile Left(Form1.Drive1.Drive, 1) & ":\", Form1.Combo1.Text
Timer1.Enabled = False
Unload Me
End Sub
