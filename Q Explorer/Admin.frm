VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Q Explorer"
   ClientHeight    =   3465
   ClientLeft      =   1755
   ClientTop       =   1770
   ClientWidth     =   4800
   Icon            =   "Admin.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Text            =   "Program Name"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Delete Program"
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
      Begin VB.CommandButton Command2 
         Caption         =   "Delete program"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Program Location"
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Insert program"
      Height          =   1935
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command1 
         Caption         =   "Save Program"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "Admin.frx":0CCA
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   4800
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Double click on a program to run it"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileAddProgram 
         Caption         =   "Add Program"
      End
      Begin VB.Menu mnuFileDeleteProgram 
         Caption         =   "Delete Program"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuLinks 
      Caption         =   "Links"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
file$ = InputBox("Program Name:", "Add a program")
file$ = file$ + ".qep"
Open file$ For Append As #1
Print #1, Text1.Text
Print #1, Text2.Text
Close #1
MsgBox "Your Program has been saved", 0, ""
GoTo en
er:
MsgBox " Uh oh, You did something wrong!! Make sure you filled out everything!!!"
en:

End Sub


Private Sub Command2_Click()
    Kill Text11.Text
    MsgBox "The program has been deleted", 0, ""
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub File1_Click()
Text11.Text = File1.FileName
End Sub

Private Sub File1_DblClick()
Open File1.FileName For Input As #1
Input #1, pn
Input #1, pl
Text7.Text = pl
Shell (Text7.Text), vbNormalFocus
Close #1

End Sub


Private Sub Form_Load()
File1.Pattern = "*.qep"
End Sub



Private Sub mnuFileAddProgram_Click()
    MsgBox "In the Program Location box, next to the list, enter the location of the program and click on the Save Program button", vbOKOnly, "How To Add A Program"
End Sub

Private Sub mnuFileDeleteProgram_Click()
    MsgBox "Click on a program in the list and click on Delete Program", vbOKOnly, "Delete A Program"
End Sub

Private Sub Timer1_Timer()
File1.Refresh
End Sub


