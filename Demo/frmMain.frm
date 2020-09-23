VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Resizing Demo"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "View 2nd example"
      Height          =   345
      Left            =   1830
      TabIndex        =   10
      Tag             =   "_Label3:::§~§"
      Top             =   0
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore initial size"
      Height          =   345
      Left            =   180
      TabIndex        =   9
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "....which can be typically used for a status bar (static top and width)"
      Height          =   435
      Left            =   30
      TabIndex        =   8
      Tag             =   "0:2:2:0§~§"
      Top             =   4200
      Width           =   4845
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anchoring the control height to the form lets you preserve a fixed zone below..."
      Height          =   435
      Left            =   150
      TabIndex        =   7
      Tag             =   "0:75_Label5:1:915§~§"
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "This label grows and moves proportionally, see how the left gap grows too"
      Height          =   1395
      Left            =   1410
      TabIndex        =   6
      Tag             =   "1:1:1:1§~§"
      Top             =   2250
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "This label has the top anchored to the orange one, and the left is right-anchored to the cyan one"
      Height          =   1065
      Left            =   2670
      TabIndex        =   5
      Tag             =   "5_Label1:135_Label4:0:0§~§"
      Top             =   2130
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "This label grows and moves down proportionally, see how the upper gap grows too"
      Height          =   1395
      Left            =   150
      TabIndex        =   4
      Tag             =   "0:1:1:1§~§"
      Top             =   2250
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"frmMain.frx":0000
      Height          =   1275
      Left            =   2670
      TabIndex        =   3
      Tag             =   "135_Label3:45_Label1:0:5_Label2§~§"
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "This label has the top anchored to the cyan one and the left to the yellow one"
      Height          =   1395
      Left            =   1410
      TabIndex        =   2
      Tag             =   "135_Label2:45_Label1:0:0§~§"
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "This label grows proportionally and moves down proportionally"
      Height          =   1395
      Left            =   150
      TabIndex        =   1
      Tag             =   "0:1:1:1§~§"
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "This label grows proportionally and doesn't move"
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Tag             =   "0:0:1:1§~§"
      Top             =   390
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Resizer As New ControlResizer

Private lWidth As Long
Private lHeight As Long


Private Sub Command1_Click()
On Error GoTo ShowErr
    Me.Move Me.Left, Me.Top, lWidth, lHeight
    Exit Sub
ShowErr:
    MsgBox "Error " & Err.Number & " : " & Err.Description, vbInformation, "Alert"
    Err.Clear
End Sub

Private Sub Command2_Click()
    frmEntry.Show vbModal
End Sub

Private Sub Form_Load()
    lWidth = Me.Width
    lHeight = Me.Height
    Resizer.InitResizer Me, Me.Width, Me.Height, True
End Sub

Private Sub Form_Resize()
    Resizer.FormResized Me
End Sub
