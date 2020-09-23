VERSION 5.00
Begin VB.Form frmEntry 
   Caption         =   "Data Entry"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPriv 
      Height          =   345
      Left            =   60
      TabIndex        =   16
      Tag             =   "::-225_txtZip:§~§"
      Text            =   "Anchored to Usercode"
      Top             =   1710
      Width           =   1755
   End
   Begin VB.TextBox txtUser 
      Height          =   345
      Left            =   1860
      TabIndex        =   14
      Tag             =   "5_txtZip:::§~§"
      Text            =   "Aligned to Zip"
      Top             =   1710
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fixed height frame, always on the bottom..."
      Height          =   675
      Left            =   30
      TabIndex        =   11
      Tag             =   ":2:2:§~§"
      Top             =   4050
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   345
         Left            =   3600
         TabIndex        =   13
         Tag             =   "2:::§~§"
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "...where to put action buttons for instance"
         Height          =   225
         Left            =   90
         TabIndex        =   12
         Top             =   300
         Width           =   2985
      End
   End
   Begin VB.TextBox txtDescr 
      Height          =   645
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   10
      Tag             =   ":105_lblHint2:2:1260§~§"
      Text            =   "frmEntry.frx":0000
      Top             =   3300
      Width           =   4695
   End
   Begin VB.TextBox txtState 
      Height          =   345
      Left            =   1500
      TabIndex        =   7
      Tag             =   "45_txtName:::§~§"
      Text            =   "Fixed"
      Top             =   990
      Width           =   495
   End
   Begin VB.TextBox txtCity 
      Height          =   345
      Left            =   3030
      TabIndex        =   2
      Tag             =   "45_txtZip::180:§~§"
      Text            =   "Grow by what remains"
      Top             =   990
      Width           =   1695
   End
   Begin VB.TextBox txtZip 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Tag             =   "15_txtState:::§~§"
      Text            =   "Fixed"
      Top             =   990
      Width           =   945
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Tag             =   "::1:§~§"
      Text            =   "Grow in proportion"
      Top             =   990
      Width           =   1395
   End
   Begin VB.Label lblPriv 
      Caption         =   "Privileges"
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Tag             =   "4_txtPriv:::§~§"
      Top             =   1470
      Width           =   825
   End
   Begin VB.Label lblUser 
      Caption         =   "Usercode"
      Height          =   195
      Left            =   1890
      TabIndex        =   15
      Tag             =   "4_txtUser:::§~§"
      Top             =   1470
      Width           =   825
   End
   Begin VB.Label lblHint2 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmEntry.frx":005C
      Height          =   1065
      Left            =   60
      TabIndex        =   9
      Tag             =   "::2:1§~§"
      Top             =   2130
      Width           =   4665
   End
   Begin VB.Label lblState 
      Caption         =   "State"
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Tag             =   "4_txtState:::§~§"
      Top             =   750
      Width           =   465
   End
   Begin VB.Label lblHint1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Using the Align action easily lets you keep the labels together with their textboxes as they move and/or grow."
      Height          =   465
      Left            =   60
      TabIndex        =   6
      Tag             =   "::2:§~§"
      Top             =   180
      Width           =   4665
   End
   Begin VB.Label lblCity 
      Caption         =   "City"
      Height          =   195
      Left            =   3060
      TabIndex        =   5
      Tag             =   "4_txtCity:::§~§"
      Top             =   750
      Width           =   825
   End
   Begin VB.Label lblZip 
      Caption         =   "Zip Code"
      Height          =   195
      Left            =   2070
      TabIndex        =   4
      Tag             =   "4_txtZip:::§~§"
      Top             =   750
      Width           =   825
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   750
      Width           =   825
   End
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Resizer As New ControlResizer

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Resizer.InitResizer Me, Me.Width, Me.Height
End Sub

Private Sub Form_Resize()
    Resizer.FormResized Me
End Sub
