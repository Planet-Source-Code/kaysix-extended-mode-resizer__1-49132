VERSION 5.00
Begin VB.Form frmAddIn 
   Caption         =   "Resizer Add In"
   ClientHeight    =   6330
   ClientLeft      =   2190
   ClientTop       =   1950
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.Frame fraFormProperties 
      Caption         =   "Form Properties"
      Height          =   855
      Left            =   2400
      TabIndex        =   35
      Tag             =   "0"
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox chkFreeResize 
         Caption         =   "Allow form shrinking below minimum sizes"
         Height          =   495
         Left            =   2490
         TabIndex        =   44
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton cmdAddCode 
         Caption         =   "Add Necessary code to form"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.ComboBox cmbContainers 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Tag             =   "0000"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Resize and Move Properties"
      Height          =   5295
      Left            =   2400
      TabIndex        =   25
      Tag             =   "0011"
      Top             =   960
      Width           =   4815
      Begin VB.CommandButton cmdUndo 
         Caption         =   "UNDO Changes"
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton cmdNoAction 
         Caption         =   "NO Action"
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdResizeProp 
         Caption         =   "Only Resize(prop)"
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdMoveProp 
         Caption         =   "Only Move(prop)"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdOnlyResize 
         Caption         =   "Only Resize(static)"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Tag             =   "1111"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton cmdOnlyMove 
         Caption         =   "Only Move(static)"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Tag             =   "1111"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Frame fraTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Tag             =   "0000"
         Top             =   1560
         Width           =   4575
         Begin VB.ComboBox cmbAnchor 
            Height          =   315
            Index           =   1
            ItemData        =   "frmAddIn.frx":0000
            Left            =   1410
            List            =   "frmAddIn.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   180
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CommandButton cmdCalc 
            Height          =   315
            Index           =   1
            Left            =   4200
            Picture         =   "frmAddIn.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Calculate the current gap"
            Top             =   180
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtTwips 
            Height          =   315
            Index           =   1
            Left            =   3360
            TabIndex        =   9
            Top             =   180
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.ComboBox cmbDo 
            Height          =   315
            Index           =   1
            ItemData        =   "frmAddIn.frx":014E
            Left            =   0
            List            =   "frmAddIn.frx":0167
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label lblAnchor 
            AutoSize        =   -1  'True
            Caption         =   "Anchor to"
            Height          =   195
            Index           =   1
            Left            =   1410
            TabIndex        =   46
            Top             =   -30
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblTwips 
            AutoSize        =   -1  'True
            Caption         =   "Twips"
            Height          =   195
            Index           =   1
            Left            =   3360
            TabIndex        =   37
            Top             =   -30
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Action"
            Height          =   195
            Left            =   30
            TabIndex        =   36
            Top             =   -30
            Width           =   450
         End
      End
      Begin VB.Frame fraLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   120
         TabIndex        =   32
         Tag             =   "0000"
         Top             =   600
         Width           =   4575
         Begin VB.ComboBox cmbAnchor 
            Height          =   315
            Index           =   0
            ItemData        =   "frmAddIn.frx":01B7
            Left            =   1410
            List            =   "frmAddIn.frx":01B9
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   180
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CommandButton cmdCalc 
            Height          =   315
            Index           =   0
            Left            =   4200
            Picture         =   "frmAddIn.frx":01BB
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Calculate the current gap"
            Top             =   180
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtTwips 
            Height          =   315
            Index           =   0
            Left            =   3360
            TabIndex        =   5
            Top             =   180
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.ComboBox cmbDo 
            Height          =   315
            Index           =   0
            ItemData        =   "frmAddIn.frx":0305
            Left            =   0
            List            =   "frmAddIn.frx":031E
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label lblAnchor 
            AutoSize        =   -1  'True
            Caption         =   "Anchor to"
            Height          =   195
            Index           =   0
            Left            =   1410
            TabIndex        =   45
            Top             =   -30
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblTwips 
            AutoSize        =   -1  'True
            Caption         =   "Twips"
            Height          =   195
            Index           =   0
            Left            =   3360
            TabIndex        =   39
            Top             =   -30
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Action"
            Height          =   195
            Left            =   30
            TabIndex        =   38
            Top             =   -30
            Width           =   450
         End
      End
      Begin VB.Frame fraWidth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Tag             =   "0000"
         Top             =   2520
         Width           =   4575
         Begin VB.ComboBox cmbAnchor 
            Height          =   315
            Index           =   2
            ItemData        =   "frmAddIn.frx":036E
            Left            =   1410
            List            =   "frmAddIn.frx":0370
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   180
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CommandButton cmdCalc 
            Height          =   315
            Index           =   2
            Left            =   4200
            Picture         =   "frmAddIn.frx":0372
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Calculate the current gap"
            Top             =   180
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtTwips 
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   13
            Top             =   180
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.ComboBox cmbDo 
            Height          =   315
            Index           =   2
            ItemData        =   "frmAddIn.frx":04BC
            Left            =   0
            List            =   "frmAddIn.frx":04D5
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label lblAnchor 
            AutoSize        =   -1  'True
            Caption         =   "Anchor to"
            Height          =   195
            Index           =   2
            Left            =   1410
            TabIndex        =   47
            Top             =   -30
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblTwips 
            AutoSize        =   -1  'True
            Caption         =   "Twips"
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   41
            Top             =   -30
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Action"
            Height          =   195
            Left            =   30
            TabIndex        =   40
            Top             =   -30
            Width           =   450
         End
      End
      Begin VB.Frame fraHeight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Tag             =   "0000"
         Top             =   3480
         Width           =   4575
         Begin VB.ComboBox cmbAnchor 
            Height          =   315
            Index           =   3
            ItemData        =   "frmAddIn.frx":052A
            Left            =   1410
            List            =   "frmAddIn.frx":052C
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   180
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CommandButton cmdCalc 
            Height          =   315
            Index           =   3
            Left            =   4200
            Picture         =   "frmAddIn.frx":052E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Calculate the current gap"
            Top             =   180
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtTwips 
            Height          =   315
            Index           =   3
            Left            =   3360
            TabIndex        =   17
            Top             =   180
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.ComboBox cmbDo 
            Height          =   315
            Index           =   3
            ItemData        =   "frmAddIn.frx":0678
            Left            =   0
            List            =   "frmAddIn.frx":0691
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label lblAnchor 
            AutoSize        =   -1  'True
            Caption         =   "Anchor to"
            Height          =   195
            Index           =   3
            Left            =   1410
            TabIndex        =   48
            Top             =   -30
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblTwips 
            AutoSize        =   -1  'True
            Caption         =   "Twips"
            Height          =   195
            Index           =   3
            Left            =   3360
            TabIndex        =   43
            Top             =   -30
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Action"
            Height          =   195
            Left            =   30
            TabIndex        =   42
            Top             =   -30
            Width           =   450
         End
      End
      Begin VB.Label lblHeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Height"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Tag             =   "0000"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblWidth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Width"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Tag             =   "0000"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Left"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Tag             =   "0000"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Top"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Tag             =   "0000"
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.ListBox lstControls 
      Height          =   5520
      Left            =   0
      TabIndex        =   1
      Tag             =   "0001"
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox cmbForms 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "0000"
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Resizer As New ControlResizer
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Const CLEFT = 0
Private Const CTOP = 1
Private Const CWIDTH = 2
Private Const CHEIGHT = 3

' used to separate resizing data from the original tag,
' it should be weird enough not to appear in the designed tags,
' but if it does, just change it here *and in the class* (K6)
Private Const TAG_SEPARATOR = "§~§"

' New variables (K6)
Dim szUserTag   As String ' used to preserve the Tag originally set up at design time
Dim bForcing    As Boolean ' used to prevent Tag update when cleaning data controls
Dim lFrmWidth   As Long ' design time width of the selected form
Dim lFrmHeight  As Long ' design time height of the selected form
Dim lCtlLeft    As Long ' design time x offset of the selected component
Dim lCtlTop     As Long ' design time y offset of the selected component
Dim lCtlWidth   As Long ' design time width of the selected component
Dim lCtlHeight  As Long ' design time height of the selected component
Dim lCtlTag     As String ' current tag of the selected component for Undo purposes

Private Function ValNoEmpty(ByVal vValue As String) As Long
' a shortcut function to treat empty strings as 0 = no action (K6)
    If IsNumeric(vValue) Then
        ValNoEmpty = CLng(vValue)
    Else
        ValNoEmpty = 0
    End If
End Function

Private Function LeftOf(ByVal aszText As String, ByVal aszSeparator As String) As String
' Returns the leftmost part of aszText before aszSeparator (K6)
' LeftOf("1234!-!5678", "!-!") = "1234"
' LeftOf("1234!-!5678", "A") = "1234!5678"
Dim llPos As Long
    LeftOf = ""
    If aszText = "" Then Exit Function
    llPos = InStr(aszText, aszSeparator)
    If llPos <> 0 Then
        LeftOf = Left$(aszText, llPos - 1)
    Else
        LeftOf = aszText
    End If
End Function

Private Function RightOf(ByVal aszText As String, ByVal aszSeparator As String) As String
' Returns the rightmost part of aszText after aszSeparator (K6)
' RightOf("1234!-!5678", "!-!") = "5678"
' RightOf("1234!-!5678", "A") = ""
Dim llPos As Long
    RightOf = ""
    If aszText = "" Then Exit Function
    llPos = InStr(aszText, aszSeparator)
    If llPos <> 0 Then
        RightOf = Mid$(aszText, llPos + Len(aszSeparator))
    Else
        RightOf = ""
    End If
End Function

Private Function IsSameControl(ByRef ctrl As VBControl, ByVal szName As String) As Boolean
' checks that the szName string matches Name and eventually Index of the
' specified control, as in "txtName" or "txtDate(4)" (K6)
Dim szIndex As String
    IsSameControl = False
    szIndex = RightOf(szName, "(")
    If szIndex = "" Then
        IsSameControl = (ctrl.Properties("Name") = szName)
    Else
        szIndex = Left$(szIndex, Len(szIndex) - 1)
        szName = LeftOf(szName, "(")
        IsSameControl = (ctrl.Properties("Name") = szName) And (ctrl.Properties("Index") = szIndex)
    End If
End Function

Private Function NameOf(ByRef ctrl As VBControl) As String
' returns the composite name of the control, Index included if present (K6)
    NameOf = ctrl.Properties("Name")
    If ctrl.Properties("Index") <> "-1" Then NameOf = NameOf & "(" & ctrl.Properties("Index") & ")"
End Function

Private Function NameReformat(ByVal szName As String, ByVal bUseDots As Boolean) As String
' changes an indexed name like Frame1(0) into Frame1.0 and vice-versa
' non-indexed names are unchanged (K6)
    If bUseDots Then
        NameReformat = Replace(LeftOf(szName, ")"), "(", ".")
    Else
        If InStr(szName, ".") > 0 Then
            NameReformat = Replace(szName, ".", "(") & ")"
        Else
            NameReformat = szName
        End If
    End If
End Function

Private Function CtrlByName(ByVal szName As String) As VBControl
' retrieves the index in the form of the current control named szName (K6)
Dim comp As VBComponent
Dim ctrl As VBControl
Dim ctrl2 As VBControl
Dim vbf As VBForm
Dim lCount As Long
    Set CtrlByName = Nothing
    If cmbForms.ListIndex = -1 Then
        Exit Function
    End If
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbForms.Text Then Exit For
    Next
    Set vbf = comp.Designer
    For lCount = 1 To vbf.VBControls.Count
        Set ctrl = vbf.VBControls(lCount)
        If NameOf(ctrl) = szName Then
            Set CtrlByName = ctrl
            Exit Function
        End If
    Next
End Function

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub GetAllForms()
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
    cmbForms.Clear
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If (comp.Type = vbext_ct_MSForm Or _
            comp.Type = vbext_ct_UserControl _
              Or comp.Type = vbext_ct_VBForm Or _
             comp.Type = vbext_ct_VBMDIForm) Then
            cmbForms.AddItem comp.Name
        End If
    Next
    If cmbForms.ListCount <> 0 Then
        cmbForms.ListIndex = 0
    End If
End Sub

Private Sub GetContainers()
' modified to include the Index value if present (K6)
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
    If cmbForms.ListIndex = -1 Then
        Exit Sub
    End If
    cmbContainers.Clear
    cmbContainers.AddItem cmbForms.Text
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbForms.Text Then
            Set vbf = comp.Designer
            For Each ctrl In vbf.VBControls
                If ctrl.ClassName = "Frame" Or ctrl.ClassName = "PictureBox" Then
                    cmbContainers.AddItem NameOf(ctrl)
                End If
            Next
        End If
    Next
    cmbContainers.ListIndex = 0
End Sub

Private Sub GetFormControls()
' modified to include the Index value if present (K6)
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
Dim cmb As ComboBox
    If cmbForms.ListIndex = -1 Then
        Exit Sub
    End If
    lstControls.Clear
    lstControls.AddItem cmbContainers.Text
    For Each cmb In cmbAnchor
        cmb.Clear
        cmb.AddItem cmbContainers.Text
    Next
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbContainers.Text Then
            Set vbf = comp.Designer
            For Each ctrl In vbf.ContainedVBControls
                lstControls.AddItem NameOf(ctrl)
                For Each cmb In cmbAnchor
                    cmb.AddItem NameOf(ctrl)
                Next
            Next
        End If
    Next
End Sub

Private Sub GetControlControls()
' modified to include the Index value if present (K6)
Dim comp As VBComponent
Dim ctrl As VBControl
Dim ctrl2 As VBControl
Dim ctrl3 As VBControl
Dim vbf As VBForm
Dim cmb As ComboBox
    If cmbForms.ListIndex = -1 Then
        Exit Sub
    End If
    lstControls.Clear
    lstControls.AddItem cmbContainers.Text
    For Each cmb In cmbAnchor
        cmb.Clear
        cmb.AddItem cmbContainers.Text
    Next
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbForms.Text Then
            Set vbf = comp.Designer
            For Each ctrl2 In vbf.VBControls
                If IsSameControl(ctrl2, cmbContainers.Text) Then
                    For Each ctrl3 In ctrl2.ContainedVBControls
                        lstControls.AddItem NameOf(ctrl3)
                        For Each cmb In cmbAnchor
                            cmb.AddItem NameOf(ctrl3)
                        Next
                    Next
                    Exit Sub
                End If
            Next
        End If
    Next
End Sub

Private Sub cmbAnchor_Click(Index As Integer)
    If Not bForcing Then Call ApplyResizeTag
End Sub

Private Sub cmbContainers_Click()
    If cmbContainers.ListIndex = 0 Then
        GetFormControls
    Else
        GetControlControls
    End If
End Sub

Private Sub cmbForms_Click()
    VBInstance.ActiveVBProject.VBComponents(cmbForms.Text).Activate
    GetContainers
    GetFormControls
    ' The current form size is retrieved to allow auto calculation
    'of the Anchor values (K6)
    With VBInstance.ActiveVBProject.VBComponents(cmbForms.Text)
        lFrmWidth = .Properties("Width").Value
        lFrmHeight = .Properties("Height").Value
    End With
End Sub

Private Sub cmbDo_Click(Index As Integer)
' show the eventually needed extra controls for the selected action (K6)
    lblAnchor(Index).Visible = (cmbDo(Index).ListIndex > 3)
    cmbAnchor(Index).Visible = lblAnchor(Index).Visible
    txtTwips(Index).Visible = (cmbDo(Index).ListIndex = cmbDo(Index).ListCount - 1)
    lblTwips(Index).Visible = txtTwips(Index).Visible
    cmdCalc(Index).Visible = txtTwips(Index).Visible
    If Not bForcing Then Call ApplyResizeTag
End Sub

Private Sub cmdCalc_Click(Index As Integer)
' calculate the gap according to the present design state (K6)
Dim ctrl As VBControl
    Select Case Index
    Case CLEFT
        Select Case cmbAnchor(CLEFT).ListIndex
        Case 0, -1 ' anchored to the form
            txtTwips(CLEFT).Text = lFrmWidth - lCtlLeft
        Case Else ' anchored to a control
            Set ctrl = CtrlByName(cmbAnchor(CLEFT).Text)
            txtTwips(CLEFT).Text = lCtlLeft - ctrl.Properties("Left").Value - ctrl.Properties("Width").Value
        End Select
    Case CTOP
        Select Case cmbAnchor(CTOP).ListIndex
        Case 0, -1 ' anchored to the form
            txtTwips(CTOP).Text = lFrmHeight - lCtlTop
        Case Else ' anchored to a control
            Set ctrl = CtrlByName(cmbAnchor(CTOP).Text)
            txtTwips(CTOP).Text = lCtlTop - ctrl.Properties("Top").Value - ctrl.Properties("Height").Value
        End Select
    Case CWIDTH
        Select Case cmbAnchor(CWIDTH).ListIndex
        Case 0, -1 ' anchored to the form
            txtTwips(CWIDTH).Text = lFrmWidth - lCtlLeft - lCtlWidth
        Case Else ' anchored to a control
            Set ctrl = CtrlByName(cmbAnchor(CWIDTH).Text)
            txtTwips(CWIDTH).Text = lCtlLeft + lCtlWidth - ctrl.Properties("Left").Value
        End Select
    Case CHEIGHT
        Select Case cmbAnchor(CHEIGHT).ListIndex
        Case 0, -1 ' anchored to the form
            txtTwips(CHEIGHT).Text = lFrmHeight - lCtlTop - lCtlHeight
        Case Else ' anchored to a control
            Set ctrl = CtrlByName(cmbAnchor(CHEIGHT).Text)
            txtTwips(CHEIGHT).Text = lCtlTop + lCtlHeight - ctrl.Properties("Top").Value
        End Select
    End Select
    Call ApplyResizeTag
End Sub

Private Sub cmdMoveProp_Click()
    cmbDo(CTOP).ListIndex = 1
    cmbDo(CLEFT).ListIndex = 1
    cmbDo(CWIDTH).ListIndex = 0
    cmbDo(CHEIGHT).ListIndex = 0
    Call ApplyResizeTag
End Sub

Private Sub cmdNoAction_Click()
    cmbDo(CTOP).ListIndex = -1
    cmbDo(CLEFT).ListIndex = -1
    cmbDo(CWIDTH).ListIndex = -1
    cmbDo(CHEIGHT).ListIndex = -1
    Call ApplyResizeTag
End Sub

Private Sub cmdOnlyMove_Click()
    cmbDo(CTOP).ListIndex = 2
    cmbDo(CLEFT).ListIndex = 2
    cmbDo(CWIDTH).ListIndex = 0
    cmbDo(CHEIGHT).ListIndex = 0
    Call ApplyResizeTag
End Sub

Private Sub cmdOnlyResize_Click()
    cmbDo(CTOP).ListIndex = 0
    cmbDo(CLEFT).ListIndex = 0
    cmbDo(CWIDTH).ListIndex = 2
    cmbDo(CHEIGHT).ListIndex = 2
    Call ApplyResizeTag
End Sub

Private Sub cmdAddCode_Click()
    AddCode
End Sub

Private Sub cmdResizeProp_Click()
    cmbDo(CTOP).ListIndex = 0
    cmbDo(CLEFT).ListIndex = 0
    cmbDo(CWIDTH).ListIndex = 1
    cmbDo(CHEIGHT).ListIndex = 1
    Call ApplyResizeTag
End Sub

Private Sub cmdUndo_Click()
    Call ApplyResizeTag(lCtlTag)
    Call ShowCurrentState(lCtlTag)
End Sub

Private Sub Form_Load()
    Resizer.InitResizer Me, Me.Width, Me.Height
    AddClass
    GetAllForms
    GetFormControls
    GetContainers
End Sub

Private Function ControlTag() As String
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
    If cmbForms.ListIndex = -1 Then
        Exit Function
    End If
    If cmbContainers.ListIndex = 0 Then
        If lstControls.ListIndex = 0 Then
            For Each comp In VBInstance.ActiveVBProject.VBComponents
                If comp.Name = cmbForms.Text Then
                    ControlTag = comp.Properties("Tag")
                    Exit Function
                End If
            Next
        Else
            GoTo NormalControl
        End If
    Else
NormalControl:
        For Each comp In VBInstance.ActiveVBProject.VBComponents
            If comp.Name = cmbForms.Text Then
                Set vbf = comp.Designer
                For Each ctrl In vbf.VBControls
                    If IsSameControl(ctrl, lstControls.Text) Then
                        ControlTag = ctrl.Properties("Tag")
                        ' retrieving current control data for Anchor calculation (K6)
                        lCtlLeft = ctrl.Properties("Left")
                        lCtlTop = ctrl.Properties("Top")
                        lCtlWidth = ctrl.Properties("Width")
                        lCtlHeight = ctrl.Properties("Height")
                        lCtlTag = ctrl.Properties("Tag")
                        Exit Function
                    End If
                Next
            End If
        Next
    End If
End Function

Private Sub lstControls_Click()
    If lstControls.ListIndex = 0 And cmbContainers.ListIndex = 0 Then
        fraFormProperties.Enabled = True
    Else
        fraFormProperties.Enabled = False
    End If
    Call ShowCurrentState(ControlTag)
End Sub

Private Sub ShowCurrentState(szStr As String)
' modified to handle the new extended syntax and the new controls (K6)
' in the line marked [1], the comparison value of 15 is a sort of borderline
' between mode id and amount of twips; if it's below 15 it directly represents
' the operation mode; if 15 or above, it's the amount of twips for mode Anchor
Dim i As Integer, ResizeString  As String, lVal As Long, vValues
    If InStr(szStr, TAG_SEPARATOR) > 0 Then
        szUserTag = RightOf(szStr, TAG_SEPARATOR)
        ResizeString = LeftOf(szStr, TAG_SEPARATOR)
    Else
        szUserTag = szStr
        ResizeString = ""
    End If
    
    bForcing = True
    For i = CLEFT To CHEIGHT
        txtTwips(i).Text = ""
        cmbDo(i).ListIndex = -1
        cmbAnchor(i).ListIndex = -1
    Next
    
    If ResizeString = "" Then
        bForcing = False
        Exit Sub
    End If
    
    vValues = Split(ResizeString, ":")
    For i = CLEFT To CHEIGHT
    If IsNumeric(vValues(i)) And (vValues(i) < 4) Then ' original modes by the previous authors
        cmbDo(i).ListIndex = CInt(vValues(i))
    ElseIf vValues(i) <> "" Then
        lVal = ValNoEmpty(LeftOf(vValues(i), "_"))
        If lVal < 15 Then ' [1] new extended cases, Align/Same size (K6)
            txtTwips(i).Text = ""
            cmbDo(i).ListIndex = lVal
        Else ' new Anchor case (K6)
            txtTwips(i).Text = lVal
            cmbDo(i).ListIndex = 6
        End If
        If Not IsNumeric(vValues(i)) Then
            On Error GoTo Anchor_Err
            cmbAnchor(i).Text = NameReformat(RightOf(vValues(i), "_"), False)
        Else
            cmbAnchor(i).ListIndex = -1
        End If
    End If
    Next
    bForcing = False
    Exit Sub
Anchor_Err:
    Select Case Err.Number
    Case 383
        MsgBox NameReformat(RightOf(vValues(i), "_"), False) & " : invalid anchor control name." _
            & vbLf & "Choose it again manually.", vbInformation, "Alert"
        cmbAnchor(i).SetFocus
    Case Else
        MsgBox "Error " & Err.Number & " : " & Err.Description, vbInformation, "Alert"
    End Select
    Err.Clear
End Sub

Private Sub ApplyResizeTag(Optional ByVal szForceTag As String = "")
' modified to force the eventually specified tag used by UNDO (K6)
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
    If cmbForms.ListIndex = -1 Then
        Exit Sub
    End If
    If cmbContainers.ListIndex = 0 Then
        If lstControls.ListIndex = 0 Then
            For Each comp In VBInstance.ActiveVBProject.VBComponents
                If comp.Name = cmbForms.Text Then
                    Set vbf = comp.Designer
                    comp.Properties("Tag") = ResizeTag
                    Exit Sub
                End If
            Next
            Exit Sub
        End If
    End If
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbForms.Text Then
            Set vbf = comp.Designer
            For Each ctrl In vbf.VBControls
                If IsSameControl(ctrl, lstControls.Text) Then
                    ctrl.Properties("Tag") = IIf(szForceTag = "", ResizeTag, szForceTag)
                    Exit Sub
                End If
            Next
        End If
    Next
End Sub

Private Function ResizeTag() As String
' modified to compose the tag according to the new extended syntax (K6)
Dim i As Integer
Dim ResizeString As String
    ResizeString = ""
    For i = CLEFT To CHEIGHT
        ResizeString = ResizeString & ":"
        Select Case cmbDo(i).ListIndex
        Case -1
        Case 0 To 3
            ResizeString = ResizeString & cmbDo(i).ListIndex
        Case cmbDo(i).ListCount - 1
            ' the Anchor case is treated immediately as its index may vary between
            ' combos if you add new modes but always keeping Anchor as the last (K6)
            ResizeString = ResizeString & txtTwips(i)
            If cmbAnchor(i).ListIndex > 0 Then
                ResizeString = ResizeString & "_" & NameReformat(cmbAnchor(i).Text, True)
            End If
        Case 4, 5
            ' then the Align/Same size cases are treated (K6)
            ResizeString = ResizeString & CStr(cmbDo(i).ListIndex)
            If cmbAnchor(i).ListIndex > 0 Then
                ResizeString = ResizeString & "_" & NameReformat(cmbAnchor(i).Text, True)
            End If
        End Select
    Next
    ResizeString = Mid$(ResizeString, 2)
    If ResizeString <> ":::" Then
        ResizeString = ResizeString & TAG_SEPARATOR
    Else
        ResizeString = ""
    End If
    If szUserTag <> "" Then ResizeString = ResizeString & szUserTag
    ResizeTag = ResizeString
End Function

Private Sub AddCode()
' added the handling of the checkbox chkFreeResize that allows free form shrinking (K6)
Dim comp As VBComponent
Dim ctrl As VBControl
Dim ctrl2 As VBControl
Dim vbf As VBForm
    If cmbForms.ListIndex = -1 Then
        Exit Sub
    End If
    lstControls.Clear
    lstControls.AddItem cmbContainers.Text
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbContainers.Text Then
            Set vbf = comp.Designer
            If comp.CodeModule.Find("Dim Resizer As New ControlResizer", 1, 1, comp.CodeModule.CountOfDeclarationLines, -1, True, True) = False Then
                comp.CodeModule.InsertLines 1, "Private Resizer As New ControlResizer"
            End If
            If comp.CodeModule.Find("Sub Form_Load", 1, 1, -1, -1) = False Then
                comp.CodeModule.CreateEventProc "Load", "Form"
            End If
            If comp.CodeModule.Find("Resizer.InitResizer Me,Me.Width,Me.Height", comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc), 1, comp.CodeModule.ProcCountLines("Form_Load", vbext_pk_Proc) + comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc), -1) = False Then
                comp.CodeModule.InsertLines comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc) + 1, "Resizer.InitResizer Me,Me.Width,Me.Height" & IIf(chkFreeResize.Value = vbChecked, ", True", "")
                If comp.Properties("MDIChild") = True Then
                    comp.CodeModule.InsertLines comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc) + 1, "Me.Height=" & comp.Properties("Height") & vbCrLf & "Me.Width=" & comp.Properties("Width")
                End If
            End If
            If comp.CodeModule.Find("Sub Form_Resize", 1, 1, -1, -1) = False Then
                comp.CodeModule.CreateEventProc "Resize", "Form"
            End If
            If comp.CodeModule.Find("Resizer.InitResizer Me,Me.Width,Me.Height", comp.CodeModule.ProcBodyLine("Form_Resize", vbext_pk_Proc), 1, comp.CodeModule.ProcCountLines("Form_Resize", vbext_pk_Proc) + comp.CodeModule.ProcBodyLine("Form_Resize", vbext_pk_Proc), -1) = False Then
                comp.CodeModule.InsertLines comp.CodeModule.ProcBodyLine("Form_Resize", vbext_pk_Proc) + 1, "Resizer.FormResized Me"
            End If
        End If
    Next
    cmbContainers_Click
    Me.Show
End Sub

Private Sub AddClass()
Dim comp As VBComponent
Dim AlreadyExists As Boolean
    FileCopy App.Path & "\ControlResizer.cls", ParsePath(VBInstance.ActiveVBProject.FileName, vbDirectory) & "\ControlResizer.cls"
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = "ControlResizer" Then
            AlreadyExists = True
            Exit For
        End If
    Next
    If Not AlreadyExists Then
        VBInstance.ActiveVBProject.VBComponents.AddFile (App.Path & "\ControlResizer.cls")
    End If
End Sub

Public Function ParsePath(strFullPathName As String, ReturnType As Integer, Optional StripLastBackslash) As String
    Dim strTemp As String, intX As Integer, strPathName As String, strFileName As String

    If IsMissing(StripLastBackslash) Then StripLastBackslash = False
    If Len(strFullPathName) > 0 Then
        strTemp = ""
        intX = Len(strFullPathName)
        Do While strTemp <> "\"
            strTemp = Mid(strFullPathName, intX, 1)
            If strTemp = "\" Then
                strPathName = Left(strFullPathName, intX + StripLastBackslash)
                strFileName = Right(strFullPathName, Len(strFullPathName) - intX)
            End If
            intX = intX - 1
        Loop

        Select Case ReturnType
        Case vbDirectory
            ParsePath = strPathName
        Case vbNormal
            ParsePath = strFileName
        Case Else
            ParsePath = strFullPathName
        End Select
    Else
        ParsePath = ""
    End If

End Function

Private Sub txtTwips_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
' added to dynamically update the tag as the user types on (K6)
    Call ApplyResizeTag
End Sub
