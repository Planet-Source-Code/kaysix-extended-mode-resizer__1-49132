VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'0=Do nothing
'1=Proportional change
'2=Static change
'3=Limited Static change
'>3=Anchor (K6)
Dim HPercent                As Double
Dim WPercent                As Double
Dim TPercent                As Double
Dim LPercent                As Double
Private FormMinWidth        As Long
Private FormMinHeight       As Long
Private FormOldWidth        As Long 'stores the old width of the owner form
Private FormOldHeight       As Long 'stores the old height of the owner form

Private Inited              As Boolean 'detect if resizer has been initiated
Private IgnoreNextIfNotMax  As Boolean 'ignore next resize if its not a maximise command
Private FreeResize          As Boolean 'allow free form shrink, leaving controls in place

' Record to store the resizing data for a single control (K6)
Private Type ControlResizeData
    ID As String ' index of the control in the .Controls set of the form (K6)
    Name As String
    Left As String
    Top As String
    Width As String
    Height As String
    AnchorLeft As String
    AnchorTop As String
    AnchorWidth As String
    AnchorHeight As String
    AnchorDepth As Integer
End Type

Private Const TAG_SEPARATOR = "§~§" ' to separate resizing data from the original tag (K6)
Private vData() As ControlResizeData ' array of resizing data (K6)
Private collKeys As Collection ' Form ID - vData Index pairs (K6)

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
' LeftOf("1234!-!5678", "A") = "1234!-!5678"
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


Public Sub InitResizer(frm As Form, MinWidth As Long, MinHeight As Long, Optional ByVal UseFreeResize As Boolean = False)
'this method will set the FormOldWidth and FormOldHeight
'and show that the resizer has been initiated
'Parameters:    frm: The owner form object
'               MinWidth: Minimum form width
'               MinHeight: Minimum form height
'               UseFreeResize: Allow form shrinking under minimum sizes (K6)
'Usage:         Simply add to Form_load event eg: frmResize.InitResizer(Me)

    FormOldWidth = frm.Width 'sets the current width of the form
    FormOldHeight = frm.Height 'sets the current height of the form
    FormMinWidth = MinWidth
    FormMinHeight = MinHeight
    
    ' the new parameter, if true, allows shrinking the form below minimum sizes,
    ' but leaving the controls in place (K6)
    FreeResize = UseFreeResize
    
    ' this migrates the resizing data in the array and restores the original tags (K6)
    Call GetResizeData(frm)
    
    Inited = True 'will show that the resizer has been initated
    
End Sub

Public Sub FormResized(frm As Form)
'this method will resize controls accordingly
'when the form is resized
'Parameters:    frm: The owner form object
'Usage:         simply add to form_resize event eg: Call frmResize.FormResized(Me)

    On Error GoTo Err_Handle 'error handler
    
    Dim cControl        As Control 'the variable holding a control from the owner form
    
    Dim WidthChange     As Long 'change in width of the owner form
    Dim HeightChange    As Long 'change in height of the owner form
    Dim WidthPercent    As Long 'percentage width change of the owner form
    Dim HeightPercent   As Long 'percentage height change of the owner form
    
    ' the next 4 variables are used to store the calculated data, to
    ' apply just one .Move to the control being resized, rather than setting
    ' each property individually (K6)
    Dim lCLeft          As Long
    Dim lCTop           As Long
    Dim lCWidth         As Long
    Dim lCHeight        As Long
    Dim lCount As Long ' just a counter (K6)
    
    If Inited = False Then 'detect if InitResizer has been called
        'raise error if it hasnt
        Call Err.Raise(80000, , "Initiate resize module before resizing")
        Exit Sub 'stop further execution of this method to prevent errors
    End If
    
    If frm.WindowState = vbMinimized Then 'check to see if resize is a minimisation
        IgnoreNextIfNotMax = True 'set to ignore resize if its not to a maximise command
        Exit Sub 'stop further execution of this method to prevent errors
    End If
    
    If IgnoreNextIfNotMax = True Then 'detect if set to ignore next non-max command
        IgnoreNextIfNotMax = False 'set to perform resize next time
        If frm.WindowState <> vbMaximized Then 'will cause resize to display wrongly if not maximised
            Exit Sub 'stop further execution of this method to prevent errors
        End If
    End If

    If Not FreeResize Then
        'if new formsize is smaller then the specified minimumsize, then set to the minimum formsize
        If frm.Width < FormMinWidth Then
            frm.Width = FormMinWidth
        End If
        If frm.Height < FormMinHeight Then
            frm.Height = FormMinHeight
        End If
    Else
        If (frm.Width < FormMinWidth) Or (frm.Height < FormMinHeight) Then Exit Sub
    End If
    
    ' early abort if there are no resizing data for the whole form (K6)
    If UBound(vData) = -1 Then Exit Sub
    
    WidthChange = frm.Width - FormOldWidth 'calculate width change
    HeightChange = frm.Height - FormOldHeight 'calculate height change
    
    WidthPercent = (((frm.Width - FormOldWidth) / FormOldWidth) * 100) 'calculate width percent change
    HeightPercent = (((frm.Height - FormOldHeight) / FormOldHeight) * 100) 'calculate height percent change
    
    ' the new cycle now only affects the controls having resizing data in the vData array (K6)
    For lCount = LBound(vData) To UBound(vData)
        With vData(lCount)
            Set cControl = frm.Controls(.ID)
            'change left property
            Select Case .Left 'check left mode
            Case 0
                lCLeft = cControl.Left
            Case Is = 1 'Proportional change
                'change the left of control proportionally
                'cControl.Left = ((100 + WidthPercent) / 100) * cControl.Left
                LPercent = cControl.Left / FormOldWidth
                lCLeft = frm.Width * LPercent
            Case Is = 2 'Static change
                lCLeft = cControl.Left + WidthChange 'add change in width
            Case Is = 3 'Limited static change
                If cControl.Left + WidthChange < 0 Then 'detect if change will reduce below 0
                    lCLeft = 0 'set left to 0
                Else
                    lCLeft = cControl.Left + WidthChange 'add change as normal
                End If
            Case Is = 4 ' Align Left (K6)
                lCLeft = frm.Controls(ValNoEmpty(.AnchorLeft)).Left
            Case Is = 5 ' Align Right (K6)
                With frm.Controls(ValNoEmpty(.AnchorLeft))
                    lCLeft = .Left + .Width - cControl.Width
                End With
            Case Else ' Anchor (K6)
                If .AnchorLeft = "" Then ' anchor to form
                    lCLeft = frm.Width - .Left
                Else ' anchor to another control
                    With frm.Controls(ValNoEmpty(.AnchorLeft))
                        lCLeft = .Left + .Width + vData(lCount).Left
                    End With
                End If
            End Select
            
            'change top property
            Select Case .Top 'check top mode
            Case 0
                lCTop = cControl.Top
            Case Is = 1 'Proportional change
                'change the top of control proportionally
                TPercent = cControl.Top / FormOldHeight
                lCTop = frm.Height * TPercent
                'cControl.Top = ((100 + HeightPercent) / 100) * cControl.Top
            Case Is = 2 'Static change
                lCTop = cControl.Top + HeightChange 'add change in height
            Case Is = 3 'Limited static change
                If cControl.Top + HeightChange < 0 Then 'detect if change will reduce below 0
                    lCTop = 0 'set top to 0
                Else
                    lCTop = cControl.Top + HeightChange 'add change as normal
                End If
            Case Is = 4 ' Align Top (K6)
                lCTop = frm.Controls(ValNoEmpty(.AnchorTop)).Top
            Case Is = 5 ' Align Bottom (K6)
                With frm.Controls(ValNoEmpty(.AnchorTop))
                    lCTop = .Top + .Height - cControl.Height
                End With
            Case Else ' Anchor (K6)
                If .AnchorTop = "" Then ' anchor to form
                    lCTop = frm.Height - .Top
                Else ' anchor to another control
                    With frm.Controls(ValNoEmpty(.AnchorTop))
                        lCTop = .Top + .Height + vData(lCount).Top
                    End With
                End If
            End Select
                
            'change width property
            Select Case .Width 'check width mode
            Case 0
                lCWidth = cControl.Width
            Case Is = 1 'Proportional change
                'change the left of control proportionally
                WPercent = cControl.Width / FormOldWidth
                lCWidth = frm.Width * WPercent
            Case Is = 2 'Static change
                lCWidth = cControl.Width + WidthChange 'add change in width
            Case Is = 3 'Limited static change
                If cControl.Width + WidthChange < 0 Then 'detect if change will reduce below 0
                    lCWidth = 0 'set left to 0
                Else
                    lCWidth = cControl.Width + WidthChange 'add change as normal
                End If
            Case Is = 4 ' Same Width As (K6)
                lCWidth = frm.Controls(ValNoEmpty(.AnchorWidth)).Width
            Case Is = 5 ' Grow Left To (K6)
                With frm.Controls(ValNoEmpty(.AnchorWidth))
                    lCWidth = .Left + .Width - cControl.Left
                End With
            Case Else ' Anchor (K6)
                If .AnchorWidth = "" Then ' anchor to form
                    lCWidth = frm.Width - lCLeft - .Width
                Else ' anchor to another control
                    With frm.Controls(ValNoEmpty(.AnchorWidth))
                        If vData(lCount).Height > 0 Then
                            lCWidth = .Left - vData(lCount).Left
                        Else
                            lCWidth = vData(lCount).Width + .Left - lCLeft
                        End If
                    End With
                End If
            End Select
            
            'change height property
            Select Case .Height 'check height mode
            Case 0
                lCHeight = cControl.Height
            Case Is = 1 'Proportional change
                'change the height of control proportionally
                HPercent = cControl.Height / FormOldHeight
                'cControl.Height = ((100 + HeightPercent) / 100) * cControl.Height
                lCHeight = frm.Height * HPercent
            Case Is = 2 'Static change
                lCHeight = cControl.Height + HeightChange 'add change in height
            Case Is = 3 'Limited static change
                If cControl.Height + HeightChange < 0 Then 'detect if change will reduce below 0
                    lCHeight = 0 'set top to 0
                Else
                    lCHeight = cControl.Height + HeightChange 'add change as normal
                End If
            Case Is = 4 ' Same Height As (K6)
                lCHeight = frm.Controls(ValNoEmpty(.AnchorHeight)).Height
            Case Is = 5 ' Grow Bottom To (K6)
                With frm.Controls(ValNoEmpty(.AnchorHeight))
                    lCHeight = .Top + .Height - cControl.Top
                End With
            Case Else ' Anchor (K6)
                If .AnchorHeight = "" Then ' anchor to form
                    lCHeight = frm.Height - lCTop - .Height
                Else ' anchor to another control
                    With frm.Controls(ValNoEmpty(.AnchorHeight))
                        If vData(lCount).Height > 0 Then
                            lCHeight = .Top - vData(lCount).Top
                        Else
                            lCHeight = vData(lCount).Height + .Top - lCTop
                        End If
                    End With
                End If
            End Select
            
            ' the .Move is faster than setting the properties individually (K6)
            cControl.Move lCLeft, lCTop, lCWidth, lCHeight
            
        End With
        
    Next
    
    FormOldWidth = frm.Width 'change old width to current width
    FormOldHeight = frm.Height 'change old height to current height

    Exit Sub

Err_Handle:
    If Err.Number = 438 Then '"object doesn't support this property or method" will occur on controls like timers
        Err.Clear 'clear error
        Resume Next 'ignore line with error and continue
    ElseIf Err.Number = 11 Then '"division by zero" will occur after restoring from minimize
        Err.Clear 'clear error
        Resume Next 'ignore line with error and continue
    ElseIf Err.Number = 80000 Then
        Exit Sub
    ElseIf Err.Number = 5 Then
        Exit Sub
    Else
        'show error box
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbExclamation + vbOKOnly
    End If
    
End Sub

Private Sub GetResizeData(frm As Form)
'new internal sub (K6)
'It retrieves the resizing data in a dynamic array
'and restores the control Tags originally set by
'the developer at design time
'Parameters:    frm: The owner form object
'Usage:         None, used internally by InitResizer

    Dim cControl    As Control 'the variable holding a control from the owner form
    Dim lCount      As Long ' just a counter
    Dim lPos        As Long ' holder of the separator position in the tag
    Dim collCtrls   As Collection
    Dim vValues ' variant array that holds the resizing data after the Split
    
    ReDim vData(0)
    Set collKeys = New Collection
    Set collCtrls = New Collection
    ' cycling via index so the index can be stored in the .ID field
    For lCount = 0 To frm.Controls.Count - 1
        Set cControl = frm.Controls(lCount)
        lPos = InStr(cControl.Tag, TAG_SEPARATOR)
        If lPos > 0 Then ' the separator is present, so extract the resizing data
            ' the values are separated by :'s
            vValues = Split(Left$(cControl.Tag, lPos - 1), ":")
            With vData(UBound(vData))
                .ID = lCount
                .Name = cControl.Name
                On Error Resume Next
                .Name = .Name & "." & cControl.Index
                On Error GoTo 0
                .Left = ValNoEmpty(LeftOf(vValues(0), "_"))
                .Top = ValNoEmpty(LeftOf(vValues(1), "_"))
                .Width = ValNoEmpty(LeftOf(vValues(2), "_"))
                .Height = ValNoEmpty(LeftOf(vValues(3), "_"))
                .AnchorLeft = RightOf(vValues(0), "_")
                .AnchorTop = RightOf(vValues(1), "_")
                .AnchorWidth = RightOf(vValues(2), "_")
                .AnchorHeight = RightOf(vValues(3), "_")
                ' initial depth, -1 for free controls, 0 for those related to others
                .AnchorDepth = CInt((.Left < 4) And (.Top < 4) And (.Width < 4) And (.Height < 4))
                collKeys.Add UBound(vData), CStr(.ID)
                collCtrls.Add lCount, .Name
            End With
            cControl.Tag = Mid$(cControl.Tag, lPos + 3) ' restore the original Tag
            ReDim Preserve vData(UBound(vData) + 1)
        End If
    Next
    ' replace the anchor control names with their current index in the Form.Controls array
    For lCount = LBound(vData) To UBound(vData)
        With vData(lCount)
            If .AnchorLeft <> "" Then .AnchorLeft = collCtrls(.AnchorLeft)
            If .AnchorTop <> "" Then .AnchorTop = collCtrls(.AnchorTop)
            If .AnchorWidth <> "" Then .AnchorWidth = collCtrls(.AnchorWidth)
            If .AnchorHeight <> "" Then .AnchorHeight = collCtrls(.AnchorHeight)
        End With
    Next
    Do While collCtrls.Count > 0
        collCtrls.Remove 1
    Loop
    Set collCtrls = Nothing
    If UBound(vData) > 0 Then
        'array elements are created in advance, so the last is not used, cut it
        ReDim Preserve vData(UBound(vData) - 1)
        Call SortResizeData
    Else ' there are no resizing data at all, the array is wiped
        Erase vData
    End If
    ' the collection is wiped anyway, it was only needed by GetAnchorDepth
    Do While collKeys.Count > 0
        collKeys.Remove 1
    Loop
    Set collKeys = Nothing
End Sub

Private Sub SortResizeData()
'new internal sub (K6)
' sorts the resizing data array according to the anchor dependencies
'Usage:         None, used internally by GetResizeData
Dim lCount As Long
    ' calculate the anchor depth for each control that needs resizing
    For lCount = LBound(vData) To UBound(vData)
        If vData(lCount).AnchorDepth = 0 Then vData(lCount).AnchorDepth = GetAnchorDepth(vData(lCount))
    Next
    ' rearrange the array in ascending order of anchor depth
    QuickSort vData, LBound(vData), UBound(vData)
End Sub

Private Function GetAnchorDepth(ByRef vNode As ControlResizeData) As Long
'new internal function (K6)
' returns *and sets* the nested anchors count for the specified control
'Parameters:    vNode: The ControlResizeData record to calculate the anchor depth for
'Usage:         None, used internally by SortResizeData
Dim lALeft As Long
Dim lATop As Long
Dim lAWidth As Long
Dim lAHeight As Long
Dim lAMax As Long
    If vNode.AnchorDepth <> 0 Then ' leaf or already calculated node, we're done
        GetAnchorDepth = vNode.AnchorDepth
        Exit Function
    Else ' recurse on any branch that has an anchor...
        ' (the Anchor* property values refer to the Form based index,
        ' so we retrieve the vData Array based index via the collKeys collection)
        If vNode.AnchorLeft <> "" Then lALeft = GetAnchorDepth(vData(collKeys(vNode.AnchorLeft)))
        If vNode.AnchorTop <> "" Then lATop = GetAnchorDepth(vData(collKeys(vNode.AnchorTop)))
        If vNode.AnchorWidth <> "" Then lAWidth = GetAnchorDepth(vData(collKeys(vNode.AnchorWidth)))
        If vNode.AnchorHeight <> "" Then lAHeight = GetAnchorDepth(vData(collKeys(vNode.AnchorHeight)))
        ' ...and return the max
        lAMax = lALeft
        If lAMax < lATop Then lAMax = lATop
        If lAMax < lAWidth Then lAMax = lAWidth
        If lAMax < lAHeight Then lAMax = lAHeight
        GetAnchorDepth = lAMax + 1
    End If
    ' set the current node depth, even inside recursions, to avoid examining
    ' the same control both as parent of other controls and by itself
    vNode.AnchorDepth = lAMax + 1
End Function

Private Function QuickSort(ByRef vData() As ControlResizeData, ByVal Low As Long, ByVal Hi As Long)
' Found in the company codebase, unknown origin, comments by the original author;
' modified to work explicitly on arrays of ControlResizeData records (K6)
' ---------------------------------------------------------
'
' Syntax:     QuickSort TmpArray(), Low, Hi
'
' Parameters:
'     vData - A variant pointing to an array to be sorted.
'       Low - LBounds(vData) low number of elements in the array
'        Hi - UBounds(vData) high number of elements in the array
'
' ---------------------------------------------------------

' ---------------------------------------------------------
' Define local variables
' ---------------------------------------------------------
  Dim lTmpLow As Long
  Dim lTmpHi As Long
  Dim lTmpMid As Long
  Dim vTempVal As ControlResizeData
  Dim vTmpHold As ControlResizeData

' ---------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------
  lTmpLow = Low
  lTmpHi = Hi

' ---------------------------------------------------------
' Leave if there is nothing to sort
' ---------------------------------------------------------
  If Hi <= Low Then Exit Function

' ---------------------------------------------------------
' Find the middle to start comparing values
' ---------------------------------------------------------
  lTmpMid = (Low + Hi) \ 2

' ---------------------------------------------------------
' Move the item in the middle of the array to the
' temporary holding area as a point of reference while
' sorting.  This will change each time we make a recursive
' call to this routine.
' ---------------------------------------------------------
  vTempVal = vData(lTmpMid)

' ---------------------------------------------------------
' Loop until we eventually meet in the middle
' ---------------------------------------------------------
  Do While (lTmpLow <= lTmpHi)

     ' Always process the low end first.  Loop as long
     ' the array data element is less than the data in
     ' the temporary holding area and the temporary low
     ' value is less than the maximum number of array
     ' elements.
     Do While (vData(lTmpLow).AnchorDepth < vTempVal.AnchorDepth And lTmpLow < Hi)
           lTmpLow = lTmpLow + 1
     Loop

     ' Now, we will process the high end.  Loop as long
     ' the data in the temporary holding area is less
     ' than the array data element and the temporary high
     ' value is greater than the minimum number of array
     ' elements.
     Do While (vTempVal.AnchorDepth < vData(lTmpHi).AnchorDepth And lTmpHi > Low)
           lTmpHi = lTmpHi - 1
     Loop

     ' if the temp low end is less than or equal
     ' to the temp high end, then swap places
     If (lTmpLow <= lTmpHi) Then
         vTmpHold = vData(lTmpLow)          ' Move the Low value to Temp Hold
         vData(lTmpLow) = vData(lTmpHi)     ' Move the high value to the low
         vData(lTmpHi) = vTmpHold           ' move the Temp Hod to the High
         lTmpLow = lTmpLow + 1              ' Increment the temp low counter
         lTmpHi = lTmpHi - 1                ' Dcrement the temp high counter
     End If

  Loop

' ---------------------------------------------------------
' If the minimum number of elements in the array is
' less than the temp high end, then make a recursive
' call to this routine.  I always sort the low end
' of the array first.
' ---------------------------------------------------------
  If (Low < lTmpHi) Then
      QuickSort vData, Low, lTmpHi
  End If

' ---------------------------------------------------------
' If the temp low end is less than the maximum number
' of elements in the array, then make a recursive call
' to this routine.  The high end is always sorted last.
' ---------------------------------------------------------
  If (lTmpLow < Hi) Then
       QuickSort vData, lTmpLow, Hi
  End If

End Function
