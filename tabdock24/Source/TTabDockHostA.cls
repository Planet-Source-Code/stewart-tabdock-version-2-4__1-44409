VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TTabDockHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************************
' Project     :  TabDock
' Module      :  TTabDockHost.cls
' Description :  Host for Docked Forms
' Created by  :  Marclei V Silva
' Machine     :  ZEUS
' Date-Time   :  12/05/2000 0:40:08
' ********************************************************************
Option Explicit

' Keep up with the errors
Const g_ErrConstant As Long = vbObjectError + 1000
Const m_constClassName = "TTabDockHost"

Private m_lngErrNum As Long
Private m_strErrStr As String
Private m_strErrSource As String

' cursor enumeration from the resource file
Private Enum EnumCursors
    tdSizeWE = 103
    tdSizeNS = 102
End Enum

' Splitters Size/Height
Private Const SPLITTER_HEIGHT = 80
Private Const SPLITTER_WIDTH = 80

' Minimum allowed sizes for any PictureBox.
Private Const MIN_HORIZONTAL = 50
Private Const MIN_VERTICAL = 50

' Drag directions.
Private Const DRAG_NONE = 0
Private Const DRAG_HORIZONTAL = 1
Private Const DRAG_VERTICAL = 2
Private Const DRAG_BOTH = 3

Private m_bSplitting As Boolean     ' Flag that indicates the
Private m_bDirty As Boolean         ' Flag that indicates the
                                    ' docked forms must be resized
Private DragDir As Integer          ' dragging direction
Private cHS As cSplitDDC            ' horizontal and
Private cVS As cSplitDDC            ' vertical splitters
Private deltaY As Integer           ' X offset when splitting
Private deltaX As Integer           ' Y offset when splitting
Private m_FormIndex As Integer      ' Index of the form over form drag area
Private m_WindowList As TDockForms  ' forms collection
Private m_Index As Integer
Private m_Resizable As Integer
Private m_ParentPtr As Long
Private m_Align As tdAlignProperty

Private WithEvents Extender As VB.PictureBox ' Used as the host
Attribute Extender.VB_VarHelpID = -1

Public Property Get Object() As TTabDockHost
    Set Object = Me
End Property

Public Property Get Resizable() As Boolean
    Resizable = m_Resizable
End Property

Public Property Let Resizable(New_Resizable As Boolean)
    m_Resizable = New_Resizable
End Property

Public Property Get Index() As Integer
    Index = m_Index
End Property

Friend Property Let Index(New_Index As Integer)
    m_Index = New_Index
End Property

Public Property Get Align() As tdAlignProperty
    Align = m_Align
End Property

Public Property Let Align(New_Align As tdAlignProperty)
    Extender.Align = New_Align
    m_Align = New_Align
End Property

Public Property Get WindowList() As TDockForms
    Set WindowList = m_WindowList
End Property

Public Property Set Container(NewValue As Object)
    Set Extender = NewValue
End Property

Public Property Get Container() As Object
    Set Container = Extender
End Property

Public Property Get Parent() As TTabDock
    Set Parent = ObjectFromPtr(m_ParentPtr)
End Property

Public Property Set Parent(NewValue As TTabDock)
    m_ParentPtr = PtrFromObject(NewValue)
End Property

Public Property Let Visible(ByVal vData As Boolean)
    On Error Resume Next
'    LockWindowUpdate Parent.Parent.hWnd
    Extender.Visible = vData
    If vData = True Then
        If IsVisible = False Then
            Extender.Visible = False
        End If
    End If
'    LockWindowUpdate ByVal 0&
End Property

Public Property Get Visible() As Boolean
    Visible = Extender.Visible
End Property

Public Property Let Top(ByVal vData As Variant)
On Error Resume Next
    Extender.Top = vData
End Property

Public Property Get Top() As Variant
    Top = Extender.Top
End Property

Public Property Let Left(ByVal vData As Variant)
    Extender.Left = vData
End Property

Public Property Get Left() As Variant
    Left = Extender.Left
End Property

Public Property Let width(ByVal vData As Variant)
    Extender.width = vData
End Property

Public Property Get width() As Variant
    width = Extender.width
End Property

Public Property Let Height(ByVal vData As Variant)
    Extender.Height = vData
End Property

Public Property Get Height() As Variant
    Height = Extender.Height
End Property

Public Property Get hwnd() As Long
    hwnd = Extender.hwnd
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Extender.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Extender.BackColor() = New_BackColor
End Property

Public Property Get BorderStyle() As Integer
    BorderStyle = Extender.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Extender.BorderStyle() = New_BorderStyle
End Property


' ******************************************************************************
' Routine       : (Sub) Class_Initialize
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:35:31
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Initialize the class
' ******************************************************************************
Private Sub Class_Initialize()
On Error GoTo Err_Class_Initialize
    Const constSource As String = m_constClassName & ".Class_Initialize"
    
    ' initially it is resizable
    m_Resizable = True
    ' configure splitters
    Set cHS = New cSplitDDC
    cHS.Orientation = espVertical
    Set cVS = New cSplitDDC
    cVS.Orientation = espHorizontal
    Set m_WindowList = New TDockForms

Exit Sub
Err_Class_Initialize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Class_Terminate
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:35:21
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Terminate the class
' ******************************************************************************
Private Sub Class_Terminate()
On Error GoTo Err_Class_Terminate
    Const constSource As String = m_constClassName & ".Class_Terminate"
    
    Set cHS = Nothing
    Set cVS = Nothing
    Set m_WindowList = Nothing

Exit Sub
Err_Class_Terminate:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Extender_MouseDown
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:34:16
' Inputs        :
' Outputs       :
' Modifications :
'
' Description   : The Extender (picturebox) mouse down event is used
'                 to perform splittings
' ******************************************************************************
Private Sub Extender_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo Err_Extender_MouseDown
    Const constSource As String = m_constClassName & ".Extender_MouseDown"
    
    Dim pt As POINTAPI
    Dim Rc As RECT
    Dim rcPanel As RECT
    Dim hWndA As Long
    
    ' if it is not resizable then exit
    If m_Resizable = False Then Exit Sub
    ' get the form handle that will receive mouse events
    hWndA = Parent.Parent.hwnd
    ' activate flag for sizing purposes
    m_bSplitting = True
    ' If we have vertical splitting then
    ' get the SplitForm rect else get the main form rect
    If DragDir = DRAG_VERTICAL Then
        'debug.Print "mdown dragvert"
        If Align = tdAlignTop Or Align = tdAlignBottom Then
            OverFormDragArea x, Y, m_FormIndex
            GetFormRect m_FormIndex, Rc
        
'            GetClipRect Rc
        Else
            OverFormDragArea x, Y, m_FormIndex
            GetFormRect m_FormIndex, Rc
        End If
    Else
        GetClipRect Rc
    End If
    ' Let's see what to split based on the dragdir var
    If DragDir = DRAG_HORIZONTAL Then
        cHS.SplitterMouseDown hWndA, Rc, x, Y
    ElseIf DragDir = DRAG_VERTICAL Then
        'debug.Print "extender mousedown vertical Split"
        If Align = tdAlignBottom Or Align = tdAlignTop Then
            cVS.SplitterMouseDown hWndA, Rc, x, Y
            deltaY = cVS.Offset
        Else
            cVS.SplitterMouseDown hWndA, Rc, x, Y
            deltaY = cVS.Offset
        End If
    End If
    ' lock updates
    LockWindowUpdate Extender.Parent.hwnd
    ' check drag direction
    If DragDir = DRAG_HORIZONTAL Then
        'debug.Print "extender mousedown"
        If Extender.Align = tdAlignLeft Or Extender.Align = tdAlignRight Then
            If Extender.Align = tdAlignLeft Then
                Extender.width = Extender.width + cHS.Offset
            Else
                Extender.width = Extender.width - cHS.Offset
            End If
        ElseIf Extender.Align = tdAlignTop Or Extender.Align = tdAlignBottom Then
            Extender.Visible = False
            If Extender.Align = tdAlignTop Then
                Extender.Top = 1000
                Extender.Height = Extender.Height + cHS.Offset
            ElseIf Extender.Align = tdAlignBottom Then
                Extender.Top = 0
                Extender.Height = Extender.Height - cHS.Offset
            End If
            Extender.Visible = True
            'm_bDirty = True
        End If
    End If
    ' repaint the host
    Extender_Paint
    ' activate flag for sizing purposes
    m_bSplitting = False
    ' reset updating
    LockWindowUpdate ByVal 0&
    ' raise sizing event
    Parent.TriggerEvent "ResizePanel", Object

Exit Sub
Err_Extender_MouseDown:
    ' reset updating
    LockWindowUpdate ByVal 0&
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub GetClipRect(Rc As RECT)
    Dim rcPanel As RECT
    
    GetWindowRect Extender.hwnd, Rc
    ' restrict resizing over other Panels
    'debug.Print m_Align
    If m_Align = tdAlignRight Then
        Parent.Panels(tdAlignLeft).GetHostRect rcPanel
        Rc.Left = rcPanel.Right + MIN_HORIZONTAL
        Rc.Right = Rc.Right - MIN_HORIZONTAL
    ElseIf m_Align = tdAlignLeft Then
        Parent.Panels(tdAlignRight).GetHostRect rcPanel
        Rc.Right = rcPanel.Left - MIN_HORIZONTAL
        Rc.Left = Rc.Left + MIN_HORIZONTAL
    ElseIf m_Align = tdAlignTop Then
        Parent.Panels(tdAlignBottom).GetHostRect rcPanel
        Rc.Bottom = rcPanel.Top - MIN_VERTICAL
        Rc.Top = Rc.Top + MIN_VERTICAL
    ElseIf m_Align = tdAlignBottom Then
        Parent.Panels(tdAlignTop).GetHostRect rcPanel
        Rc.Top = rcPanel.Bottom + MIN_VERTICAL
        Rc.Bottom = Rc.Bottom - MIN_VERTICAL
    End If
End Sub

Private Sub ClipHostRect(ByRef Rc As RECT)
    Dim rcPanel As RECT
    Dim Offset As Long
    
    ' restrict resizing over other Panels
    If Align = tdAlignRight Then
        Parent.Panels(tdAlignLeft).GetHostRect rcPanel
        Rc.Left = rcPanel.Right + MIN_HORIZONTAL
    ElseIf Align = tdAlignLeft Then
        Parent.Panels(tdAlignRight).GetHostRect rcPanel
        Rc.Right = rcPanel.Left - MIN_HORIZONTAL
    ElseIf Align = tdAlignTop Then
        Parent.Panels(tdAlignBottom).GetHostRect rcPanel
        Rc.Bottom = rcPanel.Top - MIN_VERTICAL
    ElseIf Align = tdAlignBottom Then
        Parent.Panels(tdAlignTop).GetHostRect rcPanel
        Rc.Top = rcPanel.Bottom + MIN_VERTICAL
    End If
End Sub

' ******************************************************************************
' Routine       : (Sub) Extender_MouseMove
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:33:18
' Inputs        :
' Outputs       :
' Modifications :
'
' Description   : While moving over the Extender check for mouse cursor
'                 and direction
' ******************************************************************************
Private Sub Extender_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo Err_Extender_MouseMove
    Const constSource As String = m_constClassName & ".Extender_MouseMove"
    
    ' if it is not resisable exit
    If m_Resizable = False Then Exit Sub
    
    Dim over As Integer
    ' Set the correct arrow type.
    over = OverDragArea(x, Y)
    Select Case over
        Case DRAG_HORIZONTAL
            If Align = tdAlignLeft Or _
                Align = tdAlignRight Then
                Extender.MousePointer = vbCustom
                Extender.MouseIcon = LoadResPicture(tdSizeWE, vbResCursor)
                cHS.Orientation = espVertical
                'debug.Print "ext mdown leftright draghoriz chsorivert"
            ElseIf Align = tdAlignTop Or _
                Align = tdAlignBottom Then
                Extender.MousePointer = vbCustom
                Extender.MouseIcon = LoadResPicture(tdSizeNS, vbResCursor)
                cHS.Orientation = espHorizontal
                'debug.Print "ext mdown topbottom draghoriz chsorihoriz"
            End If
        Case DRAG_VERTICAL
            'debug.Print "vert cur"
            If Align = tdAlignLeft Or _
                Align = tdAlignRight Then
                Extender.MousePointer = vbCustom
                Extender.MouseIcon = LoadResPicture(tdSizeNS, vbResCursor)
                cVS.Orientation = espHorizontal
                'debug.Print "ext mdown leftright dragvert cvsorihoriz"
            ElseIf Align = tdAlignTop Or _
                Align = tdAlignBottom Then
                Extender.MousePointer = vbCustom
                Extender.MouseIcon = LoadResPicture(tdSizeWE, vbResCursor)
                cVS.Orientation = espVertical
                
                'debug.Print "ext mdown topbottom dragvert cvsorivert"
            End If
        Case Else
            Extender.MousePointer = vbDefault
    End Select
    DragDir = over

Exit Sub
Err_Extender_MouseMove:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Extender_Paint
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:33:52
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Repaint the host
' ******************************************************************************
Private Sub Extender_Paint()
On Error GoTo Err_Extender_Paint
    Const constSource As String = m_constClassName & ".Extender_Paint"

    On Error Resume Next
    If m_bDirty Then
        Extender.Cls
        DockResize
        m_bDirty = False
    End If
    DockArrange

Exit Sub
Err_Extender_Paint:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Extender_Resize
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:34:02
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Resize the host
' ******************************************************************************
Private Sub Extender_Resize()
    If m_bSplitting = False Then
        m_bDirty = True
    End If
    Extender_Paint
End Sub

' ******************************************************************************
' Routine       : (Function) OverDragArea
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:24:22
' Inputs        :
' Outputs       :
' Modifications :
'
' Description   : Return drag flags indicating the area at this point
' ******************************************************************************
Private Function OverDragArea(ByVal x As Single, ByVal Y As Single)
On Error GoTo Err_OverDragArea
    Const constSource As String = m_constClassName & ".OverDragArea"

    Dim over As Integer
    ' initially, dragging is none
    over = DRAG_NONE
    If Align = tdAlignLeft And _
          x > Extender.width - (SPLITTER_WIDTH) And _
          x < Extender.width Then
       over = DRAG_HORIZONTAL
       'debug.Print "overdrag left horiz"
    ElseIf Align = tdAlignRight And _
          x > 0 And _
          x < (5 * SPLITTER_WIDTH) Then
       over = DRAG_HORIZONTAL
       'debug.Print "overdrag right horiz"
    ElseIf Align = tdAlignTop And _
        Y > Extender.Height - (SPLITTER_HEIGHT) And _
        Y < Extender.Height Then
       over = DRAG_HORIZONTAL
       'debug.Print "overdrag top horiz"
        ElseIf OverFormDragArea(x, Y) Then
               'debug.Print "overdrag bottom vert"
            over = DRAG_VERTICAL
    ElseIf Align = tdAlignBottom Then
        If Y > 0 And _
          Y < (SPLITTER_HEIGHT) Then
            over = DRAG_HORIZONTAL
                   'debug.Print "overdrag bottom horiz"
        ElseIf OverFormDragArea(x, Y) Then
               'debug.Print "overdrag bottom vert"
            over = DRAG_VERTICAL
        End If
    Else
        If OverFormDragArea(x, Y) Then
            over = DRAG_VERTICAL
        End If
    End If
    OverDragArea = over

Exit Function
Err_OverDragArea:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Function

Public Sub Refresh()
    Extender.Refresh
End Sub

' ******************************************************************************
' Routine       : (Sub) DockResize
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:21:50
' Inputs        :
' Outputs       :
' Modifications :
'
' Description   : Used to DockArrange all the docked forms when a resize
'                 method is performed
' ******************************************************************************
Public Sub DockResize()
On Error GoTo Err_DockResize
    Const constSource As String = m_constClassName & ".DockResize"

    Dim lHeight As Single
    Dim lWidth As Single
    Dim lTOp As Single
    Dim lLeft As Single
    Dim i As Integer
    Dim Count As Integer
    Dim lTotalPercent As Long
    Dim lMaxHeight As Long
    Dim lMaxWidth As Long

    
    Count = m_WindowList.VisibleCount
    ' no controls then exit
    If Count = 0 Then Exit Sub
    ' update coordinates based on align method of the host
    lLeft = 0
    lTOp = 0
    lHeight = Extender.ScaleHeight
    'debug.Print "Align "; Align
    ' determine maximum percentage for panel
    For i = 1 To WindowList.Count
        If m_WindowList(i).Visible Then
            lTotalPercent = lTotalPercent + m_WindowList(i).Percent
        End If
    Next i
    
    
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lWidth = Extender.ScaleWidth - SPLITTER_WIDTH
        lMaxHeight = ((lHeight - ((Count - 1) * SPLITTER_HEIGHT)))
        If Align = tdAlignRight Then
            lLeft = lLeft + SPLITTER_WIDTH
        End If
    ElseIf Align = tdAlignTop Or Align = tdAlignBottom Then
        lWidth = Extender.ScaleWidth
        lMaxHeight = (lHeight - SPLITTER_HEIGHT)
        lMaxWidth = (lWidth - ((Count - 1) * SPLITTER_WIDTH)) / Count
        If Align = tdAlignBottom Then
            lTOp = lTOp + SPLITTER_HEIGHT
        End If
    End If
    
    ' this is necessary because i don't wanna check if the parent
    ' form is minimize or not
    
    
    
    On Error Resume Next
    For i = 1 To WindowList.Count
        If m_WindowList(i).Visible Then
        
            If Align = tdAlignTop Or Align = tdAlignBottom Then
                ' height of window = percent / total percent * maxheight
                lHeight = lMaxHeight
                
                m_WindowList(i).Move lLeft, lTOp, lMaxWidth, lHeight
                lLeft = lLeft + lMaxWidth + SPLITTER_WIDTH
            Else
                ' height of window = percent / total percent * maxheight
                lHeight = Int((m_WindowList(i).Percent / lTotalPercent) * lMaxHeight)
                
                m_WindowList(i).Move lLeft, lTOp, lWidth, lHeight
                lTOp = lTOp + lHeight + SPLITTER_HEIGHT
            End If
        End If
    Next
    
Exit Sub
Err_DockResize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) DockArrange
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:15:59
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Reorder the visible forms within this host
' ******************************************************************************
Public Sub DockArrange(Optional formIndex As Integer = -1, Optional deltaOffset As Integer = 0)
On Error GoTo Err_DockArrange
    Const constSource As String = m_constClassName & ".DockArrange"

    Dim lHeight As Single
    Dim lWidth As Single
    Dim lTOp As Single
    Dim lLeft As Single
    Dim i As Integer
    Dim CurrentH As Long
    Dim NextH As Long
    Dim Count As Integer
    
    If formIndex > -1 Then
        m_FormIndex = formIndex
    End If
    
    If deltaOffset <> 0 Then
        deltaY = deltaOffset
    End If
    
    Count = m_WindowList.VisibleCount
    ' no controls (forms) then exit
    If Count = 0 Then Exit Sub
    ' reset position
    lLeft = 0
    lTOp = 0
    ' based on align property, re-ordering will take different actions
    If Align = tdAlignLeft Or _
        Align = tdAlignRight Then
        lWidth = Extender.ScaleWidth - SPLITTER_WIDTH
        If Align = tdAlignRight Then
            lLeft = lLeft + SPLITTER_WIDTH
        End If
    ElseIf Align = tdAlignTop Or _
        Align = tdAlignBottom Then
        lHeight = Extender.ScaleHeight - SPLITTER_HEIGHT
        If Align = tdAlignBottom Then
            lTOp = lTOp + SPLITTER_HEIGHT
        End If
    End If
    ' loop the form array to draw each control in its proper
    ' position
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then
            If Align = tdAlignLeft Or Align = tdAlignRight Then
                ' if this form is a splitform then resize it
                If i = m_FormIndex Then
                    ' current height
                    CurrentH = (m_WindowList(i).Height + deltaY)
                   ' height of the next form
                    NextH = (m_WindowList(i + 1).Height - deltaY)
                   ' clear split form
                    m_FormIndex = 0
                ElseIf NextH > 0 Then
                    ' update next form height
                    CurrentH = NextH
                    ' reset next height
                    NextH = 0
                Else
                    ' update current height
                    CurrentH = m_WindowList(i).Height
                End If
                ' update properties for moving the form
                lHeight = CurrentH
                m_WindowList(i).Move lLeft, lTOp, lWidth, lHeight
                m_WindowList(i).Extender.Refresh
                ' set next top position
                lTOp = lTOp + lHeight + SPLITTER_HEIGHT
            Else
                ' if this form is a splitform then resize it
                If i = m_FormIndex Then
                    ' current height
                    'Debug.Print deltaY
                    CurrentH = (m_WindowList(i).width + deltaY)
                    
                    If CurrentH < (getCaptionHeight + 3) * Screen.TwipsPerPixelX Then
                        CurrentH = (getCaptionHeight + 3) * Screen.TwipsPerPixelX
                    End If
                    ' height of the next form
                    NextH = (m_WindowList(i + 1).width - deltaY)
                    ' clear split form
                    m_FormIndex = 0
                ElseIf NextH > 0 Then
                    ' update next form height
                    CurrentH = NextH
                    ' reset next height
                    NextH = 0
                Else
                    ' update current height
                    CurrentH = m_WindowList(i).width
                End If
                ' update properties for moving the form
                lWidth = CurrentH
                'debug.Print lWidth
                m_WindowList(i).Move lLeft, lTOp, lWidth, lHeight
                m_WindowList(i).Extender.Refresh
                ' set next top position
                lLeft = lLeft + lWidth + SPLITTER_HEIGHT
            End If
            
        End If
    Next

Exit Sub
Err_DockArrange:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub
Public Sub DockCollapse(formIndex As Integer)

On Error GoTo Err_DockArrange

    Const constSource As String = m_constClassName & ".DockArrange"

    Dim lHeight As Single
    Dim lWidth As Single
    Dim lTOp As Single
    Dim lLeft As Single
    Dim i As Integer
    Dim CurrentH As Long
    Dim NextH As Long
    Dim Count As Integer
    Dim lMaxWidth As Integer
    Dim lMaxHeight As Integer
    Dim lCaptionHeightTwips As Integer
        
    Count = m_WindowList.VisibleCount
    ' no controls (forms) then exit
    If Count <= 1 Then Exit Sub
    ' reset position
    
    If Align = tdAlignBottom Or Align = tdAlignTop Then
        lCaptionHeightTwips = (getCaptionHeight + 3) * Screen.TwipsPerPixelX
    Else
        lCaptionHeightTwips = GetSystemMetrics(SM_CYMIN) * Screen.TwipsPerPixelX
    End If
    
    lLeft = 0
    lTOp = 0
    ' based on align property, re-ordering will take different actions
    
    
    If Align = tdAlignLeft Or _
        Align = tdAlignRight Then
        lWidth = Extender.ScaleWidth - SPLITTER_WIDTH
        If Align = tdAlignRight Then
            lLeft = lLeft + SPLITTER_WIDTH
        End If
        
        'Debug.Print GetSystemMetrics(SM_CXMIN), GetSystemMetrics(SM_CYMIN)
        
        lMaxHeight = Extender.ScaleHeight - ((lCaptionHeightTwips + SPLITTER_WIDTH) * (Count - 1))
        
    ElseIf Align = tdAlignTop Or _
        Align = tdAlignBottom Then
        lHeight = Extender.ScaleHeight - SPLITTER_HEIGHT
        If Align = tdAlignBottom Then
            lTOp = lTOp + SPLITTER_HEIGHT
        End If
        
        'Debug.Print "dockcoll"; Count; Me.width
        lMaxWidth = Extender.ScaleWidth - ((lCaptionHeightTwips + SPLITTER_WIDTH) * (Count - 1))
        
    End If
    
    ' loop the form array to draw each control in its proper
    ' position
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then
            If Align = tdAlignLeft Or Align = tdAlignRight Then
                ' if this form is a splitform then resize it
                If i = formIndex Then
                    CurrentH = lMaxHeight
                   ' clear split form
                    formIndex = 0
                Else
                    ' update current height
                    CurrentH = lCaptionHeightTwips
                End If
                
                ' update properties for moving the form
                lHeight = CurrentH
                
                m_WindowList(i).Move lLeft, lTOp, lWidth, lHeight
                m_WindowList(i).Extender.Refresh
                ' set next top position
                lTOp = lTOp + lHeight + SPLITTER_HEIGHT
            Else
                ' if this form is a splitform then resize it
                If i = formIndex Then
                    CurrentH = lMaxWidth '(m_WindowList(i).width + deltaY)
                    ' clear split form
                    formIndex = 0
                Else
                    ' update current height
                    CurrentH = lCaptionHeightTwips
                End If
                
                ' update properties for moving the form
                lWidth = CurrentH

                m_WindowList(i).Move lLeft, lTOp, lWidth, lHeight
                m_WindowList(i).Extender.Refresh
                ' set next top position
                lLeft = lLeft + lWidth + SPLITTER_HEIGHT
            End If
            
        End If
    Next

Exit Sub
Err_DockArrange:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub


' ******************************************************************************
' Routine       : (Function) ClassName
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:14:30
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Returns the Classname
' ******************************************************************************
Private Function ClassName(ByVal lhWnd As Long) As String
    Dim lLen As Long
    Dim sBuf As String
    lLen = 260
    sBuf = String(lLen, 0)
    lLen = GetClassName(lhWnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left(sBuf, lLen)
    End If
End Function

' ******************************************************************************
' Routine       : (Sub) DockHide
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:08:24
' Inputs        :
' Outputs       :
' Modifications :
'
' Description   : Simply hide a docked form and re-DockArrange the host
' ******************************************************************************
Private Sub DockHide(hWndA As Long)
On Error GoTo Err_DockHide
    Const constSource As String = m_constClassName & ".DockHide"

    ' hide window
    Call ShowWindow(hWndA, SW_HIDE)
    ' re-paint the host
    Extender_Paint

Exit Sub
Err_DockHide:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Function) OverFormDragArea
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:06:34
' Inputs        :
' Outputs       :
' Modifications :
'
' Description   : this will check if the cursor position is over
'                 a draggable are between forms
' ******************************************************************************
Private Function OverFormDragArea(x As Single, Y As Single, Optional wIndex As Integer) As Boolean
On Error GoTo Err_OverFormDragArea
    Const constSource As String = m_constClassName & ".OverFormDragArea"

    Dim i As Integer
    Dim lTOp As Long
    Dim Count As Integer
    Dim idx As Integer
    Dim lLeft As Long
    
    OverFormDragArea = False
    Count = m_WindowList.VisibleCount
    If Count = 0 Then Exit Function
    lTOp = 0
    wIndex = 0
    
    If Align = tdAlignTop Then
        Count = Count - 1
    ElseIf Align = tdAlignBottom Then
        lTOp = lTOp + SPLITTER_HEIGHT
    End If
    
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then
        
            If Align = tdAlignLeft Or Align = tdAlignRight Then
            
                lTOp = lTOp + m_WindowList(i).Height
                If Y >= lTOp And Y <= lTOp + SPLITTER_HEIGHT Then
                    OverFormDragArea = True
                    wIndex = i
                    Exit Function
                End If
                lTOp = lTOp + SPLITTER_HEIGHT
            Else
                lLeft = lLeft + m_WindowList(i).width
                If x >= lLeft And x <= lLeft + SPLITTER_WIDTH Then
                    OverFormDragArea = True
                    wIndex = i
                    Exit Function
                End If
                lLeft = lLeft + SPLITTER_WIDTH
            End If
        End If
    Next
    
Exit Function
Err_OverFormDragArea:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Function

' ******************************************************************************
' Routine       : (Sub) GetFormRect
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:15:07
' Inputs        :
' Outputs       :
' Modifications :
'
' Description   : Get the form rect in screen coordinates based on th index
' ******************************************************************************
Private Sub GetFormRect(ByVal Index As Integer, Rc As RECT)
On Error GoTo Err_GetFormRect
    Const constSource As String = m_constClassName & ".GetFormRect"

    Dim i As Integer
    Dim lTOp As Long
    Dim lBottom As Long
    Dim Rw As RECT
    
    If m_WindowList.Count = 0 Then Exit Sub
    
    GetWindowRect Extender.hwnd, Rw
    GetWindowRect m_WindowList(Index).hwnd, Rc
    
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lTOp = Rc.Top + (GetSystemMetrics(SM_CYCAPTION) + (10 * GetSystemMetrics(SM_CYBORDER)))
        'lTOp = Rw.Top
    End If
    
    GetWindowRect m_WindowList(Index + 1).hwnd, Rc
    
    If Rc.Bottom > Rw.Bottom Then
        Rc.Bottom = Rw.Bottom
    End If
    
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lBottom = Rc.Bottom - (GetSystemMetrics(SM_CYCAPTION) + (10 * GetSystemMetrics(SM_CYBORDER)))
        Rc.Top = lTOp
        Rc.Bottom = lBottom
    Else
        Rc.Left = Rw.Left
    End If

Exit Sub
Err_GetFormRect:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) UnDock
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 09/06/2000 - 14:59:00
' Inputs        : Byval df: TDockForm
' Outputs       :
' Modifications :
' Credits       : All credits for part of this code to vbWeb:
'
' Description   : Undock and show the form
' ******************************************************************************
Public Sub UnDock(ByVal Df As TDockForm)
On Error GoTo Err_UnDock
    Const constSource As String = m_constClassName & ".UnDock"

    Dim hWndA As Long
    Dim Style As Long
    Dim hWndDesktop As Long
    
    ' LockWindowUpdate Extender.Parent.hWnd
    
    ' change the dock form state
    Df.State = tdUndocked
    Df.Redraw = False
    
    ' get form handle
    hWndA = Df.hwnd
    ' if this form can not float then dock it
    If (Df.Style And tdDockFloat) = False Then
        ' if it can not float and it is undocked...
        If Df.State = tdUndocked Then
            ' hide the form
            Call ShowWindow(hWndA, SW_HIDE)
        End If
        Exit Sub
    End If
    ' remove the form from the control list
    m_WindowList.RemoveByHandle Df.hwnd
    ' This is necessary in order to make top and bottom
    ' panels to align after/before any existing
    ' aligned controls on your form (Statusbars, Toolbars, etc...)
    If Extender.Align = tdAlignTop Then Top = 1000
    If Extender.Align = tdAlignBottom Then Top = 0
    ' first of all let's hide this window
    ShowWindow hWndA, SW_HIDE
    ' change form's style
    Style = GetWindowLong(hWndA, GWL_STYLE)
    Style = Style And Not WS_DLGFRAME ' dialog frame
    Style = Style Or WS_CAPTION       ' has a caption
    Style = Style Or WS_THICKFRAME    ' has sizing border
    Style = Style Or WS_SYSMENU       ' add system menu
    ' set new style
    SetWindowLong hWndA, GWL_STYLE, Style
    ' This must be a toolwindow
    Style = GetWindowLong(hWndA, GWL_EXSTYLE)
    Style = Style Or WS_EX_TOOLWINDOW
    Style = Style And Not (WS_EX_APPWINDOW)
    SetWindowLong hWndA, GWL_EXSTYLE, Style
'    ' good to use it after SetWindowLong() function
'    SetWindowPos hWndA, 0, 0, 0, 0, 0, _
'        SWP_FRAMECHANGED Or SWP_NOMOVE Or _
'        SWP_NOZORDER Or SWP_NOSIZE
    ' set the parent to 0->Desktop (tricky!)
    Call SetParent(hWndA, hWndDesktop)
    ' move window to its floating position
    MoveWindow Df.hwnd, _
               Df.FloatingLeft, _
               Df.FloatingTop, _
               Df.FloatingWidth, _
               Df.FloatingHeight, _
1
    ' show the window
    ShowWindow hWndA, SW_SHOW
    ' make the form 'float' above the MDI form
    SetWindowLong hWndA, GWL_HWNDPARENT, Parent.Parent.hwnd
    ' update window pos
    SetWindowPos hWndA, 0, 0, 0, 0, 0, _
        SWP_FRAMECHANGED Or SWP_NOMOVE Or _
        SWP_NOZORDER Or SWP_NOSIZE
    ' extender is visible if there is more forms docked...
    If Not IsVisible Then
        Extender.Visible = False
    End If
    ' chage visiblity
    Df.Visible = True
    ' Call this function to raise TabDock Event FormUnDocked
    Parent.TriggerEvent "UnDock", Df
    ' needs to re-calc heights
    m_bDirty = True
    ' repaint the control
    Extender_Paint
    Df.Redraw = True
    ' free windows
'    LockWindowUpdate ByVal 0&

Exit Sub
Err_UnDock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Dock
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 09/06/2000 - 14:57:02
' Inputs        : df: TDockForm
' Outputs       :
' Modifications :
' Credits       : All credits for this code to VbWeb resources:
'
' Description   : Dock a form within host boundaries
' ******************************************************************************
Public Sub Dock(ByVal Df As TDockForm)

On Error GoTo Err_Dock
    Const constSource As String = m_constClassName & ".Dock"
    'debug.Print "Dock"
    Dim hWndA As Long
    Dim Style As Long
    Dim idx As Integer
    
'    LockWindowUpdate Extender.Parent.hWnd
    ' change the dock form state
    Df.State = tdDocked

    Df.Redraw = False
    ' get the form handle
    hWndA = Df.hwnd
    ' check if the form may dock here
    If AllowDocking(Df) = False Then
        ' if not just show the form and good bye
        Call ShowWindow(hWndA, SW_SHOW)
        Exit Sub
    End If
    ' We have to rebuild window list order
    If m_WindowList.HandleExists(hWndA) Then
        m_WindowList.RemoveByHandle hWndA
    End If
    ' get a new index for this window
    idx = GetFormIndex(Df.FloatingLeft, Df.FloatingTop)
    ' add this form to the collection
    m_WindowList.Append Df, idx
    ' This is necessary in order to make top and bottom
    ' panels to align after/before any
    ' aligned controls on your form
    If Extender.Align = tdAlignTop Then Top = 1000
    If Extender.Align = tdAlignBottom Then Top = 0
    ' hide window
    Call ShowWindow(hWndA, SW_HIDE)
    ' get the form style
    Style = GetWindowLong(hWndA, GWL_STYLE)
    ' set new style for the form
    Style = Style Or WS_DLGFRAME            ' have dialog frame
    Style = Style Or WS_SYSMENU             ' have sysmenu close button
    Style = Style Or WS_OVERLAPPED
    Style = Style Or WS_CLIPCHILDREN
    Style = Style Or WS_CLIPSIBLINGS
    Style = Style Or WS_OVERLAPPED
    Style = Style And Not WS_MAXIMIZE       ' No maximize here
    Style = Style And Not WS_MINIMIZE       ' No minimize too
    Style = Style Xor WS_MAXIMIZEBOX        ' No maximize box
    Style = Style Xor WS_MINIMIZEBOX        ' No minimize box
    Style = Style And Not WS_THICKFRAME     ' No sizing frame
    Style = Style And Not WS_SYSMENU        ' No sizing frame
    If Extender.Align = tdAlignTop Or Extender.Align = tdAlignBottom Then
        Style = Style And Not WS_CAPTION    ' no caption
        Df.Extender.BorderStyle = 0
    Else
        Style = Style And Not WS_CAPTION    ' no caption
        Df.Extender.BorderStyle = 0
        
        'Df.Extender.Refresh
    End If
    
    ' set the style
    SetWindowLong hWndA, GWL_STYLE, Style
    
    ' this must be a toolwindow
    Style = GetWindowLong(hWndA, GWL_EXSTYLE)
    Style = Style Or WS_EX_TOOLWINDOW
    Style = Style And Not (WS_EX_APPWINDOW)
    SetWindowLong hWndA, GWL_EXSTYLE, Style
    SetWindowPos hWndA, 0, 0, 0, 0, 0, _
        SWP_FRAMECHANGED Or SWP_NOMOVE Or _
        SWP_NOZORDER Or SWP_NOSIZE
    ' set the parent to the dock host
    Call SetParent(hWndA, hwnd)
    ' show the container once we have a new form
    Extender.Visible = True
    ' show docked window
    Call ShowWindow(hWndA, SW_SHOW)
    ' activate it please!
    'debug.Print "dock ncact"
    'Call SendMessage(hWndA, WM_NCACTIVATE, 1, 0)
    ' change the dock form state
    'Df.State = tdDocked
    ' chage visibility
    Df.Visible = True
    ' Call this procedure to raise a TabDock event
    'Parent.TriggerEvent "Dock", Df
    ' needs to re-calc heights then set this var to True
    m_bDirty = True
    ' repaint the control
    Extender_Paint
    ' style and update may take place now
    Df.Redraw = True
'    LockWindowUpdate ByVal 0&

    'Debug.Print "end docking"
    'SendMessage Df.hwnd, WM_SIZE, 0, 130

    Df.Extender.Height = Df.Extender.Height + 10
    'Df.Extender.height = Df.Extender.height - 10
    Df.Extender.width = Df.Extender.width + 10
    'Df.Extender.width = Df.Extender.width - 10
    ' Call this procedure to raise a TabDock event
    Parent.TriggerEvent "Dock", Df
    
Exit Sub
Err_Dock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Function) IsDocked
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 09/06/2000 - 14:56:37
' Inputs        : df : TDockform
' Outputs       :
' Modifications :
'
' Description   : check if the given docked form is docked in this host panel
' ******************************************************************************
Public Function IsDocked(Df As TDockForm) As Boolean
On Error GoTo Err_IsDocked
    Const constSource As String = m_constClassName & ".IsDocked"

    Dim idx As Integer
    
    IsDocked = False
    If m_WindowList.Count = 0 Then Exit Function
    For idx = 1 To m_WindowList.Count
        If m_WindowList(idx).hwnd = Df.hwnd Then
            IsDocked = True
            Exit Function
        End If
    Next

Exit Function
Err_IsDocked:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Function

' ******************************************************************************
' Routine       : (Function) IsVisible
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 09/06/2000 - 14:18:33
' Inputs        :
' Outputs       :
' Modifications :
' Description   : check if the given form is docked in here
'                 this is a friend function the end-user has no access
'                 to it
' ******************************************************************************
Friend Function IsVisible() As Boolean
On Error GoTo Err_IsVisible
    Const constSource As String = m_constClassName & ".IsVisible"

    Dim idx As Integer
    Dim Count As Integer
    
    IsVisible = False
    If m_WindowList.Count = 0 Then Exit Function
    For idx = 1 To m_WindowList.Count
        If m_WindowList(idx).Visible Then
            Count = Count + 1
        End If
    Next
    IsVisible = Count > 0
    
Exit Function
Err_IsVisible:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Function

' ******************************************************************************
' Routine       : (Sub) Repaint
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 09/06/2000 - 14:17:09
' Inputs        : N/A
' Outputs       : N/A
' Modifications :
' Description   : Repaint current host panel
' ******************************************************************************
Friend Sub Repaint()
On Error GoTo Err_Repaint
    Const constSource As String = m_constClassName & ".Repaint"

    m_bDirty = True
    Extender_Paint

Exit Sub
Err_Repaint:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Function) AllowDocking
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 6:50:30
' Inputs        : Byval df: TDockForm
' Outputs       :
' Modifications :
' Description   : Returns true if the specified form is allowed to dock
' ******************************************************************************
Friend Function AllowDocking(ByVal Df As TDockForm) As Boolean
On Error GoTo Err_AllowDocking
    Const constSource As String = m_constClassName & ".AllowDocking"

    Select Case Align
        Case tdAlignLeft
            AllowDocking = Df.Style And tdDockLeft
        Case tdAlignRight
            AllowDocking = Df.Style And tdDockRight
        Case tdAlignTop
            AllowDocking = Df.Style And tdDockTop
        Case tdAlignBottom
            AllowDocking = Df.Style And tdDockBottom
    End Select

Exit Function
Err_AllowDocking:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Function

' ******************************************************************************
' Routine       : GetHostRect
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:54:02
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Get the host rect based on its align property
'                 I use this function to figure out the position of the
'                 picture box object, once when it is invisible it is not
'                 possible to say where it really is
' ******************************************************************************
Friend Sub GetHostRect(ByRef Rc As RECT)
    ' if the Extender is visible no problem,
    ' just grab the visible rect
    If Extender.Visible Then
        GetWindowRect Extender.hwnd, Rc
    Else
        ' here the problem is that we don't know exactly where the
        ' Extender is so get container rect
        GetWindowRect Parent.Parent.hwnd, Rc
        ' get a rect based ona supposed position
        If Extender.Align = tdAlignRight Then
            Rc.Left = Rc.Right - (Parent.PanelWidth / Screen.TwipsPerPixelY)
        ElseIf Extender.Align = tdAlignBottom Then
            Rc.Top = Rc.Bottom - (Parent.PanelHeight / Screen.TwipsPerPixelY)
        ElseIf Extender.Align = tdAlignTop Then
            Rc.Bottom = Rc.Top + (Parent.PanelHeight / Screen.TwipsPerPixelY)
        ElseIf Extender.Align = tdAlignLeft Then
            Rc.Right = Rc.Left + (Parent.PanelWidth / Screen.TwipsPerPixelY)
        End If
    End If
End Sub

Private Function GetFormIndex(x As Long, Y As Long) As Integer
On Error GoTo Err_GetFormIndex
    Const constSource As String = m_constClassName & ".GetFormIndex"

    Dim i As Integer
    Dim Rc As RECT
    Dim Count As Integer
    
    GetFormIndex = 1
    Count = m_WindowList.VisibleCount
    If Count = 0 Then Exit Function
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then
            GetWindowRect m_WindowList(i).hwnd, Rc
            If PtInRect(Rc, x, Y) > 0 Then
                GetFormIndex = i + 1
                Exit Function
            End If
        End If
    Next
    GetWindowRect Extender.hwnd, Rc
    If Y < Rc.Top Then
        GetFormIndex = 1
    Else
        GetFormIndex = m_WindowList.Count + 1
    End If
Exit Function
Err_GetFormIndex:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Function

Friend Sub DockSwap(Df As TDockForm, x As Long, Y As Long)
On Error GoTo Err_DockSwap
    Const constSource As String = m_constClassName & ".DockSwap"

    Dim Index As Integer
    
    m_WindowList.RemoveByHandle Df.hwnd
    Index = GetFormIndex(x, Y)
    m_WindowList.Append Df, Index

Exit Sub
Err_DockSwap:
    Err.Raise Description:="Unexpected Error: " & Err.Description, _
             Number:=Err.Number, _
             Source:=constSource
End Sub
'-- end code

