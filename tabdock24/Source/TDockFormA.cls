VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************************
' Project     : TabDock
' Module      : TDockForm
' Description : Save the form properties and information to be used
'               by the docking engine
' Created by  : Marclei V Silva
' Machine     : ZEUS
' Date-Time   : 11/05/2000 22:15:58
' Revisions   : See revisions.rtf shipped with this version
' ********************************************************************
Option Explicit
'Dim Isubclass As New Isubclass
Private m_wndprcNext As Long
#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

' Keep up with the errors
Const g_ErrConstant As Long = vbObjectError + 1000
Const m_constClassName = "TDockForm"
Private PushDown As Boolean
Private StoreHWND As Long

Private m_lngErrNum As Long
Private m_strErrStr As String
Private m_strErrSource As String

Private m_bActive As Boolean
Private m_bStyleChanged As Boolean
Private m_bUnloaded As Boolean

' class properties
Private m_bDragging As Boolean      ' We are dragging?
Private m_SelectedItem As Integer   ' Host that was active during moving
Private m_lFloatingWidth As Long    ' Floating dimensions
Private m_lFloatingHeight As Long
Private m_lFloatingLeft As Long
Private m_lFloatingTop As Long
Private m_State As tdDockedState    ' form state dock/undocked
Private m_Style As tdDockStyles     ' availlable style
Private m_PanelPtr As Long          ' pointer to panel object
Private m_Key As String             ' key associatate to the form
Private m_Percent As Integer        ' percent is how much of Panel form will occupy
Private m_hWnd As Long              ' handle for the form
Private m_FormPtr As Long           ' pointer to the form object
Private m_bRedraw As Boolean        ' flag that stop form painting or style change



'*********************************
Public Enum EHitTestAreas
   HTERROR = (-2)
   HTTRANSPARENT = (-1)
   HTNOWHERE = 0
   HTCLIENT = 1
   HTCAPTION = 2
   HTSYSMENU = 3
   HTGROWBOX = 4
   HTMENU = 5
   HTHSCROLL = 6
   HTVSCROLL = 7
   HTMINBUTTON = 8
   HTMAXBUTTON = 9
   HTLEFT = 10
   HTRIGHT = 11
   HTTOP = 12
   HTTOPLEFT = 13
   HTBOTTOM = 15
   HTBOTTOMLEFT = 16
   HTBOTTOMRIGHT = 17
   HTBORDER = 18
End Enum

' Array of objects to perform WM_NCHITTEST on:
Private m_lHTPtr As Long
' Array of hWnds to perform WM_NCHITTEST on:
Private m_hWndh As Long
' Number of objects to perform WM_NCHITTEST on:
Private m_iCount As Long

' Rectangles to perform hit testing on:
Private Type tHitTestRectangle
   hwnd As Long
   sKey As String
   lLeft As Long
   lTOp As Long
   lRight As Long
   lBottom As Long
   eResponse As EHitTestAreas
End Type
Private m_tHTR As tHitTestRectangle
Private m_iHTRCount As Long

Private Const SPLITTER_HEIGHT = 80
Private Const SPLITTER_WIDTH = 80

'**********************************
Implements ISubclass

Public Property Let ParentHWND(ByVal hwnd As Long)
  StoreHWND = hwnd
End Property

Public Property Get ParentHWND() As Long
  ParentHWND = StoreHWND
End Property

Public Property Get Object() As TDockForm
    Set Object = Me
End Property

Friend Property Get Extender() As Form
    Set Extender = ObjectFromPtr(m_FormPtr)
End Property

Friend Property Set Extender(ByVal vData As Form)
    m_FormPtr = PtrFromObject(vData)
    Attach vData.hwnd
End Property

Friend Function Move(Left, Optional Top, Optional width, Optional Height)
    On Error Resume Next
    Extender.Move Left, Top, width, Height
End Function

Friend Property Get Left() As Variant
    Left = Extender.Left
End Property

Friend Property Let Left(ByVal vData As Variant)
    Extender.Left = vData
End Property

Friend Property Get Top() As Variant
    Top = Extender.Top
End Property

Friend Property Let Top(ByVal vData As Variant)
    Extender.Top = vData
End Property

Friend Property Get width() As Variant
    width = Extender.width
End Property

Friend Property Let width(ByVal vData As Variant)
    Extender.width = vData
End Property

Friend Property Get Height() As Variant
    Height = Extender.Height
End Property

Friend Property Let Height(ByVal vData As Variant)
    Extender.Height = vData
End Property

Friend Property Let Visible(ByVal vData As Boolean)
    Extender.Visible = vData
    If vData Then
        ' trigger TabDock.FormShow() Event
        Panel.Parent.TriggerEvent "ShowForm", Object
    Else
        ' trigger TabDock.Formhide() Event
        Panel.Parent.TriggerEvent "HideForm", Object
    End If
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns wether the form is visible or not"
    Visible = Extender.Visible
End Property

Public Property Let Style(ByVal vData As tdDockStyles)
Attribute Style.VB_Description = "Returns or sets the form style. The style may allow or avoid some features of a docked window. Check all the possible styles to choose the apropriate one."
    m_Style = vData
End Property

Public Property Get Style() As tdDockStyles
Attribute Style.VB_Description = "Returns or sets the form style. The style may allow or avoid some features of a docked window. Check all the possible styles to choose the apropriate one."
    Style = m_Style
End Property

Public Property Let Key(ByVal vData As String)
Attribute Key.VB_Description = "Returns the key of the form to be used with TabDock functions like FormShow() or FormHide. Instead of the key you may use the Form index to retrieve form reference"
    m_Key = vData
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key of the form to be used with TabDock functions like FormShow() or FormHide. Instead of the key you may use the Form index to retrieve form reference"
    Key = m_Key
End Property

Public Property Let FloatingHeight(ByVal vData As Long)
Attribute FloatingHeight.VB_Description = "Returns the floating height of the form"
    m_lFloatingHeight = vData
End Property

Public Property Get FloatingHeight() As Long
Attribute FloatingHeight.VB_Description = "Returns the floating height of the form"
    FloatingHeight = m_lFloatingHeight
End Property

Public Property Let FloatingWidth(ByVal vData As Long)
Attribute FloatingWidth.VB_Description = "Returns the floating Width of the form"
    m_lFloatingWidth = vData
End Property

Public Property Get FloatingWidth() As Long
Attribute FloatingWidth.VB_Description = "Returns the floating Width of the form"
    FloatingWidth = m_lFloatingWidth
End Property

Public Property Let FloatingLeft(ByVal vData As Long)
Attribute FloatingLeft.VB_Description = "Returns the floating Left of the form"
    m_lFloatingLeft = vData
End Property

Public Property Get FloatingLeft() As Long
Attribute FloatingLeft.VB_Description = "Returns the floating Left of the form"
    FloatingLeft = m_lFloatingLeft
End Property

Public Property Let FloatingTop(ByVal vData As Long)
Attribute FloatingTop.VB_Description = "Returns the floating Top of the form"
    m_lFloatingTop = vData
End Property

Public Property Get FloatingTop() As Long
Attribute FloatingTop.VB_Description = "Returns the floating Top of the form"
    FloatingTop = m_lFloatingTop
End Property

Friend Property Let Redraw(ByVal vData As Boolean)
    m_bRedraw = vData
End Property

Friend Property Get Redraw() As Boolean
    Redraw = m_bRedraw
End Property

Friend Property Set Panel(ByVal vData As TTabDockHost)
    m_PanelPtr = PtrFromObject(vData)
End Property

Public Property Get Panel() As TTabDockHost
Attribute Panel.VB_Description = "Returns the panel that is responsible for the form. This property changes when you dock the form into a different panel."
    Set Panel = ObjectFromPtr(m_PanelPtr)
End Property
Public Property Let Percent(ByVal vData As Integer)
    m_Percent = vData
End Property

Public Property Get Percent() As Integer
    Percent = m_Percent
End Property
Public Property Let State(ByVal vData As tdDockedState)
Attribute State.VB_Description = "Returns the state of the form. Docked or Undocked."
    m_State = vData
End Property

Public Property Get State() As tdDockedState
Attribute State.VB_Description = "Returns the state of the form. Docked or Undocked."
    State = m_State
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns the form handle. This may not be used in api callings once the docked form is already subclassed and another subclassing of this form may crash Vb IDE or unpredictable results may take place."
    hwnd = Extender.hwnd
End Property

' ******************************************************************************
' Routine       : (Sub) Class_Initialize
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:36:23
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : Initialize class
' ******************************************************************************
Private Sub Class_Initialize()
    On Error GoTo Err_Class_Initialize
    Const constSource As String = m_constClassName & ".Class_Initialize"

    m_bActive = False
    m_bStyleChanged = False
    m_SelectedItem = 0
    
    Exit Sub
Err_Class_Initialize:
    Err.Raise Description:=Err.Description, _
       Number:=Err.Number, _
       Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Class_Terminate
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:36:03
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : Terminate the class (clean it up)
' ******************************************************************************
Private Sub Class_Terminate()
    On Error GoTo Err_Class_Terminate
    Const constSource As String = m_constClassName & ".Class_Terminate"
    
    ' detach all window messages
    Detach
    ' try to unload the extender
    Unload Extender
    
    Exit Sub
Err_Class_Terminate:
    Err.Raise Description:=Err.Description, _
       Number:=Err.Number, _
       Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) DockForm_StoreFloatingPos
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 09/06/2000 12:47:34
' Inputs        :
' Outputs       :
' Modifications :
' Description   : Store the floating position of the form
' ******************************************************************************
Private Sub DockForm_StoreFloatingPos(lLeft As Long, lTOp As Long, lWidth As Long, lHeight As Long)
    m_lFloatingLeft = lLeft
    m_lFloatingTop = lTOp
    m_lFloatingWidth = lWidth
    m_lFloatingHeight = lHeight
End Sub

' ******************************************************************************
' Routine       : (Sub) DockForm_Show
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:35:36
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : shows the docked form no matter it is docked or undocked
' ******************************************************************************
Friend Sub DockForm_Show()
    On Error GoTo Err_DockForm_Show
    Const constSource As String = m_constClassName & ".DockForm_Show"

    If State = tdDocked Then
        Panel.Dock Me
    Else
        Panel.UnDock Me
    End If
    Attach Me.hwnd
    m_bUnloaded = False
    DockForm_ResetStyle
    DockForm_Paint
    
    Exit Sub
Err_DockForm_Show:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : (Sub) Hide
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:35:19
' Inputs        : N/A
' Outputs       : N/A
' Modifications : N/A
' Description   : hides the docked form no matter it is docked or undocked
' ******************************************************************************
Friend Sub DockForm_Hide()
    On Error GoTo Err_Hide
    Const constSource As String = m_constClassName & ".Hide"

    Visible = False
    If Panel.IsVisible Then
        Panel.Repaint
    Else
        Panel.Visible = False
    End If
    Exit Sub
Err_Hide:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

' ******************************************************************************
' Routine       : Dock
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:22:12
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : dock the form
' ******************************************************************************
Public Sub DockForm_Dock()
    
    Panel.Dock Me

End Sub

' ******************************************************************************
' Routine       : UnDock
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:22:26
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : undock the form
' ******************************************************************************
Public Sub DockForm_UnDock()
    Panel.UnDock Me
End Sub

Private Sub DockForm_Activate(ByVal bByMouse As Boolean)
    DockForm_Paint
End Sub

Private Sub DockForm_Paint()
    Dim Rc As RECT
    Dim bdrStyle As Long
    Dim bdrSides As Long
    Dim BorderStyle  As tdBorderStyles
    Dim CaptionStyle As tdCaptionStyles
    Dim hDC As Long
    Dim captionHeight As Long
    
    If Me.Extender.Name = "Form6" Then
        'debug.Print Panel.Align, Me.Extender.Name, Me.State
    End If
    
    If Me.State = tdUndocked Or _
        m_bRedraw = False Or _
        Extender.Visible = False Then
        If Me.Extender.Name = "Form6" Then
            'debug.Print "No Repaint"
        End If
        Exit Sub
    End If
    
    ' draw a custom border based on parante's color
    DrawBorder Extender, Panel.Parent.BackColor
    ' retrieve TabDock border style
    BorderStyle = Panel.Parent.BorderStyle
    CaptionStyle = Panel.Parent.CaptionStyle
    ' all sides must be updated
    bdrSides = BF_RECT
    ' update border styles
    If BorderStyle = bdrFlat Then bdrSides = bdrSides Or BF_FLAT
    If BorderStyle = bdrMono Then bdrSides = bdrSides Or BF_MONO
    If BorderStyle = bdrSoft Then bdrSides = bdrSides Or BF_SOFT
    Select Case BorderStyle
        Case bdrRaisedOuter: bdrStyle = BDR_RAISEDOUTER
        Case bdrRaisedInner: bdrStyle = BDR_RAISEDINNER
        Case bdrRaised: bdrStyle = EDGE_RAISED
        Case bdrSunkenOuter: bdrStyle = BDR_SUNKENOUTER
        Case bdrSunkenInner: bdrStyle = BDR_SUNKENINNER
        Case bdrSunken: bdrStyle = EDGE_SUNKEN
        Case bdrEtched: bdrStyle = EDGE_ETCHED
        Case bdrBump: bdrStyle = EDGE_BUMP
        Case bdrFlat: bdrStyle = BDR_SUNKEN
        Case bdrMono: bdrStyle = BDR_SUNKEN
        Case bdrSoft: bdrStyle = BDR_RAISED
    End Select
    
    ' get a window rect by hand
    ' GetWindowRect will not work here!
    Rc.Left = 0
    Rc.Top = 0
    Rc.Bottom = Extender.Height / Screen.TwipsPerPixelY
    Rc.Right = Extender.width / Screen.TwipsPerPixelY
    ' First get the window DC
    hDC = GetWindowDC(hwnd)
    ' Simply call the API and draw the edge.
    DrawEdge hDC, Rc, bdrStyle, bdrSides
    
    ' get a window rect by hand
    ' GetWindowRect will not work here!
    Rc.Left = 0
    Rc.Top = 0
    Rc.Bottom = Panel.Height / Screen.TwipsPerPixelY + 1
    Rc.Right = Panel.width / Screen.TwipsPerPixelX
    
    If Panel.Align = tdAlignLeft Then
        Rc.Right = (Panel.width - SPLITTER_WIDTH) / Screen.TwipsPerPixelX + 1
        Rc.Bottom = Panel.Height / Screen.TwipsPerPixelY
    End If
    
    If Panel.Align = tdAlignRight Then
        Rc.Left = SPLITTER_WIDTH / Screen.TwipsPerPixelX
        Rc.Bottom = Panel.Height / Screen.TwipsPerPixelY
    End If
    
    If Panel.Align = tdAlignTop Then
        Rc.Bottom = (Panel.Height - SPLITTER_HEIGHT) / Screen.TwipsPerPixelY + 1
    End If
    
    If Panel.Align = tdAlignBottom Then
        Rc.Top = SPLITTER_HEIGHT / Screen.TwipsPerPixelY
        Rc.Bottom = Panel.Height / Screen.TwipsPerPixelY
    End If
    
    ' First get the window DC
    hDC = GetWindowDC(Panel.hwnd)
    ' Simply call the API and draw the edge.
    DrawEdge hDC, Rc, bdrStyle, bdrSides
    
    Rc.Left = 0
    Rc.Top = 0
    Rc.Bottom = Extender.Height / Screen.TwipsPerPixelY
    Rc.Right = Extender.width / Screen.TwipsPerPixelY
    
    hDC = GetWindowDC(hwnd)
    '************************************************

    If Me.Extender.BorderStyle = vbBSNone Or Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
        'draw custom caption here!!!!

        Dim captionRect As RECT
        Dim frameRect As RECT
        Dim hBr As Long
           
        captionHeight = getCaptionHeight + 6

        hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
        
        With captionRect
            .Top = Rc.Top + 3
            .Left = Rc.Left + 3
            .Right = Rc.Right - 2
            .Bottom = captionHeight - 4
            
            If Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
                .Top = .Top
                .Bottom = Rc.Bottom - 3
                .Right = captionHeight - 4
            End If
            
        End With

        
        'blank out current caption
        
        FillRect hDC, captionRect, hBr
        DeleteObject hBr
        
         Select Case CaptionStyle
            Case tdCaptionNormal ' = 0
                drawGradient captionRect, hDC, Me.Extender.Caption, m_bActive, False, Panel.Align, Me.Extender
            Case tdCaptionEtched ' = 1
                drawGripper captionRect, hDC, EDGE_ETCHED, BF_RECT, False, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSoft ' = 2
                drawGripper captionRect, hDC, BDR_RAISED, BF_RECT Or BF_SOFT, False, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionRaised ' = 3
                drawGripper captionRect, hDC, EDGE_RAISED, BF_RECT, False, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionRaisedInner ' = 4
                drawGripper captionRect, hDC, BDR_RAISEDINNER, BF_RECT, False, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSunkenOuter ' = 5
                drawGripper captionRect, hDC, BDR_SUNKENOUTER, BF_RECT, False, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSunken ' = 6
                drawGripper captionRect, hDC, EDGE_SUNKEN, BF_RECT, False, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSingleRaisedBar ' = 7
                drawGripper captionRect, hDC, EDGE_RAISED, BF_RECT, True, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSingleRaisedInner ' = 9
                drawGripper captionRect, hDC, BDR_RAISEDINNER, BF_RECT, True, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSingleSoft ' = 10
                drawGripper captionRect, hDC, BDR_RAISED, BF_RECT Or BF_SOFT, True, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSingleEtched ' = 11
                drawGripper captionRect, hDC, EDGE_ETCHED, BF_RECT, True, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSingleSunken ' = 12
                drawGripper captionRect, hDC, EDGE_SUNKEN, BF_RECT, True, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionSingleSunkenOuter ' = 13
                drawGripper captionRect, hDC, BDR_SUNKENOUTER, BF_RECT, True, captionHeight, Panel.Align, Panel.Parent.MaximizeButton
            Case tdCaptionGradient ' = 8
                drawGradient captionRect, hDC, Me.Extender.Caption, m_bActive, True, Panel.Align, Me.Extender
            Case tdCaptionOfficeXP ' = 14
                drawOfficeXP captionRect, hDC, Me.Extender.Caption, m_bActive, True, Panel.Align, Me.Extender

        End Select
        
        'draw close box on form
                
        With frameRect
            .Top = Rc.Top + 5
            .Left = Rc.Right - captionHeight + 6
            .Right = Rc.Right - 4
            .Bottom = captionHeight - 5
            
            If Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
                .Right = Rc.Left + captionHeight - 6
                .Left = Rc.Left + 4
            End If
            
        End With

        If Me.Extender.Name = "Form6" Then
            'debug.Print frameRect.Top, frameRect.Left, frameRect.Right, frameRect.Bottom
        End If
        
        DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONCLOSE 'Or DFCS_FLAT
        
        If Panel.Parent.MaximizeButton Then
            If Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
                With frameRect
                    .Top = .Bottom + 2
                    .Bottom = .Top + captionHeight - 10
                    '.Right = .Right + 1
                    '.Left = .Left - 1
                End With
                DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE 'Or DFCS_FLAT
                
            Else
                With frameRect
                    .Right = .Left - 2
                    .Left = .Right - captionHeight + 10
                End With
                DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE 'Or DFCS_FLAT
                
            End If
                
        End If

    End If
    '*********************************************************
    ' release it

    ReleaseDC hwnd, hDC
End Sub

Private Sub DockForm_DblClick()
    'Debug.Print "dblclick"
    If (Me.Style And tdDockFloat) <> tdDockFloat Then
        Exit Sub
    End If
    If State = tdDocked Then
        Panel.UnDock Me
    Else
        Panel.Dock Me
    End If
End Sub

Private Sub DockForm_CaptionClose()
    Dim pt As POINTAPI
    Dim ptPrev As POINTAPI
    Dim objRect As RECT
    Dim lObjLeft As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    Dim captionHeight As Long
    Dim hDC As Long
    
    ReleaseCapture
    'debug.Print Me.Extender.Caption
    GetWindowRect Me.Extender.hwnd, objRect
    
    'objRect.Left = 0
    'objRect.Top = 0
    'objRect.Bottom = Extender.height / Screen.TwipsPerPixelY
    'objRect.Right = Extender.width / Screen.TwipsPerPixelY
    ' First get the window DC
    
    captionHeight = getCaptionButtonHeight
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt
    
    'Store the initial cursor position
    ptPrev.x = pt.x
    ptPrev.Y = pt.Y
    
'****************************************************
    Dim frameRect As RECT
        
    With frameRect
        .Top = objRect.Top + 5
        .Left = objRect.Right - captionHeight
        .Right = objRect.Right - 3
        .Bottom = .Top + captionHeight
    End With
    'debug.Print pt.x, pt.Y
    'debug.Print frameRect.Top, frameRect.Left, frameRect.Right, frameRect.Bottom
'*****************************************************
    If PtInRect(frameRect, pt.x, pt.Y) Then
        
        hDC = GetWindowDC(Me.Extender.hwnd)
        
        GetClientRect Me.Extender.hwnd, objRect
        
    With frameRect
        .Top = objRect.Top + 5
        .Left = objRect.Right - captionHeight + 5
        .Right = objRect.Right + 2
        .Bottom = .Top + captionHeight - 3
    End With
    
        
        DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONCLOSE + DFCS_PUSHED
        'debug.Print "redraw"
        'SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
        'Exit Sub
    End If
'*****************************************************

End Sub
' ******************************************************************************
' Routine       : DockForm_MouseDown
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:23:00
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Do some actions before dragging a form
' ******************************************************************************
Private Sub DockForm_MouseDown(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long, Cancel As Boolean)
    'debug.Print "mousedown"
    
    If Style And tdDockFloat Then
        m_bDragging = True
        ' if form is undocked then save the floating position
        If State = tdUndocked Then
            DockForm_StoreFloatingPos FormLeft, FormTop, FormWidth, FormHeight
        End If
    Else
        Cancel = True
    End If
End Sub

' ******************************************************************************
' Routine       : DockForm_MouseUp
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:23:29
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Actions after a form is dropped somewhere
' ******************************************************************************
Private Sub DockForm_MouseUp(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    ' If no host was found to dock the form then
    ' undock it if it is not undocked
    If m_SelectedItem = 0 Then
        If m_State = tdDocked Then
            ' undock the form only if it can float
            If m_Style And tdDockFloat Then
                Panel.UnDock Me
                ' repaint the host
                Panel.Repaint
            End If
        End If
        Exit Sub
    End If
    ' if form is undocked then dock it
    If m_State = tdUndocked Then
        ' dock the form
        If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
            Panel.Parent.Panels(m_SelectedItem).Dock Me
        End If
    ElseIf m_State = tdDocked Then
        ' if the selected host is the active host
        ' just repaint the host panel
        If m_SelectedItem = Panel.Index Then
            Panel.DockSwap Me, FormLeft, FormTop
            Panel.Repaint
        Else
            ' if selected host is different from
            ' active host, first undock the form
            If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
                Panel.UnDock Me
                ' dock the form into the selected host
                Panel.Parent.Panels(m_SelectedItem).Dock Me
            End If
        End If
    End If
    ' A docking site changed then we must set
    ' the new dock host for this form
    If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
        Set Panel = Panel.Parent.Panels(m_SelectedItem)
    End If
End Sub

' ******************************************************************************
' Routine       : DockForm_MouseMove
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:25:03
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Called when the form is moving around
' ******************************************************************************
Private Sub DockForm_MouseMove(ByRef FormLeft As Long, ByRef FormTop As Long, ByRef FormWidth As Long, ByRef FormHeight As Long, ByRef lBorderWidth As Long)
    Dim Rc As RECT
    Dim i As Integer
    Dim Found As Boolean
    Dim pt As POINTAPI
    Dim lX1Offset As Long
    Dim lX2Offset As Long
    Dim lYOffset As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    
    ' flag that indicates we found a dock host
    Found = False
    ' get the cursor position
    GetCursorPos pt
    ' loop the panels to see we find a docking host
    ' for this form
    For i = 1 To 4
        ' get the host rect
        Panel.Parent.Panels(i).GetHostRect Rc
        ' check if the form is within the panel's boundaries
        If PtInRect(Rc, pt.x, pt.Y) Then
            If Panel.Parent.Panels(i).AllowDocking(Me) Then
                ' a host was found and it is different from the
                ' the current selected one
                If Not Panel.Parent.Panels(i).Index = m_SelectedItem Then
                    ' host selection changed here
                    m_SelectedItem = Panel.Parent.Panels(i).Index
                    ' we don't have a dragging boundary set yet
                    ' for this host then update with the host's rect
                    lObjWidth = Rc.Right - Rc.Left
                    lObjHeight = Rc.Bottom - Rc.Top
                    If Panel.Parent.Panels(i).Align = tdAlignLeft Or Panel.Parent.Panels(i).Align = tdAlignRight Then
                        lX1Offset = pt.x - FormLeft
                        lX2Offset = lObjWidth * lX1Offset / FormWidth 'pt.Y - rc.Top
                        FormLeft = pt.x - lX2Offset
                        FormWidth = lObjWidth
                        FormHeight = lObjHeight
                    Else
                        lX1Offset = pt.Y - FormTop
                        lX2Offset = lObjHeight * lX1Offset / FormHeight 'pt.Y - rc.Top
                        FormTop = pt.Y - lX2Offset
                        FormWidth = lObjWidth
                        FormHeight = lObjHeight
                    End If
                End If
                ' change border width (special effect)
                lBorderWidth = 2
                ' we found a host, set flag = true
                Found = True
            End If
        End If
        ' we have found a host so exit
        If Found Then Exit For
    Next
    ' if the form has not found any host to dock in
    ' then let the rectangle image the same as its
    ' floating original rectangle
    If Not Found Then
        ' if we are dragging and the form is docked
        If m_bDragging And m_State = tdDocked Then
            ' get active host rect
            GetWindowRect Panel.hwnd, Rc
            ' restore form left property to its
            ' floating left position
            lObjWidth = Rc.Right - Rc.Left
            lObjHeight = Rc.Bottom - Rc.Top
            lX1Offset = pt.x - FormLeft
            lX2Offset = m_lFloatingWidth * lX1Offset / lObjWidth 'pt.Y - rc.Top
            FormLeft = pt.x - lX2Offset
            ' flag must be false
            m_bDragging = False
        End If
        ' change border width (just a special effect)
        lBorderWidth = 3
        ' update width and height with floating dimensions
        FormWidth = m_lFloatingWidth
        FormHeight = m_lFloatingHeight
        ' store floating position
        DockForm_StoreFloatingPos FormLeft, FormTop, FormWidth, FormHeight
        ' clear host selection
        m_SelectedItem = 0
    End If
End Sub

' ******************************************************************************
' Routine       : DockForm_Unload
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:27:09
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Hides the form
' ******************************************************************************
Private Sub DockForm_Unload()
    DockForm_Hide
End Sub

' ******************************************************************************
' Routine       : DockForm_ResetStyle
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:27:39
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Resets the style to a tollwindow
'                 This routine is necessay when the style of
'                 the form changes to force it to be a toolwindow
' ******************************************************************************
Private Sub DockForm_ResetStyle()
    Dim New_Style As Long
    
    ' if redraw is disable do not set any style
    ' or else you may get a loop here
    If m_bRedraw = False Then Exit Sub
    ' set this flag to true so that the next set style event will
    ' not be handled (trick!)
    m_bStyleChanged = True
    ' get current window style
    New_Style = GetWindowLong(hwnd, GWL_STYLE)
    ' check form state
    If State = tdDocked Then
        New_Style = New_Style And Not WS_THICKFRAME
        SetWindowLong hwnd, GWL_STYLE, New_Style
    Else
        New_Style = New_Style And Not WS_DLGFRAME
        New_Style = New_Style Or WS_CAPTION
        New_Style = New_Style Or WS_THICKFRAME
    End If
    ' get extended style
    New_Style = GetWindowLong(hwnd, GWL_EXSTYLE)
    ' set it to toolwindow
    New_Style = New_Style Or WS_EX_TOOLWINDOW
    New_Style = New_Style And Not (WS_EX_APPWINDOW)
    SetWindowLong hwnd, GWL_EXSTYLE, New_Style
    ' reset flag here
    m_bStyleChanged = False
End Sub

' ******************************************************************************
' Routine       : DockForm_Deactivate
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:29:52
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Called when the form is deactivated
' ******************************************************************************
Private Sub DockForm_Deactivate()
    DockForm_Paint
End Sub


' ******************************************************************************
' Routine       : (Sub) DragObject
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:38:15
' Inputs        :
' Outputs       :
' Modifications :
' Credits       : This rotine was extract from the DockForm_Dragger project
'               : by Steve
'               : Steve81@ mediaone.net
'               : http://www.vbtutor.com
'               : Thanks Steve!
'               :
' Modifications : Added the ability to change the dragging border width
'               : Removed reposition
'               :
' Description   : Procedure which simulates windows dragging of an object
' ******************************************************************************
Private Sub DragObject(ByVal hwnd As Long)
    Dim pt As POINTAPI
    Dim ptPrev As POINTAPI
    Dim objRect As RECT
    Dim DragRect As RECT
    Dim na As Long
    Dim lBorderWidth As Long
    Dim lObjTop As Long
    Dim lObjLeft As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim captionHeight As Long
    Dim bMoved As Boolean
    Dim bCancel As Boolean
    Dim hDC As Long
    
    ReleaseCapture
    GetWindowRect hwnd, objRect
    
    
    captionHeight = getCaptionButtonHeight
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt
    
    'Store the initial cursor position
    ptPrev.x = pt.x
    ptPrev.Y = pt.Y
    
    
'****************************************************
    Dim frameRect As RECT
    Dim frameRectShrink As RECT
        
    With frameRect
        .Top = objRect.Top + 5
        .Left = objRect.Right - captionHeight
        .Right = objRect.Right - 3
        .Bottom = .Top + captionHeight
        
        If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
            frameRectShrink.Top = .Bottom + 2
            frameRectShrink.Bottom = frameRectShrink.Top + captionHeight - 5
        Else
            frameRectShrink.Top = .Top
            frameRectShrink.Bottom = .Bottom
        End If
        
        If Me.State = tdDocked Then
            If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                .Right = objRect.Left + captionHeight
                .Left = objRect.Left + 4
                frameRectShrink.Right = .Right
                frameRectShrink.Left = .Left
            Else
                frameRectShrink.Right = .Left - 2
                frameRectShrink.Left = frameRectShrink.Right - captionHeight + 3
            End If
        End If
    End With
    

'*****************************************************
    If PtInRect(frameRect, pt.x, pt.Y) Then
        hDC = GetWindowDC(Me.Extender.hwnd)
        GetClientRect Me.Extender.hwnd, objRect
        
        With frameRect
            .Top = objRect.Top + 5
            .Left = objRect.Right - captionHeight + 5 - 6
            .Right = objRect.Right + 2 - 6
            .Bottom = .Top + captionHeight - 3
            
            If Me.State = tdDocked Then
                If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                    .Right = objRect.Left + captionHeight
                    .Left = objRect.Left + 4
                End If
            Else
                    .Right = .Right + 5
                    .Left = .Left + 5
                    .Top = .Top + 1
            End If
            
        End With
        
        DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONCLOSE Or DFCS_PUSHED
        Sleep (100)
        SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
        Exit Sub
    End If
    
    If PtInRect(frameRectShrink, pt.x, pt.Y) And Panel.Parent.MaximizeButton Then
        hDC = GetWindowDC(Me.Extender.hwnd)
        GetClientRect Me.Extender.hwnd, objRect
        
        With frameRect
            .Top = objRect.Top + 5
            .Left = objRect.Right - captionHeight + 5
            .Right = objRect.Right + 2
            .Bottom = .Top + captionHeight - 3
            
            If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                frameRectShrink.Top = .Bottom + 2
                frameRectShrink.Bottom = frameRectShrink.Top + captionHeight - 3
            Else
                frameRectShrink.Top = .Top
                frameRectShrink.Bottom = .Bottom
            End If
            
            If Me.State = tdDocked Then
                If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                    .Right = objRect.Left + captionHeight
                    .Left = objRect.Left + 4
                    frameRectShrink.Right = .Right + 1
                    frameRectShrink.Left = .Left
                Else
                    frameRectShrink.Right = .Left - 2 - 6
                    frameRectShrink.Left = frameRectShrink.Right - captionHeight + 3
                End If
            End If
                        
        End With
        
        If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
            DrawFrameControl hDC, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_PUSHED
            Sleep (100)
            DrawFrameControl hDC, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE
            
        Else
            DrawFrameControl hDC, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_PUSHED
            Sleep (200)
            DrawFrameControl hDC, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE
        End If

        
        Dim tDock As TDockForm
        Dim tDockItem As Integer
        Dim deltaY As Integer
    
        tDockItem = 1
        For Each tDock In Panel.WindowList
            If tDock Is Me Then
                Panel.DockCollapse tDockItem
            Else
                tDockItem = tDockItem + 1
            End If
        Next
        
        Exit Sub
    End If
    
'*****************************************************
    
    'Set the initial rectangle, and draw it to show the user that
    'the object can be moved
    lXOffset = pt.x - objRect.Left
    lYOffset = pt.Y - objRect.Top
    With DragRect
        .Left = pt.x - lXOffset
        .Top = pt.Y - lYOffset
        .Right = .Left + lObjWidth
        .Bottom = .Top + lObjHeight
    End With
    ' use form border width highlighting
    lBorderWidth = 3
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    ' Move the object
    DockForm_MouseDown pt.x - lXOffset, pt.Y - lYOffset, lObjWidth, lObjHeight, bCancel
    If bCancel = False Then
        Do While GetKeyState(VK_LBUTTON) < 0
            '''debug.Print "moving..."
            GetCursorPos pt
            If pt.x <> ptPrev.x Or pt.Y <> ptPrev.Y Then
                ptPrev.x = pt.x
                ptPrev.Y = pt.Y
                ' erase the previous drag rectangle if any
                DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                'Tell the user we've moved
                lObjLeft = pt.x - lXOffset
                lObjTop = pt.Y - lYOffset
                DockForm_MouseMove lObjLeft, lObjTop, lObjWidth, lObjHeight, lBorderWidth
                ' Adjust the height/width
                With DragRect
                    .Left = lObjLeft
                    .Top = lObjTop
                    .Right = .Left + lObjWidth
                    .Bottom = .Top + lObjHeight
                End With
                DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                bMoved = True
                lXOffset = pt.x - DragRect.Left
                lYOffset = pt.Y - DragRect.Top
            End If
            DoEvents
        Loop
    End If
    ' erase the previous drag rectangle if any
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    ' move and repaint the window
    If bMoved Then
        ' If m_RepositionForm Then
        MoveWindow hwnd, DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
        ' End If
        ' tell the user we've dropped the form
        DockForm_MouseUp DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top
    End If
End Sub

Private Sub DockForm_CaptionClick(ByVal Button As Integer, ByVal x As Single, ByVal Y As Single)
    ' trigger event
    Panel.Parent.TriggerEvent "CaptionClick", Object, Button, x * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    ' When I trigger this event the program locks up
    ' then I discovered that when we send the focus to the main window
    ' the problem is solved. Why? I really don't know
    ' I'm not so good at API the way you may think
    ' but for now it is working fine...
    Putfocus Panel.Parent.Parent.hwnd
End Sub

' ******************************************************************************
' Routine       : (Sub) DrawDragRectangle
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 10/06/2000 - 7:40:49
' Inputs        :
' Outputs       :
' Modifications :
' Credits       : This rotine was extract from the DockForm_Dragger project
'                 by Steve (Steve81@ mediaone.net)
'                 http://www.vbtutor.com
'                 Thanks Steve!
' Modifications : Changed the pen color to vbButtonFace (that's all)
'
' Description   : Draw a rectangle using the Win32 API
' ******************************************************************************
Private Sub DrawDragRectangle(ByVal x As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)
    Dim hDC As Long
    Dim hPen As Long
    
    hPen = CreatePen(PS_INSIDEFRAME, lWidth, vbButtonFace) '&H0) '&HE0E0E0)
    hDC = GetDC(0)
    Call SelectObject(hDC, hPen)
    Call SetROP2(hDC, R2_NOTXORPEN)
    Call Rectangle(hDC, x, Y, X1, Y1)
    Call SelectObject(hDC, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hDC, hPen)
    Call ReleaseDC(0, hDC)
End Sub

' ******************************************************************************
' Routine       : Attach
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:19:32
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : attach window messages to a specific hWnd
' ******************************************************************************
Friend Sub Attach(ByVal hWndA As Long)
    Detach
    m_hWnd = hWndA
    AttachMessage Me, m_hWnd, WM_NCACTIVATE
    AttachMessage Me, m_hWnd, WM_ACTIVATEAPP
    AttachMessage Me, m_hWnd, WM_SIZE
    AttachMessage Me, m_hWnd, WM_MOVE
    AttachMessage Me, m_hWnd, WM_SYSCOMMAND
    AttachMessage Me, m_hWnd, WM_STYLECHANGED
    AttachMessage Me, m_hWnd, WM_CLOSE
    AttachMessage Me, m_hWnd, WM_CONTEXTMENU
    AttachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
    AttachMessage Me, m_hWnd, WM_PAINT
    AttachMessage Me, m_hWnd, WM_ACTIVATE
    AttachMessage Me, m_hWnd, WM_NCHITTEST
    AttachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
End Sub

' ******************************************************************************
' Routine       : Detach
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:19:55
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Detach window messages
' ******************************************************************************
Friend Sub Detach()
    If Not m_hWnd = 0 Then
        DetachMessage Me, m_hWnd, WM_DESTROY
        DetachMessage Me, m_hWnd, WM_SIZE
        DetachMessage Me, m_hWnd, WM_ACTIVATE
        DetachMessage Me, m_hWnd, WM_MOVE
        DetachMessage Me, m_hWnd, WM_SYSCOMMAND
        DetachMessage Me, m_hWnd, WM_STYLECHANGED
        DetachMessage Me, m_hWnd, WM_CLOSE
        DetachMessage Me, m_hWnd, WM_CONTEXTMENU
        DetachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_PAINT
        DetachMessage Me, m_hWnd, WM_NCACTIVATE
        DetachMessage Me, m_hWnd, WM_ACTIVATEAPP
        DetachMessage Me, m_hWnd, WM_NCHITTEST
        DetachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
    End If
End Sub



Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
 '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ' Process before windows:
            'Debug.Print CurrentMessage
    Select Case CurrentMessage

        Case WM_SYSCOMMAND
            ISubclass_MsgResponse = emrConsume
        Case WM_CONTEXTMENU
            ISubclass_MsgResponse = emrConsume
        Case WM_NCRBUTTONDOWN
            ISubclass_MsgResponse = emrConsume
'        Case WM_NCLBUTTONDOWN
'            ISubclass_MsgResponse = emrPreprocess
'        Case WM_ENTERSIZEMOVE
'            ISubclass_MsgResponse = emrPostProcess
'        Case WM_EXITSIZEMOVE
'            ISubclass_MsgResponse = emrConsume
        Case WM_SIZE
            ISubclass_MsgResponse = emrPreprocess
        Case WM_MOVE
            ISubclass_MsgResponse = emrConsume
        Case WM_NCLBUTTONDBLCLK
            ISubclass_MsgResponse = emrConsume
        Case WM_PAINT
            ISubclass_MsgResponse = emrPreprocess
        Case WM_NCHITTEST
            ISubclass_MsgResponse = emrPreprocess
        Case Else
            ISubclass_MsgResponse = emrPreprocess
    End Select
End Property

' ******************************************************************************
' Routine       : ISubClass_WindowProc
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/005:30:36
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Window messages handler for the form
' ******************************************************************************
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pt As POINTAPI
    Dim uCmdType As Long
    Static iRefCount As Long
    
    Dim tP As POINTAPI
    Dim tR As RECT
    Dim lR As Long
    Dim bDoDefault As Boolean
    Dim retVal As Long
    Dim eType As EHitTestAreas

    Select Case iMsg
        Case WM_NCHITTEST
            'Debug.Print "nchittest", lParam, wParam
            'Debug.Print iMsg
            'With m_tHTR
'*******************
            'GetCursorPos tP
            'ScreenToClient hwnd, tP
'            eType = HTCLIENT
            'tR.Left = .lLeft
            'tR.Top = .lTOp
            'tR.Right = .lRight
            'tR.Bottom = .lBottom
            'If (PtInRect(tR, tP.x, tP.y) <> 0) Then
            retVal = CallOldWindowProc(hwnd, iMsg, wParam, lParam) 'HTCAPTION
            'Debug.Print "retval="; retVal
            If retVal = 1 Then
                ISubclass_WindowProc = HTCAPTION
            Else
                ISubclass_WindowProc = retVal
            End If
            'MsgBox SC_CLOSE & "|" & uCmdType & "|" & PushDown
            'End With
        Case WM_SYSCOMMAND
            'Debug.Print "SYSCOMMAND"
            uCmdType = wParam And &HFFF0
            If uCmdType = SC_MOVE Then

'            With m_tHTR
'*******************
'            GetCursorPos tP
'            ScreenToClient hWnd, tP
'            eType = HTCLIENT
'            tR.Left = .lLeft
'            tR.Top = .lTOp
'            tR.Right = .lRight
'            tR.Bottom = .lBottom
            'If (PtInRect(tR, tP.x, tP.y) <> 0) Then
            '        ReleaseCapture
            '        DragObject hWnd
            'End If
'            End With
'***********************
                ReleaseCapture
                DragObject hwnd
            ElseIf uCmdType = SC_CLOSE Then
                m_bUnloaded = True
          
                DockForm_Unload
            Else
                ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            End If
            'MsgBox SC_CLOSE & "|" & uCmdType & "|" & PushDown
       
       Case WM_NCACTIVATE
          'debug.Print "WM_NCACTIVATE"
            If wParam = 0 Then
                iRefCount = iRefCount + 1
                If iRefCount < 2 Then
                    ''debug.Print "Locking"
                    'debug.Print "ncact 1"
                    LockWindowUpdate ParentHWND
                    ISubclass_WindowProc = CallOldWindowProc(ParentHWND, iMsg, wParam, lParam)
                    'CallOldWindowProc m_hWnd, WM_NCACTIVATE, 1, 0
                    LockWindowUpdate 0
                    ''debug.Print "Unlocking"
                    DockForm_Paint
                Else
                    'debug.Print "ncact 2"
                    ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
                    DockForm_Paint
                End If
            Else
                'debug.Print "ncact 3"
                ISubclass_WindowProc = CallOldWindowProc(ParentHWND, iMsg, wParam, lParam)
                DockForm_Paint
          End If
        Case WM_NCLBUTTONDBLCLK
            'Debug.Print "NCLBUTTONDBLCLK"
            
            If wParam = HTCAPTION Then
                DockForm_DblClick
            End If
        Case WM_STYLECHANGED
            If m_bStyleChanged = False Then
                DockForm_ResetStyle
            End If
        Case WM_ACTIVATE
            ' Form is activated/deactivated:
            ''Debug.Print "WM_ACTIVATE"
            If wParam = 0 Then
                ''debug.Print "DEACT"
                iRefCount = 0
                m_bActive = False
                m_bRedraw = True
                DockForm_Deactivate
            Else
                ''debug.Print "ACT"
                m_bActive = True
                DockForm_Activate wParam = 2
            End If
            
        Case WM_ACTIVATEAPP
            'debug.Print "ACTIVATEAPP"
            If (wParam = 0) Then
                iRefCount = 0
                ' app being deactivated
                CallOldWindowProc m_hWnd, WM_NCACTIVATE, 0, 0
                DockForm_Paint
            Else
                ' app being activated
                ' if not the active form then we should repaint
                ' the title bar
                CallOldWindowProc m_hWnd, WM_NCACTIVATE, 1, 0
                DockForm_Paint
            End If
        Case WM_NCLBUTTONDOWN
           ' Debug.Print "ggggg"
           ' DockForm_CaptionClose
        Case WM_NCRBUTTONDOWN
            ''debug.Print "NCRBUTTONDOWN"
            If (wParam = HTCAPTION) Then
                DockForm_CaptionClick vbRightButton, LoWord(lParam), HiWord(lParam)
            End If
        Case WM_CONTEXTMENU
            ''debug.Print "WM_CONTEXTMENU"
        Case WM_CLOSE
            ''debug.Print "WM_CLOSE"
        Case WM_DESTROY
            ''debug.Print "WM_DESTROY"
            Detach
        Case WM_PAINT
            
            ''debug.Print "WM_PAINT"
            DockForm_Paint
            
        Case WM_MOVE
            'debug.Print "WM_MOVE"
            'DockForm_Paint

        Case WM_SIZE
            'debug.Print "WM_SIZE"
            If Me.State = tdUndocked And m_bDragging = False Then
              Me.FloatingWidth = (Me.width \ Screen.TwipsPerPixelX)
              Me.FloatingHeight = (Me.Height \ Screen.TwipsPerPixelY)
            End If
            'DockForm_Paint
        Case Else
            'Debug.Print "imsg"; iMsg
    End Select
End Function


'**************************************
Public Sub SetTestArea( _
      ByVal vKey As Variant, _
      ByVal lLeftPixels As Long, _
      ByVal lTopPixels As Long, _
      ByVal lWidthPixels As Long, _
      ByVal lHeightPixels As Long, _
      ByVal eResponse As EHitTestAreas _
   )
Dim iItem As Long

   iItem = AreaIndex(vKey)
   If (iItem > 0) Then
      With m_tHTR
         .lLeft = lLeftPixels
         .lTOp = lTopPixels
         .lRight = lLeftPixels + lWidthPixels
         .lBottom = lTopPixels + lHeightPixels
         .eResponse = eResponse
      End With
   End If
   
End Sub
Public Sub SetTestAreaFromObject( _
      ByVal vKey As Variant, _
      ByRef objThis As Object, _
      ByVal eResponse As EHitTestAreas _
   )
   SetTestArea vKey, objThis.Left \ Screen.TwipsPerPixelX, objThis.Top \ Screen.TwipsPerPixelY, objThis.width \ Screen.TwipsPerPixelY, objThis.Height \ Screen.TwipsPerPixelY, eResponse
End Sub

Public Function AddArea( _
      ByVal hwnd As Long, _
      Optional ByVal sKey As String = "" _
   ) As Long
      
   m_iHTRCount = m_iHTRCount + 1
   'ReDim Preserve m_tHTR(1 To m_iHTRCount) As tHitTestRectangle
   AddArea = m_iHTRCount
   With m_tHTR
      .hwnd = hwnd
      If (sKey <> "") Then
         .sKey = sKey
      Else
         .sKey = CStr(m_iHTRCount)
      End If
   End With

End Function

Private Function AreaIndex(ByVal vKey As Variant) As Long
Dim i As Long

   If (IsNumeric(vKey)) Then
      i = CLng(vKey)
      If (i > 0) And (i <= m_iHTRCount) Then
         AreaIndex = i
         Exit Function
      End If
   Else

         If (m_tHTR.sKey = vKey) Then
            AreaIndex = i
            Exit Function
         End If

   End If
   
   ' error
End Function

Public Sub AttachHitTester(ByRef objThis As Object)
Dim hwnd As Long
   hwnd = objThis.hwnd
   
   
   
   
      m_lHTPtr = ObjPtr(objThis)
      m_hWndh = hwnd
      AttachMessage Me, m_hWndh, WM_NCHITTEST
      'AttachMessage Me, m_hWndh, WM_SYSCOMMAND
   
End Sub

Public Sub DestroyHitTester()
Dim i As Long

      If (m_hWnd <> 0) Then
         DetachMessage Me, m_hWndh, WM_NCHITTEST
         DetachMessage Me, m_hWndh, WM_SYSCOMMAND
         m_hWndh = 0
      End If
      m_lHTPtr = 0


End Sub


Public Property Get PointInObject(ByRef oThis As Object, ByVal xPixels As Long, ByVal yPixels As Long) As Boolean
Dim tR As RECT
Dim tFR As RECT

   On Error Resume Next
   ' Get object position:
   tR.Left = oThis.Left \ Screen.TwipsPerPixelX
   tR.Right = tR.Left + oThis.width \ Screen.TwipsPerPixelX
   tR.Top = oThis.Top \ Screen.TwipsPerPixelX
   tR.Bottom = tR.Top + oThis.Height \ Screen.TwipsPerPixelY
   If (Err.Number = 0) Then
      ' Check for hit test:
      PointInObject = (PtInRect(tR, xPixels, yPixels) <> 0)
   End If
End Property
   



