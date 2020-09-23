VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form6"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1572
      Left            =   240
      ScaleHeight     =   1575
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   480
      Width           =   3012
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "Form6.frx":0000
         Top             =   120
         Width           =   2412
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alphabetic"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Categorized"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
    MsgBox MDIForm1.TabDock.DockedFormIndex(Me.Name)
End Sub

Private Sub Form_Load()
    Text1.Text = Replace(Text1.Text, vbCrLf, Chr(32))
    Form_Resize
End Sub

Private Sub Form_Resize_old()
    On Error Resume Next
    
    Dim formItem As Integer
    Dim formDockedTopBottom As Boolean
    
    formItem = MDIForm1.TabDock.DockedFormIndex(Me.Name)
    formDockedTopBottom = MDIForm1.TabDock.IsFormDockedTopBottom(Me.Name)
    MsgBox MDIForm1.TabDock.DockedFormCaptionOffset("Form6")
        
    If Not formItem Then
        If formDockedTopBottom Then
            TabStrip1.Move (MDIForm1.TabDock.DockedFormCaptionHeight + 4) * Screen.TwipsPerPixelY, 30, Me.ScaleWidth - ((MDIForm1.TabDock.DockedFormCaptionHeight + 7) * Screen.TwipsPerPixelY), Me.ScaleHeight - 60
        Else
            TabStrip1.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
        End If
    Else
        Debug.Print "no formitem"
    End If
    
    Picture1.Move TabStrip1.Left + 20, _
                TabStrip1.Top + 300, _
                TabStrip1.Width - 50, _
                TabStrip1.Height - 350
    Text1.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    TabStrip1.Move MDIForm1.TabDock.DockedFormCaptionOffsetLeft(Me.Name), MDIForm1.TabDock.DockedFormCaptionOffsetTop(Me.Name), Me.ScaleWidth - MDIForm1.TabDock.DockedFormCaptionOffsetRight(Me.Name), Me.ScaleHeight - MDIForm1.TabDock.DockedFormCaptionOffsetBottom(Me.Name)
    Picture1.Move TabStrip1.Left + 20, _
                TabStrip1.Top + 300, _
                TabStrip1.Width - 50, _
                TabStrip1.Height - 350
    Text1.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight

End Sub

'-- end code
Private Sub Text1_Change()

End Sub

Private Sub Text1_DblClick()
    MsgBox MDIForm1.TabDock.DockedFormIndex(Me.Name)
    MsgBox MDIForm1.TabDock.IsFormDocked(Me.Name)
    MsgBox MDIForm1.TabDock.DockedFormCaptionHeight
End Sub
