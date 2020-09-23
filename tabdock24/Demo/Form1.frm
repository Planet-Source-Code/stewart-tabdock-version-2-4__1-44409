VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   3180
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1572
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3132
      _ExtentX        =   5530
      _ExtentY        =   2778
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483641
      BackColor       =   -2147483639
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Items"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Add some item to listview
    With ListView1.ListItems
        .Add , , "Item 1"
        .Add , , "Item 2"
        .Add , , "Item 3"
        .Add , , "Item 4"
    End With
End Sub

Private Sub Form_Resize_old()

    On Error Resume Next
    
    Dim formItem As Integer
    Dim formDockedTopBottom As Boolean
    
    formItem = MDIForm1.TabDock.DockedFormIndex(Me.Name)
    formDockedTopBottom = MDIForm1.TabDock.IsFormDockedTopBottom(Me.Name)
    
        
    If Not formItem Then
        If formDockedTopBottom Then
            ListView1.Move (MDIForm1.TabDock.DockedFormCaptionHeight + 4) * Screen.TwipsPerPixelY, 30, Me.ScaleWidth - ((MDIForm1.TabDock.DockedFormCaptionHeight + 7) * Screen.TwipsPerPixelY), Me.ScaleHeight - 60
        Else
            ListView1.Move 10, 30, Me.ScaleWidth, Me.ScaleHeight
        

        End If
    End If
    
End Sub
'-- end code

Private Sub Form_Resize()

    On Error Resume Next
    
    ListView1.Move MDIForm1.TabDock.DockedFormCaptionOffsetLeft(Me.Name), MDIForm1.TabDock.DockedFormCaptionOffsetTop(Me.Name), Me.ScaleWidth - MDIForm1.TabDock.DockedFormCaptionOffsetRight(Me.Name), Me.ScaleHeight - MDIForm1.TabDock.DockedFormCaptionOffsetBottom(Me.Name)
    
End Sub

