VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000D&
   Caption         =   "Form7"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Rearrange"
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   2652
      Left            =   2280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Form7.frx":165EA
      Top             =   60
      Width           =   1692
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long


Private Sub Command1_Click()
    ' move form6 to left
    ' move form2 to top
    
    
    MDIForm1.TabDock.FormUndock "Form6"
    MDIForm1.TabDock.DockedForms.Remove "Form6"
    MDIForm1.TabDock.FormUndock "Form2"
    MDIForm1.TabDock.DockedForms.Remove "Form2"
    
    
    MDIForm1.TabDock.AddForm Form6, tdDocked, tdAlignLeft, "Form6", tdDockLeft Or tdDockRight Or tdDockTop Or tdDockBottom Or tdDockFloat
    MDIForm1.TabDock.AddForm Form2, tdDocked, tdAlignTop, "Form2", tdDockLeft Or tdDockRight Or tdDockTop Or tdDockBottom Or tdDockFloat
    MDIForm1.TabDock.Show
    
    
    ' Host Panels (1=left, 2=top, 3=right, 4=bottom)

    'MDIForm1.TabDock.Panels("Host1").DockArrange
    'MDIForm1.TabDock.Panels("Host2").DockArrange
    
    'LockWindowUpdate ByVal 0&

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Width = Me.ScaleWidth - (Text1.Left + 30)
End Sub
'-- end code
