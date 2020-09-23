VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Beauty Menu"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Unload Form2
End Sub

Private Sub Form_Click()
Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
 For intctr = (Forms.Count - 1) To 0 Step -1
  Unload Forms(intctr)
 Next intctr
End Sub

Private Sub Label1_Click()
Form2.Show
Form2.Top = Form1.Top + 670
Form2.Left = Form1.Left + 185
End Sub
