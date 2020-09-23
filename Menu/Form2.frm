VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   LinkTopic       =   "Form2"
   ScaleHeight     =   1590
   ScaleWidth      =   1170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   30
      ScaleHeight     =   1530
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   30
      Width           =   1095
      Begin VB.Line Line1 
         BorderColor     =   &H00FF8080&
         X1              =   0
         X2              =   1320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image4 
         Height          =   225
         Left            =   0
         Picture         =   "Form2.frx":0000
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   270
         Left            =   0
         Picture         =   "Form2.frx":0312
         Top             =   720
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   0
         Picture         =   "Form2.frx":06FC
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   0
         Picture         =   "Form2.frx":0A6E
         Top             =   0
         Width           =   210
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Height          =   2175
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   765
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ForCounter As Byte

Public Function ForBorder(Obj As Object) As Object
On Error Resume Next
For Each ctl In Controls
 ctl.BorderStyle = 0
Next
 Obj.BorderStyle = 1
 Line1.BorderStyle = 1
End Function
Public Function ForRestore()
For Each ctl In Controls
 If TypeOf ctl Is Label Then
  ctl.BorderStyle = 0
 End If
Next
ForCounter = 0
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ForRestore
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ForRestore
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ForRestore
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ForRestore
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ForRestore
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ForRestore
End Sub

Private Sub Label2_Click()
Me.Hide
Call ForRestore
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ForCounter = 1 Then Exit Sub
Call ForBorder(Label2)
ForCounter = 1
End Sub

Private Sub Label3_Click()
Me.Hide
Call ForRestore
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ForCounter = 1 Then Exit Sub
Call ForBorder(Label3)
ForCounter = 1
End Sub

Private Sub Label4_Click()
Me.Hide
Call ForRestore
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ForCounter = 1 Then Exit Sub
Call ForBorder(Label4)
ForCounter = 1
End Sub

Private Sub Label5_Click()
Me.Hide
Call ForRestore
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ForCounter = 1 Then Exit Sub
Call ForBorder(Label5)
ForCounter = 1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ForRestore
End Sub
