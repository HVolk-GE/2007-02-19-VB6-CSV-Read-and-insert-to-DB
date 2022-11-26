VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Insert Programm Number"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4185
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "None"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
VERSUCH = Me.Text1.Text
Me.Text1.Text = ""
Form3.Hide
Form2.Show

End Sub

Private Sub Form_Load()
    Me.Text1.Text = ""
    Me.Label2.Caption = "Please, have not found Programm Number, " & vbCr & "insert here :"
    Me.Label1.Caption = "Insert Programm Number : "
End Sub


