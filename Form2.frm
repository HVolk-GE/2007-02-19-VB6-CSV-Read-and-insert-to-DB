VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "For MySQL Connect"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   ScaleHeight     =   3345
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DB Close"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DB Open"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ConnectMySQL
End Sub

Private Sub Command2_Click()
CloseMySQLConn
End Sub
