VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Add/Change"
   ClientHeight    =   8985
   ClientLeft      =   2460
   ClientTop       =   6300
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtDescription 
      Height          =   4695
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Width           =   5775
   End
   Begin VB.ComboBox CmbErgebnis 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   2535
   End
   Begin VB.ComboBox CmbProblem 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtAutor 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox CmbTestTyp 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox CmbMaschine 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdAbort 
      Cancel          =   -1  'True
      Caption         =   "&Chancel"
      Height          =   330
      Left            =   4995
      TabIndex        =   10
      Top             =   8520
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3630
      TabIndex        =   9
      Top             =   8520
      Width           =   1275
   End
   Begin VB.TextBox txtTestNo 
      Height          =   285
      Left            =   2310
      TabIndex        =   4
      Top             =   1620
      Width           =   2265
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   2310
      TabIndex        =   2
      Top             =   600
      Width           =   2265
   End
   Begin VB.Label Label7 
      Caption         =   "Problem Beschreibung:"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Ergebnis:"
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Problem:"
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Autor:"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Test Typ:"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Maschine :"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Test No.:"
      Height          =   225
      Index           =   1
      Left            =   945
      TabIndex        =   11
      Top             =   1650
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   225
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   615
      Width           =   1185
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbort_Click()
  ' Abbrechen
  Me.Tag = False
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  ' OK
  Me.Tag = True
  Me.Hide
End Sub

Public Sub Form_Load()
 
IniRead
cmdOK.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' Schlieﬂen
  If UnloadMode <> 1 Then
    Cancel = True
    cmdAbort.Value = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Beenden
 ' For i = 0 To Me.CmbTestTyp.ListCount - 1
 ' Next i
  
  Set frmInput = Nothing

End Sub


