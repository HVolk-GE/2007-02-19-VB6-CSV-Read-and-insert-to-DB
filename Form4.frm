VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Dynamometers Problem report !"
   ClientHeight    =   7275
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   10800
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10800
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9240
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   9120
      Picture         =   "Form4.frx":0442
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdDropTable 
      Caption         =   "Delete Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   8
      Top             =   2655
      Width           =   1590
   End
   Begin VB.CommandButton cmdQueryTable 
      Caption         =   "Query"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   7
      Top             =   3075
      Width           =   1590
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   6
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   5
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   4
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton cmdCreateTable 
      Caption         =   "Create Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   3
      Top             =   2235
      Width           =   1590
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   2
      Top             =   1500
      Width           =   1590
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9135
      TabIndex        =   1
      Top             =   1080
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5835
      Left            =   210
      TabIndex        =   0
      Top             =   840
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   10292
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label lblURL 
      Caption         =   "www.little-tools-farm.de"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   7920
      MouseIcon       =   "Form4.frx":0A49
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   11
      Top             =   7020
      Width           =   2385
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright ©2007 by little-tools-farm.de"
      Height          =   225
      Index           =   0
      Left            =   7920
      TabIndex        =   10
      Top             =   6765
      UseMnemonic     =   0   'False
      Width           =   2760
   End
   Begin VB.Label lblWelcome 
      Caption         =   "Welcome Dynamometers Problem Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   105
      Picture         =   "Form4.frx":0D53
      Stretch         =   -1  'True
      Top             =   105
      Width           =   555
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   210
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   210
      X2              =   5775
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================

