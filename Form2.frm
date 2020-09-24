VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   3060
      TabIndex        =   4
      Top             =   2940
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "Mp3Playah, Make Music Smaller!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   180
      TabIndex        =   3
      Top             =   1380
      Width           =   4395
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "Proudly Presents"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Productions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2700
      TabIndex        =   1
      Top             =   720
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "[-AciD-]"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   2835
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOk_Click()
Unload Me
End Sub
