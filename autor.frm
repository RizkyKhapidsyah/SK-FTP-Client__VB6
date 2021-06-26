VERSION 5.00
Begin VB.Form autor 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FTP Client"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "autor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "=>"
      Height          =   315
      Left            =   5520
      MaskColor       =   &H00000000&
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6300
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   5775
   End
End
Attribute VB_Name = "autor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
autor.Hide
main.Show
End Sub

Private Sub Form_Load()
Label1.Caption = "" & vbCrLf & vbCrLf & "Contact me at:" & vbCrLf & "" & vbCrLf & ""
End Sub
