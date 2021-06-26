VERSION 5.00
Begin VB.Form verzeichnis_erstellen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verzeichnis erstellen"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   1350
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Erstellen"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Bestätigen des Namens des neuen Ordners"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Neuer Ordner"
      ToolTipText     =   "Hier ist der Name des zu erstellenden Ordners einzutragen"
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "verzeichnis_erstellen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
verzeichnis_erstellen.Hide
main.ttl = main.ttl & "MKD DIR-----"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub
