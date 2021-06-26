VERSION 5.00
Begin VB.Form RenameForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Umbenennen des ausgewählten Objektes in"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RenameForm.frx":0000
   ScaleHeight     =   1485
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Bestätigen"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Bestätigen des neuen Namens"
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Hier ist der neue Datei- bzw. Ordnername einzutragen"
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "RenameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
main.ttl = "RENAME_TO---" & main.ttl
RenameForm.Hide
main.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
main.Enabled = True
End Sub
