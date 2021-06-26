VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form analyse_fenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konsolenfenster zur Fehleranalyse"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "analyse_fenster.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   4800
      Top             =   5160
   End
   Begin VB.TextBox console 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      ToolTipText     =   "Hier werden die angekommenen / ausgehenden Daten protokolliert - ein Doppelklick löscht die Liste wieder"
      Top             =   600
      Width           =   10695
   End
   Begin VB.TextBox manualcom 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Command"
      ToolTipText     =   "Hier können manuel Befehle an den Sever versandt werden"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox timertodotext 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Ausständige Threads"
      Top             =   120
      Width           =   5535
   End
   Begin VB.CheckBox struktur 
      Caption         =   "Struktur"
      Height          =   255
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   $"analyse_fenster.frx":8C16
      Top             =   120
      Value           =   1  'Checked
      Width           =   735
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      ToolTipText     =   "Hier kann die Geschwindigkeit des Programmes geregelt werden"
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   100
      Min             =   1
      Max             =   3000
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "350"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   8880
      TabIndex        =   14
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 350 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H000000FF&
      Caption         =   "warten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   10080
      TabIndex        =   13
      ToolTipText     =   "Grün, wenn die Wartenvariable den Wert ""false"" hat, also wenn keine Aktionen noch durchzuführen sind"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H000000FF&
      Caption         =   "257"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 257 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Caption         =   "331"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 331 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H000000FF&
      Caption         =   "250"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 250 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Caption         =   "230"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 230 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "226"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 226 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "220"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 220 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Grün, wenn zuletzt angekommener Bestätigungscode 200 war"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label speed 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   10200
      TabIndex        =   5
      ToolTipText     =   "Hier kann die Geschwindigkeit des Programmes abgelesen werden"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "analyse_fenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub console_DblClick()
console.Text = ""
End Sub

Private Sub manualcom_Click()
manualcom.Text = ""
End Sub

Private Sub manualcom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If manualcom.Text = "cls" Then
        console.Text = ""
        manualcom.Text = ""
        Exit Sub
    End If
    If manualcom.Text = "refresh" Then
        main.ttl = main.ttl & "REFRESH-----"
        manualcom.Text = ""
        Exit Sub
    End If
    If main.comsock.State = 7 Then
    main.comsock.SendData manualcom.Text & vbCrLf
    main.WriteToConsole ("O U T G O I N G" & vbCrLf & manualcom.Text)
    End If
manualcom.Text = ""
End If
End Sub

Private Sub Slider1_Change()
main.workofftimer.Enabled = False
main.workofftimer.Interval = Slider1
speed.Caption = Slider1 & " ms"
main.workofftimer.Enabled = True
End Sub
