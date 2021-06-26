VERSION 5.00
Begin VB.Form login_fenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Daten Fenster"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login_fenster.frx":0000
   ScaleHeight     =   1185
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox username 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      ToolTipText     =   "Hier ist der Benutzername der FTP Accounts einzutragen"
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "login_fenster.frx":6168
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Anonymous-Login Daten generieren"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   5280
      Picture         =   "login_fenster.frx":6FAA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eintrag aus Account Liste löschen"
      Top             =   720
      Width           =   375
   End
   Begin VB.ListBox FTPservers 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   6615
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   5880
      Picture         =   "login_fenster.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eintrag in Account Liste tätigen"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   6480
      Picture         =   "login_fenster.frx":728E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Account Liste laden"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox adresse 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Hier wird die Adresse des FTP Servers eingetragen - z.B.: ""ftp.suse.de"""
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Einloggen mit den eingegeben Daten"
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "x"
      TabIndex        =   2
      ToolTipText     =   "Hier ist das Passwort des obigen FTP Accounts einzutragen"
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Passwort:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "login_fenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PWDatei As String

Private Sub Command1_Click()
With main
    If .comsock.State = 0 Then
        .comsock.RemotePort = 21
        .comsock.RemoteHost = adresse.Text
        .comsock.Connect
        .workofftimer.Enabled = True
        .ListCounter.Caption = ""
        .ListCounter.Visible = True
        .warten = True
        .ParentServerPath = "/"
        .ttl = "LOGIN-------"
        login_fenster.Hide
    End If
End With
End Sub

Public Sub LoginGo()
If MsgBox("Koneksi perintah ditutup! Jika interaksi sedang berlangsung, Anda harus mengulanginya! Apakah Anda ingin masuk lagi?", vbQuestion + vbYesNo, "Pengakhiran sisi server dari koneksi perintah") = vbYes Then
    With main
        .comsock.Close
        .comsock.RemotePort = 21
        .comsock.RemoteHost = adresse.Text
        .comsock.Connect
        .workofftimer.Enabled = True
        .ListCounter.Caption = ""
        .ListCounter.Visible = True
        .warten = True
        .ParentServerPath = "/"
        .ttl = "LOGIN-------"
        login_fenster.Hide
    End With
Else
ttl = ""
main.comsock.Close
End If
End Sub


Private Sub Command2_Click()
Dim DataIntoList As String
FTPservers.Clear
If Len(PWDatei) > 0 Then
    login_fenster.Height = 3500
    file = FreeFile
    Open PWDatei For Binary As #file
    DataIntoList = Space(FileLen(PWDatei))
    Get file, 1, DataIntoList
    Close #file
    w = 1
    For X = 1 To FileLen(PWDatei)
        If Mid$(DataIntoList, X, 4) = vbCrLf & vbCrLf Then
            zwischen = Mid$(DataIntoList, w, X + 1)
            w = X + 4
            zwischen2 = ""
            For Y = 1 To Len(zwischen)
                If Mid$(zwischen, Y, 2) = vbCrLf Then
                    If Len(zwischen2) = 80 Then
                        For z = 0 To Y - v - 1
                            If z Mod 2 = 0 Then
                                zwischen2 = zwischen2 & Chr(Asc(Mid$(zwischen, v + z, 1)) + 5)
                            Else
                                zwischen2 = zwischen2 & Chr(Asc(Mid$(zwischen, v + z, 1)) - 5)
                            End If
                        Next z
                        FTPservers.AddItem zwischen2
                    End If
                                  
                    If Len(zwischen2) = 40 Then
                        For z = 0 To Y - v - 1
                            If z Mod 2 = 0 Then
                                zwischen2 = zwischen2 & Chr(Asc(Mid$(zwischen, v + z, 1)) + 5)
                            Else
                                zwischen2 = zwischen2 & Chr(Asc(Mid$(zwischen, v + z, 1)) - 5)
                            End If
                        Next z
                        Do Until Len(zwischen2) = 80
                            zwischen2 = zwischen2 & " "
                        Loop
                        v = Y + 2
                    End If
                    
                    If zwischen2 = "" Then
                        For z = 1 To Y - 1
                            If z Mod 2 = 0 Then
                                zwischen2 = zwischen2 & Chr(Asc(Mid$(zwischen, z, 1)) - 5)
                            Else
                                zwischen2 = zwischen2 & Chr(Asc(Mid$(zwischen, z, 1)) + 5)
                            End If
                        Next z
                        Do Until Len(zwischen2) = 40
                            zwischen2 = zwischen2 & " "
                        Loop
                        v = Y + 2
                    End If
                End If
            Next Y
        End If
    Next X
End If
        
        
        
End Sub

Private Sub Command3_Click()
Dim datatowrite As String
datatowrite = ""
If adresse.Text <> "" And password.Text <> "" And username.Text <> "" Then
    'Daten verschlüsseln ===========================
    For X = 1 To Len(adresse.Text)
        Y = Asc(Mid$(adresse.Text, X, 1))
        If X Mod 2 = 1 Then
            Y = Y - 5
        Else
            Y = Y + 5
        End If
        datatowrite = datatowrite & Chr(Y)
    Next X
    
    datatowrite = datatowrite & vbCrLf
    
    For X = 1 To Len(username.Text)
        Y = Asc(Mid$(username.Text, X, 1))
        If X Mod 2 = 1 Then
            Y = Y - 5
        Else
            Y = Y + 5
        End If
        datatowrite = datatowrite & Chr(Y)
    Next X
    
    datatowrite = datatowrite & vbCrLf
        
    For X = 1 To Len(password.Text)
        Y = Asc(Mid$(password.Text, X, 1))
        If X Mod 2 = 1 Then
            Y = Y - 5
        Else
            Y = Y + 5
        End If
        datatowrite = datatowrite & Chr(Y)
    Next X
    
    datatowrite = datatowrite & vbCrLf & vbCrLf

    '==========================================
    file = FreeFile
    Open PWDatei For Binary As #file
    If FileLen(PWDatei) = 0 Then
        Put file, 1, datatowrite
    Else
        Put file, FileLen(PWDatei) + 1, datatowrite
    End If
    Close #file
    Command2_Click
End If
End Sub

Private Sub Command4_Click()
Dim zwischenrem As String
Dim zwischenremneu As String
If FTPservers.ListIndex > -1 Then
    X = FTPservers.ListIndex
    file = FreeFile
    Open PWDatei For Binary As #file
    zwischenrem = Space(FileLen(PWDatei))
    Get file, 1, zwischenrem
    Close #file
    Y = -1
    z = 0
    w = 1
    Do Until Y = X
        z = z + 1
        If Mid$(zwischenrem, z, 4) = vbCrLf & vbCrLf Then
            Y = Y + 1
            If Y = X Then Exit Do
            w = z + 4
        End If
    Loop
    zwischenrem = Mid$(zwischenrem, 1, w - 1) & Mid$(zwischenrem, z + 4)

    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.DeleteFile PWDatei

    file = FreeFile
    Open PWDatei For Binary As #file
    Put file, 1, zwischenrem
    Close #file
    Command2_Click
End If
End Sub


Private Sub Command5_Click()
If username.Enabled = True Then
    username.Text = "Anonymous"
    password.Text = "anonymous@email.com"
    username.Enabled = False
    password.Enabled = False
Else
    username.Text = ""
    password.Text = ""
    username.Enabled = True
    password.Enabled = True
End If
End Sub

Private Sub Form_Load()
If GetAttr(CurDir) Mod 2 = 1 Then
    MsgBox "Program dijalankan dari direktori yang dilindungi penulisan dan oleh karena itu tidak dapat membuat atau memuat file kata sandi di dalamnya. Jika perlu, mulai program lagi di direktori dengan akses tulis!", vbInformation, "Direktori saat ini '" & CurDir & "' dilindungi dari penulisan!"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
Else
    If Mid$(CurDir, Len(CurDir), 1) = "\" Then
        PWDatei = CurDir & "accounts.ftp"
    Else
        PWDatei = CurDir & "\accounts.ftp"
    End If
End If
End Sub

Private Sub FTPservers_Click()
adresse.Text = Mid$(FTPservers.List(FTPservers.ListIndex), 1, 40)
username.Text = Mid$(FTPservers.List(FTPservers.ListIndex), 41, 40)
password.Text = Mid$(FTPservers.List(FTPservers.ListIndex), 81)
End Sub


Private Sub password_Change()
Do Until X = Len(password.Text)
    X = X + 1
    If Mid$(password.Text, X, 1) = " " Then
        password.Text = Mid$(password.Text, 1, X - 1) & Mid$(password.Text, X + 1)
        X = X - 1
    End If
Loop
End Sub

Private Sub username_Change()
If Len(username.Text) > 50 Then MsgBox "Panjang maksimum (50 karakter) terlampaui!", vbInformation, "Entri tidak sah!"
Do Until X = Len(username.Text)
    X = X + 1
    If Mid$(username.Text, X, 1) = " " Then
        username.Text = Mid$(username.Text, 1, X - 1) & Mid$(username.Text, X + 1)
        X = X - 1
    End If
Loop
End Sub

Private Sub adresse_Change()
If Len(adresse.Text) > 50 Then MsgBox "Panjang maksimum (50 karakter) terlampaui!", vbInformation, "Entri tidak sah!"
X = 0
Do Until X = Len(adresse.Text)
    X = X + 1
    If Mid$(adresse.Text, X, 1) = " " Then
        adresse.Text = Mid$(adresse.Text, 1, X - 1) & Mid$(adresse.Text, X + 1)
        X = X - 1
    End If
Loop
End Sub
