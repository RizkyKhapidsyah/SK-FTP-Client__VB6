VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "Cswsk32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP Client"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket sockdata 
      Left            =   9240
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer waiter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   0
   End
   Begin VB.TextBox text_serverpath 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Gegenwärtiger Serverpfad - durch Editieren und abschließendes Enter kann in ein anderes Verzeichnis gewechselt werden"
      Top             =   415
      Width           =   8410
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   200
      Left            =   170
      TabIndex        =   0
      ToolTipText     =   "Markieren aller Listenelemente"
      Top             =   480
      Width           =   200
   End
   Begin VB.CommandButton Command1 
      Height          =   3800
      Left            =   11160
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Laufwerke des lokalen Dateisystems"
      Top             =   3840
      Width           =   4935
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   5160
      MultiSelect     =   2  'Extended
      System          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Lokale Dateien vom obigen Verzeichnis - mit der Taste D löschen Sie das ausgewählte Objekt"
      Top             =   3840
      Width           =   5655
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Verzeichnisbaum des lokalen Dateisystems - mit der Taste D löschen Sie das ausgewählte Objekt"
      Top             =   4200
      Width           =   4935
   End
   Begin VB.ListBox remotelist 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      ItemData        =   "Form1.frx":1B8A
      Left            =   120
      List            =   "Form1.frx":1B8C
      Style           =   1  'Checkbox
      TabIndex        =   2
      ToolTipText     =   "Die Dateien im gegenwärtigem Verzeichnis auf dem FTP Server"
      Top             =   720
      Width           =   11655
   End
   Begin MSWinsockLib.Winsock datasock 
      Left            =   9720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer workofftimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8760
      Top             =   0
   End
   Begin MSWinsockLib.Winsock comsock 
      Left            =   10200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   688
      ButtonWidth     =   767
      ButtonHeight    =   688
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "PicTBnormal(1)"
      DisabledImageList=   "PicTBnormal(1)"
      HotImageList    =   "PicTBdown"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Login"
            Object.ToolTipText     =   "Verbindung zum FTP Server herstellen..."
            ImageIndex      =   1
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DirUp"
            Object.ToolTipText     =   "Ins höhere Verzeichnis wechseln..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Laufende Operation stoppen..."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Dateiliste aktualisieren..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Down"
            Object.ToolTipText     =   "Datei downloaden..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            Object.ToolTipText     =   "Datei uploaden..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MkdDir"
            Object.ToolTipText     =   "Verzeichnis erstellen..."
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dele"
            Object.ToolTipText     =   "Objekt löschen..."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rename"
            Object.ToolTipText     =   "Objekt umbenennen..."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Analyse"
            Object.ToolTipText     =   "Analyse Fenster anzeigen..."
            ImageIndex      =   10
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "author"
            Object.ToolTipText     =   "Autor"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList PicTBdown 
      Left            =   10680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2132
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":321E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":37C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":430A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":53F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProcessProgress 
      Height          =   3795
      Left            =   10920
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   6694
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
   End
   Begin MSComctlLib.ImageList PicTBnormal 
      Index           =   1
      Left            =   11280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5972
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":64BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7002
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":75A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8692
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":91DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label DateigroesseLocalVisual 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   7800
      Width           =   4335
   End
   Begin VB.Label ListCounter 
      BackStyle       =   0  'Transparent
      Caption         =   "ListCounter"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   21
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label VD_ip 
      BackStyle       =   0  'Transparent
      Caption         =   "RemoteIP:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   7800
      Width           =   2655
   End
   Begin VB.Label VD_comstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Comsock Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Label VD_datastatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Datasock Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   8040
      Width           =   3495
   End
   Begin VB.Label VD_bytesReceived 
      BackStyle       =   0  'Transparent
      Caption         =   "\/ Bytes:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label VD_bytesSent 
      BackStyle       =   0  'Transparent
      Caption         =   "/\ Bytes:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   8280
      Width           =   2655
   End
   Begin VB.Label VD_sockstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Sockdata Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   8280
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aktueller Stand:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dateigroesse:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   13
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label ProcessAktStand 
      BackStyle       =   0  'Transparent
      Caption         =   "1 000 000 000 000"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label ProcessDateigroesse 
      BackStyle       =   0  'Transparent
      Caption         =   "1 000 000 000 000"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10320
      TabIndex        =   11
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label ProcessDateiname 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label label_serverpfad 
      BackStyle       =   0  'Transparent
      Caption         =   "Direktori saat ini:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   9855
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Zwischenport As String
Dim zwischen_download As String
Dim zwischen_WriteList As Variant
Dim zwischen_List As Variant
Dim comgot_verzögert As String
Dim dirpos(256) As Variant
Dim FileNamesLocal As String
Dim FileNamesServer As String
Dim DirsToCreate As String
Dim waitvar As Boolean
Dim DataToUpload As String
Dim zwischen_ServerPath As String
Dim Zwischen_Uppen_Bytes As String
Dim zwischenDg_Up As Variant
Dim zwischenDg_Down As Variant
Dim file As Long
Dim Currently_Uploading As Boolean
Dim DateiSchonOffen As Boolean
Dim RenameObject As String
Dim UpOrDown As String
Dim ShowDLProgress As Boolean
Dim waitervar As Integer

Dim Disconnected As Boolean
Dim PWD_Befehl As Boolean

Public Use_Datasock As Boolean
Public ParentServerPath As String
Public HideWarningsForBadNames As Boolean
Public HideWarningsForMultipleLocalFiles As Boolean
Public HideWarningsForMultipleLocalFolders As Boolean
Public CheckedString As String
Public consoleText As String
Public warten As Boolean
Public comgot_warten As Boolean
Public timertodo As String
Public ttl As String
Public datagot As String
Public comgot As String
Public filegot As String
Public whatdata As String
Public wholedata As String
Public DownloadFilename As String
Public UploadFilename As String
Public BytesAngekommenTotal As Double
Public BytesAngekommenDatei As Double
Public BytesAngekommenTotal_c As Double
Public BytesGesendetTotal_c As Double
Public BytesGesendetTotal As Double
Public Dateigroesse As Double
Public deletedaten As String
Public ZielPfad As String
Public RichtigerDateiname As String
Public save_as As String

Public markedfiles As String
Public subdir As String
Public subdircount As String

Public GeordneteZahl As String


Private Sub Check1_Click()
If Check1.Value = 1 Then
    For X = 0 To remotelist.ListCount - 1
        remotelist.Selected(X) = True
    Next X
End If
If Check1.Value = 0 Then
    For X = 0 To remotelist.ListIndex - 1
        remotelist.Selected(X) = False
    Next X
End If
End Sub

Private Sub Command1_Click()
If MsgBox("Wollen Sie den Dateitransfer wirklich abbrechen?", vbYesNo + vbQuestion, "Abbruch des Dateitransfers") = vbYes Then
    If ShowDLProgress Then StartProgressVisualisation ("")
    sockdata.Disconnect
    warten = False
    ttl = "REFRESH-----"
    Close #file
    If datasock.State = 7 Then ttl = "ABORT_DELETE" & ttl
    datasock.Close
End If
End Sub


Private Sub comsock_ConnectionRequest(ByVal requestID As Long)
If comsock.State <> sckClosed Then comsock.Close
comsock.Accept (requestID)
End Sub


Private Sub comsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Folgender Comsock Error trat auf: " & Number & "     " & Description & "     " & Scode & "     " & Source
End Sub

Private Sub comsock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
BytesGesendetTotal_c = BytesGesendetTotal_c + bytesSent
End Sub

Private Sub comsock_DataArrival(ByVal bytesTotal As Long)
If comsock.State = 7 Then
    If comgot_warten = True Then
        comsock.GetData comgot_verzögert
        comgot = comgot & comgot_verzögert
    End If
    If comgot_warten = False Then comsock.GetData comgot
    comgot_warten = True
    For X = 1 To Len(comgot) - 1
        If Mid$(comgot, X, 2) = vbCrLf Then comgot_warten = False
    Next X
    If comgot_warten = True Then Exit Sub
Else
    MsgBox "Datenankunft auf Comport, obwohl dieser nicht offen ist!", vbInformation, "?!?!?"
End If

WriteToConsole ("I N C O M I N G" & vbCrLf & Mid$(comgot, 1, Len(comgot) - 2))

BytesAngekommenTotal_c = BytesAngekommenTotal_c + bytesTotal

X = 0
zwischen = comgot
Do Until zwischen = ""
    
    If (Mid$(zwischen, 1, 4) = "226 ") And (whatdata = "listneu") Then
        If datagot = "" Then
             waitervar = 0
             waiter.Enabled = True
        End If
    End If
    
    If Mid$(zwischen, 1, 4) = "200 " Or Mid$(zwischen, 1, 4) = "220 " Or Mid$(zwischen, 1, 4) = "226 " Or Mid$(zwischen, 1, 4) = "230 " Or Mid$(zwischen, 1, 4) = "250 " Or Mid$(zwischen, 1, 4) = "257 " Or Mid$(zwischen, 1, 4) = "331 " Or Mid$(zwischen, 1, 4) = "350 " Then
        warten = False
        If Mid$(zwischen, 1, 4) = "257 " And PWD_Befehl = True Then
            PWD_Befehl = False
            Dim anfang, ende As Integer
            anfang = -1
            Dim zz As Integer
            For zz = 1 To Len(zwischen)
                If Asc(Mid$(zwischen, zz, 1)) = 34 Then
                    If anfang = -1 Then
                        anfang = zz + 1
                    Else
                        ende = zz
                    End If
                End If
            Next zz
            ParentServerPath = Mid$(zwischen, anfang, ende - anfang)
        End If
        
        
        WriteToConsole ("LDR#" & Mid$(zwischen, 1, 3))
        If Mid$(zwischen, 1, 4) = "250 " And zwischen_ServerPath <> "" Then ServerpfadNeuDefinieren "//DirStatementOk"
    End If
            
    'Fehler
    If Mid$(zwischen, 1, 1) = "5" Then
        timertodo = ""
        ttl = ""
        If comsock.State = 7 Then warten = False
        sockdata.Disconnect
        If ShowDLProgress Then StartProgressVisualisation ("")
        If Mid$(comgot, 1, 3) = "500" Then MsgBox "ERROR COMMAND UNRECOGNIZED" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "501" Then MsgBox "SYNTAX ERRORR IN PARAMETERS OR ARGUMENTS" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "502" Then MsgBox "COMMAND NOT IMPLEMENTED" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "503" Then MsgBox "BAD SEQUENCE OF COMMANDS" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "504" Then MsgBox "COMMAND NOT IMPLEMENTED FOR THAT PARAMETER" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "532" Then MsgBox "NEED ACCOUNT FOR STORING FILES" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "550" Then MsgBox "REQUESTED ACTION NOT TAKEN FILE UNAVAILABLE" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "551" Then MsgBox "REQUESTED ACTION ABORTED PAGE TYPE UNKNOWN" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "552" Then MsgBox "REQUESTED FILE ACTION ABORTED EXCEEDED STORAGE ALLOCATION" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        If Mid$(comgot, 1, 3) = "553" Then MsgBox "REQUESTED ACTION NOT TAKEN FILE NAME NOT ALLOWED" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
        ttl = ttl & "REFRESH-----"
        If Mid$(comgot, 1, 3) = "530" Then
            MsgBox "NOT LOGGED IN" & vbCrLf & comgot, vbCritical, "Aktion abgebrochen - Fehler"
            ttl = ""
        End If
    End If
    
    zwischen = "?'?"
    Do Until zwischen <> "?'?"
        X = X + 1
        If X < Len(comgot) Then
            If Asc(Mid$(comgot, X, 1)) = 13 And Asc(Mid$(comgot, X + 1, 1)) = 10 Then
                zwischen = Mid$(comgot, X + 2)
            End If
        Else
            Exit Sub
        End If
    Loop
Loop
End Sub


Private Sub datasock_ConnectionRequest(ByVal requestID As Long)
If datasock.State <> 0 Then datasock.Close
datasock.Accept requestID
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then Dir1.Refresh
If KeyCode = 46 Then
    If MsgBox("Wollen Sie den ausgewählten, lokalen Ordner '" & Dir1.Path & "' wirklich löschen?", vbYesNo + vbQuestion, "Löschvorgang bestätigen") = vbYes Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Mid$(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
            zwischenpfad = Mid$(Dir1.Path, 1, Len(Dir1.Path) - 1)
        Else
            zwischenpfad = Dir1.Path
        End If
        If Mid$(Dir1.Path, Len(Dir1.Path) - 1, 2) <> ":\" Then
            fs.deletefolder (zwischenpfad)
            Dir1.Path = Dir1.List(-2)
        Else
            MsgBox "Das ausgewählte, zu löschende Objekt ist offenbar eine Festplatte - diese können nicht gelöscht werden!", vbInformation, "Fehlerhafte Eingabe"
        End If
        Dir1.Refresh
    End If
End If
End Sub

Private Sub File1_Click()
Dim GesamtGroesse As Double
Dim PathString As String
GesamtGroesse = 0
If Mid$(Dir1.Path, Len(Dir1.Path) - 1, 1) = ":" Then
    PathString = Dir1.Path
Else
    PathString = Dir1.Path & "\"
End If

For X = 0 To File1.ListCount - 1
    If File1.Selected(X) = True Then
        GesamtGroesse = GesamtGroesse + FileLen(PathString & File1.List(X))
    End If
Next X
ZahlOrdnen (GesamtGroesse)
DateigroesseLocalVisual.Caption = "Größe der ausgewählten Datei(en): " & GeordneteZahl
End Sub


Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then File1.Refresh
If KeyCode = 46 Then
    If MsgBox("Wollen Sie die ausgewählte, lokale Datei '" & File1.FileName & "' wirklich löschen?", vbYesNo + vbQuestion, "Löschvorgang bestätigen") = vbYes Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Mid$(File1.Path, Len(File1.Path), 1) = "\" Then
            zwischenpfad = Mid$(File1.Path, 1, Len(File1.Path) - 1)
        Else
            zwischenpfad = File1.Path
        End If
        fs.DeleteFile (zwischenpfad & "\" & File1.FileName)
        File1.Refresh
    End If
End If
End Sub

Private Sub ProcessProgress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim zwischenwert As Integer
zwischenwert = ProcessProgress / ProcessProgress.Max * 100
ProcessProgress.ToolTipText = zwischenwert & "%"
End Sub


Private Sub remotelist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    If comsock.State = 7 Then
        If warten = False Then
            ttl = ttl & "REFRESH-----"
        Else
            MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
        End If
    Else
        MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
    End If
End If

If KeyCode = 46 Then
    If MsgBox("Wollen Sie die ausgewählte(n) Datei(en) wirklich löschen?", vbQuestion + vbYesNo, "Bestätigung erforderlich") = vbYes Then
        If comsock.State = 7 Then
            If warten = False Then
                deletedaten = ""
                subdir = "/"
                For X = 0 To remotelist.ListCount - 1
                    If remotelist.Selected(X) = True Then
                        If Mid$(remotelist.List(X), 62) <> "~" Then
                            deletedaten = deletedaten & Mid$(remotelist.List(X), 1, 4) & Mid$(remotelist.List(X), 62) & Chr(6)
                        Else
                            MsgBox "In das Verzeichnis '~' kann nicht gewechselt werden. Dieser Bug ist serverseitig zu begründen!", vbCritical, "Serverseitiger Fehler"
                        End If
                    End If
                Next X
                If deletedaten <> "" Then main.ttl = main.ttl & "DELETE------"
            Else
                MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
            End If
        Else
            MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
        End If
    End If
End If
End Sub

Private Sub sockdata_Disconnect()
Disconnected = True
End Sub

Private Sub sockdata_Read(DataLength As Integer, IsUrgent As Integer)
BytesAngekommenTotal = BytesAngekommenTotal + DataLength 'Anzeige der erhaltenen Bytes in einer Session
sockdata.Read datagot, DataLength

    If whatdata = "listneu" Then
        wholedata = wholedata & datagot
        WriteToConsole ("I N C O M I N G   F I L E   L I S T   P A C K A G E   -   S I Z E : " & DataLength)
        If Asc(Mid$(datagot, Len(datagot) - 1, 1)) & " " & Asc(Mid$(datagot, Len(datagot), 1)) = "13 10" Then
            ttl = "WRITELIST---" & ttl 'ENDE DES DATEILISTENTRANSFERS
            waiter.Enabled = False
            whatdata = ""
            datagot = ""
        End If
    End If
        
    If whatdata = "file" Then
        filegot = filegot & datagot
        BytesAngekommenDatei = BytesAngekommenDatei + DataLength
        RefreshForm
        WriteToConsole ("I N C O M I N G   F I L E   P A C K A G E   -   S I Z E :   " & DataLength & " / " & BytesAngekommenDatei & " / " & Dateigroesse)
        
        ZahlOrdnen (BytesAngekommenDatei)
        
        ProcessAktStand.Caption = GeordneteZahl
        
        If (BytesAngekommenDatei / Dateigroesse) * 100 < 100 Then ProcessProgress = (BytesAngekommenDatei / Dateigroesse) * 100
        
        If Len(filegot) > 65536 Then
            If DateiSchonOffen = False Then
                RichtigerDateiname = CheckedString
                For X = Len(CheckedString) To 1 Step -1
                    If Mid$(CheckedString, X, 1) = "." Then
                        CheckedString = Mid$(CheckedString, 1, X - 1) & ".xxx"
                    End If
                Next X
            
                file = FreeFile
                Open ZielPfad & CheckedString For Binary As #file
                Put file, , filegot
                zwischenDg_Down = Len(filegot) + 1
                DateiSchonOffen = True
            Else
                Put file, zwischenDg_Down, filegot
                zwischenDg_Down = zwischenDg_Down + Len(filegot)
            End If
            filegot = ""
        End If
        If (Asc(Mid$(datagot, Len(datagot) - 1, 1)) & " " & Asc(Mid$(datagot, Len(datagot), 1)) = "13 10") Or (BytesAngekommenDatei = Dateigroesse) Then ttl = "SAVEDOWNLOAD" & ttl
    End If
sockdata.Flush
End Sub



Private Sub datasock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
datasock.Close
generateport
End Sub

Private Sub datasock_SendComplete()
If Currently_Uploading Then
    FileIsBeingSent
Else
    datasock.Close
    warten = False
    If ShowDLProgress Then StartProgressVisualisation ("")
    BytesAngekommenDatei = 0
    If FileNamesLocal <> "" Then
        FilesAufServerUploaden
    Else
        ttl = ttl & "REFRESH-----"
    End If
End If
End Sub

Private Sub datasock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
BytesGesendetTotal = BytesGesendetTotal + bytesSent

If Dateigroesse > 0 Then
    Zwischen_Uppen_Bytes = Zwischen_Uppen_Bytes + bytesSent
    If Dateigroesse > 1048576 Then
        ProcessAktStand.Caption = zwischenDg_Up - 1048576 - 1 + Zwischen_Uppen_Bytes
    Else
        ProcessAktStand.Caption = Zwischen_Uppen_Bytes
    End If
    
    ProcessProgress = (ProcessAktStand.Caption) / Dateigroesse * 100
    WriteToConsole ("O U T G O I N G   F I L E   P A C K A G E   -   S I Z E :   " & bytesSent & " / " & ProcessAktStand.Caption & " / " & Dateigroesse)
    
    Y = 1
    For X = Len(ProcessAktStand.Caption) To 1 Step -1
        If (Len(ProcessAktStand.Caption) - X + Y) Mod 3 = 0 Then
            Y = Y - 1
            ProcessAktStand.Caption = Mid$(ProcessAktStand.Caption, 1, X - 1) & " " & Mid$(ProcessAktStand.Caption, X)
        End If
    Next X
    Do Until Len(ProcessAktStand.Caption) > 15
        ProcessAktStand.Caption = " " & ProcessAktStand.Caption
    Loop
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo errorhandler
    Dir1.Path = Drive1.Drive
    Exit Sub
errorhandler:
    MsgBox "Gerät oder Medium nicht verfügbar", vbInformation, "Fehler!"
    Drive1.Drive = Mid$(Dir1.Path, 1, 2)
End Sub

Private Sub Form_Load()
logoff
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub remotelist_DblClick() 'EINZELDDOWNLOAD START
If Mid$(remotelist.List(remotelist.ListIndex), 1, 4) = " DIR" Then 'Verzeichnis => In diese wechseln
    If Mid$(remotelist.List(remotelist.ListIndex), 62) <> "~" Then
        If warten = False Then
            ListCounter.Caption = ""
            warten = True
            ServerpfadNeuDefinieren Mid$(remotelist.List(remotelist.ListIndex), 62)
            comsock.SendData "CWD " & Mid$(remotelist.List(remotelist.ListIndex), 62) & vbCrLf
            WriteToConsole ("O U T G O I N G" & vbCrLf & "CWD " & Mid$(remotelist.List(remotelist.ListIndex), 62))
            ttl = ttl & "REFRESH-----"
        End If
    Else
        MsgBox "In das Verzeichnis mit dem Anfangszeichen '~' kann nicht gewechselt werden. Dieser Bug ist serverseitig zu begründen!", vbCritical, "Serverseitiger Fehler!"
    End If
End If
End Sub


Private Sub sockdata_Accept(SocketId As Integer)
sockdata.Accept = SocketId
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Login"
        If comsock.State = 7 Then
            If MsgBox("Durch Öffnen einer Verbindung zu einem FTP Server wird die alten Verbindung abgebrochen." & vbCrLf & "Wollen Sie wirklich fortfahren?", vbYesNo + vbQuestion, "Alte Verbindung schließen") = vbYes Then
                logoff
                login_fenster.Show
                Exit Sub
            End If
            Exit Sub
        End If
        logoff
        login_fenster.Show
        
    Case "DirUp"
        If comsock.State = 7 Then
            If warten = False Then
                warten = True
                comsock.SendData "CWD .." & vbCrLf
                ServerpfadNeuDefinieren "//DirUp"
                WriteToConsole ("O U T G O I N G" & vbCrLf & "CWD ..")
                ttl = ttl & "REFRESH-----"
            Else
                MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
            End If
        Else
            MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
        End If
        
    Case "Cancel"
        If warten = True Then
            If MsgBox("Wollen Sie die laufende Operation wirklich abbrechen?", vbYesNo + vbQuestion, "Abbrechen bestätigen") = vbYes Then
                    If ShowDLProgress Then StartProgressVisualisation ("")
                warten = False
                ProcessProgress.Visible = False
                Command1.Visible = False
                warten = False
                datasock.Close
                sockdata.Disconnect
                ttl = ""
            End If
        Else
            MsgBox "Keine Operation im Gange!", vbInformation, "Cancel"
        End If
        
    Case "Refresh"
        If comsock.State = 7 Then
            If warten = False Then
                ttl = ttl & "REFRESH-----"
            Else
                MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
            End If
        Else
            MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
        End If
        
    Case "Down"
        If comsock.State = 7 Then
            Dim test As Boolean
            For X = 0 To remotelist.ListCount - 1
                If remotelist.Selected(X) = True Then test = True
            Next X
            If test = False Then
                MsgBox "Sie haben keine Objekte zum Downloaden ausgewählt. Die Aktion wird abgebrochen!", vbInformation, "Ausführung nicht möglich!"
                Exit Sub
            End If
            If GetAttr(Dir1.Path) Mod 2 = 1 Then 'Schreibgeschützt
                MsgBox "Sie besitzen im Zielverzeichnis über keinerlei Schreibrechte! Wählen Sie ein anderes Verzeichnis aus. Die Aktion wird abgebrochen!", vbCritical, "Fehlende Rechte!"
                Exit Sub
            End If
            test = False
            If warten = False Then
                subdir = "/"
                markedfiles = ""
                For X = 0 To remotelist.ListCount - 1
                    If remotelist.Selected(X) = True Then markedfiles = markedfiles & Mid$(remotelist.List(X), 1, 4) & Mid$(remotelist.List(X), 62) & Chr(7) & Mid$(remotelist.List(X), 48, 12) & Chr(6)
                Next X
                If Mid$(Dir1.Path, Len(Dir1.Path) - 1, 2) = ":\" Then
                    ZielPfad = Dir1.Path
                Else
                    ZielPfad = Dir1.Path & "\"
                End If
                If MsgBox("Wollen Sie den Download in das Verzeichnis '" & ZielPfad & "' wirklich starten?", vbQuestion + vbYesNo, "Start des Downloadvorganges bestätigen") = vbNo Then Exit Sub
                ttl = ttl & "DOWNLOAD----"
            Else
                MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
            End If
        Else
            MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
            
        End If

    Case "Up"
        If comsock.State = 7 Then
            If warten = False Then
                If MsgBox("Wollen Sie den Uploadvorgang  in das Verzeichnis '" & ParentServerPath & "' wirklich starten?", vbQuestion + vbYesNo, "Start des Uploadvorganges bestätigen!") = vbNo Then Exit Sub
                UploadFilesZusammenstellen
            Else
                MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
            End If
        Else
            MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
        End If
    
    Case "MkdDir"
        If comsock.State = 7 Then
            If warten = False Then
                verzeichnis_erstellen.Show
                verzeichnis_erstellen.Text1.Text = ""
            Else
                MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
            End If
        Else
            MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
        End If
    
    Case "Dele"
        If MsgBox("Wollen Sie die ausgewählte(n) Datei(en) wirklich löschen?", vbQuestion + vbYesNo, "Bestätigung erforderlich") = vbYes Then
            If comsock.State = 7 Then
                If warten = False Then
                    deletedaten = ""
                    subdir = "/"
                    For X = 0 To remotelist.ListCount - 1
                        If remotelist.Selected(X) = True Then
                            If Mid$(remotelist.List(X), 62) <> "~" Then
                                deletedaten = deletedaten & Mid$(remotelist.List(X), 1, 4) & Mid$(remotelist.List(X), 62) & Chr(6)
                            Else
                                MsgBox "In das Verzeichnis '~' kann nicht gewechselt werden. Dieser Bug ist serverseitig zu begründen!", vbCritical, "Serverseitiger Fehler"
                            End If
                        End If
                    Next X
                    If deletedaten <> "" Then main.ttl = main.ttl & "DELETE------"
                Else
                    MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
                End If
            Else
                MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
            End If
        End If
        
    Case "Rename"
        If comsock.State = 7 Then
            If warten = False Then
                RenameObject = Mid$(remotelist.List(remotelist.ListIndex), 62)
                ttl = ttl & "RENAME_FROM-"
            Else
                MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
            End If
        Else
            MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
        End If
        
    Case "Analyse"
        If MsgBox("Durch das Öffnen des Analysefensters kann das Programm speziell bei schnellen Verbindungen Leistung einbüßen!" & vbCrLf & "Wollen Sie wirklich fortfahren?", vbYesNo + vbQuestion, "Leistungabfall zu erwarten!") = vbYes Then analyse_fenster.Show
    
    Case "author"
        main.Hide
        autor.Show
End Select
End Sub



Private Sub text_serverpath_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If comsock.State = 7 Then
        If warten = False Then
            warten = True
            ParentServerPath = text_serverpath.Text
            comsock.SendData "CWD " & ParentServerPath & vbCrLf
            WriteToConsole ("O U T G O I N G" & vbCrLf & "CWD " & ParentServerPath)
            ttl = ttl & "REFRESH-----"
        Else
            MsgBox "Es ist eine Operation im Gange. Warten Sie ab, bis diese erledigt ist!", vbInformation, "Aktion abgebrochen"
        End If
    Else
        MsgBox "Es besteht keine aktive Verbindung zu einem FTP Server - aus diesem Grunde kann diese Aktion nicht ausgeführt werden!", vbCritical, "Keine Verbindung"
    End If
End If

End Sub

Private Sub waiter_Timer()
waitervar = waitervar + 1
If waitervar = 2 Then
    ttl = "WRITELIST---" & ttl 'ENDE DES DATEILISTENTRANSFERS
    whatdata = ""
    datagot = ""
    sockdata.Disconnect
    waiter.Enabled = False
End If
End Sub

Private Sub workofftimer_Timer()
'Zur Analyse =========================================================================
If warten = True Then
    analyse_fenster.Label11.BackColor = &HFF00&
    remotelist.MousePointer = 11
Else
    analyse_fenster.Label11.BackColor = 255
    remotelist.MousePointer = 1
End If
analyse_fenster.timertodotext.Text = ttl

VD_bytesReceived = "\/ Bytes: " & BytesAngekommenTotal_c & " / " & BytesAngekommenTotal
VD_bytesSent = "/\ Bytes: " & BytesGesendetTotal_c & " / " & BytesGesendetTotal
VD_ip.Caption = "RemoteIP: " & comsock.RemoteHostIP

If comsock.State = 0 Then VD_comstatus.Caption = "Comsock Status: The Port is closed"
If comsock.State = 1 Then VD_comstatus.Caption = "Comsock Status: Connection in use"
If comsock.State = 2 Then VD_comstatus.Caption = "Comsock Status: Listening at Port " & comsock.LocalPort
If comsock.State = 3 Then VD_comstatus.Caption = "Comsock Status: Connection Pending"
If comsock.State = 4 Then VD_comstatus.Caption = "Comsock Status: Resolving Host"
If comsock.State = 5 Then VD_comstatus.Caption = "Comsock Status: Host Resolved"
If comsock.State = 6 Then VD_comstatus.Caption = "Comsock Status: Connecting"
If comsock.State = 7 Then VD_comstatus.Caption = "Comsock Status: Connected"
If comsock.State = 8 Then
    VD_comstatus.Caption = "Comsock Status: Peer is closing the connection"
    login_fenster.LoginGo
End If
If comsock.State = 9 Then VD_comstatus.Caption = "Comsock Status: Error"

If datasock.State = 0 Then VD_datastatus.Caption = "Datasock Status: The Port is closed"
If datasock.State = 1 Then VD_datastatus.Caption = "Datasock Status: Connection in use"
If datasock.State = 2 Then VD_datastatus.Caption = "Datasock Status: Listening at Port " & datasock.LocalPort
If datasock.State = 3 Then VD_datastatus.Caption = "Datasock Status: Connection Pending"
If datasock.State = 4 Then VD_datastatus.Caption = "Datasock Status: Resolving Host"
If datasock.State = 5 Then VD_datastatus.Caption = "Datasock Status: Host Resolved"
If datasock.State = 6 Then VD_datastatus.Caption = "Datasock Status: Connecting"
If datasock.State = 7 Then VD_datastatus.Caption = "Datasock Status: Connected"
If datasock.State = 8 Then VD_datastatus.Caption = "Datasock Status: Peer is closing the connection"
If datasock.State = 9 Then VD_datastatus.Caption = "Datasock Status: Error"

If sockdata.State = 0 Then VD_sockstatus.Caption = "Sockdata Status: The Port is closed"
If sockdata.State = 1 Then VD_sockstatus.Caption = "Sockdata Status: Socket ready"
If sockdata.State = 2 Then VD_sockstatus.Caption = "Sockdata Status: Socket listening at Port " & sockdata.LocalPort
If sockdata.State = 3 Then VD_sockstatus.Caption = "Sockdata Status: Socket connecting"
If sockdata.State = 4 Then VD_sockstatus.Caption = "Sockdata Status: Socket accepting"
If sockdata.State = 5 Then VD_sockstatus.Caption = "Sockdata Status: Socket receiving"
If sockdata.State = 6 Then VD_sockstatus.Caption = "Sockdata Status: Socket sending"
If sockdata.State = 7 Then VD_sockstatus.Caption = "Sockdata Status: Socket closing"
'=====================================================================================

timertodo = Mid$(ttl, 1, 12)
ttl = Mid$(ttl, 13)
' ====================== LOGINAKTIVITÄT ==========================
'#Benutzername senden
If timertodo = "LOGIN-------" Then
    If warten = False Then
        warten = True
        comsock.SendData "USER " & login_fenster.UserName.Text & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "USER " & login_fenster.UserName.Text)
        ttl = ttl & "PASSWORD----"
    Else: ttl = "LOGIN-------" & ttl
    End If
End If

'#Passwort senden
If timertodo = "PASSWORD----" Then
    If warten = False Then
        warten = True
        comsock.SendData "PASS " & login_fenster.Password.Text & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "PASS " & login_fenster.Password.Text)
        ttl = ttl & "PWD-BEFEHL--"
    Else
        ttl = "PASSWORD----" & ttl
    End If
End If

If timertodo = "PWD-BEFEHL--" Then
    If warten = False Then
        warten = True
        PWD_Befehl = True
        comsock.SendData "PWD" & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "PWD")
        ttl = ttl & "REFRESH-----"
    Else
        ttl = "PWD-BEFEHL--" & ttl
    End If
End If

If timertodo = "LIST--------" Then
    If warten = False And sockdata.State = 2 Then
        warten = True
        
        comsock.SendData "TYPE A" & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "TYPE A")
        Do Until warten = False And sockdata.State = 2
            DoEvents
        Loop
        
        whatdata = "listneu"
         wholedata = ""
        warten = True
        comsock.SendData "LIST" & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "LIST")
    Else: ttl = "LIST--------" & ttl
    End If
End If
' ===============================================================


If timertodo = "WRITELIST---" Then 'Liste in Listbox ausgeben
    remotelist.Clear
    warten = True
    Do Until sockdata.State = 0
        If wholedata <> "" Then
            If Mid$(wholedata, Len(wholedata) - 1, 2) = vbCrLf Then Exit Do
        End If
        If wholedata = "" And Disconnected Then Exit Do
        DoEvents
    Loop
    Disconnected = False
    sockdata.Disconnect

    comsock.SendData "TYPE I" & vbCrLf
    WriteToConsole ("O U T G O I N G" & vbCrLf & "TYPE I")
    Do Until wholedata = ""
        Y = 0
        X = 0
        Do Until X = 13
            Y = Y + 1
            X = Asc(Mid$(wholedata, Y, 1))
            If X = 13 Then remotelist.AddItem Mid$(wholedata, 1, Y - 1)
            If Y = Len(wholedata) Then X = 13
        Loop
        If Y = Len(wholedata) Then
            If LCase(Mid$(remotelist, 1, 5)) <> "total" Then remotelist.AddItem wholedata
            wholedata = ""
        Else
            wholedata = Mid$(wholedata, Y + 2)
        End If
    Loop
    'Dateiliste wird nun strukturiert
    If analyse_fenster.struktur.Value = 1 Then
        If Mid$(remotelist.List(0), 1, 5) = "total" Then remotelist.RemoveItem 0
        For C = 0 To remotelist.ListCount - 1
            zwischen_List = remotelist.List(C)
            e = 0
            For d = 1 To Len(zwischen_List) 'Zuerst Abspaltung des Dateinamens, denn wenn im Dateinamen selbst mehrere Space vorhanden sind werden die zu einem einzelnen space konvertiert
                If Mid$(zwischen_List, d, 1) = " " And Mid$(zwischen_List, d + 1, 1) <> " " Then e = e + 1
                If Mid$(zwischen_List, 1, 1) <> "d" And Mid$(zwischen_List, 1, 1) <> "l" And Mid$(zwischen_List, 1, 1) <> "-" And e = 3 Then Exit For
                If e > 7 Then Exit For
            Next d
            zwischen4 = Mid$(zwischen_List, d + 1)
            zwischen_List = Mid$(zwischen_List, 1, d)
            
            For Y = 25 To 1 Step -1 'Alle Spaces im String, egal wie lange durch eines ersetzen
                zwischen_List = Replace(zwischen_List, Space(Y), "/\")
            Next Y
            
            zwischen_List = zwischen_List & zwischen4
            
            If UBound(Split(zwischen_List, "/\")) > 7 Then '>7 => UNIX DATEISYSTEM
                
                ' String wird in Array gesplittet nach " " Kriterium, jedoch Dateiname nicht,
                ' weil er spaces enthalten kann=> nur 8 arrays
                zwischen_WriteList = Split(zwischen_List, "/\", 9)
            
            'DATENEIGENSCHAFT (von 1 bis 4)
                If Mid$(zwischen_WriteList(0), 1, 1) = "d" Then remotelist.List(C) = " DIR"
                If Mid$(zwischen_WriteList(0), 1, 1) = "l" Then remotelist.List(C) = "LINK"
                If Mid$(zwischen_WriteList(0), 1, 1) = "-" Then remotelist.List(C) = "FILE"
            
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
                
            'DATEIATTRIBUTE (von 7 bis 15)
                remotelist.List(C) = remotelist.List(C) & Mid$(zwischen_WriteList(0), 2)
                
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
                
            'EIGNER DER DATEN ANZEIGEN (von 18 bis 31)
                Y = 14 - Len(zwischen_WriteList(2))
                remotelist.List(C) = remotelist.List(C) & Space(Y) & zwischen_WriteList(2)
                
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
                
            'DATUM - TAG ANZEIGEN (von 34 bis 35)
                Y = 2 - Len(zwischen_WriteList(6))
                remotelist.List(C) = remotelist.List(C) & Space(Y) & zwischen_WriteList(6)
                
                remotelist.List(C) = remotelist.List(C) & " " '1 Platzhalter
                
            'DATUM - MONAT ANZEIGEN (von 37 bis 39)
                remotelist.List(C) = remotelist.List(C) & zwischen_WriteList(5)
                
                remotelist.List(C) = remotelist.List(C) & " " '1 Platzhalter
            
            'DATUM - UHRZEIT/JAHR ANZEIGEN (von 41 bis 45)
                Y = 5 - Len(zwischen_WriteList(7))
                remotelist.List(C) = remotelist.List(C) & Space(Y) & zwischen_WriteList(7)
                
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
            'DATEIGRÖßE ANZEIGEN (von 48 bis 59)
                Y = 12 - Len(zwischen_WriteList(4))
                remotelist.List(C) = remotelist.List(C) & Space(Y) & zwischen_WriteList(4)
                
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
            
            'DATEINAMEN ANZEIGEN (von 62 bis .. )
                remotelist.List(C) = remotelist.List(C) & zwischen_WriteList(8)
            
            
            Else 'MS DOS DATEISYSTEM
                
                zwischen_WriteList = Split(zwischen_List, "/\", 4)
                
                
            'DATENEIGENSCHAFT (von 1 bis 4)
                If zwischen_WriteList(2) = "<DIR>" Then
                    remotelist.List(C) = " DIR"
                Else
                    remotelist.List(C) = "FILE"
                End If
                
                remotelist.List(C) = remotelist.List(C) & Space(24) 'weil andere Information fehlen
                
            'DATUM (von 29 bis 36)
                remotelist.List(C) = remotelist.List(C) & zwischen_WriteList(0)
                
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
                
            'UHRZEIT (von 39 bis 45)
                remotelist.List(C) = remotelist.List(C) & zwischen_WriteList(1)
                
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
                
            'DATEIGRÖßE ANZEIGEN (48 bis 59)
                If zwischen_WriteList(2) = "<DIR>" Then
                    remotelist.List(C) = remotelist.List(C) & Space(12)
                Else
                    Y = 12 - Len(zwischen_WriteList(2))
                    remotelist.List(C) = remotelist.List(C) & Space(Y) & zwischen_WriteList(2)
                End If
                
                remotelist.List(C) = remotelist.List(C) & "  " '2 Platzhalter
                    
            'DATEINAMEN ANZEIGEN (von 62 bis .. )
                remotelist.List(C) = remotelist.List(C) & zwischen_WriteList(3)

            End If
        Next C
    End If
    
    X = 0
    Y = 0
    'Einzelne Sachen herausfiltern (z.B.: .., .)
    For C = 0 To remotelist.ListCount - 1
        If Mid$(remotelist.List(C), 62, 1) = "." And Len(remotelist.List(C)) = 62 Then remotelist.RemoveItem (C)
        If Mid$(remotelist.List(C), 62, 2) = ".." And Len(remotelist.List(C)) = 63 Then remotelist.RemoveItem (C)
        If Mid$(remotelist.List(C), 1, 4) = " DIR" Then X = X + 1
        If Mid$(remotelist.List(C), 1, 4) = "FILE" Then Y = Y + 1
    Next C
    ListCounter.Caption = X & " Ordner und " & Y & " Dateien"
    remotelist.Enabled = True
End If


If timertodo = "DOWNLOAD----" Then
    If warten = False And sockdata.State = 0 Then
        If markedfiles = "" Then
            remotelist.Selected(remotelist.ListIndex) = False
            whatdata = "file"
            filegot = ""
            generateport
            BytesAngekommenDatei = 0
            ttl = "DOWNLOAD_1--" & ttl
        End If
        
        If markedfiles <> "" Then
            If subdir = "/" Then
                subdircount = 0
                zwischen = "0"
                X = 0
                Do Until Asc(zwischen) = 6 'Nichterlaubtes Zeichen als Trennpunkt
                    X = X + 1
                    zwischen = Mid$(markedfiles, X, 1)
                Loop
                zwischen2 = Mid$(markedfiles, 1, X - 1)
                markedfiles = Mid$(markedfiles, X + 1)
                
                If zwischen2 = "\/-ENDE-\/" Then
                    ttl = ""
                    MsgBox "Der Downloadvorgang von einer bzw. mehreren Datei(en) wurde erfolgreich abgeschlossen", vbInformation, "Download beendet"
                    Dir1.Refresh
                    File1.Refresh
                    Exit Sub
                End If
                If markedfiles = "" Then markedfiles = "\/-ENDE-\/" & Chr(6)
                
                If Mid$(zwischen2, 1, 4) = "FILE" Then
                    zwischen2 = Mid$(zwischen2, 5)
                    X = 0
                    Do Until Asc(zwischen) = 7 'Nichterlaubtes Zeichen als Trennpunkt
                        X = X + 1
                        zwischen = Mid$(zwischen2, X, 1)
                    Loop
                    DownloadFilename = Mid$(zwischen2, 1, X - 1)
                    Dateigroesse = Mid$(zwischen2, X + 1)
                    whatdata = "file"
                    filegot = ""
                    generateport

                    save_as = DownloadFilename
                    BytesAngekommenDatei = 0
                    
                    ttl = "DOWNLOAD_1--" & ttl
                    
                    Exit Sub
                End If
                If Mid$(zwischen2, 1, 4) = " DIR" Then
                    zwischen2 = Mid$(zwischen2, 5)
                    zwischen = "0"
                    X = 0
                    Do Until Asc(zwischen) = 7 'Nichterlaubtes Zeichen als Trennpunkt
                        X = X + 1
                        zwischen = Mid$(zwischen2, X, 1)
                    Loop
                    
                    'Filter für bad DIR - Namen
                    CheckNameIfBad Mid$(zwischen2, 1, X - 1), "Ordner"
                    
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    If fs.folderexists(ZielPfad & CheckedString) = -1 Then 'Wenn der Ordner bereits lokal existiert
                        m = MsgBox("Der Ordner '" & ZielPfad & CheckedString & "' existiert bereits im Zielordner und wird gelöscht!" & vbCrLf & "Wollen Sie, dass diese Warnung auch weiterhin angezeigt wird?", vbYesNoCancel + vbQuestion, "Löschen des Ordners bestätigen")
                        If X = 2 Then 'cancel
                            ttl = "DOWNLOAD----"
                            Exit Sub
                        End If
                        If X = 7 Then HideWarningsForMultipleLocalFolders = True 'no
                        fs.deletefolder (ZielPfad & CheckedString)
                    End If
                    
                    subdir = subdir & CheckedString & "/"
                    ZielPfad = ZielPfad & CheckedString & "\"
                    
                    subdircount = "0000"
                    
                    zwischen_download = Mid$(zwischen2, 1, X - 1)
                                        
                    ttl = "DOWNLOAD_2--" & "REFRESH-----" & "DOWNLOAD----" & ttl
                    
                    MkDir Mid$(ZielPfad, 1, Len(ZielPfad) - 1)
                    Exit Sub
                End If
            End If
                
            If subdir <> "/" Then
                X = Mid$(subdircount, Len(subdircount) - 3) 'Bei welcher Datei / welchem Ordner weitergemacht werden soll
                If X = remotelist.ListCount Then
                    
                    subdir = Mid$(subdir, 1, Len(subdir) - 1)
                    For Y = 1 To Len(subdir)
                        If Mid$(subdir, Y, 1) = "/" Then z = Y
                    Next Y
                    subdir = Mid$(subdir, 1, z)
                    
                    ZielPfad = Mid$(ZielPfad, 1, Len(ZielPfad) - 1)
                    For Y = 1 To Len(ZielPfad)
                        If Mid$(ZielPfad, Y, 1) = "\" Then z = Y
                    Next Y
                    ZielPfad = Mid$(ZielPfad, 1, z)
                    
                    If Len(subdircount) = 4 Then
                        subdircount = ""
                    Else
                        subdircount = Mid$(subdircount, 1, Len(subdircount) - 5)
                    End If
                    ttl = "DOWNLOAD_3--" & "REFRESH-----" & "DOWNLOAD----" & ttl
                    Exit Sub
                Else
                    X = X + 1
                    subdircount = Mid$(subdircount, 1, Len(subdircount) - 4) & Space(4 - Len(X)) & X
                    
                    If Mid$(remotelist.List(X - 1), 1, 4) = "FILE" Then
                        DownloadFilename = Mid$(remotelist.List(X - 1), 62)
                        Dateigroesse = Mid$(remotelist.List(X - 1), 48, 12)
                        whatdata = "file"
                        filegot = ""
                        
                        generateport
                        
                        save_as = DownloadFilename
                        BytesAngekommenDatei = 0
                        ttl = "DOWNLOAD_1--" & ttl
                        
                        Exit Sub
                    End If
                    If Mid$(remotelist.List(X - 1), 1, 4) = " DIR" Then
                        CheckNameIfBad Mid$(remotelist.List(X - 1), 62), "Ordner"
                        subdir = subdir & CheckedString & "/"
                        ZielPfad = ZielPfad & CheckedString & "\"
                        
                        subdircount = subdircount & ",0000"
                        
                        zwischen_download = Mid$(remotelist.List(X - 1), 62)
                                              
                        MkDir Mid$(ZielPfad, 1, Len(ZielPfad) - 1)
                        ttl = "DOWNLOAD_2--" & "REFRESH-----" & "DOWNLOAD----" & ttl
                        
                        Exit Sub
                    End If
                End If
            End If
        End If
    Else
        ttl = "DOWNLOAD----" & ttl
    End If
End If

If timertodo = "DOWNLOAD_1--" Then
    CheckNameIfBad save_as, "Datei"
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.fileexists(ZielPfad & CheckedString) = -1 Then 'Datei existiert schon
        If HideWarningsForMultipleLocalFiles = False Then
            X = MsgBox("Der Dateiname '" & CheckedString & "' existiert bereits im Zielordner '" & ZielPfad & "' und wird überschrieben!" & vbCrLf & "Wollen Sie diese Meldung weiterhin erhalten?", vbYesNoCancel + vbQuestion, "Überschreiben der Datei bestätigen")
            If X = 2 Then 'cancel
                If markedfiles <> "" Then ttl = "DOWNLOAD----"
                Exit Sub
            End If
            If X = 7 Then HideWarningsForMultipleLocalFiles = True 'No
        End If
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.DeleteFile (ZielPfad & CheckedString)
    End If
    

    If Dateigroesse > 0 Then
        warten = True
        comsock.SendData "RETR " & DownloadFilename & vbCrLf
        If ShowDLProgress = False Then StartProgressVisualisation "down"
        WriteToConsole ("O U T G O I N G" & vbCrLf & "RETR " & DownloadFilename)
    Else
        CheckNameIfBad save_as, "Datei"
        file = FreeFile
        Open ZielPfad & CheckedString For Binary As #file
        Close #file
    End If
End If

If timertodo = "DOWNLOAD_2--" Then
    warten = True
    ServerpfadNeuDefinieren zwischen_download
    comsock.SendData "CWD " & zwischen_download & vbCrLf
    WriteToConsole ("O U T G O I N G" & vbCrLf & "CWD " & zwischen_download)
End If

If timertodo = "DOWNLOAD_3--" Then
    warten = True
    ServerpfadNeuDefinieren "//DirUp"
    comsock.SendData "CWD .." & vbCrLf
    WriteToConsole ("O U T G O I N G" & vbCrLf & "CWD ..")
End If


If timertodo = "REFRESH-----" Then
    remotelist.Enabled = False
    remotelist.Clear
    If warten = False Then
        File1.Refresh
        Dir1.Refresh
        Drive1.Refresh
        text_serverpath.Text = ParentServerPath
        ttl = "LIST--------" & ttl
        generateport
    Else
        ttl = "REFRESH-----" & ttl
    End If
End If

If timertodo = "DELETE------" Then
    If warten = False Then
        
        If subdir = "/" Then
            subdircount = ""
            zwischen = "0"
            X = 0
            Do Until Asc(zwischen) = 6
                X = X + 1
                zwischen = Mid$(deletedaten, X, 1)
            Loop
            zwischen2 = Mid$(deletedaten, 1, X - 1)
            deletedaten = Mid$(deletedaten, X + 1)
                
            If zwischen2 = "\/-ENDE-\/" Then
                ttl = "REFRESH-----"
                MsgBox "Der Löschvorgang wurde erfolgreich abgeschlossen", vbInformation, "Löschvorgang beendet"
                Exit Sub
            End If
            If deletedaten = "" Then deletedaten = "\/-ENDE-\/" & Chr(6)
                
            If Mid$(zwischen2, 1, 4) = "FILE" Then
                warten = True
                comsock.SendData "DELE " & Mid$(zwischen2, 5) & vbCrLf
                WriteToConsole ("O U T G O I N G" & vbCrLf & "DELE " & Mid$(zwischen2, 5))

                ttl = ttl & "DELETE------"
                    
                Exit Sub
            End If
            If Mid$(zwischen2, 1, 4) = " DIR" Then
                subdir = subdir & Mid$(zwischen2, 5) & "/"
                subdircount = "0000"
                
                warten = True
                ServerpfadNeuDefinieren Mid$(zwischen2, 5)
                comsock.SendData "CWD " & Mid$(zwischen2, 5) & vbCrLf
                WriteToConsole ("O U T G O I N G" & vbCrLf & "CWD " & Mid$(zwischen2, 5))

                ttl = ttl & "REFRESH-----" & "DELETE------"
                    
                Exit Sub
            End If
        End If
                
        If subdir <> "/" Then
            X = Mid$(subdircount, Len(subdircount) - 3) 'Bei welcher Datei / welchem Ordner weitergemacht werden soll
            If X = remotelist.ListCount Then
                    
                warten = True
                ServerpfadNeuDefinieren "//DirUp"
                comsock.SendData "CWD .." & vbCrLf & vbCrLf
                WriteToConsole ("O U T G O I N G" & vbCrLf & "CWD ..")
                Do Until warten = False
                    DoEvents
                Loop
                                
                subdir = Mid$(subdir, 1, Len(subdir) - 1)
                For Y = 1 To Len(subdir)
                    If Mid$(subdir, Y, 1) = "/" Then z = Y
                Next Y
                
                warten = True
                comsock.SendData "RMD " & Mid$(subdir, z + 1) & vbCrLf
                WriteToConsole ("O U T G O I N G" & vbCrLf & "RMD " & Mid$(subdir, z + 1))
                
                subdir = Mid$(subdir, 1, z)
                
                If subdir = "/" Then ttl = "DELETE------" & ttl
                If subdir <> "/" Then ttl = "REFRESH-----" & "DELETE------" & ttl
                
                If Len(subdircount) = 4 Then
                    subdircount = ""
                Else
                    subdircount = Mid$(subdircount, 1, Len(subdircount) - 9) & "0000"
                End If
                    
                Exit Sub
            Else
                X = X + 1
                If Len(X) = 1 Then subdircount = Mid$(subdircount, 1, Len(subdircount) - 4) & "000" & X
                If Len(X) = 2 Then subdircount = Mid$(subdircount, 1, Len(subdircount) - 4) & "00" & X
                If Len(X) = 3 Then subdircount = Mid$(subdircount, 1, Len(subdircount) - 4) & "0" & X
                If Len(X) = 4 Then subdircount = Mid$(subdircount, 1, Len(subdircount) - 4) & X
                    
                If Mid$(remotelist.List(X - 1), 1, 4) = "FILE" Then
                
                    warten = True
                    comsock.SendData "DELE " & Mid$(remotelist.List(X - 1), 62) & vbCrLf
                    WriteToConsole ("O U T G O I N G" & vbCrLf & "DELE " & Mid$(remotelist.List(X - 1), 62))
                                             
                    ttl = ttl & "DELETE------"
                        
                    Exit Sub
                End If
                If Mid$(remotelist.List(X - 1), 1, 4) = " DIR" Then
                    subdir = subdir & Mid$(remotelist.List(X - 1), 62) & "/"

                    subdircount = subdircount & ",0000"

                    warten = True
                    ServerpfadNeuDefinieren Mid$(remotelist.List(X - 1), 62)
                    comsock.SendData "CWD " & Mid$(remotelist.List(X - 1), 62) & vbCrLf
                    WriteToConsole ("O U T G O I N G" & vbCrLf & "DELE " & "CWD " & Mid$(remotelist.List(X - 1), 62))
                        
                    ttl = ttl & "REFRESH-----" & "DELETE------"
                        
                    Exit Sub
                End If
            End If
        End If
    Else
        ttl = "DELETE------" & ttl
    End If
End If

If timertodo = "MKD DIR-----" Then
    If warten = False Then
        warten = True
        comsock.SendData "MKD " & verzeichnis_erstellen.Text1.Text & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "MKD " & verzeichnis_erstellen.Text1.Text)
        ttl = ttl & "REFRESH-----"
    Else
        ttl = "MKD DIR-----" & ttl
    End If
End If

If timertodo = "RENAME_FROM-" Then
    If warten = False Then
        warten = True
        comsock.SendData "RNFR " & RenameObject & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "RNFR " & RenameObject)
        RenameForm.Show
        main.Enabled = False
        RenameForm.Text1.Text = RenameObject
    Else
        ttl = "RENAME_FROM" & ttl
    End If
End If

If timertodo = "RENAME_TO---" Then
    If warten = False Then
        warten = True
        comsock.SendData "RNTO " & RenameForm.Text1.Text & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "RNTO " & RenameForm.Text1.Text)
        ttl = ttl & "REFRESH-----"
    Else
        ttl = "RENAME_TO---" & ttl
    End If
End If


If timertodo = "ABORT_DELETE" Then
    If warten = False Then
        warten = True
        comsock.SendData "DELE " & UploadFilename & vbCrLf
        WriteToConsole ("O U T G O I N G" & vbCrLf & "DELE " & UploadFilename)
    Else
        ttl = "ABORT_DELETE" & ttl
    End If
End If

If timertodo = "SAVEDOWNLOAD" Then
    If whatdata = "file" Then
        If Dateigroesse = BytesAngekommenDatei Then
            If ShowDLProgress Then StartProgressVisualisation ("")
            
            If DateiSchonOffen Then
                Put file, zwischenDg_Down, filegot
                DateiSchonOffen = False
                Close #file
                Name ZielPfad & CheckedString As ZielPfad & RichtigerDateiname
            Else
                file = FreeFile
                Open ZielPfad & CheckedString For Binary As #file
                Put file, 1, filegot
                Close #file
            End If
            
            sockdata.Disconnect
            warten = False
            If markedfiles <> "" Then ttl = ttl & "DOWNLOAD----"
        Else
            ttl = "SAVEDOWNLOAD" & ttl
        End If
    End If
End If
End Sub

Public Sub generateport()
datasock.Close
sockdata.Disconnect
datasock.Listen
X = datasock.LocalPort
If Use_Datasock = True Then
    Use_Datasock = False
Else
    sockdata.LocalPort = datasock.LocalPort
    datasock.Close
    sockdata.Listen
End If
    
'Lokale IP in Beistrich Schreibweise konvertieren

Zwischenport = Replace(comsock.LocalIP, ".", ",")
Zwischenport = Zwischenport & "," & ((X - (X Mod 256)) / 256) & "," & (X Mod 256)
    
Do Until warten = False And (sockdata.State = 2 Or datasock.State = 2)
    DoEvents
Loop
    
warten = True
comsock.SendData "PORT " & Zwischenport & vbCrLf
WriteToConsole ("O U T G O I N G" & vbCrLf & "PORT " & Zwischenport)
Do Until warten = False
    DoEvents
Loop
End Sub


Public Sub logoff()
If comsock.State = 7 Then
    comsock.SendData "QUIT" & vbCrLf
    WriteToConsole ("O U T G O I N G" & vbCrLf & "QUIT")
End If
remotelist.Clear
ListCounter.Visible = False
analyse_fenster.console.Text = ""
workofftimer.Enabled = False
sockdata.Disconnect
comsock.Close
datagot = ""
markedfiles = ""
filegot = ""
ttl = ""

BytesGesendetTotal = 0
BytesAngekommenTotal = 0
BytesGesendetTotal_c = 0
BytesAngekommenTotal_c = 0
End Sub

Private Sub UploadFilesZusammenstellen()
FileNamesLocal = ""
DirsToCreate = ""


'Falls eine Datei(en) in der Dateiliste markiert wurde(n)
If File1.FileName <> "" Then
    For Y = 0 To File1.ListCount - 1
        If File1.Selected(Y) = True Then FileNamesLocal = FileNamesLocal & Dir1.Path & "\" & File1.List(Y) & vbCrLf
    Next Y
    If FileNamesLocal <> "" Then
        'Nach Doppel Backslash Bug scannen
        For X = 1 To Len(FileNamesLocal) - 1
            If Mid$(FileNamesLocal, 1, 2) = "\\" Then FileNamesLocal = Mid$(FileNamesLocal, 1, X) & Mid$(FileNamesLocal, X + 2)
        Next X
        OrdnerAufServerErstellen
        Exit Sub
    End If
End If

'Falls keine Datei(en) in der Dateiliste markiert wurde(n) - Ordner wird upgeloaded
DirsToCreate = Dir1.List(-1) & vbCrLf
For Y = 0 To File1.ListCount - 1
    FileNamesLocal = FileNamesLocal & Dir1.Path & "\" & File1.List(Y) & vbCrLf
Next Y
X = 0
'Array Leeren
z = 0
For z = 0 To 256
    dirpos(z) = "0"
Next z
Do Until X = -1
    waitvar = False
    
    If Dir1.List(dirpos(X)) <> "" And waitvar = False Then
        'Wechsel in Unterverzeichnis
        Dir1.Path = Dir1.List(dirpos(X))
        DirsToCreate = DirsToCreate & Dir1.List(-1) & vbCrLf
        X = X + 1
        dirpos(X) = 0
        
        For Y = 0 To File1.ListCount - 1
            FileNamesLocal = FileNamesLocal & Dir1.Path & "\" & File1.List(Y) & vbCrLf
        Next Y
        waitvar = True
    End If
    
    'Ende
    If X = 0 Then
        OrdnerAufServerErstellen
        Exit Sub
    End If
    
    If Dir1.List(dirpos(X)) = "" And waitvar = False Then
        'Wechsel in Oberverzeichnis
        X = X - 1
        dirpos(X) = dirpos(X) + 1
        Dir1.Path = Dir1.List(-2)
        
        waitvar = True
    End If
Loop

End Sub

Private Sub OrdnerAufServerErstellen()
If DirsToCreate <> "" Then
    'Falls ein Ordner Upgeloaded wird
    DirsToCreate = Replace(DirsToCreate, "\", "/") 'Alle Backslashes in Slashes konvertieren
    
    zwischen = ""
    X = 0
    Do Until zwischen <> "" 'Erster Pfad wird herausgesucht
        X = X + 1
        If Mid$(DirsToCreate, X, 2) = vbCrLf Then
            zwischen = Mid$(DirsToCreate, 1, X - 1)
        End If
    Loop
            
    For X = 1 To Len(zwischen) 'Bei diesem wird der Pfad geschnitten
        If Mid(zwischen, X, 1) = "/" Then Y = X
    Next X
    
    zwischen3 = Mid$(zwischen, 1, Y - 1)
    
    X = 0
    Do Until DirsToCreate = ""
        X = X + 1
        If Mid$(DirsToCreate, X, 2) = vbCrLf Then
            Do Until warten = False
                DoEvents
            Loop
            'Checken, ob der upzuloadende Ordner bereits existiert ("/" zählen)
            w = 0
            For z = 0 To Len(Mid$(DirsToCreate, Y, X - Y)) - 1
                If Mid$(DirsToCreate, Y + z, 1) = "/" Then w = w + 1
            Next z
            If w <= 1 Then
                For z = 0 To remotelist.ListCount - 1
                    If Mid$(remotelist.List(z), 62) = Mid$(DirsToCreate, Y + 1, X - Y - 1) Then
                        'Upload abbrechen
                        DirsToCreate = ""
                        FileNamesServer = ""
                        FileNamesLocal = ""
                        MsgBox "Upload dieses Objektes ist nicht möglich, da ein Ordner desselben Namens bereits besteht!", vbCritical, "Start des Uploadvorganges verweigert"
                        Exit Sub
                    End If
                Next z
                warten = True
                comsock.SendData "MKD " & Mid$(ParentServerPath, 1, Len(ParentServerPath) - 1) & Mid$(DirsToCreate, Y, X - Y) & vbCrLf
                WriteToConsole ("O U T G O I N G" & vbCrLf & "MKD " & Mid$(ParentServerPath, 1, Len(ParentServerPath) - 1) & Mid$(DirsToCreate, Y, X - Y))
                DirsToCreate = Mid$(DirsToCreate, X + 2)
                X = 0
            Else
                warten = True
                comsock.SendData "MKD " & Mid$(ParentServerPath, 1, Len(ParentServerPath) - 1) & Mid$(DirsToCreate, Y, X - Y) & vbCrLf
                WriteToConsole ("O U T G O I N G" & vbCrLf & "MKD " & Mid$(ParentServerPath, 1, Len(ParentServerPath) - 1) & Mid$(DirsToCreate, Y, X - Y))
                DirsToCreate = Mid$(DirsToCreate, X + 2)
                X = 0
            End If
        End If
    Loop
Else
    'Falls Datei(en) upgeloaded werden
    zwischen = Replace(FileNamesLocal, "\", "/") 'Alle Backslashes in Slashes konvertieren
    
    zwischen2 = ""
    X = 0
    Do Until zwischen2 <> "" 'Erster Pfad wird herausgesucht
        X = X + 1
        If Mid$(zwischen, X, 2) = vbCrLf Then
            zwischen2 = Mid$(zwischen, 1, X - 1)
        End If
    Loop
            
    For X = 1 To Len(zwischen2) 'Bei diesem wird der Pfad geschnitten
        If Mid(zwischen, X, 1) = "/" Then Y = X
    Next X
    
    zwischen3 = Mid$(zwischen, 1, Y - 1)
    
    '=== um Überschreiben der Dateien am Server beim Upload zu verhindern
    zwischen4 = Replace(zwischen, zwischen3, "")
    zwischen4 = Replace(zwischen4, "/", "")
    Do Until zwischen4 = ""
        For Y = 1 To Len(zwischen4)
            If Mid$(zwischen4, Y, 2) = vbCrLf Then
                zwischen5 = Mid$(zwischen4, 1, Y - 1)
                zwischen4 = Mid$(zwischen4, Y + 2)
                Exit For
            End If
        Next Y
        For X = 0 To remotelist.ListCount - 1
            If Mid$(remotelist.List(X), 62) = zwischen5 Then
                If MsgBox("Die Datei '" & zwischen5 & "' existiert bereits auf dem Server im gegenwärtigen Verzeichnis." & vbCrLf & "Soll diese Datei nun überschrieben werden?", vbYesNo + vbQuestion, "Überschreiben der Datei bestätigen") = vbNo Then
                    'Datei wird aus dem Filenamesserver string entfernt
                    For z = 1 To Len(FileNamesLocal) 'Datei definiert durch \ + Dateiname + vbcrlf
                        If Mid$(FileNamesLocal, z, Len(zwischen5) + 3) = "\" & zwischen5 & vbCrLf Then
                            w = z
                            Do Until zwischen6 = vbCrLf Or w = 0
                                zwischen6 = Mid$(FileNamesLocal, w, 2)
                                w = w - 1
                            Loop
                            If w = 0 Then
                                FileNamesLocal = Mid$(FileNamesLocal, z + Len(zwischen5) + 3)
                                If FileNamesLocal = "" Then Exit Sub
                            Else
                                FileNamesLocal = Mid$(FileNamesLocal, 1, w + 2) & Mid$(FileNamesLocal, z + Len(zwischen5) + 3)
                            End If
                            Exit For
                        End If
                    Next z
                End If
            End If
        Next X
    Loop

    '===
End If

'Jetzt werden die FileNames auf ServerNorm gebracht
FileNamesServer = Replace(FileNamesLocal, "\", "/")
FileNamesServer = Replace(FileNamesServer, zwischen3, Mid$(ParentServerPath, 1, Len(ParentServerPath) - 1))


'Jetzt werden die Dateien Upgeloaded
FilesAufServerUploaden
End Sub

Private Sub FilesAufServerUploaden()
Do Until warten = False
    DoEvents
Loop
zwischen2 = ""
X = 0
Do Until zwischen2 = vbCrLf
    X = X + 1
    zwischen2 = Mid$(FileNamesServer, X, 2)
Loop
UploadFilename = Mid$(FileNamesServer, 1, X - 1)
FileNamesServer = Mid$(FileNamesServer, X + 2)

zwischen2 = ""
X = 0
Do Until zwischen2 = vbCrLf
    X = X + 1
    zwischen2 = Mid$(FileNamesLocal, X, 2)
Loop
zwischen = Mid$(FileNamesLocal, 1, X - 1)
FileNamesLocal = Mid$(FileNamesLocal, X + 2)

Use_Datasock = True
generateport

warten = True
comsock.SendData "STOR " & UploadFilename & vbCrLf
WriteToConsole ("O U T G O I N G" & vbCrLf & "STOR " & File1.FileName)

Do Until datasock.State = 7
    DoEvents
Loop

warten = True

Dateigroesse = FileLen(zwischen)

If Dateigroesse > 0 Then
    zwischenDg_Up = 1
    If ShowDLProgress = False Then StartProgressVisualisation "up"
    whatdata = ""
    
    file = FreeFile
    Open zwischen For Binary As #file
    
    Currently_Uploading = True
    
    FileIsBeingSent
Else
    whatdata = ""
    datasock.Close
    ttl = ttl & "REFRESH-----"
End If
End Sub

Private Sub FileIsBeingSent()

If Dateigroesse - zwischenDg_Up < 1048576 Then 'letzes Dateipacket
    DataToUpload = Space(Dateigroesse - zwischenDg_Up + 1)
    Get file, zwischenDg_Up, DataToUpload
    datasock.SendData DataToUpload
    Zwischen_Uppen_Bytes = 0
    zwischenDg_Up = Dateigroesse
    Currently_Uploading = False
    Close #file
Else '1 MB grosse Dateipackete - falls datei >=1MB
    DataToUpload = Space(1048576)
    Get file, zwischenDg_Up, DataToUpload
    datasock.SendData DataToUpload
    Zwischen_Uppen_Bytes = 0
    zwischenDg_Up = zwischenDg_Up + 1048576
End If
    
End Sub



Public Sub WriteToConsole(consoleText As String)
If analyse_fenster.Visible = True Then
    If Mid$(consoleText, 1, 4) = "LDR#" Then
        If Mid$(consoleText, 5, 3) = "200" Then
            analyse_fenster.Label4.BackColor = &HFF00&
        Else
            analyse_fenster.Label4.BackColor = 255
        End If
        If Mid$(consoleText, 5, 3) = "220" Then
            analyse_fenster.Label5.BackColor = &HFF00&
        Else
            analyse_fenster.Label5.BackColor = 255
        End If
        If Mid$(consoleText, 5, 3) = "226" Then
            analyse_fenster.Label6.BackColor = &HFF00&
        Else
            analyse_fenster.Label6.BackColor = 255
        End If
        If Mid$(consoleText, 5, 3) = "230" Then
            analyse_fenster.Label7.BackColor = &HFF00&
        Else
            analyse_fenster.Label7.BackColor = 255
        End If
        If Mid$(consoleText, 5, 3) = "250" Then
            analyse_fenster.Label8.BackColor = &HFF00&
        Else
            analyse_fenster.Label8.BackColor = 255
        End If
        If Mid$(consoleText, 5, 3) = "331" Then
            analyse_fenster.Label9.BackColor = &HFF00&
        Else
            analyse_fenster.Label9.BackColor = 255
        End If
        If Mid$(consoleText, 5, 3) = "257" Then
            analyse_fenster.Label10.BackColor = &HFF00&
        Else
            analyse_fenster.Label10.BackColor = 255
        End If
        If Mid$(consoleText, 5, 3) = "350" Then
            analyse_fenster.Label1.BackColor = &HFF00&
        Else
            analyse_fenster.Label1.BackColor = 255
        End If
    Else
        analyse_fenster.console.Text = analyse_fenster.console.Text & consoleText & vbCrLf & vbCrLf
    End If
End If
End Sub

Public Sub CheckNameIfBad(StringToCheck, WhatDataToCheck As String)
For X = 1 To Len(StringToCheck)
    If Mid$(StringToCheck, X, 1) = "*" Or Mid$(StringToCheck, X, 1) = "/" Or Mid$(StringToCheck, X, 1) = "\" Or Mid$(StringToCheck, X, 1) = "|" Or Mid$(StringToCheck, X, 1) = "<" Or Mid$(StringToCheck, X, 1) = ">" Or Mid$(StringToCheck, X, 1) = ":" Or Mid$(StringToCheck, X, 1) = "?" Or Mid$(StringToCheck, X, 1) = Chr(34) Then 'chr(34)= "
        If WhatDataToCheck = "Ordner" And HideWarningsForBadNames = False Then If MsgBox("Das Zeichen '" & Mid$(StringToCheck, X, 1) & "' wird vom lokalen Dateisystem nicht akzeptiert! Es wird aus dem Ordnernamen entfernt! Soll diese Warnung weiterhin angezeigt werden?", vbQuestion + vbYesNo, "Unkonventioneller Ordnername!") = vbNo Then HideWarningsForBadNames = True
        If WhatDataToCheck = "Datei" And HideWarningsForBadNames = False Then If MsgBox("Das Zeichen '" & Mid$(StringToCheck, X, 1) & "' wird vom lokalen Dateisystem nicht akzeptiert! Es wird aus dem Dateinamen entfernt! Soll diese Warnung weiterhin angezeigt werden?", vbQuestion + vbYesNo, "Unkonventioneller Dateiname!") = vbNo Then HideWarningsForBadNames = True
        StringToCheck = Mid$(StringToCheck, 1, X - 1) & Mid$(StringToCheck, X + 1)
        X = X - 1
    End If
Next X
If StringToCheck = "" Then
    CheckedString = "Bad_Dir_Name"
Else
    CheckedString = StringToCheck
End If
End Sub

Public Sub ServerpfadNeuDefinieren(NeuerServerPfad As String)
    If NeuerServerPfad = "//DirStatementOk" Then
        If zwischen_ServerPath = "//DirUp" Then
            zwischen = ""
            For X = 1 To Len(ParentServerPath) - 1
                If Mid$(ParentServerPath, X, 1) = "/" Then Y = X
            Next X
            ParentServerPath = Mid$(ParentServerPath, 1, Y)
            If ParentServerPath = "" Then ParentServerPath = "/"
        Else
            ParentServerPath = ParentServerPath & zwischen_ServerPath & "/"
        End If
        zwischen_ServerPath = ""
    Else
        zwischen_ServerPath = NeuerServerPfad
    End If
End Sub

Public Sub StartProgressVisualisation(UpOrDown As String)
If ShowDLProgress = True Then
    ShowDLProgress = False
    Label1.Visible = False
    DateigroesseLocalVisual.Visible = True
    Label2.Visible = False
    ProcessDateiname.Visible = False
    ProcessAktStand.Visible = False
    ProcessDateigroesse.Visible = False
    ProcessProgress.Visible = False
    Command1.Visible = False
    main.WindowState = 0
Else
    ShowDLProgress = True
    main.WindowState = 0
    ProcessProgress.Visible = True
    ProcessProgress = 0
    Command1.Visible = True
    
    ZahlOrdnen (Dateigroesse)
    ProcessDateigroesse.Caption = GeordneteZahl
       
    Label1.Visible = True
    DateigroesseLocalVisual.Visible = False
    Label2.Visible = True
    ProcessDateiname.Visible = True
    ProcessAktStand.Visible = True
    ProcessDateigroesse.Visible = True
    If UpOrDown = "down" Then ProcessDateiname.Caption = "'" & DownloadFilename & "' wird downgeloaded..."
    If UpOrDown = "up" Then ProcessDateiname.Caption = "'" & UploadFilename & "' wird upgeloaded..."
End If
End Sub

Public Sub RefreshForm()
VD_bytesReceived.Refresh
VD_bytesSent.Refresh
ProcessAktStand.Refresh
VD_bytesSent.Refresh
VD_bytesReceived.Refresh
End Sub

Public Sub ZahlOrdnen(übergeben)
Y = 1
For X = Len(übergeben) To 1 Step -1
    If (Len(übergeben) - X + Y) Mod 3 = 0 Then
        Y = Y - 1
        übergeben = Mid$(übergeben, 1, X - 1) & " " & Mid$(übergeben, X)
    End If
Next X
Do Until Len(übergeben) > 15
    übergeben = " " & übergeben
Loop
GeordneteZahl = übergeben
End Sub
