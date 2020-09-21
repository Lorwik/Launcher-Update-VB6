VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "WinterAO Launcher"
   ClientHeight    =   7680
   ClientLeft      =   -60
   ClientTop       =   -30
   ClientWidth     =   11400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1A041
   ScaleHeight     =   7680
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterAOLauncher.ucAsyncDLHost ucAsyncDLHost 
      Height          =   3615
      Left            =   6100
      TabIndex        =   1
      Top             =   2700
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5318
   End
   Begin VB.Image cmdSalir 
      Height          =   285
      Left            =   10560
      Picture         =   "frmMain.frx":4A1B5
      Top             =   600
      Width           =   240
   End
   Begin VB.Image cmdJugar 
      Height          =   930
      Left            =   7800
      Picture         =   "frmMain.frx":4A26A
      Top             =   6600
      Width           =   3030
   End
   Begin WinterAOLauncher.ucAsyncDLStripe ucAsyncDLStripe 
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
   End
   Begin VB.Label lblPendientes 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6360
      TabIndex        =   0
      Top             =   3120
      Width           =   4365
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Skin Me, vbRed
End Sub

Private Sub cmdJugar_Click()
        
    If ActualizacionesPendientes Then
        ModUpdate.ActualizarCliente
    Else
        If FileExist(App.Path & "\WinterAO Resurrection.exe", vbNormal) Then
            DoEvents
            Call Shell(App.Path & "\WinterAO Resurrection.exe", vbNormalFocus)
            End
        Else
            MsgBox "No se encontro el ejecutable del juego ""0Winter AO Ultimate.EXE""."
            End
        End If
    End If
        
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub ucAsyncDLHost_DownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
'**********************************************************
'Descripcion: Evento cuando termina una descarga
'**********************************************************

    'Debug.Print "DownloadComplete for URL: "; Sender.URL & "; Directorio: " & Sender.LocalFileName
    
    'the complete-event delivers a temporary filename - it is up to the user
    'of the Control, to decide what to do with this TmpFile... the usual reaction will be
    'a simple File-Renaming (ensuring an implicit Move-Operation on the FileSystem then)
    'the Sender is the Control-Stripe of our DownloadListHost-Control - and in the Add-methods
    'in the Form-Load-Event above, we have defined a "target-LocalFilename" already, which is
    'associated with the matching URL (which was the Source of this completed Download here)
    'This target-filename is (so far) only stored as String within the Stripe (Sender.LocalFileName)
    
    'So, yeah - just ensure a proper Move/Rename of the delivered TmpFileName
    
    If FileExist(Sender.LocalFileName, vbNormal) Then Kill Sender.LocalFileName
  
    Name TmpFileName As Sender.LocalFileName
  
    If ComprobarHash(Sender.LocalFileName) = False Then  'Si el Hash no coincide...
    
        'Chapuza que hay que cambiar
        If Sender.LocalFileName <> App.Path & "\INIT\VersionInfo.json" Then
            MsgBox "No se ha podido comprobar la integridad del archivo " & Sender.LocalFileName & " es posible que no se haya podido descargar correctamente."
        End If
    End If
  
End Sub
 
Private Sub ucAsyncDLHost_DownloadProgress(Sender As ucAsyncDLStripe, ByVal BytesRead As Long, ByVal BytesTotal As Long)
  Sender.Caption = FormatBytes2KBMBGBTB(BytesRead) & " (" & FormatDLRate(BytesRead, DateDiff("s", Sender.StartDate, Now)) & ")"
End Sub

Function FormatBytes2KBMBGBTB(ByVal Bytes As Currency) As String
Dim i As Long
  Do While Bytes >= 1024: Bytes = Bytes / 1024: i = i + 1: Loop
  FormatBytes2KBMBGBTB = Int(Bytes * 10) / 10 & Split(",K,M,G,T", ",")(i) & "B"
End Function

Function FormatDLRate(ByVal Bytes As Long, ByVal Seconds As Long) As String
  If Seconds Then FormatDLRate = FormatBytes2KBMBGBTB(Bytes \ Seconds) & "/s"
End Function
