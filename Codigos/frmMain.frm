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
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
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
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "#0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   6510
      Width           =   570
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   255
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   615
   End
   Begin VB.Image imgOpciones 
      Height          =   690
      Left            =   5520
      Picture         =   "frmMain.frx":4A1B5
      Top             =   6790
      Width           =   2250
   End
   Begin VB.Image cmdSalir 
      Height          =   285
      Left            =   10560
      Picture         =   "frmMain.frx":4BED7
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image cmdJugar 
      Height          =   930
      Left            =   7800
      Picture         =   "frmMain.frx":4BF8C
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
        If Len(Fallaron) > 0 Then
            MsgBox "No se ha podido comprobar la integridad de uno o varios archivos es posible que no se haya podido descargar correctamente. Archivos: " & Fallaron
            
        Else
            If FileExist(App.Path & "\WinterAO Resurrection.exe", vbNormal) Then
                Call WriteVar(App.Path & "\INIT\Config.ini", "PARAMETERS", "LAUCH", "1")
                DoEvents
                Call Shell(App.Path & "\WinterAO Resurrection.exe", vbNormalFocus)
                End
                
            Else
                MsgBox "No se encontro el ejecutable del juego ""0Winter AO Ultimate.EXE""."
                
            End If
            
        End If
    End If
        
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub imgOpciones_Click()
    frmOpciones.Show
End Sub

Private Sub ucAsyncDLHost_DownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
'**********************************************************
'Descripcion: Evento cuando termina una descarga
'**********************************************************

    Static finalizados As Integer

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
            Fallaron = Fallaron + Sender.LocalFileName & ", "
        End If
    End If
  
    finalizados = finalizados + 1
    
    If finalizados >= Desactualizados Then _
        frmMain.lblPendientes.Caption = "Cliente actualizado."
  
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
