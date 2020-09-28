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
      Height          =   2955
      Left            =   6360
      TabIndex        =   0
      Top             =   3000
      Width           =   4365
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLIENTEXE As String = "\WinterAOResurrection.exe"

Private Sub Form_Load()
        
        On Error GoTo Form_Load_Err
        
100     Skin Me, vbRed
        
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.Form_Load " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Private Sub cmdJugar_Click()
        
        On Error GoTo cmdJugar_Click_Err
        
        
        Dim Integridad As Integer
    
        '¿Hay actualizaciones pendientes?
100     If ActualizacionesPendientes Or LauncherDesactualizado Then
102         ModUpdate.ActualizarCliente
        
        Else 'Si esta todo actualizado...
    
104         If Len(Fallaron) > 0 Then '¿Se actualizo antes? ¿Hubo fallos?
106             MsgBox "No se ha podido comprobar la integridad de uno o varios archivos es posible que no se haya podido descargar correctamente. Archivos: " & Fallaron
108             frmMain.lblPendientes = "No se ha podido comprobar la integridad de uno o varios archivos es posible que no se haya podido descargar correctamente. Archivos: " & Fallaron
            
            Else 'Si todo esta OK, lanzamos el juego
        
                'Ante de lanzar el cliente vamos a verificar todos los archivos
110             Integridad = ComprobarIntegridad
        
112             If Integridad <> 0 Then
                
114                 MsgBox "No se ha podido comprobar la integridad de " & Integridad & " archivos. Pulsa Jugar para volver a descargarlo. Si el problema persiste, revise su conexión a internet o contacte con los administradores del juego."
116                 frmMain.lblPendientes = "No se ha podido comprobar la integridad de " & Integridad & " archivos. Pulsa Jugar para volver a descargarlo. Si el problema persiste, revise su conexión a internet o contacte con los administradores del juego."
                
                Else
        
118                 If FileExist(App.Path & CLIENTEXE, vbNormal) Then '¿Existe el .exe del cliente?
120                     Call WriteVar(App.Path & "\INIT\Config.ini", "PARAMETERS", "LAUCH", "1")
122                     DoEvents
124                     Call Shell(App.Path & CLIENTEXE, vbNormalFocus)

126                     End
                    
                    Else 'Si no existe, no podemos lanzar nada
128                     MsgBox "No se encontro el ejecutable del juego WinterAOUltimate.exe"
                    
                    End If
                End If
            
            End If
        End If
        
        
        Exit Sub

cmdJugar_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.cmdJugar_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Private Sub cmdSalir_Click()
        
        On Error GoTo cmdSalir_Click_Err
        

100     End

        
        Exit Sub

cmdSalir_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.cmdSalir_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Private Sub imgOpciones_Click()
        
        On Error GoTo imgOpciones_Click_Err
        
100     frmOpciones.Show
        
        Exit Sub

imgOpciones_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.imgOpciones_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Private Sub ucAsyncDLHost_DownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
        '**********************************************************
        'Descripcion: Evento cuando termina una descarga
        '**********************************************************
        
        On Error GoTo ucAsyncDLHost_DownloadComplete_Err
        

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
    
100     If FileExist(Sender.LocalFileName, vbNormal) Then Kill Sender.LocalFileName
  
102     Name TmpFileName As Sender.LocalFileName
    
104     If Sender.LocalFileName <> LAUNCHEREXEUP Then

            'Si el Hash no coincide...
106         If ComprobarHash(Sender.LocalFileName) = False Then
108             Fallaron = Fallaron + Sender.LocalFileName & ", "
110             Call LauncherLog("Se descargo " & Sender.LocalFileName & " pero falló.")
            
            Else
112             Call LauncherLog("Se descargo " & Sender.LocalFileName & " correctamente")
            
            End If
        
114         finalizados = finalizados + 1
        
        Else 'Si es una actualizacion del Launcher...
        
116         ComprobarHash (Sender.LocalFileName)
118         LauncherDesactualizado = False
        
        End If
    
120     If finalizados >= Desactualizados Then
122         frmMain.lblPendientes.Caption = "Cliente actualizado."
        
124         Desactualizados = 0
126         ReDim DesactualizadosList(Desactualizados) As tArchivos
        
        End If
  
        
        Exit Sub

ucAsyncDLHost_DownloadComplete_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.ucAsyncDLHost_DownloadComplete " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub
 
Private Sub ucAsyncDLHost_DownloadProgress(Sender As ucAsyncDLStripe, ByVal BytesRead As Long, ByVal BytesTotal As Long)
        
        On Error GoTo ucAsyncDLHost_DownloadProgress_Err
        
100     Sender.Caption = FormatBytes2KBMBGBTB(BytesRead) & " (" & FormatDLRate(BytesRead, DateDiff("s", Sender.StartDate, Now)) & ")"
        
        Exit Sub

ucAsyncDLHost_DownloadProgress_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.ucAsyncDLHost_DownloadProgress " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Function FormatBytes2KBMBGBTB(ByVal Bytes As Currency) As String
        
        On Error GoTo FormatBytes2KBMBGBTB_Err
        

        Dim i As Long

100     Do While Bytes >= 1024
102         Bytes = Bytes / 1024
104         i = i + 1
        Loop
106     FormatBytes2KBMBGBTB = Int(Bytes * 10) / 10 & Split(",K,M,G,T", ",")(i) & "B"
        
        Exit Function

FormatBytes2KBMBGBTB_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.FormatBytes2KBMBGBTB " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Function FormatDLRate(ByVal Bytes As Long, ByVal Seconds As Long) As String
        
        On Error GoTo FormatDLRate_Err
        

100     If Seconds Then FormatDLRate = FormatBytes2KBMBGBTB(Bytes \ Seconds) & "/s"
        
        Exit Function

FormatDLRate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.frmMain.FormatDLRate " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function
