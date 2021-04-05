VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launcher ComunidadWinter"
   ClientHeight    =   4170
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   2520
      TabIndex        =   8
      Top             =   3840
      Width           =   990
   End
   Begin VB.Frame FraSeleccionaUn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Selecciona un Server de Comunidad Winter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton OptServer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imperium Clasico"
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton OptServer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "WinterAO Resurrection"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdJugar 
      Caption         =   "Jugar"
      Height          =   480
      Left            =   3960
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   480
      Left            =   360
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":000C
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2400
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblPendientes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizado"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   6225
   End
   Begin VB.Label LSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0 MBs de 0 MBs"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&
Dim f As Integer

Private Sub cmdCommand1_Click()
Debug.Print SinVersiones
Debug.Print ActualizacionesPendientes
End Sub

Private Sub cmdJugar_Click()
    Dim Integridad As Integer
    
    '¿Hay actualizaciones pendientes?
    If ActualizacionesPendientes Then
        modUpdate.ActualizarCliente
        
    Else 'Si esta todo actualizado...
    
        If Len(Fallaron) > 0 Then '¿Se actualizo antes? ¿Hubo fallos?
            MsgBox "No se ha podido comprobar la integridad de uno o varios archivos es posible que no se haya podido descargar correctamente. Archivos: " & Fallaron
            frmMain.lblPendientes = "No se ha podido comprobar la integridad de uno o varios archivos es posible que no se haya podido descargar correctamente. Archivos: " & Fallaron
            
        Else 'Si todo esta OK, lanzamos el juego
        
            'Ante de lanzar el cliente vamos a verificar todos los archivos
            Integridad = ComprobarIntegridad
            
            frmMain.lblPendientes.Caption = "Comprobando integridad de archivos..."
            Debug.Print "Comprobando integridad de archivos..."
        
            If Integridad <> 0 Then
                
                MsgBox "No se ha podido comprobar la integridad de " & Integridad & " archivos. Pulsa Jugar para volver a descargarlo. Si el problema persiste, revise su conexión a internet o contacte con los administradores del juego."
                frmMain.lblPendientes = "No se ha podido comprobar la integridad de " & Integridad & " archivos. Pulsa Jugar para volver a descargarlo. Si el problema persiste, revise su conexión a internet o contacte con los administradores del juego."
                
            Else
        
                If FileExist(App.Path & "\" & CLIENTE_FOLDER & "\" & CLIENTEXE, vbNormal) Then  '¿Existe el .exe del cliente?
                    Call WriteVar(App.Path & "\" & CLIENTE_FOLDER & "\INIT\Config.ini", "PARAMETERS", "LAUCH", "1")
                    DoEvents
                    Call Shell(App.Path & "\" & CLIENTE_FOLDER & "\" & CLIENTEXE, vbNormalFocus)
                    
                    End
                    
                Else 'Si no existe, no podemos lanzar nada
                    MsgBox "No se encontro el ejecutable del juego " & CLIENTEXE
                    
                End If
            End If
            
        End If
    End If
End Sub

Private Sub cmdCerrar_Click()
    End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

    Dim Progreso As Long

    Select Case State
        Case icError
            Call addConsole("Error en la conexión, descarga abortada.", 255, 0, 0, True, False)
            bDone = True
            dError = True
            
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long
            
            FileSize = Inet1.GetHeader("Content-length")

            LSize.Visible = True

            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                    
                    vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    Progreso = Progreso + Len(vtData) * 2
                    LSize.Caption = Round((Progreso + Len(vtData) * 2) / 1000000, 2) & "MBs de " & Round((FileSize / 1000000), 2) & "MBs"

                    DoEvents
                Loop
            Close #1
            
            Call addConsole("Descarga finalizada.", 0, 255, 0, True, False)
            LSize.Caption = FileSize & "bytes"
            LSize.Visible = False
            
            bDone = True
    End Select
    
End Sub

Private Sub OptServer_Click(Index As Integer)

    ServerSelect = Index
    
    Call SetURLModo
    Call IniciarChequeo
    
    Call WriteVar(App.Path & "\Init\Config.ini", "INIT", "Select", ServerSelect)
    
End Sub
