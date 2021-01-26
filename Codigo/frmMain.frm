VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comunidad Winter Launcher"
   ClientHeight    =   3480
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdJugar 
      Caption         =   "Jugar"
      Height          =   480
      Left            =   3960
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   480
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmMain.frx":000C
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblPendientes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizado"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   6225
   End
   Begin VB.Label LSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0 MBs de 0 MBs"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2280
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

Private Const CLIENTEXE As String = "\WinterAOResurrection.exe"

Private Sub cmdJugar_Click()
    Dim Integridad As Integer
    
    Debug.Print "Iniciando. Actualizaciones pendientes: " & ActualizacionesPendientes
    
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
        
                If FileExist(App.Path & CLIENTEXE, vbNormal) Then '¿Existe el .exe del cliente?
                    Call WriteVar(App.Path & "\INIT\Config.ini", "PARAMETERS", "LAUCH", "1")
                    DoEvents
                    Call Shell(App.Path & CLIENTEXE, vbNormalFocus)
                    
                    End
                    
                Else 'Si no existe, no podemos lanzar nada
                    MsgBox "No se encontro el ejecutable del juego WinterAOUltimate.exe"
                    
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
