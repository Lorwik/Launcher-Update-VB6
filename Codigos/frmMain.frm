VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   Caption         =   "WinterAO Resurrection - Launcher"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12645
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   6960
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdJugar 
      Caption         =   "Jugar"
      Height          =   1080
      Left            =   9240
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblPendientes 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizaciones pendientes: "
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   11400
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
