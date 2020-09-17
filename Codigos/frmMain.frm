VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "WinterAO Resurrection - Launcher"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11295
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image cmdSalir 
      Height          =   285
      Left            =   10440
      Picture         =   "frmMain.frx":30819
      Top             =   600
      Width           =   240
   End
   Begin VB.Image cmdJugar 
      Height          =   930
      Left            =   7560
      Picture         =   "frmMain.frx":308CE
      Top             =   6600
      Width           =   3030
   End
   Begin VB.Label lblPendientes 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizaciones pendientes: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   4560
      Width           =   9960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000

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

Sub Skin(Frm As Form, Color As Long)
    Frm.BackColor = Color
    Dim Ret As Long
    Ret = GetWindowLong(Frm.hwnd, G_E)
    Ret = Ret Or W_E
    SetWindowLong Frm.hwnd, G_E, Ret
    SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY
End Sub
