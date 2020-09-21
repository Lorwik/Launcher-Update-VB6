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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterAOLauncher.ucAsyncDLHost ucAsyncDLHost 
      Height          =   3015
      Left            =   6120
      TabIndex        =   1
      Top             =   2400
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5318
   End
   Begin VB.Image cmdSalir 
      Height          =   285
      Left            =   10560
      Picture         =   "frmMain.frx":30819
      Top             =   600
      Width           =   240
   End
   Begin VB.Image cmdJugar 
      Height          =   930
      Left            =   7800
      Picture         =   "frmMain.frx":308CE
      Top             =   6600
      Width           =   3030
   End
   Begin WinterAOLauncher.ucAsyncDLStripe ucAsyncDLStripe 
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
   End
   Begin VB.Label lblPendientes 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   4560
      Width           =   4365
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

Private Sub ucAsyncDLHost_DownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
  Debug.Print "DownloadComplete for URL: "; Sender.URL & "; Directorio: " & Sender.LocalFileName
  
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
