VERSION 5.00
Begin VB.UserControl ucAsyncDLStripe 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   219
   Windowless      =   -1  'True
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   60
      Width           =   225
   End
   Begin VB.Label lblRemove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "þ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Remove Download-Job"
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   195
   End
   Begin VB.Label lblCancelResume 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000AA00&
      Height          =   285
      Left            =   2490
      TabIndex        =   0
      Top             =   60
      Width           =   765
   End
   Begin VB.Shape shpCancelResume 
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   735
   End
   Begin VB.Shape shpProgress 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   450
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   255
   End
   Begin VB.Shape shpProgressBase 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   405
   End
End
Attribute VB_Name = "ucAsyncDLStripe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'a windowless Control-Stripe which ensures a spearate (async) Download,
'"Download-stripes of this Control-Type here are dynamically added later on in ucAsyncDLHost
Option Explicit

Private Type PointAPI
  X As Long
  Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private WithEvents tHover As VB.Timer, mDown As Boolean, TL As PointAPI
Attribute tHover.VB_VarHelpID = -1
Private mUrl As String, mLocalFileName As String, mStartDate As Date

Public Sub DownloadFile(URL As String, LocalFileName As String, Optional ByVal Mode As AsyncReadConstants = vbAsyncReadForceUpdate)
  CancelDownload
  mUrl = URL
  mLocalFileName = LocalFileName
  mStartDate = Now
  AsyncRead mUrl, vbAsyncTypeFile, mLocalFileName, Mode
  'Extender.ToolTipText = mUrl
End Sub

Public Sub CancelDownload()
  shpProgress.Visible = False
  On Error Resume Next
    CancelAsyncRead mLocalFileName 'cancel a possibly still running Download with the same Destination-Filename
  On Error GoTo 0
End Sub

Public Property Get URL() As String
  URL = mUrl
End Property

Public Property Get LocalFileName() As String
  LocalFileName = mLocalFileName
End Property

Public Property Get StartDate() As Date
  StartDate = mStartDate
End Property

Public Property Get Caption() As String
  Caption = lblCaption.Caption
End Property
Public Property Let Caption(ByVal NewValue As String)
  lblCaption.Caption = NewValue
End Property
 
Private Sub lblCancelResume_Click()
  If lblCancelResume.Caption = "Resume" Then
    lblCancelResume.Caption = "Stop": lblCancelResume.ForeColor = &HAA00&
    DownloadFile mUrl, mLocalFileName
  Else
    lblCancelResume.Caption = "Resume": lblCancelResume.ForeColor = &H88EE&
    CancelDownload
  End If
End Sub

Private Sub lblCancelResume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mDown = True
End Sub
Private Sub lblCancelResume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  GetCursorPos TL: TL.X = TL.X - X / Screen.TwipsPerPixelX: TL.Y = TL.Y - Y / Screen.TwipsPerPixelX
  If tHover Is Nothing Then Set tHover = Controls.Add("VB.Timer", "tHover"): tHover.Interval = 20
End Sub
Private Sub lblCancelResume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mDown = False
End Sub

Private Sub tHover_Timer()
Dim Pt As PointAPI, OutSide As Boolean
  GetCursorPos Pt
  OutSide = Pt.X < TL.X Or Pt.Y < TL.Y Or Pt.X >= TL.X + lblCancelResume.Width Or Pt.Y >= TL.Y + lblCancelResume.Height

  If mDown And Not OutSide Then
    lblCancelResume.Move shpCancelResume.Left + 1, shpCancelResume.Top + 2
    shpCancelResume.FillColor = &HA0A0A0: shpCancelResume.BorderColor = vbBlack
  ElseIf OutSide Then
    lblCancelResume.Move shpCancelResume.Left, shpCancelResume.Top + 1
    shpCancelResume.FillColor = &HC0C0C0: shpCancelResume.BorderColor = &H808080
    Set tHover = Nothing
    Controls.Remove "tHover"
  Else
    lblCancelResume.Move shpCancelResume.Left, shpCancelResume.Top + 1
    shpCancelResume.FillColor = &HE0E0E0: shpCancelResume.BorderColor = &H808080
  End If
End Sub

Private Sub UserControl_Initialize()
  ScaleMode = vbPixels
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Extender.ToolTipText = IIf(X < (lblRemove.Left + lblRemove.Width), "Remove Download-Job", mUrl)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If X < (lblRemove.Left + lblRemove.Width) Then Parent.RemoveByLocalFileName mLocalFileName
End Sub

Private Sub UserControl_Resize()
  shpCancelResume.Left = ScaleWidth - shpCancelResume.Width * 1.1
  lblCancelResume.Move shpCancelResume.Left, shpCancelResume.Top + 1, shpCancelResume.Width, shpCancelResume.Height
  shpProgressBase.Width = shpCancelResume.Left - 1.4 * shpProgressBase.Left
  lblCaption.Move shpProgressBase.Left, shpProgressBase.Top + 1, shpProgressBase.Width, shpProgressBase.Height
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
  Parent.RaiseDownloadProgress Me, AsyncProp.BytesRead, AsyncProp.BytesMax
  shpProgress.Visible = AsyncProp.BytesMax
  If AsyncProp.BytesMax = 0 Then Exit Sub
  With shpProgressBase
    shpProgress.Move .Left + 1, .Top + 1, (.Width - 2) * AsyncProp.BytesRead / AsyncProp.BytesMax, .Height - 2
  End With
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
  If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData Or AsyncProp.BytesRead = 0 Then
    Parent.RaiseDownloadError Me, AsyncProp.StatusCode, AsyncProp.Status
    CancelDownload
  Else
    Parent.RaiseDownloadComplete Me, AsyncProp.value
    Parent.RemoveByLocalFileName mLocalFileName 'let's remove ourselves from the List in the Parent-Control
  End If
End Sub
 
