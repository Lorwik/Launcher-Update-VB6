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

Private mUrl              As String, mLocalFileName As String, mStartDate As Date

Public Sub DownloadFile(URL As String, LocalFileName As String, Optional ByVal Mode As AsyncReadConstants = vbAsyncReadForceUpdate)
        
        On Error GoTo DownloadFile_Err
        
100     CancelDownload
102     mUrl = URL
104     mLocalFileName = LocalFileName
106     mStartDate = Now
108     AsyncRead mUrl, vbAsyncTypeFile, mLocalFileName, Mode
110     Extender.ToolTipText = mLocalFileName
        
        Exit Sub

DownloadFile_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.DownloadFile", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Public Sub CancelDownload()
    shpProgress.Visible = False

    On Error Resume Next

    CancelAsyncRead mLocalFileName 'cancel a possibly still running Download with the same Destination-Filename

    On Error GoTo 0

End Sub

Public Property Get URL() As String
        
        On Error GoTo URL_Err
        
100     URL = mUrl
        
        Exit Property

URL_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.URL", _
                  "ucAsyncDLStripe component failure"
        
End Property

Public Property Get LocalFileName() As String
        
        On Error GoTo LocalFileName_Err
        
100     LocalFileName = mLocalFileName
        
        Exit Property

LocalFileName_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.LocalFileName", _
                  "ucAsyncDLStripe component failure"
        
End Property

Public Property Get StartDate() As Date
        
        On Error GoTo StartDate_Err
        
100     StartDate = mStartDate
        
        Exit Property

StartDate_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.StartDate", _
                  "ucAsyncDLStripe component failure"
        
End Property

Public Property Get Caption() As String
        
        On Error GoTo Caption_Err
        
100     Caption = lblCaption.Caption
        
        Exit Property

Caption_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.Caption", _
                  "ucAsyncDLStripe component failure"
        
End Property

Public Property Let Caption(ByVal NewValue As String)
        
        On Error GoTo Caption_Err
        
100     lblCaption.Caption = NewValue
        
        Exit Property

Caption_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.Caption", _
                  "ucAsyncDLStripe component failure"
        
End Property
 
Private Sub lblCancelResume_Click()
        
        On Error GoTo lblCancelResume_Click_Err
        

100     If lblCancelResume.Caption = "Resume" Then
102         lblCancelResume.Caption = "Stop"
104         lblCancelResume.ForeColor = &HAA00&
106         DownloadFile mUrl, mLocalFileName
        Else
108         lblCancelResume.Caption = "Resume"
110         lblCancelResume.ForeColor = &H88EE&
112         CancelDownload
        End If

        
        Exit Sub

lblCancelResume_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.lblCancelResume_Click", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub lblCancelResume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo lblCancelResume_MouseDown_Err
        
100     mDown = True
        
        Exit Sub

lblCancelResume_MouseDown_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.lblCancelResume_MouseDown", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub lblCancelResume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo lblCancelResume_MouseMove_Err
        
100     GetCursorPos TL
102     TL.X = TL.X - X / Screen.TwipsPerPixelX
104     TL.Y = TL.Y - Y / Screen.TwipsPerPixelX

106     If tHover Is Nothing Then
108         Set tHover = Controls.Add("VB.Timer", "tHover")
110         tHover.Interval = 20
        End If

        
        Exit Sub

lblCancelResume_MouseMove_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.lblCancelResume_MouseMove", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub lblCancelResume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo lblCancelResume_MouseUp_Err
        
100     mDown = False
        
        Exit Sub

lblCancelResume_MouseUp_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.lblCancelResume_MouseUp", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub tHover_Timer()
    
    On Error Resume Next
    

    Dim Pt As PointAPI, OutSide As Boolean

    GetCursorPos Pt
    OutSide = Pt.X < TL.X Or Pt.Y < TL.Y Or Pt.X >= TL.X + lblCancelResume.Width Or Pt.Y >= TL.Y + lblCancelResume.Height

    If mDown And Not OutSide Then
        lblCancelResume.Move shpCancelResume.Left + 1, shpCancelResume.Top + 2
        shpCancelResume.FillColor = &HA0A0A0
        shpCancelResume.BorderColor = vbBlack
    ElseIf OutSide Then
        lblCancelResume.Move shpCancelResume.Left, shpCancelResume.Top + 1
        shpCancelResume.FillColor = &HC0C0C0
        shpCancelResume.BorderColor = &H808080
        Set tHover = Nothing
        Controls.Remove "tHover"
    Else
        lblCancelResume.Move shpCancelResume.Left, shpCancelResume.Top + 1
        shpCancelResume.FillColor = &HE0E0E0
        shpCancelResume.BorderColor = &H808080
    End If

End Sub

Private Sub UserControl_Initialize()
        
        On Error GoTo UserControl_Initialize_Err
        
100     ScaleMode = vbPixels
        
        Exit Sub

UserControl_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.UserControl_Initialize", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo UserControl_MouseMove_Err
        
100     Extender.ToolTipText = IIf(X < (lblRemove.Left + lblRemove.Width), "Remove Download-Job", mUrl)
        
        Exit Sub

UserControl_MouseMove_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.UserControl_MouseMove", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo UserControl_MouseUp_Err
        

100     If X < (lblRemove.Left + lblRemove.Width) Then Parent.RemoveByLocalFileName mLocalFileName
        
        Exit Sub

UserControl_MouseUp_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.UserControl_MouseUp", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
    
    shpCancelResume.Left = ScaleWidth - shpCancelResume.Width * 1.1
    lblCancelResume.Move shpCancelResume.Left, shpCancelResume.Top + 1, shpCancelResume.Width, shpCancelResume.Height
    shpProgressBase.Width = shpCancelResume.Left - 1.4 * shpProgressBase.Left
    lblCaption.Move shpProgressBase.Left, shpProgressBase.Top + 1, shpProgressBase.Width, shpProgressBase.Height
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
        
        On Error GoTo UserControl_AsyncReadProgress_Err
        
100     Parent.RaiseDownloadProgress Me, AsyncProp.BytesRead, AsyncProp.BytesMax
102     shpProgress.Visible = AsyncProp.BytesMax

104     If AsyncProp.BytesMax = 0 Then Exit Sub

106     With shpProgressBase
108         shpProgress.Move .Left + 1, .Top + 1, (.Width - 2) * AsyncProp.BytesRead / AsyncProp.BytesMax, .Height - 2
        End With

        
        Exit Sub

UserControl_AsyncReadProgress_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.UserControl_AsyncReadProgress", _
                  "ucAsyncDLStripe component failure"
        
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
        
        On Error GoTo UserControl_AsyncReadComplete_Err
        

100     If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData Or AsyncProp.BytesRead = 0 Then
102         Parent.RaiseDownloadError Me, AsyncProp.StatusCode, AsyncProp.Status
104         CancelDownload
        Else
106         Parent.RaiseDownloadComplete Me, AsyncProp.value
108         Parent.RemoveByLocalFileName mLocalFileName 'let's remove ourselves from the List in the Parent-Control
        End If

        
        Exit Sub

UserControl_AsyncReadComplete_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLStripe.UserControl_AsyncReadComplete", _
                  "ucAsyncDLStripe component failure"
        
End Sub
 
