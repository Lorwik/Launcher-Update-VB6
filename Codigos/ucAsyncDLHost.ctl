VERSION 5.00
Begin VB.UserControl ucAsyncDLHost 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "ucAsyncDLHost.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   795
      Left            =   4470
      Max             =   1
      Min             =   1
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Value           =   1
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "ucAsyncDLHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'just the Host-Container which ensures a simple Scrolling for the separate (windowless) Download-Control-Stripes
Option Explicit

Event DownloadProgress(Sender As ucAsyncDLStripe, ByVal BytesRead As Long, ByVal BytesTotal As Long)
Event DownloadError(Sender As ucAsyncDLStripe, ByVal StatusCode As Long, ByVal Status As String)
Event DownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
 
Public Sub AddDownloadJob(URL As String, LocalFileName As String)
        
        On Error GoTo AddDownloadJob_Err
        

        Dim NewStripe As ucAsyncDLStripe

        Static CC     As Currency
100     CC = CC + 1
  
102     If Len(GetCtlKeyForLocalFileName(LocalFileName)) Then Err.Raise vbObjectError, , "Ya hay una descarga con ese nombre de archivo local en la lista"
104     If Len(GetCtlKeyForURL(URL)) Then Err.Raise vbObjectError, , "Ya hay una Descarga con esa URL en la Lista"
 
106     With Controls.Add(GetProjectName & ".ucAsyncDLStripe", "K" & CC)
108         .Move 0, Controls.Count * .Height, ScaleWidth - VScroll.Width
110         .Visible = True
    
112         VScroll_Change
 
114         Set NewStripe = .Object 'just a cast to the concrete Control-Interface
116         NewStripe.DownloadFile URL, LocalFileName 'and here we trigger the Download
        End With

        
        Exit Sub

AddDownloadJob_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.AddDownloadJob", _
                  "ucAsyncDLHost component failure"
        
End Sub

Public Sub RemoveByLocalFileName(ByVal LocalFileName As String)
        
        On Error GoTo RemoveByLocalFileName_Err
        

        Dim CtlKey As String

100     CtlKey = GetCtlKeyForLocalFileName(LocalFileName)

102     If Len(CtlKey) = 0 Then Exit Sub
104     Controls.Remove CtlKey
106     VScroll_Change
        
        Exit Sub

RemoveByLocalFileName_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.RemoveByLocalFileName", _
                  "ucAsyncDLHost component failure"
        
End Sub

Public Function GetCtlKeyForURL(ByVal URL As String) As String
        
        On Error GoTo GetCtlKeyForURL_Err
        

        Dim i As Long

100     For i = 1 To Controls.Count - 1

102         If StrComp(Controls(i).URL, URL, vbTextCompare) = 0 Then
104             GetCtlKeyForURL = Controls(i).name
                Exit For
            End If

106     Next i

        
        Exit Function

GetCtlKeyForURL_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.GetCtlKeyForURL", _
                  "ucAsyncDLHost component failure"
        
End Function

Public Function GetCtlKeyForLocalFileName(ByVal LocalFileName As String) As String
        
        On Error GoTo GetCtlKeyForLocalFileName_Err
        

        Dim i As Long

100     For i = 1 To Controls.Count - 1

102         If StrComp(Controls(i).LocalFileName, LocalFileName, vbTextCompare) = 0 Then
104             GetCtlKeyForLocalFileName = Controls(i).name
                Exit For
            End If

106     Next i

        
        Exit Function

GetCtlKeyForLocalFileName_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.GetCtlKeyForLocalFileName", _
                  "ucAsyncDLHost component failure"
        
End Function

Public Property Get DownloadStripeByURL(ByVal URL As String) As ucAsyncDLStripe
        
        On Error GoTo DownloadStripeByURL_Err
        
100     Set DownloadStripeByURL = Controls(GetCtlKeyForURL(URL))
        
        Exit Property

DownloadStripeByURL_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.DownloadStripeByURL", _
                  "ucAsyncDLHost component failure"
        
End Property

Public Property Get DownloadStripeByLocalFileName(ByVal LocalFileName As String) As ucAsyncDLStripe
        
        On Error GoTo DownloadStripeByLocalFileName_Err
        
100     Set DownloadStripeByLocalFileName = Controls(GetCtlKeyForLocalFileName(LocalFileName))
        
        Exit Property

DownloadStripeByLocalFileName_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.DownloadStripeByLocalFileName", _
                  "ucAsyncDLHost component failure"
        
End Property

Public Property Get DownloadStripeByIndex(ByVal IdxOneBased As Long) As ucAsyncDLStripe
        
        On Error GoTo DownloadStripeByIndex_Err
        
100     Set DownloadStripeByIndex = Controls(IdxOneBased)
        
        Exit Property

DownloadStripeByIndex_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.DownloadStripeByIndex", _
                  "ucAsyncDLHost component failure"
        
End Property

Private Sub UserControl_Resize()
    
    On Error Resume Next
    
    VScroll.Move ScaleWidth - VScroll.Width, 0, VScroll.Width, ScaleHeight
End Sub

Private Sub VScroll_GotFocus()
        
        On Error GoTo VScroll_GotFocus_Err
        
100     UserControl.SetFocus
        
        Exit Sub

VScroll_GotFocus_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.VScroll_GotFocus", _
                  "ucAsyncDLHost component failure"
        
End Sub

Private Sub VScroll_Scroll()
        
        On Error GoTo VScroll_Scroll_Err
        
100     VScroll_Change
        
        Exit Sub

VScroll_Scroll_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.VScroll_Scroll", _
                  "ucAsyncDLHost component failure"
        
End Sub

Private Sub VScroll_Change()
        
        On Error GoTo VScroll_Change_Err
        

        Dim i&, Ctl As Control, CurTop&

100     VScroll.Enabled = (Controls.Count > 1)
102     VScroll.Min = 1
104     VScroll.Max = Controls.Count - 1
  
        'switch all "Stripes" invisible first
106     For Each Ctl In Controls

108         If Not TypeOf Ctl Is VScrollBar Then Ctl.Visible = False
110     Next Ctl
  
        'move and ensure visibility here
112     For i = VScroll.value To Controls.Count - 1

114         If CurTop > ScaleHeight Then Exit For
116         Set Ctl = Controls(i)
    
118         Ctl.Top = CurTop
120         Ctl.Visible = True
    
122         CurTop = CurTop + Ctl.Height
124     Next i

        
        Exit Sub

VScroll_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.VScroll_Change", _
                  "ucAsyncDLHost component failure"
        
End Sub

'only a small helper-function, to determine the first part of a ProgId (used for Controls.Add)
Private Function GetProjectName() As String

    On Error Resume Next

    Err.Raise 5
    GetProjectName = Err.Source

    On Error GoTo 0

End Function

're-delegate the Events from the Client-Control-Stripes
Public Sub RaiseDownloadProgress(Sender As ucAsyncDLStripe, ByVal BytesRead As Long, ByVal BytesTotal As Long)
        
        On Error GoTo RaiseDownloadProgress_Err
        
100     RaiseEvent DownloadProgress(Sender, BytesRead, BytesTotal)
        
        Exit Sub

RaiseDownloadProgress_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.RaiseDownloadProgress", _
                  "ucAsyncDLHost component failure"
        
End Sub

Public Sub RaiseDownloadError(Sender As ucAsyncDLStripe, ByVal StatusCode As Long, ByVal Status As String)
        
        On Error GoTo RaiseDownloadError_Err
        
100     RaiseEvent DownloadError(Sender, StatusCode, Status)
        
        Exit Sub

RaiseDownloadError_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.RaiseDownloadError", _
                  "ucAsyncDLHost component failure"
        
End Sub

Public Sub RaiseDownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
        
        On Error GoTo RaiseDownloadComplete_Err
        
100     RaiseEvent DownloadComplete(Sender, TmpFileName)
        
        Exit Sub

RaiseDownloadComplete_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.ucAsyncDLHost.RaiseDownloadComplete", _
                  "ucAsyncDLHost component failure"
        
End Sub
 
