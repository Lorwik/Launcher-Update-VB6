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
Dim NewStripe As ucAsyncDLStripe
Static CC As Currency: CC = CC + 1
  
  If Len(GetCtlKeyForLocalFileName(LocalFileName)) Then Err.Raise vbObjectError, , "Ya hay una descarga con ese nombre de archivo local en la lista"
  If Len(GetCtlKeyForURL(URL)) Then Err.Raise vbObjectError, , "Ya hay una Descarga con esa URL en la Lista"
 
  With Controls.Add(GetProjectName & ".ucAsyncDLStripe", "K" & CC)
    .Move 0, Controls.Count * .Height, ScaleWidth - VScroll.Width
    .Visible = True
    
    VScroll_Change
 
    Set NewStripe = .Object 'just a cast to the concrete Control-Interface
        NewStripe.DownloadFile URL, LocalFileName 'and here we trigger the Download
  End With
End Sub

Public Sub RemoveByLocalFileName(ByVal LocalFileName As String)
Dim CtlKey As String
  CtlKey = GetCtlKeyForLocalFileName(LocalFileName)
  If Len(CtlKey) = 0 Then Exit Sub
  Controls.Remove CtlKey
  VScroll_Change
End Sub

Public Function GetCtlKeyForURL(ByVal URL As String) As String
Dim i As Long
  For i = 1 To Controls.Count - 1
    If StrComp(Controls(i).URL, URL, vbTextCompare) = 0 Then
      GetCtlKeyForURL = Controls(i).name: Exit For
    End If
  Next i
End Function

Public Function GetCtlKeyForLocalFileName(ByVal LocalFileName As String) As String
Dim i As Long
  For i = 1 To Controls.Count - 1
    If StrComp(Controls(i).LocalFileName, LocalFileName, vbTextCompare) = 0 Then
      GetCtlKeyForLocalFileName = Controls(i).name: Exit For
    End If
  Next i
End Function

Public Property Get DownloadStripeByURL(ByVal URL As String) As ucAsyncDLStripe
  Set DownloadStripeByURL = Controls(GetCtlKeyForURL(URL))
End Property
Public Property Get DownloadStripeByLocalFileName(ByVal LocalFileName As String) As ucAsyncDLStripe
  Set DownloadStripeByLocalFileName = Controls(GetCtlKeyForLocalFileName(LocalFileName))
End Property
Public Property Get DownloadStripeByIndex(ByVal IdxOneBased As Long) As ucAsyncDLStripe
  Set DownloadStripeByIndex = Controls(IdxOneBased)
End Property

Private Sub UserControl_Resize()
  VScroll.Move ScaleWidth - VScroll.Width, 0, VScroll.Width, ScaleHeight
End Sub

Private Sub VScroll_GotFocus()
  UserControl.SetFocus
End Sub

Private Sub VScroll_Scroll()
  VScroll_Change
End Sub

Private Sub VScroll_Change()
Dim i&, Ctl As Control, CurTop&
  VScroll.Enabled = (Controls.Count > 1)
  VScroll.Min = 1
  VScroll.Max = Controls.Count - 1
  
  'switch all "Stripes" invisible first
  For Each Ctl In Controls
    If Not TypeOf Ctl Is VScrollBar Then Ctl.Visible = False
  Next Ctl
  
  'move and ensure visibility here
  For i = VScroll.value To Controls.Count - 1
    If CurTop > ScaleHeight Then Exit For
    Set Ctl = Controls(i)
    
    Ctl.Top = CurTop
    Ctl.Visible = True
    
    CurTop = CurTop + Ctl.Height
  Next i
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
  RaiseEvent DownloadProgress(Sender, BytesRead, BytesTotal)
End Sub
Public Sub RaiseDownloadError(Sender As ucAsyncDLStripe, ByVal StatusCode As Long, ByVal Status As String)
  RaiseEvent DownloadError(Sender, StatusCode, Status)
End Sub
Public Sub RaiseDownloadComplete(Sender As ucAsyncDLStripe, ByVal TmpFileName As String)
  RaiseEvent DownloadComplete(Sender, TmpFileName)
End Sub
 
