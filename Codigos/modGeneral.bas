Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000

Public Sub Main()
    Dim i As Integer

    'Cargamos listas de archivos y carpetas
    Call CargarListasLOCAL
    Call CargarListasREMOTE
    
    'Comprobamos y creamos carpetas
    Call CompararyCrearCarpetas
    
    '¿No existe el archivo de versiones?
    If SinVersiones Then
        ActualizacionesPendientes = True
        
        For i = 1 To updateREMOTE.TotalFiles
        
            Call NuevoDesactualizado(updateREMOTE.Archivos(i).Archivo, updateREMOTE.Archivos(i).md5)
        
        Next i
        
        frmMain.lblPendientes.Caption = "¡No se ha encontrado el cliente! Pulsa Jugar para descargar los archivos del cliente."
        Call LauncherLog("¡No se ha encontrado el cliente!")
        
    Else
    
        '¿Hay actualizaciones pendientes?
        ActualizacionesPendientes = ModUpdate.CompararArchivos
        
        '¿Hay actualizaciones para el Launcher?
        If UpdateLocal.LauncherCheck <> updateREMOTE.LauncherCheck Then _
            LauncherDesactualizado = True
        
        'Notificamos en el Main que hay actualizaciones pendientes
        If ActualizacionesPendientes Then
            frmMain.lblPendientes.Caption = "Hay " & Desactualizados & " archivos desactualizados."
            Call LauncherLog("Hay " & Desactualizados & " archivos desactualizados.")
            
        ElseIf LauncherDesactualizado Then
            frmMain.lblPendientes.Caption = "Hay una actualizacion disponible para el Launcher."
            Call LauncherLog("Hay una actualizacion disponible para el Launcher.")
            
        Else
            frmMain.lblPendientes.Caption = "Cliente actualizado. Pulsa Jugar para abrir el cliente."
            
        End If
    End If
    
    frmMain.lblVersion.Caption = updateREMOTE.updateNumber
    
    DoEvents
    
    frmMain.Show
    
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = Mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Public Function FileToString(strFileName As String) As String
    '###################################################################################
    ' Convierte un archivo entero a una cadena de texto para almacenarla en una variable
    '###################################################################################
    Dim IFile As Variant
    
    IFile = FreeFile
    Open strFileName For Input As #IFile
        FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
End Function

Public Sub Skin(Frm As Form, Color As Long)
    Frm.BackColor = Color
    Dim Ret As Long
    Ret = GetWindowLong(Frm.hwnd, G_E)
    Ret = Ret Or W_E
    SetWindowLong Frm.hwnd, G_E, Ret
    SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY
End Sub

Public Sub LauncherLog(Desc As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 25/09/2020
    '***************************************************

    On Error GoTo errhandler

    Dim nfile As Integer
        nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\launcher.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile
    
    Exit Sub
    
    Debug.Print Desc
    
errhandler:

End Sub

Public Sub ActualizarVersionInfo(ByVal Archivo As String, ByVal Check As String)
'********************************************
'Autor: Lorwik
'Fecha: 25/09/2020
'Descripción: Actualiza el archivo de versiones local
'********************************************

    Dim i As Integer
    
    For i = 1 To updateREMOTE.TotalFiles
    
        '¿Encontro el archivo?
        If updateREMOTE.Archivos(i).Archivo = Archivo Then
        
            'Actualizamos el archivo de versiones
            Call WriteVar(LocalFile, "A" & i, "ARCHIVO", Archivo)
            Call WriteVar(LocalFile, "A" & i, "CHECK", Check)
            
            UpdateLocal.Archivos(i).Archivo = updateREMOTE.Archivos(i).Archivo
            UpdateLocal.Archivos(i).md5 = updateREMOTE.Archivos(i).md5
            
        End If
    
    Next i

End Sub

