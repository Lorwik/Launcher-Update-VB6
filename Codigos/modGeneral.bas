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
        
        On Error GoTo Main_Err
        

        Dim i As Integer

        'Cargamos listas de archivos y carpetas
100     Call CargarListasLOCAL
102     Call CargarListasREMOTE
    
        'Comprobamos y creamos carpetas
104     Call CompararyCrearCarpetas
    
        '¿No existe el archivo de versiones?
106     If SinVersiones Then
108         ActualizacionesPendientes = True
        
110         For i = 0 To updateREMOTE.TotalFiles
112             Call NuevoDesactualizado(updateREMOTE.Archivos(i).Archivo, updateREMOTE.Archivos(i).md5)
114         Next i
        
116         frmMain.lblPendientes.Caption = "¡No se ha encontrado el cliente! Pulsa Jugar para descargar los archivos del cliente."

118         Call LauncherLog("¡No se ha encontrado el cliente!")
        
        Else
    
            '¿Hay actualizaciones pendientes?
120         ActualizacionesPendientes = ModUpdate.CompararArchivos
        
            '¿Hay actualizaciones para el Launcher?
122         If UpdateLocal.LauncherCheck <> updateREMOTE.LauncherCheck Then LauncherDesactualizado = True
        
            'Notificamos en el Main que hay actualizaciones pendientes
124         If ActualizacionesPendientes Then
126             frmMain.lblPendientes.Caption = "Hay " & Desactualizados & " archivos desactualizados."
128             Call LauncherLog("Hay " & Desactualizados & " archivos desactualizados.")
            
130         ElseIf LauncherDesactualizado Then
132             frmMain.lblPendientes.Caption = "Hay una actualizacion disponible para el Launcher."
134             Call LauncherLog("Hay una actualizacion disponible para el Launcher.")
            
            Else
136             frmMain.lblPendientes.Caption = "Cliente actualizado. Pulsa Jugar para abrir el cliente."
            
            End If
            
        End If

138     DoEvents
    
140     frmMain.Show
    
        
        Exit Sub

Main_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modGeneral.Main " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
        
    On Error GoTo WriteVar_Err
        
    Call writeprivateprofilestring(Main, Var, value, File)
        
    Exit Sub

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
        
    Exit Function

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
        
    FileExist = (Dir$(File, FileType) <> vbNullString)
    
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
        
        On Error GoTo ReadField_Err
        

        '*****************************************************************
        'Gets a field from a delimited string
        'Author: Juan Martin Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        '*****************************************************************
        Dim i          As Long
        Dim lastPos    As Long
        Dim CurrentPos As Long
        Dim delimiter  As String * 1
    
100     delimiter = Chr$(SepASCII)
    
102     For i = 1 To Pos
104         lastPos = CurrentPos
106         CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
108     Next i
    
110     If CurrentPos = 0 Then
112         ReadField = Mid$(Text, lastPos + 1, Len(Text) - lastPos)
        Else
114         ReadField = Mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
        End If

        
        Exit Function

ReadField_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modGeneral.ReadField " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Function FileToString(strFileName As String) As String
        
        On Error GoTo FileToString_Err
        

        '###################################################################################
        ' Convierte un archivo entero a una cadena de texto para almacenarla en una variable
        '###################################################################################
        Dim IFile As Variant
    
100     IFile = FreeFile
102     Open strFileName For Input As #IFile
104     FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
106     Close #IFile
        
        Exit Function

FileToString_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modGeneral.FileToString " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Sub Skin(Frm As Form, Color As Long)
        
        On Error GoTo Skin_Err
        
100     Frm.BackColor = Color

        Dim Ret As Long

102     Ret = GetWindowLong(Frm.hwnd, G_E)
104     Ret = Ret Or W_E
106     SetWindowLong Frm.hwnd, G_E, Ret
108     SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY
        
        Exit Sub

Skin_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modGeneral.Skin " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Public Sub LauncherLog(Desc As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 25/09/2020
    '***************************************************

    On Error GoTo errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
        
    '¿No existe la carpeta logs?
    If Not FileExist(App.Path & "\Logs", vbDirectory) Then MkDir App.Path & "\Logs"
    
    Open App.Path & "\Logs\launcher.log" For Append Shared As #nfile
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
        
        On Error GoTo ActualizarVersionInfo_Err
        

        Dim i As Integer
    
100     If SinVersiones Then Exit Sub
    
102     For i = 0 To updateREMOTE.TotalFiles
    
            '¿Encontro el archivo?
104         If updateREMOTE.Archivos(i).Archivo = Archivo Then
        
                'Actualizamos el archivo de versiones
106             Call WriteVar(LocalFile, "File" & i, "name", Archivo)
108             Call WriteVar(LocalFile, "File" & i, "checksum", Check)
            
110             UpdateLocal.Archivos(i).Archivo = updateREMOTE.Archivos(i).Archivo
112             UpdateLocal.Archivos(i).md5 = updateREMOTE.Archivos(i).md5
            
            End If
    
114     Next i

        
        Exit Sub

ActualizarVersionInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modGeneral.ActualizarVersionInfo " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

