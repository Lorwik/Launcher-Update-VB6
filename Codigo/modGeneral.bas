Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Public Directory As String
Public bDone As Boolean
Public dError As Boolean

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
        
        'Listamos todos los archivos excepto la posicion 0 que corresponde al propio Launcher.
        For i = 1 To updateREMOTE.TotalFiles
        
            Call NuevoDesactualizado(updateREMOTE.Archivos(i).Archivo, updateREMOTE.Archivos(i).md5)
        
        Next i
        
        frmMain.lblPendientes.Caption = "¡No se ha encontrado el cliente! Pulsa Jugar para descargar los archivos del cliente."
        Call LauncherLog("¡No se ha encontrado el cliente!")
        
    Else
    
        '¿Hay actualizaciones pendientes?
        ActualizacionesPendientes = modUpdate.CompararArchivos
        
        'Notificamos en el Main que hay actualizaciones pendientes
        If ActualizacionesPendientes Then
            frmMain.lblPendientes.Caption = "Hay " & Desactualizados & " archivos desactualizados."
            Call LauncherLog("Hay " & Desactualizados & " archivos desactualizados.")
            
        Else
            frmMain.lblPendientes.Caption = "Cliente actualizado. Pulsa Jugar para abrir el cliente."
            
        End If
    End If

    DoEvents
    
    frmMain.Show

End Sub

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function
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

Public Sub addConsole(Texto As String, Rojo As Byte, Verde As Byte, Azul As Byte, Bold As Boolean, Italic As Boolean, Optional ByVal Enter As Boolean = False)
    With frmMain.RichTextBox1
        If (Len(.Text)) > 700 Then .Text = ""
        
        .SelStart = Len(.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        .SelColor = RGB(Rojo, Verde, Azul)
        
        .SelText = IIf(Enter, Texto, Texto & vbCrLf)
        
        .Refresh
    End With
frmMain.Caption = "Aut" & "oup" & "date" & " Winter" & "AO, v" & App.Major & "." & App.Minor
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
    If Not FileExist(App.Path & "\Logs", vbDirectory) Then _
        MkDir App.Path & "\Logs"
    
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

    Dim i As Integer
    
    If SinVersiones Then Exit Sub
    
    For i = 0 To updateREMOTE.TotalFiles
    
        '¿Encontro el archivo?
        If updateREMOTE.Archivos(i).Archivo = Archivo Then
        
            'Actualizamos el archivo de versiones
            Call WriteVar(LocalFile, "File" & i, "name", Archivo)
            Call WriteVar(LocalFile, "File" & i, "checksum", Check)
            
            UpdateLocal.Archivos(i).Archivo = updateREMOTE.Archivos(i).Archivo
            UpdateLocal.Archivos(i).md5 = updateREMOTE.Archivos(i).md5
            
        End If
    
    Next i

End Sub

