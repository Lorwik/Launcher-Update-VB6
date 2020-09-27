Attribute VB_Name = "ModUpdate"
Option Explicit

Public Inet As clsInet

Private Const URLUPDATE As String = "http://winterao.com.ar/update/"
Private Const VERSIONINFOJSON As String = "VersionInfo.json"
Private Const VERSIONINFOINI As String = "VersionInfo.ini"
Public Const LAUNCHEREXEUP As String = "WinterAOLauncher.exe.up"

Type tArchivos
    md5 As String
    Archivo As String
End Type

Type tUpdate
    TotalFiles As Integer
    TotalCarpetas As Integer
    Archivos() As tArchivos
    Carpetas() As String
    JsonListas As Object
    LauncherCheck As String
End Type

'************************
Public UpdateLocal As tUpdate
Public updateREMOTE As tUpdate
'************************

Public SinVersiones As Boolean 'Si no encontro el VersionInfo
Public Desactualizados As Integer 'Numero de archivos desactualizados
Public DesactualizadosList() As tArchivos

'Indica si hay actualizaciones pendientes
Public ActualizacionesPendientes As Boolean
Public LauncherDesactualizado As Boolean

Public Fallaron As String

Public Function LocalFile() As String
    LocalFile = App.Path & "\Init\" & VERSIONINFOINI
End Function

Public Sub CargarListasLOCAL()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Carga la lista de versiones LOCAL
'********************************************

    Dim Archivo     As String
    Dim i           As Integer

    '¿Existe el archivo de versiones en el directorio local?
    If Not FileExist(LocalFile, vbArchive) Then
        SinVersiones = True
        Exit Sub
    End If

    UpdateLocal.TotalFiles = Val(GetVar(LocalFile, "MANIFEST", "TOTALFILES"))
    UpdateLocal.TotalCarpetas = Val(GetVar(LocalFile, "MANIFEST", "TOTALFOLDERS"))
    UpdateLocal.LauncherCheck = GetVar(LocalFile, "MANIFEST", "checksum")

    ReDim UpdateLocal.Archivos(0 To UpdateLocal.TotalFiles) As tArchivos

    For i = 0 To UpdateLocal.TotalFiles

        UpdateLocal.Archivos(i).Archivo = GetVar(LocalFile, "File" & i, "name")
        UpdateLocal.Archivos(i).md5 = UCase(GetVar(LocalFile, "File" & i, "checksum"))

    Next i

    ReDim UpdateLocal.Carpetas(0 To UpdateLocal.TotalCarpetas) As String

    For i = 0 To UpdateLocal.TotalCarpetas

        UpdateLocal.Carpetas(i) = GetVar(LocalFile, "Folder" & i, "name")

    Next i

End Sub

Public Sub CargarListasREMOTE()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Carga la lista de versiones REMOTA
'********************************************

    Dim responseServer  As String
    Dim Archivo         As String
    Dim i               As Integer

    Set Inet = New clsInet

    responseServer = Inet.OpenRequest(URLUPDATE & VERSIONINFOJSON, "GET")
    responseServer = Inet.Execute
    responseServer = Inet.GetResponseAsString

    Set updateREMOTE.JsonListas = ModJson.parse(responseServer)

    updateREMOTE.TotalFiles = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALFILES"))
    updateREMOTE.TotalCarpetas = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TotalFolders"))
    updateREMOTE.LauncherCheck = updateREMOTE.JsonListas.Item("MANIFEST").Item("checksum")

    ReDim updateREMOTE.Archivos(0 To updateREMOTE.TotalFiles) As tArchivos

    For i = 0 To updateREMOTE.TotalFiles

        updateREMOTE.Archivos(i).Archivo = updateREMOTE.JsonListas.Item("Files").Item("File" & i).Item("name")
        updateREMOTE.Archivos(i).md5 = UCase(updateREMOTE.JsonListas.Item("Files").Item("File" & i).Item("checksum"))

    Next i

    ReDim updateREMOTE.Carpetas(0 To updateREMOTE.TotalCarpetas) As String

    For i = 0 To updateREMOTE.TotalCarpetas

        updateREMOTE.Carpetas(i) = updateREMOTE.JsonListas.Item("Folders").Item("Folder" & i).Item("name")


    Next i

    Set Inet = Nothing

End Sub

Public Function CompararArchivos() As Boolean
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Comprueba si el CHECK de los archivos coinciden con el del remoto
'********************************************

    Dim i               As Integer
    Dim flag            As Boolean
    Dim Archivo         As String

    'El total de archivos remoto es diferente al de local? 'Hay que actualizar seguro.
    If updateREMOTE.TotalFiles <> UpdateLocal.TotalFiles Then
        CompararArchivos = True
        Exit Function
    End If

    For i = 0 To updateREMOTE.TotalFiles

        'Comprobamos todos los CHECK
        If updateREMOTE.Archivos(i).md5 <> UpdateLocal.Archivos(i).md5 Or FileExist(App.Path & "\" & updateREMOTE.Archivos(i).Archivo, vbNormal) = False Then

            Call NuevoDesactualizado(updateREMOTE.Archivos(i).Archivo, updateREMOTE.Archivos(i).md5)

            flag = True 'Activamos el flag
        End If

    Next i

    '¿Hubo algun archivo desactualizado?
    If flag Then
        CompararArchivos = True
    Else
        CompararArchivos = False
    End If

End Function

Public Sub CompararyCrearCarpetas()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Crea las carpetas que no existan
'********************************************
    On Error Resume Next

    Dim i As Integer

    For i = 0 To updateREMOTE.TotalCarpetas
    
        If Not FileExist(App.Path & "\" & updateREMOTE.Carpetas(i), vbDirectory) Then

            MkDir App.Path & "\" & updateREMOTE.Carpetas(i)
            DoEvents

        End If

    Next i

End Sub

Public Function ActualizarCliente() As Boolean
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Descarga los archivos guardados en la lista de desactualizados
'********************************************

    Dim i As Integer
    Dim Archivo As String
    Dim archivoURL As String

    If LauncherDesactualizado Then
        If FileExist(App.Path & "\" & LAUNCHEREXEUP, vbNormal) Then Kill App.Path & "\" & LAUNCHEREXEUP
        frmMain.ucAsyncDLHost.AddDownloadJob URLUPDATE & "cliente/" & LAUNCHEREXEUP, LAUNCHEREXEUP

        DoEvents
    End If
    
    If Desactualizados > 0 Then
        For i = 1 To Desactualizados
    
            'Lo adaptamos a URL
            archivoURL = Replace$(DesactualizadosList(i).Archivo, "\\", "/")
            
            frmMain.ucAsyncDLHost.AddDownloadJob URLUPDATE & "cliente/" & archivoURL, DesactualizadosList(i).Archivo
    
            DoEvents
    
        Next i

    End If
    
    If SinVersiones Then Call ObtenerVersionFile
    
    ActualizacionesPendientes = False

End Function

Private Sub ObtenerVersionFile()
'********************************************
'Autor: Lorwik
'Fecha: 25/09/2020
'Descripción: Descarga el archivo de VersionInfo y lo transforma al ini del cliente
'********************************************

    On Error Resume Next
    
    Dim i As Integer
    
    With updateREMOTE

        Call WriteVar(LocalFile, "MANIFEST", "TOTALFILES", .TotalFiles)
        Call WriteVar(LocalFile, "MANIFEST", "TOTALCARPETAS", .TotalCarpetas)
        Call WriteVar(LocalFile, "MANIFEST", "checksum", .LauncherCheck)
        
        For i = 0 To .TotalFiles
        
            Call WriteVar(LocalFile, "File" & i, "name", .Archivos(i).Archivo)
            Call WriteVar(LocalFile, "File" & i, "checksum", .Archivos(i).md5)
        
        Next i
        
        For i = 0 To .TotalCarpetas
        
            Call WriteVar(LocalFile, "Folder" & i, "name", .Carpetas(i))
        
        Next i
        
        'Ya tenemos el archivo en local, ahora lo cargamos
        Call CargarListasLOCAL
    
    End With
    
End Sub

Public Sub NuevoDesactualizado(ByVal File As String, ByVal Checksum As String)
'********************************************
'Autor: Lorwik
'Fecha: 26/09/2020
'Descripción: Añade un elemento a la lista de desactualizados
'********************************************

    Dim i As Integer
    
    'Si ya existe el archivo en la lista de desactualizados, no lo agregamos
    If Desactualizados > 0 Then
        For i = 1 To Desactualizados
            If DesactualizadosList(i).Archivo = File Then Exit Sub
        Next i
    End If

    ReDim Preserve DesactualizadosList(Desactualizados + 1) As tArchivos

    'Aumentamos el contador de la cantidad de archivos para actualizar
    Desactualizados = Desactualizados + 1

    'Añadimos el archivo a la lista para actualizar mas tarde
    DesactualizadosList(Desactualizados).Archivo = File
    DesactualizadosList(Desactualizados).md5 = Checksum
            
End Sub
