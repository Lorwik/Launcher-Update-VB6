Attribute VB_Name = "ModUpdate"
Option Explicit

Public Inet As clsInet

Private Const URLUpdate As String = "http://winterao.com.ar/update/"
Private Const VersionInfoJSON As String = "VersionInfo.json"
Private Const VersionInfoINI As String = "VersionInfo.ini"

Type tArchivos
    md5 As String
    Archivo As String
End Type

Type tUpdate
    updateNumber As Integer
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
    LocalFile = App.Path & "\Init\" & VersionInfoINI
End Function

Public Sub CargarListasLOCAL()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripci�n: Carga la lista de versiones LOCAL
'********************************************

    Dim Archivo     As String
    Dim i           As Integer

    '�Existe el archivo de versiones en el directorio local?
    If Not FileExist(LocalFile, vbArchive) Then
        SinVersiones = True
        Exit Sub
    End If

    UpdateLocal.updateNumber = Val(GetVar(LocalFile, "MANIFEST", "UPDATENUMBER"))
    UpdateLocal.TotalFiles = Val(GetVar(LocalFile, "MANIFEST", "TOTALFILES"))
    UpdateLocal.TotalCarpetas = Val(GetVar(LocalFile, "MANIFEST", "TOTALCARPETAS"))
    UpdateLocal.LauncherCheck = GetVar(LocalFile, "MANIFEST", "CHECK")

    ReDim UpdateLocal.Archivos(1 To UpdateLocal.TotalFiles) As tArchivos

    For i = 1 To UpdateLocal.TotalFiles

        UpdateLocal.Archivos(i).Archivo = GetVar(LocalFile, "A" & i, "ARCHIVO")
        UpdateLocal.Archivos(i).md5 = UCase(GetVar(LocalFile, "A" & i, "CHECK"))

    Next i

    ReDim UpdateLocal.Carpetas(1 To UpdateLocal.TotalCarpetas) As String

    For i = 1 To UpdateLocal.TotalCarpetas

        UpdateLocal.Carpetas(i) = GetVar(LocalFile, "C" & i, "CARPETA")

    Next i

End Sub

Public Sub CargarListasREMOTE()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripci�n: Carga la lista de versiones REMOTA
'********************************************

    Dim responseServer  As String
    Dim Archivo         As String
    Dim i               As Integer

    Set Inet = New clsInet

    responseServer = Inet.OpenRequest(URLUpdate & VersionInfoJSON, "GET")
    responseServer = Inet.Execute
    responseServer = Inet.GetResponseAsString

    Set updateREMOTE.JsonListas = ModJson.parse(responseServer)

    updateREMOTE.updateNumber = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("UPDATENUMBER"))
    updateREMOTE.TotalFiles = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALFILES"))
    updateREMOTE.TotalCarpetas = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALCARPETAS"))
    updateREMOTE.LauncherCheck = updateREMOTE.JsonListas.Item("MANIFEST").Item("CHECK")

    ReDim updateREMOTE.Archivos(1 To updateREMOTE.TotalFiles) As tArchivos

    For i = 1 To updateREMOTE.TotalFiles

        updateREMOTE.Archivos(i).Archivo = updateREMOTE.JsonListas.Item("A" & i).Item("ARCHIVO")
        updateREMOTE.Archivos(i).md5 = UCase(updateREMOTE.JsonListas.Item("A" & i).Item("CHECK"))

    Next i

    ReDim updateREMOTE.Carpetas(1 To updateREMOTE.TotalCarpetas) As String

    For i = 1 To updateREMOTE.TotalCarpetas

        updateREMOTE.Carpetas(i) = updateREMOTE.JsonListas.Item("C" & i).Item("CARPETA")

    Next i

    Set Inet = Nothing

End Sub

Public Function CompararVersiones() As Boolean
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripci�n: Comprueba si hay actualizaciones
'********************************************

    If UpdateLocal.updateNumber <> updateREMOTE.updateNumber Then
        CompararVersiones = False
        Exit Function
    End If

    CompararVersiones = True

End Function

Public Function CompararArchivos() As Boolean
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripci�n: Comprueba si el CHECK de los archivos coinciden con el del remoto
'********************************************

    Dim i               As Integer
    Dim flag            As Boolean
    Dim Archivo         As String

    'El total de archivos remoto es diferente al de local? 'Hay que actualizar seguro.
    If updateREMOTE.TotalFiles <> UpdateLocal.TotalFiles Then
        CompararArchivos = True
        Exit Function
    End If

    For i = 1 To updateREMOTE.TotalFiles

        Archivo = Replace$(updateREMOTE.Archivos(i).Archivo, "-", "\")
        'Comprobamos todos los CHECK
        If updateREMOTE.Archivos(i).md5 <> UpdateLocal.Archivos(i).md5 Or FileExist(App.Path & "\" & Archivo, vbNormal) = False Then

            ReDim Preserve DesactualizadosList(Desactualizados + 1) As tArchivos

            'Aumentamos el contador de la cantidad de archivos para actualizar
            Desactualizados = Desactualizados + 1

            'A�adimos el archivo a la lista para actualizar mas tarde
            DesactualizadosList(Desactualizados).Archivo = updateREMOTE.Archivos(i).Archivo
            DesactualizadosList(Desactualizados).md5 = updateREMOTE.Archivos(i).md5

            flag = True 'Activamos el flag
        End If

    Next i

    '�Hubo algun archivo desactualizado?
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
'Descripci�n: Crea las carpetas que no existan
'********************************************
    On Error Resume Next

    Dim i As Integer
    Dim Directorio As String

    For i = 1 To updateREMOTE.TotalCarpetas

        Directorio = Replace(updateREMOTE.Carpetas(i), "-", "\")

        If Not FileExist(App.Path & Directorio, vbDirectory) Then

            MkDir Directorio
            DoEvents

        End If

    Next i

End Sub

Public Function ActualizarCliente() As Boolean
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripci�n: Descarga los archivos guardados en la lista de desactualizados
'********************************************

    Dim i As Integer
    Dim Archivo As String
    Dim archivoURL As String

    If LauncherDesactualizado Then
        If FileExist(App.Path & "\WinterAOLauncher.exe.up", vbNormal) Then Kill App.Path & "\WinterAOLauncher.exe.up"
        frmMain.ucAsyncDLHost.AddDownloadJob URLUpdate & "launcher/WinterAOLauncher.exe.up", "WinterAOLauncher.exe.up"

        DoEvents
        LauncherDesactualizado = False
    End If
    
    If Desactualizados > 0 Then
        For i = 1 To Desactualizados
    
            'Primero lo adaptamos a URL
            archivoURL = Replace$(DesactualizadosList(i).Archivo, "-", "/")
    
            'Luego a directorio de Windows
            Archivo = Replace$(DesactualizadosList(i).Archivo, "-", "\")
            
            frmMain.ucAsyncDLHost.AddDownloadJob URLUpdate & "cliente/" & archivoURL, Archivo
    
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
'Descripci�n: Descarga el archivo de VersionInfo y lo transforma al ini del cliente
'********************************************

    On Error Resume Next
    
    Dim i As Integer
    
    With updateREMOTE

        Call WriteVar(LocalFile, "MANIFEST", "UPDATENUMBER", .updateNumber)
        Call WriteVar(LocalFile, "MANIFEST", "TOTALFILES", .TotalFiles)
        Call WriteVar(LocalFile, "MANIFEST", "TOTALCARPETAS", .TotalCarpetas)
        Call WriteVar(LocalFile, "MANIFEST", "CHECK", .LauncherCheck)
        
        For i = 1 To .TotalFiles
        
            Call WriteVar(LocalFile, "A" & i, "ARCHIVO", .Archivos(i).Archivo)
            Call WriteVar(LocalFile, "A" & i, "CHECK", .Archivos(i).md5)
        
        Next i
        
        For i = 1 To .TotalCarpetas
        
            Call WriteVar(LocalFile, "C" & i, "ARCHIVO", .Carpetas(i))
        
        Next i
    
    End With
    
End Sub
