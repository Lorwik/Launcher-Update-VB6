Attribute VB_Name = "ModUpdate"
Option Explicit

Public Inet As clsInet

Private Const URLUpdate As String = "http://winterao.com.ar/update/"
Private Const VersionInfoFile As String = "VersionInfo.json"

Type tArchivos
    md5 As String
    archivo As String
End Type

Type tUpdate
    updateNumber As Integer
    TotalFiles As Integer
    TotalCarpetas As Integer
    Archivos() As tArchivos
    Carpetas() As String
    JsonListas As Object
End Type

'************************
'LOCAL
'************************
Public UpdateLocal As tUpdate
'************************

'************************
'REMOTO
'************************
Public updateREMOTE As tUpdate
'************************

Public SinVersiones As Boolean
Public Desactualizados As Integer
Public ListaActualizar() As String

'Indica si hay actualizaciones pendientes
Public ActualizacionesPendientes As Boolean

Public Fallaron As String

Public Sub CargarListasLOCAL()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Carga la lista de versiones LOCAL
'********************************************

    Dim archivo     As String
    Dim FileVer     As String
    Dim LocalFile   As String
    Dim i           As Integer

    LocalFile = App.Path & "\Init\" & VersionInfoFile

    '¿Existe el archivo de versiones en el directorio local?
    If Not FileExist(LocalFile, vbArchive) Then
        SinVersiones = True
        Exit Sub
    End If

    FileVer = FileToString(LocalFile)

    Set UpdateLocal.JsonListas = ModJson.parse(FileVer)

    UpdateLocal.updateNumber = Val(UpdateLocal.JsonListas.Item("MANIFEST").Item("UPDATENUMBER"))
    UpdateLocal.TotalFiles = Val(UpdateLocal.JsonListas.Item("MANIFEST").Item("TOTALFILES"))
    UpdateLocal.TotalCarpetas = Val(UpdateLocal.JsonListas.Item("MANIFEST").Item("TOTALCARPETAS"))

    ReDim UpdateLocal.Archivos(1 To UpdateLocal.TotalFiles) As tArchivos

    For i = 1 To UpdateLocal.TotalFiles

        UpdateLocal.Archivos(i).archivo = UpdateLocal.JsonListas.Item("A" & i).Item("ARCHIVO")
        UpdateLocal.Archivos(i).md5 = UCase(UpdateLocal.JsonListas.Item("A" & i).Item("CHECK"))

    Next i

    ReDim UpdateLocal.Carpetas(1 To UpdateLocal.TotalCarpetas) As String

    For i = 1 To UpdateLocal.TotalCarpetas

        UpdateLocal.Carpetas(i) = UpdateLocal.JsonListas.Item("C" & i).Item("CARPETA")

    Next i

End Sub

Public Sub CargarListasREMOTE()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Carga la lista de versiones REMOTA
'********************************************

    Dim responseServer  As String
    Dim archivo         As String
    Dim i               As Integer

    Set Inet = New clsInet

    responseServer = Inet.OpenRequest(URLUpdate & VersionInfoFile, "GET")
    responseServer = Inet.Execute
    responseServer = Inet.GetResponseAsString

    Set updateREMOTE.JsonListas = ModJson.parse(responseServer)

    updateREMOTE.updateNumber = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("UPDATENUMBER"))
    updateREMOTE.TotalFiles = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALFILES"))
    updateREMOTE.TotalCarpetas = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALCARPETAS"))

    ReDim updateREMOTE.Archivos(1 To updateREMOTE.TotalFiles) As tArchivos

    For i = 1 To updateREMOTE.TotalFiles

        updateREMOTE.Archivos(i).archivo = updateREMOTE.JsonListas.Item("A" & i).Item("ARCHIVO")
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
'Descripción: Comprueba si hay actualizaciones
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
'Descripción: Comprueba si el CHECK de los archivos coinciden con el del remoto
'********************************************

    Dim i               As Integer
    Dim flag            As Boolean
    Dim archivo         As String

    'El total de archivos remoto es diferente al de local? 'Hay que actualizar seguro.
    If updateREMOTE.TotalFiles <> UpdateLocal.TotalFiles Then
        CompararArchivos = True
        Exit Function
    End If

    For i = 1 To updateREMOTE.TotalFiles

        archivo = Replace$(updateREMOTE.Archivos(i).archivo, "-", "\")
        'Comprobamos todos los CHECK
        If updateREMOTE.Archivos(i).md5 <> UpdateLocal.Archivos(i).md5 Or FileExist(App.Path & "\" & archivo, vbNormal) = False Then

            ReDim Preserve ListaActualizar(Desactualizados + 1) As String

            'Aumentamos el contador de la cantidad de archivos para actualizar
            Desactualizados = Desactualizados + 1

            'Añadimos el archivo a la lista para actualizar mas tarde
            ListaActualizar(Desactualizados) = updateREMOTE.Archivos(i).archivo

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
'Descripción: Descarga los archivos guardados en la lista de desactualizados
'********************************************

    Dim i As Integer
    Dim archivo As String
    Dim archivoURL As String

    For i = 1 To Desactualizados

        'Primero lo adaptamos a URL
        archivoURL = Replace$(ListaActualizar(i), "-", "/")

        'Luego a directorio de Windows
        archivo = Replace$(ListaActualizar(i), "-", "\")
        
        frmMain.ucAsyncDLHost.AddDownloadJob URLUpdate & "cliente/" & archivoURL, archivo

        DoEvents

    Next i
    
    'Esto se tiene que mejorar
    frmMain.ucAsyncDLHost.AddDownloadJob URLUpdate & "VersionInfo.json", App.Path & "\INIT\VersionInfo.json"
    
    ActualizacionesPendientes = False

End Function
