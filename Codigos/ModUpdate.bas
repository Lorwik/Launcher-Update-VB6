Attribute VB_Name = "ModUpdate"
Option Explicit

Public Inet As clsInet

Private Const URLUpdate As String = "http://winterao.com.ar/update/"
Private Const VersionInfoFile As String = "VersionInfo.json"

Type tArchivos
    md5 As String
    archivo As String
End Type

'************************
'LOCAL
'************************
Private UpdateNumberLOCAL As Integer
Public TotalFilesLOCAL As Integer
Public TotalCarpetasLOCAL As Integer
Public ArchivosLOCAL() As tArchivos
Public CarpetasLOCAL() As String
Private JsonListasLOCAL As Object
'************************

'************************
'REMOTO
'************************
Private UpdateNumberREMOTE As Integer
Public TotalFilesREMOTE As Integer
Public TotalCarpetasREMOTE As Integer
Public ArchivosREMOTE() As tArchivos
Public CarpetasREMOTE() As String
Private JsonListasREMOTE As Object
'************************

Public SinVersiones As Boolean
Public Desactualizados As Integer
Public ListaActualizar() As String

'Indica si hay actualizaciones pendientes
Public ActualizacionesPendientes As Boolean

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

    Set JsonListasLOCAL = ModJson.parse(FileVer)

    UpdateNumberLOCAL = Val(JsonListasLOCAL.Item("MANIFEST").Item("UPDATENUMBER"))
    TotalFilesLOCAL = Val(JsonListasLOCAL.Item("MANIFEST").Item("TOTALFILES"))
    TotalCarpetasLOCAL = Val(JsonListasLOCAL.Item("MANIFEST").Item("TOTALCARPETAS"))

    ReDim ArchivosLOCAL(1 To TotalFilesLOCAL) As tArchivos

    For i = 1 To TotalFilesLOCAL

        ArchivosLOCAL(i).archivo = JsonListasLOCAL.Item("A" & i).Item("ARCHIVO")
        ArchivosLOCAL(i).md5 = JsonListasLOCAL.Item("A" & i).Item("CHECK")

    Next i

    ReDim CarpetasLOCAL(1 To TotalCarpetasLOCAL) As String

    For i = 1 To TotalCarpetasLOCAL

        CarpetasLOCAL(i) = JsonListasLOCAL.Item("C" & i).Item("CARPETA")

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

    Set JsonListasREMOTE = ModJson.parse(responseServer)

    UpdateNumberREMOTE = Val(JsonListasREMOTE.Item("MANIFEST").Item("UPDATENUMBER"))
    TotalFilesREMOTE = Val(JsonListasREMOTE.Item("MANIFEST").Item("TOTALFILES"))
    TotalCarpetasREMOTE = Val(JsonListasREMOTE.Item("MANIFEST").Item("TOTALCARPETAS"))

    ReDim ArchivosREMOTE(1 To TotalFilesREMOTE) As tArchivos

    For i = 1 To TotalFilesREMOTE

        ArchivosREMOTE(i).archivo = JsonListasREMOTE.Item("A" & i).Item("ARCHIVO")
        ArchivosREMOTE(i).md5 = JsonListasREMOTE.Item("A" & i).Item("CHECK")

    Next i

    ReDim CarpetasREMOTE(1 To TotalCarpetasREMOTE) As String

    For i = 1 To TotalCarpetasREMOTE

        CarpetasREMOTE(i) = JsonListasREMOTE.Item("C" & i).Item("CARPETA")

    Next i

    Set Inet = Nothing

End Sub

Public Function CompararVersiones() As Boolean
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Comprueba si hay actualizaciones
'********************************************

    If UpdateNumberLOCAL <> UpdateNumberREMOTE Then
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
    If TotalFilesREMOTE <> TotalFilesLOCAL Then
        CompararArchivos = True
        Exit Function
    End If

    For i = 1 To TotalFilesREMOTE

        archivo = Replace$(ArchivosREMOTE(i).archivo, "-", "\")
        'Comprobamos todos los CHECK
        If ArchivosREMOTE(i).md5 <> ArchivosLOCAL(i).md5 Or FileExist(App.Path & "\" & archivo, vbNormal) = False Then

            ReDim Preserve ListaActualizar(Desactualizados + 1) As String

            'Aumentamos el contador de la cantidad de archivos para actualizar
            Desactualizados = Desactualizados + 1

            'Añadimos el archivo a la lista para actualizar mas tarde
            ListaActualizar(Desactualizados) = ArchivosREMOTE(i).archivo

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

    For i = 1 To TotalCarpetasREMOTE

        Directorio = Replace(CarpetasREMOTE(i), "-", "\")

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

        frmMain.lblPendientes.Caption = "Descargando '" & archivo & "', archivo " & i & " de " & Desactualizados & ". Por favor, espere."

        'Primero lo adaptamos a URL
        archivoURL = Replace$(ListaActualizar(i), "-", "/")

        'Luego a directorio de Windows
        archivo = Replace$(ListaActualizar(i), "-", "\")
        
        frmMain.ucAsyncDLHost.AddDownloadJob URLUpdate & "cliente/" & UpdateNumberREMOTE & "/" & archivoURL, archivo

        DoEvents

    Next i
    
    'Esto se tiene que mejorar
    frmMain.ucAsyncDLHost.AddDownloadJob URLUpdate & "VersionInfo.json", App.Path & "\INIT\VersionInfo.json"
    
    ActualizacionesPendientes = False
    
    'frmMain.lblPendientes.Caption = "Cliente actualizado."

End Function
