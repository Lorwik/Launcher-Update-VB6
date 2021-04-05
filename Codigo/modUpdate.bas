Attribute VB_Name = "modUpdate"
Option Explicit

Private Const URLUPDATE As String = "http://winterao.com.ar/update/"
Private Const VERSIONINFOJSON As String = "VersionInfo.json"
Private Const VERSIONINFOINI As String = "VersionInfo.ini"
Public CLIENTE_FOLDER As String
Public CLIENTEXE As String

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

Public Fallaron As String

Public Function LocalFile() As String
    LocalFile = App.Path & "\" & CLIENTE_FOLDER & "\Init\" & VERSIONINFOINI
End Function

Public Sub SetURLModo()

    Select Case ServerSelect
    
        Case eLaunchMode.Winter
            CLIENTE_FOLDER = "WinterClient"
            CLIENTEXE = "\WinterAOResurrection.exe"
            
        Case eLaunchMode.ImpC
            CLIENTE_FOLDER = "ImpcClient"
            CLIENTEXE = "\ImperiumClasico.exe"
    End Select

End Sub

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

    '**********************
    'ARCHIVOS
    '**********************
    ReDim UpdateLocal.Archivos(0 To UpdateLocal.TotalFiles) As tArchivos

    For i = 1 To UpdateLocal.TotalFiles

        UpdateLocal.Archivos(i).Archivo = GetVar(LocalFile, "File" & i, "name")
        UpdateLocal.Archivos(i).md5 = UCase(GetVar(LocalFile, "File" & i, "checksum"))

    Next i

    '**********************
    'CARPETAS
    '**********************
    ReDim UpdateLocal.Carpetas(0 To UpdateLocal.TotalCarpetas) As String

    For i = 1 To UpdateLocal.TotalCarpetas

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

    responseServer = frmMain.Inet1.OpenURL(URLUPDATE & CLIENTE_FOLDER & "/" & VERSIONINFOJSON)

    Set updateREMOTE.JsonListas = ModJson.parse(responseServer)

    updateREMOTE.TotalFiles = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALFILES"))
    updateREMOTE.TotalCarpetas = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALFOLDERS"))
    
    '**********************
    'ARCHIVOS
    '**********************
    ReDim updateREMOTE.Archivos(0 To updateREMOTE.TotalFiles) As tArchivos

    For i = 1 To updateREMOTE.TotalFiles

        updateREMOTE.Archivos(i).Archivo = updateREMOTE.JsonListas.Item("FILES").Item("FILE" & i).Item("NAME")
        updateREMOTE.Archivos(i).Archivo = Replace$(updateREMOTE.Archivos(i).Archivo, "\/", "\\")

        updateREMOTE.Archivos(i).md5 = UCase(updateREMOTE.JsonListas.Item("FILES").Item("FILE" & i).Item("CHECKSUM"))

    Next i

    '**********************
    'CARPETAS
    '**********************
    ReDim updateREMOTE.Carpetas(0 To updateREMOTE.TotalCarpetas) As String

    For i = 1 To updateREMOTE.TotalCarpetas

        updateREMOTE.Carpetas(i) = updateREMOTE.JsonListas.Item("FOLDERS").Item("FOLDER" & i - 1).Item("NAME")
        updateREMOTE.Carpetas(i) = Replace$(updateREMOTE.Carpetas(i), "\/", "\\")

    Next i
    
End Sub

Public Sub CompararyCrearCarpetas()
'********************************************
'Autor: Lorwik
'Fecha: 17/09/2020
'Descripción: Crea las carpetas que no existan
'********************************************
    On Error Resume Next

    Dim i As Integer

    For i = 1 To updateREMOTE.TotalCarpetas
    
        If Not FileExist(App.Path & "\" & CLIENTE_FOLDER & "\" & updateREMOTE.Carpetas(i), vbDirectory) Then

            MkDir App.Path & "\" & CLIENTE_FOLDER & "\" & updateREMOTE.Carpetas(i)
            DoEvents

        End If

    Next i

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

    '**********************
    'LAUNCHER
    '**********************
    If updateREMOTE.Archivos(0).md5 <> UpdateLocal.Archivos(0).md5 Then
    
        Call NuevoDesactualizado(updateREMOTE.Archivos(0).Archivo, updateREMOTE.Archivos(0).md5)
        
    End If

    '**********************
    'RESTO DE ARCHIVOS
    '**********************
    For i = 1 To updateREMOTE.TotalFiles

        'Comprobamos todos los CHECK
        If updateREMOTE.Archivos(i).md5 <> UpdateLocal.Archivos(i).md5 Or FileExist(App.Path & "\" & CLIENTE_FOLDER & "\" & updateREMOTE.Archivos(i).Archivo, vbNormal) = False Then

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

Public Sub ActualizarCliente()
'***************************************************
'Autor: Lorwik
'Fecha: 25/01/2021
'Descripcion: Descarga los archivos guardados en el array
'***************************************************

    Static Finalizados As Integer
    Dim i As Integer

    Call addConsole("Buscando Actualizaciones...", 255, 255, 255, True, False)
    
    If Not (Desactualizados = 0) Then
        Call addConsole("Iniciando, se descargarán " & Desactualizados & " actualizaciones.", 200, 200, 200, True, False)   '>> Informacion
            
        For i = 1 To Desactualizados
            frmMain.Inet1.AccessType = icUseDefault
                  
            Call addConsole("Descargando " & DesactualizadosList(i).Archivo, 0, 255, 0, True, False)
            Call LauncherLog("Descargando " & DesactualizadosList(i).Archivo)
            frmMain.lblPendientes.Caption = "Descargando archivo " & i & " de " & Desactualizados
            
            frmMain.Inet1.URL = URLUPDATE & CLIENTE_FOLDER & "/" & DesactualizadosList(i).Archivo  'Host
            Directory = App.Path & "\" & CLIENTE_FOLDER & "\" & DesactualizadosList(i).Archivo
            
            bDone = False
            dError = False
                    
            frmMain.Inet1.Execute , "GET"
                
            Do While bDone = False
                DoEvents
            Loop
                
            If dError Then Exit Sub
            
            Call ComprobarHash(DesactualizadosList(i).Archivo)
            Finalizados = Finalizados + 1
        Next i
    End If
    
    '¿El cliente se actualizo correctamente?
    If Finalizados >= Desactualizados Then
        frmMain.lblPendientes.Caption = "Cliente actualizado."
        Call addConsole("Cliente actualizado correctamente, listo para jugar.", 255, 255, 0, True, False)
        
        ActualizacionesPendientes = False

        Desactualizados = 0
        ReDim DesactualizadosList(Desactualizados) As tArchivos
        
    Else
        frmMain.lblPendientes.Caption = "Error en la actualización."
        Call addConsole("Alguno de los archivos no se pudo descargar correctamente.", 255, 0, 0, True, False)
        
    End If
    
    If SinVersiones Then Call ObtenerVersionFile
    
    frmMain.cmdJugar.Enabled = True

End Sub

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
        
        For i = 1 To .TotalFiles
        
            Call WriteVar(LocalFile, "File" & i, "name", .Archivos(i).Archivo)
            Call WriteVar(LocalFile, "File" & i, "checksum", .Archivos(i).md5)
        
        Next i
        
        For i = 1 To .TotalCarpetas
        
            Call WriteVar(LocalFile, "Folder" & i, "name", .Carpetas(i))
        
        Next i
        
        'Ya tenemos el archivo en local, ahora lo cargamos
        Call CargarListasLOCAL
    
    End With
    
End Sub
