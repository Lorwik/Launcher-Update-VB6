Attribute VB_Name = "ModUpdate"
Option Explicit

Public Inet                   As clsInet

Private Const URLUPDATE       As String = "http://winterao.com.ar/update/"
Private Const VERSIONINFOJSON As String = "VersionInfo.json"
Private Const VERSIONINFOINI  As String = "VersionInfo.ini"
Public Const LAUNCHEREXEUP    As String = "WinterAOLauncher.exe.up"

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
Public UpdateLocal               As tUpdate
Public updateREMOTE              As tUpdate
'************************

Public SinVersiones              As Boolean 'Si no encontro el VersionInfo
Public Desactualizados           As Integer 'Numero de archivos desactualizados
Public DesactualizadosList()     As tArchivos

'Indica si hay actualizaciones pendientes
Public ActualizacionesPendientes As Boolean
Public LauncherDesactualizado    As Boolean
Public Fallaron                  As String

Public Function LocalFile() As String
        
        On Error GoTo LocalFile_Err
        
100     LocalFile = App.Path & "\Init\" & VERSIONINFOINI
        
        Exit Function

LocalFile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModUpdate.LocalFile " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Sub CargarListasLOCAL()
        '********************************************
        'Autor: Lorwik
        'Fecha: 17/09/2020
        'Descripción: Carga la lista de versiones LOCAL
        '********************************************
        
        On Error GoTo CargarListasLOCAL_Err
        

        Dim Archivo As String

        Dim i       As Integer

        '¿Existe el archivo de versiones en el directorio local?
100     If Not FileExist(LocalFile, vbArchive) Then
102         SinVersiones = True

            Exit Sub

        End If

104     UpdateLocal.TotalFiles = Val(GetVar(LocalFile, "MANIFEST", "TOTALFILES"))
106     UpdateLocal.TotalCarpetas = Val(GetVar(LocalFile, "MANIFEST", "TOTALFOLDERS"))
108     UpdateLocal.LauncherCheck = GetVar(LocalFile, "MANIFEST", "checksum")

110     ReDim UpdateLocal.Archivos(0 To UpdateLocal.TotalFiles) As tArchivos

112     For i = 0 To UpdateLocal.TotalFiles

114         UpdateLocal.Archivos(i).Archivo = GetVar(LocalFile, "File" & i, "name")
116         UpdateLocal.Archivos(i).md5 = UCase(GetVar(LocalFile, "File" & i, "checksum"))

118     Next i

120     ReDim UpdateLocal.Carpetas(0 To UpdateLocal.TotalCarpetas) As String

122     For i = 0 To UpdateLocal.TotalCarpetas

124         UpdateLocal.Carpetas(i) = GetVar(LocalFile, "Folder" & i, "name")

126     Next i

        
        Exit Sub

CargarListasLOCAL_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModUpdate.CargarListasLOCAL " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Public Sub CargarListasREMOTE()
        '********************************************
        'Autor: Lorwik
        'Fecha: 17/09/2020
        'Descripción: Carga la lista de versiones REMOTA
        '********************************************
        
        On Error GoTo CargarListasREMOTE_Err
        

        Dim responseServer As String

        Dim Archivo        As String

        Dim i              As Integer

100     Set Inet = New clsInet

102     responseServer = Inet.OpenRequest(URLUPDATE & VERSIONINFOJSON, "GET")
104     responseServer = Inet.Execute
106     responseServer = Inet.GetResponseAsString
        
        ' El servidor no respondio nada?
        If LenB(responseServer) = 0 Then
            Call MsgBox("Hemos recibido una respuesta invalida del servidor", "Carga la lista de versiones Remota")
            Exit Sub
        End If
        
108     Set updateREMOTE.JsonListas = ModJson.parse(responseServer)

110     updateREMOTE.TotalFiles = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TOTALFILES"))
112     updateREMOTE.TotalCarpetas = Val(updateREMOTE.JsonListas.Item("MANIFEST").Item("TotalFolders"))
114     updateREMOTE.LauncherCheck = updateREMOTE.JsonListas.Item("MANIFEST").Item("checksum")

116     ReDim updateREMOTE.Archivos(0 To updateREMOTE.TotalFiles) As tArchivos

118     For i = 0 To updateREMOTE.TotalFiles

120         updateREMOTE.Archivos(i).Archivo = updateREMOTE.JsonListas.Item("Files").Item("File" & i).Item("name")
122         updateREMOTE.Archivos(i).md5 = UCase(updateREMOTE.JsonListas.Item("Files").Item("File" & i).Item("checksum"))

124     Next i

126     ReDim updateREMOTE.Carpetas(0 To updateREMOTE.TotalCarpetas) As String

128     For i = 0 To updateREMOTE.TotalCarpetas

130         updateREMOTE.Carpetas(i) = updateREMOTE.JsonListas.Item("Folders").Item("Folder" & i).Item("name")

132     Next i

134     Set Inet = Nothing

        
        Exit Sub

CargarListasREMOTE_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModUpdate.CargarListasREMOTE " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Public Function CompararArchivos() As Boolean
        '********************************************
        'Autor: Lorwik
        'Fecha: 17/09/2020
        'Descripción: Comprueba si el CHECK de los archivos coinciden con el del remoto
        '********************************************
        
        On Error GoTo CompararArchivos_Err
        

        Dim i       As Integer

        Dim flag    As Boolean

        Dim Archivo As String

        'El total de archivos remoto es diferente al de local? 'Hay que actualizar seguro.
100     If updateREMOTE.TotalFiles <> UpdateLocal.TotalFiles Then
102         CompararArchivos = True

            Exit Function

        End If

104     For i = 0 To updateREMOTE.TotalFiles

            'Comprobamos todos los CHECK
106         If updateREMOTE.Archivos(i).md5 <> UpdateLocal.Archivos(i).md5 Or FileExist(App.Path & "\" & updateREMOTE.Archivos(i).Archivo, vbNormal) = False Then

108             Call NuevoDesactualizado(updateREMOTE.Archivos(i).Archivo, updateREMOTE.Archivos(i).md5)

110             flag = True 'Activamos el flag
            End If

112     Next i

        '¿Hubo algun archivo desactualizado?
114     If flag Then
116         CompararArchivos = True
        Else
118         CompararArchivos = False
        End If

        
        Exit Function

CompararArchivos_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModUpdate.CompararArchivos " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
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
        
        On Error GoTo ActualizarCliente_Err
        

        Dim i          As Integer

        Dim Archivo    As String

        Dim archivoURL As String

100     If LauncherDesactualizado Then
102         If FileExist(App.Path & "\" & LAUNCHEREXEUP, vbNormal) Then Kill App.Path & "\" & LAUNCHEREXEUP
104         frmMain.ucAsyncDLHost.AddDownloadJob URLUPDATE & "cliente/" & LAUNCHEREXEUP, LAUNCHEREXEUP

106         DoEvents
        End If
    
108     If Desactualizados > 0 Then

110         For i = 1 To Desactualizados
    
                'Lo adaptamos a URL
112             archivoURL = Replace$(DesactualizadosList(i).Archivo, "\\", "/")
            
114             frmMain.ucAsyncDLHost.AddDownloadJob URLUPDATE & "cliente/" & archivoURL, DesactualizadosList(i).Archivo
    
116             DoEvents
    
118         Next i

        End If
    
120     If SinVersiones Then Call ObtenerVersionFile
    
122     ActualizacionesPendientes = False

        
        Exit Function

ActualizarCliente_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModUpdate.ActualizarCliente " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
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
        
        On Error GoTo NuevoDesactualizado_Err
        

        Dim i As Integer
    
        'Si ya existe el archivo en la lista de desactualizados, no lo agregamos
100     If Desactualizados > 0 Then

102         For i = 1 To Desactualizados

104             If DesactualizadosList(i).Archivo = File Then Exit Sub
106         Next i

        End If

108     ReDim Preserve DesactualizadosList(Desactualizados + 1) As tArchivos

        'Aumentamos el contador de la cantidad de archivos para actualizar
110     Desactualizados = Desactualizados + 1

        'Añadimos el archivo a la lista para actualizar mas tarde
112     DesactualizadosList(Desactualizados).Archivo = File
114     DesactualizadosList(Desactualizados).md5 = Checksum
            
        
        Exit Sub

NuevoDesactualizado_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModUpdate.NuevoDesactualizado " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub
