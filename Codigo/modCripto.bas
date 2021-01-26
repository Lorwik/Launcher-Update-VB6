Attribute VB_Name = "modCripto"
Option Explicit
 
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)
 
Private Function MD5String(ByVal p As String) As String
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function
 
Private Function MD5File(ByVal f As String) As String

    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function

Public Function ComprobarHash(ByVal File As String) As Boolean
'***************************************************
'Autor: Lorwik
'Fecha: 26/06/2020
'Descripcion: Comprueba la integridad de un archivo recientemente descargado
'***************************************************

    Dim i As Integer
    Dim Encontrado As Integer
    Dim Hash As String
    
    Encontrado = 0

    For i = 1 To Desactualizados
        If DesactualizadosList(i).Archivo = File Then
            Encontrado = i
            Exit For
        End If
        
    Next i
        
    If Encontrado > 0 Then '¿Lo encontro?
        Hash = MD5File(File)
            
        If DesactualizadosList(Encontrado).md5 <> UCase(Trim(Hash)) Then '¿No Coincide?
            Call LauncherLog("Hash del archivo " & DesactualizadosList(Encontrado).Archivo & " no coincide " & " Hash del archivo: " & UCase(Trim(Hash)))
            ComprobarHash = False
            Exit Function
                
        Else '¿Coincide?
                
            Call ActualizarVersionInfo(DesactualizadosList(Encontrado).Archivo, DesactualizadosList(Encontrado).md5)
                
            ComprobarHash = True
            Exit Function
                
        End If
            
    Else '¿No lo encontro?
        
        Debug.Print "No se encontro el archivo " & File & " con Indice: " & Encontrado
        Call LauncherLog("No se encontro el archivo " & File & " con Indice: " & Encontrado)
        ComprobarHash = False
        Exit Function
    End If

End Function

Public Function ComprobarIntegridad() As Integer
'***************************************************
'Autor: Lorwik
'Fecha: 26/06/2020
'Descripcion: Comprueba la integridad de los archivos existente con la version Local
'***************************************************

    Dim i As Integer
    Dim Count As Integer
    
    '***********
    'LAUNCHER
    '***********
    
    '¿Se descargo una actualización para el launcher?
    If FileExist(App.Path & UpdateLocal.Archivos(0).Archivo, vbNormal) Then
        '¿El MD5 guardado en local NO coincide con el obtenido del archivo?
        If UCase(UpdateLocal.Archivos(0).md5) <> UCase(MD5File(UpdateLocal.Archivos(0).Archivo)) Then
            
            Call NuevoDesactualizado(UpdateLocal.Archivos(0).Archivo, UpdateLocal.Archivos(0).md5)
            Count = Count + 1 'Llevamos el control de archivos que no se pudieron comprobar
            ActualizacionesPendientes = True
        End If
    End If
    
    '***********
    'ARCHIVOS
    '***********
    
    For i = 1 To UpdateLocal.TotalFiles
    
        'Exclusion de archivos (hay que cambiarlo)
        If UpdateLocal.Archivos(i).Archivo <> "Init\Config.ini" And UpdateLocal.Archivos(i).Archivo <> "Init\BindKeys.bin" Then
    
            '¿El MD5 guardado en local NO coincide con el obtenido del archivo?
            If UCase(UpdateLocal.Archivos(i).md5) <> UCase(MD5File(UpdateLocal.Archivos(i).Archivo)) Then
            
                Call NuevoDesactualizado(UpdateLocal.Archivos(i).Archivo, UpdateLocal.Archivos(i).md5)
                Count = Count + 1 'Llevamos el control de archivos que no se pudieron comprobar
                ActualizacionesPendientes = True
            End If
        End If
    
    Next i
    
    ComprobarIntegridad = Count
    
End Function

