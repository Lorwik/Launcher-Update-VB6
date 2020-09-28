Attribute VB_Name = "modCripto"
Option Explicit
 
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)

Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)
 
Private Function MD5String(ByVal p As String) As String
        
        On Error GoTo MD5String_Err
        

        Dim r As String * 32, t As Long

100     r = Space(32)
102     t = Len(p)
104     MDStringFix p, t, r
106     MD5String = r
        
        Exit Function

MD5String_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modCripto.MD5String " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function
 
Private Function MD5File(ByVal f As String) As String
        
        On Error GoTo MD5File_Err
        

        Dim r As String * 32

100     r = Space(32)
102     MDFile f, r
104     MD5File = r
        
        Exit Function

MD5File_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modCripto.MD5File " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Function ComprobarHash(ByVal File As String) As Boolean
        '***************************************************
        'Autor: Lorwik
        'Fecha: 26/06/2020
        'Descripcion: Comprueba la integridad de un archivo recientemente descargado
        '***************************************************
        
        On Error GoTo ComprobarHash_Err
        

        Dim i          As Integer

        Dim Encontrado As Integer

        Dim Hash       As String
    
100     Encontrado = 0
    
        '********************************************
        'ACTUALIZAR LAUNCHER
        '********************************************
102     If LauncherDesactualizado Then
104         Hash = MD5File(File)
            
106         If updateREMOTE.LauncherCheck <> UCase(Trim(Hash)) Then '¿No Coincide?
108             Call LauncherLog("Hash del Launcher no coincide. Hash del archivo: " & UCase(Trim(Hash)))
110             ComprobarHash = False

                Exit Function
                
            Else '¿Coincide?
            
112             Call WriteVar(LocalFile, "MANIFEST", "checksum", UCase(Trim(Hash)))
                
114             ComprobarHash = True

                Exit Function
                
            End If
            
            Exit Function

        End If
    
        '********************************************
        'ACTUALIZAR RESTO DE ARCHIVOS
        '********************************************
    
116     For i = 1 To Desactualizados

118         If DesactualizadosList(i).Archivo = File Then
120             Encontrado = i

                Exit For

            End If
        
122     Next i
        
124     If Encontrado > 0 Then '¿Lo encontro?
126         Hash = MD5File(File)
            
128         If DesactualizadosList(Encontrado).md5 <> UCase(Trim(Hash)) Then '¿No Coincide?
130             Call LauncherLog("Hash del archivo " & DesactualizadosList(Encontrado).Archivo & " no coincide " & " Hash del archivo: " & UCase(Trim(Hash)))
132             ComprobarHash = False

                Exit Function
                
            Else '¿Coincide?
                
134             Call ActualizarVersionInfo(DesactualizadosList(Encontrado).Archivo, DesactualizadosList(Encontrado).md5)
                
136             ComprobarHash = True

                Exit Function
                
            End If
            
        Else '¿No lo encontro?
        
138         Debug.Print "No se encontro el archivo " & File & " con Indice: " & Encontrado
140         Call LauncherLog("No se encontro el archivo " & File & " con Indice: " & Encontrado)
142         ComprobarHash = False

            Exit Function

        End If

        
        Exit Function

ComprobarHash_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modCripto.ComprobarHash " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Function ComprobarIntegridad() As Integer
        '***************************************************
        'Autor: Lorwik
        'Fecha: 26/06/2020
        'Descripcion: Comprueba la integridad de los archivos existente con la version Local
        '***************************************************
        
        On Error GoTo ComprobarIntegridad_Err
        

        Dim i     As Integer

        Dim Count As Integer
    
100     For i = 0 To UpdateLocal.TotalFiles
    
            'Exclusion de archivos (hay que cambiarlo)
102         If UpdateLocal.Archivos(i).Archivo <> "Init\Config.ini" And UpdateLocal.Archivos(i).Archivo <> "Init\BindKeys.bin" Then
    
                '¿El MD5 guardado en local NO coincide con el obtenido del archivo?
104             If UCase(UpdateLocal.Archivos(i).md5) <> UCase(MD5File(UpdateLocal.Archivos(i).Archivo)) Then
            
106                 Call NuevoDesactualizado(UpdateLocal.Archivos(i).Archivo, UpdateLocal.Archivos(i).md5)
108                 Count = Count + 1 'Llevamos el control de archivos que no se pudieron comprobar
110                 ActualizacionesPendientes = True
                End If
            End If
    
112     Next i
    
114     ComprobarIntegridad = Count
    
        
        Exit Function

ComprobarIntegridad_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.modCripto.ComprobarIntegridad " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function
