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
    Dim i As Integer
    Dim Encontrado As Integer
    Dim Hash As String
    
    Encontrado = 0
    
    '********************************************
    'ACTUALIZAR LAUNCHER
    '********************************************
    If LauncherDesactualizado Then
        Hash = MD5File(File)
            
        If updateREMOTE.LauncherCheck <> UCase(Trim(Hash)) Then '�No Coincide?
            Call LauncherLog("Hash del Launcher no coincide. Hash del archivo: " & UCase(Trim(Hash)))
            ComprobarHash = False
            Exit Function
                
        Else '�Coincide?
            
            Call WriteVar(LocalFile, "MANIFEST", "CHECK", UCase(Trim(Hash)))
                
            ComprobarHash = True
            Exit Function
                
        End If
            
        Exit Function
    End If
    
    '********************************************
    'ACTUALIZAR RESTO DE ARCHIVOS
    '********************************************
    
    For i = 1 To Desactualizados
        If DesactualizadosList(i).Archivo = Replace$(File, "\", "-") Then
            Encontrado = i
            Exit For
        End If
        
    Next i
        
    If Encontrado > 0 Then '�Lo encontro?
        Hash = MD5File(File)
            
        If DesactualizadosList(Encontrado).md5 <> UCase(Trim(Hash)) Then '�No Coincide?
            Call LauncherLog("Hash del archivo " & DesactualizadosList(Encontrado).Archivo & " no coincide " & " Hash del archivo: " & UCase(Trim(Hash)))
            ComprobarHash = False
            Exit Function
                
        Else '�Coincide?
                
            Call ActualizarVersionInfo(DesactualizadosList(Encontrado).Archivo, DesactualizadosList(Encontrado).md5)
                
            ComprobarHash = True
            Exit Function
                
        End If
            
    Else '�No lo encontro?
        
        Debug.Print "No se encontro el archivo " & File & " con Indice: " & Encontrado
        Call LauncherLog("No se encontro el archivo " & File & " con Indice: " & Encontrado)
        ComprobarHash = False
        Exit Function
    End If

End Function
