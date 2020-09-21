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
    
    For i = 1 To Desactualizados

        If ArchivosREMOTE(i).archivo = Replace$(File, "\", "-") Then
            Encontrado = i
            Exit For
        End If
    
    Next i
    
    If Encontrado > 0 Then '¿Lo encontro?
        Hash = MD5File(File)
        
        If ArchivosREMOTE(Encontrado).md5 <> Trim(Hash) Then '¿No Coincide?
            Debug.Print "Hash no coincide Original: " & ArchivosREMOTE(Encontrado).md5 & " Hash del archivo: " & Hash
            ComprobarHash = False
            Exit Function
            
        Else '¿Coincide?
            Debug.Print File & " coincide!"
            ComprobarHash = True
            Exit Function
            
        End If
        
    Else '¿No lo encontro?
    
        Debug.Print "No se encontrol el archivo " & File & " con Indice: " & Encontrado
        ComprobarHash = False
        Exit Function
    End If
    
End Function
