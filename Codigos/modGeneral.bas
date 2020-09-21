Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Public Sub Main()
    Dim i As Integer

    'Cargamos listas de archivos y carpetas
    Call CargarListasLOCAL
    Call CargarListasREMOTE
    
    'Comprobamos y creamos carpetas
    Call CompararyCrearCarpetas
    
    '¿No existe el archivo de versiones?
    If SinVersiones Then
        ActualizacionesPendientes = True
        
        Desactualizados = TotalFilesREMOTE
        
        ReDim Preserve ListaActualizar(Desactualizados) As String
        
        For i = 1 To Desactualizados
            'Añadimos el archivo a la lista para actualizar mas tarde
            ListaActualizar(i) = ArchivosREMOTE(i).archivo
        Next i
        
        frmMain.lblPendientes.Caption = "¡No se ha encontrado el cliente! Pulsa Jugar para descargar los archivos del cliente."
        
    Else
        ActualizacionesPendientes = ModUpdate.CompararArchivos
        
        If ActualizacionesPendientes Then
            frmMain.lblPendientes.Caption = "Hay " & Desactualizados & " archivos desactualizados."
        Else
            frmMain.lblPendientes.Caption = "Cliente actualizado. Pulsa Jugar para abrir el cliente."
        End If
    End If
    
    DoEvents
    
    frmMain.Show
    
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = Mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Public Function FileToString(strFileName As String) As String
    '###################################################################################
    ' Convierte un archivo entero a una cadena de texto para almacenarla en una variable
    '###################################################################################
    Dim IFile As Variant
    
    IFile = FreeFile
    Open strFileName For Input As #IFile
        FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
End Function

