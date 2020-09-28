Attribute VB_Name = "ModJson"

' VBJSONDeserializer is a VB6 adaptation of the VB-JSON project @
' Fuente: https://www.codeproject.com/Articles/720368/VB-JSON-Parser-Improved-Performance

' BSD Licensed

Option Explicit

' DECLARACIONES API
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

' CONSTANTES LOCALE API
Public Const LOCALE_SDECIMAL = &HE

Public Const LOCALE_SGROUPING = &H10

' CONSTANTES JSON
Private Const A_CURLY_BRACKET_OPEN   As Integer = 123       ' AscW("{")

Private Const A_CURLY_BRACKET_CLOSE  As Integer = 125       ' AscW("}")

Private Const A_SQUARE_BRACKET_OPEN  As Integer = 91        ' AscW("[")

Private Const A_SQUARE_BRACKET_CLOSE As Integer = 93        ' AscW("]")

Private Const A_BRACKET_OPEN         As Integer = 40        ' AscW("(")

Private Const A_BRACKET_CLOSE        As Integer = 41        ' AscW(")")

Private Const A_COMMA                As Integer = 44        ' AscW(",")

Private Const A_DOUBLE_QUOTE         As Integer = 34        ' AscW("""")

Private Const A_SINGLE_QUOTE         As Integer = 39        ' AscW("'")

Private Const A_BACKSLASH            As Integer = 92        ' AscW("\")

Private Const A_FORWARDSLASH         As Integer = 47        ' AscW("/")

Private Const A_COLON                As Integer = 58        ' AscW(":")

Private Const A_SPACE                As Integer = 32        ' AscW(" ")

Private Const A_ASTERIX              As Integer = 42        ' AscW("*")

Private Const A_VBCR                 As Integer = 13        ' AscW("vbcr")

Private Const A_VBLF                 As Integer = 10        ' AscW("vblf")

Private Const A_VBTAB                As Integer = 9         ' AscW("vbTab")

Private Const A_VBCRLF               As Integer = 13        ' AscW("vbcrlf")

Private Const A_b                    As Integer = 98        ' AscW("b")

Private Const A_f                    As Integer = 102       ' AscW("f")

Private Const A_n                    As Integer = 110       ' AscW("n")

Private Const A_r                    As Integer = 114       ' AscW("r"

Private Const A_t                    As Integer = 116       ' AscW("t"))

Private Const A_u                    As Integer = 117       ' AscW("u")

Private m_decSep                     As String

Private m_groupSep                   As String

Private m_parserrors                 As String

Private m_str()                      As Integer

Private m_length                     As Long

Public Function GetParserErrors() As String
        
        On Error GoTo GetParserErrors_Err
        
100     GetParserErrors = m_parserrors
        
        Exit Function

GetParserErrors_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.GetParserErrors " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Function parse(ByRef str As String) As Object

    m_decSep = GetRegionalSettings(LOCALE_SDECIMAL)
    m_groupSep = GetRegionalSettings(LOCALE_SGROUPING)

    Dim Index As Long

    Index = 1

    GenerateStringArray str

    m_parserrors = vbNullString

    On Error Resume Next

    Call skipChar(Index)

    Select Case m_str(Index)

        Case A_SQUARE_BRACKET_OPEN
            Set parse = parseArray(str, Index)

        Case A_CURLY_BRACKET_OPEN
            Set parse = parseObject(str, Index)

        Case Else
            m_parserrors = "JSON Invalido"

    End Select

    'clean array
    ReDim m_str(1)

End Function

Private Sub GenerateStringArray(ByRef str As String)
        
        On Error GoTo GenerateStringArray_Err
        

        Dim i As Long

100     m_length = Len(str)
102     ReDim m_str(1 To m_length)

104     For i = 1 To m_length
106         m_str(i) = AscW(Mid$(str, i, 1))
108     Next i

        
        Exit Sub

GenerateStringArray_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.GenerateStringArray " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Private Function parseObject(ByRef str As String, ByRef Index As Long) As Dictionary

    Set parseObject = New Dictionary

    Dim sKey    As String

    Dim charint As Integer

    Call skipChar(Index)

    If m_str(Index) <> A_CURLY_BRACKET_OPEN Then
        m_parserrors = m_parserrors & "Objeto invalido en la posicion " & Index & " : " & Mid$(str, Index) & vbCrLf

        Exit Function

    End If

    Index = Index + 1

    Do
    
        Call skipChar(Index)
    
        charint = m_str(Index)
    
        If charint = A_COMMA Then
            Index = Index + 1
            Call skipChar(Index)
        ElseIf charint = A_CURLY_BRACKET_CLOSE Then
            Index = Index + 1

            Exit Do

        ElseIf Index > m_length Then
            m_parserrors = m_parserrors & "Falta '}': " & Right$(str, 20) & vbCrLf

            Exit Do

        End If

        ' add key/value pair
        sKey = parseKey(Index)

        On Error Resume Next

        Call parseObject.Add(sKey, parseValue(str, Index))

        If Err.Number <> 0 Then
            m_parserrors = m_parserrors & Err.Description & ": " & sKey & vbCrLf
            Exit Do
        End If

    Loop

End Function

Private Function parseArray(ByRef str As String, ByRef Index As Long) As Collection

    Dim charint As Integer

    Set parseArray = New Collection

    Call skipChar(Index)

    If Mid$(str, Index, 1) <> "[" Then
        m_parserrors = m_parserrors & "Array invalido en la posicion " & Index & " : " + Mid$(str, Index, 20) & vbCrLf

        Exit Function

    End If
   
    Index = Index + 1

    Do
        Call skipChar(Index)
    
        charint = m_str(Index)
    
        If charint = A_SQUARE_BRACKET_CLOSE Then
            Index = Index + 1

            Exit Do

        ElseIf charint = A_COMMA Then
            Index = Index + 1
            Call skipChar(Index)
        ElseIf Index > m_length Then
            m_parserrors = m_parserrors & "Falta ']': " & Right$(str, 20) & vbCrLf

            Exit Do

        End If
    
        'add value
        On Error Resume Next

        parseArray.Add parseValue(str, Index)

        If Err.Number <> 0 Then
            m_parserrors = m_parserrors & Err.Description & ": " & Mid$(str, Index, 20) & vbCrLf

            Exit Do

        End If

    Loop

End Function

Private Function parseValue(ByRef str As String, ByRef Index As Long)
        
        On Error GoTo parseValue_Err
        

100     Call skipChar(Index)

102     Select Case m_str(Index)

            Case A_DOUBLE_QUOTE, A_SINGLE_QUOTE
104             parseValue = parseString(str, Index)

                Exit Function

106         Case A_SQUARE_BRACKET_OPEN
108             Set parseValue = parseArray(str, Index)

                Exit Function

110         Case A_t, A_f
112             parseValue = parseBoolean(str, Index)

                Exit Function

114         Case A_n
116             parseValue = parseNull(str, Index)

                Exit Function

118         Case A_CURLY_BRACKET_OPEN
120             Set parseValue = parseObject(str, Index)

                Exit Function

122         Case Else
124             parseValue = parseNumber(str, Index)

                Exit Function

        End Select

        
        Exit Function

parseValue_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.parseValue " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Private Function parseString(ByRef str As String, ByRef Index As Long) As String
        
        On Error GoTo parseString_Err
        

        Dim quoteint As Integer

        Dim charint  As Integer

        Dim code     As String
   
100     Call skipChar(Index)
   
102     quoteint = m_str(Index)
   
104     Index = Index + 1
   
106     Do While Index > 0 And Index <= m_length
   
108         charint = m_str(Index)
      
110         Select Case charint

                Case A_BACKSLASH

112                 Index = Index + 1
114                 charint = m_str(Index)

116                 Select Case charint

                        Case A_DOUBLE_QUOTE, A_BACKSLASH, A_FORWARDSLASH, A_SINGLE_QUOTE
118                         parseString = parseString & ChrW$(charint)
120                         Index = Index + 1

122                     Case A_b
124                         parseString = parseString & vbBack
126                         Index = Index + 1

128                     Case A_f
130                         parseString = parseString & vbFormFeed
132                         Index = Index + 1

134                     Case A_n
136                         parseString = parseString & vbLf
138                         Index = Index + 1

140                     Case A_r
142                         parseString = parseString & vbCr
144                         Index = Index + 1

146                     Case A_t
148                         parseString = parseString & vbTab
150                         Index = Index + 1

152                     Case A_u
154                         Index = Index + 1
156                         code = Mid$(str, Index, 4)

158                         parseString = parseString & ChrW$(Val("&h" + code))
160                         Index = Index + 4

                    End Select

162             Case quoteint
        
164                 Index = Index + 1

                    Exit Function

166             Case Else
168                 parseString = parseString & ChrW$(charint)
170                 Index = Index + 1

            End Select

        Loop
   
        
        Exit Function

parseString_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.parseString " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Private Function parseNumber(ByRef str As String, ByRef Index As Long)
        
        On Error GoTo parseNumber_Err
        

        Dim value As String

        Dim Char  As String

100     Call skipChar(Index)

102     Do While Index > 0 And Index <= m_length
104         Char = Mid$(str, Index, 1)

106         If InStr("+-0123456789.eE", Char) Then
108             value = value & Char
110             Index = Index + 1
            Else

                'check what is the grouping seperator
112             If Not m_decSep = "." Then
114                 value = Replace(value, ".", m_decSep)

                End If
     
116             If m_groupSep = "." Then
118                 value = Replace(value, ".", m_decSep)

                End If
     
120             parseNumber = CDec(value)

                Exit Function

            End If

        Loop
   
        
        Exit Function

parseNumber_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.parseNumber " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean
        
        On Error GoTo parseBoolean_Err
        

100     Call skipChar(Index)
   
102     If Mid$(str, Index, 4) = "true" Then
104         parseBoolean = True
106         Index = Index + 4

108     ElseIf Mid$(str, Index, 5) = "false" Then
110         parseBoolean = False
112         Index = Index + 5

        Else
114         m_parserrors = m_parserrors & "Boolean invalido en la posicion " & Index & " : " & Mid$(str, Index) & vbCrLf

        End If

        
        Exit Function

parseBoolean_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.parseBoolean " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Private Function parseNull(ByRef str As String, ByRef Index As Long)
        
        On Error GoTo parseNull_Err
        

100     Call skipChar(Index)
   
102     If Mid$(str, Index, 4) = "null" Then
104         parseNull = Null
106         Index = Index + 4

        Else
108         m_parserrors = m_parserrors & "Valor nulo invalido en la posicion " & Index & " : " & Mid$(str, Index) & vbCrLf

        End If

        
        Exit Function

parseNull_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.parseNull " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Private Function parseKey(ByRef Index As Long) As String
        
        On Error GoTo parseKey_Err
        

        Dim dquote  As Boolean
        Dim squote  As Boolean
        Dim charint As Integer
   
100     Call skipChar(Index)
   
102     Do While Index > 0 And Index <= m_length
    
104         charint = m_str(Index)
        
106         Select Case charint

                Case A_DOUBLE_QUOTE
108                 dquote = Not dquote
110                 Index = Index + 1

112                 If Not dquote Then
            
114                     Call skipChar(Index)
                
116                     If m_str(Index) <> A_COLON Then
118                         m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & Index & " : " & parseKey & vbCrLf

                            Exit Do

                        End If

                    End If

120             Case A_SINGLE_QUOTE
122                 squote = Not squote
124                 Index = Index + 1

126                 If Not squote Then
128                     Call skipChar(Index)
                
130                     If m_str(Index) <> A_COLON Then
132                         m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & Index & " : " & parseKey & vbCrLf

                            Exit Do

                        End If
                
                    End If
        
134             Case A_COLON
136                 Index = Index + 1

138                 If Not dquote And Not squote Then

                        Exit Do

                    Else
140                     parseKey = parseKey & ChrW$(charint)

                    End If

142             Case Else
            
144                 If A_VBCRLF = charint Then
146                 ElseIf A_VBCR = charint Then
148                 ElseIf A_VBLF = charint Then
150                 ElseIf A_VBTAB = charint Then
152                 ElseIf A_SPACE = charint Then
                    Else
154                     parseKey = parseKey & ChrW$(charint)

                    End If

156                 Index = Index + 1

            End Select

        Loop

        
        Exit Function

parseKey_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.parseKey " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Private Sub skipChar(ByRef Index As Long)
        
        On Error GoTo skipChar_Err
        

        Dim bComment      As Boolean

        Dim bStartComment As Boolean

        Dim bLongComment  As Boolean

100     Do While Index > 0 And Index <= m_length
    
102         Select Case m_str(Index)

                Case A_VBCR, A_VBLF

104                 If Not bLongComment Then
106                     bStartComment = False
108                     bComment = False

                    End If
    
110             Case A_VBTAB, A_SPACE, A_BRACKET_OPEN, A_BRACKET_CLOSE
                    'do nothing
        
112             Case A_FORWARDSLASH

114                 If Not bLongComment Then
116                     If bStartComment Then
118                         bStartComment = False
120                         bComment = True
                        Else
122                         bStartComment = True
124                         bComment = False
126                         bLongComment = False

                        End If

                    Else

128                     If bStartComment Then
130                         bLongComment = False
132                         bStartComment = False
134                         bComment = False

                        End If

                    End If

136             Case A_ASTERIX

138                 If bStartComment Then
140                     bStartComment = False
142                     bComment = True
144                     bLongComment = True
                    Else
146                     bStartComment = True

                    End If

148             Case Else
        
150                 If Not bComment Then

                        Exit Do

                    End If

            End Select

152         Index = Index + 1
        Loop

        
        Exit Sub

skipChar_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.skipChar " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Public Function GetRegionalSettings(ByVal regionalsetting As Long) As String
    ' Devuelve la configuracion regional del sistema

    On Error GoTo ErrorHandler

    Dim Locale      As Long

    Dim symbol      As String

    Dim iRet1       As Long

    Dim iRet2       As Long

    Dim lpLCDataVar As String

    Dim Pos         As Integer
      
    Locale = GetUserDefaultLCID()

    iRet1 = GetLocaleInfo(Locale, regionalsetting, lpLCDataVar, 0)
    symbol = String$(iRet1, 0)
    iRet2 = GetLocaleInfo(Locale, regionalsetting, symbol, iRet1)
    Pos = InStr(symbol, Chr$(0))

    If Pos > 0 Then
        symbol = Left$(symbol, Pos - 1)

    End If
      
ErrorHandler:
    GetRegionalSettings = symbol

    Select Case Err.Number

        Case 0

        Case Else
            Err.Raise 123, "GetRegionalSetting", "GetRegionalSetting: " & regionalsetting

    End Select

End Function

'********************************************************************************************************
'                   FUNCIONES MISCELANEAS DE LA ANTERIOR VERSION DEL MODULO
'********************************************************************************************************

Private Function Encode(ByVal str As String) As String
        
        On Error GoTo Encode_Err
        

        Dim SB      As New cStringBuilder

        Dim i       As Long

        Dim j       As Long

        Dim aL1     As Variant

        Dim aL2     As Variant

        Dim c       As String

        Dim p       As Boolean

        Dim Len_str As Long

100     aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
102     aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    
104     Len_str = LenB(str)
    
106     For i = 1 To Len_str
108         p = True
110         c = Mid$(str, i, 1)

112         For j = 0 To 7

114             If c = Chr$(aL1(j)) Then
116                 SB.Append "\" & Chr$(aL2(j))
118                 p = False

                    Exit For

                End If

            Next

120         If p Then

                Dim a As Integer
122             a = AscW(c)

124             If a > 31 And a < 127 Then
126                 SB.Append c
128             ElseIf a > -1 Or a < 65535 Then
130                 SB.Append "\u" & String$(4 - LenB(Hex$(a)), "0") & Hex$(a)

                End If

            End If

        Next
   
132     Encode = SB.toString
134     Set SB = Nothing
   
        
        Exit Function

Encode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.Encode " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Function StringToJSON(st As String) As String
        
        On Error GoTo StringToJSON_Err
        
   
        Const FIELD_SEP = "~"

        Const RECORD_SEP = "|"

        Dim sFlds      As String

        Dim sRecs      As New cStringBuilder

        Dim lRecCnt    As Long

        Dim lFld       As Long

        Dim fld        As Variant

        Dim rows       As Variant
    
        Dim Lower_rows As Long, Upper_rows As Long

        Dim Lower_fld  As Long, Upper_fld As Long

100     lRecCnt = 0

102     If LenB(st) = 0 Then
104         StringToJSON = "null"
        Else
106         rows = Split(st, RECORD_SEP)
        
108         Lower_rows = LBound(rows)
110         Upper_rows = UBound(rows)
        
112         For lRecCnt = Lower_rows To Upper_rows
114             sFlds = vbNullString
116             fld = Split(rows(lRecCnt), FIELD_SEP)
            
118             Lower_fld = LBound(fld)
120             Upper_fld = UBound(fld)
            
122             For lFld = Lower_fld To Upper_fld Step 2
124                 sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
                Next 'fld

126             sRecs.Append IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}"
            Next 'rec

128         StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

        End If

        
        Exit Function

StringToJSON_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.StringToJSON " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

Public Function RStoJSON(rs As ADODB.Recordset) As String

    On Error GoTo errhandler

    Dim sFlds   As String

    Dim sRecs   As New cStringBuilder

    Dim lRecCnt As Long

    Dim fld     As ADODB.Field

    lRecCnt = 0

    If rs.State = adStateClosed Then
        RStoJSON = "null"
    Else

        If rs.EOF Or rs.BOF Then
            RStoJSON = "null"
        Else

            Do While Not rs.EOF And Not rs.BOF
                lRecCnt = lRecCnt + 1
                sFlds = vbNullString

                For Each fld In rs.Fields

                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.name & """:""" & toUnicode(fld.value & "") & """")
                Next 'fld

                sRecs.Append IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}"
                rs.MoveNext
            Loop

            RStoJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

        End If

    End If

    Exit Function

errhandler:

End Function

Public Function toUnicode(str As String) As String
        
        On Error GoTo toUnicode_Err
        

        Dim X        As Long

        Dim uStr     As New cStringBuilder

        Dim uChrCode As Integer
    
        Dim Len_str  As Long

100     Len_str = LenB(str)

102     For X = 1 To Len_str
104         uChrCode = Asc(Mid$(str, X, 1))

106         Select Case uChrCode

                Case 8
                    ' backspace
108                 uStr.Append "\b"

110             Case 9
                    ' tab
112                 uStr.Append "\t"

114             Case 10
                    ' line feed
116                 uStr.Append "\n"

118             Case 12
                    ' formfeed
120                 uStr.Append "\f"

122             Case 13
                    ' carriage return
124                 uStr.Append "\r"

126             Case 34
                    ' quote
128                 uStr.Append "\"""

130             Case 39
                    ' apostrophe
132                 uStr.Append "\'"

134             Case 92
                    ' backslash
136                 uStr.Append "\\"

138             Case 123, 125
                    ' "{" and "}"
140                 uStr.Append ("\u" & Right$("0000" & Hex$(uChrCode), 4))

142             Case Is < 32, Is > 127
                    ' non-ascii characters
144                 uStr.Append ("\u" & Right$("0000" & Hex$(uChrCode), 4))

146             Case Else
148                 uStr.Append Chr$(uChrCode)

            End Select

        Next

150     toUnicode = uStr.toString

        Exit Function

        
        Exit Function

toUnicode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in WinterAOLauncher.ModJson.toUnicode " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Function

