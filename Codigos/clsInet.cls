VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*** DECLARATIONS FOUND IN WININET.H ***
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetQueryDataAvailable Lib "wininet.dll" (ByVal hHttpRequest As Long, ByRef lpdwNumberOfBytesAvailable As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" (ByVal hHttpRequest As Long, ByVal dwOption As Long, lpBuffer As Any, ByRef dwBufferLength As Long) As Long
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hHttpRequest As Long, ByVal dwOption As Long, lpBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hInternetSession As Long, ByVal lpHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal sAcceptTypes As String, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Long '**********************************************

'**********************************************
' Constants for InternetOpen()
'**********************************************
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0

'**********************************************
' Constants for InternetConnect()
'**********************************************
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443
Private Const INTERNET_SERVICE_HTTP = 3

'**********************************************
' Constants for HttpOpenRequest()
'**********************************************
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100
Private Const INTERNET_FLAG_NO_COOKIES = &H80000
Private Const INTERNET_FLAG_SECURE = &H800000

'**********************************************
' Constants for InternetQueryOption() or InternetSetOption()
'**********************************************
Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Private Const INTERNET_OPTION_SEND_TIMEOUT = 5
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Private Const INTERNET_OPTION_SECURITY_FLAGS = 31
Private Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100

'**********************************************
' Constants for HttpAddRequestHeaders()
'**********************************************
Private Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA = &H40000000
Private Const HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON = &H1000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
Private Const ERROR_INSUFFICIENT_BUFFER = 122
 
Public Enum enumPostHeader

    HTTP_QUERY_ACCEPT = 24
    HTTP_QUERY_ACCEPT_CHARSET = 25
    HTTP_QUERY_ACCEPT_ENCODING = 26
    HTTP_QUERY_ACCEPT_LANGUAGE = 27
    HTTP_QUERY_ACCEPT_RANGES = 42
    HTTP_QUERY_AGE = 48
    HTTP_QUERY_ALLOW = 7
    HTTP_QUERY_AUTHORIZATION = 28
    HTTP_QUERY_CACHE_CONTROL = 49
    HTTP_QUERY_CONNECTION = 23
    HTTP_QUERY_CONTENT_BASE = 50
    HTTP_QUERY_CONTENT_DESCRIPTION = 4
    HTTP_QUERY_CONTENT_DISPOSITION = 47
    HTTP_QUERY_CONTENT_ENCODING = 29
    HTTP_QUERY_CONTENT_ID = 3
    HTTP_QUERY_CONTENT_LANGUAGE = 6
    HTTP_QUERY_CONTENT_LENGTH = 5
    HTTP_QUERY_CONTENT_LOCATION = 51
    HTTP_QUERY_CONTENT_MD5 = 52
    HTTP_QUERY_CONTENT_RANGE = 53
    HTTP_QUERY_CONTENT_TRANSFER_ENCODING = 2
    HTTP_QUERY_CONTENT_TYPE = 1
    HTTP_QUERY_COOKIE = 44
    HTTP_QUERY_COST = 15
    HTTP_QUERY_CUSTOM = 65535
    HTTP_QUERY_DATE = 9
    HTTP_QUERY_DERIVED_FROM = 14
    HTTP_QUERY_ECHO_HEADERS = 73
    HTTP_QUERY_ECHO_HEADERS_CRLF = 74
    HTTP_QUERY_ECHO_REPLY = 72
    HTTP_QUERY_ECHO_REQUEST = 71
    HTTP_QUERY_ETAG = 54
    HTTP_QUERY_EXPECT = 68
    HTTP_QUERY_EXPIRES = 10
    HTTP_QUERY_FORWARDED = 30
    HTTP_QUERY_FROM = 31
    HTTP_QUERY_HOST = 55
    HTTP_QUERY_IF_MATCH = 56
    HTTP_QUERY_IF_MODIFIED_SINCE = 32
    HTTP_QUERY_IF_NONE_MATCH = 57
    HTTP_QUERY_IF_RANGE = 58
    HTTP_QUERY_IF_UNMODIFIED_SINCE = 59
    HTTP_QUERY_LAST_MODIFIED = 11
    HTTP_QUERY_LINK = 16
    HTTP_QUERY_LOCATION = 33
    HTTP_QUERY_MAX = 78
    HTTP_QUERY_MAX_FORWARDS = 60
    HTTP_QUERY_MESSAGE_ID = 12
    HTTP_QUERY_MIME_VERSION = 0
    HTTP_QUERY_ORIG_URI = 34
    HTTP_QUERY_PRAGMA = 17
    HTTP_QUERY_PROXY_AUTHENTICATE = 41
    HTTP_QUERY_PROXY_AUTHORIZATION = 61
    HTTP_QUERY_PROXY_CONNECTION = 69
    HTTP_QUERY_PUBLIC = 8
    HTTP_QUERY_RANGE = 62
    HTTP_QUERY_RAW_HEADERS = 21
    HTTP_QUERY_RAW_HEADERS_CRLF = 22
    HTTP_QUERY_REFERER = 35
    HTTP_QUERY_REFRESH = 46
    HTTP_QUERY_REQUEST_METHOD = 45
    HTTP_QUERY_RETRY_AFTER = 36
    HTTP_QUERY_SERVER = 37
    HTTP_QUERY_SET_COOKIE = 43
    HTTP_QUERY_STATUS_CODE = 19
    HTTP_QUERY_STATUS_TEXT = 20
    HTTP_QUERY_TITLE = 38
    HTTP_QUERY_TRANSFER_ENCODING = 63
    HTTP_QUERY_UNLESS_MODIFIED_SINCE = 70
    HTTP_QUERY_UPGRADE = 64
    HTTP_QUERY_URI = 13
    HTTP_QUERY_USER_AGENT = 39
    HTTP_QUERY_VARY = 65
    HTTP_QUERY_VERSION = 18
    HTTP_QUERY_VIA = 66
    HTTP_QUERY_WARNING = 67
    HTTP_QUERY_WWW_AUTHENTICATE = 40

End Enum
 
Private Const CONNECTION_TIMEOUT_MILLIS As Long = 60000
 
Private m_sURLDomain                    As String
Private m_sURLPort                      As String
Private m_sURLPath                      As String
Private m_bURLSSL                       As Boolean
Private m_sPOSTData                     As String
Private m_hInetSession                  As Long
Private m_hInetConnect                  As Long
Private m_hInetRequest                  As Long
Private m_sHttpMethod                   As String
 
Public Property Let PostParameters(ByVal sData As String)
        
        On Error GoTo PostParameters_Err
        
100     m_sPOSTData = sData
        
        Exit Property

PostParameters_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.PostParameters", _
                  "clsInet component failure"
        
End Property
 
Public Property Get PostParameters() As String
        
        On Error GoTo PostParameters_Err
        
100     PostParameters = m_sPOSTData
        
        Exit Property

PostParameters_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.PostParameters", _
                  "clsInet component failure"
        
End Property
 
Public Function OpenRequest(ByVal sURL As String, Optional ByVal sHttpMethod As String = "POST", Optional ByVal sUserName As String = vbNullString, Optional ByVal sPassword As String = vbNullString, Optional ByRef sErrorMsg As Variant) As Boolean
        
        On Error GoTo OpenRequest_Err
        
 
100     m_sPOSTData = ""
102     m_sHttpMethod = sHttpMethod

104     If Not CrackURL(sURL) Then Exit Function
 
106     m_hInetSession = InternetOpen("Http Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

108     If m_hInetSession <= 0 Then
110         sErrorMsg = "InternetOpen() failed (System Error " & Err.LastDllError & ")"
112         GoTo EXIT_LABEL
        End If
 
114     m_hInetConnect = InternetConnect(m_hInetSession, m_sURLDomain, m_sURLPort, sUserName, sPassword, INTERNET_SERVICE_HTTP, 0, 0)

116     If m_hInetSession <= 0 Then
118         sErrorMsg = "InternetConnect() failed (System Error " & Err.LastDllError & ")"
120         GoTo EXIT_LABEL
        End If
 
        Dim lFlags As Long

122     lFlags = (INTERNET_FLAG_NO_COOKIES Or INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE)

124     If m_bURLSSL Then lFlags = (lFlags Or INTERNET_FLAG_SECURE)
 
126     m_hInetRequest = HttpOpenRequest(m_hInetConnect, sHttpMethod, m_sURLPath, "HTTP/1.0", vbNullString, vbNullString, lFlags, 0)

128     If m_hInetRequest <= 0 Then
130         sErrorMsg = "HttpOpenRequest() failed (System Error " & Err.LastDllError & ")"
132         GoTo EXIT_LABEL
        End If
    
134     If Not CBool(InternetSetOption(m_hInetRequest, INTERNET_OPTION_CONNECT_TIMEOUT, CONNECTION_TIMEOUT_MILLIS, 4)) Then
136         sErrorMsg = "InternetSetOption(INTERNET_OPTION_CONNECT_TIMEOUT) failed (System Error " & Err.LastDllError & ")"
138         GoTo EXIT_LABEL
        End If
    
140     If Not CBool(InternetSetOption(m_hInetRequest, INTERNET_OPTION_SEND_TIMEOUT, CONNECTION_TIMEOUT_MILLIS, 4)) Then
142         sErrorMsg = "InternetSetOption(INTERNET_OPTION_SEND_TIMEOUT) failed (System Error " & Err.LastDllError & ")"
144         GoTo EXIT_LABEL
        End If
    
146     If Not CBool(InternetSetOption(m_hInetRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, CONNECTION_TIMEOUT_MILLIS, 4)) Then
148         sErrorMsg = "InternetSetOption(INTERNET_OPTION_RECEIVE_TIMEOUT) failed (System Error " & Err.LastDllError & ")"
150         GoTo EXIT_LABEL
        End If
    
152     If sHttpMethod = "POST" Then
154         If Not AddHeader("Content-Type", "application/x-www-form-urlencoded") Then
156             sErrorMsg = "AddHeader('Content-Type') failed"
158             GoTo EXIT_LABEL
            End If
        End If
    
160     If m_bURLSSL Then

            Dim lBuffer As Long

162         Call InternetQueryOption(m_hInetRequest, INTERNET_OPTION_SECURITY_FLAGS, lFlags, lBuffer)
164         lFlags = lFlags Or SECURITY_FLAG_IGNORE_UNKNOWN_CA
166         Call InternetSetOption(m_hInetRequest, INTERNET_OPTION_SECURITY_FLAGS, lFlags, lBuffer)

        End If
    
168     OpenRequest = True
    
EXIT_LABEL:

        
        Exit Function

OpenRequest_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.OpenRequest", _
                  "clsInet component failure"
        
End Function
    
Public Function AddHeader(ByVal sName As String, ByVal sValue As String) As Boolean
        
        On Error GoTo AddHeader_Err
        

        Dim sHeaderFmt As String
    
100     sHeaderFmt = Trim$(sName) & ": " & Trim$(sValue) & vbCrLf
102     AddHeader = CBool(HttpAddRequestHeaders(m_hInetRequest, sHeaderFmt, Len(sHeaderFmt), HTTP_ADDREQ_FLAG_ADD Or HTTP_ADDREQ_FLAG_REPLACE))
    
        
        Exit Function

AddHeader_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.AddHeader", _
                  "clsInet component failure"
        
End Function
    
Public Function AddPOSTParameter(ByVal name As String, ByVal value As String) As Boolean
        
        On Error GoTo AddPOSTParameter_Err
        

        Dim sParam As String

100     name = Trim$(name)
102     value = Trim$(value)

104     If m_sHttpMethod = "POST" Then value = UrlEncode(value)

106     If Len(name) > 0 Then
108         sParam = name & "=" & value
        Else
110         sParam = value
        End If
 
112     If Len(m_sPOSTData) = 0 Then
114         m_sPOSTData = sParam
        Else
116         m_sPOSTData = m_sPOSTData & "&" & sParam
        End If
    
118     AddPOSTParameter = True
        
        Exit Function

AddPOSTParameter_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.AddPOSTParameter", _
                  "clsInet component failure"
        
End Function
 
Public Function Execute() As Boolean
        
        On Error GoTo Execute_Err
        

100     If Len(m_sPOSTData) > 0 Then
102         Call AddHeader("Content-Length", Len(m_sPOSTData))
104         Execute = CBool(HttpSendRequest(m_hInetRequest, vbNullString, 0, m_sPOSTData, Len(m_sPOSTData)))

        Else
106         Execute = CBool(HttpSendRequest(m_hInetRequest, vbNullString, 0, vbNullString, 0))

        End If

        
        Exit Function

Execute_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.Execute", _
                  "clsInet component failure"
        
End Function
 
Public Function GetResponseAsString() As String

    On Error GoTo errhandler

    Dim sBody      As String
    Dim sBuffer    As String
    Dim lBytesRead As Long

    Do
        sBuffer = Space$(4048)
        InternetReadFile m_hInetRequest, sBuffer, Len(sBuffer), lBytesRead

        If lBytesRead <= 0 Then Exit Do
        sBody = sBody & Left$(sBuffer, lBytesRead)
    Loop

    GetResponseAsString = sBody

    Exit Function

errhandler:

    GetResponseAsString = ""
End Function

Public Function GetResponseHeader(ByVal lHeaderType As enumPostHeader) As String

    On Error GoTo errhandler

    Dim sValue       As String
    Dim lBytesRead   As Long
    Dim lHeaderIndex As Long
    Dim bSuccess     As Boolean
    
    sValue = String$(2000, Chr$(0))
    lBytesRead = Len(sValue)
    bSuccess = CBool(HttpQueryInfo(m_hInetRequest, lHeaderType, ByVal sValue, lBytesRead, lHeaderIndex))

    If bSuccess And lBytesRead > 0 Then
        GetResponseHeader = Left$(sValue, InStr(1, sValue, Chr$(0)) - 1)
    End If

    Exit Function
 
errhandler:
End Function
 
Private Function CrackURL(ByVal strURL As String) As Boolean
        
        On Error GoTo CrackURL_Err
        

        Dim intPos      As Integer

        Dim strProtocol As String
 
100     m_sURLDomain = vbNullString
102     m_sURLPort = vbNullString
104     m_sURLPath = vbNullString
106     m_bURLSSL = False
 
        ' search for double forward slash and remove protocol prefix
108     intPos = InStr(1, strURL, "//")

110     If intPos > 0 Then
112         strProtocol = Left$(strURL, intPos - 1)
114         strURL = Mid$(strURL, intPos + 2)
        End If
 
116     m_bURLSSL = (InStr(LCase$(strProtocol), "https") > 0)
 
        ' split the URL into domain and path using the first forward slash found
        ' If no slash is found, then the url is the domain only
118     intPos = InStr(1, strURL, "/")

120     If intPos > 0 Then
122         m_sURLDomain = Left$(strURL, intPos - 1)
124         m_sURLPath = Mid$(strURL, intPos + 1)
        Else
126         m_sURLDomain = strURL
        End If
 
        ' Now see if a port is attached to the domain; split if found
128     intPos = InStr(1, m_sURLDomain, ":")

130     If intPos > 0 Then
132         m_sURLPort = Mid$(m_sURLDomain, intPos + 1)
134         m_sURLDomain = Left$(m_sURLDomain, intPos - 1)
        Else
136         m_sURLPort = IIf(m_bURLSSL, INTERNET_DEFAULT_HTTPS_PORT, INTERNET_DEFAULT_HTTP_PORT)
        End If

138     CrackURL = True
        
        Exit Function

CrackURL_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.CrackURL", _
                  "clsInet component failure"
        
End Function
 
Private Function TrimString(ByVal s As String) As String
        
        On Error GoTo TrimString_Err
        
100     TrimString = Left$(s, InStr(1, s, Chr$(0)) - 1)
        
        Exit Function

TrimString_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.TrimString", _
                  "clsInet component failure"
        
End Function
 
Private Function UrlEncode(sText As String) As String
        
        On Error GoTo UrlEncode_Err
        

        Dim sResult As String

        Dim sFinal  As String

        Dim sChar   As String

        Dim i       As Long
 
100     sResult = ""
102     sFinal = ""

104     For i = 1 To Len(sText)
106         sChar = Mid$(sText, i, 1)

108         If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", sChar) <> 0 Then
110             sResult = sResult & sChar
112         ElseIf sChar = " " Then
114             sResult = sResult & "+"
116         ElseIf sChar = vbCr Or sChar = vbLf Or sChar = vbCrLf Then
            Else
118             sResult = sResult & "%" & Hex$(Asc(sChar))
            End If

120         If Len(sResult) >= 1000 Then
122             sFinal = sFinal & sResult
124             sResult = ""
            End If

        Next

126     UrlEncode = sFinal & sResult
        
        Exit Function

UrlEncode_Err:
        Err.Raise vbObjectError + 100, _
                  "WinterAOLauncher.clsInet.UrlEncode", _
                  "clsInet component failure"
        
End Function
 
Public Sub CloseConnection()

    On Error GoTo ERROR_HANDLER

    If m_hInetRequest > 0 Then Call InternetCloseHandle(m_hInetRequest)
    If m_hInetConnect > 0 Then Call InternetCloseHandle(m_hInetConnect)
    If m_hInetSession > 0 Then Call InternetCloseHandle(m_hInetSession)
    m_hInetRequest = 0
    m_hInetConnect = 0
    m_hInetSession = 0
    m_sPOSTData = ""
    m_sURLDomain = ""
    m_sURLPort = ""
    m_sURLPath = ""
    m_bURLSSL = False
ERROR_HANDLER:
End Sub
 
Private Sub Class_Terminate()
    
    On Error Resume Next
    
    Call CloseConnection
End Sub

