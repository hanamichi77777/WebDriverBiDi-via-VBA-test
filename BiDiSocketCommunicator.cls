VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BiDiSocketCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ========================================================================================
' Class Name : BiDiSocketCommunicator
' Purpose    : Professional WebSocket Client for WebDriver BiDi.
' Features   :
'   - Adaptive Polling & Autonomous Heartbeat (No burst-mode conflict).
'   - Strict Deadline-based Timeout Management (No timeout leakage).
'   - Peak-holding Buffer Growth (Optimized for large JSON payloads).
'   - Type-safe operations and standardized constant definitions.
' ========================================================================================

Private Const DEBUG_MODE As Boolean = False

' --- WinHttp API Declarations ---
Private Declare PtrSafe Function WinHttpOpen Lib "WinHttp" (ByVal pszAgentW As LongPtr, ByVal dwAccessType As Long, ByVal pszProxyW As LongPtr, ByVal pszProxyBypassW As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpConnect Lib "WinHttp" (ByVal hSession As LongPtr, ByVal pswzServerName As LongPtr, ByVal nServerPort As Long, ByVal dwReserved As Long) As LongPtr
Private Declare PtrSafe Function WinHttpOpenRequest Lib "WinHttp" (ByVal hConnect As LongPtr, ByVal pwszVerb As LongPtr, ByVal pwszObjectName As LongPtr, ByVal pwszVersion As LongPtr, ByVal pwszReferrer As LongPtr, ByVal ppwszAcceptTypes As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpSendRequest Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As Long) As Long
Private Declare PtrSafe Function WinHttpReceiveResponse Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal lpReserved As LongPtr) As Long
Private Declare PtrSafe Function WinHttpQueryHeaders Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, ByVal pwszName As LongPtr, ByRef lpBuffer As Any, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long
Private Declare PtrSafe Function WinHttpSetOption Lib "WinHttp" (ByVal hInternet As LongPtr, ByVal dwOption As Long, ByVal lpBuffer As LongPtr, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpSetTimeouts Lib "WinHttp" (ByVal hInternet As LongPtr, ByVal nResolveTimeout As Long, ByVal nConnectTimeout As Long, ByVal nSendTimeout As Long, ByVal nReceiveTimeout As Long) As Long
Private Declare PtrSafe Function WinHttpCloseHandle Lib "WinHttp" (ByVal hInternet As LongPtr) As Long

' --- WebSocket API Declarations ---
Private Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal pContext As LongPtr) As LongPtr
Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal eBufferType As Long, ByVal pvBuffer As LongPtr, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByRef pvBuffer As Any, ByVal dwBufferLength As Long, ByRef pdwBytesRead As Long, ByRef peBufferType As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketClose Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal usStatus As Integer, ByVal pvReason As LongPtr, ByVal dwReasonLength As Long) As Long

' --- Utility API Declarations ---
Private Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- Standard WinHttp Constants ---
Private Const WINHTTP_ACCESS_TYPE_NO_PROXY As Long = 1
Private Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET As Long = 114
Private Const WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL As Long = 150
Private Const WINHTTP_QUERY_STATUS_CODE As Long = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER As Long = &H20000000
Private Const WINHTTP_HEADER_NAME_BY_INDEX As Long = 0
Private Const WINHTTP_NO_HEADER_INDEX As Long = 0
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE As Long = 2
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE As Long = 4

' --- Custom Library Constants ---
Private Const NO_ERROR As Long = 0
Private Const WINHTTP_ERROR_CONNECTION_RESET As Long = 12030
Private Const AGENT_NAME As String = "SeleniumVBA_BiDi_Client"
Private Const SYSTEM_PING_ID As Long = -999

' --- Configuration Defaults ---
Private Const DEFAULT_BUFFER_SIZE As Long = 8192
Private Const MAX_BUFFER_SIZE As Long = 65536
Private Const DEFAULT_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_POLLING_INTERVAL_MS As Long = 1000
Private Const KEEPALIVE_INTERVAL_MS As Long = 15000

' --- Member Variables ---
Private hSession_ As LongPtr
Private hConnection_ As LongPtr
Private hRequest_ As LongPtr
Private hWebSocket_ As LongPtr

Private m_EventBuffer As Collection
Private m_DefaultTimeout As Long
Private m_PollingInterval As Long

Private m_BufferSize As Long
Private m_MaxBufferSize As Long
Private m_GrowThresholdCount As Long
Private m_RecentFullReads As Long

Private p_lastActivityTime As Double
Private m_Stream As ADODB.Stream
Private p_RegExp As RegExp
Private m_AllowUiPump As Boolean

' ========================================================================================
' Initialization & Termination
' ========================================================================================
Private Sub Class_Initialize()
    m_DefaultTimeout = DEFAULT_TIMEOUT_MS
    m_PollingInterval = DEFAULT_POLLING_INTERVAL_MS
    p_lastActivityTime = 0
    
    m_BufferSize = DEFAULT_BUFFER_SIZE
    m_MaxBufferSize = MAX_BUFFER_SIZE
    m_GrowThresholdCount = 1
    m_RecentFullReads = 0

    Set m_EventBuffer = New Collection
    Set m_Stream = New ADODB.Stream
    m_Stream.Mode = adModeReadWrite
    
    Set p_RegExp = New RegExp
    With p_RegExp
        .Global = False
        .IgnoreCase = True
        .Multiline = False
        .Pattern = """id""\s*:\s*(-?\d+)"
    End With

    m_AllowUiPump = False
End Sub

Private Sub Class_Terminate()
    Disconnect
    If Not m_Stream Is Nothing Then
        If m_Stream.State = adStateOpen Then m_Stream.Close
        Set m_Stream = Nothing
    End If
    Set p_RegExp = Nothing
End Sub

' ========================================================================================
' Public Properties
' ========================================================================================
Public Property Get IsConnected() As Boolean: IsConnected = (hWebSocket_ <> 0): End Property
Public Property Let PollingInterval(ByVal v As Long): m_PollingInterval = v: End Property
Public Property Let AllowUiPump(ByVal v As Boolean): m_AllowUiPump = v: End Property
Public Property Get AllowUiPump() As Boolean: AllowUiPump = m_AllowUiPump: End Property

' ========================================================================================
' Connection Logic
' ========================================================================================
Public Function Connect(ByVal wsUrl As String, ByVal localPort As Long) As Boolean
    On Error GoTo ErrorHandler
    Disconnect
    If wsUrl = "" Or localPort = 0 Then Exit Function
    
    Dim Path As String: Path = ExtractPath(wsUrl)
    Dim host As String: host = "localhost"

    hSession_ = WinHttpOpen(StrPtr(AGENT_NAME), WINHTTP_ACCESS_TYPE_NO_PROXY, 0, 0, 0)
    If hSession_ = 0 Then Err.Raise 513, "BiDiSocket", "WinHttpOpen failed"

    ' Initial API layer timeouts
    WinHttpSetTimeouts hSession_, 2000, 2000, 10000, 60000

    hConnection_ = WinHttpConnect(hSession_, StrPtr(host), localPort, 0)
    If hConnection_ = 0 Then Err.Raise 514, "BiDiSocket", "WinHttpConnect failed"

    hRequest_ = WinHttpOpenRequest(hConnection_, StrPtr("GET"), StrPtr(Path), StrPtr("HTTP/1.1"), 0, 0, 0)
    If hRequest_ = 0 Then Err.Raise 515, "BiDiSocket", "WinHttpOpenRequest failed"

    WinHttpSetOption hRequest_, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0

    If WinHttpSendRequest(hRequest_, 0, 0, 0, 0, 0, 0) = 0 Then Err.Raise 517, "BiDiSocket", "SendRequest failed"
    If WinHttpReceiveResponse(hRequest_, 0) = 0 Then Err.Raise 518, "BiDiSocket", "ReceiveResponse failed"

    Dim status As Long, Size As Long: Size = 4
    WinHttpQueryHeaders hRequest_, WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, status, Size, WINHTTP_NO_HEADER_INDEX
    If status <> 101 Then Err.Raise 519, "BiDiSocket", "Upgrade rejected: " & status

    hWebSocket_ = WinHttpWebSocketCompleteUpgrade(hRequest_, 0)
    If hWebSocket_ = 0 Then Err.Raise 520, "BiDiSocket", "Upgrade failed"

    Dim ka As Long: ka = KEEPALIVE_INTERVAL_MS
    WinHttpSetOption hWebSocket_, WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL, VarPtr(ka), 4
    
    p_lastActivityTime = GetTickMS()
    Connect = True
    Exit Function

ErrorHandler:
    Me.Disconnect
    Connect = False
End Function

Public Sub Disconnect()
    On Error Resume Next
    If hWebSocket_ <> 0 Then WinHttpWebSocketClose hWebSocket_, 1000, 0, 0: WinHttpCloseHandle hWebSocket_: hWebSocket_ = 0
    If hRequest_ <> 0 Then WinHttpCloseHandle hRequest_: hRequest_ = 0
    If hConnection_ <> 0 Then WinHttpCloseHandle hConnection_: hConnection_ = 0
    If hSession_ <> 0 Then WinHttpCloseHandle hSession_: hSession_ = 0
    On Error GoTo 0
    Set m_EventBuffer = New Collection
End Sub

' ========================================================================================
' Communication Logic (High Level)
' ========================================================================================

Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    If Not IsConnected Then Err.Raise 600, "BiDiSocket", "Not connected"
    
    Dim cmdId As String: cmdId = ExtractJsonId(jsonMessage)
    If Not SendMessage(jsonMessage) Then Err.Raise 601, "BiDiSocket", "Send failed"
    
    Dim timeout As Long: timeout = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    Dim deadline As Double: deadline = GetTickMS() + timeout
    
    Dim msg As String, inId As String
    
    Do
        msg = ReadFromWire(deadline)
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            If inId = cmdId Then
                SendAndReceive = msg
                Exit Function
            Else
                m_EventBuffer.Add msg
            End If
        End If
        
        If GetTickMS() > deadline Then Err.Raise 604, "BiDiSocket", "Timeout waiting for ID: " & cmdId

        YieldExecution 2, m_AllowUiPump
    Loop
End Function

Public Function GetMessage() As String
    ' 1. Check Event Buffer first (Priority for burst data)
    If m_EventBuffer.Count > 0 Then
        GetMessage = m_EventBuffer(1)
        m_EventBuffer.Remove 1
        Exit Function
    End If
    
    If Not IsConnected Then Exit Function
    
    ' 2. Idle Monitoring (Corrected: No p_lastActivityTime = 0)
    Dim nowT As Double: nowT = GetTickMS()
    If (nowT - p_lastActivityTime) < m_PollingInterval Then
        Exit Function
    End If
    
    ' 3. Active Polling (Heartbeat Ping)
    Dim pingCmd As String
    pingCmd = "{""id"":" & SYSTEM_PING_ID & ", ""method"":""session.status"", ""params"":{}}"
    
    If Not SendMessage(pingCmd) Then Exit Function
    
    ' 4. Read Response with strict deadline
    Dim rawMsg As String: rawMsg = ReadFromWire(GetTickMS() + 500)
    
    If rawMsg = "" Then
        YieldExecution 1, m_AllowUiPump
        Exit Function
    End If
    
    ' 5. Dispatching
    Dim inId As String: inId = ExtractJsonId(rawMsg)
    
    If inId = CStr(SYSTEM_PING_ID) Then
        GetMessage = "" ' Swallow heartbeat response
    Else
        GetMessage = rawMsg ' Real event arrived during ping wait
    End If
End Function

Public Function SendMessage(ByVal msg As String) As Boolean
    If hWebSocket_ = 0 Then Exit Function
    Dim data() As Byte: data = ToUtf8Bytes(msg)
    
    If WinHttpWebSocketSend(hWebSocket_, WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, VarPtr(data(0)), UBound(data) + 1) = NO_ERROR Then
        SendMessage = True
        p_lastActivityTime = GetTickMS()
    End If
End Function

' ========================================================================================
' Low Level Wire Reading
' ========================================================================================
Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    If hWebSocket_ = 0 Then Exit Function

    Dim buffer() As Byte: ReDim buffer(m_BufferSize - 1)
    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim hasData As Boolean: hasData = False

    If m_Stream.State = adStateOpen Then m_Stream.Close
    m_Stream.Type = adTypeBinary
    m_Stream.Open

    Do
        ' Deadline check to prevent hang on incomplete frames
        If deadline > 0 Then
            If GetTickMS() > deadline Then Exit Do
        End If

        ret = WinHttpWebSocketReceive(hWebSocket_, buffer(0), UBound(buffer) + 1, bytesRead, bufType)
        
        If ret <> NO_ERROR Then
            If ret = WINHTTP_ERROR_CONNECTION_RESET Then Disconnect
            Exit Function
        End If

        If bufType = WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE Then
            Disconnect
            Exit Function
        End If

        If bytesRead > 0 Then
            hasData = True
            If bytesRead < (UBound(buffer) + 1) Then
                Dim temp() As Byte: ReDim temp(bytesRead - 1)
                Dim i As Long: For i = 0 To bytesRead - 1: temp(i) = buffer(i): Next
                m_Stream.Write temp
            Else
                m_Stream.Write buffer
            End If
            MaybeGrowBuffer bytesRead
        End If

        If bufType = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then Exit Do
        
        YieldExecution 0, False
    Loop

    If hasData Then
        m_Stream.Position = 0
        m_Stream.Type = adTypeText: m_Stream.Charset = "UTF-8"
        ReadFromWire = m_Stream.ReadText
        p_lastActivityTime = GetTickMS()
    End If
End Function

' ========================================================================================
' Helpers & Utils
' ========================================================================================
Private Function GetTickMS() As Double
    Static freq As Currency
    Dim Count As Currency
    If freq = 0 Then QueryPerformanceFrequency freq
    QueryPerformanceCounter Count
    GetTickMS = (Count / freq) * 1000
End Function

' Peak-holding buffer growth logic
Private Sub MaybeGrowBuffer(ByVal bytesRead As Long)
    If bytesRead >= m_BufferSize Then
        m_RecentFullReads = m_RecentFullReads + 1
        If m_RecentFullReads >= m_GrowThresholdCount Then
            Dim newSize As Long: newSize = m_BufferSize * 2
            If newSize > m_MaxBufferSize Then newSize = m_MaxBufferSize
            If newSize > m_BufferSize Then
                m_BufferSize = newSize
                m_RecentFullReads = 0
                If DEBUG_MODE Then Debug.Print "BiDi: Buffer grown to " & m_BufferSize
            End If
        End If
    End If
End Sub

' Standardized port retrieval with type-safety
Public Function AttemptAutoConnect(ByVal wsUrl As String) As Boolean
    Dim port As Long
    port = GetPortFromUrl(wsUrl)
    If wsUrl <> "" And port <> 0 Then
        AttemptAutoConnect = Connect(wsUrl, port)
    End If
End Function

Public Function GetPortFromUrl(ByVal wsUrl As String) As Long
    Dim posSession As Long, posColon As Long
    Dim portStr As String
    
    posSession = InStr(wsUrl, "/session")
    If posSession = 0 Then Exit Function
    
    posColon = InStrRev(wsUrl, ":", posSession)
    If posColon = 0 Then Exit Function
    
    portStr = Mid(wsUrl, posColon + 1, posSession - posColon - 1)
    If IsNumeric(portStr) Then
        GetPortFromUrl = CLng(portStr)
    End If
End Function

Private Function ExtractJsonId(ByVal json As String) As String
    If Len(json) < 5 Then Exit Function
    Dim matches As MatchCollection
    Set matches = p_RegExp.execute(json)
    If matches.Count > 0 Then ExtractJsonId = matches(0).SubMatches(0)
End Function

Private Function ExtractPath(ByVal url As String) As String
    Dim p As Long: p = InStr(1, url, "://"): If p > 0 Then p = p + 3 Else p = 1
    p = InStr(p, url, "/"): ExtractPath = IIf(p > 0, Mid$(url, p), "/")
End Function

Private Function ToUtf8Bytes(ByVal txt As String) As Byte()
    With m_Stream
        If .State = adStateOpen Then .Close
        .Type = adTypeText: .Charset = "UTF-8": .Open
        .WriteText txt
        .Position = 0: .Type = adTypeBinary: .Position = 3 ' Skip BOM
        If .Size > 3 Then ToUtf8Bytes = .Read Else ToUtf8Bytes = ""
    End With
End Function

Private Sub YieldExecution(Optional ByVal ms As Long = 1, Optional ByVal allowDoEvents As Boolean = False)
    If allowDoEvents Then DoEvents
    If ms >= 0 Then SleepWinAPI ms
End Sub
