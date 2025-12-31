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
' Class Name: BiDiSocketCommunicator
' Purpose   : WebSocket Client with "Adaptive Polling" & Autonomous Heartbeat.
'             - Idle: Sends "session.status" (ID: -999) every 1s to keep connection alive.
'             - Burst: Switches to immediate processing when data is received.
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
' [Updated] Using QPC to avoid the 49.7-day limit of GetTickCount
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- Constants ---
Private Const WINHTTP_ACCESS_TYPE_NO_PROXY As Long = 1
Private Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET As Long = 114
Private Const WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL As Long = 150
Private Const WINHTTP_QUERY_STATUS_CODE As Long = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER As Long = &H20000000
Private Const WINHTTP_HEADER_NAME_BY_INDEX As Long = 0
Private Const WINHTTP_NO_HEADER_INDEX As Long = 0

Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE As Long = 2
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE As Long = 4
Private Const NO_ERROR As Long = 0
Private Const AGENT_NAME As String = "SeleniumVBA_BiDi_Client"

' --- System Reserved ID for Heartbeat ---
Private Const SYSTEM_PING_ID As Long = -999

' --- Member Variables ---
Private hSession_ As LongPtr
Private hConnection_ As LongPtr
Private hRequest_ As LongPtr
Private hWebSocket_ As LongPtr

Private m_EventBuffer As Collection
Private m_DefaultTimeout As Long
Private m_PollingInterval As Long

' Buffer Management
Private m_BufferSize As Long
Private m_MaxBufferSize As Long
Private m_GrowThresholdCount As Long
Private m_RecentFullReads As Long

' Optimization: Ping Throttling
' [Updated] Changed to Double to support millisecond precision from QPC
Private p_lastActivityTime As Double

' Optimization: Reusable Objects
Private m_Stream As ADODB.Stream
Private p_RegExp As RegExp

' UI pump control
Private m_AllowUiPump As Boolean

' ========================================================================================
' Initialization & Termination
' ========================================================================================
Private Sub Class_Initialize()
    m_DefaultTimeout = 30000
    m_PollingInterval = 1000 ' 1 second idle interval
    p_lastActivityTime = 0
    
    ' Buffer settings
    m_BufferSize = 8192
    m_MaxBufferSize = 65536
    m_GrowThresholdCount = 1
    m_RecentFullReads = 0

    Set m_EventBuffer = New Collection
    
    ' Initialize ADODB.Stream (Binary/Text Converter)
    Set m_Stream = New ADODB.Stream
    m_Stream.Mode = adModeReadWrite
    
    ' Initialize RegExp (Early Binding)
    Set p_RegExp = New RegExp
    With p_RegExp
        .Global = False
        .IgnoreCase = True
        .Multiline = False
        ' Allows negative integers (e.g. -999) for System Pings.
        .Pattern = """id""\s*:\s*(-?\d+)"
    End With

    ' Default: do not pump UI inside communication loops
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
Public Property Get IsConnected() As Boolean
    IsConnected = (hWebSocket_ <> 0)
End Property

Public Property Let PollingInterval(ByVal v As Long)
    m_PollingInterval = v
End Property

Public Property Let AllowUiPump(ByVal v As Boolean)
    m_AllowUiPump = v
End Property

Public Property Get AllowUiPump() As Boolean
    AllowUiPump = m_AllowUiPump
End Property

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

    ' Set generic timeouts
    WinHttpSetTimeouts hSession_, 2000, 2000, 10000, 60000

    hConnection_ = WinHttpConnect(hSession_, StrPtr(host), localPort, 0)
    If hConnection_ = 0 Then Err.Raise 514, "BiDiSocket", "WinHttpConnect failed"

    hRequest_ = WinHttpOpenRequest(hConnection_, StrPtr("GET"), StrPtr(Path), StrPtr("HTTP/1.1"), 0, 0, 0)
    If hRequest_ = 0 Then Err.Raise 515, "BiDiSocket", "WinHttpOpenRequest failed"

    ' Request upgrade to WebSocket
    WinHttpSetOption hRequest_, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0

    If WinHttpSendRequest(hRequest_, 0, 0, 0, 0, 0, 0) = 0 Then Err.Raise 517, "BiDiSocket", "SendRequest failed"
    If WinHttpReceiveResponse(hRequest_, 0) = 0 Then Err.Raise 518, "BiDiSocket", "ReceiveResponse failed"

    Dim status As Long, Size As Long: Size = 4
    WinHttpQueryHeaders hRequest_, WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, status, Size, WINHTTP_NO_HEADER_INDEX
    If status <> 101 Then Err.Raise 519, "BiDiSocket", "Upgrade rejected: " & status

    hWebSocket_ = WinHttpWebSocketCompleteUpgrade(hRequest_, 0)
    If hWebSocket_ = 0 Then Err.Raise 520, "BiDiSocket", "Upgrade failed"

    ' Set TCP KeepAlive (distinct from our Application-level Ping)
    Dim ka As Long: ka = 15000
    WinHttpSetOption hWebSocket_, WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL, VarPtr(ka), 4
    
    p_lastActivityTime = GetTickMS()
    Connect = True
    Exit Function

ErrorHandler:
    Me.Disconnect
    Connect = False
End Function

Public Sub Disconnect()
    ' 1. Cleanup API Handles (Ignore errors here as handles might be already closed)
    On Error Resume Next
    If hWebSocket_ <> 0 Then WinHttpWebSocketClose hWebSocket_, 1000, 0, 0: WinHttpCloseHandle hWebSocket_: hWebSocket_ = 0
    If hRequest_ <> 0 Then WinHttpCloseHandle hRequest_: hRequest_ = 0
    If hConnection_ <> 0 Then WinHttpCloseHandle hConnection_: hConnection_ = 0
    If hSession_ <> 0 Then WinHttpCloseHandle hSession_: hSession_ = 0
    
    ' 2. Restore Error Handling (Critical fix: Do not hide internal errors)
    On Error GoTo 0
    
    ' 3. Reset internal state
    Set m_EventBuffer = New Collection
End Sub

' ========================================================================================
' Communication Logic (High Level)
' ========================================================================================

' Synchronous Request-Response (Blocking)
Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    If Not IsConnected Then Err.Raise 600, "BiDiSocket", "Not connected"
    
    Dim cmdId As String: cmdId = ExtractJsonId(jsonMessage)
    If Not SendMessage(jsonMessage) Then Err.Raise 601, "BiDiSocket", "Send failed"
    
    Dim timeout As Long: timeout = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    ' [Updated] Use Double for start time
    Dim startT As Double: startT = GetTickMS()
    Dim msg As String, inId As String
    
    Do
        msg = ReadFromWire()
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            If inId = cmdId Then
                ' Match found! Return result.
                SendAndReceive = msg
                Exit Function
            Else
                ' Unrelated event/message. Buffer it.
                m_EventBuffer.Add msg
            End If
        End If
        
        If (GetTickMS() - startT) > timeout Then Err.Raise 604, "BiDiSocket", "Timeout waiting for ID: " & cmdId

        YieldExecution 2, m_AllowUiPump
    Loop
End Function

' Asynchronous Message Retrieval (Adaptive Polling)
' Used by the loop in the main Wrapper.
Public Function GetMessage() As String
    ' 1. Check Event Buffer first (Return immediately if data exists)
    If m_EventBuffer.Count > 0 Then
        GetMessage = m_EventBuffer(1)
        m_EventBuffer.Remove 1
        Exit Function
    End If
    
    If Not IsConnected Then Exit Function
    
    ' [Updated] Use Double for current time
    Dim nowT As Double: nowT = GetTickMS()
    
    ' 2. Adaptive Throttling Logic
    ' If no activity recently, wait until the interval passes before pinging again.
    If (nowT - p_lastActivityTime) < m_PollingInterval Then
        Exit Function
    End If
    
    ' 3. Active Polling (Heartbeat Ping)
    ' Using ID: -999 to guarantee no collision with user commands (positive integers).
    Dim pingCmd As String
    pingCmd = "{""id"":" & SYSTEM_PING_ID & ", ""method"":""session.status"", ""params"":{}}"
    
    If Not SendMessage(pingCmd) Then Exit Function
    
    ' Attempt to read response immediately
    Dim rawMsg As String: rawMsg = ReadFromWire()
    
    If rawMsg = "" Then
        ' No data? Yield CPU and return nothing.
        YieldExecution 1, m_AllowUiPump
        Exit Function
    End If
    
    Dim inId As String: inId = ExtractJsonId(rawMsg)
    
    If inId = CStr(SYSTEM_PING_ID) Then
        ' CASE: Heartbeat Response
        ' Swallow it. It served its purpose (keeping the connection open).
        GetMessage = ""
    Else
        ' CASE: Real Event or Async Message
        ' Switch to "Burst Mode":
        ' [Fix] Set timestamp to 0 to force the throttling condition
        ' (nowT - lastActivity < interval) to result in FALSE next time.
        p_lastActivityTime = 0
        
        GetMessage = rawMsg
    End If
End Function

Public Function SendMessage(ByVal msg As String) As Boolean
    If hWebSocket_ = 0 Then Exit Function
    Dim data() As Byte: data = ToUtf8Bytes(msg)
    
    If WinHttpWebSocketSend(hWebSocket_, WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, VarPtr(data(0)), UBound(data) + 1) = NO_ERROR Then
        SendMessage = True
        ' Update activity timestamp to prevent immediate re-pinging
        ' [Updated] Use GetTickMS()
        p_lastActivityTime = GetTickMS()
    End If
End Function

' ========================================================================================
' Low Level Wire Reading
' ========================================================================================
Private Function ReadFromWire() As String
    If hWebSocket_ = 0 Then Exit Function

    Dim buffer() As Byte: ReDim buffer(m_BufferSize - 1)
    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim hasData As Boolean: hasData = False

    If m_Stream.State = adStateOpen Then m_Stream.Close
    m_Stream.Type = adTypeBinary
    m_Stream.Open

    Do
        ret = WinHttpWebSocketReceive(hWebSocket_, buffer(0), UBound(buffer) + 1, bytesRead, bufType)
        
        If ret <> NO_ERROR Then
            ' 12030 = Connection terminated usually
            If ret = 12030 Then Disconnect
            Exit Function
        End If

        If bufType = WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE Then
            Disconnect
            Exit Function
        End If

        If bytesRead > 0 Then
            hasData = True
            ' Optimization: Write only actual bytes read
            If bytesRead < (UBound(buffer) + 1) Then
                Dim temp() As Byte: ReDim temp(bytesRead - 1)
                Dim i As Long: For i = 0 To bytesRead - 1: temp(i) = buffer(i): Next
                m_Stream.Write temp
            Else
                m_Stream.Write buffer
            End If
            MaybeGrowBuffer bytesRead
        End If

        ' Break if end of message frame
        If bufType = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then Exit Do
        YieldExecution 0, False
    Loop

    If hasData Then
        m_Stream.Position = 0
        m_Stream.Type = adTypeText: m_Stream.Charset = "UTF-8"
        ReadFromWire = m_Stream.ReadText
        
        ' [Fix] Update activity timestamp even on reads.
        ' If data is flowing in from server, connection is alive, no need to ping.
        ' [Updated] Use GetTickMS()
        p_lastActivityTime = GetTickMS()
    End If
End Function

' ========================================================================================
' Helpers & Utils
' ========================================================================================
' Prevents the 49-day overflow issue found in GetTickCount.
Private Function GetTickMS() As Double
    Static freq As Currency
    Dim Count As Currency
    If freq = 0 Then QueryPerformanceFrequency freq
    QueryPerformanceCounter Count
    GetTickMS = (Count / freq) * 1000
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

' Dynamic Buffer resizing to handle large payloads (e.g. huge page source)
Private Sub MaybeGrowBuffer(ByVal bytesRead As Long)
    If bytesRead >= m_BufferSize Then
        ' Full read: Increment counter
        m_RecentFullReads = m_RecentFullReads + 1
        
        If m_RecentFullReads >= m_GrowThresholdCount Then
            Dim newSize As Long: newSize = m_BufferSize * 2
            If newSize > m_MaxBufferSize Then newSize = m_MaxBufferSize
            
            ' Only apply if actually changing
            If newSize > m_BufferSize Then
                m_BufferSize = newSize
                m_RecentFullReads = 0 ' Reset counter
                If DEBUG_MODE Then Debug.Print "BiDi: Buffer grown to " & m_BufferSize
            End If
        End If
    Else
        ' Partial read: Decay the counter instead of hard reset
        ' This prevents a single fragmented packet from cancelling out previous full reads.
        If m_RecentFullReads > 0 Then
            m_RecentFullReads = m_RecentFullReads - 1
        End If
    End If
End Sub

' Attempt Connect
Public Function AttemptAutoConnect(ByVal wsUrl As String) As Boolean
    Dim port As Long
    port = GetPortFromUrl(wsUrl)

    If wsUrl <> "" And port <> 0 Then AttemptAutoConnect = Connect(wsUrl, port)
End Function

' Extracts the port number from a URL string using the Split function
Public Function GetPortFromUrl(ByVal wsUrl As String) As String
    Dim posSession As Long
    Dim posColon As Long
    
    ' Locate "/session"
    posSession = InStr(wsUrl, "/session")
    If posSession = 0 Then Exit Function
    
    ' Locate the colon before "/session"
    posColon = InStrRev(wsUrl, ":", posSession)
    If posColon = 0 Then Exit Function
    
    ' Extract the substring
    GetPortFromUrl = Mid(wsUrl, posColon + 1, posSession - posColon - 1)
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
