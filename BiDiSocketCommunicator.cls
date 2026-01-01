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

' --- Reconnection Configuration ---
Private m_LastWsUrl As String
Private m_LastPort As Long
Private m_RetryCount As Integer
Private m_MaxRetryAttempts As Integer
Private m_AutoReconnect As Boolean

' --- State Machine Definitions ---
Public Enum BiDiSocketState
    StateDisconnected = 0
    StateConnecting = 1
    StateConnected = 2
    StateDisconnecting = 3
    StateError = 4
End Enum
Private m_State As BiDiSocketState

' Event for UI or Controller to monitor state changes
Public Event StateChanged(ByVal NewState As BiDiSocketState, ByVal stateName As String)

Private m_EventBuffer As Collection
Private m_DefaultTimeout As Long
Private m_PollingInterval As Long

Private m_BufferSize As Long
Private m_MaxBufferSize As Long
Private m_ReceiveBuffer() As Byte
Private m_GrowThresholdCount As Long
Private m_RecentFullReads As Long

Private p_lastActivityTime As Double
Private p_lastPollTime As Double
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
    ReDim m_ReceiveBuffer(m_BufferSize - 1)
    
    m_MaxBufferSize = MAX_BUFFER_SIZE
    m_GrowThresholdCount = 1
    m_RecentFullReads = 0

    Set m_EventBuffer = New Collection
    Set m_Stream = New ADODB.Stream
    m_Stream.Mode = adModeReadWrite
    m_Stream.Open ' Pre-open to avoid allocation overhead during first transmission
    
    Set p_RegExp = New RegExp
    With p_RegExp
        .Global = False
        .IgnoreCase = True
        .Multiline = False
        .Pattern = "(?:\{|,)\s*""id""\s*:\s*(-?\d+)\s*(?=[,}])"
    End With
    
    p_lastPollTime = GetTickMS()
    m_AllowUiPump = False
    
    m_MaxRetryAttempts = 3
    m_AutoReconnect = True
    m_RetryCount = 0
    SetState StateDisconnected
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
Public Property Get State() As BiDiSocketState: State = m_State: End Property
Public Property Let AutoReconnect(ByVal v As Boolean): m_AutoReconnect = v: End Property
Public Property Get AutoReconnect() As Boolean: AutoReconnect = m_AutoReconnect: End Property

Private Sub SetState(ByVal NewState As BiDiSocketState)
    If m_State = NewState Then Exit Sub
    
    m_State = NewState
    
    Dim stateName As String
    Select Case m_State
        Case StateDisconnected: stateName = "DISCONNECTED"
        Case StateConnecting:   stateName = "CONNECTING"
        Case StateConnected:    stateName = "CONNECTED"
        Case StateDisconnecting: stateName = "DISCONNECTING"
        Case StateError:         stateName = "ERROR"
    End Select
    
    If DEBUG_MODE Then Debug.Print "BiDi State Changed: " & stateName
    RaiseEvent StateChanged(m_State, stateName)
End Sub

' ========================================================================================
' Establishes a WebSocket connection to the WebDriver BiDi endpoint.
' Utilizes a state machine to prevent redundant connection attempts and manages lifecycle.
' ========================================================================================
Public Function Connect(ByVal wsUrl As String, ByVal localPort As Long) As Boolean
    ' Guard: Prevent connection attempts if already connected or in progress
    If m_State = StateConnecting Or m_State = StateConnected Then
        If DEBUG_MODE Then Debug.Print "BiDi: Connect aborted. Current state is " & m_State
        Exit Function
    End If

    ' Store connection parameters for potential recovery attempts
    m_LastWsUrl = wsUrl
    m_LastPort = localPort

    On Error GoTo ErrorHandler

    ' Step 1: Initialize transition to Connecting state
    SetState StateConnecting
    
    ' Ensure any residual handles are cleaned up before a new attempt
    ' This is internal cleanup and doesn't trigger a state-change event loop
    Call InternalCleanup
    
    ' Basic parameter validation
    If wsUrl = "" Or localPort = 0 Then
        If DEBUG_MODE Then Debug.Print "BiDi: Connect failed. Invalid URL or Port."
        GoTo ErrorHandler
    End If
    
    Dim Path As String: Path = ExtractPath(wsUrl)
    Dim host As String: host = "localhost"

    ' Step 2: Initialize WinHttp Session
    hSession_ = WinHttpOpen(StrPtr(AGENT_NAME), WINHTTP_ACCESS_TYPE_NO_PROXY, 0, 0, 0)
    If hSession_ = 0 Then Err.Raise 513, "BiDiSocket", "WinHttpOpen failed"

    ' Set connection-related timeouts (Resolve, Connect, Send, Receive)
    WinHttpSetTimeouts hSession_, 2000, 2000, 5000, 5000

    ' Step 3: Connect to the WebDriver host
    hConnection_ = WinHttpConnect(hSession_, StrPtr(host), localPort, 0)
    If hConnection_ = 0 Then Err.Raise 514, "BiDiSocket", "WinHttpConnect failed"

    ' Step 4: Open an HTTP request to initiate the WebSocket handshake
    hRequest_ = WinHttpOpenRequest(hConnection_, StrPtr("GET"), StrPtr(Path), StrPtr("HTTP/1.1"), 0, 0, 0)
    If hRequest_ = 0 Then Err.Raise 515, "BiDiSocket", "WinHttpOpenRequest failed"

    ' Step 5: Configure the request for WebSocket upgrade
    If WinHttpSetOption(hRequest_, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0) = 0 Then
        Err.Raise 516, "BiDiSocket", "Failed to set WebSocket upgrade option"
    End If

    ' Step 6: Send the upgrade request and wait for the response
    If WinHttpSendRequest(hRequest_, 0, 0, 0, 0, 0, 0) = 0 Then Err.Raise 517, "BiDiSocket", "SendRequest failed"
    If WinHttpReceiveResponse(hRequest_, 0) = 0 Then Err.Raise 518, "BiDiSocket", "ReceiveResponse failed"

    ' Step 7: Verify Status 101 (Switching Protocols)
    Dim status As Long, Size As Long: Size = 4
    WinHttpQueryHeaders hRequest_, WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, status, Size, WINHTTP_NO_HEADER_INDEX
    
    If status = 101 Then
        ' Step 8: Complete the upgrade and get the WebSocket handle
        hWebSocket_ = WinHttpWebSocketCompleteUpgrade(hRequest_, 0)
        If hWebSocket_ = 0 Then Err.Raise 520, "BiDiSocket", "WebSocket upgrade failed"

        ' Step 9: Configure Keep-Alive to maintain connection
        Dim ka As Long: ka = KEEPALIVE_INTERVAL_MS
        WinHttpSetOption hWebSocket_, WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL, VarPtr(ka), 4
        
        ' Finalize: Success transition
        p_lastActivityTime = GetTickMS()
        m_RetryCount = 0 ' Reset retry counter on successful handshake
        SetState StateConnected
        Connect = True
    Else
        Err.Raise 519, "BiDiSocket", "Upgrade rejected: " & status
    End If
    Exit Function

ErrorHandler:
    ' On any failure, transition to Error state
    SetState StateError
    Call InternalCleanup
    Connect = False
End Function

' ========================================================================================
' Sends a raw UTF-8 string message through the WebSocket.
' Uses a linear (non-recursive) retry logic to handle transient connection failures.
' ========================================================================================
Public Function SendMessage(ByVal msg As String) As Boolean
    ' 1. Initial Recovery: If the state is already Error, attempt to restore connection
    If m_State = StateError Then
        If Not AttemptRecovery() Then Exit Function
    End If

    ' 2. Connectivity Guard
    If m_State <> StateConnected Then
        If DEBUG_MODE Then Debug.Print "BiDi: SendMessage aborted. Not connected."
        Exit Function
    End If

    ' 3. Data Conversion: VBA String (UTF-16) to raw Byte array (UTF-8)
    ' Performed only once to ensure efficiency.
    Dim data() As Byte: data = ToUtf8Bytes(msg)
    Dim dataSize As Long: dataSize = UBound(data) + 1
    
    ' 4. Primary Transmission Attempt
    Dim ret As Long
    ret = WinHttpWebSocketSend(hWebSocket_, WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, VarPtr(data(0)), dataSize)
    
    If ret = NO_ERROR Then
        SendMessage = True
        p_lastActivityTime = GetTickMS()
        Exit Function
    End If

    ' 5. Failure Handling: If the first attempt failed, trigger recovery
    If DEBUG_MODE Then Debug.Print "BiDi: Primary SendMessage failed (Error: " & ret & "). Attempting recovery..."
    
    ' Mark as Error and clean up handles (preserving the state for AttemptRecovery)
    SetState StateError
    Disconnect keepErrorState:=True
    
    ' 6. Recovery & Final Retry Attempt
    ' AttemptRecovery calls Connect(), which refreshes the hWebSocket_ handle.
    If AttemptRecovery() Then
        ' Perform the second attempt using the same binary data buffer
        ret = WinHttpWebSocketSend(hWebSocket_, WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, VarPtr(data(0)), dataSize)
        
        If ret = NO_ERROR Then
            If DEBUG_MODE Then Debug.Print "BiDi: SendMessage succeeded after recovery."
            SendMessage = True
            p_lastActivityTime = GetTickMS()
        Else
            If DEBUG_MODE Then Debug.Print "BiDi: SendMessage failed again after recovery (Error: " & ret & ")."
            ' Ensure we remain in Error state if the retry also fails
            SetState StateError
            Disconnect keepErrorState:=True
        End If
    End If
End Function

' ========================================================================================
' Properly closes the WebSocket and all associated WinHttp handles.
' If keepErrorState is True, it performs cleanup without overwriting the Error state.
' ========================================================================================
Public Sub Disconnect(Optional ByVal keepErrorState As Boolean = False)
    ' Guard: Prevent redundant calls, but allow cleanup if we need to preserve an Error state
    If Not keepErrorState Then
        If m_State = StateDisconnecting Or m_State = StateDisconnected Then Exit Sub
        SetState StateDisconnecting
    End If
    
    ' Release all WinHttp and WebSocket API handles
    Call InternalCleanup
    
    ' Clear the local event queue
    Set m_EventBuffer = New Collection
    
    ' Finalize: Transition to fully Disconnected only if not in an Error recovery flow
    If Not keepErrorState Then
        SetState StateDisconnected
    End If
End Sub

' ========================================================================================
' Private helper to release WinHttp handles without triggering state events.
' ========================================================================================
Private Sub InternalCleanup()
    On Error Resume Next
    If hWebSocket_ <> 0 Then
        ' Graceful WebSocket close (Status 1000: Normal Closure)
        WinHttpWebSocketClose hWebSocket_, 1000, 0, 0
        WinHttpCloseHandle hWebSocket_
        hWebSocket_ = 0
    End If
    If hRequest_ <> 0 Then WinHttpCloseHandle hRequest_: hRequest_ = 0
    If hConnection_ <> 0 Then WinHttpCloseHandle hConnection_: hConnection_ = 0
    If hSession_ <> 0 Then WinHttpCloseHandle hSession_: hSession_ = 0
    On Error GoTo 0
End Sub

' ========================================================================================
' Sends a JSON message and waits for a specific response ID.
' Features fail-fast state monitoring and infinite loop prevention.
' ========================================================================================
Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    ' Connectivity Guard: Ensure operational state before sending
    If m_State <> StateConnected Then
        If Not AttemptRecovery() Then Err.Raise 600, "BiDiSocket", "Not connected"
    End If
    
    Dim cmdId As String: cmdId = ExtractJsonId(jsonMessage)
    If Not SendMessage(jsonMessage) Then Err.Raise 601, "BiDiSocket", "Send failed"
    
    Dim timeout As Long: timeout = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    Dim deadline As Double: deadline = GetTickMS() + timeout
    Dim msg As String, inId As String
    Dim loopCount As Long: loopCount = 0
    
    Do
        ' State Machine Guard: If state shifts to Error/Disconnected, abort immediately
        If m_State <> StateConnected Then
            Err.Raise 605, "BiDiSocket", "Connection lost while waiting for ID: " & cmdId
        End If

        ' Infinite Loop Insurance: Hard cap to prevent CPU hang
        loopCount = loopCount + 1
        If loopCount > 50000 Then
            Err.Raise 606, "BiDiSocket", "Internal loop overflow: Potential network or logic hang."
        End If

        ' Attempt to retrieve assembled message from wire
        msg = ReadFromWire(deadline)
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            If inId = cmdId Then
                SendAndReceive = msg
                Exit Function
            Else
                ' Cache unrelated events for GetMessage processing
                m_EventBuffer.Add msg
            End If
        End If
        
        ' Strict Deadline Check
        If GetTickMS() > deadline Then
            Err.Raise 604, "BiDiSocket", "Timeout waiting for response ID: " & cmdId
        End If

        YieldExecution 2, m_AllowUiPump
    Loop
End Function

' ========================================================================================
' Internal recovery logic using exponential backoff.
' Prevents flooding the driver during its restart or temporary network glitches.
' ========================================================================================
Private Function AttemptRecovery() As Boolean
    If Not m_AutoReconnect Then Exit Function
    If m_RetryCount >= m_MaxRetryAttempts Then
        If DEBUG_MODE Then Debug.Print "BiDi: Max recovery attempts reached."
        Exit Function
    End If

    m_RetryCount = m_RetryCount + 1
    
    ' Exponential backoff: 1000ms, 2000ms, 4000ms...
    Dim waitTime As Long: waitTime = (2 ^ (m_RetryCount - 1)) * 1000
    If DEBUG_MODE Then Debug.Print "BiDi: Recovery attempt " & m_RetryCount & " (Waiting " & waitTime & "ms)"
    
    YieldExecution waitTime, True ' Keep UI responsive during wait
    
    ' Try to reconnect using the last known valid credentials
    If Connect(m_LastWsUrl, m_LastPort) Then
        If DEBUG_MODE Then Debug.Print "BiDi: Recovery successful."
        m_RetryCount = 0 ' Reset on success
        AttemptRecovery = True
    Else
        AttemptRecovery = False
    End If
End Function

' ========================================================================================
' Retrieves a message from the local buffer or polls the wire via heartbeat.
' Updated to support automated recovery during polling.
' ========================================================================================
Public Function GetMessage() As String
    ' 1. Check the local event buffer first (No network activity required)
    If m_EventBuffer.Count > 0 Then
        GetMessage = m_EventBuffer(1)
        m_EventBuffer.Remove 1
        Exit Function
    End If
    
    ' 2. Recovery Guard: If in error, try to restore connection before polling
    If m_State = StateError Then
        If Not AttemptRecovery() Then Exit Function
    End If
    
    If m_State <> StateConnected Then Exit Function
    
    ' 3. Polling Interval Check
    Dim nowT As Double: nowT = GetTickMS()
    If (nowT - p_lastPollTime) < m_PollingInterval Then Exit Function
    
    p_lastPollTime = nowT
    
    ' 4. Active Polling (Heartbeat Ping)
    Dim pingCmd As String
    pingCmd = "{""id"":" & SYSTEM_PING_ID & ", ""method"":""session.status"", ""params"":{}}"
    
    ' If the send fails, it will trigger SetState StateError internally
    If Not SendMessage(pingCmd) Then Exit Function
    
    ' 5. Read Response (Short deadline to keep UI responsive)
    Dim rawMsg As String: rawMsg = ReadFromWire(GetTickMS() + 500)
    
    If rawMsg = "" Then
        YieldExecution 1, m_AllowUiPump
        Exit Function
    End If
    
    ' 6. Filter and Return
    Dim inId As String: inId = ExtractJsonId(rawMsg)
    If inId = CStr(SYSTEM_PING_ID) Then
        GetMessage = ""
    Else
        GetMessage = rawMsg
    End If
End Function

' ========================================================================================
' Reads incoming WebSocket frames using a class-level persistent buffer.
' Optimized to minimize heap allocations and handles errors without clearing the Error state.
' ========================================================================================
Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    ' Guard: Only allow reading if in a valid operational state or during intentional disconnect
    If m_State <> StateConnected And m_State <> StateDisconnecting Then Exit Function

    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim hasData As Boolean: hasData = False

    ' Prepare the persistent ADODB.Stream (Binary mode) to receive raw frames
    PrepareStream adTypeBinary

    Do
        ' Receive a fragment into the pre-allocated buffer
        ret = WinHttpWebSocketReceive(hWebSocket_, m_ReceiveBuffer(0), m_BufferSize, bytesRead, bufType)
        
        ' --- Error Handling Logic ---
        If ret <> NO_ERROR Then
            If DEBUG_MODE Then Debug.Print "BiDi: Receive error. Code: " & ret
            
            ' Transition to Error state to signal that the connection is broken
            SetState StateError
            
            ' Perform internal handle cleanup while preserving the StateError
            ' This allows SendMessage/GetMessage to trigger AttemptRecovery later
            Disconnect keepErrorState:=True
            Exit Function
        End If

        ' --- Remote Closure Handling ---
        If bufType = WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE Then
            If DEBUG_MODE Then Debug.Print "BiDi: Connection closed by remote host."
            Disconnect
            Exit Function
        End If

        ' --- Data Processing ---
        If bytesRead > 0 Then
            hasData = True
            
            ' Write the fragment to the stream
            ' Optimization: Write directly from buffer if full, otherwise slice via temp array
            If bytesRead = m_BufferSize Then
                m_Stream.Write m_ReceiveBuffer
            Else
                Dim temp() As Byte: ReDim temp(bytesRead - 1)
                Dim i As Long
                For i = 0 To bytesRead - 1: temp(i) = m_ReceiveBuffer(i): Next i
                m_Stream.Write temp
            End If
            
            ' Check if the persistent buffer size needs to be increased for future large payloads
            MaybeGrowBuffer bytesRead
        End If

        ' --- Frame Control ---
        ' Break the loop once a complete UTF-8 message frame is fully assembled
        If bufType = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then Exit Do
        
        ' Safety: Exit if the provided deadline is exceeded during a fragmented read
        If deadline > 0 Then
            If GetTickMS() > deadline Then Exit Do
        End If

        ' Allow short pause to prevent CPU pegging during long fragment sequences
        YieldExecution 0, False
    Loop

    ' --- Final Assembly & UTF-8 Decoding ---
    If hasData Then
        m_Stream.Position = 0
        m_Stream.Type = adTypeText
        m_Stream.Charset = "UTF-8"
        
        ' Convert the binary stream back into a VBA UTF-16 String
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
                
                ' Sync the actual byte array with the new buffer size
                ReDim m_ReceiveBuffer(m_BufferSize - 1)
                
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

' ========================================================================================
' UTF-8 Conversion Utility
' ========================================================================================
Private Function ToUtf8Bytes(ByVal txt As String) As Byte()
    ' 1. Prepare stream in text mode for encoding conversion
    PrepareStream adTypeText
    m_Stream.Charset = "UTF-8"
    m_Stream.WriteText txt
    
    ' 2. Switch to binary mode to evaluate the raw byte structure
    m_Stream.Position = 0
    m_Stream.Type = adTypeBinary
    
    ' 3. Dynamic BOM Detection (EF BB BF)
    ' ADODB.Stream typically adds a 3-byte BOM for UTF-8.
    ' We verify its existence to avoid accidentally stripping valid data.
    If m_Stream.Size >= 3 Then
        Dim header() As Byte
        header = m_Stream.Read(3) ' Retrieve the first 3 bytes
        
        ' Check for UTF-8 BOM hex sequence: EF BB BF
        If header(0) = &HEF And header(1) = &HBB And header(2) = &HBF Then
            ' BOM found: set position to skip it
            m_Stream.Position = 3
        Else
            ' BOM not found: reset position to the start
            m_Stream.Position = 0
        End If
    Else
        m_Stream.Position = 0
    End If
    
    ' 4. Extract the payload
    ' Ensure we don't attempt to read beyond the stream size
    If m_Stream.Size > m_Stream.Position Then
        ToUtf8Bytes = m_Stream.Read
    Else
        ' Return an initialized empty byte array to prevent UBound errors in SendMessage
        ToUtf8Bytes = StrConv("", vbFromUnicode)
    End If
End Function
Private Sub YieldExecution(Optional ByVal ms As Long = 1, Optional ByVal allowDoEvents As Boolean = False)
    If allowDoEvents Then DoEvents
    If ms >= 0 Then SleepWinAPI ms
End Sub

' ========================================================================================
' Stream Management Helper
' ========================================================================================

' Resets the stream state and sets the required type (binary/text) without closing/reopening.
' This improves performance by reusing the allocated memory buffer and avoiding handle overhead.
Private Sub PrepareStream(ByVal sType As StreamTypeEnum)
    If m_Stream.State = adStateClosed Then m_Stream.Open
    m_Stream.Position = 0
    m_Stream.SetEOS ' Truncate existing data to clear the buffer
    m_Stream.Type = sType
End Sub
