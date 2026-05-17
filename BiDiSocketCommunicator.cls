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
' Purpose    : Professional WebSocket Client for WebDriver BiDi (W3C Standard).
'
' [SECURITY & COMPLIANCE - ANALYST REFERENCE]
' 1. SANDBOXED IPC SCOPE: Strictly restricted to the Local Loopback interface (127.0.0.1).
'    Engineered exclusively for Inter-Process Communication (IPC) between the VBA host
'    and the local WebDriver/Browser instance.
' 2. ANTI-RAT DESIGN: This implementation enforces a strict URL validation policy that
'    blocks non-localhost routing, preventing its use as a Remote Access Tool (RAT).
' 3. NO DATA EXFILTRATION: By architectural design, this module cannot initiate connections
'    to remote C2 (Command & Control) servers, ensuring zero risk of outbound exfiltration.
' 4. NATIVE OS INTEGRITY: Utilizes 'WinHttp.dll' (a signed, core Windows System Component)
'    for socket handshaking, complying with enterprise-level EDR auditing and OS policies.
' 5. OPEN-SOURCE TRANSPARENCY: Part of the SeleniumVBA project (GCuser99).
'    The full source code and logic are available for independent security audits:
'    Audit URL: https://github.com/GCuser99/SeleniumVBA
' 6. ARCHITECTURAL NECESSITY (W3C BiDi): Required for compliance with the W3C WebDriver BiDi
'    standard. Unlike legacy one-way REST APIs, BiDi mandates a persistent WebSocket for
'    bidirectional events (e.g., real-time network interception, console logs).
'    The communication is inherently safe as the WebDriver process only exposes its
'    BiDi endpoint to the local loopback interface.
' ========================================================================================
' --- Configuration ---
Private Const DEBUG_MODE As Boolean = False
Private Const AGENT_NAME As String = "SeleniumVBA/BiDi-Client (Open-Source; Local-IPC-Only)"

' --- WinHttp API Declarations ---
Private Declare PtrSafe Function WinHttpOpen Lib "WinHttp" (ByVal pszAgentW As LongPtr, ByVal dwAccessType As Long, ByVal pszProxyW As LongPtr, ByVal pszProxyBypassW As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpConnect Lib "WinHttp" (ByVal hSession As LongPtr, ByVal pswzServerName As LongPtr, ByVal nServerPort As Long, ByVal dwReserved As Long) As LongPtr
Private Declare PtrSafe Function WinHttpOpenRequest Lib "WinHttp" (ByVal hConnect As LongPtr, ByVal pwszVerb As LongPtr, ByVal pwszObjectName As LongPtr, ByVal pwszVersion As LongPtr, ByVal pwszReferrer As LongPtr, ByVal ppwszAcceptTypes As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpSendRequest Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As Long) As Long
Private Declare PtrSafe Function WinHttpReceiveResponse Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal lpReserved As LongPtr) As Long
Private Declare PtrSafe Function WinHttpQueryHeaders Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, ByVal pwszName As LongPtr, ByRef lpBuffer As Any, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Long
Private Declare PtrSafe Function WinHttpSetOption Lib "WinHttp" (ByVal hInternet As LongPtr, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpSetTimeouts Lib "WinHttp" (ByVal hInternet As LongPtr, ByVal nResolveTimeout As Long, ByVal nConnectTimeout As Long, ByVal nSendTimeout As Long, ByVal nReceiveTimeout As Long) As Long
Private Declare PtrSafe Function WinHttpCloseHandle Lib "WinHttp" (ByVal hInternet As LongPtr) As Long

' --- WebSocket API Declarations ---
Private Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "WinHttp" (ByVal hRequest As LongPtr, ByVal pContext As LongPtr) As LongPtr
Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal eBufferType As Long, ByRef pvBuffer As Any, ByVal dwBufferLength As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByRef pvBuffer As Any, ByVal dwBufferLength As Long, ByRef pdwBytesRead As Long, ByRef peBufferType As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketClose Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal usStatus As Long, ByRef pvReason As Any, ByVal dwReasonLength As Long) As Long

' --- Utility API Declarations ---
Private Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- WinHttp Constants ---
Private Const WINHTTP_ACCESS_TYPE_NO_PROXY As Long = 1
Private Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET As Long = 114
Private Const WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL As Long = 116
Private Const WINHTTP_QUERY_STATUS_CODE As Long = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER As Long = &H20000000
Private Const WINHTTP_HEADER_NAME_BY_INDEX As Long = 0
Private Const WINHTTP_NO_HEADER_INDEX As Long = 0

' --- WebSocket Buffer Types ---
Private Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE  As Long = 0
Private Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE As Long = 1
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE    As Long = 2
Private Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE   As Long = 3
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE           As Long = 4

' --- Library Settings ---
Private Const NO_ERROR As Long = 0
Private Const SYSTEM_PING_ID As Long = -999
Private Const DEFAULT_BUFFER_SIZE As Long = 8192
Private Const MAX_BUFFER_SIZE As Long = 10485760 ' 10MB
Private Const DEFAULT_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_POLLING_INTERVAL_MS As Long = 100
Private Const KEEPALIVE_INTERVAL_MS As Long = 15000
Private Const HEARTBEAT_INTERVAL_MS As Long = 20000
Private Const ENABLE_HEARTBEAT_LOGGING As Boolean = True

' --- State & Handle Variables ---
Private hSession_ As LongPtr
Private hConnection_ As LongPtr
Private hRequest_ As LongPtr
Private hWebSocket_ As LongPtr
Private m_LastWsUrl As String
Private m_LastPort As Long
Private m_RetryCount As Long
Private m_MaxRetryAttempts As Long
Private m_AutoReconnect As Boolean
Private m_consecutiveInternalErrors As Long

' --- Robustness Semaphores & Timestamps ---
Private m_isReading As Boolean
Private m_isReadingSince As Double ' Tracks the timestamp of the last read lock
Private m_isSendReceiving As Boolean ' True when a command sequence is active

Public Enum BiDiSocketState
    StateDisconnected = 0
    StateConnecting = 1
    StateConnected = 2
    StateDisconnecting = 3
    StateError = 4
End Enum
Private m_State As BiDiSocketState
Public Event StateChanged(ByVal NewState As BiDiSocketState, ByVal stateName As String)

' --- Internal Buffers & Objects ---
Private m_EventBuffer As Collection
Private m_DefaultTimeout As Long
Private m_PollingInterval As Long
Private m_BufferSize As Long
Private m_MaxBufferSize As Long
Private m_ReceiveBuffer() As Byte
Private m_GrowThresholdCount As Long
Private m_RecentFullReads As Long
Private m_InStream As ADODB.Stream  ' Dedicated Read Stream
Private m_OutStream As ADODB.Stream ' Dedicated Write Stream
Private m_DroppedEventsCount As Long
Private p_lastPollTime As Double
Private p_RegExp As RegExp
Private m_AllowUiPump As Boolean
Private m_qpcFreq As Currency ' QueryPerformanceFrequency cache (set in Class_Initialize)
Private m_lastPumpTicks As Currency ' Throttle state for PumpMessages

' ========================================================================================
' Initialization & Termination
' ========================================================================================
Private Sub Class_Initialize()
    ' Fail-fast QPC initialization. Required for all timer APIs (NowTicks/NowMs/ElapsedMs).
    If QueryPerformanceFrequency(m_qpcFreq) = 0 Or m_qpcFreq <= 0 Then
        Err.Raise vbObjectError + 9001, "BiDiSocketCommunicator", _
                  "QueryPerformanceFrequency is unavailable on this system. QPC support is required."
    End If
    m_DefaultTimeout = DEFAULT_TIMEOUT_MS
    m_PollingInterval = DEFAULT_POLLING_INTERVAL_MS
    m_AllowUiPump = True
    
    m_BufferSize = DEFAULT_BUFFER_SIZE
    ReDim m_ReceiveBuffer(m_BufferSize - 1)
    m_MaxBufferSize = MAX_BUFFER_SIZE

    Set m_EventBuffer = New Collection
    
    ' Maintain separate streams for Read and Write to prevent reentrancy collisions
    Set m_InStream = New ADODB.Stream
    m_InStream.mode = adModeReadWrite
    m_InStream.Open

    Set m_OutStream = New ADODB.Stream
    m_OutStream.mode = adModeReadWrite
    m_OutStream.Open
    
    ' Initial pattern for ExtractJsonId (will be updated JIT if needed)
    Set p_RegExp = New RegExp
    p_RegExp.Global = False
    p_RegExp.IgnoreCase = True
    
    ' Enable controlled auto-reconnect by default
    m_MaxRetryAttempts = 5      ' up to 5 attempts with exponential backoff
    m_AutoReconnect = True      ' allow recovery

    ' Moderate buffer growth threshold (3 consecutive full reads)
    m_GrowThresholdCount = 3

    SetState StateDisconnected
End Sub

Private Sub Class_Terminate()
' Ensure all handles, streams, and semaphores are cleaned up.
    ' [FIX] Reset communication semaphores BEFORE Disconnect.
    ' If the macro was forcefully stopped (Stop button, End statement, Ctrl+Break)
    ' during an active SendAndReceive or ReadFromWire cycle, these flags remain True.
    ' Without resetting them here, any subsequent reuse of the object (or a new instance
    ' that inherits the module-level state in edge cases) would deadlock permanently.
    m_isSendReceiving = False
    m_isReading = False
    m_isReadingSince = 0
    
    Disconnect
    If Not m_InStream Is Nothing Then
        If m_InStream.state = adStateOpen Then m_InStream.Close
        Set m_InStream = Nothing
    End If
    If Not m_OutStream Is Nothing Then
        If m_OutStream.state = adStateOpen Then m_OutStream.Close
        Set m_OutStream = Nothing
    End If
    Set p_RegExp = Nothing
    Set m_EventBuffer = Nothing
End Sub

' Resets the target stream state and sets the required type without cross-talk
Private Sub PrepareStream(ByVal strm As ADODB.Stream, ByVal sType As StreamTypeEnum)
    If strm.state = adStateClosed Then strm.Open
    strm.Position = 0
    strm.SetEOS ' Truncate existing data
    strm.Type = sType
End Sub

' Helper to safely release stream resources
Private Sub CloseStream(ByRef strm As ADODB.Stream)
    If Not strm Is Nothing Then
        If strm.state = adStateOpen Then strm.Close
        Set strm = Nothing
    End If
End Sub

' ========================================================================================
' Public Properties
' ========================================================================================
Public Property Get IsConnected() As Boolean: IsConnected = (hWebSocket_ <> 0): End Property
Public Property Let PollingInterval(ByVal v As Long): m_PollingInterval = v: End Property
Public Property Let AllowUiPump(ByVal v As Boolean): m_AllowUiPump = v: End Property
Public Property Get AllowUiPump() As Boolean: AllowUiPump = m_AllowUiPump: End Property
Public Property Get state() As BiDiSocketState: state = m_State: End Property
Public Property Let AutoReconnect(ByVal v As Boolean): m_AutoReconnect = v: End Property
Public Property Get AutoReconnect() As Boolean: AutoReconnect = m_AutoReconnect: End Property
Public Property Get ConsecutiveInternalErrors() As Long: ConsecutiveInternalErrors = m_consecutiveInternalErrors: End Property
Public Property Get DroppedEventsCount() As Long: DroppedEventsCount = m_DroppedEventsCount: End Property

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
    
    ' --- SECURITY ENFORCEMENT: STRICT LOOPBACK VALIDATION ---
    ' [INTENT] Ensure the WebSocket connection is strictly bound to the local machine.
    ' [FIX] Updated Regular Expression to include IPv6 loopback (\[::1\]) and prevent domain spoofing.
    p_RegExp.pattern = "^wss?://(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?(?:/|$)"
    p_RegExp.IgnoreCase = True
    
    If Not p_RegExp.Test(wsUrl) Then
        Trace "[Security] Blocked: Connection attempt to non-local address. " & _
              "BiDi communication is restricted to localhost for safety.", True
        SetState StateError
        Exit Function
    End If
    
    ' Store connection parameters for potential recovery attempts
    m_LastWsUrl = wsUrl
    m_LastPort = localPort

    On Error GoTo ErrorHandler

    ' Step 1: Initialize transition to Connecting state
    SetState StateConnecting
    
    ' Ensure any residual handles are cleaned up before a new attempt
    Call InternalCleanup
    
    ' Basic parameter validation
    If wsUrl = "" Or localPort = 0 Then
        If DEBUG_MODE Then Debug.Print "BiDi: Connect failed. Invalid URL or Port."
        GoTo ErrorHandler
    End If
    
    Dim Path As String: Path = ExtractPath(wsUrl)
    
    ' [FIX] Dynamically extract the host to match the actual binding
    Dim host As String: host = ExtractHost(wsUrl)

    ' Step 2: Initialize WinHttp Session
    hSession_ = WinHttpOpen(StrPtr(AGENT_NAME), WINHTTP_ACCESS_TYPE_NO_PROXY, 0, 0, 0)
    If hSession_ = 0 Then Err.Raise 513, "BiDiSocket", "WinHttpOpen failed"

    ' --- CRITICAL FIX: UI RESPONSIVENESS vs CPU BALANCE ---
    ' Set connection-related timeouts (Resolve, Connect, Send, Receive)
    WinHttpSetTimeouts hSession_, 2000, 2000, 5000, 100
    
    ' Step 3: Connect to the WebDriver host
    hConnection_ = WinHttpConnect(hSession_, StrPtr(host), localPort, 0)
    If hConnection_ = 0 Then Err.Raise 514, "BiDiSocket", "WinHttpConnect failed"

    ' Step 4: Open an HTTP request to initiate the WebSocket handshake
    hRequest_ = WinHttpOpenRequest(hConnection_, StrPtr("GET"), StrPtr(Path), StrPtr("HTTP/1.1"), 0, 0, 0)
    If hRequest_ = 0 Then Err.Raise 515, "BiDiSocket", "WinHttpOpenRequest failed"

    ' Step 5: Configure the request for WebSocket upgrade
    Dim nullPtr As LongPtr: nullPtr = 0
    If WinHttpSetOption(hRequest_, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, ByVal nullPtr, 0) = 0 Then
        Err.Raise 516, "BiDiSocket", "Failed to set WebSocket upgrade option"
    End If

    ' Step 6: Send the upgrade request and wait for the response
    If WinHttpSendRequest(hRequest_, 0, 0, 0, 0, 0, 0) = 0 Then Err.Raise 517, "BiDiSocket", "SendRequest failed"
    If WinHttpReceiveResponse(hRequest_, 0) = 0 Then Err.Raise 518, "BiDiSocket", "ReceiveResponse failed"

    ' Step 7: Verify Status 101 (Switching Protocols)
    Dim status As Long, Size As Long: Size = 4
    Dim idx As Long: idx = WINHTTP_NO_HEADER_INDEX
    Dim headerRet As Long
    
    headerRet = WinHttpQueryHeaders(hRequest_, _
                        WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, _
                        WINHTTP_HEADER_NAME_BY_INDEX, _
                        status, Size, idx)
    
    If headerRet <> 0 And status = 101 Then
        ' Step 8: Complete the upgrade and get the WebSocket handle
        hWebSocket_ = WinHttpWebSocketCompleteUpgrade(hRequest_, 0)
        
        ' [CRITICAL 4 FIX] Defensive handling for partially transferred handles.
        ' If CompleteUpgrade fails, explicit closure of hRequest_ is attempted
        ' under Resume Next to prevent double-free or leaks regardless of OS implementation.
        If hWebSocket_ = 0 Then
            On Error Resume Next
            WinHttpCloseHandle hRequest_
            On Error GoTo ErrorHandler
            hRequest_ = 0
            Err.Raise 520, "BiDiSocket", "WebSocket upgrade failed"
        End If

        ' Step 9: Configure Keep-Alive to maintain connection
        Dim ka As Long: ka = KEEPALIVE_INTERVAL_MS
        WinHttpSetOption hWebSocket_, WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL, ka, 4
        
        ' Reset session-specific statistics.
        m_DroppedEventsCount = 0
        
        ' Finalize: Success transition
        p_lastPollTime = 0
        m_RetryCount = 0
        SetState StateConnected
        Connect = True
    Else
        Err.Raise 519, "BiDiSocket", "Upgrade rejected. API Result: " & headerRet & " Status: " & status
    End If
    Exit Function

ErrorHandler:
    ' On any failure, transition to Error state and perform handle cleanup
    SetState StateError
    Call InternalCleanup
    Connect = False
End Function

' ========================================================================================
' Sends a raw UTF-8 string message through the WebSocket.
' ========================================================================================
Public Function SendMessage(ByVal msg As String) As Boolean
    If LenB(msg) = 0 Then Exit Function
    
    ' NEW: Log every send attempt before it happens
    Trace ">>> [Socket] Attempting to send message..."

    If m_State = StateError Then
        If Not AttemptRecovery() Then Exit Function
    End If

    If m_State <> StateConnected Then
        Trace ">>> [Socket] Send aborted. Not connected.", True
        Exit Function
    End If

    Dim data() As Byte: data = ToUtf8Bytes(msg)
    Dim dataSize As Long: dataSize = UBound(data) + 1
    
    Dim ret As Long
    ret = WinHttpWebSocketSend(hWebSocket_, _
                               WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                               data(0), _
                               dataSize)
    
    If ret = NO_ERROR Then
        SendMessage = True
        Exit Function
    End If

    Trace ">>> [Socket] SendMessage failed. Error: " & ret, True
    SetState StateError
    Disconnect keepErrorState:=True
    
    ' [FIX] PROHIBIT RESENDING (Protocol Compliance)
    ' W3C WebDriver BiDi mandates unique request IDs. Resending an identical JSON payload
    ' after a socket drop violates the protocol and corrupts the Wrapper's request map.
    ' We heal the connection here, but strictly return False so the upper layer
    ' handles the failure (e.g., re-generating the ID or raising an exception).
    Call AttemptRecovery
    SendMessage = False
End Function

' ========================================================================================
' PostScript: Fire & Forget transport. Sends a message without expecting a response.
' Rationale: Decouples the wrapper from the physical SendMessage implementation.
' ========================================================================================
Public Sub PostScript(ByVal msg As String)
    SendMessage msg
End Sub

' ========================================================================================
' Properly closes the WebSocket and all associated WinHttp handles.
' ========================================================================================
Public Sub Disconnect(Optional ByVal keepErrorState As Boolean = False)
    ' If we are already in the middle of a connection attempt,
    ' we should be cautious about forced disconnection unless it's a hard reset.
    If m_State = StateDisconnecting Or m_State = StateDisconnected Then Exit Sub

    ' [STEP 1] Log the final count of dropped events to bidi_socket_trace.log.
    ' This acts as a "Post-Mortem" report for debugging high-load scenarios.
    ' We do not reset the counter here so it remains accessible via property after disconnect.
    If m_DroppedEventsCount > 0 Then
        Trace ">>> [Session Summary] Connection closing. Total events dropped during this session: " & m_DroppedEventsCount, True
    End If

    If Not keepErrorState Then SetState StateDisconnecting
    
    ' Perform the actual handle cleanup
    Call InternalCleanup
    
    ' Clear the event buffer
    Set m_EventBuffer = New Collection
    
    ' Finalize state only if we are not preserving an Error state for recovery
    If Not keepErrorState Then
        SetState StateDisconnected
    End If
End Sub

' ========================================================================================
' Private helper to release WinHttp handles.
' STRICT RULE: This method MUST NOT modify m_State.
' ========================================================================================
Private Sub InternalCleanup()
    On Error Resume Next
    
    ' Track if ownership of the request handle was transferred to the WebSocket
    Dim wasUpgraded As Boolean
    wasUpgraded = (hWebSocket_ <> 0)
    
    ' 1. Close WebSocket Handle
    If hWebSocket_ <> 0 Then
        ' Attempt to send a close frame (Status 1000: Normal Closure).
        ' We don't wait for the server's response here to ensure the cleanup
        ' remains fast and non-blocking during error or disconnect sequences.
        WinHttpWebSocketClose hWebSocket_, 1000, ByVal 0&, 0
        WinHttpCloseHandle hWebSocket_
        hWebSocket_ = 0
    End If
    
    ' 2. Close HTTP Request Handle
    If hRequest_ <> 0 Then
        ' [FIX] Request handle ownership is transferred after CompleteUpgrade.
        ' Closing it after a successful upgrade can cause double-free corruption.
        ' It must only be explicitly closed if the upgrade never occurred.
        If Not wasUpgraded Then
            WinHttpCloseHandle hRequest_
        End If
        hRequest_ = 0
    End If
    
    ' 3. Close HTTP Connection Handle
    If hConnection_ <> 0 Then
        WinHttpCloseHandle hConnection_
        hConnection_ = 0
    End If
    
    ' 4. Close WinHttp Session Handle
    If hSession_ <> 0 Then
        WinHttpCloseHandle hSession_
        hSession_ = 0
    End If
    
    ' 5. Purge the input stream to prevent cross-session data corruption
    ' If a connection drops mid-fragment, the residual bytes must be truncated
    ' before the next socket connection is established.
    If Not m_InStream Is Nothing Then
        If m_InStream.state = adStateOpen Then
            m_InStream.Position = 0
            m_InStream.SetEOS
            m_InStream.Type = adTypeBinary
        End If
    End If
    
    On Error GoTo 0
End Sub

' ========================================================================================
' Sends a JSON message and waits for a specific response ID with a strict hard deadline.
' Refactored to establish a "High-Purity Loop" that prevents WinHttp state corruption
' by suppressing DoEvents during the active command-response cycle.
' ========================================================================================
Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    ' 1. Immediate Re-entry Guard: Lock the communication channel
    If m_isSendReceiving Then Exit Function
    m_isSendReceiving = True
    
    Dim succeeded As Boolean: succeeded = False
    Dim msg As String, inId As String
    Dim cmdId As String
    
    On Error GoTo Cleanup

    ' 2. Connectivity Guard
    If m_State <> StateConnected Then
        If Not AttemptRecovery() Then
            Err.Raise 600, "BiDiSocket", "Not connected and recovery failed."
        End If
    End If
    
    cmdId = ExtractJsonId(jsonMessage)
    
    ' 3. Send Operation
    If Not SendMessage(jsonMessage) Then
        If m_State <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost during send operation."
        Else
            Err.Raise 601, "BiDiSocket", "Send failed due to internal WinHttp error."
        End If
    End If
    
    ' 4. Timeout and Deadline Calculation
    Dim timeout As Long: timeout = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    Dim deadline As Double: deadline = NowMs() + timeout
    
    ' Memory Protection Constants
    Const MAX_EVENT_BUFFER_SIZE As Long = 1000
    
    ' 5. Receive Loop (High-Purity)
    Do
        If m_State <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost while waiting for ID: " & cmdId
        End If

        ' Attempt to read a frame from the wire
        msg = ReadFromWire(deadline)
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            If inId = cmdId Then
                ' Success: Target ID matched
                SendAndReceive = msg
                succeeded = True
                GoTo Cleanup
            Else
                ' Async events or pongs: Queue for GetMessage processing
                ' [FIX] Buffer all messages except heartbeat pongs.
                ' BiDi events (network.beforeRequestSent etc.) have no "id" field,
                ' so the old condition (inId <> "") silently discarded them.
                ' Losing these one-shot events degraded SPA idle detection accuracy.
                If inId <> CStr(SYSTEM_PING_ID) Then
                    m_EventBuffer.Add msg
                    
                    ' FIFO Protection to prevent memory exhaustion
                    If m_EventBuffer.count > MAX_EVENT_BUFFER_SIZE Then
                        m_EventBuffer.Remove 1
                        m_DroppedEventsCount = m_DroppedEventsCount + 1
                    End If
                End If
            End If
        End If
        
        ' Hard Timeout Check
        If NowMs() > deadline Then
            ' [FIX] Transition to Error state and raise exception
            SetState StateError
            Err.Raise 604, "BiDiSocket", "Socket Silence: Hard timeout waiting for ID: " & cmdId
        End If

        ' 6. Yield execution (Strict Guarded Yield)
        YieldExecution 5, False
    Loop

Cleanup:
    ' 7. Finalization: Release the communication lock
    ' [FIX] Clear any residual Err state from intermediate operations (e.g., SetState
    ' or InternalCleanup under Resume Next) before evaluating the success/error path.
    ' Without this, a stale Err.Number could cause false error propagation on success.
    If succeeded Then Err.Clear
    m_isSendReceiving = False
    
    ' [FIX] If a timeout occurred and auto-reconnect is enabled, clear stale handles.
    ' This ensures that the next attempt starts with a clean slate (Recovery).
    If m_State = StateError And m_AutoReconnect Then
        Call InternalCleanup
    End If
    
    ' Explicit cleanup of potentially large strings
    msg = ""
    inId = ""
    
    ' Error Propagation
    If Not succeeded And Err.Number <> 0 Then
        Dim errNum As Long: errNum = Err.Number
        Dim errDesc As String: errDesc = Err.description
        Dim errSource As String: errSource = Err.Source
        On Error GoTo 0
        Err.Raise errNum, errSource, errDesc
    End If
End Function

' ========================================================================================
' Internal recovery logic with re-entrancy protection.
' ========================================================================================
Private Function AttemptRecovery() As Boolean
    Static isRecovering As Boolean
    
    ' 1. Re-entrancy guard: Acquire the lock IMMEDIATELY.
    ' Moving this to the very top ensures atomic-like locking before any other
    ' evaluations or error handler setups occur.
    If isRecovering Then Exit Function
    isRecovering = True
    
    ' [FIX] Route all runtime errors to the ErrorHandler to guarantee flag reset
    On Error GoTo ErrorHandler

    ' Pre-condition checks now route to CleanExit to properly release the lock
    If Not m_AutoReconnect Then GoTo CleanExit
    If m_RetryCount >= m_MaxRetryAttempts Then
        Trace ">>> [Recovery] Max retry attempts reached.", True
        GoTo CleanExit
    End If

    m_RetryCount = m_RetryCount + 1
    
    Dim waitTime As Long: waitTime = (2 ^ (m_RetryCount - 1)) * 1000
    Trace ">>> [Recovery] Attempt #" & m_RetryCount & " (Waiting " & waitTime & "ms)", True
    
    ' During YieldExecution, DoEvents is pumped (allowDoEvents = True).
    ' External COM events or Timers could attempt re-entrancy here.
    ' Because isRecovering is already True, they will be safely rejected.
    YieldExecution waitTime, True
    
    If Connect(m_LastWsUrl, m_LastPort) Then
        Trace ">>> [Recovery] Success.", True
        m_RetryCount = 0
        AttemptRecovery = True
    Else
        Trace ">>> [Recovery] Failed to reconnect.", True
        AttemptRecovery = False
    End If
    
CleanExit:
    ' [CRITICAL] Always release the recovery lock regardless of success, failure, or runtime errors
    isRecovering = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    ' [CRITICAL 3 FIX] Prohibit Trace (File I/O) within the error handler.
    ' If File I/O fails (e.g., file lock, permissions), a new exception occurs here,
    ' preventing Resume CleanExit from executing and leaving isRecovering=True forever.
    ' Debug.Print is non-throwing and safe in this context.
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [Recovery] Critical Error during attempt: " & Err.description
    
    AttemptRecovery = False
    ' Clear the error state and jump to the cleanup block to release the lock
    Resume CleanExit
End Function

' ========================================================================================
' Non-blocking BiDi message pump; processes queued events and manages heartbeat.
' Optimized to prevent Excel UI freezing by using a micro-deadline for socket I/O.
' ========================================================================================
Public Function GetMessage() As String
    ' Block heartbeat/ping while a command response is pending in SendAndReceive
    If m_isSendReceiving Then Exit Function
    
    ' 1) Serve from local event buffer first
    If m_EventBuffer.count > 0 Then
        GetMessage = m_EventBuffer(1): m_EventBuffer.Remove 1: Exit Function
    End If

    ' 2) Autonomous recovery and connectivity guard
    If m_State = StateError Then If Not AttemptRecovery() Then Exit Function
    If m_State <> StateConnected Then Exit Function

    ' 3) Enforce polling cadence (100ms) to prevent CPU spikes
    Dim nowT As Double: nowT = NowMs()
    If (nowT - p_lastPollTime) < m_PollingInterval Then Exit Function
    p_lastPollTime = nowT

    ' 4) [SMART HEARTBEAT]
    ' Send a ping, but use a static variable to only LOG it every 20 seconds.
    ' This keeps the pipe open at 100ms speed without the console spam.
    Static lastLoggedTime As Double
    Dim isLoggingTime As Boolean
    isLoggingTime = (nowT - lastLoggedTime >= HEARTBEAT_INTERVAL_MS Or lastLoggedTime = 0)
    
    Dim pingCmd As String: pingCmd = "{""id"":" & SYSTEM_PING_ID & ", ""method"":""session.status"", ""params"":{}}"

    If ENABLE_HEARTBEAT_LOGGING And isLoggingTime Then
        lastLoggedTime = nowT
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [HB-REQ] session.status (Keep-alive active)"
    End If

    If Not SendMessage(pingCmd) Then Exit Function

    ' 5) Wait briefly for response
    Dim rawMsg As String: rawMsg = ReadFromWire(nowT + 10)

    If LenB(rawMsg) = 0 Then
        YieldExecution 1, m_AllowUiPump: Exit Function
    End If

    ' 6) Data Classification
    Dim inId As String: inId = ExtractJsonId(rawMsg)
    If inId = CStr(SYSTEM_PING_ID) Then
        ' Only log the response if we logged the request
        If ENABLE_HEARTBEAT_LOGGING And isLoggingTime Then
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [HB-RES] session.status (Connection healthy)"
        End If
        GetMessage = "" ' Swallow heartbeat response
    Else
        ' Return the actual BiDi message (Event/Response)
        GetMessage = rawMsg
    End If
End Function

' ========================================================================================
' Reads a WebSocket frame.
' Refactored to include a Liveness Check (deadlock detection) and strict memory management.
' ========================================================================================
Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    If m_State <> StateConnected And m_State <> StateDisconnecting Then Exit Function
    
    ' --- 1. Liveness Check (Deadlock Detection) ---
    If m_isReading Then
        If (NowMs() - m_isReadingSince) > 5000 Then
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] BiDi Warning: Read lock stuck. Force-resetting."
            m_isReading = False
        Else
            Exit Function
        End If
    End If

    On Error GoTo ErrorHandler

    ' --- 2. Acquire Lock & Record Timestamp ---
    m_isReading = True
    m_isReadingSince = NowMs()

    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim isMessageFinished As Boolean
    Dim hasValidData As Boolean
    ' [FIX] Counter to track fragment throughput
    Dim fragmentCount As Long: fragmentCount = 0

    If m_InStream.Position = 0 Then
        PrepareStream m_InStream, adTypeBinary
    End If
    
    Do
        ' Receive data from the WebSocket
        ret = WinHttpWebSocketReceive(hWebSocket_, _
                                      m_ReceiveBuffer(0), _
                                      m_BufferSize, _
                                      bytesRead, _
                                      bufType)
        
        If ret <> NO_ERROR Then
            If ret <> 12002 Then
                Trace ">>> [WinHttp] Receive Failed. Code: " & ret, True
                SetState StateError
                Disconnect keepErrorState:=True
            End If
            GoTo CleanExit
        End If

        If bytesRead > 0 Then
            hasValidData = True
            
            If bytesRead = m_BufferSize Then
                m_InStream.Write m_ReceiveBuffer
            Else
                Dim chunk() As Byte
                chunk = m_ReceiveBuffer
                ReDim Preserve chunk(bytesRead - 1)
                m_InStream.Write chunk
                Erase chunk
            End If
            
            MaybeGrowBuffer bytesRead
        End If

        ' Frame Assembly Logic
        Select Case bufType
            Case WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                 WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE
                isMessageFinished = True
                GoTo CleanExit

            Case WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE
                ' Server initiated close. Handle is now invalid.
                Debug.Print "[BiDi] WebSocket Close frame received from server."
                SetState StateError
                Disconnect keepErrorState:=True
                GoTo CleanExit
                
            Case WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE, _
                 WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE
                ' [FIX] Prevent Excel freezing during high-volume fragment assembly
                fragmentCount = fragmentCount + 1
                If fragmentCount Mod 20 = 0 Then
                    ' Yield to the OS message pump only if allowed by class settings
                    ' This prevents "Not Responding" while assembling large payloads
                    If m_AllowUiPump And Not m_isSendReceiving Then DoEvents
                    ' [FIX] Removed "fragmentCount = 0" here. The Mod operator naturally
                    ' cycles every 20 iterations. Resetting to 0 caused the next DoEvents
                    ' to fire after 21 fragments instead of 20 (off-by-one).
                End If
                
            Case Else
                GoTo CleanExit
        End Select
        
        ' Overall Deadline Check
        If deadline > 0 And NowMs() > deadline Then
            Trace ">>> [WinHttp] Read loop hit internal deadline.", True
            GoTo CleanExit
        End If

        ' Yield control with allowDoEvents:=False during fragment assembly for precision
        YieldExecution 1, False
    Loop

CleanExit:
    ' --- 3. Assembly & Integrity Check ---
    On Error GoTo StreamErrorHandler

    If isMessageFinished And hasValidData Then
        m_InStream.Position = 0
        m_InStream.Type = adTypeText
        m_InStream.Charset = "UTF-8"
        ReadFromWire = m_InStream.ReadText

        PrepareStream m_InStream, adTypeBinary
        
        Dim isSuspicious As Boolean
        isSuspicious = (InStr(1, ReadFromWire, """error"":", vbTextCompare) > 0) Or _
                       (InStr(1, ReadFromWire, "chrome-error", vbTextCompare) > 0)
        
        If isSuspicious Then m_consecutiveInternalErrors = m_consecutiveInternalErrors + 1
        Trace ">>> [WinHttp] Message assembled. Length: " & Len(ReadFromWire), isSuspicious
    End If
    
    GoTo FinalCleanup

StreamErrorHandler:
    On Error Resume Next
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [WinHttp] Stream Error: " & Err.description
    CloseStream m_InStream
    Set m_InStream = New ADODB.Stream
    m_InStream.mode = adModeReadWrite
    m_InStream.Open
    PrepareStream m_InStream, adTypeBinary
    Resume FinalCleanup

FinalCleanup:
    ' --- 4. Release Lock & Reset Timestamp ---
    m_isReading = False
    m_isReadingSince = 0
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] [WinHttp] Critical Read Error: " & Err.description
    Resume CleanExit
End Function

' ========================================================================================
' Helpers & Utils
' ========================================================================================
' ========================================================================================
' Unified Timer API (QPC-based, monotonic)
' NowTicks   : Opaque Currency snapshot (raw QPC counter). Use as a token for ElapsedMs.
' ElapsedMs  : Difference between a NowTicks snapshot and the current counter, in ms.
' NowMs      : Absolute monotonic timestamp in ms (since QPC origin).
' Note: m_qpcFreq is initialized in Class_Initialize with fail-fast on missing QPC support.
' ========================================================================================
Public Function NowTicks() As Currency
    QueryPerformanceCounter NowTicks
End Function

Public Function ElapsedMs(ByVal startTicks As Currency) As Double
    Dim cur As Currency
    QueryPerformanceCounter cur
    ElapsedMs = (CDbl(cur - startTicks) / CDbl(m_qpcFreq)) * 1000#
End Function

Public Function NowMs() As Double
    Dim cur As Currency
    QueryPerformanceCounter cur
    NowMs = (CDbl(cur) / CDbl(m_qpcFreq)) * 1000#
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
            Else
                ' [FIX] Reset consecutive counter even when MAX_BUFFER_SIZE is reached
                ' to prevent infinite accumulation.
                m_RecentFullReads = 0
            End If
        End If
    Else
        ' [FIX] Reset consecutive counter when a partial read occurs
        m_RecentFullReads = 0
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

' ========================================================================================
' Extracts the root-level "id" value from a JSON string.
' Designed to be robust against variations in key order and large payloads.
' ========================================================================================
Private Function ExtractJsonId(ByVal json As String) As String
    Dim s As String: s = LTrim$(json)
    If Len(s) < 10 Or Left$(s, 1) <> "{" Then Exit Function

    ' Fast pre-check: Avoid running the scanner on payloads without "id"
    If InStr(1, s, """id""", vbTextCompare) = 0 Then Exit Function

    Dim depth As Long: depth = 0
    Dim inString As Boolean: inString = False
    Dim isEscaped As Boolean: isEscaped = False
    Dim i As Long
    
    ' Map the string directly to a Byte array (UTF-16 LE) for blazing fast memory access
    Dim b() As Byte: b = s
    Dim bLen As Long: bLen = UBound(b)
    
    ' Constants for ASCII character codes
    Const CHAR_QUOTE As Byte = 34  ' "
    Const CHAR_LBRACE As Byte = 123 ' {
    Const CHAR_RBRACE As Byte = 125 ' }
    Const CHAR_SLASH As Byte = 92  ' \
    
    ' In VBA, strings are UTF-16, so each character is 2 bytes. We step by 2.
    For i = 0 To bLen Step 2
        ' Check if we are currently escaping a character
        If isEscaped Then
            isEscaped = False
            GoTo ContinueLoop
        End If

        Select Case b(i)
            Case CHAR_SLASH
                ' If inside a string, the next character is escaped
                If inString Then isEscaped = True
                
            Case CHAR_QUOTE
                inString = Not inString
                
                ' If we just OPENED a string at the root level, check if it's "id"
                If inString And depth = 1 Then
                    ' Ensure we don't go out of bounds (need 4 chars: i, d, ", :)
                    If i + 8 <= bLen Then
                        ' Check for 'i' (105) and 'd' (100) and '"' (34)
                        If b(i + 2) = 105 And b(i + 4) = 100 And b(i + 6) = 34 Then
                            
                            Dim j As Long: j = i + 8 ' Start right after "id"
                            Dim idVal As String
                            
                            ' [HIGH FIX] Explicitly check both j and j+1 bounds in every iteration.
                            ' This ensures no reliance on implicit mathematical safety (j <= bLen-1).
                            Do While j + 1 <= bLen
                                ' Validate that the character is strictly within ASCII range
                                ' (Upper byte must be 0 in UTF-16 LE).
                                If b(j + 1) = 0 Then
                                    Select Case b(j)
                                        Case 48 To 57, 45 ' '0'-'9', '-'
                                            idVal = idVal & Chr$(b(j))
                                        Case 32, 58, 9, 10, 13 ' spaces, colon, tabs, CRLF
                                            ' Stop parsing if we hit a delimiter AFTER numbers have started
                                            If Len(idVal) > 0 Then Exit Do
                                        Case Else
                                            ' Hit an invalid ASCII char (e.g., letters, quote, bracket)
                                            Exit Do
                                    End Select
                                Else
                                    ' Non-ASCII character encountered; terminate extraction immediately.
                                    Exit Do
                                End If
                                j = j + 2
                            Loop
                            
                            If Len(idVal) > 0 Then
                                ExtractJsonId = idVal
                                Exit Function
                            End If
                            
                        End If
                    End If
                End If
                
            Case CHAR_LBRACE
                If Not inString Then depth = depth + 1
                
            Case CHAR_RBRACE
                If Not inString Then
                    depth = depth - 1
                    ' Optimization: If depth drops to 0, we have parsed the root object
                    If depth = 0 Then Exit Function
                End If
        End Select
        
ContinueLoop:
    Next i
End Function

Private Function ExtractPath(ByVal url As String) As String
    Dim p As Long: p = InStr(1, url, "://"): If p > 0 Then p = p + 3 Else p = 1
    p = InStr(p, url, "/"): ExtractPath = IIf(p > 0, Mid$(url, p), "/")
End Function

' ========================================================================================
' UTF-8 Conversion Utility
' ========================================================================================
Private Function ToUtf8Bytes(ByVal txt As String) As Byte()
    ' Use the dedicated output stream to avoid corrupting m_InStream
    PrepareStream m_OutStream, adTypeText
    m_OutStream.Charset = "UTF-8"
    m_OutStream.WriteText txt
    
    m_OutStream.Position = 0
    m_OutStream.Type = adTypeBinary
    
    ' Dynamic BOM Detection (EF BB BF)
    If m_OutStream.Size >= 3 Then
        Dim Header() As Byte: Header = m_OutStream.Read(3)
        If Header(0) = &HEF And Header(1) = &HBB And Header(2) = &HBF Then
            m_OutStream.Position = 3 ' Skip BOM
        Else
            m_OutStream.Position = 0
        End If
    Else
        m_OutStream.Position = 0
    End If
    
    If m_OutStream.Size > m_OutStream.Position Then
        ToUtf8Bytes = m_OutStream.Read
    Else
        ToUtf8Bytes = StrConv("", vbFromUnicode)
    End If
End Function

' ========================================================================================
' PumpMessages: Drains the Windows message queue (DoEvents) without sleeping.
'   throttleMs = 0 : pump unconditionally
'   throttleMs > 0 : pump only if at least throttleMs has elapsed since the last pump
' Skipped while m_isSendReceiving is True to prevent WinHttp re-entrancy.
' On the very first call (m_lastPumpTicks = 0), DoEvents always runs.
' ========================================================================================
Public Sub PumpMessages(Optional ByVal throttleMs As Long = 0)
    If m_isSendReceiving Then Exit Sub
    If throttleMs > 0 And m_lastPumpTicks <> 0 Then
        If ElapsedMs(m_lastPumpTicks) < throttleMs Then Exit Sub
    End If
    DoEvents
    m_lastPumpTicks = NowTicks()
End Sub

' ========================================================================================
' YieldExecution: Sleeps for the given duration, optionally pumping the Windows
' message queue. Companion to PumpMessages (which only pumps, no sleep).
'   ms <= 50  : direct SleepWinAPI for precision
'   ms > 50   : delegates to InterruptibleSleep for chunked, responsive waiting
' DoEvents fires only when allowDoEvents=True AND m_isSendReceiving=False
' (re-entrancy guard against WinHttp corruption during active I/O).
' ========================================================================================
Public Sub YieldExecution(Optional ByVal ms As Long = 1, Optional ByVal allowDoEvents As Boolean = False)
    ' Floor Guard: Prevent negative values from causing undefined SleepWinAPI behavior.
    ' This is the only addition; the optimized 50ms branching logic below is preserved as-is.
    Dim vMs As Long: vMs = ms
    If vMs < 0 Then vMs = 1

    ' 1. Use the robust chunked sleep for any duration longer than 50ms.
    ' This ensures that long waits (like in AttemptRecovery) do not freeze Excel.
    If vMs > 50 Then
        Call InterruptibleSleep(vMs, allowDoEvents)
    Else
        ' 2. For very short durations, perform a direct sleep to maintain precision.
        ' Even for short sleeps, we apply the m_isSendReceiving guard to prevent
        ' unintended re-entrancy during active command cycles.
        If allowDoEvents And Not m_isSendReceiving Then
            DoEvents
        End If

        If vMs > 0 Then SleepWinAPI vMs
    End If
End Sub

' Leaves breadcrumbs in both the external file and the Immediate Window.
Public Sub Trace(ByVal msg As String, Optional ByVal isCritical As Boolean = False)
    If Not DEBUG_MODE Then Exit Sub
    
    Dim logMsg As String: logMsg = "[" & Format(Now, "hh:mm:ss") & "] " & msg
    
    ' Only print the message itself to the Immediate Window
    Debug.Print logMsg
    
    ' Detection Logic
    Dim shouldLogToFile As Boolean
    shouldLogToFile = isCritical Or _
                      (InStr(1, msg, "Return Code", vbTextCompare) > 0) Or _
                      (InStr(1, msg, "failed", vbTextCompare) > 0) Or _
                      (InStr(1, msg, "error", vbTextCompare) > 0)

    If shouldLogToFile Then
        Dim localDir As String: localDir = getFullLocalPath(ThisWorkbook.Path)
        If localDir = "" Then localDir = Environ("USERPROFILE") & "\Desktop"
        If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"
        
        Dim logPath As String: logPath = localDir & "bidi_socket_trace.log"
        Dim fNum As Integer: fNum = FreeFile
        
        On Error Resume Next
        Open logPath For Append As #fNum
        Print #fNum, logMsg & " (LOG_EVENT)"
        Close #fNum
        On Error GoTo 0
        ' [SILENT] Removed LOG_SUCCESS print for clean operation
    End If
End Sub

' ========================================================================================
' LogEvent: Unified logging with QPC timestamps and chunked output for long strings.
' ========================================================================================
Public Sub LogEvent(ByVal Header As String, ByVal Body As String, _
                   Optional ByVal Tag As String = "[INFO]", _
                   Optional ByVal LogToFile As Boolean = False)
    Const CHUNK_SIZE As Long = 250
    Dim i As Long
    Dim cleanText As String
    Dim timePrefix As String
    Dim currentTick As Currency
    
    ' High-resolution timestamp for synchronization and performance analysis
    currentTick = Me.NowTicks()
    timePrefix = "[" & Format(Now, "hh:mm:ss") & " | " & Format(currentTick, "0") & "] "

    ' Normalize text: remove null characters and trailing whitespace/newlines
    cleanText = Replace(Body, vbNullChar, "")
    Do While Len(cleanText) > 0
        Dim lastChar As String
        lastChar = Right$(cleanText, 1)
        If lastChar = vbCr Or lastChar = vbLf Then
            cleanText = Left$(cleanText, Len(cleanText) - 1)
        Else
            Exit Do
        End If
    Loop

    ' Output the header with the structured tag.
    ' Format: [Time | Ticks] [TAG] Header (Length: N)
    Dim sfx As String: sfx = "": If Len(cleanText) > 0 Then sfx = " (Length: " & Len(cleanText) & ")"
    Dim headerLine As String: headerLine = timePrefix & Tag & " " & Header & sfx
    Debug.Print headerLine

    ' Chunked output to bypass VBA's Immediate Window line length limits
    If Len(cleanText) > 0 Then
        For i = 1 To Len(cleanText) Step CHUNK_SIZE
            ' Print parts of the string without an automatic newline
            Debug.Print Mid$(cleanText, i, CHUNK_SIZE);
        Next i
        Debug.Print ""
    End If
    
    ' Optional file logging
    If LogToFile Then
        WriteLogToFile headerLine & vbCrLf & cleanText
    End If
End Sub

' Internal helper for LogToFile
Private Sub WriteLogToFile(ByVal msg As String)
    Dim localDir As String
    On Error Resume Next
    localDir = getFullLocalPath(ThisWorkbook.Path)
    If localDir = "" Then localDir = Environ("USERPROFILE") & "\Desktop"
    If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"
    
    Dim logPath As String: logPath = localDir & "bidi_socket_trace.log"
    Dim fNum As Integer: fNum = FreeFile
    
    Open logPath For Append As #fNum
    Print #fNum, msg
    Close #fNum
    On Error GoTo 0
End Sub

' Explicitly resets the consecutive error counter.
' Called by the Wrapper only when a truly healthy response is confirmed.
Public Sub ClearInternalErrorCounter()
    m_consecutiveInternalErrors = 0
End Sub

' ========================================================================================
' Extracts the host name/IP from a URL.
' Designed to be future-proof against IPv6 loopbacks (e.g., [::1]) and custom ports.
' ========================================================================================
Private Function ExtractHost(ByVal url As String) As String
    Dim startPos As Long
    Dim relPos As Long
    Dim authEnd As Long
    Dim qPos As Long
    Dim fPos As Long
    Dim atPos As Long
    Dim portPos As Long
    Dim closeBracket As Long
    Dim authority As String
    Dim hostPart As String

    ' 1. Defensive trimming
    url = Trim$(url)
    If url = "" Then Exit Function

    ' 2. Determine the starting position of the authority part
    ' Check for "://" (standard), "//" (protocol-relative), or none
    startPos = InStr(1, url, "://", vbTextCompare)
    relPos = InStr(1, url, "//")

    If startPos > 0 Then
        startPos = startPos + 3
    ElseIf relPos = 1 Then
        startPos = 3
    Else
        startPos = 1
    End If

    ' 3. Identify the end of the authority part (terminated by '/', '?', or '#')
    authEnd = InStr(startPos, url, "/")
    
    ' Check for Query string indicator (?)
    qPos = InStr(startPos, url, "?")
    If qPos > 0 And (authEnd = 0 Or qPos < authEnd) Then authEnd = qPos
    
    ' Check for Fragment indicator (#)
    fPos = InStr(startPos, url, "#")
    If fPos > 0 And (authEnd = 0 Or fPos < authEnd) Then authEnd = fPos
    
    If authEnd = 0 Then authEnd = Len(url) + 1
    
    ' 4. Extract the authority string (user:pass@host:port)
    authority = Mid$(url, startPos, authEnd - startPos)
    If authority = "" Then Exit Function
    
    ' 5. Strip user credentials if present
    ' Use InStrRev to handle potential '@' characters inside passwords
    atPos = InStrRev(authority, "@")
    If atPos > 0 Then authority = Mid$(authority, atPos + 1)
    
    ' 6. Separate host from port
    If Left$(authority, 1) = "[" Then
        ' Case: IPv6 address [::1]:port
        closeBracket = InStr(1, authority, "]")
        If closeBracket > 0 Then
            hostPart = Mid$(authority, 2, closeBracket - 2)
        Else
            ' Invalid IPv6 format (missing closing bracket)
            Exit Function
        End If
    Else
        ' Case: Standard domain or IPv4 address
        ' Locate the last colon to identify the port number
        portPos = InStrRev(authority, ":")
        If portPos > 0 Then
            hostPart = Left$(authority, portPos - 1)
        Else
            hostPart = authority
        End If
    End If
    
    ' 7. Normalization and final check
    ' Convert host to lowercase for consistency
    If hostPart <> "" Then ExtractHost = LCase$(hostPart)
End Function

' ========================================================================================
' Extracts the origin (scheme + host + port) from a URL string.
' Example: "https://example.com:8080/path" -> "https://example.com:8080"
' Returns empty string for non-parseable or empty URLs.
' ========================================================================================
Public Function ExtractOrigin(ByVal url As String) As String
    If Len(url) = 0 Then Exit Function
    Dim schemeEnd As Long: schemeEnd = InStr(1, url, "://", vbTextCompare)
    If schemeEnd = 0 Then Exit Function
    ' Find the first "/" after "://" which terminates the authority part
    Dim authStart As Long: authStart = schemeEnd + 3
    Dim pathStart As Long: pathStart = InStr(authStart, url, "/")
    If pathStart = 0 Then
        ExtractOrigin = LCase$(url)
    Else
        ExtractOrigin = LCase$(Left$(url, pathStart - 1))
    End If
End Function

' ========================================================================================
' Internal Helper: Performs a non-blocking sleep in small chunks.
' It intelligently yields to the OS (DoEvents) only when it's safe to do so,
' ensuring that the WebSocket state machine is not corrupted by re-entrancy.
' ========================================================================================
Private Sub InterruptibleSleep(ByVal totalMs As Long, Optional ByVal allowDoEvents As Boolean = True)
    Dim startTs As Double: startTs = NowMs()
    Const CHUNK_MS As Long = 50 ' Small granularity for better responsiveness
    
    Do While (NowMs() - startTs) < totalMs
        ' CRITICAL GUARD:
        ' We only call DoEvents if:
        ' 1. The caller explicitly allows it (allowDoEvents).
        ' 2. We are NOT in the middle of a SendAndReceive operation (m_isSendReceiving).
        ' This prevents Excel/Add-ins from intercepting the thread and
        ' corrupting the WinHttp handle state during active I/O.
        If allowDoEvents And Not m_isSendReceiving Then
            DoEvents
        End If
        
        ' Physical sleep to prevent high CPU usage during the wait period.
        SleepWinAPI CHUNK_MS
    Loop
End Sub

' Releases all WinHttp handles without sending a WebSocket close frame.
Public Sub forceClose()
    On Error Resume Next
    ' [FIX] Clear communication locks first, before any handle operations.
    ' forceClose is the emergency exit path (called from Wrapper.Shutdown and
    ' Class_Terminate). If the caller was interrupted mid-SendAndReceive,
    ' these flags must be reset to prevent deadlock on reconnection.
    m_isSendReceiving = False
    m_isReading = False
    m_isReadingSince = 0
    
    ' Skip WinHttpWebSocketClose (no close frame, no timeout).
    ' Directly release all WinHttp handles.
    
    ' [FIX] Track whether the HTTP request handle's ownership was transferred
    ' to the WebSocket via WinHttpWebSocketCompleteUpgrade. After a successful
    ' upgrade, hRequest_ is owned by hWebSocket_ and must NOT be closed separately.
    ' Closing it here would cause a double-free, potentially corrupting WinHttp
    ' internal state and making subsequent reconnection attempts fail silently.
    ' This mirrors the same ownership logic used in InternalCleanup.
    Dim wasUpgraded As Boolean
    wasUpgraded = (hWebSocket_ <> 0)
    
    If hWebSocket_ <> 0 Then
        WinHttpCloseHandle hWebSocket_
        hWebSocket_ = 0
    End If
    If hRequest_ <> 0 Then
        If Not wasUpgraded Then
            WinHttpCloseHandle hRequest_
        End If
        hRequest_ = 0
    End If
    If hConnection_ <> 0 Then
        WinHttpCloseHandle hConnection_
        hConnection_ = 0
    End If
    If hSession_ <> 0 Then
        WinHttpCloseHandle hSession_
        hSession_ = 0
    End If
    Set m_EventBuffer = New Collection
    SetState StateDisconnected
    On Error GoTo 0
End Sub
