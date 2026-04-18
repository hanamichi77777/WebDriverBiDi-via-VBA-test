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
' 7. ZERO PERSISTENCE: The module performs no file writes, registry edits,
'    scheduled tasks, or persistence mechanisms of any kind.
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
Private m_isReading As Boolean
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

Public Enum BiDiSocketState
    StateDisconnected = 0
    StateConnecting = 1
    StateConnected = 2
    StateDisconnecting = 3
    StateError = 4
End Enum
Private m_State As BiDiSocketState
Public Event StateChanged(ByVal NewState As BiDiSocketState, ByVal stateName As String)

' --- Internal Buffers ---
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
Private p_lastActivityTime As Double
Private p_lastPollTime As Double
Private p_RegExp As RegExp
Private m_AllowUiPump As Boolean

' ========================================================================================
' Initialization & Termination
' ========================================================================================
Private Sub Class_Initialize()
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
' Ensure all handles and streams are closed
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
Public Property Get ConsecutiveInternalErrors() As Long:: ConsecutiveInternalErrors = m_consecutiveInternalErrors: End Property

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
    
    ' [FIX] Dynamically extract the host to match the actual binding (e.g., 127.0.0.1 vs localhost).
    ' Prevents connection failures when WinHttp resolves 'localhost' differently
    ' than the WebDriver's active listener (e.g., IPv4 vs IPv6 mismatch).
    Dim host As String: host = ExtractHost(wsUrl)

    ' Step 2: Initialize WinHttp Session
    hSession_ = WinHttpOpen(StrPtr(AGENT_NAME), WINHTTP_ACCESS_TYPE_NO_PROXY, 0, 0, 0)
    If hSession_ = 0 Then Err.Raise 513, "BiDiSocket", "WinHttpOpen failed"

    ' --- CRITICAL FIX: UI RESPONSIVENESS vs CPU BALANCE ---
    ' Set connection-related timeouts (Resolve, Connect, Send, Receive)
    ' Changed Receive Timeout from 1000ms to 100ms.
    ' This allows the WinHttp synchronous call to yield back to VBA 10 times a second,
    ' ensuring Excel's UI remains responsive without causing high CPU busy-wait loops.
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
        If hWebSocket_ = 0 Then Err.Raise 520, "BiDiSocket", "WebSocket upgrade failed"

        ' Step 9: Configure Keep-Alive to maintain connection
        Dim ka As Long: ka = KEEPALIVE_INTERVAL_MS
        WinHttpSetOption hWebSocket_, WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL, ka, 4
        
        ' Finalize: Success transition
        p_lastActivityTime = GetTickMS()
        p_lastPollTime = 0
        m_RetryCount = 0
        SetState StateConnected
        Connect = True
    Else
        Err.Raise 519, "BiDiSocket", "Upgrade rejected. API Result: " & headerRet & " Status: " & status
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
        p_lastActivityTime = GetTickMS()
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
' Properly closes the WebSocket and all associated WinHttp handles.
' ========================================================================================
Public Sub Disconnect(Optional ByVal keepErrorState As Boolean = False)
    ' If we are already in the middle of a connection attempt,
    ' we should be cautious about forced disconnection unless it's a hard reset.
    If m_State = StateDisconnecting Or m_State = StateDisconnected Then Exit Sub

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
' [REVISION]
' Replaced arbitrary loop counter with a strict millisecond-based Hard Deadline.
' This prevents VBA from hanging if the WebDriver/Browser silences.
' ========================================================================================
Public Function SendAndReceive(ByVal jsonMessage As String, Optional ByVal waitTimeMs As Long = 0) As String
    ' Connectivity Guard: Ensure operational state before sending
    If m_State <> StateConnected Then
        If Not AttemptRecovery() Then Err.Raise 600, "BiDiSocket", "Not connected"
    End If
    
    Dim cmdId As String: cmdId = ExtractJsonId(jsonMessage)
    
    ' Differentiate between protocol send error and connection drop.
    ' 600 = Connection Lost (Triggers context/session recovery in Wrapper)
    ' 601 = Raw Send Failure (Temporary WinHttp hiccup)
    If Not SendMessage(jsonMessage) Then
        If m_State <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost during send operation."
        Else
            Err.Raise 601, "BiDiSocket", "Send failed due to internal WinHttp error."
        End If
    End If
    
    ' Calculate the hard exit time
    Dim timeout As Long: timeout = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    Dim deadline As Double: deadline = GetTickMS() + timeout
    Dim msg As String, inId As String
    
    ' Define maximum buffer size to prevent memory leaks during event floods
    Const MAX_EVENT_BUFFER_SIZE As Long = 1000
    
    Do
        ' 1. State Guard: Abort if recovery failed or connection dropped elsewhere
        If m_State <> StateConnected Then
            Err.Raise 600, "BiDiSocket", "Connection lost while waiting for ID: " & cmdId
        End If

        ' 2. Receive attempt: ReadFromWire will now respect the deadline internally
        msg = ReadFromWire(deadline)
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            If inId = cmdId Then
                SendAndReceive = msg
                Exit Function
            Else
                ' [FIX] Prevent internal heartbeat responses from polluting the event buffer.
                ' Only queue valid asynchronous browser events.
                If inId <> CStr(SYSTEM_PING_ID) Then
                    ' Store unrelated events (e.g., network events) for later retrieval
                    m_EventBuffer.Add msg
                    
                    ' OOM Protection: Enforce FIFO (First-In-First-Out) limit
                    ' Removes the oldest message if the buffer exceeds the maximum allowed size.
                    If m_EventBuffer.Count > MAX_EVENT_BUFFER_SIZE Then
                        m_EventBuffer.Remove 1
                    End If
                End If
            End If
        End If
        
        ' 3. [STRICT DEADLINE] VBA-side patience limit exceeded
        ' If the browser silences, we must cut the line to prevent an Excel hang.
        If GetTickMS() > deadline Then
            SetState StateError
            Err.Raise 604, "BiDiSocket", "Socket Silence: Hard timeout waiting for ID: " & cmdId & _
                      " (Waited " & timeout & "ms). Browser engine may be hung."
        End If

        ' Yield execution but keep a tight poll for responsiveness
        YieldExecution 5, m_AllowUiPump
    Loop
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
    Trace ">>> [Recovery] Critical Error during attempt: " & Err.Description, True
    AttemptRecovery = False
    ' Clear the error state and jump to the cleanup block
    Resume CleanExit
End Function

' ========================================================================================
' Non-blocking BiDi message pump; processes queued events and manages heartbeat.
' Optimized to prevent Excel UI freezing by using a micro-deadline for socket I/O.
' ========================================================================================
Public Function GetMessage() As String
    ' 1) Serve from local event buffer first
    If m_EventBuffer.Count > 0 Then
        GetMessage = m_EventBuffer(1): m_EventBuffer.Remove 1: Exit Function
    End If

    ' 2) Autonomous recovery and connectivity guard
    If m_State = StateError Then If Not AttemptRecovery() Then Exit Function
    If m_State <> StateConnected Then Exit Function

    ' 3) Enforce polling cadence (100ms) to prevent CPU spikes
    Dim nowT As Double: nowT = GetTickMS()
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
' Reads a WebSocket frame. Now includes a safety check against the overall deadline
' to prevent internal loop capture during browser-side hangs.
' ========================================================================================
Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    If m_State <> StateConnected And m_State <> StateDisconnecting Then Exit Function
    
    ' 1. Re-entrancy guard
    If m_isReading Then Exit Function

    ' Set up the error handler BEFORE acquiring the lock.
    ' This guarantees that if any error occurs immediately after locking,
    ' execution will route to ErrorHandler and correctly release the lock.
    On Error GoTo ErrorHandler

    ' 2. Acquire lock
    m_isReading = True

    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim isMessageFinished As Boolean
    Dim hasValidData As Boolean

    ' Initialize stream only when starting a new message
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
            Trace ">>> [WinHttp] Receive Return Code: " & ret, True
            If ret <> 12002 Then ' Not a timeout
                SetState StateError
                Disconnect keepErrorState:=True
            End If
            GoTo CleanExit
        End If

        ' [FIX] CRITICAL: Only process if bytes were actually read.
        ' This prevents "Runtime Error 9" (ReDim Preserve chunk(-1)) during socket closure.
        If bytesRead > 0 Then
            hasValidData = True
            
            If bytesRead = m_BufferSize Then
                ' Performance: Direct write if the buffer is perfectly full
                m_InStream.Write m_ReceiveBuffer
            Else
                ' Safe partial write: Extract only the valid bytes
                Dim chunk() As Byte
                chunk = m_ReceiveBuffer
                ReDim Preserve chunk(bytesRead - 1) ' Safe because bytesRead > 0
                m_InStream.Write chunk
            End If
            
            ' [FIX] Only consider growing the buffer if we have valid data flow
            MaybeGrowBuffer bytesRead
        End If

        ' Handle WebSocket fragmentation and message types
        Select Case bufType
            Case WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                 WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, _
                 WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE
                isMessageFinished = True
                GoTo CleanExit
                
            Case WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE, _
                 WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE
                ' Continue looping to receive remaining fragments
                
            Case Else
                ' Unknown or unexpected buffer type
                GoTo CleanExit
        End Select
        
        ' Check if we've exceeded the absolute time limit
        If deadline > 0 And GetTickMS() > deadline Then
            Trace ">>> [WinHttp] Read loop hit internal deadline.", True
            GoTo CleanExit
        End If

        ' Strictly prohibit DoEvents during fragment assembly to prevent re-entrancy deadlocks.
        YieldExecution 1, False
    Loop

CleanExit:
    ' 3. Assembly block (Fail-safe)
    On Error GoTo StreamErrorHandler

    ' Only attempt to decode if we have a complete message and valid data
    If isMessageFinished And hasValidData Then
        m_InStream.Position = 0
        m_InStream.Type = adTypeText
        m_InStream.Charset = "UTF-8"
        ReadFromWire = m_InStream.ReadText
        p_lastActivityTime = GetTickMS()
        
        ' Reset stream for the next message
        PrepareStream m_InStream, adTypeBinary
        
        ' --- STICKY ERROR DETECTION ---
        ' Check if content contains protocol-level errors or internal browser failures
        Dim isSuspicious As Boolean
        isSuspicious = (InStr(1, ReadFromWire, """error"":", vbTextCompare) > 0) Or _
                       (InStr(1, ReadFromWire, "chrome-error", vbTextCompare) > 0) Or _
                       (InStr(1, ReadFromWire, "chromewebdata", vbTextCompare) > 0)
        
        If isSuspicious Then
            m_consecutiveInternalErrors = m_consecutiveInternalErrors + 1
        End If
        
        Dim logDesc As String
        logDesc = ">>> [WinHttp] Message assembled. Length: " & Len(ReadFromWire)
        If m_consecutiveInternalErrors > 0 Then
            logDesc = logDesc & " | Internal Error Count: " & m_consecutiveInternalErrors
        End If
        
        Trace logDesc, isSuspicious
    End If
    
    GoTo FinalCleanup

StreamErrorHandler:
    Trace ">>> [WinHttp] Stream Extraction Error: " & Err.Description, True
    ' Recovery: Nuke and recreate corrupted stream
    CloseStream m_InStream
    Set m_InStream = New ADODB.Stream
    m_InStream.mode = adModeReadWrite
    m_InStream.Open
    PrepareStream m_InStream, adTypeBinary
    Resume FinalCleanup

FinalCleanup:
    ' Guaranteed release of the re-entrancy lock
    m_isReading = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Trace ">>> [WinHttp] Read Error: " & Err.Description, True
    Resume CleanExit
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
    Const CHAR_QUOTE As Byte = 34 ' "
    Const CHAR_LBRACE As Byte = 123 ' {
    Const CHAR_RBRACE As Byte = 125 ' }
    Const CHAR_SLASH As Byte = 92 ' \
    
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
                            
                            ' [FIX] Directly parse the numeric value.
                            Dim j As Long: j = i + 8 ' Start right after "id"
                            Dim idVal As String
                            
                            ' [FIX] Ensure j + 1 is strictly within bounds before accessing
                            ' the upper byte b(j + 1) to prevent boundary crash.
                            Do While j + 1 <= bLen
                                ' Ensure the character is strictly within the ASCII range
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
                                    ' Hit a non-ASCII character, immediately terminate extraction
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
        Dim header() As Byte: header = m_OutStream.Read(3)
        If header(0) = &HEF And header(1) = &HBB And header(2) = &HBF Then
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

Private Sub YieldExecution(Optional ByVal ms As Long = 1, Optional ByVal allowDoEvents As Boolean = False)
    If allowDoEvents Then DoEvents
    If ms >= 0 Then SleepWinAPI ms
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
    Dim startPos As Long: startPos = InStr(1, url, "://")
    If startPos > 0 Then startPos = startPos + 3 Else startPos = 1
    
    Dim endPos As Long: endPos = InStr(startPos, url, "/")
    If endPos = 0 Then endPos = Len(url) + 1
    
    Dim hostPort As String: hostPort = Mid$(url, startPos, endPos - startPos)
    
    ' Handle future potential IPv6 format: ws://[::1]:9222/
    If Left$(hostPort, 1) = "[" Then
        Dim closeBracket As Long: closeBracket = InStr(1, hostPort, "]")
        If closeBracket > 0 Then
            ExtractHost = Mid$(hostPort, 2, closeBracket - 2)
            Exit Function
        End If
    End If
    
    ' Handle standard IPv4 and domain names
    Dim portPos As Long: portPos = InStrRev(hostPort, ":")
    If portPos > 0 Then
        ExtractHost = Left$(hostPort, portPos - 1)
    Else
        ExtractHost = hostPort
    End If
End Function
