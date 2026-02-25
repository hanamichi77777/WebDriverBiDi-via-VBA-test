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
'   - Adaptive Polling & Autonomous Heartbeat.
'   - Reentrancy-safe Stream Management (Separate Read/Write Buffers).
'   - Strict Deadline-based Timeout Management.
'   - Peak-holding Buffer Growth.
'
' NOTE ON ERROR HANDLING STRATEGY (Strategic Buffer):
' This class utilizes "On Error Resume Next" (OERN) in low-level I/O operations.
' [Rationale]
' Real-time WebSocket streams can be interrupted by browser-side context changes.
' Strict error trapping during frame reassembly can cause permanent communication
' breaks (Error 601) when the driver is in a transitional state.
' [Implementation]
' OERN is used to absorb mid-stream fragmentation and race conditions during
' navigation. This "Optimistic Concurrency" approach prevents the socket from
' saturating with redundant recovery requests, maintaining a clean pipe between
' the VBA host and the WebDriver instance.
' ========================================================================================

' ========================================================================================
' Class Name : BiDiSocketCommunicator
' Purpose    : Professional WebSocket Client for WebDriver BiDi.
' Features   :
'    - Adaptive Polling & Autonomous Heartbeat.
'    - Reentrancy-safe Stream Management (Separate Read/Write Buffers).
'    - Strict Deadline-based Timeout Management.
'    - Peak-holding Buffer Growth.
' ========================================================================================
Private Const DEBUG_MODE As Boolean = False

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

' --- Standard WinHttp Constants ---
Private Const WINHTTP_ACCESS_TYPE_NO_PROXY As Long = 1
Private Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET As Long = 114
Private Const WINHTTP_OPTION_WEB_SOCKET_KEEPALIVE_INTERVAL As Long = 116
Private Const WINHTTP_QUERY_STATUS_CODE As Long = 19
Private Const WINHTTP_QUERY_FLAG_NUMBER As Long = &H20000000
Private Const WINHTTP_HEADER_NAME_BY_INDEX As Long = 0
Private Const WINHTTP_NO_HEADER_INDEX As Long = 0

' --- WebSocket Buffer Types (Standard WinHttp Definitions) ---
Private Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE  As Long = 0
Private Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE As Long = 1
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE    As Long = 2
Private Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE   As Long = 3
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE           As Long = 4

' --- Custom Library Constants ---
Private Const NO_ERROR As Long = 0
Private Const AGENT_NAME As String = "SeleniumVBA_BiDi_Client"
Private Const SYSTEM_PING_ID As Long = -999

' --- Configuration Defaults ---
Private Const DEFAULT_BUFFER_SIZE As Long = 8192
Private Const MAX_BUFFER_SIZE As Long = 10485760 ' 10MB to handle large DOM/BiDi responses
Private Const DEFAULT_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_POLLING_INTERVAL_MS As Long = 1000
Private Const KEEPALIVE_INTERVAL_MS As Long = 15000 ' 15,000 ms Å‡ 15 seconds

' --- Reentrancy Protection Flag ---
Private m_isReading As Boolean

' --- WinHttp Handles ---
Private hSession_ As LongPtr
Private hConnection_ As LongPtr
Private hRequest_ As LongPtr
Private hWebSocket_ As LongPtr

' --- Reconnection Configuration ---
Private m_LastWsUrl As String
Private m_LastPort As Long
Private m_RetryCount As Long
Private m_MaxRetryAttempts As Long
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

' --- Internal Buffers & Utilities ---
Private m_EventBuffer As Collection
Private m_DefaultTimeout As Long
Private m_PollingInterval As Long

' Persistent buffers to minimize heap fragmentation
Private m_BufferSize As Long
Private m_MaxBufferSize As Long
Private m_ReceiveBuffer() As Byte
Private m_GrowThresholdCount As Long
Private m_RecentFullReads As Long

' REENTRANCY PROTECTION: Separate streams for Read and Write
Private m_InStream As ADODB.Stream  ' Dedicated to ReadFromWire
Private m_OutStream As ADODB.Stream ' Dedicated to ToUtf8Bytes

Private p_lastActivityTime As Double
Private p_lastPollTime As Double
Private p_RegExp As RegExp
Private m_AllowUiPump As Boolean

' Enable/disable app heartbeat logging (HB-REQ/HB-RES)
Private Const ENABLE_HEARTBEAT_LOGGING As Boolean = True

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
    If Not m_InStream Is Nothing Then If m_InStream.state = adStateOpen Then m_InStream.Close: Set m_InStream = Nothing
    If Not m_OutStream Is Nothing Then If m_OutStream.state = adStateOpen Then m_OutStream.Close: Set m_OutStream = Nothing
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
    Call InternalCleanup
    
    ' Basic parameter validation
    If wsUrl = "" Or localPort = 0 Then
        If DEBUG_MODE Then Debug.Print "BiDi: Connect failed. Invalid URL or Port."
        GoTo ErrorHandler
    End If
    
    Dim path As String: path = ExtractPath(wsUrl)
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
    hRequest_ = WinHttpOpenRequest(hConnection_, StrPtr("GET"), StrPtr(path), StrPtr("HTTP/1.1"), 0, 0, 0)
    If hRequest_ = 0 Then Err.Raise 515, "BiDiSocket", "WinHttpOpenRequest failed"

    ' Step 5: Configure the request for WebSocket upgrade
    If WinHttpSetOption(hRequest_, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, ByVal 0&, 0) = 0 Then
        Err.Raise 516, "BiDiSocket", "Failed to set WebSocket upgrade option"
    End If

    ' Step 6: Send the upgrade request and wait for the response
    If WinHttpSendRequest(hRequest_, 0, 0, 0, 0, 0, 0) = 0 Then Err.Raise 517, "BiDiSocket", "SendRequest failed"
    If WinHttpReceiveResponse(hRequest_, 0) = 0 Then Err.Raise 518, "BiDiSocket", "ReceiveResponse failed"

    ' Step 7: Verify Status 101 (Switching Protocols)
    Dim status As Long, Size As Long: Size = 4
    Dim idx As Long: idx = WINHTTP_NO_HEADER_INDEX
    WinHttpQueryHeaders hRequest_, _
                        WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER, _
                        WINHTTP_HEADER_NAME_BY_INDEX, _
                        status, Size, idx
    
    If status = 101 Then
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
' ========================================================================================
Public Function SendMessage(ByVal msg As String) As Boolean
    ' [FIX] Early guard for empty payload (prevents UBound on empty byte array)
    If LenB(msg) = 0 Then
        If DEBUG_MODE Then Debug.Print "BiDi: SendMessage aborted. Empty payload."
        Exit Function
    End If

    ' 1. Initial State Check: If in error, try to restore connection before sending
    If m_State = StateError Then
        If Not AttemptRecovery() Then Exit Function
    End If

    ' 2. Connectivity Guard
    If m_State <> StateConnected Then
        If DEBUG_MODE Then Debug.Print "BiDi: SendMessage aborted. Not connected."
        Exit Function
    End If

    ' 3. Data Conversion: VBA String (UTF-16) to Byte array (UTF-8)
    Dim data() As Byte: data = ToUtf8Bytes(msg)
    Dim dataSize As Long: dataSize = UBound(data) + 1
    
    ' 4. Primary Transmission Attempt
    ' Passing the first element of the byte array directly to ByRef pvBuffer As Any.
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

    ' 5. Failure Handling: If the first attempt failed, trigger recovery
    If DEBUG_MODE Then Debug.Print "BiDi: SendMessage failed (Error: " & ret & "). Attempting recovery..."
    
    ' Mark state as Error and clean up handles (preserving the error state for recovery)
    SetState StateError
    Disconnect keepErrorState:=True
    
    ' 6. Recovery & Final Retry Attempt
    If AttemptRecovery() Then
        ' Re-attempt sending using the same data buffer.
        ret = WinHttpWebSocketSend(hWebSocket_, _
                                   WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                                   data(0), _
                                   dataSize)
        
        If ret = NO_ERROR Then
            If DEBUG_MODE Then Debug.Print "BiDi: SendMessage succeeded after recovery."
            SendMessage = True
            p_lastActivityTime = GetTickMS()
        Else
            If DEBUG_MODE Then Debug.Print "BiDi: SendMessage failed again after recovery (Error: " & ret & ")."
            SetState StateError
            Disconnect keepErrorState:=True
        End If
    End If
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
        WinHttpCloseHandle hRequest_
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
    If Not SendMessage(jsonMessage) Then Err.Raise 601, "BiDiSocket", "Send failed"
    
    ' Calculate the hard exit time
    Dim timeout As Long: timeout = IIf(waitTimeMs > 0, waitTimeMs, m_DefaultTimeout)
    Dim deadline As Double: deadline = GetTickMS() + timeout
    Dim msg As String, inId As String
    
    Do
        ' 1. State Guard: Abort if recovery failed or connection dropped elsewhere
        If m_State <> StateConnected Then
            Err.Raise 605, "BiDiSocket", "Connection lost while waiting for ID: " & cmdId
        End If

        ' 2. Receive attempt: ReadFromWire will now respect the deadline internally
        msg = ReadFromWire(deadline)
        
        If Len(msg) > 0 Then
            inId = ExtractJsonId(msg)
            If inId = cmdId Then
                SendAndReceive = msg
                Exit Function
            Else
                ' Store unrelated events (e.g., network events) for later retrieval
                m_EventBuffer.Add msg
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
    ' Guard: Prevent multiple recovery attempts at the same time
    Static isRecovering As Boolean
    If isRecovering Then Exit Function
    
    If Not m_AutoReconnect Then Exit Function
    If m_RetryCount >= m_MaxRetryAttempts Then
        If DEBUG_MODE Then Debug.Print "BiDi: Max recovery attempts reached."
        Exit Function
    End If

    isRecovering = True
    m_RetryCount = m_RetryCount + 1
    
    Dim waitTime As Long: waitTime = (2 ^ (m_RetryCount - 1)) * 1000
    If DEBUG_MODE Then Debug.Print "BiDi: Recovery attempt " & m_RetryCount & " (Waiting " & waitTime & "ms)"
    
    ' Yield while keeping the UI responsive
    YieldExecution waitTime, True
    
    ' Connect handles the state transition to Connected or Error internally
    If Connect(m_LastWsUrl, m_LastPort) Then
        If DEBUG_MODE Then Debug.Print "BiDi: Recovery successful."
        m_RetryCount = 0
        AttemptRecovery = True
    Else
        AttemptRecovery = False
    End If
    
    isRecovering = False
End Function

' ========================================================================================
' Non-blocking BiDi message pump; returns queued events and runs a periodic
' heartbeat to confirm socket liveness.
' ========================================================================================
Public Function GetMessage() As String
    ' 1) Serve from local event buffer first (no socket I/O required).
    If m_EventBuffer.Count > 0 Then
        GetMessage = m_EventBuffer(1)
        m_EventBuffer.Remove 1
        Exit Function
    End If

    ' 2) If we recently hit an error, attempt an auto-recovery before polling.
    If m_State = StateError Then
        If Not AttemptRecovery() Then Exit Function
    End If

    ' 3) Only proceed when connected.
    If m_State <> StateConnected Then Exit Function

    ' 4) Respect the polling cadence (lightweight app heartbeat frequency).
    Dim nowT As Double: nowT = GetTickMS()
    If (nowT - p_lastPollTime) < m_PollingInterval Then Exit Function
    p_lastPollTime = nowT

    ' 5) Build a tiny heartbeat ping to keep the pipe exercised.
    Dim pingCmd As String
    pingCmd = "{""id"":" & SYSTEM_PING_ID & ", ""method"":""session.status"", ""params"":{}}"

    ' Optional logging: heartbeat request
    If ENABLE_HEARTBEAT_LOGGING Then
        Debug.Print "[" & Format(Now, "hh:nn:ss") & " | " & Format(GetTickMS, "0") & _
                    "] [HB-REQ] session.status (id=" & SYSTEM_PING_ID & ")"
    End If

    ' Transmit the heartbeat. If it fails, state will flip to Error inside SendMessage.
    If Not SendMessage(pingCmd) Then Exit Function

    ' 6) Read response within a short deadline; do not block the UI.
    Dim rawMsg As String
    rawMsg = ReadFromWire(GetTickMS() + 500)  ' ~0.5s micro-deadline

    ' Early return when no data (brief yield keeps UI smooth)
    If LenB(rawMsg) = 0 Then
        YieldExecution 1, m_AllowUiPump
        Exit Function
    End If

    ' 7) Classify: heartbeat vs non-heartbeat.
    Dim inId As String
    inId = ExtractJsonId(rawMsg)

    If inId = CStr(SYSTEM_PING_ID) Then
        If ENABLE_HEARTBEAT_LOGGING Then
            Debug.Print "[" & Format(Now, "hh:nn:ss") & " | " & Format(GetTickMS, "0") & _
                        "] [HB-RES] session.status (id=" & SYSTEM_PING_ID & ")"
        End If
        GetMessage = ""   ' swallow heartbeat reply
    Else
        GetMessage = rawMsg
    End If
End Function

' ========================================================================================
' Reads a WebSocket frame. Now includes a safety check against the overall deadline
' to prevent internal loop capture during browser-side hangs.
' ========================================================================================
Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    If m_State <> StateConnected And m_State <> StateDisconnecting Then Exit Function
    If m_isReading Then Exit Function

    m_isReading = True
    On Error Resume Next ' Optimistic: absorb mid-stream fragmentation

    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim hasData As Boolean, isMessageFinished As Boolean
    Dim transientRetries As Long: transientRetries = 0

    PrepareStream m_InStream, adTypeBinary

    Do
        ' Blocking call: WinHttp receives a fragment
        ret = WinHttpWebSocketReceive(hWebSocket_, _
                                      m_ReceiveBuffer(0), _
                                      m_BufferSize, _
                                      bytesRead, _
                                      bufType)
        
        If ret <> NO_ERROR Then
            transientRetries = transientRetries + 1
            If transientRetries <= 3 Then
                SleepWinAPI 2
                If deadline > 0 And GetTickMS() > deadline Then Exit Do
                GoTo ContinueLoop
            End If
            
            ' Hard failure: Terminate connection on persistent I/O error
            On Error GoTo 0
            SetState StateError
            Disconnect keepErrorState:=True
            Exit Do
        End If

        ' Handle server-initiated close frame
        If bufType = WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE Then
            On Error GoTo 0
            Call Disconnect
            Exit Do
        End If

        ' Assemble fragments into the stream
        If bytesRead > 0 Then
            hasData = True
            If bytesRead = m_BufferSize Then
                m_InStream.Write m_ReceiveBuffer
            Else
                Dim chunk() As Byte: ReDim chunk(bytesRead - 1)
                Dim i As Long: For i = 0 To bytesRead - 1: chunk(i) = m_ReceiveBuffer(i): Next i
                m_InStream.Write chunk
            End If
            MaybeGrowBuffer bytesRead
        End If

        ' Exit condition: Message fully reassembled
        If bufType = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Or _
           bufType = WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE Then
            isMessageFinished = True: Exit Do
        End If
        
        ' Safety Abort
        If deadline > 0 Then
            If GetTickMS() > deadline Then Exit Do
        End If

ContinueLoop:
        SleepWinAPI 1
    Loop

    On Error GoTo 0
    If isMessageFinished And hasData Then
        m_InStream.Position = 0
        m_InStream.Type = adTypeText
        m_InStream.Charset = "UTF-8"
        ReadFromWire = m_InStream.ReadText
        p_lastActivityTime = GetTickMS()
    End If

    m_isReading = False
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

' ========================================================================================
' Extracts the root-level "id" value from a JSON string.
' Designed to be robust against variations in key order and large payloads.
' ========================================================================================
Private Function ExtractJsonId(ByVal json As String) As String
    Dim s As String: s = LTrim$(json)
    If Len(s) < 10 Or Left$(s, 1) <> "{" Then Exit Function

    ' Scan boundaries to handle multi-MB responses efficiently
    Dim searchArea As String
    If Len(s) > 2048 Then searchArea = Left$(s, 1024) & Right$(s, 1024) Else searchArea = s

    ' Pattern matches "id" only at the root level (start of object or after a comma)
    p_RegExp.Pattern = "(?:^\{|,)\s*""id""\s*:\s*(-?\d+)(?=[,}])"
    
    Dim matches As MatchCollection
    Set matches = p_RegExp.execute(searchArea)
    
    If matches.Count > 0 Then
        ExtractJsonId = matches(0).SubMatches(0)
    End If
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
