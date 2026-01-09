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
Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal pvBuffer As LongPtr, ByVal dwBufferLength As Long, ByRef pdwBytesRead As Long, ByRef peBufferType As Long) As Long
Private Declare PtrSafe Function WinHttpWebSocketClose Lib "WinHttp" (ByVal hWebSocket As LongPtr, ByVal usStatus As Long, ByVal pvReason As LongPtr, ByVal dwReasonLength As Long) As Long

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
Private Const KEEPALIVE_INTERVAL_MS As Long = 15000

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
    
    ' Initialize specialized streams to prevent reentrancy collisions
    Set m_InStream = New ADODB.Stream
    m_InStream.mode = adModeReadWrite
    m_InStream.Open

    Set m_OutStream = New ADODB.Stream
    m_OutStream.mode = adModeReadWrite
    m_OutStream.Open
    
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
    CloseStream m_InStream
    CloseStream m_OutStream
    Set p_RegExp = Nothing
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
    ' This is internal cleanup and doesn't trigger a state-change event loop
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
' Optimized for 64-bit pointer safety with built-in linear recovery logic.
' ========================================================================================
Public Function SendMessage(ByVal msg As String) As Boolean
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
    
    ' 4. [FIX] Pointer Acquisition: Explicitly store address as LongPtr
    ' Prevents access violations in 64-bit environments by ensuring the full
    ' 64-bit address is correctly passed to the WinHttp API.
    Dim pBuffer As LongPtr
    pBuffer = VarPtr(data(LBound(data)))
    
    ' 5. Primary Transmission Attempt
    Dim ret As Long
    ret = WinHttpWebSocketSend(hWebSocket_, _
                               WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                               pBuffer, _
                               dataSize)
    
    If ret = NO_ERROR Then
        SendMessage = True
        p_lastActivityTime = GetTickMS()
        Exit Function
    End If

    ' 6. Failure Handling: If the first attempt failed, trigger recovery
    If DEBUG_MODE Then Debug.Print "BiDi: SendMessage failed (Error: " & ret & "). Attempting recovery..."
    
    ' Mark state as Error and clean up handles (preserving the error state for recovery)
    SetState StateError
    Disconnect keepErrorState:=True
    
    ' 7. Recovery & Final Retry Attempt
    If AttemptRecovery() Then
        ' Re-attempt sending using the same data buffer address
        ' AttemptRecovery() ensures hWebSocket_ is updated with a fresh handle.
        ret = WinHttpWebSocketSend(hWebSocket_, _
                                   WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                                   pBuffer, _
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
' State management is the responsibility of the caller (Connect/Disconnect/etc.)
' to avoid race conditions during recovery flows.
' ========================================================================================
Private Sub InternalCleanup()
    On Error Resume Next
    
    ' 1. Close WebSocket Handle
    If hWebSocket_ <> 0 Then
        ' Attempt to send a close frame (Status 1000: Normal Closure).
        ' We don't wait for the server's response here to ensure the cleanup
        ' remains fast and non-blocking during error or disconnect sequences.
        WinHttpWebSocketClose hWebSocket_, 1000, 0, 0
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
' Sends a JSON message and waits for a specific response ID.
' [ARCHITECTURE NOTE]
' This is a SYNCHRONOUS blocking call. While it uses DoEvents to mitigate UI freezing,
' heavy payloads or slow driver responses will cause UI stuttering.
' Avoid calling this during intensive UI interactions or within tight loops
' where high UI responsiveness is required.
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
' Reads incoming WebSocket frames and assembles them into a complete UTF-8 string.
' Features:
'   - Reentrancy waiting (prevents data loss during concurrent calls).
'   - 64-bit Pointer safety (uses LongPtr for memory addressing).
'   - Fragment reassembly (loops until a full MESSAGE type is received).
'   - Atomic decoding (discards partial data on deadline to prevent UTF-8 corruption).
' ========================================================================================
Private Function ReadFromWire(Optional ByVal deadline As Double = 0) As String
    ' 1. Connectivity Check
    If m_State <> StateConnected And m_State <> StateDisconnecting Then Exit Function

    ' 2. Reentrancy Guard (Wait-and-Retry logic)
    ' If another process is already reading, wait up to 2 seconds to maintain
    ' message order and prevent data loss.
    Dim waitStart As Double: waitStart = GetTickMS()
    Do While m_isReading
        DoEvents
        SleepWinAPI 1
        If GetTickMS() - waitStart > 2000 Then
            If DEBUG_MODE Then Debug.Print "BiDi: ReadFromWire reentrancy timeout."
            Exit Function
        End If
    Loop

    m_isReading = True
    On Error GoTo Cleanup

    Dim ret As Long, bytesRead As Long, bufType As Long
    Dim hasData As Boolean: hasData = False
    Dim isMessageFinished As Boolean: isMessageFinished = False
    
    ' 3. Pointer Acquisition for 64-bit Memory Safety
    ' Use LongPtr to ensure the buffer address is not truncated in 64-bit environments.
    Dim pReceiveBuffer As LongPtr
    pReceiveBuffer = VarPtr(m_ReceiveBuffer(0))

    ' Initialize stream to accumulate binary fragments
    PrepareStream m_InStream, adTypeBinary

    ' 4. Frame Receiving Loop
    Do
        ' Receive a single frame using the secured memory pointer
        ret = WinHttpWebSocketReceive(hWebSocket_, pReceiveBuffer, m_BufferSize, bytesRead, bufType)
        
        If ret <> NO_ERROR Then
            If DEBUG_MODE Then Debug.Print "BiDi: WinHttpWebSocketReceive failed (Error: " & ret & ")"
            SetState StateError
            Disconnect keepErrorState:=True
            GoTo Cleanup
        End If

        ' Handle Remote Close Signal
        If bufType = WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE Then
            Disconnect
            GoTo Cleanup
        End If

        ' Accumulate incoming bytes
        If bytesRead > 0 Then
            hasData = True
            
            ' AV-Safe manual copy to avoid RtlMoveMemory/CopyMemory detection
            If bytesRead = m_BufferSize Then
                m_InStream.Write m_ReceiveBuffer
            Else
                Dim temp() As Byte: ReDim temp(bytesRead - 1)
                Dim i As Long: For i = 0 To bytesRead - 1: temp(i) = m_ReceiveBuffer(i): Next i
                m_InStream.Write temp
            End If
            
            ' Peak-holding buffer scaling
            MaybeGrowBuffer bytesRead
        End If

        ' 5. Terminal Condition: Full Message Received
        ' Exit loop only when a complete UTF8 or Binary message type is identified.
        ' Fragments (types 1 and 3) are accumulated until the terminal frame arrives.
        If bufType = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Or _
           bufType = WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE Then
            isMessageFinished = True
            Exit Do
        End If
        
        ' 6. Deadline Monitoring (Partial Data Guard)
        ' If deadline expires before a message is finished, we exit the loop.
        ' isMessageFinished remains False, causing the partial/corrupt data to be discarded.
        If deadline > 0 Then
            If GetTickMS() > deadline Then
                If DEBUG_MODE Then Debug.Print "BiDi: ReadFromWire deadline reached. Discarding partial message."
                Exit Do
            End If
        End If

        ' Yield to system while maintaining stream state
        YieldExecution 0, False
    Loop

    ' 7. Safe Decoding
    ' Convert binary stream to UTF-8 text ONLY if the full message was successfully received.
    If isMessageFinished And hasData Then
        m_InStream.Position = 0
        m_InStream.Type = adTypeText
        m_InStream.Charset = "UTF-8"
        ReadFromWire = m_InStream.ReadText
        p_lastActivityTime = GetTickMS()
    Else
        ' Discard partial sequences to avoid JSON parsing errors or UTF-8 corruption
        ReadFromWire = ""
    End If

Cleanup:
    ' 8. Release Reentrancy Guard
    m_isReading = False
    If Err.Number <> 0 Then
        Debug.Print "BiDi Socket Critical Error in ReadFromWire: " & Err.Description
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

' ========================================================================================
' Extracts the root-level "id" value from a JSON string.
' Optimized to prevent collision with nested IDs in BiDi events/results.
' ========================================================================================
Private Function ExtractJsonId(ByVal json As String) As String
    ' 1. Structural Guard: Must be a JSON object
    Dim s As String: s = LTrim$(json)
    Dim length As Long: length = Len(s)
    If length < 10 Then Exit Function
    If Left$(s, 1) <> "{" Then Exit Function

    ' 2. Scan Area Optimization
    ' Root-level "id" in BiDi/CDP is always near the start.
    ' Limiting scan to 1KB prevents performance hits and deep-nested ID collisions.
    Dim searchArea As String
    If length > 1024 Then
        searchArea = Left$(s, 1024)
    Else
        searchArea = s
    End If

    ' 3. Pre-check for "id" position (Performance optimization)
    Dim idPos As Long: idPos = InStr(searchArea, """id""")
    If idPos = 0 Then Exit Function
    
    ' Heuristic Guard: If "id" is too deep (e.g., after 512 chars),
    ' it's likely not the root command ID in this protocol's context.
    If idPos > 512 Then Exit Function

    ' 4. Execute Regex Search
    Dim matches As MatchCollection
    Set matches = p_RegExp.execute(searchArea)
    
    If matches.Count > 0 Then
        ' Check if this match is truly at the root level.
        ' Logic: There should be no nested object start "{" before the first "id" key.
        Dim matchPos As Long: matchPos = matches(0).FirstIndex + 1
        Dim firstNestedBrace As Long: firstNestedBrace = InStr(2, searchArea, "{")
        
        ' If a nested object starts BEFORE our "id", this "id" is probably nested.
        ' (e.g., {"result":{"id":999}, "id":3} -> matches the first "id" 999)
        If firstNestedBrace = 0 Or matchPos < firstNestedBrace Then
            ExtractJsonId = matches(0).SubMatches(0)
        Else
            ' This was a nested ID. For safety in SendAndReceive, we return Empty.
            If DEBUG_MODE Then Debug.Print "BiDi: Nested ID detected and ignored at pos " & matchPos
        End If
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
