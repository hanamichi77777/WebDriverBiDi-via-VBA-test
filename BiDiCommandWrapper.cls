VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BiDiCommandWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ========================================================================================
' Class Name: BiDiCommandWrapper (CDP Tunneling Enabled & AI-Optimized)
' Summary   : High-Level Wrapper for Selenium WebDriver BiDi with CDP Integration
' Features  :
'   - Resource Blocking via CDP Tunneling (Fetch.enable / Network.setBlockedURLs)
'   - Smart Frame Recovery (Auto-handles frame reloads/stale contexts)
'   - Centralized Event Dispatching (Fixes event conflicts)
'   - Anti-Freeze Alert Handling
'   - AI-Ready Discovery Recorder
'   - Advanced Wait Strategies (Network Idle, DOM Stability, SPA Idle)
'
' NOTE ON ERROR HANDLING STRATEGY (Strategic Buffer):
' This class utilizes "On Error Resume Next" (OERN) in specific interaction segments.
' [Rationale]
' The WebDriver BiDi protocol is highly asynchronous. In high-speed scenarios (e.g.,
' ServiceNow redirects), a "Race Condition" can occur where a Context ID or Realm
' becomes invalid between the verification and execution phases.
' [Implementation]
' Instead of complex recovery loops that often lead to "601 Send failed" by
' overwhelming the socket, OERN acts as a "Strategic Buffer." It allows the
' process to bypass transient state mismatches and naturally synchronize with
' the next stable state, ensuring maximum reliability in dynamic SPAs.
'
' Dependencies: BiDiSocketCommunicator, WebJsonConverter, Microsoft Scripting Runtime
' ========================================================================================
Private socket_ As BiDiSocketCommunicator
Private Const DEBUG_MODE As Boolean = True

' Recovery & Stability
Private m_isRecovering As Boolean
Private m_lastDoEventsTime As Currency
Private Const MAX_CONTEXT_RECOVERY_RETRIES As Long = 5

Private m_autoAcceptAlerts As Boolean ' Flag to control automatic alert handling
Private m_autoPromptText As String    ' Text to be entered into prompt dialogs
Private m_activeSubscriptions As Dictionary  'Event Management (Aggregation)
Private m_originalInteractive As Boolean
Private m_originalCursor As Long
Private m_timerPrecisionBuffer As Long
Private m_inflightRequests As Dictionary ' Track active network requests by requestId

' [Configuration Constants]
Private Const DEFAULT_IDLE_STABLE_MS As Long = 500
Private Const DEFAULT_IDLE_TIMEOUT_MS As Long = 10000
Private Const DEFAULT_STALE_RETRIES As Long = 3
Private Const SHORT_WAIT_MS As Long = 500
Private Const MAX_LOG_SIZE As Long = 500

' Adjusted for heavy SPAs (ServiceNow, etc.) to handle redirect instability
Private Const RECOVERY_BASE_WAIT_MS As Long = 500    ' Base wait for recovery sequence
Private Const ACTION_STABILIZE_MS As Long = 150      ' Buffer after clicks/inputs
Private Const PROBE_INJECT_WAIT_MS As Long = 50     ' Buffer after JS probe injection

' --- Noise Filtering State ---
' Default filters are built-in to handle common noise (Analytics, Hidden inputs, etc.)
Private Const DEFAULT_IGNORE_SELECTORS As String = "script, link, input[type='hidden'], .vba-ignore, #vba-bidi-hud"
Private Const DEFAULT_IGNORE_NETWORK As String = "google-analytics, doubleclick, /collect, /beacon, .woff2, favicon.ico"

' User-defined filters set from Standard Modules
Private p_userIgnoreSelectors As Collection
Private p_userIgnoreNetworkPatterns As Collection

' --- State flags for persistent registrations ---
Private m_waitStackDepth As Long     'Tracks recursion depth of active wait loops to allow nested stability checks and manage Excel UI state via the outermost caller.
Private m_lastJsNow As Double        ' Heartbeat tracker (performance.now from JS)
Private m_cachedRealmIds As Dictionary  ' Cached Execution Realm ID

' Internal State Variables
Private p_nextRequestId As Long
Private p_mainContextId As String
Private p_cdpSessionId As String
Private p_pendingRequestCount As Long

' Internal Recorder Storage
Private p_recorderLogs As Collection
Private p_isRecording As Boolean
Private p_recorderStartTime As Double
Private p_excludeResources As Boolean

' Timer Frequency Cache
Private p_timerFreq As Currency

' Threshold for skipping messages to prevent VBA hanging during JSON parsing (approx. 300KB)
Private Const MAX_MSG_SIZE As Long = 300000

' Global retry counter for CDP session failures
Private m_consecutiveCdpFailures As Long
Private Const MAX_CONSECUTIVE_CDP_FAILURES As Long = 5

' Ignore extremely short messages to reduce unnecessary processing overhead
Private Const MIN_MSG_SIZE As Long = 10

' ========================================================================================
' API Declarations (Replaced GetTickCount with QueryPerformanceCounter)
' ========================================================================================
Private Declare PtrSafe Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- Use this (already exists in your variable list) ---
Private m_isProcessingQueue As Boolean

' Error Codes (Centralized for better debugging)
Private Const ERR_BIDI_PROTOCOL As Long = vbObjectError + 1001  ' Protocol-level error (e.g., Invalid params)
Private Const ERR_JS_EXCEPTION As Long = vbObjectError + 1002   ' Runtime JS exception (e.g., SyntaxError)
Private Const ERR_ACTION_FAILED As Long = vbObjectError + 1003  ' Functional failure (e.g., Element not found)

' Public Enumerations for better Intellisense and Type Safety
Public Enum BiDiMatchBy
    MatchByUrl = 0
    MatchByTitle = 1
End Enum

Public Enum BiDiContextType
    ContextTypeTab = 0
    ContextTypeWindow = 1
End Enum

' --- Waiting Engine Internals ---
Private Enum waitMode
    WaitMode_UrlContains
    WaitMode_SpaIdle
End Enum
' ========================================================================================
' DEVELOPMENT & EXTENSION RULES FOR STABILITY
' ----------------------------------------------------------------------------------------
' To maintain "Anti-Freeze" architecture and "AI-Ready" predictability,
' all new interaction methods MUST adhere to the following 4 Golden Rules:
' 1. [BASE UTILIZATION]: All JavaScript execution MUST go through 'Me.CallScript'.
'    - Rationale: This ensures the 'WrapInSafetyShell' is automatically applied,
'      activating the browser-side watchdog timer to prevent VBA hanging.
' 2. [HEARTBEAT INTEGRATION (hb)]: For JS involving loops, polling, or long waits:
'    - Action: Call 'hb()' inside the loop.
'    - Rationale: hb() is provided via closure from WrapInSafetyShell.
'      Do NOT write 'const hb = arguments[0]' ? hb is NOT passed as an argument,
'      and doing so will overwrite your actual parameters (like sharedId).
'      Calling hb() resets the browser-side watchdog timer to prevent task killing.
' 3. [TIMEOUT DELEGATION]: Every Public method MUST expose a 'timeoutMs' parameter.
'    - Action: Pass this parameter down to 'CallScript'.
'    - Rationale: This allows the caller to adjust the watchdog limit based on
'      network latency or site-specific heavy processing.
' 4. [VBA UI RESPONSIVENESS]: For any loop logic created within VBA (Do...Loop):
'    - Action: You MUST call 'WaitHeartbeat' or 'ThrottledDoEvents'.
'    - Rationale: This pumps the Windows message queue, preventing Excel from
'      entering a "Not Responding" state during high-precision waits.

' ========================================================================================
' Initialization: Setup WebSocket and validate High-Resolution Timer frequency.
' ========================================================================================
Private Sub Class_Initialize()
    Set socket_ = New BiDiSocketCommunicator
    Set m_activeSubscriptions = New Dictionary ' Initialize the registry
    Set m_inflightRequests = New Dictionary
    p_nextRequestId = 1
    QueryPerformanceFrequency p_timerFreq
End Sub

' ========================================================================================
' METHOD: ConnectTo
' SUMMARY: Connects to the WebSocket and subscribes to prompt events.
' ========================================================================================
Public Sub ConnectTo(ByVal webSocketUrl As String)
    ' Reset wrapper state before connecting.
    ' This ensures old subscription records don't block mandatory session.subscribe calls.
    Me.ClearState
    
    socket_.AttemptAutoConnect webSocketUrl
    
    ' This will now execute session.subscribe because the dictionary was cleared above.
    Me.ExecuteSessionSubscribe Array("log.entryAdded"), , True
End Sub

' ========================================================================================
' PROPERTY: AutoAcceptAlerts
' DESCRIPTION: If True, the wrapper will automatically click "OK" on any JS alert/confirm.
' ========================================================================================
Public Property Let AutoAcceptAlerts(ByVal Value As Boolean)
    m_autoAcceptAlerts = Value
End Property

' ========================================================================================
' PROPERTY: AutoPromptText
' DESCRIPTION: Sets the default text to be entered when a JS prompt appearing.
' ========================================================================================
Public Property Let AutoPromptText(ByVal Value As String)
    m_autoPromptText = Value
End Property

' ========================================================================================
' Class_Terminate: Resource cleanup including explicit CDP session detachment
' ========================================================================================
Private Sub Class_Terminate()
    On Error Resume Next
    
    ' EMERGENCY UI RESTORATION: In case the macro was stopped via the "End" button
    Application.Interactive = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    
    ' ATTEMPT HUD REMOVAL
    If p_mainContextId <> "" Then
        Call Me.CallScript(GetHudJs("remove"), Nothing, p_mainContextId, False, False, 1000, True)
    End If
    
    ' Explicitly release BiDi/CDP resources
    Call Shutdown
    
    If DEBUG_MODE Then Debug.Print "BiDiCommandWrapper: Terminated and UI protection released."
End Sub

Public Property Set SetSocket(socket As BiDiSocketCommunicator)
    Set socket_ = socket
End Property

' ========================================================================================
' Central Event Dispatcher (Modified with Smart Filtering)
' ========================================================================================
Private Sub DispatchGlobalEvents(ByVal json As Object)
    If Not json.Exists("method") Then Exit Sub
    Dim method As String: method = json("method")

    ' Process CDP-specific blocking events via tunneling
    ProcessCDPBlockingEvent json

    Select Case method
        ' --- BROWSER CONSOLE LOGS ---
        Case "log.entryAdded"
            Call HandleConsoleLog(json)
            
        ' --- [FIXED] NETWORK REQUEST START ---
        Case "network.beforeRequestSent"
            Dim reqId As String: reqId = GetPathString(json, "params.request.requestId")
            Dim reqUrl As String: reqUrl = GetPathString(json, "params.request.url")
            
            ' Only track significant (non-noise) requests.
            ' We store the ID in a dictionary to ensure the same request is decremented later.
            If Not IsNoiseUrl(reqUrl) Then
                If Not m_inflightRequests.Exists(reqId) Then
                    m_inflightRequests.Add reqId, True
                    p_pendingRequestCount = p_pendingRequestCount + 1
                End If
            End If

        ' --- [FIXED] NETWORK REQUEST END (Success or Error) ---
        Case "network.responseCompleted", "network.fetchError"
            Dim resId As String: resId = GetPathString(json, "params.request.requestId")
            
            ' If this requestId was being tracked, decrement the counter.
            ' We check the Dictionary, NOT the URL here, to prevent redirect asymmetry.
            If m_inflightRequests.Exists(resId) Then
                m_inflightRequests.Remove resId
                p_pendingRequestCount = p_pendingRequestCount - 1
            End If
            
            ' Safety Guard: Prevent count from ever falling below zero due to protocol desync.
            If p_pendingRequestCount < 0 Then p_pendingRequestCount = 0

    End Select

    ' Store event in Discovery Recorder if active
    If p_isRecording Then
        FormatAndStoreLog json
    End If
End Sub

Public Property Get PendingRequestCount() As Long
    PendingRequestCount = p_pendingRequestCount
End Property

' ========================================================================================
' [PUBLIC] GetIframeContextIdByUrl
' DESCRIPTION: Searches for a specific iframe context ID by its partial URL.
'              This method implements a "Fail-Fast" strategy by raising a VBA error
'              if the target frame is not located within the specified timeout.
' ========================================================================================
Public Function GetIframeContextIdByUrl(ByVal partialUrl As String, Optional ByVal timeOutms As Long = 5000) As String
    ' --- Input Validation ---
    ' Ensures timeout is not negative and stays within a safe range.
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeOutms, "timeoutMs", 5000)
    
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim ctxId As String

    Do
        ' 1. Fetch the browsing context tree and perform a recursive search
        Dim strRes As String: strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Dim contexts As Collection: Set contexts = GetJsonValue(strRes, "result.contexts")
        
        If Not contexts Is Nothing Then
            ctxId = RecursiveFrameSearch(contexts, partialUrl)
            If ctxId <> "" Then
                ' Return found Context ID and exit function
                GetIframeContextIdByUrl = ctxId
                Exit Function
            End If
        End If

        ' Immediate exit if the user specified no wait (0ms)
        If vTimeout <= 0 Then Exit Do

        ' 2. Maintain UI responsiveness and check for timeout
        ' WaitHeartbeat calls DoEvents internally to prevent Excel from freezing.
        WaitHeartbeat SHORT_WAIT_MS
        
        ' --- Fail-Fast Logic ---
        ' Raise a runtime error if the element is not found within the timeout period.
        ' This prevents subsequent actions from executing on an invalid/empty context.
        If IsTimeout(startTick, CDbl(vTimeout)) Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:GetIframeContextIdByUrl", _
                      "Iframe context not found within " & vTimeout & "ms (Target URL pattern: " & partialUrl & ")"
        End If
    Loop
End Function

' ========================================================================================
' [FIXED] Recursive search with Stack Guard and Depth Limit
' ========================================================================================
Private Function RecursiveFrameSearch(ByVal contextList As Collection, _
                                      ByVal partialUrl As String, _
                                      Optional ByVal currentDepth As Long = 0) As String
    
    ' Guard: Recursion depth limit
    Const MAX_RECURSION_DEPTH As Long = 20
    
    On Error GoTo StackGuard
    
    If currentDepth > MAX_RECURSION_DEPTH Then
        Debug.Print "BiDi Warning: Max recursion depth reached (" & MAX_RECURSION_DEPTH & ")"
        RecursiveFrameSearch = ""
        Exit Function
    End If
    
    Dim ctx As Dictionary
    Dim childRes As String

    For Each ctx In contextList
        ' URL Check
        If InStr(ctx("url"), partialUrl) > 0 Then
            RecursiveFrameSearch = ctx("context")
            Exit Function
        End If

        ' Recursive search for children (iframes)
        If ctx.Exists("children") Then
            If IsObject(ctx("children")) Then
                If Not ctx("children") Is Nothing Then
                    childRes = RecursiveFrameSearch(ctx("children"), partialUrl, currentDepth + 1)
                    If childRes <> "" Then
                        RecursiveFrameSearch = childRes
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ctx

    RecursiveFrameSearch = ""
    Exit Function

StackGuard:
    If Err.Number = 28 Then ' Out of stack space
        Debug.Print "BiDi Critical: Stack overflow detected in RecursiveFrameSearch"
        RecursiveFrameSearch = ""
    Else
        Debug.Print "BiDi Error in RecursiveFrameSearch: " & Err.Description
        RecursiveFrameSearch = ""
    End If
End Function

' ========================================================================================
' [PUBLIC] ExecuteFindWindowContextId
' DESCRIPTION: Locates a window context (tab/window) by URL or Title.
' ========================================================================================
Public Function ExecuteFindWindowContextId(ByVal matchType As BiDiMatchBy, _
                                           ByVal matchValue As String, _
                                           Optional ByVal timeOutms As Long = 5000, _
                                           Optional ByVal waitForComplete As Boolean = True) As String
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeOutms, "timeoutMs", 5000)
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim strRes As String, res As Variant
    Dim jsonTree As Object, ctxList As Collection, ctx As Dictionary
    
    Do
        strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Set jsonTree = WebJsonConverter.ParseJson(strRes)
        Set ctxList = GetPathObject(jsonTree, "result.contexts")

        If Not ctxList Is Nothing Then
            For Each ctx In ctxList
                Dim isMatch As Boolean
                If matchType = MatchByUrl Then
                    isMatch = (InStr(ctx("url"), matchValue) > 0)
                ElseIf matchType = MatchByTitle Then
                    isMatch = (InStr(ExecuteGetTitleByContextId(ctx("context")), matchValue) > 0)
                End If

                If isMatch Then
                    If Not waitForComplete Then
                        ExecuteFindWindowContextId = ctx("context"): Exit Function
                    Else
                        ' Leverage the new simplified pipeline
                        res = GetCleanResult(CallScript("return document.readyState;", , ctx("context")))
                        If CStr(res) = "complete" Then
                            ExecuteFindWindowContextId = ctx("context"): Exit Function
                        End If
                    End If
                End If
            Next ctx
        End If

        WaitHeartbeat SHORT_WAIT_MS
        If IsTimeout(startTick, CDbl(vTimeout)) Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteFindWindowContextId", _
                      "Target window not found within " & vTimeout & "ms (Match: " & matchValue & ")"
        End If
    Loop
End Function

' ========================================================================================
' FUNCTION: ExecuteCreateContext
' SUMMARY: Creates a new browsing context (tab or window).
' ========================================================================================
Public Function ExecuteCreateContext(Optional ByVal createType As BiDiContextType = ContextTypeTab) As String
    Dim typeStr As String
    If createType = ContextTypeWindow Then
        typeStr = "window"
    Else
        typeStr = "tab"
    End If

    Dim params As New Dictionary
    params.Add "type", typeStr
    
    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.create", params)
    
    ' Parse once and extract ID using the new path getter
    Dim json As Object
    Set json = WebJsonConverter.ParseJson(strRes)
    ExecuteCreateContext = GetPathString(json, "result.context")
End Function

Public Function ExecuteCloseContext(ByVal contextId As String) As String
    Dim params As New Dictionary
    params.Add "context", contextId
    ExecuteCloseContext = ExecuteBiDiCommand("browsingContext.close", params)
    If contextId = p_mainContextId Then ClearState
End Function

' ========================================================================================
' [PUBLIC] ExecuteGetTitleByContextId (Optimized)
' ========================================================================================
Public Function ExecuteGetTitleByContextId(ByVal contextId As String) As String
    ' Note: CallScript internally handles the realm discovery and shell unwrapping
    Dim strRes As String: strRes = CallScript("return document.title || '';", , contextId)
    
    ' Return the plain string extracted from the JSON response
    ExecuteGetTitleByContextId = CStr(GetCleanResult(strRes))
End Function

Public Function ExecuteHandleUserPrompt(ByVal Accept As Boolean, Optional ByVal userText As String = "", Optional ByVal contextId As String = "") As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "accept", Accept
    If Len(userText) > 0 Then params.Add "userText", userText

    On Error Resume Next
    ExecuteHandleUserPrompt = ExecuteBiDiCommand("browsingContext.handleUserPrompt", params)
    On Error GoTo 0
End Function

' ========================================================================================
' [PUBLIC] ExecuteIsUrlContains
' DESCRIPTION: Checks if the current URL contains the partial string. Polls until match
'              is found or timeout occurs. Optionally waits for SPA stability.
' RETURNS    : True if URL contains the string within timeout, False otherwise.
' ========================================================================================
Public Function ExecuteIsUrlContains(ByVal partialUrl As String, _
                                     Optional ByVal waitNetworkIdle As Boolean = True, _
                                     Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                     Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                     Optional ByVal contextId As String = "") As Boolean
    
    ' --- [ENTRY GUARD] Guard against 0ms when sync is requested ---
    ' Rationale: If waitNetworkIdle is True, we must ensure that the synchronization
    ' parameters are valid positive integers to prevent the SPA sync from failing
    ' silently or exiting prematurely due to the 0ms-allowance in SanitizeMs.
    If waitNetworkIdle Then
        If minStableMs = 0 Then minStableMs = DEFAULT_IDLE_STABLE_MS
        If maxTimeoutMs = 0 Then maxTimeoutMs = DEFAULT_IDLE_TIMEOUT_MS
    End If
    ' --------------------------------------------------------------
    
    ' 1. Sanitize timing parameters
    ' SanitizeMs now respects 0ms for immediate checks, but our Entry Guard
    ' has already ensured they are non-zero if waitNetworkIdle is True.
    Dim vMinStable As Long: vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vMaxTimeout As Long: vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    
    ' 2. Resolve target context
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 3. Subscribe to navigation events required for monitoring
    Me.ExecuteSessionSubscribe Array("network.responseCompleted", "browsingContext.load", "browsingContext.fragmentNavigated")
    
    ' 4. Initialize wait parameters for the polling engine
    Dim waitCtx As New Dictionary
    waitCtx.Add "contextId", targetCtx
    waitCtx.Add "partialUrl", partialUrl
    
    ' 5. Start the engine. Internally calls EvaluateWaitCondition -> GetCurrentUrl (Clean)
    Dim isMatched As Boolean
    ' Note: If vMaxTimeout is 0 (and waitNetworkIdle was False), this will perform a single immediate check.
    isMatched = MasterWaitEngine(WaitMode_UrlContains, vMaxTimeout, 200, waitCtx)
    
    ' 6. Synchronize environment if requested and match was successful
    If isMatched And waitNetworkIdle Then
        ' vMinStable and vMaxTimeout are guaranteed to be valid due to our Entry Guard.
        SyncAfterAction targetCtx, vMinStable, vMaxTimeout
    End If

    ExecuteIsUrlContains = isMatched
End Function

' ========================================================================================
' FUNCTION: ExecuteNavigateAndGetStatus
' SUMMARY: Public entry point for navigation. Returns "200" on success or "500" on error.
'          This high-level method leverages the private navigation engine.
' ========================================================================================
Public Function ExecuteNavigateAndGetStatus(ByVal targetUrl As String, _
                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "") As String

    Dim vMinStable As Long:  vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vMaxTimeout As Long: vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    Dim attempt As Long
    Const MAX_NAV_RETRIES As Long = 2
    
    Dim isNavigated As Boolean
    Dim modeStr As String
    If waitNetworkIdle Then
        modeStr = "complete"
    Else: modeStr = "none"
    End If

    For attempt = 0 To MAX_NAV_RETRIES
        isNavigated = False
        On Error Resume Next
        
        ' Initiate navigation.
        ' BiDi requestId is automatically refreshed inside ExecuteBiDiCommand per attempt.
        Call ExecuteBrowsingContextNavigate(targetUrl, modeStr, vMinStable, vMaxTimeout, targetCtx)
        
        If Err.Number = 0 Then
            ' --- [FIXED] Conditional Liveness Check ---
            If Not waitNetworkIdle Then
                ' Mode "none": Perform quick URL validation since the protocol returns immediately.
                Dim curUrl As String: curUrl = GetCurrentUrl(targetCtx, True)
                If Not (InStr(1, curUrl, "chrome-error", vbTextCompare) > 0 Or curUrl = "") Then
                    isNavigated = True
                End If
            Else
                ' Mode "complete": Success is implied by a zero Err.Number after stability sync.
                isNavigated = True
            End If
        End If
        
        On Error GoTo 0
        
        If isNavigated Then
            ExecuteNavigateAndGetStatus = "200"
            Exit Function
        End If

        ' Wait for the environment to settle before retry
        WaitHeartbeat 1000
    Next attempt
    
    ExecuteNavigateAndGetStatus = "500"
End Function

Private Function ExecuteGetReadyState(ByVal ctxId As String) As String
    Dim p As New Dictionary, t As New Dictionary
    t.Add "context", ctxId
    p.Add "functionDeclaration", "function(){ return document.readyState; }"
    p.Add "target", t: p.Add "awaitPromise", False

    Dim ret As String: ret = ExecuteBiDiCommand("script.callFunction", p)
    Dim json As Object: Set json = WebJsonConverter.ParseJson(ret)
    ExecuteGetReadyState = GetPathString(json, "result.result.value")
End Function

' ========================================================================================
' EnsureCdpBlockingInitialized: Ensures CDP domains required for blocking are enabled.
' This is required before calling Network.setBlockedURLs or Fetch.failRequest.
' ========================================================================================
Private Function EnsureCdpBlockingInitialized() As Boolean
    Dim res As Object
    
    ' Force Silent:=True for infrastructure setup commands
    Set res = ExecuteCDPCommand("Network.enable", New Dictionary, True, True)
    If Not res("Ok") Then Exit Function

    Set res = ExecuteCDPCommand("Fetch.enable", New Dictionary, True, True)
    EnsureCdpBlockingInitialized = res("Ok")
End Function

' ========================================================================================
' ExecuteEnableResourceBlocking: Sets URL patterns to be blocked via CDP.
' ========================================================================================
Public Function ExecuteEnableResourceBlocking(ByVal patterns As Variant) As String
    Dim params As New Dictionary
    Dim urlList As New Collection
    Dim Item As Variant

    ' Normalize input patterns into a Collection
    If IsArray(patterns) Then
        For Each Item In patterns: urlList.Add CStr(Item): Next
    Else
        urlList.Add CStr(patterns)
    End If
    params.Add "urls", urlList

    ' --- Proactive Initialization ---
    ' Ensure the browser is ready to intercept and block network requests.
    
    If Not EnsureCdpBlockingInitialized() Then
        ExecuteEnableResourceBlocking = "{""error"":""CDP_INIT_FAILURE"",""message"":""Failed to enable Network/Fetch domains.""}"
        Exit Function
    End If

    ' Execute the blocking command via CDP tunneling
    Dim resObj As Object: Set resObj = ExecuteCDPCommand("Network.setBlockedURLs", params, True, True)
    
    If Not resObj("Ok") Then
        ExecuteEnableResourceBlocking = "{""error"":""CDP_FAILURE""}"
    Else
        ExecuteEnableResourceBlocking = resObj("Json")
    End If
End Function

Public Sub ProcessCDPBlockingEvent(ByVal json As Object)
    If Not json.Exists("method") Then Exit Sub

    Dim method As String
    Dim eventParams As Object
    method = json("method")

    If method = "goog:cdp.event" Then
        If json.Exists("params") Then
            Set eventParams = json("params")
            If eventParams.Exists("method") Then
                method = eventParams("method")
                If eventParams.Exists("params") Then
                    Set eventParams = eventParams("params")
                End If
            End If
        End If
    ElseIf json.Exists("params") Then
        Set eventParams = json("params")
    End If

    If method = "Fetch.requestPaused" Then
        Dim requestId As String
        requestId = eventParams("requestId")

        Dim failParams As New Dictionary
        failParams.Add "requestId", requestId
        failParams.Add "errorReason", "BlockedByClient"

        If DEBUG_MODE Then
            Dim url As String
            On Error Resume Next
            url = eventParams("request")("url")
            On Error GoTo 0
            Debug.Print "[CDP Blocked] " & url
        End If

        ExecuteCDPCommand "Fetch.failRequest", failParams
    End If
End Sub

' ========================================================================================
' [PUBLIC] ExecuteIsElementVisible
' DESCRIPTION: Checks if an element is rendered and visible in the viewport.
'              Returns True only if the element exists in the DOM and is functionally visible.
'              Supports 0ms searchTimeoutMs for high-speed immediate polling.
' ========================================================================================
Public Function ExecuteIsElementVisible(ByVal xpath As String, _
                                        Optional ByVal searchTimeoutMs As Long = 10000, _
                                        Optional ByVal waitNetworkIdle As Boolean = False, _
                                        Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                        Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                        Optional ByVal contextId As String = "") As Boolean
    
    ' --- [ENTRY GUARD] Guard against 0ms only when sync is requested ---
    ' Enforce safety floor for stability consensus parameters.
    If waitNetworkIdle Then
        If minStableMs <= 0 Then minStableMs = DEFAULT_IDLE_STABLE_MS
        If maxTimeoutMs <= 0 Then maxTimeoutMs = DEFAULT_IDLE_TIMEOUT_MS
    End If
    ' -------------------------------------------------------------------
    
    ' 1. Consolidate validation for all timing parameters at the entry point.
    ' Uses the enhanced SanitizeMs to ensure safe floor values.
    Dim vSearchTimeout As Long: vSearchTimeout = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 2000)
    Dim vMinStable As Long:     vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vMaxTimeout As Long:    vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    
    ' 2. Resolve the target browsing context.
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 3. Locate the element.
    ' Returns a sharedId if found within vSearchTimeout.
    Dim sharedId As String
    sharedId = GetFindElementSharedIdByXPath(xpath, vSearchTimeout, targetCtx)
    
    ' 4. Early Exit: If the element is not found, it cannot be visible.
    If sharedId = "" Then
        ExecuteIsElementVisible = False
        Exit Function
    End If

    ' 5. Prepare JavaScript arguments (RemoteObject reference).
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "sharedId", sharedId: args.Add arg1

    ' 6. Execution: Perform the visual state check via the central executor.
    ' strRes will be a flattened JSON string returned by the new ProcessResponse.
    Dim strRes As String
    strRes = ExecuteBaseAction(targetCtx, "visibility_check", args, waitNetworkIdle, vMinStable, True, DEFAULT_STALE_RETRIES, vMaxTimeout)
    
    ' 7. [FIXED] Parse Result using the unified pipeline.
    ' Rationale: We no longer manually parse "result.result.value" because GetCleanResult
    ' handles the new flattened JSON structure {"status":"ok", "value":...} correctly.
    On Error Resume Next
    Dim resultVal As Variant
    resultVal = GetCleanResult(strRes)
    
    ' If the result is not Empty, convert the JS boolean to a VBA Boolean.
    If Not IsEmpty(resultVal) Then
        ExecuteIsElementVisible = CBool(resultVal)
    End If
    On Error GoTo 0
End Function

' ========================================================================================
' ExecuteSessionSubscribe: Subscribes to BiDi events with optional context filtering.
' ========================================================================================
Public Function ExecuteSessionSubscribe(ByVal events As Variant, _
                                         Optional ByVal contexts As Collection = Nothing, _
                                         Optional ByVal Silent As Boolean = False) As String
    Dim eventArray() As String
    Dim eventsToRequest As New Collection
    Dim subKey As String, ctxId As String
    Dim i As Long, v As Variant

    ' 1. Resolve context identifier for keying
    ctxId = ""
    If Not contexts Is Nothing Then
        If contexts.Count > 0 Then ctxId = "|" & CStr(contexts(1))
    End If

    ' 2. Normalize input to a String Array
    If IsArray(events) Then
        ReDim eventArray(UBound(events))
        For i = LBound(events) To UBound(events)
            eventArray(i) = CStr(events(i))
        Next i
    Else
        ReDim eventArray(0)
        eventArray(0) = CStr(events)
    End If

    ' 3. Delta Check: Identify only new, unsubscribed events
    For i = LBound(eventArray) To UBound(eventArray)
        subKey = eventArray(i) & ctxId
        If Not m_activeSubscriptions.Exists(subKey) Then
            eventsToRequest.Add eventArray(i)
        End If
    Next i

    ' 4. Early Exit: If all requested events are already active, skip the protocol call.
    If eventsToRequest.Count = 0 Then
        ' [S-Rank Fix] Respect the Silent flag even during early exit to keep logs clean.
        ' This prevents redundant "Subscription skipped" logs when re-connecting or registering tools.
        If DEBUG_MODE And Not Silent Then
            Debug.Print "BiDi: Subscription skipped (Events already active)."
        End If
        ExecuteSessionSubscribe = "{""status"":""already_subscribed""}"
        Exit Function
    End If

    ' 5. Prepare Payload for protocol-level subscription
    Dim params As New Dictionary
    Dim finalArr() As String
    ReDim finalArr(eventsToRequest.Count - 1)
    
    i = 0
    For Each v In eventsToRequest
        finalArr(i) = CStr(v)
        i = i + 1
    Next v
    
    params.Add "events", finalArr
    If Not contexts Is Nothing Then params.Add "contexts", contexts

    ' 6. Execute BiDi Command
    Dim strRes As String
    strRes = ExecuteBiDiCommand("session.subscribe", params, Silent)
    
    ' 7. Persistence: Store the SubscriptionID upon success
    Dim subId As String
    subId = GetPathString(WebJsonConverter.ParseJson(strRes), "result.subscription")
    
    If subId <> "" Then
        For i = 0 To UBound(finalArr)
            m_activeSubscriptions(finalArr(i) & ctxId) = subId
        Next i
    End If

    ExecuteSessionSubscribe = strRes
End Function

' ========================================================================================
' [PRIVATE] ExecuteBrowsingContextNavigate (Refactored for Anti-Freeze)
' SUMMARY: Initiates navigation without blocking at the socket level.
'          Uses VBA-side polling to monitor stability and handle timeouts.
' ========================================================================================
Private Function ExecuteBrowsingContextNavigate(ByVal url As String, _
                                    Optional ByVal waitMode As String = "complete", _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "") As String
                                                
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' [FIX] Reset only transient navigation/network state.
    ' Rationale: Do NOT call ClearState() here. ClearState nukes p_mainContextId and
    ' m_cachedRealmIds, forcing unnecessary re-discovery trips and causing state loss
    ' if waitMode is not "complete".
    p_pendingRequestCount = 0
    p_cdpSessionId = ""
    
    Dim bidiWaitLevel As String
    bidiWaitLevel = "interactive"
    
    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "url", url
    params.Add "wait", bidiWaitLevel
    
    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.navigate", params)
    Call ProcessResponse(strRes, "browsingContext.navigate")

    ' Stability Consensus Phase
    If LCase(waitMode) = "complete" Then
        ' Wait for the URL and document state to settle
        Me.WaitForUrlStability targetCtx, 1000, maxTimeoutMs
        
        ' Refresh the context tree to ensure mainContextId reflects the current document
        UpdateMainContextId True, True
        
        ' Perform deep SPA synchronization
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx
    End If
    
    ExecuteBrowsingContextNavigate = strRes
End Function

' ========================================================================================
' [PUBLIC] ExecuteClickByXPath
' ========================================================================================
Public Function ExecuteClickByXPath(ByVal xpath As String, _
                                    Optional ByVal searchTimeoutMs As Long = 10000, _
                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "", _
                                    Optional ByVal checkUrl As Boolean = False) As String
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 1. Resolve element (Raises ERR_ACTION_FAILED if not found)
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteClickByXPath")

    ' 2. Pack sharedId into arguments for the JavaScript action
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "sharedId", sharedId: args.Add arg1

    ' 3. Execute action via the base pipeline
    '    The checkUrl flag is passed down to SyncAfterAction to determine the wait strategy.
    ExecuteClickByXPath = ExecuteBaseAction(targetCtx, _
                                            "click", _
                                            args, _
                                            waitNetworkIdle, _
                                            minStableMs, _
                                            True, _
                                            DEFAULT_STALE_RETRIES, _
                                            maxTimeoutMs, _
                                            checkUrl)
End Function

' ========================================================================================
' [PUBLIC] ExecuteInputValueByXPath
' ========================================================================================
Public Function ExecuteInputValueByXPath(ByVal xpath As String, _
                                         ByVal valueToSet As String, _
                                         Optional ByVal searchTimeoutMs As Long = 10000, _
                                         Optional ByVal waitNetworkIdle As Boolean = True, _
                                         Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                         Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                         Optional ByVal contextId As String = "") As String
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 1. Resolve element
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteInputValueByXPath")

    ' 2. Pack arguments: [sharedId, textValue]
    Dim args As New Collection
    Dim a1 As New Dictionary: a1.Add "sharedId", sharedId: args.Add a1
    Dim a2 As New Dictionary: a2.Add "type", "string": a2.Add "value", valueToSet: args.Add a2

    ' 3. Execute input action (Uses JS execCommand for better SPA event firing)
    ExecuteInputValueByXPath = ExecuteBaseAction(targetCtx, "input", args, waitNetworkIdle, minStableMs, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
End Function

' ========================================================================================
' [PUBLIC] ExecuteSelectValueByXPath
' ========================================================================================
Public Function ExecuteSelectValueByXPath(ByVal xpath As String, _
                                          ByVal valueOrText As String, _
                                          Optional ByVal selectByText As Boolean = False, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal waitNetworkIdle As Boolean = True, _
                                          Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                          Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                          Optional ByVal contextId As String = "") As String
    
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId()
    
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteSelectValueByXPath")

    Dim jsKey As String
    If selectByText Then jsKey = "select_text" Else jsKey = "select"

    Dim args As New Collection
    Dim a1 As New Dictionary: a1.Add "sharedId", sharedId: args.Add a1
    Dim a2 As New Dictionary: a2.Add "type", "string": a2.Add "value", valueOrText: args.Add a2

    ExecuteSelectValueByXPath = ExecuteBaseAction(targetCtx, jsKey, args, waitNetworkIdle, minStableMs, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
End Function

' ========================================================================================
' Locator Helpers
' ========================================================================================
Private Function ExecuteBrowsingContextLocateNodes(ByVal xpath As String, _
                                                  Optional ByVal contextId As String = "", _
                                                  Optional ByVal Silent As Boolean = False) As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim params As New Dictionary, locator As New Dictionary
    locator.Add "type", "xpath": locator.Add "value", xpath
    params.Add "context", targetCtx: params.Add "locator", locator
    
    ' Pass Silent flag down
    ExecuteBrowsingContextLocateNodes = ExecuteBiDiCommand("browsingContext.locateNodes", params, Silent)
End Function

' ========================================================================================
' [INTERNAL] GetFindElementSharedIdByXPath
' DESCRIPTION: Elements may not be ready immediately. This loop polls with a throttle
'              to prevent CPU saturation and "Unresponsive" Excel state.
' ========================================================================================
Private Function GetFindElementSharedIdByXPath(ByVal xpath As String, _
                                               ByVal searchTimeoutMs As Long, _
                                               ByVal targetCtx As String) As String
    
    ' 1. Logic Guard: Determine if this is an immediate (one-shot) check.
    ' IMPORTANT: We must check the ORIGINAL searchTimeoutMs here.
    ' SanitizeMs will enforce a minimum floor (e.g., 50ms), which would otherwise
    ' make (vTimeoutMs <= 0) always False.
    Dim isImmediate As Boolean: isImmediate = (searchTimeoutMs <= 0)
    
    ' 2. Sanitization: Ensure a valid timeout for the IsTimeout comparison.
    Dim vTimeoutMs As Long: vTimeoutMs = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 2000)
    
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim strRes As String, json As Object, nodes As Collection
    Dim sid As String
    Dim attemptCount As Long: attemptCount = 0
    
    If DEBUG_MODE Then
        Debug.Print "BiDi: Searching for [" & xpath & "] (Timeout: " & _
                    IIf(isImmediate, "Immediate", vTimeoutMs & "ms") & ")"
    End If

    ' 3. Polling Loop
    Do
        attemptCount = attemptCount + 1
        
        ' --- Step A: Navigation Guard ---
        ' Abort search every 5 attempts if the tab has navigated to a blank/invalid page.
        ' This prevents wasting time on a dead context.
        If attemptCount Mod 5 = 0 Then
            Dim currentUrl As String: currentUrl = GetCurrentUrl(targetCtx, True)
            If InStr(1, currentUrl, "about:blank", vbTextCompare) > 0 Or currentUrl = "" Then
                Debug.Print "BiDi: Aborting search. Target context navigated away."
                Exit Do
            End If
        End If

        ' --- Step B: Protocol Search ---
        ' Use Silent:=True to prevent log flooding during high-speed polling.
        strRes = ExecuteBrowsingContextLocateNodes(xpath, targetCtx, True)
        Set json = WebJsonConverter.ParseJson(strRes)
        Set nodes = GetPathObject(json, "result.nodes")

        ' --- Step C: Result Extraction ---
        If Not nodes Is Nothing Then
            If nodes.Count > 0 Then
                If nodes(1).Exists("sharedId") Then
                    sid = CStr(nodes(1)("sharedId"))
                    ' SUCCESS: Return the ID and exit immediately.
                    GetFindElementSharedIdByXPath = sid
                    Exit Function
                End If
            End If
        End If

        ' --- Step D: Exit Strategy ---
        ' If 0ms (or less) was requested, exit NOW without hitting WaitHeartbeat.
        If isImmediate Then Exit Do

        ' Yield control to Windows and process background socket messages.
        ' Using 500ms for element discovery polling is a balanced trade-off
        ' between CPU usage and responsiveness.
        WaitHeartbeat 500
        
        ' Standard Timeout Check for designated wait periods.
        If IsTimeout(startTick, CDbl(vTimeoutMs)) Then Exit Do
    Loop
    
    ' FAILURE: Return empty string if not found within the designating window.
    GetFindElementSharedIdByXPath = ""
End Function

' ========================================================================================
' [PUBLIC] GetCurrentUrl (Robust Version)
' ========================================================================================
Public Function GetCurrentUrl(Optional ByVal contextId As String = "", _
                              Optional ByVal Silent As Boolean = False) As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' Relay the Silent flag to the 7th parameter (Silent) of CallScript
    Dim strRes As String: strRes = CallScript("return window.location.href;", Nothing, targetCtx, False, False, 10000, Silent)
    
    GetCurrentUrl = CStr(GetCleanResult(strRes))
End Function

Public Function ExecuteWebExtensionInstall(ByVal extensionPath As String) As String
    Dim params As New Dictionary, extData As New Dictionary
    extData.Add "type", "path"
    extData.Add "path", extensionPath
    params.Add "extensionData", extData
    ExecuteWebExtensionInstall = ExecuteBiDiCommand("webExtension.install", params)
End Function

' ========================================================================================
' [PUBLIC] ExecuteShadowClick
' DESCRIPTION: Penetrates Shadow DOM boundaries to click a nested element.
' ========================================================================================
Public Function ExecuteShadowClick(ByVal selectorsArray As Variant, _
                                   Optional ByVal searchTimeoutMs As Long = 10000, _
                                   Optional ByVal waitNetworkIdle As Boolean = True, _
                                   Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                   Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                   Optional ByVal contextId As String = "") As String
    
    ' --- Input Validation Layer ---
    ' 1. Validate searchTimeoutMs locally because it bypasses the ResolveXPath bottleneck.
    Dim vSearchTimeout As Long: vSearchTimeout = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 5000)
    
    ' 2. Validation for other parameters will be handled by ExecuteBaseAction,
    '    but we pass minStableMs and maxTimeoutMs directly through.
    ' ------------------------------
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' Prepare JS arguments: [selectorsArray, searchTimeout]
    Dim args As New Collection
    args.Add BuildShadowSelectorArg(selectorsArray)
    
    ' Use the validated vSearchTimeout for the browser-side script
    Dim argTime As New Dictionary
    argTime.Add "type", "number": argTime.Add "value", vSearchTimeout
    args.Add argTime
    
    ' Execute via Master Runner (vMinStable and vMaxTimeout are validated inside this call)
    ExecuteShadowClick = ExecuteBaseAction(targetCtx, "shadow_click", args, waitNetworkIdle, minStableMs, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
End Function

' ========================================================================================
' [INTERNAL] ExecuteBiDiCommand (Reverted to High-Efficiency SendAndReceive)
' SUMMARY: Sends a BiDi command and waits for a direct response from the socket.
'          Includes an automated recovery loop for lost contexts or socket timeouts.
' ========================================================================================
Private Function ExecuteBiDiCommand(ByVal method As String, _
                                    ByRef params As Dictionary, _
                                    Optional ByVal Silent As Boolean = False) As String
    
    If socket_ Is Nothing Then Err.Raise 91, "ExecuteBiDiCommand", "Socket object is not set."

    Dim currentRetry As Long
    Dim strRes As String, sendMsg As String, requestDic As Dictionary
    Dim lastErr As Long
    Dim requestId As Long
    
    currentRetry = 0

RecoveryExecution:
    ' --- PHASE 1: Error Handler Initialization ---
    ' Activate the recovery handler before any state modification to ensure
    ' that retries generate a clean, new state for every attempt.
    On Error GoTo RecoveryCleanup
    
    ' --- PHASE 2: Unique ID Generation (Defensive) ---
    ' Assign current ID and increment. Wrap-around at 2 billion to prevent
    ' Long overflow (Err 6) and maintain protocol compatibility.
    requestId = p_nextRequestId
    p_nextRequestId = p_nextRequestId + 1
    If p_nextRequestId > 2000000000 Then p_nextRequestId = 1
        
    ' --- PHASE 3: Request Construction ---
    Set requestDic = New Dictionary
    requestDic.Add "id", requestId
    requestDic.Add "method", method
    requestDic.Add "params", params
    
    sendMsg = WebJsonConverter.ConvertToJson(requestDic)
    
    ' Conditional logging: Suppress if Silent flag is active.
    If DEBUG_MODE And Not Silent Then
        Dim logMsg As String: logMsg = sendMsg
        If Len(logMsg) > 500 Then logMsg = Left$(logMsg, 497) & "..."
        DebugPrintLong method & " (ID: " & requestId & ", Attempt " & currentRetry & ")", logMsg, "[REQ]"
    End If
    
    ' --- PHASE 4: Socket Transmission & Reception ---
    ' Direct communication with the BiDiSocketCommunicator.
    strRes = socket_.SendAndReceive(sendMsg)
    
    ' --- PHASE 5: Environment Validation ---
    ' Check for "Stale" environments or specific browser-side navigation errors.
    If IsContextLostError(strRes) Or (InStr(1, strRes, "chrome-error", vbTextCompare) > 0) Then
        Err.Raise 604, "BiDi:Protocol", "Environment desync detected during transmission."
    End If

    ' --- SUCCESS PATH ---
    ' Deactivate error handler before final result processing to avoid circular traps.
    On Error GoTo 0
    
    If DEBUG_MODE And Not Silent Then
        Dim resLog As String: resLog = strRes
        If Len(resLog) > 500 Then resLog = Left$(resLog, 497) & "..."
        DebugPrintLong method & " (ID: " & requestId & ", Attempt " & currentRetry & ")", resLog, "[RES]"
    End If
    
    ' Clear internal socket error state and pass result to the centralized parser.
    socket_.ClearInternalErrorCounter
    Call ProcessResponse(strRes, method)
    
    ExecuteBiDiCommand = strRes
    Exit Function

' ========================================================================================
' RECOVERY HANDLER (Self-Healing Block)
' ========================================================================================
RecoveryCleanup:
    lastErr = Err.Number
    currentRetry = currentRetry + 1
    
    ' Attempt recovery if maximum retry limit is not reached and not already in recovery mode.
    If Not m_isRecovering And currentRetry <= MAX_CONTEXT_RECOVERY_RETRIES Then
        m_isRecovering = True
        Err.Clear
        
        ' 1. Wait with an adaptive backoff (800ms multiplier)
        Dim adaptiveWaitMs As Long: adaptiveWaitMs = 800 * currentRetry
        ProcessRecorderQueue
        WaitHeartbeat adaptiveWaitMs
        
        ' 2. Re-discover browsing contexts and execute state refresh
        If DEBUG_MODE Then Debug.Print "BiDi: Recovery attempt " & currentRetry & "/" & MAX_CONTEXT_RECOVERY_RETRIES
        UpdateMainContextId True, True
        If p_mainContextId <> "" Then UpdateContextAndRealm params, p_mainContextId
        
        m_isRecovering = False
        
        ' 3. RE-ENTRY: Jump back to the execution block to try with a fresh requestId.
        Resume RecoveryExecution
    Else
        ' FATAL: Maximum retries exceeded or critical failure.
        m_isRecovering = False
        On Error GoTo 0
        Err.Raise lastErr, "BiDi:ExecuteBiDiCommand", _
                  "Command '" & method & "' failed after " & (currentRetry - 1) & " recovery attempts. Last Error: " & lastErr
    End If
End Function

' ========================================================================================
' INTERNAL HELPER: IsContextLostError (STRICT VERSION)
' ========================================================================================
Private Function IsContextLostError(ByVal response As String) As Boolean
    ' Early Exit: If the word "result" exists at the root, it is a protocol SUCCESS.
    ' This prevents treating "Shadow element not found" (a success response carrying an error string)
    ' as a context loss error.
    If InStr(1, response, """result"":", vbBinaryCompare) > 0 Then
        IsContextLostError = False
        Exit Function
    End If

    On Error Resume Next
    Dim json As Object: Set json = WebJsonConverter.ParseJson(response)
    If json Is Nothing Then Exit Function
    
    ' ONLY trigger if there is a root "error" key (BiDi Protocol level error)
    If json.Exists("error") Then
        Dim errType As String: errType = LCase(CStr(json("error")))
        Dim errMsg As String: errMsg = LCase(CStr(json("message")))
        
        Select Case errType
            Case "no such frame", "no such window", "unknown error"
                If InStr(errMsg, "context") > 0 Or InStr(errMsg, "gone") > 0 Or InStr(errMsg, "destroyed") > 0 Then
                    IsContextLostError = True
                    Exit Function
                End If
        End Select
    End If
    On Error GoTo 0
End Function

' ========================================================================================
' UpdateContextAndRealmRecursive: Recursively scans a node to update context and realm IDs.
' This ensures that even deeply nested IDs in script arguments are refreshed after recovery.
' ========================================================================================
Private Function UpdateContextAndRealmRecursive(ByRef node As Variant, _
                                                ByVal newCtx As String, _
                                                Optional ByVal newRealm As String = "") As Boolean
    On Error GoTo ErrH

    If IsObject(node) Then
        If TypeName(node) = "Dictionary" Then
            Dim k As Variant
            For Each k In node.keys
                If LCase(CStr(k)) = "context" Then
                    node(k) = newCtx
                ElseIf LCase(CStr(k)) = "realm" And Len(newRealm) > 0 Then
                    node(k) = newRealm
                End If
                ' Deep walk through nested parameters
                Call UpdateContextAndRealmRecursive(node(k), newCtx, newRealm)
            Next k
            UpdateContextAndRealmRecursive = True
            
        ElseIf TypeName(node) = "Collection" Then
            Dim i As Long
            For i = 1 To node.Count
                ' Recurse through collection items
                Call UpdateContextAndRealmRecursive(node(i), newCtx, newRealm)
            Next i
            UpdateContextAndRealmRecursive = True
        Else
            UpdateContextAndRealmRecursive = True
        End If
    Else
        UpdateContextAndRealmRecursive = True
    End If
    Exit Function

ErrH:
    UpdateContextAndRealmRecursive = False
End Function

' ========================================================================================
' UpdateContextAndRealm: Entry point for refreshing IDs within a parameter dictionary.
' ========================================================================================
Private Function UpdateContextAndRealm(ByRef params As Dictionary, ByVal newCtx As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Fetch a fresh Realm ID silently to prevent log spam during recovery
    Dim newRealm As String
    newRealm = GetScriptRealmId(newCtx, True, True) ' Silent:=True
    
    If newRealm = "" Then
        UpdateContextAndRealm = False
        Exit Function
    End If

    ' Delegate to the recursive walker to update all instances of context/realm keys
    UpdateContextAndRealm = UpdateContextAndRealmRecursive(params, newCtx, newRealm)
    Exit Function

ErrorHandler:
    UpdateContextAndRealm = False
End Function

' ========================================================================================
' DebugPrintLong: Prints structured logs with tags and handles long strings by chunking.
' Note: The 'tag' parameter is optional to maintain compatibility with existing calls.
' ========================================================================================
Private Sub DebugPrintLong(ByVal header As String, ByVal longText As String, _
                           Optional ByVal tag As String = "[INFO]")
    Const CHUNK_SIZE As Long = 250
    Dim i As Long
    Dim cleanText As String
    Dim timePrefix As String
    Dim currentTick As Double
    
    ' High-resolution timestamp for synchronization and performance analysis
    currentTick = GetRawTicks()
    timePrefix = "[" & Format(Now, "hh:mm:ss") & " | " & Format(currentTick, "0") & "] "

    ' Normalize text: remove null characters and trailing whitespace/newlines
    cleanText = Replace(longText, vbNullChar, "")
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
    Debug.Print timePrefix & tag & " " & header & sfx

    ' Chunked output to bypass VBA's Immediate Window line length limits
    If Len(cleanText) > 0 Then
        For i = 1 To Len(cleanText) Step CHUNK_SIZE
            ' Print parts of the string without an automatic newline
            Debug.Print Mid$(cleanText, i, CHUNK_SIZE);
        Next i
        Debug.Print ""
    End If
End Sub

' ========================================================================================
' ClearState: Resets internal state and cleans up browser-side side effects.
' ========================================================================================
Public Sub ClearState()
    p_cdpSessionId = ""
    p_mainContextId = ""
    p_pendingRequestCount = 0
    Set m_cachedRealmIds = Nothing
    Set m_inflightRequests = New Dictionary
    p_pendingRequestCount = 0
    
    'Reset the subscription registry.
    ' Using a New Dictionary ensures that the wrapper re-subscribes on the next connection.
    Set m_activeSubscriptions = New Dictionary
End Sub

' ========================================================================================
' FUNCTION: GetMainContextId
' DESCRIPTION: Returns the currently active main context ID.
'              If not yet set, it triggers a discovery of the browser tree.
' ========================================================================================
Public Function GetMainContextId() As String
    If p_mainContextId = "" Then
        ' MACHINERY: Pass True to the 'Silent' parameter to hide id:2
        GetMainContextId = UpdateMainContextId(True, True)
    Else
        GetMainContextId = p_mainContextId
    End If
End Function

' ========================================================================================
' FUNCTION: UpdateMainContextId
' SUMMARY: Refreshes the browsing context tree. Raises error if no context is available.
' ========================================================================================
Public Function UpdateMainContextId(Optional ByVal forceRefresh As Boolean = False, _
                                    Optional ByVal Silent As Boolean = False) As String
                                    
    If Not forceRefresh And p_mainContextId <> "" Then
        UpdateMainContextId = p_mainContextId
        Exit Function
    End If

    Dim jsonRes As String
    ' Suppress background noise during tree discovery
    jsonRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary, Silent)
    
    Dim jsonTree As Object: Set jsonTree = WebJsonConverter.ParseJson(jsonRes)
    Dim contextId As String: contextId = GetPathString(jsonTree, "result.contexts.0.context")
    
    If contextId <> "" Then
        p_mainContextId = contextId
        UpdateMainContextId = contextId
    Else
        ' Error must be reported even if Silent is True
        Err.Raise ERR_ACTION_FAILED, "BiDi:UpdateMainContextId", "Failed to identify any browsing context."
    End If
End Function

' Resolves the target browsing context ID.
' Returns the provided contextId if not empty, otherwise falls back to the main context.
Private Function ResolveTargetContext(ByVal contextId As String) As String
    If contextId <> "" Then
        ResolveTargetContext = contextId
    Else
        ' Returns cached main context ID or discovers a new one if necessary
        ResolveTargetContext = GetMainContextId()
    End If
End Function

' ========================================================================================
' FUNCTION: GetScriptRealmId (Improved for Service Worker environments)
' DESCRIPTION: Retrieves and caches the Execution Realm ID for a specific context.
'              Prioritizes "window" type realms to ensure UI/DOM script compatibility.
' ========================================================================================
Private Function GetScriptRealmId(Optional ByVal contextId As String = "", _
                                 Optional ByVal forceRefresh As Boolean = False, _
                                 Optional ByVal Silent As Boolean = False) As String
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = GetMainContextId()
    If targetCtx = "" Then Exit Function

    ' [FIX] Initialize the cache dictionary JIT (Just-In-Time)
    If m_cachedRealmIds Is Nothing Then Set m_cachedRealmIds = New Dictionary

    ' [FIX] Check for context-specific cached realm ID
    If Not forceRefresh Then
        If m_cachedRealmIds.Exists(targetCtx) Then
            GetScriptRealmId = m_cachedRealmIds(targetCtx)
            Exit Function
        End If
    End If

    Dim params As New Dictionary
    params.Add "context", targetCtx
    
    Dim jsonRes As String
    On Error Resume Next
    ' RELAY: Pass the Silent flag to suppress REQ/RES logs for realm discovery
    jsonRes = ExecuteBiDiCommand("script.getRealms", params, Silent)
    On Error GoTo 0
    
    Dim realms As Collection
    Set realms = GetJsonValue(jsonRes, "result.realms")
    
    Dim foundRealmId As String: foundRealmId = ""
    
    If Not realms Is Nothing Then
        If realms.Count > 0 Then
            Dim r As Dictionary
            For Each r In realms
                If r.Exists("type") Then
                    ' Priority: Window realms are standard for DOM manipulation
                    If r("type") = "window" Then
                        foundRealmId = CStr(r("realm"))
                        Exit For
                    End If
                End If
            Next r
            ' Fallback: Use the first available realm if no window realm is found
            If foundRealmId = "" Then foundRealmId = CStr(realms(1)("realm"))
        End If
    End If
    
    ' [FIX] Cache the found Realm ID using the contextId as the unique key
    If foundRealmId <> "" Then
        m_cachedRealmIds(targetCtx) = foundRealmId
    End If
    
    GetScriptRealmId = foundRealmId
End Function

' Gets the current CDP session ID, creating a new one if necessary.
Private Function GetCdpSessionId(Optional ByVal forceUpdate As Boolean = False) As String
    If p_mainContextId = "" Then UpdateMainContextId True, True
    
    If Not forceUpdate And p_cdpSessionId <> "" Then
        GetCdpSessionId = p_cdpSessionId: Exit Function
    End If

    Dim params As New Dictionary: params.Add "context", p_mainContextId
    
    ' Execute BiDi command Silently to hide internal session acquisition
    Dim strRes As String: strRes = ExecuteBiDiCommand("goog:cdp.getSession", params, True)
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    Dim sessionId As String: sessionId = GetPathString(json, "result.session")
    
    If sessionId <> "" Then
        p_cdpSessionId = sessionId: m_consecutiveCdpFailures = 0: GetCdpSessionId = p_cdpSessionId
    Else
        m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1: GetCdpSessionId = ""
    End If
End Function

' ========================================================================================
' FUNCTION: DetachCurrentCdpSession
' SUMMARY: Manually releases the CDP tunnel. Useful for switching contexts or resetting.
' ========================================================================================
Public Sub DetachCurrentCdpSession()
    If p_cdpSessionId = "" Then Exit Sub
    
    On Error Resume Next
    Dim params As New Dictionary
    params.Add "session", p_cdpSessionId
    ExecuteBiDiCommand "goog:cdp.detachSession", params
    On Error GoTo 0
    
    p_cdpSessionId = ""
    m_consecutiveCdpFailures = 0
End Sub

' ========================================================================================
' [PUBLIC] ExecuteRegisterAutoClickerByXPath
' DESCRIPTION: Registers a background observer via CDP to click a button as soon as it appears.
' ========================================================================================
Public Sub ExecuteRegisterAutoClickerByXPath(ByVal xpath As String, _
                                             Optional ByVal timeOutms As Long = 30000)
                                             
    ' Subscribe silently to log entries for AutoClicker confirmation
    Me.ExecuteSessionSubscribe "log.entryAdded", , True
    
    Dim sourceJs As String: sourceJs = GetActionJs("autoclicker")
    Dim safeXpath As String: safeXpath = Replace(xpath, "'", "\'")
    sourceJs = Replace(sourceJs, "{{XPATH}}", safeXpath)
    sourceJs = Replace(sourceJs, "{{TIMEOUT}}", CStr(timeOutms))
    
    Dim params As New Dictionary: params.Add "source", sourceJs
    
    Dim res As Object
    ' PASS Silent:=True to suppress the massive goog:cdp.sendCommand log
    Set res = ExecuteCDPCommand("Page.addScriptToEvaluateOnNewDocument", params, True, True)
    
    If res("Ok") Then
        Debug.Print "BiDi: AutoClicker registered for XPath: " & xpath
    Else
        Debug.Print "BiDi Error: AutoClicker registration failed: " & res("Error")
    End If
End Sub

' ========================================================================================
' INTERNAL HELPER: BuildShadowSelectorArg
' DESCRIPTION: Normalizes input to a JSON array. Accepts either a single String
'              or a VBA Array of Strings for nested Shadow DOM penetration.
' ========================================================================================
Private Function BuildShadowSelectorArg(ByVal selectorsArray As Variant) As Dictionary
    Dim arg As New Dictionary
    arg.Add "type", "array"
    
    Dim valList As New Collection
    Dim valItem As Dictionary
    
    ' Determine if input is a single value or an array
    If IsArray(selectorsArray) Then
        ' Case: Multiple selectors (e.g. Array("shadow-host", "button"))
        Dim i As Long
        For i = LBound(selectorsArray) To UBound(selectorsArray)
            Set valItem = New Dictionary
            valItem.Add "type", "string"
            valItem.Add "value", CStr(selectorsArray(i))
            valList.Add valItem
        Next i
    Else
        ' Case: Single selector (e.g. "button" or ".my-class")
        Set valItem = New Dictionary
        valItem.Add "type", "string"
        valItem.Add "value", CStr(selectorsArray)
        valList.Add valItem
    End If
    
    arg.Add "value", valList
    Set BuildShadowSelectorArg = arg
End Function

' ========================================================================================
' Discovery & Recorder API
' ========================================================================================
Public Sub StartDiscoveryLog(Optional ByVal excludeImagesAndCss As Boolean = True)
    Set p_recorderLogs = New Collection
    p_isRecording = True
    p_recorderStartTime = GetRawTicks() / 1000
    p_excludeResources = excludeImagesAndCss

    Dim events(3) As String
    events(0) = "network.beforeRequestSent"
    events(1) = "network.responseCompleted"
    events(2) = "browsingContext.load"
    events(3) = "browsingContext.domContentLoaded"
    Me.ExecuteSessionSubscribe events

    Debug.Print "BiDi Recorder: Started. (Filter Resources: " & excludeImagesAndCss & ")"
End Sub

' ========================================================================================
' [FIXED] StopAndSaveDiscoveryLog with explicit memory cleanup
' ========================================================================================
Public Sub StopAndSaveDiscoveryLog(ByVal filePath As String)
    ' [FIX] Flush the incoming message queue first.
    ' Since ProcessRecorderQueue calls DispatchGlobalEvents, which checks p_isRecording,
    ' we must process the queue while the flag is still True.
    Call ProcessRecorderQueue

    ' Now safely stop the recording flag.
    p_isRecording = False

    If p_recorderLogs Is Nothing Then Exit Sub
    
    ' Save logic
    If p_recorderLogs.Count > 0 Then
        Dim fso As New FileSystemObject
        Dim ts As TextStream
        
        ' Open file as UTF-16 (Unicode) to support various web content characters.
        Set ts = fso.CreateTextFile(filePath, True, True)

        ts.WriteLine "=== BiDi Discovery Log (Optimized for AI Analysis) ==="
        ts.WriteLine "Timestamp | Type | Method/Status | MIME | URL/Detail"
        ts.WriteLine "--------------------------------------------------------"

        Dim logItem As Variant
        For Each logItem In p_recorderLogs
            ts.WriteLine logItem
        Next logItem
        ts.Close
        Debug.Print "BiDi Recorder: Log saved to " & filePath
    End If

    ' Explicit cleanup to prevent memory leaks with nested objects in the Collection.
    Dim i As Long
    If Not p_recorderLogs Is Nothing Then
        For i = p_recorderLogs.Count To 1 Step -1
            p_recorderLogs.Remove i
        Next i
        Set p_recorderLogs = Nothing
    End If
End Sub

' ========================================================================================
' [PUBLIC] RecordEventsForSeconds
' SUMMARY: Keeps the WebSocket queue moving for a specified duration to record events.
' ========================================================================================
Public Sub RecordEventsForSeconds(ByVal seconds As Long)
    ' 1. Logic Check
    If seconds <= 0 Then Exit Sub

    ' --- [FIXED] UI LOCKING ---
    ' Disable user interaction to prevent entering "Cell Edit Mode" during recording.
    ' This is critical because VBA thread suspension during DoEvents can lead to
    ' protocol desync or crashes in the BiDi communicator.
    Application.Interactive = False
    Application.Cursor = xlWait

    ' 2. Initialize timing
    Dim startMs As Double: startMs = nowMs()
    Dim durationMs As Double: durationMs = CDbl(seconds) * 1000#

    On Error GoTo Cleanup

    ' 3. Polling Loop
    Do While (nowMs() - startMs) < durationMs
        ' Process the incoming WebSocket message queue
        ProcessRecorderQueue
        
        ' Release CPU briefly
        sleep 100
        
        ' Maintain Windows message pump
        ' Using ThrottledDoEvents for consistency with the class's stability policy.
        If durationMs > 5000 Then
            ThrottledDoEvents
        Else
            DoEvents
        End If
    Loop

Cleanup:
    ' --- [FIXED] UI RESTORATION ---
    ' Always ensure the UI is unlocked before returning control to the user.
    Application.Interactive = True
    Application.Cursor = xlDefault
    
    If Err.Number <> 0 Then
        Debug.Print "BiDi Warning in RecordEventsForSeconds: " & Err.Description
    End If
End Sub

' ========================================================================================
' PUBLIC SUB: ProcessRecorderQueue
' SUMMARY: Processes the incoming message queue with re-entrancy protection.
'          Uses a simple Boolean flag for state management (VBA is single-threaded).
' ========================================================================================
Public Sub ProcessRecorderQueue()
    ' --- 1. Re-entrancy Guard ---
    If m_isProcessingQueue Then Exit Sub
    
    ' Set the lock
    m_isProcessingQueue = True
    
    Dim msg As String
    Dim json As Object
    Dim processedCount As Long
    Dim consecutiveErrors As Long
    Const MAX_MESSAGES_PER_BATCH As Long = 1000
    
    On Error GoTo Cleanup
    
    ' Main processing loop
    Do While processedCount < MAX_MESSAGES_PER_BATCH
        msg = socket_.GetMessage()
        If Len(msg) = 0 Then Exit Do ' Buffer is empty
        
        ' Basic size/structure validation
        If Len(msg) >= MIN_MSG_SIZE And Len(msg) <= MAX_MSG_SIZE Then
            If InStr(msg, """method"":") > 0 Or InStr(msg, """id"":") > 0 Then
                
                Dim parseOk As Boolean: parseOk = False
                On Error Resume Next
                Set json = WebJsonConverter.ParseJson(msg)
                parseOk = (Err.Number = 0 And Not json Is Nothing)
                Err.Clear
                On Error GoTo Cleanup ' Restore main handler
                
                If parseOk Then
                    ' --- ISOLATED DISPATCH ---
                    ' Protect the loop from individual event processing failures
                    On Error Resume Next
                    DispatchGlobalEvents json
                    
                    If Err.Number <> 0 Then
                        consecutiveErrors = consecutiveErrors + 1
                        If DEBUG_MODE Then
                            Debug.Print "BiDi: Dispatch error #" & consecutiveErrors & ": " & Err.Description
                        End If
                        Err.Clear
                    Else
                        ' Reset counter on any successful dispatch
                        consecutiveErrors = 0
                    End If
                    On Error GoTo Cleanup ' Restore main handler
                    
                    ' --- CIRCUIT BREAKER ---
                    ' If the protocol is consistently failing, stop to prevent Excel hang.
                    If consecutiveErrors >= 10 Then
                        Debug.Print "BiDi Critical: 10 consecutive dispatch errors. Abandoning queue."
                        Exit Do
                    End If
                End If
            End If
        End If
        
        processedCount = processedCount + 1
        
        ' Yield to OS every 100 messages to keep UI responsive
        If (processedCount Mod 100) = 0 Then ThrottledDoEvents
    Loop

Cleanup:
    ' --- 2. Release Lock ---
    m_isProcessingQueue = False
    
    If Err.Number <> 0 Then
        Debug.Print "BiDi Critical Error in ProcessRecorderQueue: " & Err.Description
    End If
End Sub

Private Function IsIgnorableError(ByVal errType As String, ByVal errMsg As String) As Boolean
    Dim lMsg As String: lMsg = LCase(errMsg)

    Select Case errType
        Case "no such node", "no such frame", "stale element reference"
            IsIgnorableError = True
            Exit Function
    End Select

    If InStr(lMsg, "inspected target navigated or closed") > 0 Then IsIgnorableError = True: Exit Function
    If InStr(lMsg, "execution context was destroyed") > 0 Then IsIgnorableError = True: Exit Function
    If InStr(lMsg, "node is detached from document") > 0 Then IsIgnorableError = True: Exit Function
    If InStr(lMsg, "cannot find context") > 0 Then IsIgnorableError = True: Exit Function
    If InStr(lMsg, "stale element reference") > 0 Then IsIgnorableError = True: Exit Function

    IsIgnorableError = False
End Function

Private Sub FormatAndStoreLog(ByVal json As Object)
    ' Store selected BiDi events into the discovery log with optional noise filtering.
    If Not json.Exists("method") Then Exit Sub
    Dim method As String: method = json("method")

    Select Case method
        Case "network.beforeRequestSent", "network.responseCompleted", "browsingContext.load", "browsingContext.domContentLoaded", "log.entryAdded"
        Case Else
            Exit Sub
    End Select

    Dim params As Object
    Dim logLine As String
    Dim relTime As String

    If Not json.Exists("params") Then Exit Sub
    Set params = json("params")

    relTime = Format((GetRawTicks() / 1000) - p_recorderStartTime, "0.00") & "s"

    Select Case method
        Case "network.beforeRequestSent"
            ' Request logging with resource noise filtering when enabled.
            Dim req As Object: Set req = params("request")
            Dim url As String: url = req("url")

            If p_excludeResources Then
                ' URL-based generic noise filter
                If IsNoiseUrl(url) Then Exit Sub

                ' Header-based optional filter: skip pure image requests when Accept header indicates image/*
                On Error Resume Next
                If req.Exists("headers") Then
                    Dim acceptHdr As String
                    acceptHdr = LCase(req("headers")("Accept"))
                    If InStr(acceptHdr, "image/") > 0 Then Exit Sub
                End If
                On Error GoTo 0
            End If

            logLine = "[" & relTime & "] [REQ] " & req("method") & " - " & url

        Case "network.responseCompleted"
            ' Response logging with MIME and URL noise filtering when enabled.
            Dim res As Object: Set res = params("response")
            Dim resUrl As String: resUrl = res("url")
            Dim mime As String: mime = res("mimeType")

            If p_excludeResources Then
                If IsNoiseUrl(resUrl) Or IsNoiseMime(mime) Then Exit Sub
            End If

            logLine = "[" & relTime & "] [RES] " & res("status") & " [" & mime & "] " & resUrl

        Case "browsingContext.load", "browsingContext.domContentLoaded"
            ' Navigation milestones
            logLine = "[" & relTime & "] [NAV] " & Replace(method, "browsingContext.", "") & " - " & params("url")

        Case "log.entryAdded"
            ' Console error entries only
            If params("level") = "error" Then
                logLine = "[" & relTime & "] [CONSOLE-ERR] " & params("text")
            End If
    End Select

    If logLine <> "" Then
        If p_recorderLogs.Count >= MAX_LOG_SIZE Then p_recorderLogs.Remove 1
        p_recorderLogs.Add logLine
    End If
End Sub

' Determines whether a URL should be excluded from Discovery Logs and stability checks.
' This version is updated to filter out framework-specific noise (Next.js) and ads found in recent logs.
Private Function IsNoiseUrl(ByVal url As String) As Boolean
    Dim lUrl As String: lUrl = LCase(url)
    
    ' --- 1. Framework & Static Resource Noise ---
    If InStr(lUrl, "/_next/static/") > 0 Or _
       InStr(lUrl, "webworker.js") > 0 Or _
       InStr(lUrl, ".png") > 0 Or InStr(lUrl, ".jpg") > 0 Or InStr(lUrl, ".jpeg") > 0 Or _
       InStr(lUrl, ".gif") > 0 Or InStr(lUrl, ".webp") > 0 Or InStr(lUrl, ".svg") > 0 Or _
       InStr(lUrl, ".css") > 0 Or InStr(lUrl, ".woff") > 0 Or InStr(lUrl, ".woff2") > 0 Or _
       InStr(lUrl, ".ico") > 0 Or Left(lUrl, 5) = "data:" Then
        IsNoiseUrl = True
        Exit Function
    End If
    
    ' --- 2. Ads, Analytics & Telemetry Noise ---
    If InStr(lUrl, "google-analytics") > 0 Or _
       InStr(lUrl, "googletagmanager") > 0 Or _
       InStr(lUrl, "doubleclick") > 0 Or _
       InStr(lUrl, "pagead") > 0 Or _
       InStr(lUrl, "conversion") > 0 Or _
       InStr(lUrl, "/collect") > 0 Or _
       InStr(lUrl, "/analytics") > 0 Or _
       InStr(lUrl, "/pixel") > 0 Or _
       InStr(lUrl, "/beacon") > 0 Then
        IsNoiseUrl = True
        Exit Function
    End If

    ' --- 3. Dynamic User-Defined Filters ---
    ' Ensures patterns added via AddIdleIgnoreNetworkPattern are also respected here.
    If Not p_userIgnoreNetworkPatterns Is Nothing Then
        Dim pattern As Variant
        For Each pattern In p_userIgnoreNetworkPatterns
            If InStr(1, lUrl, CStr(pattern), vbTextCompare) > 0 Then
                IsNoiseUrl = True
                Exit Function
            End If
        Next pattern
    End If

    IsNoiseUrl = False
End Function

Private Function IsNoiseMime(ByVal mime As String) As Boolean
    Dim lMime As String: lMime = LCase(mime)
    IsNoiseMime = (InStr(lMime, "image/") > 0 Or InStr(lMime, "font/") > 0 Or InStr(lMime, "css") > 0)
End Function

' ========================================================================================
' FUNCTION: QuerySpaProbe
' DESCRIPTION: Executes the JS query snippet and parses the result.
'              Returns Nothing if the probe is missing or reports 'active: false'.
' ========================================================================================
Private Function QuerySpaProbe(Optional ByVal contextId As String = "") As Dictionary
    Dim strRes As String
    ' Pass Silent:=True to CallScript
    strRes = CallScript(GetProbeJs("query"), Nothing, contextId, False, False, 10000, True)
    
    Dim innerJsonStr As String: innerJsonStr = CStr(GetCleanResult(strRes))
    If innerJsonStr = "" Then Set QuerySpaProbe = Nothing: Exit Function

    On Error Resume Next
    Set QuerySpaProbe = WebJsonConverter.ParseJson(innerJsonStr)
    On Error GoTo 0
End Function

' ========================================================================================
' PROCEDURE: RecoverSpaProbe
' SUMMARY: Nukes the state cache and re-injects the probe into the fresh context.
' ========================================================================================
Private Sub RecoverSpaProbe(ByVal contextId As String)
    Dim targetId As String: targetId = contextId
    Dim isMainTarget As Boolean
    
    ' Determine if we are recovering the main context (empty or matching current main)
    isMainTarget = (targetId = "" Or targetId = p_mainContextId)
    
    ' 1. Invalidate ONLY the specific context's realm ID cache.
    ' Rationale: By using m_cachedRealmIds (Dictionary), we ensure that recovering
    ' one tab does not accidentally destroy valid sessions/realms in other tabs.
    If Not m_cachedRealmIds Is Nothing Then
        If m_cachedRealmIds.Exists(targetId) Then
            m_cachedRealmIds.Remove targetId
        End If
    End If
    
    ' 2. Force a full tree refresh to handle cases where the tab navigated to
    ' a completely different origin or the internal BiDi context mapping was lost.
    ' We capture the result of the discovery to get the most current ID.
    Dim refreshedMainId As String: refreshedMainId = UpdateMainContextId(True, True)
    
    ' 3. [FIX] ID Synchronization
    ' If we were targeting the main context, we must use the refreshed ID.
    ' Otherwise, StartSpaIdleProbe would target a stale/dead context ID.
    If isMainTarget Then
        targetId = refreshedMainId
    End If
    
    ' 4. Re-inject. StartSpaIdleProbe will now target the confirmed active context.
    ' This ensures the SPA monitoring script is active on the new page state.
    StartSpaIdleProbe targetId
    
    If DEBUG_MODE Then Debug.Print "BiDi: Recovery completed for context: " & targetId
End Sub

' ========================================================================================
' PUBLIC SUB: WaitForNetworkIdleSpa
' SUMMARY: Entry point for deep SPA-aware synchronization.
'          This version includes Entry Guards to handle 0ms inputs caused by the
'          updated SanitizeMs logic, ensuring stability consensus isn't bypassed.
' ========================================================================================
Public Sub WaitForNetworkIdleSpa(ByVal minStableMs As Long, _
                                 ByVal maxTimeoutMs As Long, _
                                 Optional ByVal contextId As String = "")
    
    ' --- [ENTRY GUARD] Guard against 0ms in stability check ---
    ' Rationale: 0ms stability is logically inconsistent for SPA synchronization.
    ' If SanitizeMs passes 0, we must revert to defaults for this specific logic
    ' to prevent the consensus engine from exiting prematurely.
    If minStableMs = 0 Then minStableMs = DEFAULT_IDLE_STABLE_MS
    If maxTimeoutMs = 0 Then maxTimeoutMs = DEFAULT_IDLE_TIMEOUT_MS
    ' ----------------------------------------------------------
    
    ' Resolve the target browsing context (tab/window)
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' Reset the JS heartbeat tracker to ensure fresh monitoring for this loop
    m_lastJsNow = 0
    
    ' Process any pending WebSocket messages to sync current network request counts
    ' before entering the stability consensus phase.
    ProcessRecorderQueue

    ' Initialize the wait context dictionary with tracking metrics for SPA stability.
    ' These keys are used defensively by EvaluateWaitCondition.
    Dim waitCtx As New Dictionary
    waitCtx.Add "contextId", targetCtx
    waitCtx.Add "minStableMs", minStableMs
    waitCtx.Add "stableStart", nowMs()
    waitCtx.Add "missingProbeCount", 0 ' Used for multi-stage JIT recovery
    
    ' Execute the high-precision wait engine with SPA Idle mode.
    ' MasterWaitEngine will perform final sanitization using the guarded values.
    MasterWaitEngine WaitMode_SpaIdle, maxTimeoutMs, 100, waitCtx
End Sub

' ========================================================================================
' FUNCTION: EvaluateSpaIdleCondition
' DESCRIPTION: Final consensus engine for SPA stability.
'              Integrates:
'              1. Network Silence (BiDi Protocol + JS Hooks)
'              2. DOM Stability (MutationObserver + Prototype Hooks)
'              3. Diagnostic Logging (Immediate Window + Discovery Recorder)
' ========================================================================================
Private Function EvaluateSpaIdleCondition(ByVal probe As Dictionary, _
                                          ByRef ctx As Dictionary, _
                                          ByVal currentMs As Double) As Boolean
    Dim jsNow As Double, lastActivityTs As Double, jsIdleDuration As Double
    Dim inflFetch As Long, inflXhr As Long
    Dim netIdle As Boolean, activityStable As Boolean

    ' 1. Extract threshold and timing states from the context container
    Dim minStableMs As Long: minStableMs = CLng(ctx("minStableMs"))
    
    ' 2. Extract performance metrics from the JavaScript probe
    jsNow = CDbl(probe("now"))
    lastActivityTs = CDbl(probe("lastMutationTs"))
    inflFetch = CLng(probe("inflightFetchCount"))
    inflXhr = CLng(probe("inflightXhrCount"))
    
    ' Calculate current idle duration reported by the browser
    jsIdleDuration = jsNow - lastActivityTs

    ' --- 3. Activity Logging (Discovery Recorder Integration) ---
    If probe.Exists("changes") Then
        Dim changes As Collection: Set changes = probe("changes")
        If Not changes Is Nothing Then
            If changes.Count > 0 Then
                Dim v As Variant, logSummary As String
                For Each v In changes: logSummary = logSummary & v & " ": Next
                logSummary = Trim(logSummary)
                
                If Len(logSummary) > 500 Then logSummary = Left$(logSummary, 497) & "..."
                If DEBUG_MODE Then Debug.Print "[BiDi SPA Activity] " & logSummary
                
                If p_isRecording Then
                    Dim relTime As String
                    relTime = Format((GetRawTicks() / 1000) - p_recorderStartTime, "0.00") & "s"
                    p_recorderLogs.Add "[" & relTime & "] [SPA-ACT] " & logSummary
                End If
            End If
        End If
    End If

    ' --- 4. Self-Correction Logic (Hard-Sync) ---
    ' [FIX] Execute sync BEFORE netIdle evaluation to prevent Consensus Yo-Yo.
    ' If JS reports zero activity but VBA has phantom counts, force sync after 2x silence.
    If (inflFetch = 0 And inflXhr = 0) And (p_pendingRequestCount > 0) Then
        If jsIdleDuration > (CDbl(minStableMs) * 2#) Then
            If DEBUG_MODE Then Debug.Print "BiDi: Hard-Syncing pendingRequestCount to 0 due to prolonged silence."
            p_pendingRequestCount = 0
        End If
    End If

    ' --- 5. Consensus Evaluation ---
    ' Network Idle: Both protocol and JS hooks must report zero activity.
    netIdle = (p_pendingRequestCount <= 0) And (inflFetch = 0) And (inflXhr = 0)

    ' DOM Stability: No mutations within the required threshold.
    activityStable = (jsIdleDuration >= CDbl(minStableMs))

    ' --- 6. Final Consensus Window Check ---
    If netIdle And activityStable Then
        ' Check if the stable state has persisted for the designated window.
        If (currentMs - CDbl(ctx("stableStart"))) >= CDbl(minStableMs) Then
            If DEBUG_MODE Then
                Call DebugPrintLong("Consensus reached (" & minStableMs & "ms window satisfied)", "", ">>> [BiDi SPA Stable]")
            End If
            
            If p_isRecording Then
                Dim successRelTime As String
                successRelTime = Format((GetRawTicks() / 1000) - p_recorderStartTime, "0.00") & "s"
                p_recorderLogs.Add "[" & successRelTime & "] [SPA-STABLE] Consensus reached"
            End If
            
            EvaluateSpaIdleCondition = True ' Stability confirmed
        Else
            EvaluateSpaIdleCondition = False ' Cooling down...
        End If
    Else
        ' Active state detected: Reset the window start timer.
        ctx("stableStart") = currentMs
        EvaluateSpaIdleCondition = False
    End If
End Function

' ========================================================================================
' ExecuteCDPCommand: Sends a raw CDP command via BiDi tunneling.
' ========================================================================================
Private Function ExecuteCDPCommand(ByVal cdpMethod As String, _
                                   ByVal cdpParams As Dictionary, _
                                   Optional ByVal useSession As Boolean = True, _
                                   Optional ByVal Silent As Boolean = False) As Object
    Dim res As New Dictionary
    res.Add "Ok", False: res.Add "Json", "": res.Add "Error", ""

    If socket_ Is Nothing Then
        res("Error") = "Socket object is not set."
        Set ExecuteCDPCommand = res: Exit Function
    End If

    ' --- [FIX] Trigger state reset if CDP failures exceed limit ---
    If m_consecutiveCdpFailures >= MAX_CONSECUTIVE_CDP_FAILURES Then
        ' Correctly calls the actual Sub PerformFatalReset instead of jumping to a label
        Call PerformFatalReset
        res("Error") = "Fatal: CDP instability. Context state has been cleared."
        Set ExecuteCDPCommand = res
        Exit Function
    End If

    Dim params As New Dictionary
    params.Add "method", cdpMethod
    params.Add "params", EnsureDictionary(cdpParams)

    If useSession Then
        Dim sessionId As String: sessionId = GetCdpSessionId()
        If sessionId <> "" Then
            params.Add "session", sessionId
        Else
            m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
            res("Error") = "CDP session unavailable."
            Set ExecuteCDPCommand = res: Exit Function
        End If
    End If

    Dim strRes As String
    On Error Resume Next
    strRes = ExecuteBiDiCommand("goog:cdp.sendCommand", params, Silent)
    
    If Err.Number <> 0 Then
        res("Error") = "Socket Error: " & Err.Description
        m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
        On Error GoTo 0: Set ExecuteCDPCommand = res: Exit Function
    End If
    On Error GoTo 0

    ' Evaluate internal CDP protocol errors
    If InStr(strRes, """error"":") > 0 Then
        res("Error") = "CDP Error: " & strRes: res("Json") = strRes
        m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
        DebugPrintLong "CDP Command FAILED: " & cdpMethod, strRes, "[CDP:ERR]"
    Else
        res("Ok") = True: res("Json") = strRes: m_consecutiveCdpFailures = 0
        If DEBUG_MODE And Not Silent Then
            DebugPrintLong "CDP Command OK: " & cdpMethod, strRes, "[CDP:OK]"
        End If
    End If

    Set ExecuteCDPCommand = res
End Function

' Internal helper to clear all state and stop the bleed
Private Sub PerformFatalReset()
    Debug.Print "BiDi Critical: Performing Fatal Reset due to protocol instability."
    ClearState
    m_consecutiveCdpFailures = 0
    ' Optionally, close the socket if the connection itself is suspect
    ' If Not socket_ Is Nothing Then socket_.CloseConnection
End Sub

' ========================================================================================
' PROCEDURE: StartSpaIdleProbe
' DESCRIPTION: Executes the JavaScript probe injection.
' ========================================================================================
Private Sub StartSpaIdleProbe(Optional ByVal contextId As String = "")
    Dim targetCtx As String
    ' Ensure the tree refresh is also silent if triggered
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId(True, True)
    
    ' MANDATORY: Pass Silent:=True to GetScriptRealmId to kill id:16
    Dim targetRealm As String
    targetRealm = GetScriptRealmId(targetCtx, True, True)
    
    Dim p As New Dictionary
    p.Add "expression", GetProbeJs("inject")
    p.Add "target", BuildContextTarget(targetCtx, targetRealm)
    p.Add "awaitPromise", False

    ' Execute the injection silently
    Call ExecuteBiDiCommand("script.evaluate", p, True)
End Sub

' ========================================================================================
' JavaScript Error Decoder (Enhanced Debugging)
' ========================================================================================

' Formats BiDi exceptionDetails into a readable string including stack trace
Private Function FormatJavaScriptError(ByVal details As Object) As String
    Dim msg As String
    Dim stack As String
    Dim frames As Collection
    Dim frame As Object
    Dim i As Long

    ' 1. Basic error info
    msg = "JS Exception: " & details("text")
    
    ' 2. Location info
    If details.Exists("lineNumber") Then
        msg = msg & " [Line: " & details("lineNumber") & ", Col: " & details("columnNumber") & "]"
    End If

    ' 3. Exception object description (if exists)
    If details.Exists("exception") Then
        If details("exception").Exists("description") Then
            msg = msg & vbCrLf & "Description: " & details("exception")("description")
        End If
    End If

    ' 4. Stack Trace decoding
    If details.Exists("stackTrace") Then
        Set frames = details("stackTrace")("callFrames")
        stack = "--- JS Stack Trace ---"
        For i = 1 To frames.Count
            Set frame = frames(i)
            stack = stack & vbCrLf & "  at " & frame("functionName") & " (" & frame("url") & ":" & frame("lineNumber") & ":" & frame("columnNumber") & ")"
        Next i
        msg = msg & vbCrLf & stack
    End If

    FormatJavaScriptError = msg
End Function

' ========================================================================================
' ThrottledDoEvents: Processes Windows events at a controlled frequency.
' Prevents UI flickering and reduces the risk of re-entrancy during tight wait loops.
' ========================================================================================
Private Sub ThrottledDoEvents()
    ' Check if more than 100ms has elapsed since the last event pump.
    ' Now uses Currency-based ticks for high-precision comparison via GetElapsedMs.
    If GetElapsedMs(m_lastDoEventsTime) > 100 Then
        DoEvents
        ' Record the current performance counter for the next interval check.
        m_lastDoEventsTime = GetRawTicks()
    End If
End Sub

' ========================================================================================
' AddIdleIgnoreSelector: Registers a CSS selector to be ignored by MutationObserver.
' Users can call this from a Standard Module to bypass noisy elements (e.g., clocks).
' ========================================================================================
Public Sub AddIdleIgnoreSelector(ByVal cssSelector As String)
    If p_userIgnoreSelectors Is Nothing Then Set p_userIgnoreSelectors = New Collection
    p_userIgnoreSelectors.Add cssSelector
End Sub

' ========================================================================================
' AddIdleIgnoreNetworkPattern: Registers a URL substring to be ignored by Network tracking.
' Useful for bypassing background logging, telemetry, or specific third-party ads.
' ========================================================================================
Public Sub AddIdleIgnoreNetworkPattern(ByVal urlPattern As String)
    If p_userIgnoreNetworkPatterns Is Nothing Then Set p_userIgnoreNetworkPatterns = New Collection
    p_userIgnoreNetworkPatterns.Add urlPattern
End Sub

' ========================================================================================
' ClearIdleFilters: Resets all user-defined noise filters.
' ========================================================================================
Public Sub ClearIdleFilters()
    Set p_userIgnoreSelectors = Nothing
    Set p_userIgnoreNetworkPatterns = Nothing
End Sub

' ========================================================================================
' MergeToJsArray: Combines a comma-separated default string and a VBA Collection
' into a JSON-formatted JavaScript array string.
' ========================================================================================
Private Function MergeToJsArray(ByVal defaultItems As String, ByVal userItems As Collection) As String
    Dim items As New Collection
    Dim dArr() As String: dArr = Split(defaultItems, ",")
    Dim i As Long
    
    ' 1. Add defaults
    For i = LBound(dArr) To UBound(dArr)
        items.Add Trim(dArr(i))
    Next i
    
    ' 2. Add user definitions
    If Not userItems Is Nothing Then
        Dim v As Variant
        For Each v In userItems
            items.Add CStr(v)
        Next v
    End If
    
    ' 3. Serialize to JS Array [ "a", "b" ]
    Dim result As String: result = "["
    For i = 1 To items.Count
        result = result & """" & Replace(items(i), """", "\""") & """"
        If i < items.Count Then result = result & ","
    Next i
    result = result & "]"
    
    MergeToJsArray = result
End Function

' ========================================================================================
' FUNCTION: GetJsonValue
' DESCRIPTION: Extracts a value from a nested JSON string using dot-notation.
' [FIXED] Object-Safe Traversal: Prevents Error 91/424 when encountering primitive nodes.
' ========================================================================================
Private Function GetJsonValue(ByVal jsonStr As String, ByVal keyPath As String) As Variant
    Dim json As Object
    On Error Resume Next
    
    ' 1. Parse the JSON string
    Set json = WebJsonConverter.ParseJson(jsonStr)
    If json Is Nothing Then Exit Function

    Dim keys() As String: keys = Split(keyPath, ".")
    Dim target As Object: Set target = json
    Dim i As Long
    Dim currentKey As String

    ' 2. Traverse the tree
    ' Loop until the second-to-last key in the path
    For i = 0 To UBound(keys) - 1
        currentKey = keys(i)
        
        If TypeName(target) = "Dictionary" Then
            If Not target.Exists(currentKey) Then Exit Function
            
            ' [CRITICAL FIX] Verify node is an Object before using 'Set'
            ' to avoid Error 91 if currentKey holds a primitive value.
            If IsObject(target(currentKey)) Then
                Set target = target(currentKey)
            Else
                Exit Function ' Path demands deeper traversal, but we hit a primitive
            End If
        
        ElseIf TypeName(target) = "Collection" Then
            ' Support array indexing. JSON is 0-based, VBA Collection is 1-based.
            If IsNumeric(currentKey) Then
                Dim idx As Long: idx = CLng(currentKey) + 1
                If idx < 1 Or idx > target.Count Then Exit Function
                
                ' [CRITICAL FIX] Verify node is an Object before using 'Set'
                If IsObject(target(idx)) Then
                    Set target = target(idx)
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Else
            ' target is neither Dictionary nor Collection; cannot traverse further
            Exit Function
        End If
    Next i

    ' 3. Extract the final leaf value
    If Not target Is Nothing Then
        Dim lastKey As String: lastKey = keys(UBound(keys))
        
        If TypeName(target) = "Dictionary" Then
            If target.Exists(lastKey) Then
                ' Handle the leaf value correctly (Set for objects, Let for primitives)
                If IsObject(target(lastKey)) Then
                    Set GetJsonValue = target(lastKey)
                Else
                    GetJsonValue = target(lastKey)
                End If
            End If
        ElseIf TypeName(target) = "Collection" And IsNumeric(lastKey) Then
            Dim fIdx As Long: fIdx = CLng(lastKey) + 1
            If fIdx >= 1 And fIdx <= target.Count Then
                If IsObject(target(fIdx)) Then
                    Set GetJsonValue = target(fIdx)
                Else
                    GetJsonValue = target(fIdx)
                End If
            End If
        End If
    End If
    On Error GoTo 0
End Function

' ========================================================================================
' INTERNAL JS HELPER: GetActionJs
' DESCRIPTION: Provides centralized JavaScript snippets for all browser actions.
'              All actions return a stringified JSON to ensure robust error handling in VBA.
' ========================================================================================
Private Function GetActionJs(ByVal actionType As String) As String
    Dim js As String: js = ""
    Select Case LCase(actionType)
        Case "click"
            ' Scrolls, focuses, and clicks. Standardized JSON response.
            js = js & "try { const e = arguments[0]; e.scrollIntoView({block:'center'}); e.focus(); e.click(); return JSON.stringify({status: 'ok'}); } "
            js = js & "catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"
                 
        Case "input"
            ' SPA-compatible input using execCommand with bubbling event fallback.
            js = js & "try { const e = arguments[0]; const v = arguments[1]; e.scrollIntoView({block:'center',inline:'center'}); e.click(); e.focus(); e.value=''; "
            js = js & "const s = document.execCommand('insertText', false, v); if(!s){ e.value = v; e.dispatchEvent(new Event('input', {bubbles:true})); e.dispatchEvent(new Event('change', {bubbles:true})); } "
            js = js & "e.blur(); return JSON.stringify({status: 'ok'}); } catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"
                 
        Case "select"
            ' Standard Select-Box handling with event dispatching.
            js = js & "try { const s = arguments[0]; const v = arguments[1]; s.scrollIntoView({block:'center',inline:'center'}); s.focus(); s.value = v; "
            js = js & "s.dispatchEvent(new Event('input', {bubbles:true})); s.dispatchEvent(new Event('change', {bubbles:true})); s.blur(); return JSON.stringify({status: 'ok'}); } "
            js = js & "catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"

        Case "select_text"
            ' Normalizes text nodes to handle non-breaking spaces & dynamic content.
            js = js & "try { const s = arguments[0]; const t = arguments[1]; const n = (str) => str.replace(/[\s\u00A0]+/g, ' ').trim(); let f = false; const target = n(t); "
            js = js & "for (let i = 0; i < s.options.length; i++) { if (n(s.options[i].text) === target) { s.value = s.options[i].value; f = true; break; } } "
            js = js & "if (!f) return JSON.stringify({status: 'error', message: 'Option text not found: ' + t}); s.scrollIntoView({block:'center'}); s.focus(); "
            js = js & "s.dispatchEvent(new Event('input', {bubbles:true})); s.dispatchEvent(new Event('change', {bubbles:true})); s.blur(); return JSON.stringify({status: 'ok'}); } "
            js = js & "catch(err) { return JSON.stringify({status: 'error', message: err.message}); }"

        Case "visibility_check"
            ' Visual state validation (Display, Visibility, Opacity, and DOM attachment).
            js = js & "const e = arguments[0]; if (!e || !e.isConnected) return JSON.stringify({status: 'ok', value: false}); const s = window.getComputedStyle(e), r = e.getBoundingClientRect(); "
            js = js & "const v = (s.display !== 'none' && s.visibility !== 'hidden' && parseFloat(s.opacity || '1') > 0 && (r.width > 0 || r.height > 0 || e.getClientRects().length > 0)); "
            js = js & "return JSON.stringify({status: 'ok', value: v});"
   
        Case "ready_state"
            ' Standardized readyState query for the GetCleanResult pipeline.
            js = js & "return JSON.stringify({status: 'ok', value: document.readyState});"

        Case "shadow_click"
            ' ASYNC: Recursively penetrates ShadowRoots with heartbeat and timeout support.
            ' [FIX] Removed 'const hb = arguments[0];'. hb() is now provided via closure.
            js = js & "const selectors = arguments[0]; const timeout = arguments[1]; const find = (arr, t) => { "
            js = js & "return new Promise((res) => { const end = Date.now() + t; const check = () => { hb(); let el = document.querySelector(arr[0]); "
            js = js & "if(el) { for(let i=1; i<arr.length; i++){ if(el.shadowRoot) el = el.shadowRoot.querySelector(arr[i]); else { el=null; break; } if(!el) break; } } "
            js = js & "if(el) res(el); else if(Date.now() < end) setTimeout(check, 100); else res(null); }; check(); }); }; "
            js = js & "const targetEl = await find(selectors, timeout); if(!targetEl) throw new Error('Shadow element not found'); "
            js = js & "targetEl.scrollIntoView({block:'center'}); targetEl.focus(); targetEl.click(); return JSON.stringify({status: 'ok'});"
        
        Case "autoclicker"
            ' IIFE: Background MutationObserver that fires a click as soon as target appears.
            js = js & "(function(x,t){const s=Date.now(),f=()=>{const e=document.evaluate(x,document,null,9,null).singleNodeValue;"
            js = js & "if(e&&(e.offsetWidth>0||e.offsetHeight>0)){e.click();e.dispatchEvent(new MouseEvent('click',{bubbles:true}));"
            js = js & "console.log('BiDi-AutoClicker: Target clicked');return 1}return 0};"
            js = js & "if(f())return;const o=new MutationObserver(()=>{if(f()||(Date.now()-s>t))o.disconnect()});"
            js = js & "o.observe(document.body||document,{childList:1,subtree:1,attributes:1})})('{{XPATH}}',{{TIMEOUT}});"
            
    End Select
    GetActionJs = js
End Function

' ========================================================================================
' [PRIVATE] CallScript (Refactored for Anti-Freeze)
' NOTE TO DEVELOPERS:
' This method automatically wraps your JS code in a 'SafetyShell' watchdog.
' If your script is expected to run longer than 10 seconds, you MUST use the heartbeat reset.
' ========================================================================================
Public Function CallScript(ByVal js As String, _
                           Optional ByVal args As Collection = Nothing, _
                           Optional ByVal contextId As String = "", _
                           Optional ByVal awaitPromise As Boolean = False, _
                           Optional ByVal forceRealmTarget As Boolean = False, _
                           Optional ByVal safetyTimeoutMs As Long = 10000, _
                           Optional ByVal Silent As Boolean = False) As String
    
    Dim params As New Dictionary, targetDic As New Dictionary
    If args Is Nothing Then Set args = New Collection
    Dim resolvedCtx As String: resolvedCtx = ResolveTargetContext(contextId)
    
    If resolvedCtx = "" Then
        CallScript = "{""error"":""no_context"",""message"":""Failed to resolve contextId.""}"
        Exit Function
    End If

    Dim finalJs As String: finalJs = WrapInSafetyShell(js, safetyTimeoutMs)

    If forceRealmTarget Then
        ' FIX: Pass the Silent flag to the internal realm discovery call
        targetDic.Add "realm", GetScriptRealmId(resolvedCtx, False, Silent)
    Else
        targetDic.Add "context", resolvedCtx
    End If
    
    params.Add "functionDeclaration", finalJs
    params.Add "arguments", args
    params.Add "target", targetDic
    params.Add "awaitPromise", True
    
    ' Execute command with the Silent flag
    Dim rawRes As String: rawRes = ExecuteBiDiCommand("script.callFunction", params, Silent)
    
    Dim procObj As Object: Set procObj = ProcessResponse(rawRes, "script.callFunction")
    CallScript = WebJsonConverter.ConvertToJson(procObj)
End Function

' ========================================================================================
' IsTimeout: Precision-safe timeout check.
' ========================================================================================
Private Function IsTimeout(ByVal startTicks As Currency, ByVal timeOutms As Double) As Boolean
    If timeOutms <= 0 Then Exit Function
    ' Simple comparison
    IsTimeout = GetElapsedMs(startTicks) > timeOutms
End Function

' ========================================================================================
' Internal Helper: Centralizes background tasks during a wait loop.
' Keeps the WebSocket queue moving and the Excel UI responsive.
' ========================================================================================
Private Sub WaitHeartbeat(Optional ByVal sleepMs As Long = 50)
    ' 1. Process incoming BiDi events (Updates pending request counts, etc.)
    ' Rationale: This ensures the VBA state stays in sync with the browser's network/DOM state.
    ProcessRecorderQueue
    
    ' 2. Process Windows events (Keeps Excel from freezing)
    ' Rationale: Allows Excel to paint and handle background messages.
    ThrottledDoEvents
    
    ' 3. [FIXED] Release CPU briefly with Floor protection.
    ' Rationale: Even if 0 or a very low value is requested, we enforce the safety floor
    ' (default 50ms) to ensure the Windows message pump has sufficient time to breathe.
    Dim vSleep As Long
    vSleep = SanitizeMs(sleepMs, "WaitHeartbeat", 50, , 50)
    
    If vSleep > 0 Then sleep vSleep
End Sub

' ========================================================================================
' FUNCTION: GetProbeJs
' DESCRIPTION: Provides the comprehensive JavaScript payload for SPA monitoring.
'              Integrated with Smart XPath generation (Ancestor-ID based).
' ========================================================================================
Private Function GetProbeJs(ByVal probeType As String) As String
    Dim js As String: js = ""
    Dim selArray As String: selArray = MergeToJsArray(DEFAULT_IGNORE_SELECTORS, p_userIgnoreSelectors)
    Dim netArray As String: netArray = MergeToJsArray(DEFAULT_IGNORE_NETWORK, p_userIgnoreNetworkPatterns)

    Select Case LCase$(probeType)
        Case "inject"
            ' --- SPA Monitoring Probe: Now filters out Sizzle noise ---
            js = js & "(function(){if(window.__vbaSpaProbeInstalled)return true;window.__vbaSpaProbeInstalled=true;var selectors=" & selArray & ",ignoreNetwork=" & netArray & ","
            js = js & "p=window.__vbaIdleProbe||{inflightFetchCount:0,inflightXhrCount:0,lastMutationTs:performance.now(),changeList:[]};window.__vbaIdleProbe=p;"
            js = js & "var gX=function(e){if(!e||e.nodeType!==1)return '';var s=[];while(e&&e.nodeType===1){if(e.id){s.unshift('//*[@id=\''+e.id+'\']');return s.join('/')}"
            js = js & "var i=1,b=e.previousSibling;while(b){if(b.nodeType===1&&b.tagName===e.tagName)i++;b=b.previousSibling}s.unshift(e.tagName.toLowerCase()+'['+i+']');e=e.parentNode}return '/'+s.join('/')};"
            
            ' LG Modification: Ignore technical noise like "sizzle" or dynamic session markers
            js = js & "var uT=()=>{p.lastMutationTs=performance.now()},lg=(e)=>{"
            js = js & "if(e && !/sizzle|signon_javascript|conversation_button/.test(e) && p.changeList.indexOf(e)===-1){"
            js = js & "if(p.changeList.length<50)p.changeList.push(e);uT()}};"
            js = js & "iN=(n,s)=>{if(!n)return 1;var t=(n.nodeType===1)?n:n.parentElement;return t?s.some((x)=>t.closest(x)!==null):0};"
            js = js & "try{const d=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value');Object.defineProperty(HTMLInputElement.prototype,'value',{set:function(v){lg('PROP:'+gX(this)+'[value]');return d.set.call(this,v)},get:function(){return d.get.call(this)},configurable:true})}catch(e){}"
            js = js & "try{new MutationObserver((ms)=>{ms.forEach((m)=>{if(iN(m.target,selectors))return;var path=gX(m.target),entry=(m.type==='attributes')?'ATTR:'+path+'['+m.attributeName+']':'DOM:'+path;"
            
            ' Attribute Filtering: Ignore mutations on technical/transient attributes
            js = js & "if(entry.indexOf('sizzle')===-1) lg(entry);})}).observe(document,{childList:true,subtree:true,attributes:true})}catch(e){}"
            
            js = js & "if(typeof window.fetch==='function'&&!window.fetch.__vbaWrapped){var oF=window.fetch;window.fetch=function(u,o){var r=(typeof u==='string')?u:(u&&u.url)||'';if(ignoreNetwork.some((pat)=>r.indexOf(pat)!==-1))return oF.apply(this,arguments);"
            js = js & "p.inflightFetchCount++;var dec=()=>{p.inflightFetchCount--;uT()};return oF.apply(this,arguments).then((r)=>{dec();return r}).catch((e)=>{dec();throw e})};window.fetch.__vbaWrapped=true}"
            js = js & "if(!XMLHttpRequest.prototype.__vbaWrapped){var oO=XMLHttpRequest.prototype.open,oS=XMLHttpRequest.prototype.send;XMLHttpRequest.prototype.open=function(m,u){this.__vbaUrl=u;this.__vbaHooked=true;return oO.apply(this,arguments)};"
            js = js & "XMLHttpRequest.prototype.send=function(){var s=this,u=this.__vbaUrl||'';if(ignoreNetwork.some((pat)=>u.indexOf(pat)!==-1)){s.__vbaHooked=false;return oS.apply(this,arguments)}p.inflightXhrCount++;"
            js = js & "var dn=()=>{if(s.__vbaHooked){p.inflightXhrCount--;s.__vbaHooked=false;uT()}s.removeEventListener('loadend',dn)};this.addEventListener('loadend',dn);return oS.apply(this,arguments)};XMLHttpRequest.prototype.__vbaWrapped=true}return true})();"
        Case "query"
            js = js & "var p=window.__vbaIdleProbe;if(!p)return JSON.stringify({active:false});"
            js = js & "var res={active:true,now:performance.now(),lastMutationTs:p.lastMutationTs,inflightFetchCount:p.inflightFetchCount,inflightXhrCount:p.inflightXhrCount,changes:p.changeList};"
            js = js & "p.changeList=[];return JSON.stringify(res);"
    End Select
    GetProbeJs = js
End Function

' ========================================================================================
' Internal Helper: Builds a target dictionary for BiDi script commands.
' Priority: Uses Realm ID if provided for execution stability, falls back to Context ID.
' ========================================================================================
Private Function BuildContextTarget(ByVal contextId As String, _
                                   Optional ByVal realmId As String = "") As Dictionary
    Dim target As New Dictionary
    
    If realmId <> "" Then
        ' Execution in a specific realm is more stable for persistent probes
        target.Add "realm", realmId
    Else
        ' Fallback to context-level targeting
        target.Add "context", contextId
    End If
    
    Set BuildContextTarget = target
End Function

' ========================================================================================
' GetRawTicks: Returns the current performance counter value as Currency.
' ========================================================================================
Private Function GetRawTicks() As Currency
    Dim val As Currency
    QueryPerformanceCounter val
    GetRawTicks = val
End Function

' ========================================================================================
' FUNCTION: GetElapsedMs
' SUMMARY: Calculates elapsed time in milliseconds.
' ========================================================================================
Private Function GetElapsedMs(ByVal startTicks As Currency) As Double
    Dim currentTicks As Currency
    QueryPerformanceCounter currentTicks
    ' Calculate ms based on hardware frequency without fallback logic
    GetElapsedMs = (CDbl(currentTicks - startTicks) / CDbl(p_timerFreq)) * 1000
End Function

' ========================================================================================
' Configuration: Public properties to manage noise filtering from Standard Modules.
' ========================================================================================
Public Property Let IgnoreSelectors(ByVal cssSelectors As String)
    If p_userIgnoreSelectors Is Nothing Then Set p_userIgnoreSelectors = New Collection
    p_userIgnoreSelectors.Add cssSelectors
End Property

Public Property Let IgnoreNetworkPatterns(ByVal urlPattern As String)
    If p_userIgnoreNetworkPatterns Is Nothing Then Set p_userIgnoreNetworkPatterns = New Collection
    p_userIgnoreNetworkPatterns.Add urlPattern
End Property

' ========================================================================================
' Internal Helper: Ensures a Dictionary object is initialized.
' Returns a new Dictionary if the input is Nothing, otherwise returns the input.
' ========================================================================================
Private Function EnsureDictionary(ByVal dic As Dictionary) As Dictionary
    If dic Is Nothing Then
        Set EnsureDictionary = New Dictionary
    Else
        Set EnsureDictionary = dic
    End If
End Function

' ========================================================================================
' [INTERNAL] MasterWaitEngine (Enhanced with UI Protection & Telemetry)
' SUMMARY: High-precision polling engine for all wait operations.
' FEATURES:
'   - Physical UI locking (Application.Interactive) to prevent cell-edit freezes.
'   - Real-time telemetry in the Status Bar (Throttled at 100ms).
'   - Stack-depth aware state management for nested wait stability.
' ========================================================================================
Private Function MasterWaitEngine(ByVal mode As waitMode, _
                                  ByVal timeOutms As Long, _
                                  ByVal intervalMs As Long, _
                                  ByRef ctx As Dictionary) As Boolean
    
    ' 1. Parameter Sanitization
    ' Ensures timing values respect the hardware floor for stable Windows messaging.
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeOutms, "timeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    Dim vInterval As Long: vInterval = SanitizeMs(intervalMs, "intervalMs", 150, , 100)
    
    ' Execution state tracking
    Dim errNum As Long, errDesc As String, errSrc As String
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim success As Boolean: success = False

    ' 2. Re-entrancy & UI State Management
    m_waitStackDepth = m_waitStackDepth + 1
    
    If m_waitStackDepth = 1 Then
        ' --- ROOT ENTRY ---
        ' Save initial UI state for cleanup, but priority is always "Restore to True".
        m_originalInteractive = Application.Interactive
        m_originalCursor = Application.Cursor
        
        ' PHYSICAL LOCK: Prevent cell-edit mode from suspending the VBA thread.
        Application.Interactive = False
        Application.Cursor = xlWait
        
        ' JIT HUD: Show real-time status in the browser window.
        On Error Resume Next
        Call Me.CallScript(GetHudJs("init"), Nothing, ResolveTargetContext(CStr(ctx("contextId"))), False, False, 2000, True)
        On Error GoTo 0
    End If

    ' 3. Main Polling Loop
    On Error GoTo ErrorHandler
    Do
        ' Yield to OS: Process WebSocket queue and maintain Excel responsiveness.
        WaitHeartbeat vInterval
        
        ' THROTTLED HUD UPDATE: Minimize traffic while keeping display fresh.
        If GetElapsedMs(m_lastDoEventsTime) > 100 Then
            Call UpdateWaitTelemetry(mode, ctx)
        End If
        
        ' Check condition (URL, SPA stability, etc.)
        success = EvaluateWaitCondition(mode, ctx)
        
        ' Exit if goal reached or timed out.
        If success Or IsTimeout(startTick, CDbl(vTimeout)) Then Exit Do
    Loop

' ========================================================================================
' CLEANUP & RESTORATION (The "Finally" Block)
' ========================================================================================
CleanExit:
    On Error Resume Next
    
    ' 4. Safe Stack Restoration
    ' If error occurred (errNum <> 0), force full UI reset to prevent permanent lock.
    If errNum <> 0 Or m_waitStackDepth <= 1 Then
        ' --- EMERGENCY / ROOT EXIT ---
        Call Me.CallScript(GetHudJs("remove"), Nothing, ResolveTargetContext(CStr(ctx("contextId"))), False, False, 1000, True)
        
        ' FORCE RESTORE: Ensure Excel is usable after execution.
        Application.Interactive = True
        Application.Cursor = xlDefault
        Application.StatusBar = False
        m_waitStackDepth = 0 ' Reset stack counter
    Else
        ' Nested exit: decrement stack depth normally.
        m_waitStackDepth = m_waitStackDepth - 1
    End If
    
    ' 5. Error Propagation
    If errNum <> 0 Then
        On Error GoTo 0
        Err.Raise errNum, "BiDi:MasterWaitEngine:" & errSrc, errDesc
    End If
    
    MasterWaitEngine = success
    Exit Function

' ========================================================================================
' ERROR HANDLER
' ========================================================================================
ErrorHandler:
    errNum = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    Resume CleanExit
End Function

' ========================================================================================
' FUNCTION: EvaluateWaitCondition (Version 4.2 - Fixed ByRef and IIf issues)
' ========================================================================================
Private Function EvaluateWaitCondition(ByVal mode As waitMode, ByRef ctx As Dictionary) As Boolean
    Const MAX_RECOVERY_ATTEMPTS As Long = 8
    If ctx Is Nothing Then Err.Raise 520, "BiDi", "Wait context is Nothing."
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(CStr(ctx("contextId")))
    
    Select Case mode
        Case WaitMode_UrlContains
            ' Pass Silent:=True to GetCurrentUrl during polling
            EvaluateWaitCondition = (InStr(GetCurrentUrl(targetCtx, True), CStr(ctx("partialUrl"))) > 0)
        
        Case WaitMode_SpaIdle
            Dim probe As Dictionary: Set probe = QuerySpaProbe(targetCtx)
            
            Dim isProbeActive As Boolean: isProbeActive = False
            If Not probe Is Nothing Then
                If probe.Exists("active") Then isProbeActive = CBool(probe("active"))
            End If

            If isProbeActive Then
                ' --- PATH A: Probe Active ---
                ctx("missingProbeCount") = 0
                
                ' Check for Stalled Renderer via heartbeat (performance.now() check)
                Dim jsNowVal As Double: jsNowVal = CDbl(probe("now"))
                If jsNowVal = m_lastJsNow And (p_pendingRequestCount > 0) Then
                    RecoverSpaProbe targetCtx
                End If
                m_lastJsNow = jsNowVal
                
                ' [FIX] Core Consensus Engine (Network + DOM)
                ' State management is fully delegated to EvaluateSpaIdleCondition via 'ctx'
                EvaluateWaitCondition = EvaluateSpaIdleCondition(probe, ctx, nowMs())
            Else
                ' --- PATH B: Probe Recovery ---
                Dim bCount As Long: bCount = CLng(ctx("missingProbeCount")) + 1
                ctx("missingProbeCount") = bCount
                ctx("stableStart") = nowMs()
                
                If bCount > MAX_RECOVERY_ATTEMPTS Then
                    Err.Raise 516, "BiDi", "SPA Probe failed after " & MAX_RECOVERY_ATTEMPTS & " attempts."
                End If
                
                ' Log recovery attempt for visibility in debug window
                If DEBUG_MODE Then Debug.Print "BiDi: SPA Probe inactive. Re-injecting... (Attempt " & bCount & ")"
                StartSpaIdleProbe targetCtx
                EvaluateWaitCondition = False
            End If
    End Select
End Function

' ========================================================================================
' [INTERNAL] ExecuteBaseAction (STOP-ON-FAIL VERSION)
' ========================================================================================
Private Function ExecuteBaseAction(ByVal contextId As String, _
                                    ByVal jsKey As String, _
                                    ByVal args As Collection, _
                                    Optional ByVal waitNetwork As Boolean = True, _
                                    Optional ByVal minStable As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal awaitPromise As Boolean = True, _
                                    Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                    Optional ByVal maxTimeout As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal checkUrl As Boolean = False) As String
    Dim retryCount As Long
    Dim strRes As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim lastErrNum As Long, lastErrDesc As String, lastErrSrc As String
    Dim actualAttempts As Long

    ' 1. Sanitize timing parameters (Enforce safe floor values)
    Dim vStable As Long:  vStable = SanitizeMs(minStable, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vTimeout As Long: vTimeout = SanitizeMs(maxTimeout, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)

    ' 2. Primary Retry Loop
    For retryCount = 0 To maxRetries
        actualAttempts = retryCount + 1
        Err.Clear
        On Error Resume Next
        
        ' --- PHASE A: PREPARATION ---
        ' Ensure SPA probe is active. Jump to retry if internal injection fails.
        If waitNetwork Then
            PrepareProbeForAction targetCtx
            If Err.Number <> 0 Then GoTo ContinueRetry
        End If
        
        ' On retries, refresh context tree to clear any stale element/context references.
        If retryCount > 0 Then
            UpdateMainContextId True, True ' Silent refresh
            If Err.Number <> 0 Then GoTo ContinueRetry
            WaitHeartbeat 300
        End If
        
        ' --- PHASE B: EXECUTION ---
        ' Dispatch JavaScript action. CallScript triggers the internal watchdog.
        strRes = CallScript(GetActionJs(jsKey), args, targetCtx, awaitPromise, , vTimeout)
        
        ' --- PHASE C: EVALUATION ---
        ' Verify overall success. Err.Number is only 0 if ALL preceding steps succeeded.
        If Err.Number = 0 Then
            ' Final environment stabilization (Wait for Network/DOM Idle)
            If waitNetwork Then SyncAfterAction targetCtx, vStable, vTimeout, checkUrl
            
            ExecuteBaseAction = strRes
            On Error GoTo 0 ' Deactivate OERN before exit
            Exit Function
        End If

ContinueRetry:
        ' --- PHASE D: RECOVERY & LOGGING ---
        ' Capture error state immediately before it's cleared by subsequent cleanup logic.
        lastErrNum = Err.Number: lastErrDesc = Err.Description: lastErrSrc = Err.Source
        
        ' Reset error handler to allow safe execution of retry logic
        On Error GoTo 0
        
        ' Decision A: Abort if it's a non-retryable functional error (e.g., Element not found)
        If lastErrNum = ERR_ACTION_FAILED Then Exit For
        
        ' Decision B: Retry if it's a protocol desync and attempts remain
        If lastErrNum = ERR_BIDI_PROTOCOL And retryCount < maxRetries Then
            If DEBUG_MODE Then
                Debug.Print "BiDi: Action '" & jsKey & "' failed (Attempt " & _
                            actualAttempts & "/" & (maxRetries + 1) & "). Retrying after 500ms..."
            End If
            WaitHeartbeat 500
        Else
            ' Final attempt failed or fatal error encountered
            Exit For
        End If
    Next retryCount

    ' 3. Final Error Propagation (The Safety Guardrail)
    ' Ensure we never return a false positive or an empty error state.
    If lastErrNum = 0 Then lastErrNum = ERR_ACTION_FAILED
    If lastErrDesc = "" Then lastErrDesc = "Unexpected execution state in ExecuteBaseAction."
    
    Err.Raise lastErrNum, lastErrSrc, "Action '" & jsKey & "' failed after " & _
              actualAttempts & " attempt(s). [Context: " & targetCtx & "] Detail: " & lastErrDesc
End Function

' ========================================================================================
' [INTERNAL] SyncAfterAction (Updated with URL Stability)
' ========================================================================================
Private Sub SyncAfterAction(ByVal contextId As String, _
                            ByVal minStableMs As Long, _
                            ByVal maxTimeoutMs As Long, _
                            Optional ByVal checkUrl As Boolean = False)
    
    ' 1. Brief pause to let the initial event fire (Non-blocking)
    WaitHeartbeat 50
    
    ' 2. Conditional URL/ReadyState monitoring
    '    Only executed when the developer explicitly expects a page transition.
    If checkUrl Then
        If DEBUG_MODE Then Debug.Print ">>> [Sync] Page navigation mode active."
        On Error Resume Next
        ' Wait for the physical address and document state to settle
        WaitForUrlStability contextId, 200, 3000
        ' Refresh the context map for the new document
        UpdateMainContextId True
        On Error GoTo 0
    End If
    
    ' 3. Flush the BiDi event buffer
    ProcessRecorderQueue
    
    ' 4. Core SPA Synchronization (The Engine's Heart)
    '    Wait for Network requests and DOM Mutations to cease.
    WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, contextId
End Sub

' ========================================================================================
' [PRIVATE] ProcessResponse (Upgraded with Auto-Unwrap)
' SUMMARY: Validates BiDi responses and automatically unwraps the Safety Timeout Shell.
' ========================================================================================
Private Function ProcessResponse(ByVal jsonRaw As String, ByVal methodName As String) As Object
    Dim json As Object
    Dim resultObj As New Dictionary ' The fresh, clean output container
    
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonRaw)
    On Error GoTo 0
    
    If json Is Nothing Then
        Err.Raise vbObjectError + 512, "BiDi:" & methodName, "Response is not valid JSON."
    End If

    ' --- Layer 1: Protocol Level Errors ---
    ' Checks for BiDi-level errors (e.g., "no such window", "invalid session id")
    If json.Exists("error") Then
        Dim errType As String: errType = json("error")
        Dim errMsg As String: errMsg = json("message")
        If Not IsIgnorableError(errType, errMsg) Then
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:" & methodName, "[" & errType & "] " & errMsg
        End If
    End If

    ' --- Layer 2: JavaScript Runtime Exceptions ---
    ' Handles exceptions thrown during JS evaluation in the browser.
    Dim excDetails As Object
    Set excDetails = GetPathObject(json, "result.exceptionDetails")
    If Not excDetails Is Nothing Then
        If DEBUG_MODE Then
            Debug.Print "--- BIDI JS ERROR REPORT | Method: " & methodName & " ---"
            Debug.Print FormatJavaScriptError(excDetails)
        End If
        
        Err.Raise ERR_JS_EXCEPTION, "BiDi:JSException", _
                  "JS Error in " & methodName & ": " & GetPathString(json, "result.exceptionDetails.text")
    End If

    ' --- Layer 3: Safety Shell Unwrapping (Non-Destructive) ---
    ' Extracts the value without modifying the original 'json' object.
    ' This prevents data contamination during retries.
    Dim rawVal As Variant
    rawVal = InternalTraverseFromObject(json, "result.result.value")
    
    Dim isWrapped As Boolean: isWrapped = False
    
    If VarType(rawVal) = vbString Then
        Dim sVal As String: sVal = Trim(CStr(rawVal))
        ' Check for the custom JSON shell: { "status": "...", "value": ... }
        If Left(sVal, 1) = "{" And Right(sVal, 1) = "}" Then
            Dim innerJson As Object
            On Error Resume Next
            Set innerJson = WebJsonConverter.ParseJson(sVal)
            On Error GoTo 0
            
            If Not innerJson Is Nothing Then
                If innerJson.Exists("status") Then
                    isWrapped = True
                    ' If the browser-side watchdog triggered a timeout, raise it as an action failure.
                    If innerJson("status") = "error" Then
                        Err.Raise ERR_ACTION_FAILED, "BiDi:SafetyTimeout:" & methodName, CStr(innerJson("message"))
                    End If
                    
                    ' Transfer the unwrapped value to our new result container.
                    Call AssignValue(resultObj, "value", innerJson("value"))
                End If
            End If
        End If
    End If

    ' Fallback: If no shell was detected, use the original raw value.
    If Not isWrapped Then
        Call AssignValue(resultObj, "value", rawVal)
    End If

    resultObj("status") = "ok"
    Set ProcessResponse = resultObj
End Function

' ========================================================================================
' [INTERNAL] AssignValue
' SUMMARY: Safely assigns a value to a Dictionary key using Let or Set as appropriate.
' ========================================================================================
Private Sub AssignValue(ByRef dic As Dictionary, ByVal key As String, ByRef val As Variant)
    If IsObject(val) Then
        Set dic(key) = val
    Else
        dic(key) = val
    End If
End Sub

' ========================================================================================
' JSON Path Utilities (Object-Safe Version)
' ========================================================================================

Private Function GetPathObject(ByVal root As Object, ByVal Path As String) As Object
    On Error Resume Next
    ' Always use Set for objects to prevent Error 450
    Set GetPathObject = InternalTraverseFromObject(root, Path)
    On Error GoTo 0
End Function

Private Function GetPathString(ByVal root As Object, ByVal Path As String) As String
    Dim v As Variant
    ' Check if it is an object first to avoid 450
    If IsObject(InternalTraverseFromObject(root, Path)) Then
        GetPathString = "" ' Cannot convert object to string directly
    Else
        v = InternalTraverseFromObject(root, Path)
        GetPathString = CStr(v & "")
    End If
End Function

Private Function GetPathValue(ByVal root As Object, ByVal Path As String) As Variant
    ' Use a temporary variant to safely determine assignment type
    Dim raw As Variant
    
    ' Intermediate check using IsObject
    If IsObject(InternalTraverseFromObject(root, Path)) Then
        Set GetPathValue = InternalTraverseFromObject(root, Path)
    Else
        GetPathValue = InternalTraverseFromObject(root, Path)
    End If
End Function

' ========================================================================================
' FUNCTION: InternalJsonTraverse
' SUMMARY: Entry point to parse a JSON string and traverse to a specific path.
'          Handles the conversion from String to Object before path resolution.
' ========================================================================================
Private Function InternalJsonTraverse(ByVal jsonStr As String, ByVal Path As String) As Variant
    Dim json As Object
    
    ' 1. Defensive check for empty input
    If Len(jsonStr) = 0 Then Exit Function

    On Error Resume Next
    ' 2. Parse the raw JSON string into a VBA-accessible Object (Dictionary/Collection)
    Set json = WebJsonConverter.ParseJson(jsonStr)
    On Error GoTo 0
    
    ' 3. Proceed only if parsing was successful
    If Not json Is Nothing Then
        ' 4. Call the traversal engine to find the value at the specified dot-notation path.
        ' To avoid Error 450, we must determine if the result is an Object or a Primitive.
        If IsObject(InternalTraverseFromObject(json, Path)) Then
            ' Use SET for Dictionary or Collection objects
            Set InternalJsonTraverse = InternalTraverseFromObject(json, Path)
        Else
            ' Use standard assignment for String, Double, Boolean, etc.
            InternalJsonTraverse = InternalTraverseFromObject(json, Path)
        End If
    End If
End Function

' ========================================================================================
' FUNCTION: InternalTraverseFromObject
' SUMMARY: Recursively traverses a VBA JSON object structure using dot-notation.
'          Supports both Dictionary keys and Collection indices (0-based for JSON).
' ========================================================================================
Private Function InternalTraverseFromObject(ByVal root As Object, ByVal Path As String) As Variant
    On Error Resume Next
    If root Is Nothing Or Path = "" Then Exit Function
    
    Dim keys() As String: keys = Split(Path, ".")
    Dim current As Object: Set current = root
    Dim i As Long, idx As Long
    
    ' 1. Traverse the hierarchy until the second-to-last segment.
    '    This loop resolves all parent nodes in the path.
    For i = 0 To UBound(keys) - 1
        If TypeName(current) = "Dictionary" Then
            If Not current.Exists(keys(i)) Then Exit Function
            ' [CRITICAL FIX] Verify node is an Object before using 'Set'
            If IsObject(current(keys(i))) Then
                Set current = current(keys(i))
            Else
                Exit Function ' Expected an object for further traversal, but hit a primitive
            End If
            
        ElseIf TypeName(current) = "Collection" And IsNumeric(keys(i)) Then
            ' Convert 0-based JSON index to 1-based VBA Collection index.
            idx = CLng(keys(i)) + 1
            If idx < 1 Or idx > current.Count Then Exit Function
            
            ' [CRITICAL FIX] Verify node is an Object before using 'Set'
            ' Prevents crash if the collection contains primitive values.
            If IsObject(current(idx)) Then
                Set current = current(idx)
            Else
                Exit Function
            End If
        Else
            ' Invalid path or non-traversable object (primitive) encountered.
            Exit Function
        End If
        If current Is Nothing Then Exit Function
    Next i
    
    ' 2. Resolve the final leaf node.
    Dim lastKey As String: lastKey = keys(UBound(keys))
    
    If TypeName(current) = "Dictionary" Then
        If current.Exists(lastKey) Then
            ' Handle the leaf value correctly (Set for objects, Let for primitives)
            If IsObject(current(lastKey)) Then
                Set InternalTraverseFromObject = current(lastKey)
            Else
                InternalTraverseFromObject = current(lastKey)
            End If
        End If
    ElseIf TypeName(current) = "Collection" And IsNumeric(lastKey) Then
        idx = CLng(lastKey) + 1
        If idx >= 1 And idx <= current.Count Then
            If IsObject(current(idx)) Then
                Set InternalTraverseFromObject = current(idx)
            Else
                InternalTraverseFromObject = current(idx)
            End If
        End If
    End If
    On Error GoTo 0
End Function

' ========================================================================================
' [INTERNAL] ResolveXPathOrThrow
' SUMMARY: A bottleneck function to resolve an XPath to a sharedId.
'          Features:
'          - Integrated validation for 'searchTimeoutMs' to prevent loop logic errors.
'          - Centralized error raising if the element is not found within the timeout.
'          - Ensures that the target context is fresh before the search.
' ========================================================================================
Private Function ResolveXPathOrThrow(ByVal xpath As String, _
                                     ByVal contextId As String, _
                                     ByVal timeOutms As Long, _
                                     ByVal methodName As String) As String
    Dim sid As String
    Dim targetCtx As String
    
    ' 1. Consolidated Validation: Prevent logic errors from 'True' or negative inputs.
    ' Even if a public method receives 'True', it is sanitized here into a safe Long.
    Dim vTimeout As Long
    vTimeout = SanitizeMs(timeOutms, "searchTimeoutMs", 10000)
    
    ' 2. Resolve the active browsing context
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId()
    End If
    
    ' 3. Execution: Attempt to locate the element using the sanitized timeout.
    ' This calls the internal private locator which handles the polling loop.
    sid = GetFindElementSharedIdByXPath(xpath, vTimeout, targetCtx)
    
    ' 4. Error Handling: Raise a standardized exception if resolution fails.
    If sid = "" Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
                  "Element not found within " & vTimeout & "ms for XPath: " & xpath
    End If
    
    ResolveXPathOrThrow = sid
End Function

' ========================================================================================
' [HELPER] PrepareProbeForAction
' Summary: Ensures the SPA monitoring probe is installed and active.
'          Called JIT (Just-In-Time) before any interaction to prevent count leaks.
' ========================================================================================
Private Sub PrepareProbeForAction(ByVal contextId As String)
    Dim probe As Dictionary
    Dim isReady As Boolean: isReady = False
    
    ' Internal probe status check is Silent
    Set probe = QuerySpaProbe(contextId)
    
    If Not probe Is Nothing Then
        If probe.Exists("active") Then isReady = CBool(probe("active"))
    End If

    If Not isReady Then
        ' Probe injection is an internal management task, so it should be Silent
        StartSpaIdleProbe contextId
        WaitHeartbeat 150
    End If
End Sub

' ========================================================================================
' [INTERNAL] SanitizeMs
' SUMMARY: Central gatekeeper to validate millisecond arguments.
'          Strictly rejects negative values (preventing True/-1 errors)
'          and caps values to prevent permanent Excel hangs.
' ========================================================================================
Private Function SanitizeMs(ByVal ms As Long, ByVal argName As String, _
                             ByVal defaultMs As Long, _
                             Optional ByVal applyBuffer As Boolean = False, _
                             Optional ByVal minMs As Long = 50) As Long
    Dim vMs As Long: vMs = ms
    
    ' 1. Logic Guard: Rejects zero or negative values.
    ' Rationale: 0ms in polling loops (like WaitForUrlStability) triggers
    ' accidental immediate timeouts or CPU-intensive "empty loops."
    ' Also prevents logic errors where a Boolean True (-1) is passed as Long.
    If vMs <= 0 Then
        If DEBUG_MODE Then
            Debug.Print "BiDi Warning: Invalid value (" & ms & ") for '" & argName & "'. Using default: " & defaultMs
        End If
        vMs = defaultMs
    End If

    ' 2. Minimum Floor Guard:
    ' Ensure the interval is not too short for Windows message pumping.
    ' Values under 50ms can cause "Not Responding" states in heavy SPAs.
    If vMs < minMs Then vMs = minMs

    ' 3. Range Guard: Circuit breaker for UI responsiveness (10 min cap)
    If vMs > 600000 Then vMs = 600000
    
    ' 4. Strategic Buffer:
    ' Applied for low-resolution timers if requested.
    If applyBuffer And vMs > 0 Then
        SanitizeMs = vMs + m_timerPrecisionBuffer
    Else
        SanitizeMs = vMs
    End If
End Function

' ========================================================================================
' [PUBLIC] WaitForUrlStability
' DESCRIPTION: Polls the current URL and Document.readyState until they remain constant
'              for a specified 'stableDurationMs'.
' STRATEGY   : Uses a "Fail-Fast" approach. If the environment fails to stabilize
'              within the timeout, the loop exits.
' ========================================================================================
Public Sub WaitForUrlStability(Optional ByVal contextId As String = "", _
                               Optional ByVal stableDurationMs As Long = 500, _
                               Optional ByVal timeOutms As Long = 10000)

    ' --- [ENTRY GUARD] Guard against 0ms inputs ---
    ' Rationale: Stability confirmation requires a positive duration to compare against.
    ' If 0 is passed (due to updated SanitizeMs logic), we revert to safe defaults.
    If stableDurationMs <= 0 Then stableDurationMs = 500
    If timeOutms <= 0 Then timeOutms = 10000
    ' ----------------------------------------------

    Dim vStableMs As Long: vStableMs = SanitizeMs(stableDurationMs, "stableDurationMs", 500)
    Dim vTimeoutMs As Long: vTimeoutMs = SanitizeMs(timeOutms, "timeoutMs", 10000)

    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim startMs As Double: startMs = nowMs()
    Dim stableStartMs As Double: stableStartMs = nowMs()
    Dim lastStateKey As String: lastStateKey = ""
    Dim checkJs As String: checkJs = "return window.location.href + '|' + document.readyState;"
    
    Do
        ' FORCE Silent:=True to prevent technically repetitive REQ/RES logs during polling
        Dim currentResult As String
        currentResult = CStr(GetCleanResult(CallScript(checkJs, Nothing, targetCtx, False, False, vTimeoutMs, True)))
        
        ' Determine if the page state is still in flux
        Dim isMoving As Boolean
        ' State is "moving" if the URL/ReadyState changed, or if not yet "complete"
        isMoving = (currentResult <> lastStateKey) Or (Right$(currentResult, 8) <> "complete")
        
        If isMoving Then
            ' Reset the stability timer
            lastStateKey = currentResult
            stableStartMs = nowMs()
        Else
            ' Check if the state has been stable for the required duration
            If (nowMs() - stableStartMs) >= CDbl(vStableMs) Then Exit Do
        End If
        
        ' Global timeout check to prevent infinite loops
        If (nowMs() - startMs) > CDbl(vTimeoutMs) Then Exit Do
        
        ' Yield to OS and process WebSocket queue
        WaitHeartbeat 50
    Loop
End Sub

' ========================================================================================
' [PRIVATE] nowMs
' DESCRIPTION: Returns monotonic time in milliseconds based on QueryPerformanceCounter.
' ========================================================================================
Private Function nowMs() As Double
    Dim t As Currency
    QueryPerformanceCounter t
    ' p_timerFreq is initialized in Class_Initialize
    nowMs = (CDbl(t) / CDbl(p_timerFreq)) * 1000#
End Function

' ========================================================================================
' [PRIVATE] WrapInSafetyShell
' DESCRIPTION: Wraps JS code in a Promise.race to enforce a browser-side timeout.
'              Ensures that the browser always sends a response even if the task hangs.
' ========================================================================================
Private Function WrapInSafetyShell(ByVal jsBody As String, ByVal timeOutms As Long) As String
    Dim js As String
    
    js = "async function(){ "
    js = js & "const _t=" & timeOutms & "; let _la=Date.now(); "
    js = js & "const hb=()=>{_la=Date.now()}; "
    
    ' -----------------------------------------------------------------------------------
    ' IMPORTANT: Arrow function is intentional.
    ' _tk MUST remain an arrow function for two reasons:
    '   1. arguments[]: Arrow functions inherit the outer async function()'s arguments,
    '      which contains the BiDi sharedId references passed via script.callFunction.
    '   2. hb(): Arrow functions capture hb via lexical closure from the outer scope.
    ' Changing to "async function(){}" would break hb() access.
    ' Changing to "async function(){}" with .apply() would break hb() access differently.
    ' This design is load-bearing. Do not refactor without testing both click and shadow_click.
    ' -----------------------------------------------------------------------------------
    js = js & "const _tk = async () => { " & jsBody & " }; "
    
    js = js & "let _wdTimer; "
    js = js & "const _wd = new Promise((_,r)=>{ "
    js = js & "  _wdTimer=setInterval(()=>{"
    js = js & "    if(Date.now()-_la>_t){clearInterval(_wdTimer);r(new Error('VBA_TIMEOUT'))}"
    js = js & "  },100); "
    js = js & "}); "
    
    js = js & "try{ "
    js = js & "  const res = await Promise.race([_tk(), _wd]); "
    js = js & "  clearInterval(_wdTimer); "
    js = js & "  return JSON.stringify({status:'ok',value:res}); "
    js = js & "}catch(e){ "
    js = js & "  clearInterval(_wdTimer); "
    js = js & "  return JSON.stringify({status:'error',message:e.message}); "
    js = js & "} "
    js = js & "}"
    
    WrapInSafetyShell = js
End Function

' ========================================================================================
' [PRIVATE] HandleConsoleLog
' DESCRIPTION: Formats and prints browser console logs to the VBA Immediate Window.
' ========================================================================================
Private Sub HandleConsoleLog(ByVal json As Object)
    On Error Resume Next
    Dim text As String: text = GetPathString(json, "params.text")
    
    ' Only output logs from our specific AutoClicker to keep the window clean.
    If InStr(text, "BiDi-AutoClicker") > 0 Then
        Debug.Print ">>> [BROWSER] " & text
    End If
    On Error GoTo 0
End Sub

' ========================================================================================
' [PRIVATE] GetCleanResult
' SUMMARY: Safely extracts the final result value from a processed BiDi JSON string.
' ========================================================================================
Private Function GetCleanResult(ByVal jsonRaw As String) As Variant
    On Error Resume Next
    
    ' jsonRaw is the output of CallScript -> ProcessResponse -> ConvertToJson.
    ' It no longer contains the original BiDi wrapper layers.
    Dim json As Object: Set json = WebJsonConverter.ParseJson(jsonRaw)
    
    ' [FIX] Target the root "value" key directly.
    ' InternalTraverseFromObject handles Dictionary/Collection resolution.
    Dim val As Variant: val = InternalTraverseFromObject(json, "value")
    
    ' Object-Safe Assignment (Let vs Set)
    If IsObject(val) Then
        If val Is Nothing Then
            GetCleanResult = Empty
        Else
            Set GetCleanResult = val
        End If
    Else
        GetCleanResult = val
    End If
    
    On Error GoTo 0
End Function

' ========================================================================================
' METHOD: Shutdown
' SUMMARY: Explicitly releases all BiDi/CDP resources.
'          This should be called before setting the object to Nothing.
' ========================================================================================
Public Sub Shutdown()
    ' --- Gatekeeper: Prevent double execution ---
    ' If the socket is already released, we have nothing left to clean up.
    If socket_ Is Nothing Then Exit Sub

    On Error Resume Next
    
    ' --- 1. CDP Session Detachment ---
    ' Explicitly detach the CDP session while the browser/driver is still alive.
    If p_cdpSessionId <> "" Then
        Dim params As New Dictionary
        params.Add "session", p_cdpSessionId
        
        ' Send the detach command silently to prevent logging during cleanup.
        ' This is a "best-effort" action.
        ExecuteBiDiCommand "goog:cdp.detachSession", params, True
        p_cdpSessionId = ""
    End If

    ' --- 2. Recorder Memory Cleanup ---
    ' Release collection items to prevent potential memory leaks with nested objects.
    If Not p_recorderLogs Is Nothing Then
        Dim i As Long
        For i = p_recorderLogs.Count To 1 Step -1
            p_recorderLogs.Remove i
        Next i
        Set p_recorderLogs = Nothing
    End If

    ' --- 3. Internal State Reset ---
    ClearState

    ' --- 4. Communication Socket Release ---
    ' Setting the socket to Nothing here triggers the gatekeeper in subsequent calls.
    Set socket_ = Nothing
    
    If DEBUG_MODE Then Debug.Print "BiDi: Explicit Shutdown sequence completed."
End Sub

' ========================================================================================
' [PRIVATE] UpdateWaitTelemetry
' SUMMARY: Displays real-time browser synchronization status in the Excel Status Bar.
' ========================================================================================
Private Sub UpdateWaitTelemetry(ByVal mode As waitMode, ByRef ctx As Dictionary)
    On Error Resume Next
    Dim msg As String, colorCode As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(CStr(ctx("contextId")))
    
    Select Case mode
        Case WaitMode_UrlContains
            msg = "BiDi: https://www.sync.com/" & vbCrLf & "Waiting for: " & ctx("partialUrl")
            colorCode = "#00FF00" ' Green (Stable check)
            
        Case WaitMode_SpaIdle
            Dim netCount As Long: netCount = p_pendingRequestCount
            Dim idleTime As Double: idleTime = (nowMs() - CDbl(ctx("stableStart"))) / 1000
            Dim targetIdle As Double: targetIdle = CLng(ctx("minStableMs")) / 1000
            
            ' VISUAL FEEDBACK: Red border if the environment is still unstable
            If netCount > 0 Or idleTime < targetIdle Then
                colorCode = "#FF4444"
            Else: colorCode = "#00FF00"
            End If
            
            msg = "BiDi: [SPA SYNC]" & vbCrLf & _
                  "Network: " & netCount & " active" & vbCrLf & _
                  "Idle: " & Format(idleTime, "0.0") & "s / " & Format(targetIdle, "0.0") & "s"
    End Select
    
    ' Inject update via CallScript with Silent:=True to prevent log flooding
    Dim updateScript As String
    updateScript = GetHudJs("update")
    updateScript = Replace(updateScript, "[[MSG]]", Replace(msg, vbCrLf, "\n"))
    updateScript = Replace(updateScript, "[[COLOR]]", colorCode)
    
    Call Me.CallScript(updateScript, Nothing, targetCtx, False, False, 1000, True)
End Sub

' ========================================================================================
' [PRIVATE] GetHudJs
' SUMMARY: Returns JavaScript snippets for the Browser-side HUD (Head-Up Display).
' ========================================================================================
Private Function GetHudJs(ByVal snippetType As String) As String
    Select Case LCase(snippetType)
        Case "init"
            ' Injects a floating HUD. 'pointer-events:none' ensures it won't block clicks.
            GetHudJs = "(function(){" & _
                       "if(document.getElementById('vba-bidi-hud')) return;" & _
                       "const h=document.createElement('div'); h.id='vba-bidi-hud';" & _
                       "h.style.cssText='position:fixed;top:15px;right:15px;z-index:2147483647;padding:10px;font-family:monospace;font-size:12px;background:rgba(0,0,0,0.85);color:#00FF00;border:1px solid #00FF00;border-radius:6px;pointer-events:none;box-shadow:0 4px 15px rgba(0,0,0,0.5);min-width:180px;white-space:pre-wrap;';" & _
                       "h.innerText='BiDi: Initializing...'; document.body.appendChild(h);" & _
                       "})();"
                       
        Case "update"
            ' Updates content and switches border color (Red = Busy, Green = Stable)
            GetHudJs = "const h=document.getElementById('vba-bidi-hud'); " & _
                       "if(h){ h.innerText='[[MSG]]'; h.style.borderColor='[[COLOR]]'; h.style.color='[[COLOR]]'; }"
                       
        Case "remove"
            ' Cleanly removes the HUD from the webpage DOM
            GetHudJs = "(function(){const h=document.getElementById('vba-bidi-hud'); if(h)h.remove();})();"
    End Select
End Function
