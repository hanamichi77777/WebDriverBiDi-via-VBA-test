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

' Class Name: BiDiCommandWrapper (W3C BiDi Standard & AI-Optimized)
' Summary   : High-Level Wrapper for Selenium WebDriver BiDi (W3C Standard)
' Features  :
'   - Resource Blocking via BiDi script.addPreloadScript (JS-side fetch/XHR/DOM hook)
'   - Smart Frame Recovery (Auto-handles frame reloads/stale contexts)
'   - Centralized Event Dispatching (Fixes event conflicts)
'   - Anti-Freeze Alert Handling
'   - AI-Ready Discovery Recorder
'   - Advanced Wait Strategies (Network Idle, DOM Stability, SPA Idle)
'
' CDP MIGRATION STATUS:
'   Phase 1 (Complete): AutoClicker -> script.addPreloadScript
'   Phase 2 (Complete): Resource Blocking -> script.addPreloadScript (JS-side blocking)
'   Phase 3 (Complete): CDP infrastructure removed (ExecuteCDPCommand, GetCdpSessionId, etc.)
'   ALL CDP DEPENDENCIES ELIMINATED. This class now operates on pure W3C BiDi standard.
'
' NOTE ON ERROR HANDLING STRATEGY (Strategic Buffer):
' This class utilizes "On Error Resume Next" (OERN) in specific interaction segments.
' The WebDriver BiDi protocol is highly asynchronous. In high-speed scenarios (e.g.,
' ServiceNow redirects), a "Race Condition" can occur where a Context ID or Realm
' becomes invalid between the verification and execution phases.
' OERN is used alongside ExecuteBiDiCommand's retry loop (MAX_CONTEXT_RECOVERY_RETRIES)
' to provide two-tier resilience: the retry loop handles recoverable context-loss errors,
' while OERN absorbs transient mismatches in non-critical paths (HUD updates, probe
' injections, discovery logging) where a failed call is acceptable.
'
' Dependencies: BiDiSocketCommunicator, WebJsonConverter, Microsoft Scripting Runtime
Private socket_ As BiDiSocketCommunicator
Private Const DEBUG_MODE As Boolean = True

' --- Foundation: ID Management ---
' [Phase 1] p_nextRequestId has been migrated to BiDiSocketCommunicator.
'           Wrapper accesses it via socket_.GetNextRequestId() when needed.
' [Phase 3] m_inflightRequests and p_pendingRequestCount have been migrated to
'           BiDiSocketCommunicator. Wrapper accesses via socket_.TrackInflightRequest,
'           socket_.UntrackInflightRequest, socket_.ResetInflightTracking, and
'           socket_.PendingRequestCount.

' --- Recovery & Stability ---
Private m_isRecovering As Boolean
Private m_lastTelemetryTick As Currency
Private Const MAX_CONTEXT_RECOVERY_RETRIES As Long = 5

' --- Navigation Detection (Predictive Probe Recovery) ---
Private m_navigationDetected As Boolean   ' Flag: navigationStarted event received
Private m_navigationContextId As String   ' Context ID where navigation was detected

' --- UI & Interaction Settings ---
Private m_waitStackDepth As Long ' Tracks recursion depth of active wait loops

' --- Internal State Variables ---
Private p_mainContextId As String
' [Phase 3] p_pendingRequestCount moved to Communicator (socket_.PendingRequestCount)
Private m_lastJsNow As Double ' Heartbeat tracker (performance.now from JS)
Private m_cachedRealmIds As Dictionary ' Cached Execution Realm ID
Private m_activeSubscriptions As Dictionary ' Event Management (Aggregation)
Private m_isProcessingQueue As Boolean
Private m_resourceBlockerScriptId As String ' [Phase 2] PreloadScript ID for resource blocking

' --- Configuration Constants ---
Private Const DEFAULT_IDLE_STABLE_MS As Long = 500
Private Const DEFAULT_IDLE_TIMEOUT_MS As Long = 10000
Private Const DEFAULT_STALE_RETRIES As Long = 3
Private Const SHORT_WAIT_MS As Long = 500
Private Const MAX_LOG_SIZE As Long = 500
Private Const MAX_MSG_SIZE As Long = 300000 ' Threshold to prevent VBA hanging
Private Const MIN_MSG_SIZE As Long = 10     ' Ignore extremely short messages
Private Const POST_SCRIPT_ID As Long = -1   ' Fixed ID for non-blocking Fire & Forget messages

' --- Noise Filtering State ---
Private Const DEFAULT_IGNORE_SELECTORS As String = "script, link, input[type='hidden'], .vba-ignore, #vba-bidi-hud"
Private Const DEFAULT_IGNORE_NETWORK As String = "google-analytics, doubleclick, /collect, /beacon, .woff2, favicon.ico"
Private p_userIgnoreSelectors As Collection
Private p_userIgnoreNetworkPatterns As Collection

' --- Internal Recorder Storage ---
Private p_recorderLogs As Collection
Private p_isRecording As Boolean
Private p_recorderStartTime As Double
Private p_excludeResources As Boolean

' Error Codes & Enumerations
Private Const ERR_BIDI_PROTOCOL As Long = vbObjectError + 1001
Private Const ERR_JS_EXCEPTION As Long = vbObjectError + 1002
Private Const ERR_ACTION_FAILED As Long = vbObjectError + 1003
Private Const ERR_CONSENSUS_HALT As Long = vbObjectError + 1010  ' Consensus failure: process halted for review

Public Enum BiDiMatchBy
    MatchByUrl = 0
    MatchByTitle = 1
End Enum

Public Enum BiDiContextType
    ContextTypeTab = 0
    ContextTypeWindow = 1
End Enum

Private Enum waitMode
    WaitMode_UrlContains
    WaitMode_SpaIdle
End Enum
' DEVELOPMENT & EXTENSION RULES FOR STABILITY
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
' 3. [TIMEOUT DELEGATION]: Public interaction methods SHOULD expose a 'timeoutMs' parameter.
'    - Action: Pass this parameter down to 'CallScript'.
'    Utility methods (ClearState, GetMainContextId, etc.) are exempt.
' 4. [VBA UI RESPONSIVENESS]: For any loop logic created within VBA (Do...Loop):
'    - Action: You MUST call 'WaitHeartbeat' or 'socket_.PumpMessages'.
'    - Rationale: This pumps the Windows message queue, preventing Excel from
'      entering a "Not Responding" state during high-precision waits.

' Initialization: Setup WebSocket and validate High-Resolution Timer frequency.
Private Sub Class_Initialize()
    ' --- 1. Collection & Dictionary Initializations ---
    ' Ensures all tracking objects are instantiated for re-entrancy safety.
    Set socket_ = New BiDiSocketCommunicator
    Set m_activeSubscriptions = New Dictionary
    ' [Phase 3] m_inflightRequests moved to Communicator (initialized in its Class_Initialize)
    Set m_cachedRealmIds = New Dictionary
    
    ' --- 2. Request ID Seeding ---
    ' [Phase 1] Request ID generation is now managed by BiDiSocketCommunicator.
    ' [Phase 3] Inflight tracking is now managed by BiDiSocketCommunicator.

    ' --- 3. State Reset ---
    m_waitStackDepth = 0
    m_isProcessingQueue = False
End Sub

' METHOD: ConnectTo
'  Connects to the WebSocket and subscribes to prompt events.
Public Sub ConnectTo(ByVal webSocketUrl As String)
    ' 1. Reset wrapper state before connecting.
    Me.ClearState
    
    ' 2. Establish the WebSocket connection.
    socket_.AttemptAutoConnect webSocketUrl
    
    ' 3. Mandatory Session Subscriptions (Internal logs).
    Me.ExecuteSessionSubscribe Array("log.entryAdded", "browsingContext.navigationStarted"), , True
    
    ' 4. GLOBAL SHADOW DOM UNLOCKER (Auto-Injection):
    ' Ensures all 'closed' shadow roots are accessible as 'open' across all frames.
    ' Called silently to keep the connection phase logs clean.
    Me.ExecuteEnableGlobalShadowUnlocker
End Sub

' Class_Terminate: Resource cleanup
Private Sub Class_Terminate()
    On Error Resume Next
    
    ' EMERGENCY UI RESTORATION: Ensures Excel returns to a usable state
    ' even if the macro was forcefully stopped.
    Application.Interactive = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    
    ' ASYNCHRONOUS HUD REMOVAL
    ' Uses Fire & Forget mode to attempt a clean browser-side cleanup
    ' without risking a final hang/whiteout during object destruction.
    If p_mainContextId <> "" Then
        Call Me.PostScript(GetHudJs("remove"), Nothing, p_mainContextId)
    End If
    
    ' RESOURCE RELEASE
    ' Explicitly shuts down the socket.
    Call Shutdown
    
    If DEBUG_MODE Then
        Debug.Print "BiDiCommandWrapper: Terminated and UI protection released."
    End If
End Sub
Public Property Get GetSocket() As BiDiSocketCommunicator
    Set GetSocket = socket_
End Property
Public Property Set SetSocket(Socket As BiDiSocketCommunicator)
    Set socket_ = Socket
End Property

' Central Event Dispatcher (Modified with Smart Filtering)
Private Sub DispatchGlobalEvents(ByVal json As Object)
    If Not json.Exists("method") Then Exit Sub
    Dim method As String: method = json("method")

    Select Case method
        ' --- BROWSER CONSOLE LOGS ---
        Case "log.entryAdded"
            Call HandleConsoleLog(json)

        ' --- NAVIGATION DETECTION (Predictive Probe Recovery) ---
        Case "browsingContext.navigationStarted"
            m_navigationDetected = True
            m_navigationContextId = GetPathString(json, "params.context")

            ' Invalidate the cached realm for the navigating context,
            ' as the execution realm will be destroyed by the navigation.
            If m_cachedRealmIds.Exists(m_navigationContextId) Then
                m_cachedRealmIds.Remove m_navigationContextId
            End If

            ' [Phase 3] Reset inflight network request tracking via Communicator.
            If m_navigationContextId = p_mainContextId Then
                socket_.ResetInflightTracking
            End If

            ' Log only when the navigation occurs in the main context.
            ' Sub-frame navigations (iframes) are handled silently to reduce log noise,
            ' as they are not consumed by the Predictive Recovery path (context mismatch).
            If DEBUG_MODE Then
                If m_navigationContextId = p_mainContextId Then
                    Debug.Print "BiDi: Navigation detected in main context " & _
                                Left$(m_navigationContextId, 16) & "... (Probe pre-invalidated, inflight reset)"
                End If
            End If

        ' --- [FIXED] NETWORK REQUEST START ---
        Case "network.beforeRequestSent"
            Dim reqId As String: reqId = GetPathString(json, "params.request.requestId")
            Dim reqUrl As String: reqUrl = GetPathString(json, "params.request.url")
            
            ' Only track significant (non-noise) requests.
            ' [Phase 3] Delegate storage to Communicator; noise filtering stays here.
            If Not IsNoiseUrl(reqUrl) Then
                socket_.TrackInflightRequest reqId
            End If

        ' --- [FIXED] NETWORK REQUEST END (Success or Error) ---
        Case "network.responseCompleted", "network.fetchError"
            Dim resId As String: resId = GetPathString(json, "params.request.requestId")
            
            ' [Phase 3] Delegate removal to Communicator.
            ' Safety guard (negative count prevention) is handled inside UntrackInflightRequest.
            socket_.UntrackInflightRequest resId

    End Select

    ' Store event in Discovery Recorder if active
    If p_isRecording Then
        FormatAndStoreLog json
    End If
End Sub

Public Property Get PendingRequestCount() As Long
    ' [Phase 3] Delegated to Communicator
    PendingRequestCount = socket_.PendingRequestCount
End Property

' [PUBLIC] GetIframeContextIdByUrl
'  Searches for a specific iframe context ID by its partial URL.
'              This method implements a "Fail-Fast" strategy by raising a VBA error
'              if the target frame is not located within the specified timeout.
Public Function GetIframeContextIdByUrl(ByVal partialUrl As String, Optional ByVal timeoutMs As Long = 5000) As String
    ' --- Input Validation ---
    Dim vTimeout As Long
    If timeoutMs = 0 Then
        vTimeout = 0
    Else
        vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 5000)
    End If
    
    Dim StartTick As Currency: StartTick = socket_.NowTicks()
    Dim ctxId As String

    Do
        ' 1. Fetch the browsing context tree and perform a recursive search
        Dim strRes As String: strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Dim contexts As Collection: Set contexts = GetPathObject(WebJsonConverter.ParseJson(strRes), "result.contexts")
        
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
        If IsTimeout(StartTick, CDbl(vTimeout)) Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:GetIframeContextIdByUrl", _
                      "Iframe context not found within " & vTimeout & "ms (Target URL pattern: " & partialUrl & ")"
        End If
    Loop
End Function

' [FIXED] Recursive search with Stack Guard and Depth Limit
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
        Debug.Print "BiDi Error in RecursiveFrameSearch: " & Err.description
        RecursiveFrameSearch = ""
    End If
End Function

' [PUBLIC] ExecuteFindWindowContextId
'  Locates a window context (tab/window) by URL or Title.
Public Function ExecuteFindWindowContextId(ByVal matchType As BiDiMatchBy, _
                                           ByVal matchValue As String, _
                                           Optional ByVal timeoutMs As Long = 5000, _
                                           Optional ByVal waitForComplete As Boolean = True) As String
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 5000)
    Dim StartTick As Currency: StartTick = socket_.NowTicks()
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
        If IsTimeout(StartTick, CDbl(vTimeout)) Then
            Err.Raise ERR_ACTION_FAILED, "BiDi:ExecuteFindWindowContextId", _
                      "Target window not found within " & vTimeout & "ms (Match: " & matchValue & ")"
        End If
    Loop
End Function

' FUNCTION: ExecuteCreateContext
'  Creates a new browsing context (tab or window).
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

' [PUBLIC] ExecuteGetTitleByContextId (Optimized)
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
    Dim capturedErr As Long: capturedErr = Err.Number
    Dim capturedDesc As String: capturedDesc = Err.description
    Dim capturedSrc As String: capturedSrc = Err.Source
    On Error GoTo 0
    
    ' Tolerate "no such alert" (prompt absent) but propagate socket/fatal errors
    If capturedErr <> 0 Then
        If InStr(1, capturedDesc, "no such alert", vbTextCompare) = 0 And _
           InStr(1, capturedDesc, "no user prompt", vbTextCompare) = 0 Then
            Err.Raise capturedErr, capturedSrc, capturedDesc
        End If
    End If
End Function

' [PUBLIC] ExecuteIsUrlContains
'  Checks if the current URL contains the partial string. Polls until match
'              is found or timeout occurs. Optionally waits for SPA stability.
' RETURNS    : True if URL contains the string within timeout, False otherwise.
Public Function ExecuteIsUrlContains(ByVal partialUrl As String, _
                                     Optional ByVal waitNetworkIdle As Boolean = True, _
                                     Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                     Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                     Optional ByVal contextId As String = "") As Boolean
    
    ' --- [ENTRY GUARD] Guard against 0ms when sync is requested ---
    If waitNetworkIdle Then
        If minStableMs = 0 Then minStableMs = DEFAULT_IDLE_STABLE_MS
        If maxTimeoutMs = 0 Then maxTimeoutMs = DEFAULT_IDLE_TIMEOUT_MS
    End If
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

' FUNCTION: ExecuteNavigateAndGetStatus
'  Public entry point for navigation. Returns "200" on success or "500" on error.
'          This high-level method leverages the private navigation engine.
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
        
        ' --- PHASE A: Navigation Command ---
        On Error Resume Next
        Call ExecuteBrowsingContextNavigate(targetUrl, modeStr, vMinStable, vMaxTimeout, targetCtx)
        Dim navErr As Long: navErr = Err.Number
        Dim navDesc As String: navDesc = Err.description
        On Error GoTo 0
        
        ' Fatal socket errors (600+) should not be retried
        If navErr >= 600 Then
            ExecuteNavigateAndGetStatus = "500"
            Exit Function
        End If
        
        If navErr = 0 Then
            ' --- PHASE B: Liveness Verification ---
            If Not waitNetworkIdle Then
                ' Mode "none": Perform quick URL validation since the protocol returns immediately.
                On Error Resume Next
                Dim curUrl As String: curUrl = GetCurrentUrl(targetCtx, True)
                On Error GoTo 0
                If Not (InStr(1, curUrl, "chrome-error", vbTextCompare) > 0 Or curUrl = "") Then
                    isNavigated = True
                End If
            Else
                ' Mode "complete": Success is implied by a zero navErr after stability sync.
                isNavigated = True
            End If
        End If
        
        If isNavigated Then
            ExecuteNavigateAndGetStatus = "200"
            Exit Function
        End If

        ' Wait for the environment to settle before retry
        WaitHeartbeat 1000
    Next attempt
    
    ExecuteNavigateAndGetStatus = "500"
End Function

' [PUBLIC] ExecuteEnableResourceBlocking (Phase 2: JS Injection via preloadScript)
' MIGRATION : Replaces CDP Network.enable + Fetch.enable + Network.setBlockedURLs pipeline.
'
' ARCHITECTURE DECISION:
'   BiDi's network.addIntercept + UrlPattern has fundamental incompatibility with
'   CDP-style wildcard patterns ("*.png", "*doubleclick*"). The UrlPattern API
'   rejects '*' in hostname/pathname as "Forbidden characters". A blanket intercept
'   (no urlPatterns) causes WebSocket deadlock because every request blocks until
'   VBA responds synchronously, saturating the single-threaded message loop.
'
'   SOLUTION: Use script.addPreloadScript to inject a JS-side blocker that:
'     1. Wraps window.fetch and XMLHttpRequest.prototype.open/send
'     2. Checks each request URL against the blocking patterns (purely in JS)
'     3. Aborts matching requests immediately with AbortController (fetch) or
'        xhr.abort() (XHR), never leaving the browser event loop
'     4. Runs in every new document automatically (preloadScript lifecycle)
'
'   This approach has ZERO VBA-side overhead per request, no WebSocket traffic
'   for blocking decisions, and no risk of deadlock. It's architecturally identical
'   to ad-blocker browser extensions.
'
' BEHAVIOR CONTRACT (unchanged for callers):
'   - Accepts a single URL pattern string or a VBA Array of pattern strings.
'   - Patterns use CDP-style wildcards: * = match zero or more characters.
'   - Returns the preloadScript registration result on success.
Public Function ExecuteEnableResourceBlocking(ByVal patterns As Variant) As String
    ' 1. Normalize patterns into a JS array literal
    Dim jsPats As String: jsPats = "["
    Dim Item As Variant, first As Boolean: first = True
    
    If IsArray(patterns) Then
        For Each Item In patterns
            If Not first Then jsPats = jsPats & ","
            jsPats = jsPats & """" & EscapeJsString(CStr(Item)) & """"
            first = False
        Next
    Else
        jsPats = jsPats & """" & EscapeJsString(CStr(patterns)) & """"
    End If
    jsPats = jsPats & "]"
    
    ' 2. Build the preloadScript JS payload
    '    This function is called by the browser for every new document.
    '    It installs fetch/XHR wrappers that check URLs against the pattern list.
    Dim js As String
    js = "function() { "
    js = js & "if(window.__vbaBlockerInstalled) return; window.__vbaBlockerInstalled = true; "
    js = js & "var pats = " & jsPats & "; "
    
    ' Pattern matcher: converts CDP wildcards to a test function
    ' Each '*' becomes '.*' in a RegExp for full glob support
    js = js & "var rxs = pats.map(function(p) { "
    js = js & "  var e = p.replace(/[.+^${}()|[\\]\\\\]/g, '\\\\$&'); "
    js = js & "  e = e.replace(/\\*/g, '.*'); "
    js = js & "  return new RegExp(e, 'i'); "
    js = js & "}); "
    js = js & "var isBlocked = function(u) { return rxs.some(function(r) { return r.test(u); }); }; "
    
    ' Wrap fetch: abort blocked requests via AbortController
    js = js & "if(typeof window.fetch === 'function' && !window.fetch.__vbaBlocked) { "
    js = js & "  var oF = window.fetch; "
    js = js & "  window.fetch = function(input, init) { "
    js = js & "    var u = (typeof input === 'string') ? input : (input && input.url) || ''; "
    js = js & "    if(isBlocked(u)) { var c = new AbortController(); c.abort(); "
    js = js & "      return Promise.reject(new DOMException('Blocked by VBA','AbortError')); } "
    js = js & "    return oF.apply(this, arguments); "
    js = js & "  }; window.fetch.__vbaBlocked = true; "
    js = js & "} "
    
    ' Wrap XHR: abort blocked requests in send()
    js = js & "if(!XMLHttpRequest.prototype.__vbaBlocked) { "
    js = js & "  var oOpen = XMLHttpRequest.prototype.open; "
    js = js & "  XMLHttpRequest.prototype.open = function(m, u) { "
    js = js & "    this.__vbaUrl = u; return oOpen.apply(this, arguments); }; "
    js = js & "  var oSend = XMLHttpRequest.prototype.send; "
    js = js & "  XMLHttpRequest.prototype.send = function() { "
    js = js & "    if(this.__vbaUrl && isBlocked(this.__vbaUrl)) { this.abort(); return; } "
    js = js & "    return oSend.apply(this, arguments); }; "
    js = js & "  XMLHttpRequest.prototype.__vbaBlocked = true; "
    js = js & "} "
    
    ' Block image/script/link loads via a MutationObserver on src/href attributes
    js = js & "var blockAttr = function(el) { "
    js = js & "  var s = el.src || el.href || ''; "
    js = js & "  if(s && isBlocked(s)) { el.src = ''; el.href = ''; "
    js = js & "    if(el.srcset) el.srcset = ''; "
    js = js & "    el.remove(); } }; "
    js = js & "try { new MutationObserver(function(ms) { "
    js = js & "  ms.forEach(function(m) { "
    js = js & "    m.addedNodes.forEach(function(n) { "
    js = js & "      if(n.nodeType === 1) { blockAttr(n); "
    js = js & "        n.querySelectorAll && n.querySelectorAll('img,script,link,iframe').forEach(blockAttr); } "
    js = js & "    }); }); "
    js = js & "}).observe(document.documentElement || document, {childList:true, subtree:true}); } catch(e) {} "
    
    js = js & "}"
    
    ' 3. Register via BiDi preloadScript (same pattern as Shadow DOM Unlocker)
    Dim params As New Dictionary
    params.Add "functionDeclaration", js
    
    Dim strRes As String
    On Error Resume Next
    strRes = ExecuteBiDiCommand("script.addPreloadScript", params, True)
    Dim regErr As Long: regErr = Err.Number
    Dim regErrDesc As String: regErrDesc = Err.description
    On Error GoTo 0
    
    If regErr = 0 Then
        Dim scriptId As String
        scriptId = GetPathString(WebJsonConverter.ParseJson(strRes), "result.script")
        
        If scriptId <> "" Then
            m_resourceBlockerScriptId = scriptId
            Debug.Print "BiDi: Resource blocker registered via preloadScript (ID: " & scriptId & ")"
        Else
            Debug.Print "BiDi Warning: Resource blocker registration succeeded but no script ID returned."
        End If
    Else
        Debug.Print "BiDi Error: Resource blocker registration failed: " & regErrDesc
    End If
    
    ' 4. Also inject into the CURRENT document immediately
    '    (preloadScript only fires on NEW documents)
    '    Strategy: Wrap the entire function declaration in an IIFE call.
    '    CallScript -> WrapInSafetyShell wraps in async function(){ _tk = async () => { jsBody } }
    '    so jsBody = "(" & js & ")()" invokes the blocker function inside the safety shell.
    If regErr = 0 Then
        On Error Resume Next
        Call Me.CallScript("(" & js & ")()", Nothing, , False, False, 5000, True)
        On Error GoTo 0
    End If
    
    ExecuteEnableResourceBlocking = strRes
End Function

' [PUBLIC] ExecuteDisableResourceBlocking
'  Removes the resource blocker preloadScript.
'              Note: Already-injected scripts in open documents cannot be unwrapped,
'              but new navigations will not receive the blocker.
Public Sub ExecuteDisableResourceBlocking()
    If m_resourceBlockerScriptId = "" Then Exit Sub
    
    Dim params As New Dictionary
    params.Add "script", m_resourceBlockerScriptId
    
    On Error Resume Next
    ExecuteBiDiCommand "script.removePreloadScript", params, True
    On Error GoTo 0
    
    If DEBUG_MODE Then
        Debug.Print "BiDi: Resource blocker preloadScript removed (ID: " & m_resourceBlockerScriptId & ")"
    End If
    
    m_resourceBlockerScriptId = ""
End Sub

' [PRIVATE] EscapeJsString
'  Escapes a VBA string for safe embedding inside a JS string literal.
Private Function EscapeJsString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, "'", "\'")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    EscapeJsString = s
End Function

' [PUBLIC] ExecuteIsElementVisible
'  Checks if an element is rendered and visible in the viewport.
'              Returns True only if the element exists in the DOM and is functionally visible.
'              Supports 0ms searchTimeoutMs for high-speed immediate polling.
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
    ' strRes will be a flattened JSON string returned by the response pipeline.
    Dim strRes As String
    strRes = ExecuteBaseAction(targetCtx, "visibility_check", args, waitNetworkIdle, vMinStable, True, DEFAULT_STALE_RETRIES, vMaxTimeout)
    
    ' 7. [FIXED] Parse Result using the unified pipeline.
    On Error Resume Next
    Dim resultVal As Variant
    resultVal = GetCleanResult(strRes)
    
    ' If the result is not Empty, convert the JS boolean to a VBA Boolean.
    If Not IsEmpty(resultVal) Then
        ExecuteIsElementVisible = CBool(resultVal)
    End If
    On Error GoTo 0
End Function

' ExecuteSessionSubscribe: Subscribes to BiDi events with optional context filtering.
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
        If contexts.count > 0 Then ctxId = "|" & CStr(contexts(1))
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
    If eventsToRequest.count = 0 Then
        ' [S-Rank Fix] Respect the Silent flag even during early exit to keep logs clean.
        If DEBUG_MODE And Not Silent Then
            Debug.Print "BiDi: Subscription skipped (Events already active)."
        End If
        ExecuteSessionSubscribe = "{""status"":""already_subscribed""}"
        Exit Function
    End If
    ' 5. Prepare Payload for protocol-level subscription
    Dim params As New Dictionary
    Dim finalArr() As String
    ReDim finalArr(eventsToRequest.count - 1)
    
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

' [PRIVATE] ExecuteBrowsingContextNavigate (Refactored for Anti-Freeze)
'  Initiates navigation without blocking at the socket level.
'          Uses VBA-side polling to monitor stability and handle timeouts.
Private Function ExecuteBrowsingContextNavigate(ByVal url As String, _
                                    Optional ByVal waitMode As String = "complete", _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "") As String
                                                
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' [FIX] Reset only transient navigation/network state.
    ' [Phase 3] Reset inflight tracking via Communicator so stale request IDs
    ' from the previous page do not cause count mismatches.
    socket_.ResetInflightTracking
    
    Dim bidiWaitLevel As String
    bidiWaitLevel = "interactive"
    
    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "url", url
    params.Add "wait", bidiWaitLevel
    
    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.navigate", params)
    Dim navJson As Object: Set navJson = ValidateProtocolResponse(strRes, "browsingContext.navigate")

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

' [PUBLIC] ExecuteClickByXPath
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

' [PUBLIC] ExecuteInputValueByXPath
Public Function ExecuteInputValueByXPath(ByVal xpath As String, _
                                         ByVal valueToSet As String, _
                                         Optional ByVal searchTimeoutMs As Long = 10000, _
                                         Optional ByVal waitNetworkIdle As Boolean = True, _
                                         Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                         Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                         Optional ByVal contextId As String = "", _
                                         Optional ByVal useKeyEvents As Boolean = True) As String
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 1. Resolve element
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteInputValueByXPath")

    ' 2. Pack arguments: [sharedId, textValue]
    Dim args As New Collection
    Dim a1 As New Dictionary: a1.Add "sharedId", sharedId: args.Add a1
    Dim a2 As New Dictionary: a2.Add "type", "string": a2.Add "value", valueToSet: args.Add a2

    If Not useKeyEvents Then
        ' FAST MODE: Bulk execCommand('insertText') via standard pipeline.
        ExecuteInputValueByXPath = ExecuteBaseAction(targetCtx, "input", args, waitNetworkIdle, minStableMs, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
    Else
        ' KEY EVENTS MODE: Bypasses ExecuteBaseAction to prevent SPA probe/sync
        ' from interfering with the combobox active state. The entire
        ' click -> delete -> type sequence runs as one uninterrupted JS function.
        '
        Dim keySafetyMs As Long
        keySafetyMs = maxTimeoutMs
        If keySafetyMs < 20000 Then keySafetyMs = 20000
        
        ExecuteInputValueByXPath = CallScript( _
            GetActionJs("input_keys"), args, targetCtx, False, False, keySafetyMs)
    End If
End Function

' [PUBLIC] ExecuteSelectValueByXPath (Upgraded: Auto-Detect Hybrid Engine)
'  Centrally handles both standard HTML <select> tags and semantic pseudo
'              dropdowns (Wiz/React comboboxes) seamlessly without user awareness.
Public Function ExecuteSelectValueByXPath(ByVal xpath As String, _
                                          ByVal valueOrText As String, _
                                          Optional ByVal selectByText As Boolean = False, _
                                          Optional ByVal searchTimeoutMs As Long = 10000, _
                                          Optional ByVal waitNetworkIdle As Boolean = True, _
                                          Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                          Optional ByVal contextId As String = "") As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 1. Resolve target element identifier securely
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteSelectValueByXPath")
    
    ' 2. Inspect element tag name via low-cost script to detect architecture type
    Dim chkArgs As New Collection, tDic As New Dictionary: tDic.Add "sharedId", sharedId: chkArgs.Add tDic
    Dim isStandardSelect As Boolean
    isStandardSelect = CBool(GetCleanResult(Me.CallScript("return arguments[0].tagName==='SELECT';", chkArgs, targetCtx, False, False, 2000, True)))
    
    ' 3. Route to the optimized execution path based on detection result
    If isStandardSelect Then
        ' Path A: Native HTML <select> element tracking
        Dim jsKey As String: If selectByText Then jsKey = "select_text" Else jsKey = "select"
        Dim args As New Collection: args.Add tDic
        Dim a2 As New Dictionary: a2.Add "type", "string": a2.Add "value", valueOrText: args.Add a2
        ExecuteSelectValueByXPath = ExecuteBaseAction(targetCtx, jsKey, args, waitNetworkIdle, DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
    Else
        ' Path B: Semantic Pseudo Dropdown element (Wiz/React / W3C Combobox)
        ' Phase 1: Click the combo box wrapper to trigger option list population
        Dim clickArgs As New Collection: clickArgs.Add tDic
        Call ExecuteBaseAction(targetCtx, "click", clickArgs, True, DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
        ' Phase 2: Resolve the dynamic option node and fire the final selection click
        '   - No quotes:    'text'
        '   - Single only:  "text"
        '   - Both:         concat('part1',"'","part2")
        Dim optXPath As String: optXPath = "//*[@role='option'][contains(., " & EscapeXPathString(valueOrText) & ")]"
        Dim optSharedId As String: optSharedId = ResolveXPathOrThrow(optXPath, targetCtx, searchTimeoutMs, "ExecuteSelectValueByXPath")
        Dim optClickArgs As New Collection, oDic As New Dictionary: oDic.Add "sharedId", optSharedId: optClickArgs.Add oDic
        ExecuteSelectValueByXPath = ExecuteBaseAction(targetCtx, "click", optClickArgs, waitNetworkIdle, DEFAULT_IDLE_STABLE_MS, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
    End If
    
End Function

' Locator Helpers
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

' [INTERNAL] GetFindElementSharedIdByXPath
'  Elements may not be ready immediately. This loop polls with a throttle
'              to prevent CPU saturation and "Unresponsive" Excel state.
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
    Dim StartTick As Currency: StartTick = socket_.NowTicks()
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
            If nodes.count > 0 Then
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
        If IsTimeout(StartTick, CDbl(vTimeoutMs)) Then Exit Do
    Loop
    
    ' FAILURE: Return empty string if not found within the designating window.
    GetFindElementSharedIdByXPath = ""
End Function

' [PUBLIC] GetCurrentUrl (Robust Version)
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

' [PUBLIC] ExecuteShadowClick
'  Penetrates Shadow DOM boundaries to click a nested element.
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

Private Function ExecuteBiDiCommand(ByVal method As String, _
                                     ByRef params As Dictionary, _
                                     Optional ByVal Silent As Boolean = False) As String
    
    ' Guard: Ensure the communication socket is available.
    If socket_ Is Nothing Then Err.Raise 91, "ExecuteBiDiCommand", "Socket object is not set."

    Dim retryCount As Long
    Dim strRes As String
    Dim localErrNum As Long
    Dim localErrDesc As String
    
    ' Retry Loop: Attempt to recover from transient environment or socket failures.
    Do While retryCount <= MAX_CONTEXT_RECOVERY_RETRIES
        
        ' 1. [Phase 1] Delegate ID generation, payload construction, size validation,
        '    serialization, and transmission to Communicator's SendCommand.
        '    Logging is also handled by SendCommand when Silent=False.
        On Error Resume Next
        strRes = socket_.SendCommand(method, params, Silent)
        localErrNum = Err.Number: localErrDesc = Err.description
        On Error GoTo 0
        
        ' 2. Protocol & Business Validation
        If localErrNum = 0 Then
            ' Check for "Context Lost" errors which indicate the tab/frame was destroyed.
            If IsContextLostError(strRes) Or (InStr(1, strRes, "chrome-error", vbTextCompare) > 0) Then
                localErrNum = ERR_BIDI_PROTOCOL
                localErrDesc = "Environment desync (Context Lost) detected in response."
            End If
            
            ' SUCCESS PATH: If no errors occurred during transmission and validation.
            If localErrNum = 0 Then
                ' Clear the internal error counter of the socket upon a verified healthy response.
                socket_.ClearInternalErrorCounter
                ExecuteBiDiCommand = strRes
                
                ' Exit the function immediately to prevent reaching the recovery block.
                Exit Function
            End If
        End If

        ' --- FAIL-FAST LOGIC ---
        ' If the error is an explicit action failure (e.g., Element not found),
        ' do NOT attempt recovery. Stop the loop and propagate the error immediately.
        If localErrNum = ERR_ACTION_FAILED Then Exit Do

        ' --- RECOVERY LOGIC ---
        ' If we haven't exhausted our retries, attempt to heal the environment.
        If retryCount < MAX_CONTEXT_RECOVERY_RETRIES And Not m_isRecovering Then
            
            ' Pre-recovery check: If the underlying socket is dead, no point in retrying.
            If Not socket_.IsConnected Then
                Debug.Print "BiDi [FATAL]: Socket connection lost. Aborting command retries."
                Exit Do
            End If
            
            m_isRecovering = True
            
            ' Progressive Wait: Increase wait time based on the retry attempt.
            WaitHeartbeat 800 * (retryCount + 1)
            
            If DEBUG_MODE Then
                Debug.Print "BiDi [RECOVERY]: Attempting to refresh context (" & (retryCount + 1) & "/" & MAX_CONTEXT_RECOVERY_RETRIES & ")..."
            End If
            
            ' Attempt to refresh the main context mapping silently.
            On Error Resume Next
            Call UpdateMainContextId(True, True) ' forceRefresh:=True, Silent:=True
            Dim recoveryErr As Long: recoveryErr = Err.Number
            On Error GoTo 0
            
            ' If context refresh succeeded, update the IDs within the original 'params' dictionary.
            If recoveryErr = 0 Then
                If p_mainContextId <> "" Then Call UpdateContextAndRealm(params, p_mainContextId)
                m_isRecovering = False
                retryCount = retryCount + 1
            Else
                ' If we cannot even refresh the context tree, the session is likely unrecoverable.
                Debug.Print "BiDi [RECOVERY FAIL]: Failed to identify any valid browsing context."
                m_isRecovering = False
                Exit Do
            End If
        Else
            ' Max retries reached or already in a recovery state.
            Exit Do
        End If
    Loop

    ' --- FINAL ERROR PROPAGATION ---
    ' Ensure the total retry count and the last error message are propagated to the user.
    Err.Raise IIf(localErrNum = 0, ERR_ACTION_FAILED, localErrNum), _
              "BiDi:ExecuteBiDiCommand", _
              "Command '" & method & "' failed after " & retryCount & " recovery attempt(s). " & _
              "Last Error: " & localErrDesc
End Function

' --- [Phase 1] LogTransmission is retained for retry-context logging ---
' Note: Primary REQ/RES logging is now handled by socket_.SendCommand.
'       This helper is kept for recovery-phase diagnostic logging where
'       the retry attempt number is needed for analysis.
Private Sub LogTransmission(ByVal method As String, ByVal id As Long, ByVal Msg As String, ByVal prefix As String, ByVal attempt As Long)
    Dim displayMsg As String: displayMsg = Msg
    If Len(displayMsg) > 500 Then displayMsg = Left$(displayMsg, 497) & "..."
    socket_.LogEvent method & " (ID: " & id & ", Attempt: " & attempt & ")", displayMsg, "[" & prefix & "]"
End Sub

' INTERNAL HELPER: IsContextLostError (STRICT VERSION)
Private Function IsContextLostError(ByVal response As String) As Boolean
    ' Early Exit: If the word "result" exists at the root, it is a protocol SUCCESS.
    If InStr(1, response, """result"":", vbBinaryCompare) > 0 Then
        IsContextLostError = False
        Exit Function
    End If

    ' Parse under OERN; if ParseJson fails, json remains Nothing and we exit.
    On Error Resume Next
    Dim json As Object: Set json = WebJsonConverter.ParseJson(response)
    On Error GoTo 0
    
    If json Is Nothing Then Exit Function
    
    ' ONLY trigger if there is a root "error" key (BiDi Protocol level error)
    If json.Exists("error") And json.Exists("message") Then
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
End Function

' UpdateContextAndRealmRecursive: Recursively scans a node to update context and realm IDs.
' This ensures that even deeply nested IDs in script arguments are refreshed after recovery.
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
            For i = 1 To node.count
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

' UpdateContextAndRealm: Entry point for refreshing IDs within a parameter dictionary.
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

' ClearState: Resets internal state and cleans up browser-side side effects.
Public Sub ClearState()
    p_mainContextId = ""
    m_resourceBlockerScriptId = "" ' [Phase 2] Clear resource blocker reference
    
    ' [Phase 3] Reset inflight tracking via Communicator
    If Not socket_ Is Nothing Then socket_.ResetInflightTracking

    ' [FIX] Use New Dictionary instead of Nothing for m_cachedRealmIds.
    ' Prevents Error 424 if other methods access the cache before JIT initialization.
    Set m_cachedRealmIds = New Dictionary

    ' Reset the subscription registry.
    ' Using a New Dictionary ensures that the wrapper re-subscribes on the next connection.
    Set m_activeSubscriptions = New Dictionary
End Sub

' FUNCTION: GetMainContextId
'  Returns the currently active main context ID.
'              If not yet set, it triggers a discovery of the browser tree.
Public Function GetMainContextId() As String
    If p_mainContextId = "" Then
        ' MACHINERY: Pass True to the 'Silent' parameter to hide id:2
        GetMainContextId = UpdateMainContextId(True, True)
    Else
        GetMainContextId = p_mainContextId
    End If
End Function

' FUNCTION: UpdateMainContextId
'  Refreshes the browsing context tree. Raises error if no context is available.
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
        ' [FIX] Validate that the provided contextId exists in the browser tree.
        If contextId = p_mainContextId Then
            ResolveTargetContext = contextId
        Else
            ' Verify the context actually exists by checking the browser tree
            Dim treeJson As String
            treeJson = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary, True)
            If InStr(1, treeJson, contextId, vbTextCompare) > 0 Then
                ResolveTargetContext = contextId
            Else
                Err.Raise ERR_ACTION_FAILED, "BiDi:ResolveTargetContext", _
                          "Context ID not found in browser tree: " & contextId & _
                          ". The tab/window may have been closed or the ID is invalid."
            End If
        End If
    Else
        ' Returns cached main context ID or discovers a new one if necessary
        ResolveTargetContext = GetMainContextId()
    End If
End Function

' FUNCTION: GetScriptRealmId (Improved for Service Worker environments)
'  Retrieves and caches the Execution Realm ID for a specific context.
'              Prioritizes "window" type realms to ensure UI/DOM script compatibility.
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
    Dim realmErr As Long: realmErr = Err.Number
    Dim realmErrDesc As String: realmErrDesc = Err.description
    On Error GoTo 0
    
    If realmErr <> 0 Then
        If DEBUG_MODE Then
            Debug.Print "BiDi: getRealms failed for context " & Left$(targetCtx, 16) & _
                        "... (Err " & realmErr & ": " & realmErrDesc & ")"
        End If
        GetScriptRealmId = ""
        Exit Function
    End If
    
    Dim realms As Collection
    Set realms = GetPathObject(WebJsonConverter.ParseJson(jsonRes), "result.realms")
    
    Dim foundRealmId As String: foundRealmId = ""
    
    If Not realms Is Nothing Then
        If realms.count > 0 Then
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

' [PUBLIC] ExecuteRegisterAutoClickerByXPath
'  Registers a background observer via BiDi preloadScript to click a button
'              as soon as it appears. The script is automatically injected into every
'              new document (including navigations and new tabs), matching the behavior
'              of the previous CDP Page.addScriptToEvaluateOnNewDocument approach.
' MIGRATION : Phase 1 CDP->BiDi. Uses script.addPreloadScript (W3C standard) instead
'             of goog:cdp.sendCommand("Page.addScriptToEvaluateOnNewDocument").
'             Pattern is identical to ExecuteEnableGlobalShadowUnlocker.
Public Sub ExecuteRegisterAutoClickerByXPath(ByVal xpath As String, _
                                             Optional ByVal timeoutMs As Long = 30000)
                                             
    ' Subscribe silently to log entries for AutoClicker confirmation
    Me.ExecuteSessionSubscribe "log.entryAdded", , True
    
    ' Build the AutoClicker JS payload with safe XPath/timeout substitution
    Dim sourceJs As String: sourceJs = GetActionJs("autoclicker")
    Dim safeXpath As String: safeXpath = Replace(xpath, "'", "\'")
    sourceJs = Replace(sourceJs, "{{XPATH}}", safeXpath)
    sourceJs = Replace(sourceJs, "{{TIMEOUT}}", CStr(timeoutMs))
    
    ' Wrap the IIFE payload inside a function declaration for BiDi's
    ' script.addPreloadScript contract. The outer function() { ... } envelope
    ' is required by the spec; the inner IIFE executes immediately when the
    ' browser invokes this function on each new document load.
    Dim params As New Dictionary
    params.Add "functionDeclaration", "function() { " & sourceJs & " }"
    
    ' Execute via BiDi standard command (Silent:=True to reduce log noise)
    Dim strRes As String
    On Error Resume Next
    strRes = ExecuteBiDiCommand("script.addPreloadScript", params, True)
    Dim regErr As Long: regErr = Err.Number
    Dim regErrDesc As String: regErrDesc = Err.description
    On Error GoTo 0
    
    ' Validate result: Extract the PreloadScriptId for confirmation
    If regErr = 0 Then
        Dim scriptId As String
        scriptId = GetPathString(WebJsonConverter.ParseJson(strRes), "result.script")
        
        If scriptId <> "" Then
            Debug.Print "BiDi: AutoClicker registered via preloadScript (ID: " & scriptId & ") for XPath: " & xpath
        Else
            Debug.Print "BiDi Warning: AutoClicker command succeeded but no script ID returned."
        End If
    Else
        Debug.Print "BiDi Error: AutoClicker registration failed: " & regErrDesc
    End If
End Sub

' INTERNAL HELPER: BuildShadowSelectorArg
'  Normalizes input to a JSON array. Accepts either a single String
'              or a VBA Array of Strings for nested Shadow DOM penetration.
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

' Discovery & Recorder API
Public Sub StartDiscoveryLog(Optional ByVal excludeImagesAndCss As Boolean = True)
    Set p_recorderLogs = New Collection
    p_isRecording = True
    p_recorderStartTime = socket_.NowMs() / 1000
    p_excludeResources = excludeImagesAndCss
    Dim events(3) As String
    events(0) = "network.beforeRequestSent"
    events(1) = "network.responseCompleted"
    events(2) = "browsingContext.load"
    events(3) = "browsingContext.domContentLoaded"
    Me.ExecuteSessionSubscribe events

    Debug.Print "BiDi Recorder: Started. (Filter Resources: " & excludeImagesAndCss & ")"
End Sub

' [FIXED] StopAndSaveDiscoveryLog with explicit memory cleanup
Public Sub StopAndSaveDiscoveryLog(ByVal filePath As String)
    ' [FIX] Flush the incoming message queue first.
    ' Since ProcessRecorderQueue calls DispatchGlobalEvents, which checks p_isRecording,
    ' we must process the queue while the flag is still True.
    Call ProcessRecorderQueue

    ' Now safely stop the recording flag.
    p_isRecording = False

    If p_recorderLogs Is Nothing Then Exit Sub
    
    ' Save logic
    If p_recorderLogs.count > 0 Then
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

    Set p_recorderLogs = Nothing
End Sub

' [PUBLIC] RecordEventsForSeconds
'  Keeps the WebSocket queue moving for a specified duration to record events.
Public Sub RecordEventsForSeconds(ByVal seconds As Long)
    ' 1. Logic Check
    If seconds <= 0 Then Exit Sub

    ' --- UI LOCKING ---
    ' Disable user interaction to prevent entering "Cell Edit Mode" during recording.
    Application.Interactive = False
    Application.Cursor = xlWait

    ' 2. Initialize timing
    Dim startMs As Double: startMs = socket_.NowMs()
    Dim durationMs As Double: durationMs = CDbl(seconds) * 1000#

    On Error GoTo ErrorHandler

    ' 3. Polling Loop
    Do While (socket_.NowMs() - startMs) < durationMs
        ' Process the incoming WebSocket message queue
        ProcessRecorderQueue

        ' Unified wait with re-entrancy protection via YieldExecution.
        ' The 5-second threshold is preserved to suppress excessive DoEvents
        ' side-effects during long recording sessions (add-in/COM interrupts).
        If durationMs > 5000 Then
            ' Long recording: Throttled PumpMessages + sleep-only YieldExecution
            socket_.PumpMessages 100
            socket_.YieldExecution 100, False
        Else
            ' Short recording: Higher responsiveness with guarded DoEvents
            socket_.YieldExecution 100, socket_.AllowUiPump
        End If
    Loop

Cleanup:
    ' [FIX] Always reset the processing lock regardless of how the loop finished.
    ' If an error occurred inside ProcessRecorderQueue, this flag must be cleared
    ' to allow subsequent message processing.
    m_isProcessingQueue = False
    
    ' --- UI RESTORATION ---
    ' Always ensure the UI is unlocked before returning control to the user.
    Application.Interactive = True
    Application.Cursor = xlDefault
    Exit Sub

ErrorHandler:
    ' [FIX v2] Properly deactivate the error handler via Resume before cleanup.
    ' Without Resume, VBA's error handler remains "active" after GoTo Cleanup.
    ' If any statement in Cleanup raises a new error while the handler is active,
    ' VBA terminates the call stack immediately instead of handling it gracefully.
    Debug.Print "BiDi Warning in RecordEventsForSeconds: " & Err.description
    Resume Cleanup
End Sub

' PUBLIC SUB: ProcessRecorderQueue
'  Processes the incoming message queue with re-entrancy protection.
'          [Phase 2] Simplified: JSON retrieval, validation, and parsing are now
'          delegated to Communicator's GetParsedEvent(). This method focuses purely
'          on dispatch logic and error management.
Public Sub ProcessRecorderQueue()
    ' --- 1. Re-entrancy Guard ---
    If m_isProcessingQueue Then Exit Sub
    
    ' Set the lock
    m_isProcessingQueue = True
    Dim json As Object
    Dim processedCount As Long
    Dim consecutiveErrors As Long
    Const MAX_MESSAGES_PER_BATCH As Long = 1000
    
    On Error GoTo Cleanup
    
    ' Main processing loop
    Do While processedCount < MAX_MESSAGES_PER_BATCH
        ' [Phase 2] Delegate message retrieval + size validation + JSON parse
        ' to Communicator. Returns Nothing if buffer is empty or message invalid.
        Set json = socket_.GetParsedEvent(MIN_MSG_SIZE, MAX_MSG_SIZE)
        If json Is Nothing Then Exit Do
        
        ' --- ISOLATED DISPATCH ---
        ' Protect the loop from individual event processing failures
        On Error Resume Next
        DispatchGlobalEvents json
        
        If Err.Number <> 0 Then
            consecutiveErrors = consecutiveErrors + 1
            If DEBUG_MODE Then
                Debug.Print "BiDi: Dispatch error #" & consecutiveErrors & ": " & Err.description
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
        
        processedCount = processedCount + 1
        
        ' Yield to OS every 100 messages to keep UI responsive
        If (processedCount Mod 100) = 0 Then socket_.PumpMessages 100
    Loop

Cleanup:
    ' --- 2. Release Lock ---
    m_isProcessingQueue = False
    
    If Err.Number <> 0 Then
        Debug.Print "BiDi Critical Error in ProcessRecorderQueue: " & Err.description
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
    ' [FIX v2] Replaced On Error Resume Next with explicit key existence checks.
    ' OERN was masking structural errors (e.g., missing "url" key on a valid "request"
    ' object). Each access is now guarded by Dictionary.Exists() or TypeName checks,
    ' ensuring only genuinely missing payloads are skipped while real bugs surface.

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

    relTime = Format((socket_.NowMs() / 1000) - p_recorderStartTime, "0.00") & "s"

    Select Case method
        Case "network.beforeRequestSent"
            ' Guard: Validate required nested keys before access
            If Not params.Exists("request") Then Exit Sub
            If Not IsObject(params("request")) Then Exit Sub
            Dim req As Object: Set req = params("request")
            If Not req.Exists("url") Then Exit Sub
            Dim url As String: url = CStr(req("url"))

            If p_excludeResources Then
                If IsNoiseUrl(url) Then Exit Sub

                ' Header-based optional filter: skip pure image requests
                If req.Exists("headers") Then
                    On Error Resume Next
                    Dim hdrJson As String
                    hdrJson = LCase(WebJsonConverter.ConvertToJson(req("headers")))
                    If Err.Number = 0 And Len(hdrJson) > 0 Then
                        ' Check if any Accept header contains "image/"
                        ' Pattern: {"name":"accept","value":"image/..."}
                        If InStr(hdrJson, """accept""") > 0 And InStr(hdrJson, "image/") > 0 Then
                            ' Verify it's not a mixed accept (e.g. "text/html, image/webp")
                            ' by checking if image/ appears near accept
                            If InStr(hdrJson, "image/webp") > 0 Or InStr(hdrJson, "image/png") > 0 _
                                Or InStr(hdrJson, "image/jpeg") > 0 Or InStr(hdrJson, "image/gif") > 0 _
                                Or InStr(hdrJson, "image/svg") > 0 Then
                                ' Only skip if it looks like a dedicated image request
                                ' (Accept starts with image/), not a page request that also accepts images
                                If InStr(hdrJson, """value"":""image/") > 0 Then
                                    Err.Clear
                                    On Error GoTo 0
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If

            Dim reqMethod As String: reqMethod = ""
            If req.Exists("method") Then reqMethod = CStr(req("method"))
            logLine = "[" & relTime & "] [REQ] " & reqMethod & " - " & url

        Case "network.responseCompleted"
            If Not params.Exists("response") Then Exit Sub
            If Not IsObject(params("response")) Then Exit Sub
            Dim res As Object: Set res = params("response")
            If Not res.Exists("url") Then Exit Sub
            If Not res.Exists("mimeType") Then Exit Sub
            Dim resUrl As String: resUrl = CStr(res("url"))
            Dim mime As String: mime = CStr(res("mimeType"))

            If p_excludeResources Then
                If IsNoiseUrl(resUrl) Or IsNoiseMime(mime) Then Exit Sub
            End If

            Dim resStatus As String: resStatus = ""
            If res.Exists("status") Then resStatus = CStr(res("status"))
            logLine = "[" & relTime & "] [RES] " & resStatus & " [" & mime & "] " & resUrl

        Case "browsingContext.load", "browsingContext.domContentLoaded"
            If Not params.Exists("url") Then Exit Sub
            logLine = "[" & relTime & "] [NAV] " & Replace(method, "browsingContext.", "") & " - " & CStr(params("url"))

        Case "log.entryAdded"
            If Not params.Exists("level") Then Exit Sub
            If CStr(params("level")) = "error" Then
                Dim logText As String: logText = ""
                If params.Exists("text") Then logText = CStr(params("text"))
                logLine = "[" & relTime & "] [CONSOLE-ERR] " & logText
            End If
    End Select

    If logLine <> "" Then
        If p_recorderLogs.count >= MAX_LOG_SIZE Then p_recorderLogs.Remove 1
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

' FUNCTION: QuerySpaProbe
'  Executes the JS query snippet and parses the result.
'              Returns Nothing if the probe is missing or reports 'active: false'.
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

' PROCEDURE: RecoverSpaProbe
'  Nukes the state cache and re-injects the probe into the fresh context.
Private Sub RecoverSpaProbe(ByVal contextId As String)
    Dim targetId As String: targetId = contextId
    Dim isMainTarget As Boolean

    ' Determine if we are recovering the main context (empty or matching current main)
    isMainTarget = (targetId = "" Or targetId = p_mainContextId)

    ' [FIX] Resolve empty contextId to mainContextId BEFORE cache operations.
    If targetId = "" Then targetId = p_mainContextId

    ' 1. Invalidate ONLY the specific context's realm ID cache.
    If Not m_cachedRealmIds Is Nothing Then
        If targetId <> "" And m_cachedRealmIds.Exists(targetId) Then
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

' PUBLIC SUB: WaitForNetworkIdleSpa
'  Entry point for deep SPA-aware synchronization.
'          This version includes Entry Guards to handle 0ms inputs caused by the
'          updated SanitizeMs logic, ensuring stability consensus isn't bypassed.
Public Sub WaitForNetworkIdleSpa(ByVal minStableMs As Long, _
                                 ByVal maxTimeoutMs As Long, _
                                 Optional ByVal contextId As String = "")
    
    ' --- [ENTRY GUARD] Guard against 0ms in stability check ---
    If minStableMs = 0 Then minStableMs = DEFAULT_IDLE_STABLE_MS
    If maxTimeoutMs = 0 Then maxTimeoutMs = DEFAULT_IDLE_TIMEOUT_MS
    
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
    waitCtx.Add "stableStart", socket_.NowMs()
    waitCtx.Add "missingProbeCount", 0 ' Used for multi-stage JIT recovery
    
    ' Execute the high-precision wait engine with SPA Idle mode.
    ' MasterWaitEngine will perform final sanitization using the guarded values.
    MasterWaitEngine WaitMode_SpaIdle, maxTimeoutMs, 100, waitCtx
End Sub

' FUNCTION: EvaluateSpaIdleCondition
'  Final consensus engine for SPA stability.
'              Integrates:
'              1. Network Silence (BiDi Protocol + JS Hooks)
'              2. DOM Stability (MutationObserver + Prototype Hooks)
'              3. Diagnostic Logging (Immediate Window + Discovery Recorder)
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
    
    ' Cache inflight URL list into the context for diagnostic use on timeout.
    ' Overwrites each poll cycle so the list is always current.
    If probe.Exists("inflightUrls") Then
        Set ctx("lastInflightUrls") = probe("inflightUrls")
    End If
    
    ' Calculate current idle duration reported by the browser
    jsIdleDuration = jsNow - lastActivityTs

    ' --- 3. Activity Logging (Discovery Recorder Integration) ---
    If probe.Exists("changes") Then
        Dim changes As Collection: Set changes = probe("changes")
        If Not changes Is Nothing Then
            If changes.count > 0 Then
                Dim v As Variant, logSummary As String
                For Each v In changes: logSummary = logSummary & v & " ": Next
                logSummary = Trim(logSummary)
                
                If Len(logSummary) > 500 Then logSummary = Left$(logSummary, 497) & "..."
                If DEBUG_MODE Then Debug.Print "[BiDi SPA Activity] " & logSummary
                
                If p_isRecording Then
                    Dim relTime As String
                    relTime = Format((socket_.NowMs() / 1000) - p_recorderStartTime, "0.00") & "s"
                    p_recorderLogs.Add "[" & relTime & "] [SPA-ACT] " & logSummary
                End If
            End If
        End If
    End If

    ' --- 4. Self-Correction Logic (Hard-Sync) ---
    ' [Phase 3] Access and write via Communicator property.
    If (inflFetch = 0 And inflXhr = 0) And (socket_.PendingRequestCount > 0) Then
        If jsIdleDuration > (CDbl(minStableMs) * 2#) Then
            If DEBUG_MODE Then Debug.Print "BiDi: Hard-Syncing pendingRequestCount to 0 due to prolonged silence."
            socket_.PendingRequestCount = 0
        End If
    End If

    ' --- 5. Consensus Evaluation ---
    ' Network Idle: Both protocol and JS hooks must report zero activity.
    netIdle = (socket_.PendingRequestCount <= 0) And (inflFetch = 0) And (inflXhr = 0)

    ' DOM Stability: No mutations within the required threshold.
    activityStable = (jsIdleDuration >= CDbl(minStableMs))

    ' --- 6. Final Consensus Window Check ---
    If netIdle And activityStable Then
        ' Check if the stable state has persisted for the designated window.
        If (currentMs - CDbl(ctx("stableStart"))) >= CDbl(minStableMs) Then
            If DEBUG_MODE Then
                Call socket_.LogEvent("Consensus reached (" & minStableMs & "ms window satisfied)", "", ">>> [BiDi SPA Stable]")
            End If
            
            If p_isRecording Then
                Dim successRelTime As String
                successRelTime = Format((socket_.NowMs() / 1000) - p_recorderStartTime, "0.00") & "s"
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

' PROCEDURE: StartSpaIdleProbe
'  Executes the JavaScript probe injection.
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

' JavaScript Error Decoder (Enhanced Debugging)
' Formats BiDi exceptionDetails into a readable string including stack trace
Private Function FormatJavaScriptError(ByVal details As Object) As String
    Dim Msg As String
    Dim stack As String
    Dim frames As Collection
    Dim frame As Object
    Dim i As Long

    ' 1. Basic error info
    Msg = "JS Exception: " & details("text")
    
    ' 2. Location info
    If details.Exists("lineNumber") Then
        Msg = Msg & " [Line: " & details("lineNumber") & ", Col: " & details("columnNumber") & "]"
    End If
    ' 3. Exception object description (if exists)
    If details.Exists("exception") Then
        If details("exception").Exists("description") Then
            Msg = Msg & vbCrLf & "Description: " & details("exception")("description")
        End If
    End If
    ' 4. Stack Trace decoding
    If details.Exists("stackTrace") Then
        Set frames = details("stackTrace")("callFrames")
        stack = "--- JS Stack Trace ---"
        For i = 1 To frames.count
            Set frame = frames(i)
            stack = stack & vbCrLf & "  at " & frame("functionName") & " (" & frame("url") & ":" & frame("lineNumber") & ":" & frame("columnNumber") & ")"
        Next i
        Msg = Msg & vbCrLf & stack
    End If

    FormatJavaScriptError = Msg
End Function

' AddIdleIgnoreSelector: Registers a CSS selector to be ignored by MutationObserver.
' Users can call this from a Standard Module to bypass noisy elements (e.g., clocks).
Public Sub AddIdleIgnoreSelector(ByVal cssSelector As String)
    If p_userIgnoreSelectors Is Nothing Then Set p_userIgnoreSelectors = New Collection
    p_userIgnoreSelectors.Add cssSelector
End Sub

' AddIdleIgnoreNetworkPattern: Registers a URL substring to be ignored by Network tracking.
' Useful for bypassing background logging, telemetry, or specific third-party ads.
Public Sub AddIdleIgnoreNetworkPattern(ByVal urlPattern As String)
    If p_userIgnoreNetworkPatterns Is Nothing Then Set p_userIgnoreNetworkPatterns = New Collection
    p_userIgnoreNetworkPatterns.Add urlPattern
End Sub

' ClearIdleFilters: Resets all user-defined noise filters.
Public Sub ClearIdleFilters()
    Set p_userIgnoreSelectors = Nothing
    Set p_userIgnoreNetworkPatterns = Nothing
End Sub

' MergeToJsArray: Combines a comma-separated default string and a VBA Collection
' into a JSON-formatted JavaScript array string.
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
    For i = 1 To items.count
        result = result & """" & Replace(items(i), """", "\""") & """"
        If i < items.count Then result = result & ","
    Next i
    result = result & "]"
    
    MergeToJsArray = result
End Function

' INTERNAL JS HELPER: GetActionJs
'  Provides centralized JavaScript snippets for all browser actions.
'              Ensures the SPA watchdog is reset immediately before execution.
Private Function GetActionJs(ByVal actionType As String) As String
    Dim js As String: js = ""
    
    ' SPA probe reset: forces 'Busy' state in MasterWaitEngine to prevent race conditions.
    Dim rp As String
    rp = "if(window.__vbaIdleProbe) window.__vbaIdleProbe.lastMutationTs = performance.now(); "
    
    Select Case LCase(actionType)
        Case "click"
            js = js & "try { const e=arguments[0]; e.scrollIntoView({block:'center'}); e.focus(); " & rp & "e.click(); return true; } catch(err) { throw err; }"
                 
        Case "input"
            ' Bulk execCommand('insertText') with select() overwrite. Fast for standard forms.
            ' Uses e.select() to preserve framework state, then insertText replaces selected text.
            js = js & "try { const e=arguments[0],v=arguments[1]; e.scrollIntoView({block:'center',inline:'center'}); e.click(); e.focus(); " & rp & "e.select(); "
            js = js & "const s=document.execCommand('insertText',false,v); if(!s){e.value='';e.value=v;e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}));} "
            js = js & "e.blur(); return true;} catch(err){throw err;}"
                 
        Case "select"
            js = js & "try { const s=arguments[0],v=arguments[1]; s.scrollIntoView({block:'center',inline:'center'}); s.focus(); " & rp & "s.value=v; "
            js = js & "s.dispatchEvent(new Event('input',{bubbles:true})); s.dispatchEvent(new Event('change',{bubbles:true})); s.blur(); return true;} catch(err){throw err;}"

        Case "select_text"
            ' Normalized text matching for <select> option elements.
            js = js & "try { const s=arguments[0],t=arguments[1],n=(str)=>str.replace(/[\s\u00A0]+/g,' ').trim(); let f=false; const target=n(t); "
            js = js & "for(let i=0;i<s.options.length;i++){if(n(s.options[i].text)===target){s.value=s.options[i].value;f=true;break;}} "
            js = js & "if(!f)throw new Error('Option text not found: '+t); s.scrollIntoView({block:'center'}); s.focus(); " & rp
            js = js & "s.dispatchEvent(new Event('input',{bubbles:true})); s.dispatchEvent(new Event('change',{bubbles:true})); s.blur(); return true;} catch(err){throw err;}"

        Case "visibility_check"
            ' Returns true if element is connected, not display:none/hidden, has opacity,
            ' and occupies visual space (width/height > 0 or has client rects).
            js = js & "const e=arguments[0]; if(!e||!e.isConnected)return false; const s=window.getComputedStyle(e),r=e.getBoundingClientRect(); "
            js = js & "return(s.display!=='none'&&s.visibility!=='hidden'&&parseFloat(s.opacity||'1')>0&&(r.width>0||r.height>0||e.getClientRects().length>0));"
   
        Case "shadow_click"
            ' Penetrate Shadow DOM boundaries to find and click a nested element.
            js = js & "const sels=arguments[0],tout=arguments[1],start=Date.now(); "
            js = js & "const fd=(r,s)=>{let n=r.querySelector(s);if(n)return n;for(let c of r.querySelectorAll('*')){if(c.shadowRoot&&(n=fd(c.shadowRoot,s)))return n;}return null;}; "
            js = js & "while(Date.now()-start<tout){ hb(); let el=document,lastS=null; "
            js = js & "for(const s of sels){el=fd(el.shadowRoot||el,s);if(!el){lastS=s;break;}} "
            js = js & "if(el){el.scrollIntoView({block:'center'});el.focus();" & rp & "el.click();return true;} "
            js = js & "if(Date.now()-start+100>tout)throw new Error('ShadowSearchError: Failed at ['+lastS+'] within '+tout+'ms'); "
            js = js & "await new Promise(r=>setTimeout(r,100));} "
            
        Case "input_keys"
            ' --- Phase 0: activate and resolve the real input element ---
            ' Clicks the resolved element, then polls for the *actual* active input.
            ' Wiz/Lit SPAs (e.g. Google Flights) often destroy the trigger on click
            ' and spawn a replacement <input> -- sometimes multiple times rapidly.
            ' We must NOT lock onto an intermediate element the framework will discard.
            ' Strategy: require the same activeElement on two consecutive polls
            ' (stability check) so transient mid-swap elements pass through without
            ' early break, while resolving quickly (2 x 15 ms = 30 ms extra latency).
            
            ' --- Phase 0: activate and resolve the real input element ---
            js = js & "try{let e=arguments[0],v=arguments[1],o=e;e.scrollIntoView({block:'center',inline:'center'});e.click();e.focus();"
            js = js & "let prev=null,stb=0;for(let t=Date.now();Date.now()-t<3000;){hb();let a=document.activeElement;"
            js = js & "if(a&&a.tagName==='INPUT'&&(a.type==='text'||a.type===''||!a.type)&&typeof a.selectionStart==='number'){if(a===prev){stb++;if(stb>=2){e=a;break;}}else{prev=a;stb=1;}}else{prev=null;stb=0;}await new Promise(r=>setTimeout(r,15));}"
            ' --- Phase 1: clear existing value safely ---
            js = js & "(e.select?e.select():document.execCommand('selectAll'));await new Promise(r=>setTimeout(r,80));document.execCommand('delete');await new Promise(r=>setTimeout(r,80));"
            js = js & "if(e.value){if(e.select)e.select();document.execCommand('forwardDelete');await new Promise(r=>setTimeout(r,80));}if(e.value){e.value='';e.dispatchEvent(new Event('input',{bubbles:true}));await new Promise(r=>setTimeout(r,80));}" & rp & "await new Promise(r=>setTimeout(r,40));"
            ' --- Phase 2: per-char insertText with activeElement tracking ---
            js = js & "for(let c of v){hb();let s=(typeof e.selectionStart==='number'?e.selectionStart:0);document.execCommand('insertText',false,c);for(let k=0;k<80;k++){let a=document.activeElement;if(a&&a!==e&&a.tagName==='INPUT'&&(a.type==='text'||!a.type))e=a;if(document.activeElement===e&&typeof e.selectionStart==='number'&&e.selectionStart>=s+1)break;await new Promise(r=>setTimeout(r,10));}}"
            ' --- Phase 3: final validation ---
            js = js & "if(v.length>0&&(!e.value||e.value.length<v.length))throw new Error('Wiz Input Validation Failed: '+e.value);return true;}catch(err){throw err;}"
            
        Case "autoclicker"
            ' Background MutationObserver-based clicker (injected via BiDi script.addPreloadScript).
            ' The IIFE is wrapped in a function() { ... } envelope by the caller
            ' (ExecuteRegisterAutoClickerByXPath) to satisfy the addPreloadScript contract.
            js = js & "(function(x,t){const s=Date.now(),f=()=>{const e=document.evaluate(x,document,null,9,null).singleNodeValue;"
            js = js & "if(e&&(e.offsetWidth>0||e.offsetHeight>0)){e.click();e.dispatchEvent(new MouseEvent('click',{bubbles:true}));"
            js = js & "console.log('BiDi-AutoClicker: Target clicked');return 1}return 0};"
            js = js & "if(f())return 1;const o=new MutationObserver(()=>{if(f()||(Date.now()-s>t))o.disconnect()});"
            js = js & "o.observe(document.body||document,{childList:1,subtree:1,attributes:1});return 1;})('{{XPATH}}',{{TIMEOUT}});"
            
    End Select
    GetActionJs = js
End Function

' CRITICAL: JS AUTHORING CONTRACT
' Any JavaScript executed via this method must adhere to the following contract
' to maintain communication with VBA.
'
' 1. [FORBIDDEN] Synchronous Blocking Loops
'    Will trigger WinHttp timeouts and cause Excel to "Not Responding" (whiteout).
'    BAD : while(true) { ... }
'    GOOD: while(cond) { hb(); await new Promise(r => setTimeout(r, 50)); }
'
' 2. [REQUIRED] Asynchronous Chunking for Long Tasks
'    Any process exceeding 100ms must yield control to the event loop using 'await'.
'    GOOD: await heavyProcess();
'
' 3. [REQUIRED] Invoke hb() at the Loop Head
'    The 'hb' (Heartbeat) closure is essential to reset VBA-side observation timeouts.
'    GOOD: while(cond) { hb(); /* process */ }
'
' * NOTE: Due to VBA's single-threaded nature, it is impossible to forcibly terminate
'   an infinite loop originating from the JS side.
'   THIS CONTRACT IS EXPERIMENTALLY VERIFIED:
'   Confirmed that Excel whites out and fails to recover even after 30 seconds
'   when executing 'while(true) {}'.
' ========================================================================================' ========================================================================================
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
    ' [STRUCTURAL CONSTRAINT] awaitPromise is intentionally hardcoded to True.
    ' WrapInSafetyShell wraps jsBody in an "async function" that returns a Promise
    ' containing the JSON-stringified result. If awaitPromise were False, the BiDi
    ' protocol would return the raw Promise object instead of the resolved value,
    ' causing all result extraction (UnwrapSafetyShell / GetCleanResult) to fail.
    ' The caller's awaitPromise parameter is therefore NOT forwarded here.
    ' Fire-and-forget execution should use PostScript instead of CallScript.
    params.Add "awaitPromise", True
    
    ' Execute command with the Silent flag
    Dim rawRes As String: rawRes = ExecuteBiDiCommand("script.callFunction", params, Silent)
    Dim json As Object: Set json = ValidateProtocolResponse(rawRes, "script.callFunction")
    Dim procObj As Object: Set procObj = UnwrapSafetyShell(json, "script.callFunction")
    CallScript = WebJsonConverter.ConvertToJson(procObj)
End Function

' IsTimeout: Precision-safe timeout check.
Private Function IsTimeout(ByVal startTicks As Currency, ByVal timeoutMs As Double) As Boolean
    If timeoutMs <= 0 Then Exit Function
    ' Simple comparison
    IsTimeout = socket_.ElapsedMs(startTicks) > timeoutMs
End Function

' Internal Helper: Centralizes background tasks during a wait loop.
' Keeps the WebSocket queue moving and the Excel UI responsive.
Private Sub WaitHeartbeat(Optional ByVal sleepMs As Long = 50)
    ' 1. Logic-layer responsibility: Flush the incoming BiDi event queue.
    ' Ensures VBA state stays in sync with the browser's network/DOM state.
    ProcessRecorderQueue

    ' 2. Throttled DoEvents via PumpMessages (preserves existing 100ms gate).
    socket_.PumpMessages 100

    ' 3. Delegate physical sleep to Communicator's YieldExecution.
    ' DoEvents is set to False because PumpMessages above already handled it.
    ' The 50ms floor guard preserves the original SanitizeMs safety behavior.
    Dim vSleep As Long: vSleep = sleepMs
    If vSleep < 50 Then vSleep = 50

    socket_.YieldExecution vSleep, False
End Sub

' FUNCTION: GetProbeJs
'  Provides the comprehensive JavaScript payload for SPA monitoring.
'              Integrated with Smart XPath generation (Ancestor-ID based).
Private Function GetProbeJs(ByVal probeType As String) As String
    Dim js As String: js = ""
    Dim selArray As String: selArray = MergeToJsArray(DEFAULT_IGNORE_SELECTORS, p_userIgnoreSelectors)
    Dim netArray As String: netArray = MergeToJsArray(DEFAULT_IGNORE_NETWORK, p_userIgnoreNetworkPatterns)

    Select Case LCase$(probeType)
        Case "inject"
            ' --- SPA Monitoring Probe: Now filters out Sizzle noise ---
            js = js & "(function(){if(window.__vbaSpaProbeInstalled)return true;window.__vbaSpaProbeInstalled=true;var selectors=" & selArray & ",ignoreNetwork=" & netArray & ","
            js = js & "p=window.__vbaIdleProbe||{inflightFetchCount:0,inflightXhrCount:0,lastMutationTs:performance.now(),changeList:[],inflightUrls:new Set()};window.__vbaIdleProbe=p;if(!p.inflightUrls)p.inflightUrls=new Set();"
            js = js & "var gX=function(e){if(!e||e.nodeType!==1)return '';var s=[];while(e&&e.nodeType===1){if(e.id){s.unshift('//*[@id=\''+e.id+'\']');return s.join('/')}"
            js = js & "var i=1,b=e.previousSibling;while(b){if(b.nodeType===1&&b.tagName===e.tagName)i++;b=b.previousSibling}s.unshift(e.tagName.toLowerCase()+'['+i+']');e=e.parentNode}return '/'+s.join('/')};"
            
            ' --- MutationObserver: tracks DOM/attribute changes for SPA idle detection ---
            ' gX(node) computes a full XPath for the changed node.
            ' iN(node, selectors) returns truthy if node matches any ignore selector (CSS).
            ' lg(entry) appends to changeList (capped at 50) and updates lastMutationTs.
            ' uT() updates the last-mutation timestamp (used for idle consensus).
            js = js & "var uT=()=>{p.lastMutationTs=performance.now()},lg=(e)=>{"
            js = js & "if(e && !/sizzle|signon_javascript|conversation_button/.test(e) && p.changeList.indexOf(e)===-1){"
            js = js & "if(p.changeList.length<50)p.changeList.push(e);uT()}};"
            js = js & "iN=(n,s)=>{if(!n)return 1;var t=(n.nodeType===1)?n:n.parentElement;return t?s.some((x)=>t.closest(x)!==null):0};"
            ' --- HTMLInputElement.value setter hook: detects programmatic value changes ---
            js = js & "try{const d=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value');Object.defineProperty(HTMLInputElement.prototype,'value',{set:function(v){lg('PROP:'+gX(this)+'[value]');return d.set.call(this,v)},get:function(){return d.get.call(this)},configurable:true})}catch(e){}"
            ' --- Observer registration: childList + attributes on entire document ---
            js = js & "try{new MutationObserver((ms)=>{ms.forEach((m)=>{if(iN(m.target,selectors))return;var path=gX(m.target),entry=(m.type==='attributes')?'ATTR:'+path+'['+m.attributeName+']':'DOM:'+path;"
            
            ' Attribute Filtering: Ignore mutations on technical/transient attributes
            js = js & "if(entry.indexOf('sizzle')===-1) lg(entry);})}).observe(document,{childList:true,subtree:true,attributes:true})}catch(e){}"
            
            ' --- fetch() wrapper: intercepts fetch calls to track inflight count ---
            ' Wraps window.fetch once (__vbaWrapped flag). Increments inflightFetchCount on
            ' call, decrements on resolve/reject. Skipped if URL matches ignoreNetwork.
            js = js & "if(typeof window.fetch==='function'&&!window.fetch.__vbaWrapped){var oF=window.fetch;window.fetch=function(u,o){var r=(typeof u==='string')?u:(u&&u.url)||'';if(ignoreNetwork.some((pat)=>r.indexOf(pat)!==-1))return oF.apply(this,arguments);"
            js = js & "p.inflightFetchCount++;p.inflightUrls.add(r);var dec=()=>{p.inflightFetchCount--;p.inflightUrls.delete(r);uT()};return oF.apply(this,arguments).then((r2)=>{dec();return r2}).catch((e)=>{dec();throw e})};window.fetch.__vbaWrapped=true}"
            ' --- XHR wrapper: intercepts XMLHttpRequest open/send to track inflight count ---
            ' Wraps prototype.open to capture URL, prototype.send to increment counter.
            ' Decrements on 'loadend' event. Skipped if URL matches ignoreNetwork.
            js = js & "if(!XMLHttpRequest.prototype.__vbaWrapped){var oO=XMLHttpRequest.prototype.open,oS=XMLHttpRequest.prototype.send;XMLHttpRequest.prototype.open=function(m,u){this.__vbaUrl=u;this.__vbaHooked=true;return oO.apply(this,arguments)};"
            js = js & "XMLHttpRequest.prototype.send=function(){var s=this,u=this.__vbaUrl||'';if(ignoreNetwork.some((pat)=>u.indexOf(pat)!==-1)){s.__vbaHooked=false;return oS.apply(this,arguments)}p.inflightXhrCount++;p.inflightUrls.add(u);"
            js = js & "var dn=()=>{if(s.__vbaHooked){p.inflightXhrCount--;p.inflightUrls.delete(u);s.__vbaHooked=false;uT()}s.removeEventListener('loadend',dn)};this.addEventListener('loadend',dn);return oS.apply(this,arguments)};XMLHttpRequest.prototype.__vbaWrapped=true}return true})();"
        Case "query"
            ' Returns JSON snapshot of probe state: active flag, timing, inflight counts,
            ' pending DOM change list, and inflight URLs. Clears changeList after read.
            js = js & "var p=window.__vbaIdleProbe;if(!p)return JSON.stringify({active:false});"
            js = js & "var res={active:true,now:performance.now(),lastMutationTs:p.lastMutationTs,inflightFetchCount:p.inflightFetchCount,inflightXhrCount:p.inflightXhrCount,changes:p.changeList,inflightUrls:Array.from(p.inflightUrls||[])};"
            js = js & "p.changeList=[];return JSON.stringify(res);"
    End Select
    GetProbeJs = js
End Function

' Internal Helper: Builds a target dictionary for BiDi script commands.
' Priority: Uses Realm ID if provided for execution stability, falls back to Context ID.
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

' Configuration: Public properties to manage noise filtering from Standard Modules.
Public Property Let IgnoreSelectors(ByVal cssSelectors As String)
    If p_userIgnoreSelectors Is Nothing Then Set p_userIgnoreSelectors = New Collection
    p_userIgnoreSelectors.Add cssSelectors
End Property

Public Property Let IgnoreNetworkPatterns(ByVal urlPattern As String)
    If p_userIgnoreNetworkPatterns Is Nothing Then Set p_userIgnoreNetworkPatterns = New Collection
    p_userIgnoreNetworkPatterns.Add urlPattern
End Property

' Internal Helper: Ensures a Dictionary object is initialized.
' Returns a new Dictionary if the input is Nothing, otherwise returns the input.
Private Function EnsureDictionary(ByVal dic As Dictionary) As Dictionary
    If dic Is Nothing Then
        Set EnsureDictionary = New Dictionary
    Else
        Set EnsureDictionary = dic
    End If
End Function

' [INTERNAL] MasterWaitEngine (Refined Anti-Freeze & Stack-Safe Version)
'  High-precision polling engine for all wait operations.
' UPGRADED:
'   - Telemetry Gating: Uses instance-level 'm_lastTelemetryTick' to ensure consistent
'     100ms HUD updates, decoupled from Windows message pump frequency.
'   - Stack Safety: Timer is reset only at root entry (depth=1) to allow nested calls
'     to share a continuous timing context without flickering.
'   - CleanExit Fix: Restored correct logic (errNum = 0) to ensure HUD removal on success.
Private Function MasterWaitEngine(ByVal mode As waitMode, _
                                  ByVal timeoutMs As Long, _
                                  ByVal intervalMs As Long, _
                                  ByRef ctx As Dictionary) As Boolean
    
    ' 1. Parameter Sanitization
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeoutMs, "timeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    Dim vInterval As Long: vInterval = SanitizeMs(intervalMs, "intervalMs", 150, 100)
    Dim errNum As Long, errDesc As String, errSrc As String
    Dim StartTick As Currency: StartTick = socket_.NowTicks()
    Dim success As Boolean: success = False

    ' 2. Re-entrancy & UI State Management
    m_waitStackDepth = m_waitStackDepth + 1
    
    If m_waitStackDepth = 1 Then
        ' --- ROOT ENTRY: Lock UI and Initialize Timer ---
        ' PHYSICAL LOCK: Prevent user intervention during critical sync operations.
        Application.Interactive = False
        Application.Cursor = xlWait
        
        ' Initialize the instance-level timer for HUD updates.
        ' Resetting to 0 ensures the first update fires immediately.
        m_lastTelemetryTick = 0
        
        ' JIT HUD: Inject the UI container into the browser.
        On Error Resume Next
        Call Me.CallScript(GetHudJs("init"), Nothing, ResolveTargetContext(CStr(ctx("contextId"))), False, False, 2000, True)
        On Error GoTo 0
    End If
    ' 3. Main Polling Loop
    On Error GoTo ErrorHandler
    Do
        ' [CRITICAL] Processes WebSocket queue and pumps Windows messages.
        ' Maintains Excel UI responsiveness during the wait.
        WaitHeartbeat vInterval
        
        ' --- REFINED TELEMETRY GATING ---
        ' Uses dedicated instance-level timer to poll UpdateWaitTelemetry every 100ms.
        ' This is independent of DoEvents frequency to ensure a stable HUD refresh rate.
        If socket_.ElapsedMs(m_lastTelemetryTick) > 100 Then
            Call UpdateWaitTelemetry(mode, ctx)
            m_lastTelemetryTick = socket_.NowTicks()
        End If
        
        ' --- LAYER 2: LOGICAL EVALUATION & CONSENSUS CHECK ---
        ' Check if the target condition (URL match or SPA stability) is met.
        On Error Resume Next
        success = EvaluateWaitCondition(mode, ctx)
        If Err.Number <> 0 Then
            errNum = Err.Number: errDesc = Err.description: errSrc = Err.Source
            On Error GoTo ErrorHandler: Exit Do
        End If
        On Error GoTo ErrorHandler
        
        ' --- LAYER 3: EXIT & TIMEOUT LOGIC ---
        If success Then Exit Do
        
        ' Standard Logical Timeout check (Error 504).
        If IsTimeout(StartTick, CDbl(vTimeout)) Then
            errNum = 504
            errDesc = "Wait condition timed out (" & vTimeout & "ms)."
            errSrc = "LogicalTimeout"
            Exit Do
        End If
    Loop

' CLEANUP & RESTORATION
CleanExit:
    ' STEP 1: UI RESTORATION
    If errNum <> 0 Then
        m_waitStackDepth = 0
    Else
        m_waitStackDepth = m_waitStackDepth - 1
        If m_waitStackDepth < 0 Then m_waitStackDepth = 0
    End If

    If m_waitStackDepth = 0 Then
        Application.Interactive = True
        Application.Cursor = xlDefault
        Application.StatusBar = False
    End If
    
    ' STEP 2: BROWSER SIDE CLEANUP
    ' [FIXED] Restored correct removal condition:
    ' Remove HUD if operation succeeded (errNum = 0) OR if we are at the outermost call.
    On Error Resume Next
    If errNum = 0 Or m_waitStackDepth = 0 Then
        Call Me.CallScript(GetHudJs("remove"), Nothing, ResolveTargetContext(CStr(ctx("contextId"))), False, False, 1000, True)
    End If
    
    ' STEP 3: RESOURCE RELEASE
    On Error GoTo 0
    
    ' STEP 4: ERROR PROPAGATION (504, 505, or 517 from JS)
    If errNum <> 0 Then
        ' --- Consensus Failure Analysis ---
        ' Check if pending network requests caused the failure.
        ' If yes: HALT processing and log URLs for triage.
        ' If no:  Propagate as a normal error (absorbed by TrySyncPhase).
        On Error Resume Next
        Application.Interactive = True
        
        ' 1. Inflight request URLs from JS probe (last snapshot before timeout)
        Dim inflightDetail As String: inflightDetail = ""
        If ctx.Exists("lastInflightUrls") Then
            Dim urls As Collection: Set urls = ctx("lastInflightUrls")
            If Not urls Is Nothing Then
                Dim urlCount As Long: urlCount = urls.count
                Dim ui As Long
                For ui = 1 To urlCount
                    If ui <= 10 Then
                        inflightDetail = inflightDetail & vbCrLf & "  " & CStr(urls(ui))
                    End If
                Next ui
                If urlCount > 10 Then
                    inflightDetail = inflightDetail & vbCrLf & "  ... and " & (urlCount - 10) & " more"
                End If
            End If
        End If
        On Error GoTo 0
        
        If Len(inflightDetail) > 0 Then
            ' --- HALT: Pending requests are blocking consensus ---
            ' Output to Immediate Window and raise non-absorbable error.
            Debug.Print String(60, "=")
            Debug.Print "[" & Format(Now, "hh:mm:ss") & "] CONSENSUS HALT"
            Debug.Print "  Pending Requests:" & inflightDetail
            Debug.Print String(60, "=")
            
            Dim haltMsg As String
            haltMsg = "CONSENSUS HALT: SPA stability consensus was not reached." & vbCrLf & _
                      "Check Immediate Window for details."
            
            Err.Raise ERR_CONSENSUS_HALT, "BiDi:MasterWaitEngine:" & errSrc, haltMsg
        Else
            ' --- NORMAL TIMEOUT: No pending requests (e.g. DOM mutations only) ---
            ' Propagate as a standard error; TrySyncPhase will absorb this.
            Err.Raise errNum, "BiDi:MasterWaitEngine:" & errSrc, errDesc
        End If
    End If
    
    MasterWaitEngine = success
    Exit Function

ErrorHandler:
    errNum = Err.Number: errDesc = Err.description: errSrc = Err.Source
    Resume CleanExit
End Function

' FUNCTION    : EvaluateWaitCondition
'  Centralized condition checker for the wait engine.
'               Handles logical state evaluation and detects "Consensus Panic" (505).
Private Function EvaluateWaitCondition(ByVal mode As waitMode, ByRef ctx As Dictionary) As Boolean
    ' Maximum allowed re-injections of the monitoring script before raising a panic error.
    Const MAX_RECOVERY_ATTEMPTS As Long = 8
    
    ' Safety Guard: Ensure context has not been prematurely destroyed.
    If ctx Is Nothing Then Exit Function
    
    Dim targetCtx As String: targetCtx = ResolveTargetContext(CStr(ctx("contextId")))
    
    Select Case mode
        Case WaitMode_UrlContains
            ' Simple logical check: URL pattern matching.
            EvaluateWaitCondition = (InStr(GetCurrentUrl(targetCtx, True), CStr(ctx("partialUrl"))) > 0)
        
        Case WaitMode_SpaIdle
            ' Query the SPA Probe injected in the browser's JavaScript execution context.
            Dim probe As Dictionary: Set probe = QuerySpaProbe(targetCtx)
            Dim isProbeActive As Boolean: isProbeActive = False
            If Not probe Is Nothing Then
                If probe.Exists("active") Then isProbeActive = CBool(probe("active"))
            End If

            If isProbeActive Then
                ' --- Path A: Probe is alive and reporting ---
                ctx("missingProbeCount") = 0
                
                ' Check for Stalled Renderer (Heartbeat verification).
                ' If the browser's timestamp is frozen despite pending requests, it's a desync.
                Dim jsNowVal As Double: jsNowVal = CDbl(probe("now"))
                If jsNowVal = m_lastJsNow And (socket_.PendingRequestCount > 0) Then
                    If DEBUG_MODE Then Debug.Print "BiDi: Renderer stall detected via heartbeat. Forcing probe recovery..."
                    Call RecoverSpaProbe(targetCtx)
                End If
                m_lastJsNow = jsNowVal
                
                ' --- CORE CONSENSUS ENGINE ---
                ' Evaluates Network Idle + DOM Stability.
                ' If internal panic occurs (e.g., infinite mutations), this may raise 505.
                EvaluateWaitCondition = EvaluateSpaIdleCondition(probe, ctx, socket_.NowMs())
                
                ' Explicitly release the local probe reference.
                Set probe = Nothing
            Else
                ' --- Path B: Probe is missing (Navigation, Crash, or Context Swap) ---

                ' --- [OPTIMIZATION] Predictive Recovery via navigationStarted ---
                ' If we already know a navigation occurred in THIS context,
                ' skip the expensive QuerySpaProbe round-trip on the next iteration.
                ' Instead, go directly to probe re-injection with a clean state.
                If m_navigationDetected And m_navigationContextId = targetCtx Then
                    ' Consume the flag (one-shot) to prevent infinite re-injection loops.
                    m_navigationDetected = False
                    m_navigationContextId = ""

                    ' Force context refresh since the document has changed.
                    On Error Resume Next
                    Call UpdateMainContextId(True, True)
                    On Error GoTo 0

                    ' Reset stability window: new document = unknown state.
                    ctx("stableStart") = socket_.NowMs()
                    ctx("missingProbeCount") = 0

                    ' Immediate re-injection without querying first.
                    If DEBUG_MODE Then
                        Debug.Print "BiDi: Predictive probe injection (navigationStarted shortcut)"
                    End If
                    Call StartSpaIdleProbe(targetCtx)

                    EvaluateWaitCondition = False
                    Exit Function
                End If

                Dim bCount As Long: bCount = CLng(ctx("missingProbeCount")) + 1
                ctx("missingProbeCount") = bCount
                
                ' Reset stability window during recovery as the logical state is unknown.
                ctx("stableStart") = socket_.NowMs()
                
                ' ERROR 505: Logical Panic (Consensus Failure)
                ' If the probe cannot be established after multiple attempts, the page is
                ' considered too unstable to reach a consensus.
                If bCount > MAX_RECOVERY_ATTEMPTS Then
                    Err.Raise 505, "BiDi:ConsensusError", _
                        "SPA Consensus Panic: Stability probe failed to initialize after " & _
                        MAX_RECOVERY_ATTEMPTS & " attempts. The page state is too volatile."
                End If
                
                ' JIT Re-injection: Attempt to heal the logical monitoring layer.
                If DEBUG_MODE Then Debug.Print "BiDi: SPA Probe inactive. Attempting re-injection (#" & bCount & ")..."
                Call StartSpaIdleProbe(targetCtx)
                
                EvaluateWaitCondition = False
            End If
    End Select
End Function

' [INTERNAL] ExecuteBaseAction
'  The core execution engine for browser interactions (Click, Input, etc.).
' UPGRADED:
'   - Optimized memory management by explicitly clearing object references.
'   - Refined retry logic to distinguish between transient protocol desync and fatal errors.
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
    Dim localErrNum As Long:    localErrNum = 0
    Dim localErrDesc As String: localErrDesc = ""
    Dim localErrSrc As String:  localErrSrc = ""
    Dim callResult As String:   callResult = ""
    Dim status As Long:         status = 0
    
    ' Resolve the target context and sanitize timing parameters.
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim vStable As Long:     vStable = SanitizeMs(minStable, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vTimeout As Long:    vTimeout = SanitizeMs(maxTimeout, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)

    Do While retryCount <= maxRetries
        ' --- PHASE 1: PREPARE (Probe injection & Context Refresh) ---
        ' Ensures the SPA monitoring probe is active before interaction.
        ' If this is a retry (retryCount > 0), it forces a context tree refresh.
        status = TryPreparePhase(targetCtx, waitNetwork, (retryCount > 0), _
                                 localErrNum, localErrDesc, localErrSrc)
        
        ' --- PHASE 2: EXECUTE (JS Action) ---
        ' Proceed to execution only if the preparation phase succeeded.
        If status = 0 Then
            status = TryExecutePhase(targetCtx, jsKey, args, awaitPromise, vTimeout, _
                                     callResult, localErrNum, localErrDesc, localErrSrc)
        End If
        
        ' --- PHASE 3: SUCCESS & SYNC ---
        ' If interaction succeeded, perform post-action synchronization.
        If status = 0 Then
            ' Post-action sync is best-effort (non-fatal if it times out).
            Call TrySyncPhase(targetCtx, waitNetwork, vStable, vTimeout, checkUrl)
            ExecuteBaseAction = callResult
            
            ' SUCCESS EXIT: Proceed to memory cleanup.
            GoTo Finalize
        End If

        ' --- RETRY DECISION LOGIC ---
        ' If status indicates a fatal failure (e.g., Element Not Found), abort retries.
        If status = ERR_ACTION_FAILED Then Exit Do
        
        ' For transient protocol errors (Context Lost, Socket issues), wait and retry.
        retryCount = retryCount + 1
        If retryCount <= maxRetries Then
            If DEBUG_MODE Then
                Debug.Print "BiDi: Action '" & jsKey & "' failed (Status:" & status & "). " & _
                            "Attempting retry " & retryCount & "/" & maxRetries & "..."
            End If
            ' Yield control to process background messages before retrying.
            WaitHeartbeat 500
        End If
    Loop

    ' --- ERROR PROPAGATION ---
    ' If the loop completes without success, raise the captured exception.
    If localErrNum = 0 Then localErrNum = ERR_ACTION_FAILED
    If localErrDesc = "" Then localErrDesc = "Unexpected execution state in ExecuteBaseAction."
    
    Err.Raise localErrNum, localErrSrc, localErrDesc

Finalize:
    ' --- MEMORY NOTE ---
End Function

' [INTERNAL] SyncAfterAction (Updated with URL Stability)
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

' [PRIVATE] ValidateProtocolResponse
'  Parses and validates BiDi protocol responses. Detects malformed JSON,
'          top-level BiDi protocol errors, and JavaScript runtime exceptions,
'          raising errors accordingly.
' RETURNS: ParseJson-produced Dictionary on success.
Private Function ValidateProtocolResponse(ByVal jsonRaw As String, ByVal methodName As String) As Object
    Dim json As Object

    ' --- Step 1: JSON Parsing Validation ---
    ' OERN scope: ParseJson only. If it throws, json remains Nothing (VBA default),
    ' which is caught by the Nothing check below. This is intentional.
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonRaw)
    On Error GoTo 0

    If json Is Nothing Then
        Err.Raise vbObjectError + 512, "BiDi:" & methodName, "Response is not valid JSON."
    End If

    ' --- Step 2: Layer 1 - Protocol Level Errors ---
    ' Abort if the browser returns a top-level BiDi protocol error
    If json.Exists("error") Then
        Dim errType As String: errType = json("error")
        Dim errMsg As String: errMsg = json("message")
        If Not IsIgnorableError(errType, errMsg) Then
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:" & methodName, "[" & errType & "] " & errMsg
        End If
    End If

    ' --- Step 3: Layer 2 - JavaScript Runtime Exceptions ---
    ' Use Layer 3 Exit Accessors for consistent object/string extraction
    Dim excDetails As Object
    Set excDetails = GetPathObject(json, "result.exceptionDetails")

    If Not excDetails Is Nothing Then
        Dim jsErrDetail As String
        jsErrDetail = GetPathString(json, "result.exceptionDetails.text")

        If DEBUG_MODE Then
            Debug.Print "--- BIDI JS ERROR REPORT | Method: " & methodName & " ---"
            Debug.Print FormatJavaScriptError(excDetails)
        End If

        Err.Raise ERR_JS_EXCEPTION, "BiDi:JSException:" & methodName, jsErrDetail
    End If

    Set ValidateProtocolResponse = json
End Function

' [PRIVATE] UnwrapSafetyShell
'  Unwraps the Safety Shell envelope from a validated BiDi response.
'          Falls back to the raw BiDi value when the response is not wrapped,
'          and enforces the script-result integrity contract.
' RETURNS: Dictionary in normalized {"status":"ok","value":...} form.
Private Function UnwrapSafetyShell(ByVal json As Object, ByVal methodName As String) As Object
    Dim resultObj As New Dictionary ' Container for normalized output

    ' --- Step 4: Layer 3 - Safety Shell Unwrapping ---
    ' [CORE UPDATE]: Routing extraction through the Gatekeeper (GetPathValue)
    Dim rawVal As Variant
    Assign rawVal, GetPathValue(json, "result.result.value")

    Dim isWrapped As Boolean: isWrapped = False

    ' Check for the Safety Shell envelope (JSON string starting with '{')
    If VarType(rawVal) = vbString Then
        Dim sVal As String: sVal = Trim(CStr(rawVal))
        If Left(sVal, 1) = "{" And Right(sVal, 1) = "}" Then
            Dim innerJson As Object
            On Error Resume Next
            Set innerJson = WebJsonConverter.ParseJson(sVal)
            On Error GoTo 0

            If Not innerJson Is Nothing Then
                If innerJson.Exists("status") Then
                    isWrapped = True
                    ' Handle browser-side watchdog/script errors reported via the shell
                    If innerJson("status") = "error" Then
                        Dim shellMsg As String: shellMsg = CStr(innerJson("message"))

                        ' Map the specific VBA_TIMEOUT token to a dedicated Error 517
                        If shellMsg = "VBA_TIMEOUT" Then
                            Err.Raise 517, "BiDi:SafetyTimeout:" & methodName, _
                                "Execution cancelled by browser-side Safety Shell (Watchdog triggered)."
                        Else
                            Err.Raise ERR_ACTION_FAILED, "BiDi:SafetyTimeout:" & methodName, shellMsg
                        End If
                    End If

                    ' Success: Assign the unwrapped payload
                    Call AssignValue(resultObj, "value", innerJson("value"))
                End If
            End If
        End If
    End If

    ' --- Step 5: Final Integrity Check (BiDi Protocol Compliance) ---
    ' Fallback to raw BiDi value if the response was not wrapped
    If Not isWrapped Then
        ' Perform success pattern matching for methods with no return value (e.g., script actions)
        If VarType(rawVal) = vbEmpty Then
            If InStr(1, methodName, "script.", vbTextCompare) > 0 Then
                Dim resultType As String
                resultType = GetPathString(json, "result.result.type")

                Select Case resultType
                    Case "undefined", "null"
                        ' Valid script termination with no return
                    Case ""
                        ' Check for structural success indicators (scriptId or realm presence)
                        Dim hasScriptId As String: hasScriptId = GetPathString(json, "result.script")
                        Dim hasRealms As String:   hasRealms = GetPathString(json, "result.realms.0.realm")

                        If hasScriptId = "" And hasRealms = "" Then
                            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:UnwrapSafetyShell:" & methodName, _
                                "Structural failure: No recognizable result or metadata in script response."
                        End If
                    Case Else
                        ' Protocol violation: a type was reported, but the value is missing
                        Err.Raise ERR_BIDI_PROTOCOL, "BiDi:UnwrapSafetyShell:" & methodName, _
                            "Protocol violation: Result type '" & resultType & "' reported without value."
                End Select
            End If
        End If

        ' Final normalized assignment
        Call AssignValue(resultObj, "value", rawVal)
    End If

    resultObj("status") = "ok"
    Set UnwrapSafetyShell = resultObj
End Function

' [INTERNAL] AssignValue
'  Safely assigns a value to a Dictionary key using Let or Set as appropriate.
Private Sub AssignValue(ByRef dic As Dictionary, ByVal key As String, ByRef val As Variant)
    If IsObject(val) Then
        Set dic(key) = val
    Else
        dic(key) = val
    End If
End Sub

' [INTERNAL] ResolveXPathOrThrow
'  A bottleneck function to resolve an XPath to a sharedId.
'          Features:
'          - Integrated validation for 'searchTimeoutMs' to prevent loop logic errors.
'          - Centralized error raising if the element is not found within the timeout.
'          - Ensures that the target context is fresh before the search.
Private Function ResolveXPathOrThrow(ByVal xpath As String, _
                                     ByVal contextId As String, _
                                     ByVal timeoutMs As Long, _
                                     ByVal methodName As String) As String
    Dim sid As String
    Dim targetCtx As String
    
    ' 0. [FIX] Null Byte Sanitization: Reject strings containing Chr$(0).
    If InStr(xpath, Chr$(0)) > 0 Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
                  "XPath contains null byte (Chr$(0)) which would corrupt the WebSocket message. Sanitize input before calling."
    End If
    ' 1. Consolidated Validation: Prevent logic errors from 'True' or negative inputs.
    Dim vTimeout As Long
    If timeoutMs = 0 Then
        vTimeout = 0
    Else
        vTimeout = SanitizeMs(timeoutMs, "searchTimeoutMs", 10000)
    End If
    ' 2. Resolve the active browsing context
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId()
    End If
    ' 3. Execution: Attempt to locate the element using the validated timeout.
    ' This calls the internal private locator which handles the polling loop.
    sid = GetFindElementSharedIdByXPath(xpath, vTimeout, targetCtx)
    
    ' 4. Error Handling: Raise a standardized exception if resolution fails.
    If sid = "" Then
        Err.Raise ERR_ACTION_FAILED, "BiDi:" & methodName, _
                  "Element not found within " & IIf(vTimeout = 0, "0 (immediate)", vTimeout) & "ms for XPath: " & xpath
    End If
    
    ResolveXPathOrThrow = sid
End Function

' [HELPER] PrepareProbeForAction
' Summary: Ensures the SPA monitoring probe is installed and active.
'          Called JIT (Just-In-Time) before any interaction to prevent count leaks.
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

' [INTERNAL] SanitizeMs
'  Central gatekeeper to validate millisecond arguments.
'          Strictly rejects negative values (preventing True/-1 errors)
'          and caps values to prevent permanent Excel hangs.
Private Function SanitizeMs(ByVal ms As Long, ByVal argName As String, _
                             ByVal defaultMs As Long, _
                             Optional ByVal minMs As Long = 50) As Long
    Dim vMs As Long: vMs = ms
    
    ' 1. Logic Guard: Rejects zero or negative values.
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

    SanitizeMs = vMs
End Function

' [PUBLIC] WaitForUrlStability
'  Polls the current URL and Document.readyState until they remain constant
'              for a specified 'stableDurationMs'.
' STRATEGY   : Uses a "Fail-Fast" approach. If the environment fails to stabilize
'              within the timeout, the loop exits.
Public Sub WaitForUrlStability(Optional ByVal contextId As String = "", _
                               Optional ByVal stableDurationMs As Long = 500, _
                               Optional ByVal timeoutMs As Long = 10000)

    ' --- [ENTRY GUARD] Guard against 0ms inputs ---
    If stableDurationMs <= 0 Then stableDurationMs = 500
    If timeoutMs <= 0 Then timeoutMs = 10000

    Dim vStableMs As Long: vStableMs = SanitizeMs(stableDurationMs, "stableDurationMs", 500)
    Dim vTimeoutMs As Long: vTimeoutMs = SanitizeMs(timeoutMs, "timeoutMs", 10000)
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim startMs As Double: startMs = socket_.NowMs()
    Dim stableStartMs As Double: stableStartMs = socket_.NowMs()
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
            stableStartMs = socket_.NowMs()
        Else
            ' Check if the state has been stable for the required duration
            If (socket_.NowMs() - stableStartMs) >= CDbl(vStableMs) Then Exit Do
        End If
        
        ' Global timeout check to prevent infinite loops
        If (socket_.NowMs() - startMs) > CDbl(vTimeoutMs) Then Exit Do
        
        ' [FIX] Adaptive yielding:
        ' If recording is active or network requests are pending, maintain queue sync.
        ' Otherwise, yield via Communicator for high-speed URL change detection.
        If p_isRecording Or (socket_.PendingRequestCount > 0) Then
            WaitHeartbeat 50
        Else
            socket_.YieldExecution 50, True
        End If
    Loop
End Sub

' [PRIVATE] WrapInSafetyShell (REFACTORED)
'  Wraps JS code in a Promise.race to enforce a browser-side timeout.
'              This is now the SOLE point of JSON stringification in the JS layer,
'              producing a clean, single-layer response for VBA.
Private Function WrapInSafetyShell(ByVal jsBody As String, ByVal timeoutMs As Long) As String
    Dim js As String
    
    js = "async function(){ "
    js = js & "const _t=" & timeoutMs & "; let _la=Date.now(); const hb=()=>{_la=Date.now()}; "
    
    ' --- _tk: The actual task execution ---
    ' Preservation of 'arguments' and lexical closure 'hb' is key here.
    ' [CRITICAL CONTRACT] _tk MUST remain an arrow function (=>).
    ' Arrow functions do NOT create their own 'arguments' binding, so
    ' arguments[0], arguments[1] etc. inside jsBody correctly resolve to
    ' the outer async function's parameters (the BiDi-serialized args).
    ' Changing _tk to a regular "async function()" will SILENTLY BREAK
    ' every action method (click, input, select, shadow_click) because
    ' arguments would become an empty Arguments object.
    js = js & "const _tk = async () => { " & jsBody & " }; "
    
    ' --- _wd: The Autonomous Watchdog ---
    ' Heartbeat check: If silence > timeout, reject with the specific "VBA_TIMEOUT" token.
    js = js & "let _wdTimer; const _wd = new Promise((_,r)=>{ _wdTimer=setInterval(()=>{if(Date.now()-_la>_t){clearInterval(_wdTimer);r(new Error('VBA_TIMEOUT'))}},100);}); "
    
    ' Race between the task and the internal watchdog.
    js = js & "try{ const res = await Promise.race([_tk(), _wd]); clearInterval(_wdTimer); return JSON.stringify({status:'ok', value:(typeof res === 'undefined' ? null : res)}); "
    js = js & "}catch(e){ clearInterval(_wdTimer); return JSON.stringify({status:'error', message:e.message}); } }"
    
    WrapInSafetyShell = js
End Function

' [PRIVATE] HandleConsoleLog
'  Formats and prints browser console logs to the VBA Immediate Window.
Private Sub HandleConsoleLog(ByVal json As Object)
    ' [FIX v2] Removed blanket OERN. GetPathString already returns "" for missing paths
    ' without raising errors, so the OERN wrapper was unnecessary and could mask
    ' unrelated bugs (e.g., a corrupted json object reference).
    Dim text As String: text = GetPathString(json, "params.text")
    
    ' Only output logs from our specific AutoClicker to keep the window clean.
    If InStr(text, "BiDi-AutoClicker") > 0 Then
        Debug.Print ">>> [BROWSER] " & text
    End If
End Sub

' METHOD: Shutdown
'  Explicitly releases all BiDi resources.
'          This should be called before setting the object to Nothing.
Public Sub Shutdown(Optional ByVal forceClose As Boolean = True)
    If socket_ Is Nothing Then Exit Sub
    On Error Resume Next
    
    ' --- 1. Recorder Memory Cleanup ---
    Set p_recorderLogs = Nothing
    
    ' --- 2. Internal State Reset ---
    ClearState
    
    ' --- 3. Communication Socket Release ---
    If forceClose Then
        socket_.forceClose
    End If
    Set socket_ = Nothing
    
    If DEBUG_MODE Then Debug.Print "BiDi: Explicit Shutdown sequence completed."
End Sub

' [PRIVATE] UpdateWaitTelemetry (Atomic HUD Upsert - State Only)
'  Updates the browser-side HUD with real-time sync telemetry.
' REFINED:
'   - Efficiency: Removed dynamic idle counter to allow lastHudKey to block redundant traffic.
'   - ID Integrity: Remains within the CallScript framework to ensure no ID collisions.
'   - Stability: Dramatically reduces WebSocket traffic by only updating on state changes.
Private Sub UpdateWaitTelemetry(ByVal mode As waitMode, ByRef ctx As Dictionary)
    On Error Resume Next
    
    Dim Msg As String, colorCode As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(CStr(ctx("contextId")))
    
    ' --- Static State Management ---
    Static lastHudKey As String
    Static lastHudCtx As String
    Static lastRetryTick As Currency
    
    ' --- 1. Construct Message & Color ---
    Select Case mode
        Case WaitMode_UrlContains
            Msg = "BiDi: Waiting for URL..." & vbLf & "Pattern: " & ctx("partialUrl")
            colorCode = "#00FF00" ' Green
            
        Case WaitMode_SpaIdle
            Dim netCount As Long: netCount = socket_.PendingRequestCount
            Dim idleTime As Double: idleTime = (socket_.NowMs() - CDbl(ctx("stableStart"))) / 1000
            Dim targetIdle As Double: targetIdle = CLng(ctx("minStableMs")) / 1000
            
            ' Determine HUD color and text based on STATE, not COUNTER.
            If netCount > 0 Or idleTime < targetIdle Then
                colorCode = "#FF4444" ' Red (Busy)
                Msg = "BiDi: [SPA SYNC]" & vbLf & "Status: Synchronizing..."
            Else
                colorCode = "#00FF00" ' Green (Stable)
                Msg = "BiDi: [SPA SYNC]" & vbLf & "Status: Stable"
            End If
    End Select
    
    ' Create a fingerprint to avoid redundant WebSocket calls
    ' Now that the idle counter is gone, this key will remain identical
    ' throughout the entire "Busy" phase, effectively blocking all repetitive calls.
    Dim currentHudKey As String: currentHudKey = Msg & "|" & colorCode
    
    ' --- 2. Context Change Detection ---
    If targetCtx <> lastHudCtx Then
        lastHudKey = ""
        lastHudCtx = targetCtx
        lastRetryTick = 0
    End If
    
    ' --- 3. Optimization: Skip if state is identical ---
    If currentHudKey = lastHudKey Then Exit Sub
    
    ' --- 4. ERROR THROTTLING ---
    If lastRetryTick > 0 Then
        If socket_.ElapsedMs(lastRetryTick) < 500 Then Exit Sub
    End If
    
    ' --- 5. JavaScript Construction ---
    Dim safeMsg As String: safeMsg = Replace(Replace(Msg, "\", "\\"), "'", "\'")
    safeMsg = Replace(Replace(safeMsg, vbCr, ""), vbLf, "\n")
    Dim safeColor As String: safeColor = Replace(colorCode, "'", "\'")
    Dim js As String: js = ""
    js = js & "(function(){ var b=document.body; if(!b) return; var MSG='" & safeMsg & "', CLR='" & safeColor & "';"
    js = js & " var h=document.getElementById('vba-bidi-hud'); if(!h){ h=document.createElement('div'); h.id='vba-bidi-hud';"
    js = js & " h.style.cssText='position:fixed;top:15px;right:15px;z-index:2147483647;padding:10px;font-family:monospace;font-size:12px;background:rgba(0,0,0,0.85);color:#00FF00;border:1px solid #00FF00;border-radius:6px;pointer-events:none;box-shadow:0 4px 15px rgba(0,0,0,0.5);min-width:180px;white-space:pre-wrap;'; b.appendChild(h);}"
    js = js & " h.innerText=MSG; h.style.borderColor=CLR; h.style.color=CLR; return true;})()"
    
    ' --- 6. Execution via CallScript ---
    Call Me.CallScript(js, Nothing, targetCtx, False, False, 1000, True)
    
    ' --- 7. Final State Update ---
    If Err.Number = 0 Then
        lastHudKey = currentHudKey
        lastRetryTick = 0
    Else
        lastRetryTick = socket_.NowTicks()
    End If
    
    On Error GoTo 0
End Sub

' [PRIVATE] GetHudJs
'  Returns JavaScript snippets for the Browser-side HUD (Head-Up Display).
Private Function GetHudJs(ByVal snippetType As String) As String
    Dim js As String: js = ""
    Select Case LCase(snippetType)
        Case "init"
            ' Injects a floating HUD. 'pointer-events:none' ensures it won't block clicks.
            js = js & "(function(){if(document.getElementById('vba-bidi-hud')) return;"
            js = js & "const h=document.createElement('div'); h.id='vba-bidi-hud';"
            js = js & "h.style.cssText='position:fixed;top:15px;right:15px;z-index:2147483647;padding:10px;font-family:monospace;font-size:12px;background:rgba(0,0,0,0.85);color:#00FF00;border:1px solid #00FF00;border-radius:6px;pointer-events:none;box-shadow:0 4px 15px rgba(0,0,0,0.5);min-width:180px;white-space:pre-wrap;';"
            js = js & "h.innerText='BiDi: Initializing...'; document.body.appendChild(h);})();"
            GetHudJs = js
                       
        Case "remove"
            ' Cleanly removes the HUD from the webpage DOM
            GetHudJs = "(function(){const h=document.getElementById('vba-bidi-hud'); if(h)h.remove();})();"
    End Select
End Function

' [PUBLIC] ExecuteEnableGlobalShadowUnlocker
'  Uses WebDriver BiDi's preloadScript to globally force Shadow DOM (closed)
'              to 'open' mode across all domains and iframes.
' RETURNS    : The registered script ID (script.PreloadScriptId)
Public Function ExecuteEnableGlobalShadowUnlocker() As String
    Dim params As New Dictionary, js As String
    
    ' Hook attachShadow to always execute with mode: 'open'
    js = "function() { "
    js = js & "if (Element.prototype._attachShadow) return; "
    js = js & "Element.prototype._attachShadow = Element.prototype.attachShadow; "
    js = js & "Element.prototype.attachShadow = function(options) { "
    js = js & "  return this._attachShadow({ mode: 'open' }); "
    js = js & "}; }"
    
    params.Add "functionDeclaration", js

    ' Execute command silently. GetPathString handles empty/error results gracefully.
    Dim strRes As String: strRes = ExecuteBiDiCommand("script.addPreloadScript", params, True)
    ExecuteEnableGlobalShadowUnlocker = GetPathString(WebJsonConverter.ParseJson(strRes), "result.script")
    
    If DEBUG_MODE And ExecuteEnableGlobalShadowUnlocker <> "" Then
        Debug.Print "BiDi: Global Shadow DOM Unlocker enabled (ID: " & ExecuteEnableGlobalShadowUnlocker & ")"
    End If
End Function

' [Phase 1] FUNCTION: GetNextRequestId
'  Delegated to Communicator. This stub is retained for backward compatibility
'          with any external callers (e.g., PostScript's manual payload construction).
'          Internal callers (ExecuteBiDiCommand) now use socket_.SendCommand directly.
Private Function GetNextRequestId() As Long
    GetNextRequestId = socket_.GetNextRequestId()
End Function

' [PHASE] TryPreparePhase
'  Prepares the environment before an interaction.
'              Ensures the SPA probe is active and refreshes the context tree if needed.
Private Function TryPreparePhase(ByVal ctx As String, _
                                  ByVal waitNet As Boolean, _
                                  ByVal isRetry As Boolean, _
                                  ByRef outErrNum As Long, _
                                  ByRef outErrDesc As String, _
                                  ByRef outErrSrc As String) As Long
    Dim capturedErr As Long
    
    ' Step 1: Ensure SPA monitoring probe is installed if synchronization is required.
    ' JIT injection prevents state loss between page transitions.
    If waitNet Then
        On Error Resume Next
        Call PrepareProbeForAction(ctx)
        capturedErr = Err.Number
        If capturedErr <> 0 Then
            outErrNum = capturedErr: outErrDesc = Err.description: outErrSrc = Err.Source
            On Error GoTo 0
            TryPreparePhase = capturedErr: Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' Step 2: Refresh main browsing context mapping on retry attempts.
    ' This handles cases where the context was destroyed or navigated during a race condition.
    If isRetry Then
        On Error Resume Next
        Call UpdateMainContextId(True, True) ' forceRefresh:=True, Silent:=True
        capturedErr = Err.Number
        If capturedErr <> 0 Then
            outErrNum = capturedErr: outErrDesc = Err.description: outErrSrc = Err.Source
            On Error GoTo 0
            TryPreparePhase = capturedErr: Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' Return 0 indicating a successful preparation phase.
    TryPreparePhase = 0
End Function

' [PHASE] TryExecutePhase
'  Dispatches the actual JavaScript command to the browser.
'              The JS is automatically wrapped in a Safety Shell (Watchdog).
Private Function TryExecutePhase(ByVal ctx As String, _
                                  ByVal key As String, _
                                  ByVal args As Collection, _
                                  ByVal await As Boolean, _
                                  ByVal timeout As Long, _
                                  ByRef outResult As String, _
                                  ByRef outErrNum As Long, _
                                  ByRef outErrDesc As String, _
                                  ByRef outErrSrc As String) As Long
    On Error Resume Next
    
    ' Execute the script via CallScript which manages realm targeting and safety shells.
    ' The result is stored in the ByRef string variable for the caller.
    outResult = CallScript(GetActionJs(key), args, ctx, await, , timeout)
    
    ' Capture any error occurred during transmission or script execution.
    TryExecutePhase = Err.Number
    If Err.Number <> 0 Then
        outErrNum = Err.Number
        outErrDesc = Err.description
        outErrSrc = Err.Source
    End If
    On Error GoTo 0
End Function

' [PHASE] TrySyncPhase
'  Performs post-action synchronization (Network Idle / DOM stability).
'              Transient failures (normal timeouts) are logged as warnings but do not
'              invalidate the action. However, ERR_CONSENSUS_HALT is a deliberate stop
'              signal that MUST propagate to halt the entire process for review.
Private Sub TrySyncPhase(ByVal ctx As String, _
                         ByVal waitNet As Boolean, _
                         ByVal stable As Long, _
                         ByVal timeout As Long, _
                         ByVal chkUrl As Boolean)
    On Error Resume Next
    
    ' Perform SPA-aware synchronization if requested.
    If waitNet Then
        Call SyncAfterAction(ctx, stable, timeout, chkUrl)
    End If
    
    ' Capture error state before clearing
    Dim syncErrNum As Long:    syncErrNum = Err.Number
    Dim syncErrDesc As String: syncErrDesc = Err.description
    Dim syncErrSrc As String:  syncErrSrc = Err.Source
    On Error GoTo 0
    
    ' --- CRITICAL: Do NOT absorb CONSENSUS_HALT ---
    ' ERR_CONSENSUS_HALT means the failure was logged and the operator needs to
    ' review the log before continuing. Re-raise to stop the entire process.
    If syncErrNum = ERR_CONSENSUS_HALT Then
        Err.Raise syncErrNum, syncErrSrc, syncErrDesc
    End If
    
    ' Diagnostic Logging: Transient sync failures (504 timeout etc.) are non-fatal.
    ' The primary interaction (Click/Input) has already succeeded.
    If syncErrNum <> 0 And DEBUG_MODE Then
        Debug.Print "[" & Format(Now, "hh:mm:ss") & "] BiDi Warning: Post-action Sync failed (Status: " & syncErrNum & "). DOM mutations only (no pending requests). Consider extending idle timeout."
    End If
End Sub

' [PUBLIC] PostScript (Fire & Forget Mode)
'  Executes JavaScript without waiting for a response from the browser.
'              Primarily used for HUD updates to prevent UI freezing (whiteout)
'              caused by synchronous socket congestion during heavy event traffic.
Public Sub PostScript(ByVal js As String, _
                      Optional ByVal args As Collection = Nothing, _
                      Optional ByVal contextId As String = "")
    
    Dim params As New Dictionary, targetDic As New Dictionary
    Dim resolvedCtx As String: resolvedCtx = ResolveTargetContext(contextId)
    
    If resolvedCtx = "" Then Exit Sub

    ' 1. Wrap the script in the Safety Shell for browser-side protection.
    '    Timeout is set to a short duration (1000ms) as this is non-critical.
    Dim finalJs As String: finalJs = WrapInSafetyShell(js, 1000)

    ' 2. Prepare the command payload.
    targetDic.Add "context", resolvedCtx
    params.Add "functionDeclaration", finalJs
    params.Add "arguments", IIf(args Is Nothing, New Collection, args)
    params.Add "target", targetDic
    params.Add "awaitPromise", True

    ' 3. Construct the full BiDi Request.
    '    A fixed sentinel ID (POST_SCRIPT_ID = -1) helps identify these responses
    '    during background queue processing. GetNextRequestId issues only positive
    '    values (1..2,000,000,000), so -1 cannot collide with normal request IDs.
    Dim requestDic As New Dictionary
    requestDic.Add "id", POST_SCRIPT_ID
    requestDic.Add "method", "script.callFunction"
    requestDic.Add "params", params
    
    ' 4. [CRITICAL] Direct Send without Receive via Communicator's PostScript.
    '    By calling socket_.PostScript instead of ExecuteBiDiCommand,
    '    we bypass the synchronous SendAndReceive loop, ensuring the VBA
    '    main thread remains responsive to Windows messages.
    socket_.PostScript WebJsonConverter.ConvertToJson(requestDic)
End Sub

' LAYER 0: XPath String Escaper
' Summary  : Produces a safe XPath string literal from arbitrary VBA text.
'            If the text contains no single quotes, wraps in 'text'.
'            If it contains no double quotes, wraps in "text".
'            Otherwise, splits at single-quote boundaries and uses concat().
' Example  : "O'Brien"  -> concat('O',&quot;'&quot;,'Brien')
Private Function EscapeXPathString(ByVal text As String) As String
    If InStr(text, "'") = 0 Then
        ' Simple case: no single quotes, wrap in single quotes
        EscapeXPathString = "'" & text & "'"
    ElseIf InStr(text, """") = 0 Then
        ' No double quotes, wrap in double quotes
        EscapeXPathString = """" & text & """"
    Else
        ' Contains both: split on single quotes and reassemble with concat()
        Dim parts() As String: parts = Split(text, "'")
        Dim result As String: result = "concat("
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            If i > LBound(parts) Then result = result & ",""'"","
            result = result & "'" & parts(i) & "'"
        Next i
        result = result & ")"
        EscapeXPathString = result
    End If
End Function

' LAYER 0: Safe Assignment Helper (The "Shield")
' Summary  : Handles the Set/Let logic internally to prevent assignment errors.
' Warning  : Unsafe for function return variables inside deep On Error scopes.
Private Sub Assign(ByRef dest As Variant, ByRef src As Variant)
    If IsObject(src) Then Set dest = src Else dest = src
End Sub

' LAYER 1: Core Navigation Engine (Physical Search)
' Traverses a nested Dictionary/Collection tree using dot-separated path notation.
' Supports Dictionary key lookup and Collection index access (0-based in path, 1-based in VBA).
' Example: "result.contexts.0.url" navigates result -> contexts -> first item -> url.
' Returns Empty (default Variant) if any segment is missing or invalid.
Private Function InternalTraverseFromObject(ByVal root As Object, ByVal Path As String) As Variant
    If root Is Nothing Or Path = "" Then Exit Function
    
    Dim keys() As String: keys = Split(Path, ".")
    Dim current As Object: Set current = root
    Dim i As Long, idx As Long
    
    For i = 0 To UBound(keys) - 1
        If TypeName(current) = "Dictionary" Then
            If Not current.Exists(keys(i)) Then Exit Function
            If IsObject(current(keys(i))) Then Set current = current(keys(i)) Else Exit Function
        ElseIf TypeName(current) = "Collection" And IsNumeric(keys(i)) Then
            idx = CLng(keys(i)) + 1
            If idx < 1 Or idx > current.count Then Exit Function
            If IsObject(current(idx)) Then Set current = current(idx) Else Exit Function
        Else: Exit Function
        End If
        If current Is Nothing Then Exit Function
    Next i
    
    Dim lastKey As String: lastKey = keys(UBound(keys))
    If TypeName(current) = "Dictionary" Then
        If current.Exists(lastKey) Then
            ' CRITICAL: Use explicit Set/Let here to guarantee return value stability
            If IsObject(current(lastKey)) Then Set InternalTraverseFromObject = current(lastKey) _
            Else InternalTraverseFromObject = current(lastKey)
        End If
    ElseIf TypeName(current) = "Collection" And IsNumeric(lastKey) Then
        idx = CLng(lastKey) + 1
        If idx >= 1 And idx <= current.count Then
            If IsObject(current(idx)) Then Set InternalTraverseFromObject = current(idx) _
            Else InternalTraverseFromObject = current(idx)
        End If
    End If
End Function

' LAYER 2: Gatekeeper (Type Arbitration)
Private Function GetPathValue(ByVal root As Object, ByVal Path As String) As Variant
    Assign GetPathValue, InternalTraverseFromObject(root, Path)
End Function

' LAYER 3: Specialized Exit Accessors
Private Function GetPathString(ByVal root As Object, ByVal Path As String) As String
    Dim v As Variant: Assign v, GetPathValue(root, Path)
    If Not IsObject(v) Then GetPathString = CStr(v & "")
End Function

Private Function GetPathObject(ByVal root As Object, ByVal Path As String) As Object
    ' [FIX v2] Explicit type check replaces blanket OERN.
    ' GetPathValue returns Variant; we must verify it's an Object before Set.
    Dim v As Variant: Assign v, GetPathValue(root, Path)
    If IsObject(v) Then Set GetPathObject = v
    ' Returns Nothing implicitly if v is not an Object
End Function

Private Function GetCleanResult(ByVal jsonRaw As String) As Variant
    ' [FIX v2] Narrowed OERN scope to ParseJson only. ParseJson can legitimately fail
    ' on malformed/empty input, but errors in GetPathValue or Assign should propagate
    ' since they indicate structural bugs in the response pipeline.
    If Len(jsonRaw) = 0 Then GetCleanResult = Empty: Exit Function
    
    Dim parsed As Object
    On Error Resume Next
    Set parsed = WebJsonConverter.ParseJson(jsonRaw)
    Dim parseErr As Long: parseErr = Err.Number
    On Error GoTo 0
    
    If parseErr <> 0 Or parsed Is Nothing Then GetCleanResult = Empty: Exit Function
    
    Dim v As Variant: Assign v, GetPathValue(parsed, "value")
    If IsObject(v) Then
        If v Is Nothing Then GetCleanResult = Empty Else Set GetCleanResult = v
    Else: GetCleanResult = v
    End If
End Function

' [PUBLIC] ExecuteWizSelectDropdown
'  Controls semantic pseudo dropdowns (combobox/option UI architecture)
'              by leveraging the multi-phase SPA consensus click engine.
Public Sub ExecuteWizSelectDropdown(ByVal triggerXPath As String, _
                                    ByVal optionText As String, _
                                    Optional ByVal searchTimeoutMs As Long = 10000, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "")
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' Phase 1: Click the combo box trigger to dynamic spawn the option list window.
    ' SyncAfterAction will automatically lock until the pop-up mutations completely settle.
    Call Me.ExecuteClickByXPath(triggerXPath, searchTimeoutMs, True, , maxTimeoutMs, targetCtx)
    
    ' Construct W3C ARIA-compliant semantic XPath to isolate the designated list option.
    Dim optionXPath As String
    optionXPath = "//*[@role='option'][contains(., " & EscapeXPathString(optionText) & ")]"
    
    ' Phase 2: Click the resolved option inside the dynamically populated layer.
    ' SyncAfterAction will automatically trap the frame re-render and wait until the SPA returns to stable state.
    Call Me.ExecuteClickByXPath(optionXPath, searchTimeoutMs, True, , maxTimeoutMs, targetCtx)
End Sub
