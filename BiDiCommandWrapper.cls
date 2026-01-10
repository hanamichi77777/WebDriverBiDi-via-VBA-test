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
Private Const MAX_CONTEXT_RECOVERY_RETRIES As Long = 2

' [Configuration Constants]
Private Const DEFAULT_IDLE_STABLE_MS As Long = 500
Private Const DEFAULT_IDLE_TIMEOUT_MS As Long = 10000
Private Const DEFAULT_STALE_RETRIES As Long = 2
Private Const SHORT_WAIT_MS As Long = 500
Private Const MAX_LOG_SIZE As Long = 500

' --- Noise Filtering State ---
' Default filters are built-in to handle common noise (Analytics, Hidden inputs, etc.)
Private Const DEFAULT_IGNORE_SELECTORS As String = "script, link, input[type='hidden'], .vba-ignore"
Private Const DEFAULT_IGNORE_NETWORK As String = "google-analytics, doubleclick, /collect, /beacon, .woff2, favicon.ico"

' User-defined filters set from Standard Modules
Private p_userIgnoreSelectors As Collection
Private p_userIgnoreNetworkPatterns As Collection

' --- State flags for persistent registrations ---
Private m_isWaiting As Boolean       ' Guard against re-entrant calls
Private m_lastJsNow As Double        ' Heartbeat tracker (performance.now from JS)
Private m_cachedRealmId As String    ' Cached Execution Realm ID

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
' Initialization: Setup WebSocket and validate High-Resolution Timer frequency.
' ========================================================================================
Private Sub Class_Initialize()
    Set socket_ = New BiDiSocketCommunicator
    p_nextRequestId = 1
    
    ' Get frequency once. Modern systems always support this.
    QueryPerformanceFrequency p_timerFreq
End Sub

Public Sub ConnectTo(ByVal WebSocketUrl As String)
    socket_.AttemptAutoConnect WebSocketUrl
End Sub

' ========================================================================================
' Class_Terminate: Resource cleanup including explicit CDP session detachment
' ========================================================================================
Private Sub Class_Terminate()
    ' --- 1. CDP Session Detachment ---
    ' Explicitly detach the CDP session to prevent memory leaks in the browser process.
    ' If the session is not detached, the browser maintains internal debug listeners.
    If p_cdpSessionId <> "" And Not socket_ Is Nothing Then
        On Error Resume Next
        Dim params As New Dictionary
        params.Add "session", p_cdpSessionId
        
        ' Use ExecuteBiDiCommand to send the detach command before closing the socket.
        ' We do not check for success here as the class is already terminating.
        ExecuteBiDiCommand "goog:cdp.detachSession", params
        On Error GoTo 0
        
        p_cdpSessionId = ""
    End If

    ' --- 2. Recorder Memory Cleanup ---
    If Not p_recorderLogs Is Nothing Then
        On Error Resume Next
        Dim i As Long
        For i = p_recorderLogs.Count To 1 Step -1
            p_recorderLogs.Remove i
        Next i
        Set p_recorderLogs = Nothing
        On Error GoTo 0
    End If
    
    ' --- 3. Communication Socket Release ---
    Set socket_ = Nothing
    
    If DEBUG_MODE Then Debug.Print "BiDiCommandWrapper: Terminated. CDP session and socket resources released."
End Sub

Public Property Set SetSocket(socket As BiDiSocketCommunicator)
    Set socket_ = socket
End Property

' ========================================================================================
' Central Event Dispatcher (Modified with Smart Filtering)
' ========================================================================================
Private Sub DispatchGlobalEvents(ByVal json As Object)
    If Not json.Exists("method") Then Exit Sub
    Dim method As String
    method = json("method")

    ProcessCDPBlockingEvent json

    Select Case method
        Case "network.beforeRequestSent"
            Dim url As String
            url = ""
            If json.Exists("params") Then
                If json("params").Exists("request") Then
                    url = json("params")("request")("url")
                End If
            End If
            
            If Not IsNoiseUrl(url) Then
                p_pendingRequestCount = p_pendingRequestCount + 1
            End If

        Case "network.responseCompleted", "network.fetchError"
            Dim resUrl As String
            resUrl = ""
            If json.Exists("params") Then
                If json("params").Exists("response") Then
                    resUrl = json("params")("response")("url")
                End If
            End If

            p_pendingRequestCount = p_pendingRequestCount - 1
            If p_pendingRequestCount < 0 Then p_pendingRequestCount = 0
    End Select

    If p_isRecording Then
        FormatAndStoreLog json
    End If
End Sub

Public Property Get PendingRequestCount() As Long
    PendingRequestCount = p_pendingRequestCount
End Property

' ========================================================================================
' [PUBLIC] GetIframeContextIdByUrl
' DESCRIPTION: Searches for a frame context by partial URL with validated timeout.
' ========================================================================================
Public Function GetIframeContextIdByUrl(ByVal partialUrl As String, Optional ByVal timeoutMs As Long = 5000) As String
    ' --- Input Validation ---
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 5000)
    
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim ctxId As String

    Do
        ' Fetch the tree and search recursively
        Dim strRes As String: strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Dim contexts As Collection: Set contexts = GetJsonValue(strRes, "result.contexts")
        
        If Not contexts Is Nothing Then
            ctxId = RecursiveFrameSearch(contexts, partialUrl)
            If ctxId <> "" Then
                GetIframeContextIdByUrl = ctxId
                Exit Function
            End If
        End If

        ' Immediate exit if wait is not required
        If vTimeout <= 0 Then Exit Do

        ' Process events and check for timeout using high-precision timer
        WaitHeartbeat SHORT_WAIT_MS
        If IsTimeout(startTick, CDbl(vTimeout)) Then Exit Do
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
                                           Optional ByVal timeoutMs As Long = 5000, _
                                           Optional ByVal waitForComplete As Boolean = True) As String
    ' --- Input Validation ---
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 5000)
    
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim strRes As String, res As Variant
    Dim jsonTree As Object, ctxList As Collection, ctx As Dictionary
    
    Do
        ' 1. Retrieve the latest browsing context tree
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
                        ' Check document.readyState
                        Dim jsRes As String
                        jsRes = CallScript(GetActionJs("ready_state"), , ctx("context"))
                        res = GetPathValue(WebJsonConverter.ParseJson(jsRes), "result.result.value")
                        If CStr(res) = "complete" Then
                            ExecuteFindWindowContextId = ctx("context"): Exit Function
                        End If
                    End If
                End If
            Next ctx
        End If

        ' Yield and check for timeout
        WaitHeartbeat SHORT_WAIT_MS
        
        ' standardized error if target not found within vTimeout
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
Private Function ExecuteGetTitleByContextId(ByVal contextId As String) As String
    Dim js As String: js = "function() { return document.title || ''; }"
    Dim strRes As String
    strRes = CallScript(js, , contextId)
    
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    ExecuteGetTitleByContextId = GetPathString(json, "result.result.value")
End Function

' ========================================================================================
' [PUBLIC] ExecuteClickAndHandleAlertByXPath
' ========================================================================================
Public Function ExecuteClickAndHandleAlertByXPath(ByVal xpath As String, _
                                                  Optional ByVal Accept As Boolean = True, _
                                                  Optional ByVal waitNetworkIdle As Boolean = True, _
                                                  Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                                  Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                                  Optional ByVal contextId As String = "") As String
    
    ' 1. Resolve context
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 2. Resolve and Validate XPath (This replaces GetFindElementSharedIdByXPath and the IF check)
    ' searchTimeoutMs はこのメソッドの引数にないため、暫定的に 2000ms (固定値) を渡すか、
    ' 必要であればこのメソッドの引数にも searchTimeoutMs を追加してください。
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, 2000, "ExecuteClickAndHandleAlertByXPath")

    ' 3. Prepare Arguments
    Dim args As New Collection, arg1 As New Dictionary
    arg1.Add "sharedId", sharedId: args.Add arg1

    ' 4. Execute asynchronous click
    CallScript GetActionJs("click_alert"), args, targetCtx, False
    
    ' 5. Wait for UI
    WaitHeartbeat 500
    
    ' 6. Handle the alert
    ExecuteClickAndHandleAlertByXPath = ExecuteHandleUserPrompt(Accept, "", targetCtx)
    
    ' 7. Sync state (waitNetworkIdle logic)
    ' minStableMs と maxTimeoutMs のバリデーションは SyncAfterAction 内の
    ' WaitForNetworkIdleSpa で一括処理されるため、ここではそのまま渡してOKです。
    If waitNetworkIdle Then SyncAfterAction targetCtx, minStableMs, maxTimeoutMs
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
' [PUBLIC] ExecuteBiDiWaitUntilUrlContains
' DESCRIPTION: Polls the URL until it matches the partial string.
' ========================================================================================
Public Function ExecuteBiDiWaitUntilUrlContains(ByVal partialUrl As String, _
                                                Optional ByVal timeoutSec As Long = 300, _
                                                Optional ByVal contextId As String = "") As Boolean
    ' --- Input Validation ---
    ' Convert seconds to milliseconds and sanitize to prevent negative values (True/-1)
    Dim vTimeoutMs As Long
    vTimeoutMs = SanitizeMs(timeoutSec * 1000, "timeoutSec", 300000)
    
    ' Subscribe to relevant navigation events
    Me.ExecuteSessionSubscribe Array("network.responseCompleted", "browsingContext.load", "browsingContext.fragmentNavigated")
    
    Dim waitCtx As New Dictionary
    waitCtx.Add "contextId", IIf(contextId <> "", contextId, UpdateMainContextId())
    waitCtx.Add "partialUrl", partialUrl
    
    ' Execute the engine with the validated millisecond timeout
    ExecuteBiDiWaitUntilUrlContains = MasterWaitEngine(WaitMode_UrlContains, vTimeoutMs, 100, waitCtx)
End Function

' ========================================================================================
' FUNCTION: ExecuteNavigateAndGetStatus
' SUMMARY: High-level navigation that returns a status indicator (200/500).
' ========================================================================================
Public Function ExecuteNavigateAndGetStatus(ByVal targetUrl As String, _
                                            Optional ByVal waitNetworkIdle As Boolean = True, _
                                            Optional ByVal contextId As String = "") As Long
    On Error Resume Next
    
    ' Trigger navigation
    ' Note: If waitNetworkIdle is False, we only wait for 'complete' readyState.
    Dim waitMode As String
    If waitNetworkIdle Then waitMode = "complete" Else waitMode = "none"
    
    ExecuteBrowsingContextNavigate targetUrl, waitMode, contextId
    
    If Err.Number = 0 Then
        ExecuteNavigateAndGetStatus = 200 ' Success
    Else
        Debug.Print "Navigation Failed: " & Err.Description
        ExecuteNavigateAndGetStatus = 500 ' Simulated failure code
        Err.Clear
    End If
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
    
    ' 1. Enable Network domain to allow URL-based blocking
    Set res = ExecuteCDPCommand("Network.enable", New Dictionary, True)
    If Not res("Ok") Then
        EnsureCdpBlockingInitialized = False
        Exit Function
    End If

    ' 2. Enable Fetch domain to allow intercepting and failing requests (requestPaused)
    ' An empty dictionary is sufficient for default fetch interception.
    Dim p As New Dictionary
    Set res = ExecuteCDPCommand("Fetch.enable", p, True)
    
    EnsureCdpBlockingInitialized = res("Ok")
End Function

' ========================================================================================
' ExecuteEnableResourceBlocking: Sets URL patterns to be blocked via CDP.
' Updated: Now ensures CDP domains are enabled before setting patterns.
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
    Dim resObj As Object: Set resObj = ExecuteCDPCommand("Network.setBlockedURLs", params)
    
    If Not resObj("Ok") Then
        ExecuteEnableResourceBlocking = "{""error"":""CDP_FAILURE"",""message"":""" & _
                                        Replace(resObj("Error"), """", "'") & """}"
    Else
        ExecuteEnableResourceBlocking = resObj("Json")
    End If
End Function

' ========================================================================================
' ExecuteDisableResourceBlocking: Clears all current URL blocking patterns.
' ========================================================================================
Public Function ExecuteDisableResourceBlocking() As String
    Dim params As New Dictionary
    Dim emptyList As New Collection
    params.Add "urls", emptyList
    
    Dim resObj As Object
    Set resObj = ExecuteCDPCommand("Network.setBlockedURLs", params)
    
    If resObj("Ok") Then
        ExecuteDisableResourceBlocking = resObj("Json")
    Else
        ExecuteDisableResourceBlocking = "{""error"":""CDP_FAILURE"",""message"":""" & _
                                        Replace(resObj("Error"), """", "'") & """}"
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
'              This method performs a visibility check via JavaScript and returns
'              True only if the element exists and is functionally visible.
' ========================================================================================
Public Function ExecuteIsElementVisible(ByVal xpath As String, _
                                        Optional ByVal searchTimeoutMs As Long = 2000, _
                                        Optional ByVal waitNetworkIdle As Boolean = False, _
                                        Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                        Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                        Optional ByVal contextId As String = "") As Boolean
    
    ' 1. Consolidated Validation: Sanitize all timing-related parameters at the entry point.
    ' This prevents logic errors if the user mistakenly passes 'True' (-1) as an argument.
    Dim vSearchTimeout As Long: vSearchTimeout = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 2000)
    Dim vMinStable As Long:     vMinStable = SanitizeMs(minStableMs, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vMaxTimeout As Long:    vMaxTimeout = SanitizeMs(maxTimeoutMs, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    
    ' 2. Resolve the target browsing context.
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    
    ' 3. Locate the element.
    ' Note: Now calling the Private version of GetFindElementSharedIdByXPath with 3 arguments.
    Dim sharedId As String
    sharedId = GetFindElementSharedIdByXPath(xpath, vSearchTimeout, targetCtx)
    
    ' 4. Early Exit: If the element is not found within the search timeout, it is considered not visible.
    If sharedId = "" Then
        ExecuteIsElementVisible = False
        Exit Function
    End If

    ' 5. Prepare JavaScript arguments for the visibility check script.
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "sharedId", sharedId: args.Add arg1

    ' 6. Execution: Perform the check via the central executor.
    ' The result is returned as a JSON response containing the boolean visibility state.
    Dim strRes As String
    strRes = ExecuteBaseAction(targetCtx, "visibility_check", args, waitNetworkIdle, vMinStable, True, DEFAULT_STALE_RETRIES, vMaxTimeout)
    
    ' 7. Parse Result: Extract the boolean value from the returned JSON structure.
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    ExecuteIsElementVisible = CBool(GetPathValue(json, "result.result.value.value"))
End Function

' ========================================================================================
' ExecuteSessionSubscribe: Subscribes to BiDi events with optional context filtering.
' Updated: Supports 'contexts' parameter and remains idempotent.
' ========================================================================================
Public Function ExecuteSessionSubscribe(ByVal events As Variant, Optional ByVal contexts As Collection = Nothing) As String
    Dim params As New Dictionary
    
    ' Normalize events into an array
    If IsArray(events) Then
        params.Add "events", events
    Else
        Dim arr(0) As String: arr(0) = events
        params.Add "events", arr
    End If
    
    ' Optional: Filter events by specific browsing contexts (tabs)
    If Not contexts Is Nothing Then
        params.Add "contexts", contexts
    End If

    ' Execute subscription (Idempotent: repeating the same subscription is safe)
    ExecuteSessionSubscribe = ExecuteBiDiCommand("session.subscribe", params)
End Function

' ========================================================================================
' FUNCTION: ExecuteBrowsingContextNavigate
' SUMMARY: Triggers navigation and ensures the page is fully stable/interactive.
' ========================================================================================
Public Function ExecuteBrowsingContextNavigate(ByVal url As String, _
                                                Optional ByVal waitMode As String = "complete", _
                                                Optional ByVal contextId As String = "") As String
    Dim targetCtx As String
    ' Replaced IIf with If-Else to prevent redundant UpdateMainContextId calls
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId(True)
    End If

    ' Reset internal state and pending counts before navigation
    p_pendingRequestCount = 0
    ClearState
    
    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "url", url
    params.Add "wait", waitMode

    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.navigate", params)

    ' Finalize with SPA stability consensus if waitMode is 'complete'
    If waitMode = "complete" Then
        UpdateMainContextId True
        WaitForNetworkIdleSpa DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
    End If
    
    ExecuteBrowsingContextNavigate = strRes
End Function

' ========================================================================================
' [PUBLIC] ExecuteClickByXPath
' ========================================================================================
Public Function ExecuteClickByXPath(ByVal xpath As String, _
                                    Optional ByVal searchTimeoutMs As Long = 5000, _
                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                    Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                    Optional ByVal contextId As String = "") As String
    
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId()
    
    ' 1. Resolve and validate element
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteClickByXPath")

    ' 2. Prepare arguments
    Dim args As New Collection
    Dim arg1 As New Dictionary: arg1.Add "sharedId", sharedId: args.Add arg1

    ' 3. Execute action via central runner
    ExecuteClickByXPath = ExecuteBaseAction(targetCtx, "click", args, waitNetworkIdle, minStableMs, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
End Function

' ========================================================================================
' [PUBLIC] ExecuteInputValueByXPath
' ========================================================================================
Public Function ExecuteInputValueByXPath(ByVal xpath As String, _
                                         ByVal valueToSet As String, _
                                         Optional ByVal searchTimeoutMs As Long = 5000, _
                                         Optional ByVal waitNetworkIdle As Boolean = True, _
                                         Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                         Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                         Optional ByVal contextId As String = "") As String
    
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId()
    
    Dim sharedId As String
    sharedId = ResolveXPathOrThrow(xpath, targetCtx, searchTimeoutMs, "ExecuteInputValueByXPath")

    Dim args As New Collection
    Dim a1 As New Dictionary: a1.Add "sharedId", sharedId: args.Add a1
    Dim a2 As New Dictionary: a2.Add "type", "string": a2.Add "value", valueToSet: args.Add a2

    ExecuteInputValueByXPath = ExecuteBaseAction(targetCtx, "input", args, waitNetworkIdle, minStableMs, True, DEFAULT_STALE_RETRIES, maxTimeoutMs)
End Function

' ========================================================================================
' [PUBLIC] ExecuteSelectValueByXPath
' ========================================================================================
Public Function ExecuteSelectValueByXPath(ByVal xpath As String, _
                                          ByVal valueOrText As String, _
                                          Optional ByVal selectByText As Boolean = False, _
                                          Optional ByVal searchTimeoutMs As Long = 5000, _
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
Private Function ExecuteBrowsingContextLocateNodes(ByVal xpath As String, Optional ByVal contextId As String = "") As String
    Dim targetCtx As String: targetCtx = ResolveTargetContext(contextId)
    Dim params As New Dictionary, locator As New Dictionary
    locator.Add "type", "xpath"
    locator.Add "value", xpath
    params.Add "context", targetCtx
    params.Add "locator", locator
    ExecuteBrowsingContextLocateNodes = ExecuteBiDiCommand("browsingContext.locateNodes", params)
End Function

' ========================================================================================
' [INTERNAL] GetFindElementSharedIdByXPath (Now Private)
' SUMMARY: Internal locator bottleneck. Resolves XPath to sharedId with validation.
' ========================================================================================
Private Function GetFindElementSharedIdByXPath(ByVal xpath As String, _
                                               ByVal searchTimeoutMs As Long, _
                                               ByVal targetCtx As String) As String
    ' VALIDATION: Sanitize search timeout at the entry point
    Dim vTimeout As Long
    vTimeout = SanitizeMs(searchTimeoutMs, "searchTimeoutMs", 2000)
    
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim strRes As String, json As Object, nodes As Collection

    Do
        strRes = ExecuteBrowsingContextLocateNodes(xpath, targetCtx)
        Set json = WebJsonConverter.ParseJson(strRes)
        Set nodes = GetPathObject(json, "result.nodes")

        If Not nodes Is Nothing Then
            If nodes.Count > 0 Then
                If nodes(1).Exists("sharedId") Then
                    GetFindElementSharedIdByXPath = CStr(nodes(1)("sharedId"))
                    Exit Function
                End If
            End If
        End If

        If vTimeout <= 0 Then Exit Do
        WaitHeartbeat 100
        If IsTimeout(startTick, CDbl(vTimeout)) Then Exit Do
    Loop
    GetFindElementSharedIdByXPath = ""
End Function

' ========================================================================================
' [PUBLIC] GetCurrentUrl (Optimized)
' ========================================================================================
Public Function GetCurrentUrl(Optional ByVal contextId As String = "") As String
    Dim strRes As String
    strRes = CallScript("function(){ return window.location.href; }", , contextId)
    
    ' Extract value from the already validated response
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    GetCurrentUrl = GetPathString(json, "result.result.value")
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
                                   Optional ByVal searchTimeoutMs As Long = 5000, _
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
' [INTERNAL] ExecuteBiDiCommand
' SUMMARY: Core dispatcher for BiDi protocol commands.
'          Features:
'          - Automated recovery for destroyed execution contexts (e.g., page navigation).
'          - Exponential backoff with segmented sleep to maintain Excel responsiveness.
'          - Centralized error processing for protocol and script-level exceptions.
' ========================================================================================
Private Function ExecuteBiDiCommand(ByVal method As String, ByVal params As Dictionary) As String
    If socket_ Is Nothing Then Err.Raise 91, "ExecuteBiDiCommand", "Socket object is not set."

    Dim retryCount As Long
    Dim strRes As String, sendMsg As String
    Dim requestDic As Dictionary
    
    ' Retry loop for handling transient environment changes (e.g., context reloads)
    For retryCount = 0 To MAX_CONTEXT_RECOVERY_RETRIES
        ' 1. Prepare the standard request object
        Set requestDic = New Dictionary
        requestDic.Add "id", p_nextRequestId
        requestDic.Add "method", method
        requestDic.Add "params", params
        p_nextRequestId = p_nextRequestId + 1

        ' Serialize dictionary to JSON string
        sendMsg = WebJsonConverter.ConvertToJson(requestDic)
        
        If DEBUG_MODE Then DebugPrintLong method & " (Try " & retryCount & ")", sendMsg, "[REQ]"
        
        ' 2. Transmission via WebSocket
        strRes = socket_.SendAndReceive(sendMsg)
        
        If DEBUG_MODE Then DebugPrintLong method & " (Try " & retryCount & ")", strRes, "[RES]"
        
        ' 3. Context Recovery Check
        ' Determine if the error is due to the browser environment disappearing (destroyed context)
        If IsContextLostError(strRes) And Not m_isRecovering And retryCount < MAX_CONTEXT_RECOVERY_RETRIES Then
            On Error GoTo RecoveryCleanup
            m_isRecovering = True

            ' Exponential Backoff: Calculate delay based on attempt count (300ms, 600ms, 1200ms...)
            Dim waitMs As Long: waitMs = 300 * (2 ^ retryCount)
            If waitMs > 2400 Then waitMs = 2400 ' Cap recovery wait at 2.4 seconds
            
            DebugPrintLong "Context lost. Attempting recovery...", strRes, "[RECOVERY]"
            
            ' REFINEMENT: Segmented sleep with DoEvents to prevent Excel from freezing during wait
            Dim totalSlept As Long: totalSlept = 0
            Do While totalSlept < waitMs
                sleep 100    ' Sleep in 100ms chunks
                DoEvents     ' Keep Excel responsive to user interrupts (Esc key, etc.)
                totalSlept = totalSlept + 100
            Loop

            ' Refresh internal identifiers
            ClearState ' Clear cached IDs
            Dim newCtx As String: newCtx = UpdateMainContextId(True)

            ' Abort if the context cannot be found after refresh
            If newCtx = "" Then
                m_isRecovering = False
                Call ProcessResponse(strRes, method) ' This will raise the original protocol error
                Exit For
            End If

            ' Update the original params dictionary in-place for the next retry
            If Not UpdateContextAndRealm(params, newCtx) Then
                m_isRecovering = False
                Call ProcessResponse(strRes, method)
                Exit For
            End If

            m_isRecovering = False
            ' Continue to next retry with refreshed context/realm IDs
        Else
            ' 4. Final Validation & Error Dispatching
            ' Pass response to the centralized validator (detects JS exceptions & functional errors)
            Call ProcessResponse(strRes, method)
            Exit For
        End If
    Next retryCount

    ExecuteBiDiCommand = strRes
    Exit Function

RecoveryCleanup:
    m_isRecovering = False
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "BiDi:Recovery", "Critical failure during recovery: " & Err.Description
    End If
End Function

' ========================================================================================
' Internal Helpers for ExecuteBiDiCommand
' ========================================================================================

Private Function IsContextLostError(ByVal response As String) As Boolean
    ' Identify common browser protocol errors indicating that the execution environment is gone
    IsContextLostError = (InStr(response, "Execution context was destroyed") > 0) Or _
                         (InStr(response, "Inspected target navigated or closed") > 0) Or _
                         (InStr(response, "no such frame") > 0) Or _
                         (InStr(response, "Cannot find context") > 0) Or _
                         (InStr(response, "realm") > 0 And InStr(response, "not found") > 0)
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
                ' Perform case-insensitive key matching for context and realm
                If LCase(CStr(k)) = "context" Then
                    node(k) = newCtx
                ElseIf LCase(CStr(k)) = "realm" And Len(newRealm) > 0 Then
                    node(k) = newRealm
                End If
                
                ' Recurse into the value if it's a Dictionary or Collection
                Call UpdateContextAndRealmRecursive(node(k), newCtx, newRealm)
            Next k
            UpdateContextAndRealmRecursive = True
        ElseIf TypeName(node) = "Collection" Then
            Dim i As Long
            For i = 1 To node.Count
                ' VBA Collections are 1-based. Recurse into each element.
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
    
    ' Fetch a fresh Realm ID for the newly identified context
    Dim newRealm As String
    newRealm = GetScriptRealmId(newCtx, True)
    
    ' If realm retrieval fails, the recovery cannot proceed reliably
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
    Debug.Print timePrefix & tag & " " & header & " (Length: " & Len(cleanText) & ")"

    ' Chunked output to bypass VBA's Immediate Window line length limits
    If Len(cleanText) > 0 Then
        For i = 1 To Len(cleanText) Step CHUNK_SIZE
            ' Print parts of the string without an automatic newline
            Debug.Print Mid$(cleanText, i, CHUNK_SIZE);
        Next i
    End If

    ' Finalize the log entry with a newline for clarity
    Debug.Print ""
End Sub

' ========================================================================================
' ClearState: Resets internal state and cleans up browser-side side effects.
' ========================================================================================
Public Sub ClearState()
    ' No need to restore JS env manually; navigation or closing will handle it.
    p_cdpSessionId = ""
    p_mainContextId = ""
    p_pendingRequestCount = 0
    m_cachedRealmId = "" ' Also clear cached realm
End Sub

' ========================================================================================
' FUNCTION: GetMainContextId
' DESCRIPTION: Returns the currently active main context ID.
'              If not yet set, it triggers a discovery of the browser tree.
' ========================================================================================
Public Function GetMainContextId() As String
    If p_mainContextId = "" Then
        GetMainContextId = UpdateMainContextId(True)
    Else
        GetMainContextId = p_mainContextId
    End If
End Function

' ========================================================================================
' FUNCTION: UpdateMainContextId
' SUMMARY: Refreshes the browsing context tree. Raises error if no context is available.
' ========================================================================================
Public Function UpdateMainContextId(Optional ByVal forceRefresh As Boolean = False) As String
    ' Return cached ID if possible
    If Not forceRefresh And p_mainContextId <> "" Then
        UpdateMainContextId = p_mainContextId
        Exit Function
    End If

    Dim jsonRes As String
    jsonRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
    
    Dim jsonTree As Object
    Set jsonTree = WebJsonConverter.ParseJson(jsonRes)
    
    ' Extract the first available context ID
    Dim contextId As String
    contextId = GetPathString(jsonTree, "result.contexts.0.context")
    
    If contextId <> "" Then
        p_mainContextId = contextId
        UpdateMainContextId = contextId
    Else
        ' Fix: Raise error if no browsing context exists (e.g., browser was closed)
        Err.Raise ERR_ACTION_FAILED, "BiDi:UpdateMainContextId", _
                  "Failed to identify any browsing context. Ensure the browser is active and connected."
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
                                 Optional ByVal forceRefresh As Boolean = False) As String
    Dim targetCtx As String
    
    ' 1. Determine target context
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = GetMainContextId()
    End If
    
    If targetCtx = "" Then Exit Function

    ' 2. Return cached Realm ID if available and refresh not forced
    If Not forceRefresh And m_cachedRealmId <> "" Then
        GetScriptRealmId = m_cachedRealmId
        Exit Function
    End If

    ' 3. Fetch all realms associated with this context
    Dim params As New Dictionary
    params.Add "context", targetCtx
    
    Dim jsonRes As String
    On Error Resume Next
    jsonRes = ExecuteBiDiCommand("script.getRealms", params)
    On Error GoTo 0
    
    ' 4. Filter for the "window" type realm
    Dim realms As Collection
    Set realms = GetJsonValue(jsonRes, "result.realms")
    
    Dim foundRealmId As String: foundRealmId = ""
    
    If Not realms Is Nothing Then
        If realms.Count > 0 Then
            Dim r As Dictionary
            ' Iterate through all realms to find the 'window' type
            For Each r In realms
                If r.Exists("type") Then
                    If r("type") = "window" Then
                        foundRealmId = CStr(r("realm"))
                        Exit For ' Found the primary UI realm
                    End If
                End If
            Next r
            
            ' Fallback: If no window type found, take the first one as a last resort
            If foundRealmId = "" Then foundRealmId = CStr(realms(1)("realm"))
        End If
    End If
    
    ' 5. Update cache and return
    m_cachedRealmId = foundRealmId
    GetScriptRealmId = m_cachedRealmId
    
    If foundRealmId = "" And DEBUG_MODE Then
        Debug.Print "BiDi Warning: No valid Realm ID found for context: " & targetCtx
    End If
End Function

' Gets the current CDP session ID, creating a new one if necessary.
Private Function GetCdpSessionId(Optional ByVal forceUpdate As Boolean = False) As String
    If p_mainContextId = "" Then UpdateMainContextId
    
    If Not forceUpdate And p_cdpSessionId <> "" Then
        GetCdpSessionId = p_cdpSessionId
        Exit Function
    End If

    Dim params As New Dictionary
    params.Add "context", p_mainContextId
    
    ' Execute command and parse response safely
    Dim strRes As String: strRes = ExecuteBiDiCommand("goog:cdp.getSession", params)
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    
    ' Extract session ID using the path helper to avoid Runtime Error 91
    Dim sessionId As String: sessionId = GetPathString(json, "result.session")
    
    If sessionId <> "" Then
        p_cdpSessionId = sessionId
        m_consecutiveCdpFailures = 0
        GetCdpSessionId = p_cdpSessionId
    Else
        m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
        If DEBUG_MODE Then Debug.Print "BiDi Warning: Failed to acquire CDP Session."
        GetCdpSessionId = ""
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
                                             Optional ByVal timeoutMs As Long = 30000)
    ' --- Input Validation ---
    Dim vTimeout As Long: vTimeout = SanitizeMs(timeoutMs, "timeoutMs", 30000)
    
    ' Retrieve centralized JS template
    Dim js As String: js = GetActionJs("autoclicker")
    
    ' Inject parameters safely into the source code
    Dim sourceJs As String
    sourceJs = Replace(js, "arguments[0]", WebJsonConverter.ConvertToJson(xpath))
    sourceJs = Replace(sourceJs, "arguments[1]", vTimeout) ' Use validated value

    Dim cdpParams As New Dictionary
    cdpParams.Add "source", sourceJs
    
    ' Register the observer script via CDP tunnel
    Dim result As Object
    Set result = ExecuteCDPCommand("Page.addScriptToEvaluateOnNewDocument", cdpParams)
    
    If Not result("Ok") Then
        Debug.Print "BiDi Warning: AutoClicker registration failed: " & result("Error")
    End If
End Sub

Private Function BuildShadowSelectorArg(ByVal selectorsArray As Variant) As Dictionary
    Dim arg As New Dictionary
    arg.Add "type", "array"
    Dim valList As New Collection, i As Long, valItem As Dictionary
    For i = LBound(selectorsArray) To UBound(selectorsArray)
        Set valItem = New Dictionary
        valItem.Add "type", "string"
        valItem.Add "value", selectorsArray(i)
        valList.Add valItem
    Next i
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
    p_isRecording = False

    If p_recorderLogs Is Nothing Then Exit Sub
    
    ' Save logic
    If p_recorderLogs.Count > 0 Then
        Dim fso As New FileSystemObject
        Dim ts As TextStream
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

    ' Explicit cleanup to prevent memory leaks with nested objects
    Dim i As Long
    For i = p_recorderLogs.Count To 1 Step -1
        p_recorderLogs.Remove i
    Next i
    Set p_recorderLogs = Nothing
End Sub

Public Sub RecordEventsForSeconds(ByVal seconds As Long)
    Dim endTick As Double
    endTick = GetRawTicks() + (seconds * 1000)

    Do While GetRawTicks() < endTick
        ProcessRecorderQueue
        sleep 100
        DoEvents
    Loop
End Sub

' ========================================================================================
' PUBLIC SUB: ProcessRecorderQueue
' SUMMARY: Processes the incoming message queue with re-entrancy protection.
'          Uses a simple Boolean flag for state management (VBA is single-threaded).
' ========================================================================================
Public Sub ProcessRecorderQueue()
    ' --- 1. Re-entrancy Guard ---
    ' Check if the queue is already being processed.
    ' Since VBA is single-threaded, this flag is sufficient to prevent overlapping calls
    ' triggered by DoEvents or nested wait heartbeats.
    If m_isProcessingQueue Then Exit Sub
    
    ' Set the lock
    m_isProcessingQueue = True
    
    On Error GoTo Cleanup
    
    Dim msg As String
    Dim json As Object
    Dim processedCount As Long
    Const MAX_MESSAGES_PER_BATCH As Long = 1000
    
    ' Main processing loop
    Do While processedCount < MAX_MESSAGES_PER_BATCH
        msg = socket_.GetMessage()
        If Len(msg) = 0 Then Exit Do
        
        ' Basic size/structure validation
        If Len(msg) >= MIN_MSG_SIZE And Len(msg) <= MAX_MSG_SIZE Then
            If InStr(msg, """method"":") > 0 Or InStr(msg, """id"":") > 0 Then
                On Error Resume Next
                Set json = WebJsonConverter.ParseJson(msg)
                If Err.Number = 0 And Not json Is Nothing Then
                    On Error GoTo Cleanup ' Restore main handler
                    DispatchGlobalEvents json
                End If
                On Error GoTo Cleanup
            End If
        End If
        
        processedCount = processedCount + 1
        
        ' Yield to OS every 100 messages to keep UI responsive
        If (processedCount Mod 100) = 0 Then ThrottledDoEvents
    Loop

Cleanup:
    ' --- 2. Release Lock ---
    ' Always reset the flag to False so that future calls can enter.
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

Private Function IsNoiseUrl(ByVal url As String) As Boolean
    ' Decide whether a URL is noisy for discovery logging.
    ' Generic filters: images, fonts, styles, data URLs, common analytics collectors.
    Dim lUrl As String: lUrl = LCase(url)
    IsNoiseUrl = (Left(lUrl, 5) = "data:" Or InStr(lUrl, ".png") > 0 Or InStr(lUrl, ".jpg") > 0 Or InStr(lUrl, ".jpeg") > 0 Or InStr(lUrl, ".gif") > 0 Or InStr(lUrl, ".webp") > 0 Or InStr(lUrl, ".svg") > 0 Or InStr(lUrl, ".css") > 0 Or InStr(lUrl, ".woff") > 0 Or InStr(lUrl, ".woff2") > 0 Or InStr(lUrl, ".ico") > 0 Or InStr(lUrl, "google-analytics") > 0 Or InStr(lUrl, "doubleclick") > 0 Or InStr(lUrl, "googletagmanager") > 0 Or InStr(lUrl, "/collect") > 0 Or InStr(lUrl, "/analytics") > 0 Or InStr(lUrl, "/pixel") > 0 Or InStr(lUrl, "/beacon") > 0)
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
    Dim jsonResponse As String
    Dim resultValue As Variant
    Dim dict As Dictionary
    
    ' Execute the JS snippet to get current metrics
    jsonResponse = CallScript(GetProbeJs("query"), , contextId)
    resultValue = GetScriptResultValue(jsonResponse)
    
    ' Fail gracefully if BiDi returns an empty result
    If IsEmpty(resultValue) Or IsNull(resultValue) Then
        Set QuerySpaProbe = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    ' Handle different BiDi return types (stringified JSON vs direct Object)
    If VarType(resultValue) = vbString Then
        Set dict = WebJsonConverter.ParseJson(CStr(resultValue))
    ElseIf IsObject(resultValue) Then
        Set dict = resultValue
    End If
    On Error GoTo 0
    
    ' Validation Check:
    ' Even if the JS function exists, the state object (__vbaIdleProbe) might be gone.
    If Not dict Is Nothing Then
        If dict.Exists("active") Then
            If dict("active") = False Then
                ' Force a recovery by returning Nothing
                Set QuerySpaProbe = Nothing
                Exit Function
            End If
        End If
    End If
    
    Set QuerySpaProbe = dict
End Function

' ========================================================================================
' PROCEDURE: RecoverSpaProbe
' SUMMARY: Nukes the state cache and re-injects the probe into the fresh context.
' ========================================================================================
Private Sub RecoverSpaProbe(ByVal contextId As String)
    ' 1. Invalidate all cached identifiers
    m_cachedRealmId = ""
    
    ' 2. Force a full tree refresh to ensure the contextId is still valid
    ' This handles cases where the tab navigated to a completely different origin.
    UpdateMainContextId True
    
    ' 3. Re-inject. StartSpaIdleProbe will now target 'context' by default.
    StartSpaIdleProbe contextId
    
    If DEBUG_MODE Then Debug.Print "BiDi: Recovery completed for context: " & contextId
End Sub

' ========================================================================================
' PUBLIC SUB: WaitForNetworkIdleSpa (Version 3.2 - Final)
' SUMMARY: Entry point for deep SPA-aware synchronization.
'          This procedure prepares the wait context and starts the MasterWaitEngine.
'          Actual probe injection and recovery are handled JIT inside the wait loop.
' ========================================================================================
Public Sub WaitForNetworkIdleSpa(ByVal minStableMs As Long, _
                                 ByVal maxTimeoutMs As Long, _
                                 Optional ByVal contextId As String = "")
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
    waitCtx.Add "stableStart", GetRawTicks()
    waitCtx.Add "missingProbeCount", 0 ' Used for multi-stage JIT recovery
    
    ' Execute the high-precision wait engine with SPA Idle mode.
    MasterWaitEngine WaitMode_SpaIdle, maxTimeoutMs, 100, waitCtx
End Sub

' ========================================================================================
' FUNCTION: EvaluateSpaIdleCondition
' DESCRIPTION: Determines if the page has reached a stable idle state.
'              Consensus = (BiDi Net Idle) AND (JS Net Idle) AND (DOM Stable).
' ========================================================================================
Private Function EvaluateSpaIdleCondition(ByVal probe As Dictionary, _
                                          ByVal minStableMs As Long, _
                                          ByRef stableStart As Double, _
                                          ByVal nowTick As Double) As Boolean
    
    Dim jsNow As Double, lastActivityTs As Double, jsIdleDuration As Double
    Dim inflFetch As Long, inflXhr As Long
    Dim netIdle As Boolean, activityStable As Boolean

    ' 1. Extract high-res metrics from the browser probe
    jsNow = CDbl(probe("now"))
    lastActivityTs = CDbl(probe("lastMutationTs"))
    inflFetch = CLng(probe("inflightFetchCount"))
    inflXhr = CLng(probe("inflightXhrCount"))
    
    ' Time elapsed since the last structural DOM change or Network end
    jsIdleDuration = jsNow - lastActivityTs

    ' 2. Network Consensus Logic
    ' Require agreement between the protocol layer (BiDi) and the application layer (JS Hooks).
    netIdle = (p_pendingRequestCount <= 0) And (inflFetch = 0) And (inflXhr = 0)

    ' 3. Self-Correction: Safe Counter Decay
    ' If JS reports zero activity but BiDi still has a positive count,
    ' wait for twice the stability window before assuming a missed event.
    If (inflFetch = 0 And inflXhr = 0) And (p_pendingRequestCount > 0) Then
        If jsIdleDuration > (CDbl(minStableMs) * 2) Then
            ' Recover from missed 'network.responseCompleted' events
            p_pendingRequestCount = p_pendingRequestCount - 1
        End If
    End If

    ' 4. DOM Stability Check
    activityStable = (jsIdleDuration >= CDbl(minStableMs))

    ' 5. Final Threshold Evaluation
    If netIdle And activityStable Then
        ' Check if this idle state has been maintained long enough on the VBA side
        If (nowTick - stableStart) >= CDbl(minStableMs) Then
            EvaluateSpaIdleCondition = True ' CONVERGENCE ACHIEVED
        Else
            EvaluateSpaIdleCondition = False
        End If
    Else
        ' Activity detected; reset the VBA-side stability start point
        stableStart = nowTick
        EvaluateSpaIdleCondition = False
    End If
End Function

' ========================================================================================
' ExecuteCDPCommand: Sends a raw CDP command via BiDi tunneling.
' ========================================================================================
Private Function ExecuteCDPCommand(ByVal cdpMethod As String, _
                                  ByVal cdpParams As Dictionary, _
                                  Optional ByVal useSession As Boolean = True) As Object
    
    Dim res As New Dictionary: res.Add "Ok", False: res.Add "Json", "": res.Add "Error", ""

    ' 1. Connectivity Check
    If socket_ Is Nothing Then
        res("Error") = "Socket object is not set.": Set ExecuteCDPCommand = res: Exit Function
    End If

    ' 2. Health Check
    If m_consecutiveCdpFailures >= MAX_CONSECUTIVE_CDP_FAILURES Then
        PerformFatalReset
        res("Error") = "Fatal: CDP instability.": Set ExecuteCDPCommand = res: Exit Function
    End If

    ' 3. Parameter Preparation
    Dim params As New Dictionary
    params.Add "method", cdpMethod
    ' CLEANED: Use the helper instead of inline IIf logic
    params.Add "params", EnsureDictionary(cdpParams)

    ' 4. Session Attachment
    If useSession Then
        Dim sessionId As String: sessionId = GetCdpSessionId()
        If sessionId <> "" Then
            params.Add "session", sessionId
        Else
            res("Error") = "CDP session unavailable.": Set ExecuteCDPCommand = res: Exit Function
        End If
    End If

    ' 5. Execution and Error Handling
    Dim strRes As String
    On Error Resume Next
    strRes = ExecuteBiDiCommand("goog:cdp.sendCommand", params)
    
    ' Use GetJsonValue to safely check for BiDi/CDP errors
    Dim protocolErr As String: protocolErr = CStr(GetJsonValue(strRes, "error"))

    If Err.Number <> 0 Then
        res("Error") = "Socket Error: " & Err.Description
    ElseIf protocolErr <> "" Then
        res("Error") = "CDP Error: " & strRes
    Else
        res("Ok") = True: res("Json") = strRes
    End If
    On Error GoTo 0

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
    
    ' FIX: Short-circuit context identification.
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId(True)
    End If
    
    ' Retrieve the Realm ID (this will use the cache if available).
    Dim targetRealm As String
    targetRealm = GetScriptRealmId(targetCtx)
    
    Dim p As New Dictionary
    p.Add "expression", GetProbeJs("inject")
    p.Add "target", BuildContextTarget(targetCtx, targetRealm)
    p.Add "awaitPromise", False

    ' Send injection command.
    Call ExecuteBiDiCommand("script.evaluate", p)
End Sub

' ========================================================================================
' FUNCTION: GetScriptResultValue
' SUMMARY: Extracts the actual return value from a script execution.
' ========================================================================================
Private Function GetScriptResultValue(ByVal jsonResponse As String) As Variant
    ' Use the core traversal helper directly as we need the raw Variant (String, Double, Boolean)
    GetScriptResultValue = InternalJsonTraverse(jsonResponse, "result.result.value")
End Function

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
'              Supports Dictionary keys and Collection indices (e.g., "results.0.id").
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
    For i = 0 To UBound(keys) - 1
        currentKey = keys(i)
        
        If TypeName(target) = "Dictionary" Then
            If Not target.Exists(currentKey) Then Exit Function
            Set target = target(currentKey)
        
        ElseIf TypeName(target) = "Collection" Then
            ' FIX: Support array indexing. JSON is 0-based, VBA Collection is 1-based.
            If IsNumeric(currentKey) Then
                Dim idx As Long: idx = CLng(currentKey) + 1
                If idx < 1 Or idx > target.Count Then Exit Function
                Set target = target(idx)
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    Next i

    ' 3. Extract the final value
    If Not target Is Nothing Then
        Dim lastKey As String: lastKey = keys(UBound(keys))
        
        If TypeName(target) = "Dictionary" Then
            If target.Exists(lastKey) Then
                If IsObject(target(lastKey)) Then
                    Set GetJsonValue = target(lastKey)
                Else
                    GetJsonValue = target(lastKey)
                End If
            End If
        ElseIf TypeName(target) = "Collection" And IsNumeric(lastKey) Then
            Dim fIdx As Long: fIdx = CLng(lastKey) + 1
            If IsObject(target(fIdx)) Then
                Set GetJsonValue = target(fIdx)
            Else
                GetJsonValue = target(fIdx)
            End If
        End If
    End If
    On Error GoTo 0
End Function

' ========================================================================================
' INTERNAL JS HELPER: GetActionJs
' DESCRIPTION: Provides centralized JavaScript snippets for all browser actions.
'              All actions return a structured object: { status: "ok", value: ... }
'              or { status: "error", message: "..." } to ensure robustness.
' ========================================================================================
Private Function GetActionJs(ByVal actionType As String) As String
    Dim js As String
    
    Select Case LCase(actionType)
        Case "click"
            js = "function(e) { return new Promise((r) => { " & _
                 "  try { " & _
                 "    e.scrollIntoView({block:'center'}); e.focus(); e.click(); " & _
                 "    r({status: 'ok'}); " & _
                 "  } catch(err) { r({status: 'error', message: err.message}); } " & _
                 "}); }"
                 
        Case "input"
            ' Focuses, clears, and inserts text via execCommand to trigger all event listeners.
            js = "function(e,v) { return new Promise((r) => { " & _
                 "  try { " & _
                 "    e.scrollIntoView({block:'center',inline:'center'}); " & _
                 "    e.click(); e.focus(); e.value=''; " & _
                 "    var s=document.execCommand('insertText',false,v); " & _
                 "    if(!s){ e.value=v; e.dispatchEvent(new Event('input',{bubbles:true})); e.dispatchEvent(new Event('change',{bubbles:true})); } " & _
                 "    e.blur(); setTimeout(() => r({status: 'ok'}), 50); " & _
                 "  } catch(err) { r({status: 'error', message: err.message}); } " & _
                 "}); }"
                 
        Case "select"
            ' Selects an option by its value property.
            js = "function(s,v) { return new Promise((r) => { " & _
                 "  try { " & _
                 "    s.scrollIntoView({block:'center',inline:'center'}); " & _
                 "    s.focus(); s.value = v; " & _
                 "    s.dispatchEvent(new Event('input', {bubbles:true})); " & _
                 "    s.dispatchEvent(new Event('change', {bubbles:true})); " & _
                 "    s.blur(); setTimeout(() => r({status: 'ok'}), 50); " & _
                 "  } catch(err) { r({status: 'error', message: err.message}); } " & _
                 "}); }"

        Case "select_text"
            ' Selects an option by its visible innerText (normalized).
            js = "function(s, textToFind) { return new Promise((r) => { " & _
                 "  try { " & _
                 "    const normalize = (str) => str.replace(/[\\s\\u00A0]+/g, ' ').trim(); " & _
                 "    var found = false; var targetText = normalize(textToFind); " & _
                 "    for (var i = 0; i < s.options.length; i++) { " & _
                 "      if (normalize(s.options[i].text) === targetText) { s.value = s.options[i].value; found = true; break; } " & _
                 "    } " & _
                 "    if (!found) { r({status: 'error', message: 'Option text not found: ' + textToFind}); return; } " & _
                 "    s.scrollIntoView({block:'center'}); s.focus(); " & _
                 "    s.dispatchEvent(new Event('input', {bubbles:true})); s.dispatchEvent(new Event('change', {bubbles:true})); " & _
                 "    s.blur(); setTimeout(() => r({status: 'ok'}), 50); " & _
                 "  } catch(err) { r({status: 'error', message: err.message}); } " & _
                 "}); }"

        Case "visibility_check"
            ' Checks if an element is rendered and visible in the viewport.
            ' Returns the boolean result wrapped in an 'ok' status for consistency.
            js = "function(el) { " & _
                 "  if (!el || !el.isConnected) return {status: 'ok', value: false}; " & _
                 "  const rect = el.getBoundingClientRect(); " & _
                 "  const style = window.getComputedStyle(el); " & _
                 "  const isVisible = (rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0'); " & _
                 "  return {status: 'ok', value: isVisible}; " & _
                 "}"
                 
        Case "ready_state"
            ' Simple check for document completion status.
            js = "function(){ return document.readyState; }"
            
        Case "shadow_click"
            ' Recursively penetrates Shadow Roots to find and click a deep element.
            js = "function(selectors, searchTimeout) { " & _
                 "  const find = (arr, timeout) => { " & _
                 "    return new Promise((res) => { " & _
                 "      const end = Date.now() + timeout; " & _
                 "      const check = () => { " & _
                 "        let el = document.querySelector(arr[0]); " & _
                 "        if(el) { for(let i=1; i<arr.length; i++){ if(el.shadowRoot) el = el.shadowRoot.querySelector(arr[i]); else { el=null; break; } if(!el) break; } } " & _
                 "        if(el) res(el); else if(Date.now() < end) setTimeout(check, 100); else res(null); " & _
                 "      }; check(); " & _
                 "    }); " & _
                 "  }; " & _
                 "  return find(selectors, searchTimeout).then(el => { " & _
                 "    if(!el) return {status: 'error', message: 'Shadow element not found within timeout'}; " & _
                 "    el.scrollIntoView({block:'center'}); el.focus(); el.click(); " & _
                 "    return {status: 'ok'}; " & _
                 "  }).catch(err => ({status: 'error', message: err.message})); " & _
                 "}"
    
        Case "wait_options"
            ' Polls the select element until minCount options are available or timeout.
            js = "function(element, minCount, timeout) { " & _
                 "  return new Promise((resolve) => { " & _
                 "    const start = Date.now(); " & _
                 "    const check = () => { " & _
                 "      if (element.options && element.options.length >= minCount) { resolve(true); } " & _
                 "      else if (Date.now() - start > timeout) { resolve(false); } " & _
                 "      else { setTimeout(check, 100); } " & _
                 "    }; check(); " & _
                 "  }); " & _
                 "}"

        Case "get_options"
            ' Serializes all options of a select element to a JSON string.
            js = "function(element) { " & _
                 "  var opts = []; " & _
                 "  for (var i = 0; i < element.options.length; i++) { " & _
                 "    var o = element.options[i]; " & _
                 "    opts.push({ 'text': o.innerText, 'value': o.value, 'selected': o.selected }); " & _
                 "  } " & _
                 "  return JSON.stringify(opts); " & _
                 "}"
    
        Case "click_alert"
            ' Clicks an element asynchronously to prevent VBA from hanging when a modal appears.
            js = "function(e) { " & _
                 "  setTimeout(() => e.click(), 0); " & _
                 "  return {status: 'ok'}; " & _
                 "}"

        Case "autoclicker"
            ' Background MutationObserver to click a specific XPath as soon as it appears.
            js = "(function() { " & _
                 "  var xpath = arguments[0]; var timeout = arguments[1]; var start = Date.now(); " & _
                 "  const check = () => { " & _
                 "    try { " & _
                 "      var res = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null); " & _
                 "      var btn = res.singleNodeValue; " & _
                 "      if (btn && btn.offsetParent !== null) { btn.click(); return true; } " & _
                 "    } catch(e) {} return false; " & _
                 "  }; " & _
                 "  if (check()) return; " & _
                 "  var obs = new MutationObserver((mutations, o) => { " & _
                 "    if (check() || (Date.now() - start > timeout)) { o.disconnect(); } " & _
                 "  }); " & _
                 "  obs.observe(document, { childList: true, subtree: true }); " & _
                 "  setTimeout(() => obs.disconnect(), timeout + 1000); " & _
                 "})();"
                     
    End Select
    
    GetActionJs = js
End Function

' ========================================================================================
' FUNCTION: CallScript (Version 2.8 - Context-Target Support)
' SUMMARY: Executes script.callFunction. Prefers 'context' targeting over 'realm'
'          to maintain stability during navigation milestones (bfcache/prerender).
' ========================================================================================
Private Function CallScript(ByVal js As String, _
                           Optional ByVal args As Collection = Nothing, _
                           Optional ByVal contextId As String = "", _
                           Optional ByVal awaitPromise As Boolean = False, _
                           Optional ByVal forceRealmTarget As Boolean = False) As String
    Dim params As New Dictionary, targetDic As New Dictionary
    If args Is Nothing Then Set args = New Collection
    Dim resolvedCtx As String: resolvedCtx = ResolveTargetContext(contextId)
    
    If resolvedCtx = "" Then
        CallScript = "{""error"":""no_context"",""message"":""Failed to resolve contextId.""}"
        Exit Function
    End If
    
    ' --- STRATEGY: Target Resolution ---
    ' In WebDriver BiDi, targeting a 'context' is more robust than a specific 'realm' ID.
    ' If targeting a context, the browser automatically selects the active Realm.
    If forceRealmTarget Then
        ' Only use Realm ID for specific isolated environments (e.g., Workers)
        targetDic.Add "realm", GetScriptRealmId(resolvedCtx)
    Else
        ' Recommended for DOM Probes: Targets the browsing context (tab/frame) directly.
        targetDic.Add "context", resolvedCtx
    End If
    
    params.Add "functionDeclaration", js
    params.Add "arguments", args
    params.Add "target", targetDic
    params.Add "awaitPromise", awaitPromise
    
    CallScript = ExecuteBiDiCommand("script.callFunction", params)
End Function

' ========================================================================================
' IsTimeout: Precision-safe timeout check.
' ========================================================================================
Private Function IsTimeout(ByVal startTicks As Currency, ByVal timeoutMs As Double) As Boolean
    If timeoutMs <= 0 Then Exit Function
    ' Simple comparison
    IsTimeout = GetElapsedMs(startTicks) > timeoutMs
End Function

' ========================================================================================
' Internal Helper: Centralizes background tasks during a wait loop.
' Keeps the WebSocket queue moving and the Excel UI responsive.
' ========================================================================================
Private Sub WaitHeartbeat(Optional ByVal sleepMs As Long = 50)
    ' 1. Process incoming BiDi events (Updates pending request counts, etc.)
    ProcessRecorderQueue
    
    ' 2. Process Windows events (Keeps Excel from freezing)
    ThrottledDoEvents
    
    ' 3. Release CPU briefly
    If sleepMs > 0 Then sleep sleepMs
End Sub

' ========================================================================================
' FUNCTION: GetProbeJs
' DESCRIPTION: Returns centralized JavaScript for SPA monitoring.
'              Version 3.1: Bulletproof Fetch/XHR wrappers to prevent count leaks.
' ========================================================================================
Private Function GetProbeJs(ByVal probeType As String) As String
    Dim js As String
    Dim selArray As String: selArray = MergeToJsArray(DEFAULT_IGNORE_SELECTORS, p_userIgnoreSelectors)
    Dim netArray As String: netArray = MergeToJsArray(DEFAULT_IGNORE_NETWORK, p_userIgnoreNetworkPatterns)

    Select Case LCase(probeType)
        Case "inject"
            js = "(function(){"
            
            ' 1. Idempotency Check
            js = js & "  if(window.__vbaSpaProbeInstalled) return true;"
            js = js & "  window.__vbaSpaProbeInstalled = true;"
            
            ' 2. Configuration & Shared State
            js = js & "  var selectors = " & selArray & ";"
            js = js & "  var ignoreNetwork = " & netArray & ";"
            js = js & "  var p = window.__vbaIdleProbe || {"
            js = js & "    inflightFetchCount: 0,"
            js = js & "    inflightXhrCount: 0,"
            js = js & "    lastMutationTs: performance.now()"
            js = js & "  };"
            js = js & "  window.__vbaIdleProbe = p;"
            
            js = js & "  var updateTs = function(){ p.lastMutationTs = performance.now(); };"
            
            ' 3. Node Validation Helper
            ' Evaluates elements or parentElements for Text/Comment nodes to ensure accurate filtering.
            js = js & "  var isIgnoredNode = function(node, sel) {"
            js = js & "    if (!node) return true;"
            js = js & "    var target = (node.nodeType === 1) ? node : node.parentElement;"
            js = js & "    if (!target) return false;"
            js = js & "    return sel.some(function(s) { return target.closest(s) !== null; });"
            js = js & "  };"
            
            ' 4. MutationObserver: Accurate ChildList Monitoring
            js = js & "  try {"
            js = js & "    new MutationObserver(function(mutations){"
            js = js & "      var hasRealChange = mutations.some(function(m){"
            js = js & "        if (m.type !== 'childList') return false;"
            js = js & "        var checkNodes = function(nodes) {"
            js = js & "          return Array.from(nodes).some(function(n){ return !isIgnoredNode(n, selectors); });"
            js = js & "        };"
            js = js & "        return checkNodes(m.addedNodes) || checkNodes(m.removedNodes);"
            js = js & "      });"
            js = js & "      if(hasRealChange) updateTs();"
            js = js & "    }).observe(document, {childList:true, subtree:true, attributes:false});"
            js = js & "  } catch(e) { console.error('BiDi Probe: Observer error', e); }"
            
            ' 5. Network Hook: Fetch Wrapper (Version 3.1 Fix)
            ' Ensures the counter is decremented even if the request fails (catch block).
            js = js & "  if(typeof window.fetch === 'function' && !window.fetch.__vbaWrapped){"
            js = js & "    var origFetch = window.fetch;"
            js = js & "    window.fetch = function(u, o){"
            js = js & "      var url = (typeof u === 'string') ? u : (u.url || '');"
            js = js & "      if(ignoreNetwork.some(function(pat){ return url.indexOf(pat) !== -1; })) return origFetch.apply(this, arguments);"
            js = js & "      p.inflightFetchCount++;"
            js = js & "      var dec = function(){ p.inflightFetchCount--; updateTs(); };"
            js = js & "      return origFetch.apply(this, arguments)"
            js = js & "        .then(function(r){ dec(); return r; })"
            js = js & "        .catch(function(e){ dec(); throw e; });"
            js = js & "    }; window.fetch.__vbaWrapped = true;"
            js = js & "  }"
            
            ' 6. Network Hook: XHR Wrapper (Version 3.1 Fix)
            ' Uses strict flag checking in loadend to prevent double-decrement.
            js = js & "  if(!XMLHttpRequest.prototype.__vbaWrapped){"
            js = js & "    var origOpen = XMLHttpRequest.prototype.open; var origSend = XMLHttpRequest.prototype.send;"
            js = js & "    XMLHttpRequest.prototype.open = function(m, u){ this.__vbaUrl = u; this.__vbaHooked = true; return origOpen.apply(this, arguments); };"
            js = js & "    XMLHttpRequest.prototype.send = function(){"
            js = js & "      var self = this; var url = this.__vbaUrl || '';"
            js = js & "      if(ignoreNetwork.some(function(pat){ return url.indexOf(pat) !== -1; })){ self.__vbaHooked = false; return origSend.apply(this, arguments); }"
            js = js & "      p.inflightXhrCount++;"
            js = js & "      var done = function(){ "
            js = js & "        if(self.__vbaHooked){ p.inflightXhrCount--; self.__vbaHooked = false; updateTs(); }"
            js = js & "        self.removeEventListener('loadend', done); "
            js = js & "      };"
            js = js & "      this.addEventListener('loadend', done); return origSend.apply(this, arguments);"
            js = js & "    }; XMLHttpRequest.prototype.__vbaWrapped = true;"
            js = js & "  }"
            
            js = js & "  return true;"
            js = js & "})();"

        Case "query"
            js = "function(){ var p = window.__vbaIdleProbe; if(!p) return JSON.stringify({active:false});"
            js = js & "  return JSON.stringify({ active: true, now: performance.now(), lastMutationTs: p.lastMutationTs, "
            js = js & "  inflightFetchCount: p.inflightFetchCount, inflightXhrCount: p.inflightXhrCount }); }"
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
' [INTERNAL] MasterWaitEngine (Version 2.4 - Anti-Freeze & Safety Guard)
' SUMMARY: High-precision polling engine for all wait operations.
'          Features:
'          - Integrated validation to prevent negative values (True/-1) from freezing Excel.
'          - Minimum interval enforcement to ensure system responsiveness.
'          - Automatic Excel UI state preservation and restoration.
'          - Re-entrancy guard to prevent recursive loop execution.
' ========================================================================================
Private Function MasterWaitEngine(ByVal mode As waitMode, _
                                  ByVal timeoutMs As Long, _
                                  ByVal intervalMs As Long, _
                                  ByRef ctx As Dictionary) As Boolean
    ' 1. Re-entrancy Guard: Prevent starting a new wait loop if one is already active.
    If m_isWaiting Then Exit Function

    ' 2. Consolidated Validation: Sanitize parameters at the bottleneck.
    ' This prevents logic errors from 'True' or negative inputs provided by the user.
    timeoutMs = SanitizeMs(timeoutMs, "timeoutMs", DEFAULT_IDLE_TIMEOUT_MS)
    
    ' 3. Anti-Freeze Guard: Enforce minimum polling interval (min 50ms).
    ' Setting interval too low causes 100% CPU usage and makes Excel "Not Responding".
    Dim vInterval As Long
    If intervalMs < 50 Then vInterval = 100 Else vInterval = intervalMs

    ' 4. UI State Management: Save current settings for restoration.
    Dim wasInteractive As Boolean: wasInteractive = Application.Interactive
    Dim wasCursor As Long: wasCursor = Application.Cursor
    Dim errNum As Long, errDesc As String, errSrc As String

    On Error GoTo ErrorHandler
    
    ' Enter wait state: Disable interactions to prevent user-driven re-entrancy.
    m_isWaiting = True
    Application.Interactive = False
    Application.Cursor = xlWait
    
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim success As Boolean: success = False
    
    ' 5. Core Polling Loop
    Do
        ' Yield execution: Process WebSocket queue and system events (ThrottledDoEvents).
        ' This ensures Excel remains responsive and the Escape key can still be used.
        WaitHeartbeat vInterval
        
        ' Evaluate the specific condition (URL change, SPA Idle, etc.)
        success = EvaluateWaitCondition(mode, ctx)
        
        ' Exit on success or if the high-precision timeout is reached.
        If success Then Exit Do
        If IsTimeout(startTick, CDbl(timeoutMs)) Then Exit Do
    Loop

CleanExit:
    ' 6. Finally Block: Always restore the Excel UI state before exiting.
    On Error Resume Next
    Application.Interactive = wasInteractive
    Application.Cursor = wasCursor
    m_isWaiting = False
    On Error GoTo 0
    
    ' Re-raise any captured errors during the loop.
    If errNum <> 0 Then
        Err.Raise errNum, "BiDi:MasterWaitEngine:" & errSrc, errDesc
    End If
    
    MasterWaitEngine = success
    Exit Function

ErrorHandler:
    ' Capture error details before restoring the UI.
    errNum = Err.Number: errDesc = Err.Description: errSrc = Err.source
    Resume CleanExit
End Function

' ========================================================================================
' FUNCTION: EvaluateWaitCondition (Version 4.2 - Fixed ByRef and IIf issues)
' ========================================================================================
Private Function EvaluateWaitCondition(ByVal mode As waitMode, ByRef ctx As Dictionary) As Boolean
    Const MAX_RECOVERY_ATTEMPTS As Long = 8
    If ctx Is Nothing Then Err.Raise 520, "BiDi", "Wait context is Nothing."

    Dim targetCtx As String
    If ctx.Exists("contextId") Then targetCtx = CStr(ctx("contextId")) Else targetCtx = p_mainContextId
    If targetCtx = "" Then Err.Raise 521, "BiDi", "Target context is missing."

    Select Case mode
        Case WaitMode_UrlContains
            EvaluateWaitCondition = (InStr(GetCurrentUrl(targetCtx), CStr(ctx("partialUrl"))) > 0)

        Case WaitMode_SpaIdle
            Dim probe As Object: Set probe = QuerySpaProbe(targetCtx)

            If Not probe Is Nothing Then
                ctx("missingProbeCount") = 0
                
                ' Sync heartbeat
                Dim jsNow As Double: jsNow = CDbl(probe("now"))
                If jsNow = m_lastJsNow And (p_pendingRequestCount > 0) Then
                    RecoverSpaProbe targetCtx
                End If
                m_lastJsNow = jsNow

                ' Fix for Error 450: Pass dictionary values via local variables for ByRef parameters
                Dim stableStart As Double: stableStart = CDbl(ctx("stableStart"))
                Dim minStable As Long: minStable = CLng(ctx("minStableMs"))
                Dim currentTicks As Double: currentTicks = GetRawTicks()
                
                EvaluateWaitCondition = EvaluateSpaIdleCondition(probe, minStable, stableStart, currentTicks)
                
                ' Persist updated stableStart back to dictionary
                ctx("stableStart") = stableStart
            Else
                ' Probe missing handling
                Dim bCount As Long: bCount = CLng(ctx("missingProbeCount")) + 1
                ctx("missingProbeCount") = bCount
                ctx("stableStart") = CDbl(GetRawTicks()) ' Safety reset

                If bCount > MAX_RECOVERY_ATTEMPTS Then
                    Err.Raise 516, "BiDi", "SPA Probe failed after " & MAX_RECOVERY_ATTEMPTS & " attempts."
                End If

                If bCount = 1 Then
                    StartSpaIdleProbe targetCtx
                ElseIf bCount >= 3 Then
                    RecoverSpaProbe targetCtx
                End If
                EvaluateWaitCondition = False
            End If
    End Select
End Function

' ========================================================================================
' [INTERNAL] ExecuteBaseAction (Updated with Centralized Validation)
' SUMMARY: The central runner for all browser interactions.
'          Handles retries for stale elements, JIT probe injection, and
'          post-action synchronization with unified parameter validation.
' ========================================================================================
Private Function ExecuteBaseAction(ByVal contextId As String, _
                                   ByVal jsKey As String, _
                                   ByVal args As Collection, _
                                   Optional ByVal waitNetwork As Boolean = True, _
                                   Optional ByVal minStable As Long = DEFAULT_IDLE_STABLE_MS, _
                                   Optional ByVal awaitPromise As Boolean = True, _
                                   Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                   Optional ByVal maxTimeout As Long = DEFAULT_IDLE_TIMEOUT_MS) As String
    Dim retryCount As Long
    Dim strRes As String
    Dim targetCtx As String
    Dim errNum As Long, errDesc As String, errSrc As String

    ' 1. Consolidated Validation: Guard all timing parameters at the entry point
    ' This effectively blocks 'True' (-1) or invalid values from entering the logic.
    Dim vStable As Long:  vStable = SanitizeMs(minStable, "minStableMs", DEFAULT_IDLE_STABLE_MS)
    Dim vTimeout As Long: vTimeout = SanitizeMs(maxTimeout, "maxTimeoutMs", DEFAULT_IDLE_TIMEOUT_MS)

    ' 2. Resolve the target browsing context
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId()
    End If

    ' 3. Execution loop with transient error recovery (e.g., Stale Element, Frame Reloads)
    For retryCount = 0 To maxRetries
        On Error Resume Next
        
        ' --- [JIT PROBE PREPARATION] ---
        ' Ensure the SPA monitoring probe is installed BEFORE execution
        ' to capture rapid network requests triggered by the action.
        If waitNetwork Then PrepareProbeForAction targetCtx
        
        ' Refresh context identifiers on retry attempts to handle reloaded environments
        If retryCount > 0 Then
            UpdateMainContextId True
            WaitHeartbeat 300
        End If
        
        ' 4. Execution: Dispatch the JavaScript snippet via WebDriver BiDi
        strRes = CallScript(GetActionJs(jsKey), args, targetCtx, awaitPromise)
        
        If Err.Number = 0 Then
            ' SUCCESS: Finalize state with post-action synchronization if requested
            ' Uses the sanitized vStable and vTimeout values for the wait loop.
            If waitNetwork Then SyncAfterAction targetCtx, vStable, vTimeout
            
            ExecuteBaseAction = strRes
            Exit Function
        Else
            ' FAILURE: Capture error details for recovery analysis
            errNum = Err.Number: errDesc = Err.Description: errSrc = Err.source
            
            ' Check if the error is recoverable (Protocol hitch or Functional action failure)
            If (errNum = ERR_ACTION_FAILED Or errNum = ERR_BIDI_PROTOCOL) And retryCount < maxRetries Then
                Err.Clear
                If DEBUG_MODE Then
                    Debug.Print "[BiDi Action Retry] " & jsKey & " (Attempt " & retryCount + 1 & ")"
                End If
                
                ' Briefly yield to let the browser stabilize before retrying
                WaitHeartbeat 500
            Else
                ' EXHAUSTED: Re-raise the error once retries are spent or error is fatal
                On Error GoTo 0
                Err.Raise errNum, errSrc, errDesc
            End If
        End If
    Next retryCount
End Function

' ========================================================================================
' Synchronizes the state after an interaction by checking network and DOM stability.
' ========================================================================================
Private Sub SyncAfterAction(ByVal contextId As String, ByVal minStableMs As Long, ByVal maxTimeoutMs As Long)
    ' Refresh the context tree to ensure we are targeting the current document
    UpdateMainContextId True
    ' Process pending messages to clear the queue
    ProcessRecorderQueue
    ' Wait for the page to reach an idle state
    WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, contextId
End Sub

' ========================================================================================
' Private Function: ProcessResponse
' Summary: Centralized validator for all BiDi/CDP responses.
'          Checks 3 layers: Protocol errors, JS exceptions, and Functional status.
' ========================================================================================
Private Function ProcessResponse(ByVal jsonRaw As String, ByVal methodName As String) As Object
    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonRaw)
    On Error GoTo 0
    
    If json Is Nothing Then
        Err.Raise vbObjectError + 512, "BiDi:" & methodName, "Response is not valid JSON."
    End If

    ' --- Layer 1: Protocol Error Detection (BiDi/CDP level) ---
    If json.Exists("error") Then
        Dim errType As String: errType = json("error")
        Dim errMsg As String: errMsg = json("message")
        ' Skip if it's a known ignorable state (like frame detached during nav)
        If Not IsIgnorableError(errType, errMsg) Then
            Err.Raise ERR_BIDI_PROTOCOL, "BiDi:Protocol:" & methodName, "[" & errType & "] " & errMsg
        End If
    End If

    ' --- Layer 2: JS Execution Error Detection (Runtime Exception) ---
    Dim excDetails As Object
    Set excDetails = GetPathObject(json, "result.exceptionDetails")
    
    If Not excDetails Is Nothing Then
        Debug.Print "--- BIDI JS ERROR REPORT | Method: " & methodName & " ---"
        Debug.Print FormatJavaScriptError(excDetails) ' Uses existing formatter
        Err.Raise ERR_JS_EXCEPTION, "BiDi:JSException", _
                  "JS Error in " & methodName & ": " & GetPathString(json, "result.exceptionDetails.text")
    End If

    ' --- Layer 3: Functional Status Detection (Custom JS script status) ---
    ' This handles our JS snippets' return: {status: "error", message: "..."}
    Dim funcStatus As String
    funcStatus = GetPathString(json, "result.result.value.status")
    
    If funcStatus = "error" Then
        Dim actionMsg As String
        actionMsg = GetPathString(json, "result.result.value.message")
        Err.Raise ERR_ACTION_FAILED, "BiDi:ActionFailed:" & methodName, actionMsg
    End If

    ' If we reached here, the response is valid and successful.
    Set ProcessResponse = json
End Function

' ========================================================================================
' [SECTION] JSON Path Utilities (Fixed for Error 450)
' ========================================================================================

' Wrapper to return an Object (Dictionary/Collection) from a path.
Private Function GetPathObject(ByVal root As Object, ByVal path As String) As Object
    On Error Resume Next
    ' In VBA, Set is required when the returned value is an Object.
    ' If the path points to a non-object, this assignment will safely fail due to Resume Next.
    Set GetPathObject = InternalTraverseFromObject(root, path)
    On Error GoTo 0
End Function

' Wrapper to return a String from a path.
Private Function GetPathString(ByVal root As Object, ByVal path As String) As String
    Dim v As Variant
    v = InternalTraverseFromObject(root, path)
    
    If IsObject(v) Then
        GetPathString = ""
    ElseIf IsNull(v) Then
        GetPathString = ""
    Else
        GetPathString = CStr(v)
    End If
End Function

' Wrapper to return a Variant from a path (Supports both Objects and Primitives).
Private Function GetPathValue(ByVal root As Object, ByVal path As String) As Variant
    Dim temp As Variant
    ' Temporarily check if the return value is an object.
    If IsObject(InternalTraverseFromObject(root, path)) Then
        Set GetPathValue = InternalTraverseFromObject(root, path)
    Else
        GetPathValue = InternalTraverseFromObject(root, path)
    End If
End Function

' ========================================================================================
' FUNCTION: InternalJsonTraverse
' SUMMARY: Entry point to parse a JSON string and traverse to a specific path.
'          Handles the conversion from String to Object before path resolution.
' ========================================================================================
Private Function InternalJsonTraverse(ByVal jsonStr As String, ByVal path As String) As Variant
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
        If IsObject(InternalTraverseFromObject(json, path)) Then
            ' Use SET for Dictionary or Collection objects
            Set InternalJsonTraverse = InternalTraverseFromObject(json, path)
        Else
            ' Use standard assignment for String, Double, Boolean, etc.
            InternalJsonTraverse = InternalTraverseFromObject(json, path)
        End If
    End If
End Function

' ========================================================================================
' FUNCTION: InternalTraverseFromObject
' SUMMARY: Recursively traverses a VBA JSON object structure using dot-notation.
'          Supports both Dictionary keys and Collection indices (0-based for JSON).
' ========================================================================================
Private Function InternalTraverseFromObject(ByVal root As Object, ByVal path As String) As Variant
    On Error Resume Next
    If root Is Nothing Or path = "" Then Exit Function
    
    Dim keys() As String: keys = Split(path, ".")
    Dim current As Object: Set current = root
    Dim i As Long, idx As Long
    
    ' 1. Traverse the hierarchy until the second-to-last segment.
    '    This loop resolves all parent nodes in the path.
    For i = 0 To UBound(keys) - 1
        If TypeName(current) = "Dictionary" Then
            If Not current.Exists(keys(i)) Then Exit Function
            Set current = current(keys(i))
        ElseIf TypeName(current) = "Collection" And IsNumeric(keys(i)) Then
            ' Convert 0-based JSON index to 1-based VBA Collection index.
            idx = CLng(keys(i)) + 1
            If idx < 1 Or idx > current.Count Then Exit Function
            Set current = current(idx)
        Else
            ' Invalid path or non-traversable object encountered.
            Exit Function
        End If
        If current Is Nothing Then Exit Function
    Next i
    
    ' 2. Resolve the final leaf node.
    Dim lastKey As String: lastKey = keys(UBound(keys))
    Dim resultValue As Variant
    
    If TypeName(current) = "Dictionary" Then
        If current.Exists(lastKey) Then
            ' Defensive check for Object vs Primitive assignment to avoid Error 450.
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
                                     ByVal timeoutMs As Long, _
                                     ByVal methodName As String) As String
    Dim sid As String
    Dim targetCtx As String
    
    ' 1. Consolidated Validation: Prevent logic errors from 'True' or negative inputs.
    ' Even if a public method receives 'True', it is sanitized here into a safe Long.
    Dim vTimeout As Long
    vTimeout = SanitizeMs(timeoutMs, "searchTimeoutMs", 5000)
    
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
    On Error Resume Next
    ' QuerySpaProbe returns Nothing if the probe state (__vbaIdleProbe) is missing
    If QuerySpaProbe(contextId) Is Nothing Then
        ' Inject the JavaScript monitoring hooks
        StartSpaIdleProbe contextId
        ' Yield briefly to allow the JS engine to evaluate the injection
        WaitHeartbeat 50
    End If
End Sub

' ========================================================================================
' [INTERNAL] SanitizeMs
' SUMMARY: Central gatekeeper to validate millisecond arguments.
'          Strictly rejects negative values (preventing True/-1 errors)
'          and caps values to prevent permanent Excel hangs.
' ========================================================================================
Private Function SanitizeMs(ByVal ms As Long, ByVal argName As String, ByVal defaultMs As Long) As Long
    ' 1. Logic check: Rejects negative values (misused Booleans)
    If ms < 0 Then
        Err.Raise 5, "BiDi:InvalidParameter", _
            "Argument '" & argName & "' cannot be negative (" & ms & ")." & vbCrLf & _
            "Ensure you are passing milliseconds, not a Boolean (True)."
    End If

    ' 2. Range check and default fallback
    If ms = 0 Then
        SanitizeMs = 0 ' Explicit no-wait
    ElseIf ms > 600000 Then
        SanitizeMs = 600000 ' Circuit breaker: 10-minute cap
    Else
        SanitizeMs = ms
    End If
End Function
