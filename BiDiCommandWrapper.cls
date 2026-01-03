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
' Summary: High-Level Wrapper for Selenium WebDriver BiDi with CDP Integration
'          Features:
'           - Resource Blocking via CDP Tunneling (Fetch.enable / Network.setBlockedURLs)
'           - Smart Frame Recovery (Auto-handles frame reloads/stale contexts)
'           - Centralized Event Dispatching (Fixes event conflicts)
'           - Anti-Freeze Alert Handling
'           - AI-Ready Discovery Recorder
'           - Advanced Wait Strategies (Network Idle, DOM Stability, SPA Idle)
' Dependencies: BiDiSocketCommunicator, WebJsonConverter, Microsoft Scripting Runtime
' ========================================================================================
Private socket_ As BiDiSocketCommunicator
Private Const DEBUG_MODE As Boolean = True

' Recovery & Stability
Private m_isRecovering As Boolean
Private m_lastDoEventsTime As Currency
Private Const MAX_CONTEXT_RECOVERY_RETRIES As Integer = 2

' [Configuration Constants]
Private Const DEFAULT_NAV_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_IDLE_STABLE_MS As Long = 500
Private Const DEFAULT_IDLE_TIMEOUT_MS As Long = 10000
Private Const DEFAULT_STALE_RETRIES As Long = 2
Private Const FRAME_SEARCH_TIMEOUT_MS As Long = 3000
Private Const SHORT_WAIT_MS As Long = 500
Private Const MAX_LOG_SIZE As Long = 500
Private Const PROBE_VERSION As String = "1.3"

' --- Noise Filtering State ---
' Default filters are built-in to handle common noise (Analytics, Hidden inputs, etc.)
Private Const DEFAULT_IGNORE_SELECTORS As String = "script, link, input[type='hidden'], .vba-ignore"
Private Const DEFAULT_IGNORE_NETWORK As String = "google-analytics, doubleclick, /collect, /beacon, .woff2, favicon.ico"

' User-defined filters set from Standard Modules
Private p_userIgnoreSelectors As Collection
Private p_userIgnoreNetworkPatterns As Collection

' File System Object for Download Monitoring
Private fso_ As New Scripting.FileSystemObject
Private p_currentDownloadPath As String

' --- State flags for persistent registrations ---
Private m_isProbeRegistered As Boolean  ' Flag for CDP Page.addScriptToEvaluateOnNewDocument
Private m_isSubscribed As Boolean       ' Flag for session-wide BiDi event subscriptions
Private m_isProcessingQueue As Boolean

' Internal State Variables
Private p_nextRequestId As Long
Private p_mainContextId As String
Private p_mainRealmId As String
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
Private Const REALM_DEFAULT_TARGET As String = "context"
Private Const SPA_LOAD_GRACE_MS As Long = 1500

' ========================================================================================
' API Declarations (Replaced GetTickCount with QueryPerformanceCounter)
' ========================================================================================
Private Declare PtrSafe Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' Public Enumerations for better Intellisense and Type Safety
Public Enum BiDiMatchBy
    MatchByUrl = 0
    MatchByTitle = 1
End Enum

Public Enum BiDiContextType
    ContextTypeTab = 0
    ContextTypeWindow = 1
End Enum

' ========================================================================================
' Initialization: Setup WebSocket and validate High-Resolution Timer frequency.
' ========================================================================================
Private Sub Class_Initialize()
    Dim socket As New BiDiSocketCommunicator
    Set socket_ = socket
    
    p_nextRequestId = 1
    
    ' --- High-Resolution Timer Validation ---
    ' Verify if the hardware support QueryPerformanceCounter.
    ' If it fails (returns 0), set p_timerFreq to -1 to trigger the VBA Timer() fallback.
    If QueryPerformanceFrequency(p_timerFreq) = 0 Or p_timerFreq = 0 Then
        p_timerFreq = -1 ' Flag for fallback mode
        Debug.Print "BiDi Warning: High-resolution timer unavailable. Using VBA.Timer() fallback."
    End If
End Sub

Public Sub ConnectTo(ByVal WebSocketUrl As String)
    socket_.AttemptAutoConnect WebSocketUrl
End Sub

' ========================================================================================
' [FIXED] Class_Terminate with robust resource release
' ========================================================================================
Private Sub Class_Terminate()
    ' Ensure recorder logs are cleared from memory
    If Not p_recorderLogs Is Nothing Then
        On Error Resume Next
        Dim i As Long
        For i = p_recorderLogs.Count To 1 Step -1
            p_recorderLogs.Remove i
        Next i
        Set p_recorderLogs = Nothing
        On Error GoTo 0
    End If
    
    If Not fso_ Is Nothing Then Set fso_ = Nothing
    Set socket_ = Nothing
    
    If DEBUG_MODE Then Debug.Print "BiDiCommandWrapper: Class Terminated. All resources explicitly released."
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
' Frame / Iframe Helper (Enhanced with Wait Logic)
' ========================================================================================
' ========================================================================================
' GetIframeContextIdByUrl: Searches for a frame context by partial URL with timeout.
' ========================================================================================
Public Function GetIframeContextIdByUrl(ByVal partialUrl As String, Optional ByVal timeoutMs As Long = 0) As String
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim ctxId As String

    Do
        ' Get the tree and search recursively
        Dim strRes As String: strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Dim contexts As Collection: Set contexts = GetJsonValue(strRes, "result.contexts")
        
        If Not contexts Is Nothing Then
            ctxId = RecursiveFrameSearch(contexts, partialUrl)
            If ctxId <> "" Then
                GetIframeContextIdByUrl = ctxId
                Exit Function
            End If
        End If

        ' Immediate exit if no timeout is specified
        If timeoutMs <= 0 Then Exit Do

        ' Heartbeat and timeout check
        WaitHeartbeat SHORT_WAIT_MS
        If IsTimeout(startTick, CDbl(timeoutMs)) Then Exit Do
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
' Frame / Iframe Helper (Enhanced with Robust Loop Retry)
' ========================================================================================
Public Function ExecuteSmartFrameClickAndWait(ByVal framePartialUrl As String, _
                                              ByVal xpath As String, _
                                              Optional ByVal waitNetworkIdle As Boolean = True, _
                                              Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String
    Dim ctxId As String
    Dim res As String
    Dim retryCount As Integer
    Const MAX_FRAME_RETRIES As Integer = 3

    ctxId = Me.GetIframeContextIdByUrl(framePartialUrl, 2000)

    For retryCount = 0 To MAX_FRAME_RETRIES
        If ctxId = "" Then
            ctxId = Me.GetIframeContextIdByUrl(framePartialUrl, FRAME_SEARCH_TIMEOUT_MS + (retryCount * 500))
            If ctxId = "" Then
                If retryCount = MAX_FRAME_RETRIES Then
                    Err.Raise vbObjectError + 556, "ExecuteSmartFrameClickAndWait", "Target frame not found: " & framePartialUrl
                Else
                    sleep 500
                    GoTo NextLoop
                End If
            End If
        End If

        On Error Resume Next
        res = Me.ExecuteClickAndWaitByXPath(xpath, waitNetworkIdle, minStableMs, ctxId)
        If Err.Number <> 0 Then
            Dim errDesc As String: errDesc = Err.Description
            Dim isDeadFrame As Boolean
            isDeadFrame = (InStr(errDesc, "no such frame") > 0) Or (InStr(errDesc, "context destroyed") > 0) Or (InStr(errDesc, "stale") > 0)
            On Error GoTo 0

            If isDeadFrame Then
                Debug.Print "BiDi Info: Frame reload detected (" & framePartialUrl & "). Retry " & retryCount & "..."
                ctxId = ""
                sleep 500
                DoEvents
            Else
                Err.Raise Err.Number, Err.source, errDesc
            End If
        Else
            On Error GoTo 0
            ExecuteSmartFrameClickAndWait = res
            Exit Function
        End If

NextLoop:
    Next retryCount

    ExecuteSmartFrameClickAndWait = res
End Function

Public Function ExecuteSmartFrameShadowClickAndWait(ByVal framePartialUrl As String, _
                                                    ByVal selectorsArray As Variant, _
                                                    Optional ByVal waitNetworkIdle As Boolean = True, _
                                                    Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String
    Dim ctxId As String
    Dim res As String
    Dim retryCount As Integer
    Const MAX_FRAME_RETRIES As Integer = 3

    ctxId = Me.GetIframeContextIdByUrl(framePartialUrl, 2000)

    For retryCount = 0 To MAX_FRAME_RETRIES
        If ctxId = "" Then
            ctxId = Me.GetIframeContextIdByUrl(framePartialUrl, FRAME_SEARCH_TIMEOUT_MS + (retryCount * 500))
            If ctxId = "" Then
                If retryCount = MAX_FRAME_RETRIES Then
                    ExecuteSmartFrameShadowClickAndWait = "{""error"":""frame not found after retries""}"
                    Exit Function
                Else
                    sleep 500
                    GoTo NextLoop
                End If
            End If
        End If

        res = Me.ExecuteShadowClickAndWait(selectorsArray, waitNetworkIdle, minStableMs, , ctxId)

        Dim isDeadFrame As Boolean
        isDeadFrame = (InStr(res, "no such frame") > 0) Or (InStr(res, "context destroyed") > 0) Or (InStr(res, "stale") > 0) Or (InStr(res, "execution context") > 0) Or (InStr(res, "VBA_Runtime_Error") > 0)

        If isDeadFrame Then
            Debug.Print "BiDi Info: Frame reload detected during Shadow Click. Retry " & retryCount & "..."
            ctxId = ""
            sleep 500
            DoEvents
        Else
            ExecuteSmartFrameShadowClickAndWait = res
            Exit Function
        End If

NextLoop:
    Next retryCount

    ExecuteSmartFrameShadowClickAndWait = res
End Function

' ========================================================================================
' ExecuteFindWindowContextId: Locates a window context using Enums for match type.
' ========================================================================================
Public Function ExecuteFindWindowContextId(ByVal matchType As BiDiMatchBy, _
                                           ByVal matchValue As String, _
                                           Optional ByVal timeoutMs As Long = 5000, _
                                           Optional ByVal waitForComplete As Boolean = True) As String
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim strRes As String, res As Variant
    Dim ctxList As Collection, ctx As Dictionary
    
    Do
        strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Set ctxList = GetJsonValue(strRes, "result.contexts")

        If Not ctxList Is Nothing Then
            For Each ctx In ctxList
                Dim isMatch As Boolean
                ' Logic: Use Enum comparison instead of LCase string comparison
                If matchType = MatchByUrl Then
                    isMatch = (InStr(ctx("url"), matchValue) > 0)
                ElseIf matchType = MatchByTitle Then
                    isMatch = (InStr(ExecuteGetTitleByContextId(ctx("context")), matchValue) > 0)
                End If

                If isMatch Then
                    If Not waitForComplete Then
                        ExecuteFindWindowContextId = ctx("context"): Exit Function
                    Else
                        res = GetJsonValue(CallScript(GetActionJs("ready_state"), , ctx("context")), "result.result.value")
                        If CStr(res) = "complete" Then
                            ExecuteFindWindowContextId = ctx("context"): Exit Function
                        End If
                    End If
                End If
            Next ctx
        End If

        WaitHeartbeat SHORT_WAIT_MS
        If IsTimeout(startTick, CDbl(timeoutMs)) Then Exit Do
    Loop
End Function

' ========================================================================================
' ExecuteCreateContext: Creates a new tab or window using Enum for type.
' ========================================================================================
Public Function ExecuteCreateContext(Optional ByVal createType As BiDiContextType = ContextTypeTab) As String
    Dim typeStr As String
    typeStr = IIf(createType = ContextTypeWindow, "window", "tab")

    Dim params As New Dictionary: params.Add "type", typeStr
    Dim strRes As String: strRes = ExecuteBiDiCommand("browsingContext.create", params)
    
    ExecuteCreateContext = CStr(GetJsonValue(strRes, "result.context"))
End Function
Public Function ExecuteCloseContext(ByVal contextId As String) As String
    Dim params As New Dictionary
    params.Add "context", contextId
    ExecuteCloseContext = ExecuteBiDiCommand("browsingContext.close", params)
    If contextId = p_mainContextId Then ClearState
End Function

Private Function ExecuteGetTitleByContextId(ByVal contextId As String) As String
    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    If realmId = "" Then Exit Function

    Dim params As New Dictionary, target As New Dictionary
    Dim js As String: js = "function() { return document.title || ''; }"

    target.Add "realm", realmId
    params.Add "functionDeclaration", js
    params.Add "arguments", New Collection
    params.Add "target", target
    params.Add "awaitPromise", False

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.callFunction", params)

    ' Use helper to extract the title
    ExecuteGetTitleByContextId = CStr(GetJsonValue(strRes, "result.result.value"))
End Function

' ========================================================================================
' ExecuteClickAndHandleAlertByXPath: Clicks an element that triggers an alert and handles it.
' Uses a non-blocking click strategy to prevent VBA from hanging.
' ========================================================================================
Public Function ExecuteClickAndHandleAlertByXPath(ByVal xpath As String, _
                                                  Optional ByVal Accept As Boolean = True, _
                                                  Optional ByVal contextId As String = "") As String
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, targetCtx)
    
    If sharedId = "" Then
        ExecuteClickAndHandleAlertByXPath = "{""error"":""element not found""}"
        Exit Function
    End If

    ' 1. Prepare Arguments
    Dim args As New Collection, arg1 As New Dictionary
    arg1.Add "sharedId", sharedId: args.Add arg1

    ' 2. Execute non-blocking click via CallScript
    ' Note: awaitPromise is False because we return immediately after setTimeout
    CallScript GetActionJs("click_alert"), args, targetCtx, False

    ' 3. Brief wait for the alert to appear in the browser's UI thread
    WaitHeartbeat 500

    ' 4. Handle the alert (Accept or Dismiss)
    ExecuteClickAndHandleAlertByXPath = ExecuteHandleUserPrompt(Accept, "", targetCtx)
End Function

Public Function ExecuteHandleUserPrompt(ByVal Accept As Boolean, Optional ByVal userText As String = "", Optional ByVal contextId As String = "") As String
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId()

    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "accept", Accept
    If Len(userText) > 0 Then params.Add "userText", userText

    On Error Resume Next
    ExecuteHandleUserPrompt = ExecuteBiDiCommand("browsingContext.handleUserPrompt", params)
    On Error GoTo 0
End Function

' ========================================================================================
' ExecuteInputClick: Performs a physical mouse click using BiDi's input.performActions.
' Updated: Explicitly specifies pointerType as "mouse" for improved compatibility.
' ========================================================================================
Public Function ExecuteInputClick(ByVal xpath As String, Optional ByVal contextId As String = "") As String
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())

    ' 1. Resolve the element to a sharedId
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, targetCtx)
    If sharedId = "" Then
        ExecuteInputClick = "{""error"":""element not found""}"
        Exit Function
    End If

    ' 2. Construct the BiDi parameters for input.performActions
    Dim params As New Dictionary: params.Add "context", targetCtx

    ' Create the Action Sequence
    Dim actions As New Collection, actionSeq As New Dictionary
    actionSeq.Add "type", "pointer"
    actionSeq.Add "id", "mouse_input_vba"

    ' --- SPECIFY POINTER TYPE ---
    ' Explicitly set pointerType to 'mouse' to ensure correct event firing in all environments.
    Dim seqParams As New Dictionary
    seqParams.Add "pointerType", "mouse"
    actionSeq.Add "parameters", seqParams

    ' Define the steps (Move -> Down -> Up)
    Dim steps As New Collection
    Dim stepMove As New Dictionary, stepDown As New Dictionary, stepUp As New Dictionary

    ' Build Origin Object for the target element
    Dim originObj As New Dictionary, elementObj As New Dictionary
    elementObj.Add "sharedId", sharedId
    originObj.Add "type", "element"
    originObj.Add "element", elementObj

    ' Step 1: Pointer Move to the element's origin
    stepMove.Add "type", "pointerMove"
    stepMove.Add "origin", originObj
    stepMove.Add "x", 0: stepMove.Add "y", 0

    ' Step 2: Pointer Down (Left button = 0)
    stepDown.Add "type", "pointerDown"
    stepDown.Add "button", 0

    ' Step 3: Pointer Up (Left button = 0)
    stepUp.Add "type", "pointerUp"
    stepUp.Add "button", 0

    steps.Add stepMove: steps.Add stepDown: steps.Add stepUp
    actionSeq.Add "actions", steps
    actions.Add actionSeq
    params.Add "actions", actions

    ' 3. Execute the protocol command
    ExecuteInputClick = ExecuteBiDiCommand("input.performActions", params)
End Function

' ========================================================================================
' ExecuteBiDiWaitUntilUrlContains: Polls the URL until it matches the partial string.
' ========================================================================================
Public Function ExecuteBiDiWaitUntilUrlContains(ByVal partialUrl As String, _
                                                Optional ByVal timeoutSec As Long = 300, _
                                                Optional ByVal contextId As String = "") As Boolean
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim timeoutMs As Double: timeoutMs = CDbl(timeoutSec) * 1000
    
    ' Subscribe to relevant navigation events
    Me.ExecuteSessionSubscribe Array("network.responseCompleted", "browsingContext.load", "browsingContext.fragmentNavigated")

    Do
        ' Check current URL
        Dim currentUrl As String: currentUrl = GetCurrentUrl(contextId)
        
        If Len(currentUrl) > 0 And InStr(currentUrl, partialUrl) > 0 Then
            ExecuteBiDiWaitUntilUrlContains = True
            Exit Function
        End If

        ' Centralized background processing
        WaitHeartbeat 100 ' 100ms interval for URL polling
        
        ' Check for timeout using the new helper
        If IsTimeout(startTick, timeoutMs) Then
            Debug.Print "BiDi Warning: URL wait timed out."
            Exit Do
        End If
    Loop
End Function

' ========================================================================================
' ExecuteNavigateAndGetStatus: Navigates and waits for the document to reach interactive state.
' ========================================================================================
Public Function ExecuteNavigateAndGetStatus(ByVal targetUrl As String, _
                                            Optional ByVal waitNetworkIdle As Boolean = True, _
                                            Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                            Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                            Optional ByVal contextId As String = "", _
                                            Optional ByVal waitForInteractive As Boolean = True) As Long
    
    p_pendingRequestCount = 0
    ClearState
    
    ' Subscribe to lifecycle events
    ExecuteSessionSubscribe Array("network.beforeRequestSent", "network.responseCompleted", _
                                  "browsingContext.load", "browsingContext.domContentLoaded")

    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId(True))
    
    ' Trigger navigation (wait: none for manual polling)
    Dim p As New Dictionary: p.Add "context", targetCtx: p.Add "url", targetUrl: p.Add "wait", "none"
    ExecuteBiDiCommand "browsingContext.navigate", p

    Dim startTick As Currency: startTick = GetRawTicks()
    Dim isLoadComplete As Boolean

    ' Monitoring loop
    Do While Not IsTimeout(startTick, CDbl(DEFAULT_NAV_TIMEOUT_MS))
        Dim state As String: state = ExecuteGetReadyState(targetCtx)
        If state = "complete" Or (waitForInteractive And state = "interactive") Then
            isLoadComplete = True: Exit Do
        End If
        
        WaitHeartbeat 100
    Loop
    
    If waitNetworkIdle And isLoadComplete Then
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx
    End If

    ExecuteNavigateAndGetStatus = 200 ' Success fallback
End Function

Private Function ExecuteGetReadyState(ByVal ctxId As String) As String
    Dim js As String: js = "function(){ return document.readyState; }"
    Dim p As New Dictionary, t As New Dictionary
    t.Add "context", ctxId
    p.Add "functionDeclaration", js
    p.Add "target", t
    p.Add "awaitPromise", False

    Dim ret As String
    ret = ExecuteBiDiCommand("script.callFunction", p)
    
    ' Extract the readyState string (e.g., "complete", "interactive")
    ExecuteGetReadyState = CStr(GetJsonValue(ret, "result.result.value"))
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
' Visibility & State Check
' ========================================================================================
Public Function ExecuteIsElementVisible(ByVal xpath As String, Optional ByVal contextId As String = "") As Boolean
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, contextId)
    If sharedId = "" Then Exit Function

    ' Prepare only the necessary arguments
    Dim args As New Collection, arg1 As New Dictionary
    arg1.Add "sharedId", sharedId: args.Add arg1

    ' Execute using centralized helpers
    Dim strRes As String
    strRes = CallScript(GetActionJs("visibility_check"), args, contextId)
    
    ExecuteIsElementVisible = CBool(GetJsonValue(strRes, "result.result.value"))
End Function

' ========================================================================================
' ExecuteWaitForSelectOptions: Waits for a select element to populate with options.
' ========================================================================================
Public Function ExecuteWaitForSelectOptions(ByVal xpath As String, _
                                            Optional ByVal minCount As Long = 1, _
                                            Optional ByVal timeoutMs As Long = 5000, _
                                            Optional ByVal contextId As String = "") As Boolean
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, contextId)
    If sharedId = "" Then Exit Function

    ' 1. Prepare Arguments: [element, minCount, timeoutMs]
    Dim args As New Collection, arg1 As New Dictionary, arg2 As New Dictionary, arg3 As New Dictionary
    arg1.Add "sharedId", sharedId
    arg2.Add "type", "number": arg2.Add "value", minCount
    arg3.Add "type", "number": arg3.Add "value", timeoutMs
    args.Add arg1: args.Add arg2: args.Add arg3

    ' 2. Execute via CallScript (awaitPromise=True is essential for the Promise in JS)
    Dim res As String
    res = CallScript(GetActionJs("wait_options"), args, contextId, True)

    ' 3. Extract and return the boolean result
    ExecuteWaitForSelectOptions = CBool(GetJsonValue(res, "result.result.value"))
End Function

' ========================================================================================
' ExecuteGetSelectOptions: Retrieves all options from a select element as a Collection.
' ========================================================================================
Public Function ExecuteGetSelectOptions(ByVal xpath As String, Optional ByVal contextId As String = "") As Collection
    Dim resultColl As New Collection
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, contextId)
    If sharedId = "" Then Set ExecuteGetSelectOptions = resultColl: Exit Function

    ' 1. Prepare Argument: [element]
    Dim args As New Collection, arg1 As New Dictionary
    arg1.Add "sharedId", sharedId: args.Add arg1

    ' 2. Execute and get JSON string of options
    Dim res As String
    res = CallScript(GetActionJs("get_options"), args, contextId)
    
    ' 3. Parse the inner JSON string returned by JS
    Dim innerJsonStr As String: innerJsonStr = CStr(GetJsonValue(res, "result.result.value"))
    
    On Error Resume Next
    If innerJsonStr <> "" Then Set resultColl = WebJsonConverter.ParseJson(innerJsonStr)
    On Error GoTo 0
    
    Set ExecuteGetSelectOptions = resultColl
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
    
    If InStr(ExecuteSessionSubscribe, """success""") > 0 Then
        m_isSubscribed = True
    End If
End Function

' ========================================================================================
' ExecuteBrowsingContextNavigate: Triggers navigation and proactively refreshes IDs.
' Optimization: Refreshes Context and Realm IDs immediately after the navigate command
' to eliminate the need for error-triggered recovery loops.
' ========================================================================================
Public Function ExecuteBrowsingContextNavigate(ByVal url As String, _
                                               Optional ByVal waitMode As String = "complete", _
                                               Optional ByVal contextId As String = "") As String
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId(True)

    p_pendingRequestCount = 0
    ClearState
    
    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "url", url
    params.Add "wait", waitMode

    Dim strRes As String
    ' 1. Send navigation command (wait="none" allowed for speed)
    strRes = ExecuteBiDiCommand("browsingContext.navigate", params)

    ' 2. [PROACTIVE REFRESH]
    ' If navigation started successfully, refresh IDs immediately while the page is loading.
    ' This ensures subsequent commands hit the NEW realm without causing Context Lost errors.
    If InStr(strRes, """type"":""success""") > 0 Then
        ' Proactively wait for URL stability and refresh Context ID
        UpdateMainContextId True
        
        ' Pre-fetch and cache the new Realm ID for the new document
        GetScriptRealmId p_mainContextId, True
        
        ' Finalize with SPA idle detection if required
        WaitForNetworkIdleSpa DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
    End If
    
    ExecuteBrowsingContextNavigate = strRes
End Function

' ========================================================================================
' ExecuteActionWithStaleRetry: Runs JS on an element with automatic recovery from stale references.
' Now uses CallScript and GetJsonValue for a much cleaner implementation.
' ========================================================================================
Private Function ExecuteActionWithStaleRetry(ByVal xpath As String, _
                                             ByVal jsDeclaration As String, _
                                             Optional ByVal secondArgValue As Variant = Empty, _
                                             Optional ByVal awaitPromise As Boolean = True, _
                                             Optional ByVal contextId As String = "", _
                                             Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    Dim retryCount As Integer
    Dim sharedId As String, strRes As String
    Dim args As Collection, arg1 As Dictionary, arg2 As Dictionary
    Dim isStale As Boolean

    ' 1. Initial attempt to find the element
    sharedId = GetSharedIdFromXPath(xpath, contextId)

    For retryCount = 0 To maxRetries
        ' 2. Re-verify sharedId on retry if it was lost
        If sharedId = "" And retryCount > 0 Then sharedId = GetSharedIdFromXPath(xpath, contextId)

        If sharedId = "" Then
            strRes = "{""error"":""element not found""}"
        Else
            ' 3. Prepare arguments collection
            Set args = New Collection: Set arg1 = New Dictionary
            arg1.Add "sharedId", sharedId: args.Add arg1

            ' Add optional second argument (e.g., input value)
            If Not IsEmpty(secondArgValue) Then
                Set arg2 = New Dictionary
                arg2.Add "type", "string": arg2.Add "value", CStr(secondArgValue)
                args.Add arg2
            End If

            ' 4. Execute using the CallScript helper
            On Error Resume Next
            strRes = CallScript(jsDeclaration, args, contextId, awaitPromise)
            If Err.Number <> 0 Then
                strRes = "{""error"":""vba_error"", ""message"":""" & Err.Description & """}"
                Err.Clear
            End If
            On Error GoTo 0
        End If

        ' 5. Determine if we should retry (Stale or Context loss)
        ' Use GetJsonValue to check for protocol-level error or JS exception
        Dim errType As String: errType = CStr(GetJsonValue(strRes, "error"))
        Dim excText As String: excText = CStr(GetJsonValue(strRes, "result.exceptionDetails.text"))
        
        isStale = IsIgnorableError(errType, excText)

        If isStale And retryCount < maxRetries Then
            ' Recovery: Refresh the realm and wait briefly before retrying
            If DEBUG_MODE Then Debug.Print "BiDi Info: Stale detected (Try " & retryCount & "). Recovering..."
            GetScriptRealmId contextId, True ' Force refresh realm ID
            WaitHeartbeat 200 ' Exponential-like backoff
        Else
            ' Success or non-recoverable error
            Exit For
        End If
    Next retryCount

    ' 6. Final error handling
    If InStr(strRes, """error""") > 0 Or InStr(strRes, "exceptionDetails") > 0 Then
        Err.Raise vbObjectError + 555, "BiDi:ActionFailed", "Failed on XPath: " & xpath & vbCrLf & "Res: " & strRes
    End If

    ExecuteActionWithStaleRetry = strRes
End Function

' ========================================================================================
' ExecuteClickByXPath: Clicks an element identified by XPath with automatic stale retry.
' ========================================================================================
Public Function ExecuteClickByXPath(ByVal xpath As String, _
                                    Optional ByVal contextId As String = "", _
                                    Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    
    ' 1. Get the centralized JS snippet
    Dim js As String: js = GetActionJs("click")
    
    ' 2. Execute the action with built-in retry logic
    ' Note: ExecuteActionWithStaleRetry internally uses ExecuteBiDiCommand
    ExecuteClickByXPath = ExecuteActionWithStaleRetry(xpath, js, , True, contextId, maxRetries)
    
End Function

' ========================================================================================
' ExecuteClickAndWaitByXPath: Clicks and synchronizes with page/SPA navigation.
' ========================================================================================
Public Function ExecuteClickAndWaitByXPath(ByVal xpath As String, _
                                           Optional ByVal waitNetworkIdle As Boolean = True, _
                                           Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                           Optional ByVal contextId As String = "", _
                                           Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                           Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS) As String
    
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())
    Dim oldUrl As String: oldUrl = GetCurrentUrl(targetCtx)

    ' Use the refactored click function above
    Dim res As String
    res = Me.ExecuteClickByXPath(xpath, targetCtx, maxRetries)

    ' Handle navigation sync using centralized helpers
    If waitNetworkIdle And IsScriptDone(res) Then
        ' ... (Wait logic remains the same, but calling cleaner UpdateMainContextId) ...
        UpdateMainContextId True
        InitializeSpaProbe targetCtx
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx
    End If

    ExecuteClickAndWaitByXPath = res
End Function

' ========================================================================================
' ExecuteClickAndNavigateByXPath: Click and wait for navigation using centralized methods.
' ========================================================================================
Public Function ExecuteClickAndNavigateByXPath(ByVal xpath As String, _
                                               Optional ByVal waitNetworkIdle As Boolean = True, _
                                               Optional ByVal contextId As String = "", _
                                               Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                               Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                               Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS) As String
    Dim clickRes As String
    
    ' 1. Execute the basic click (already has retry logic)
    clickRes = Me.ExecuteClickByXPath(xpath, contextId, maxRetries)

    ' 2. If click was successful, follow the navigation
    If IsScriptDone(clickRes) Then
        Dim newCtxId As String: newCtxId = UpdateMainContextId(True)
        If waitNetworkIdle Then
            WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, newCtxId
        End If
        ExecuteClickAndNavigateByXPath = Me.GetCurrentUrl(newCtxId)
    Else
        Err.Raise vbObjectError + 500, "BiDi:NavigateError", "Click action did not return 'done'."
    End If
End Function

' ========================================================================================
' ExecuteInputValueAndWaitByXPath: Inputs text into an element.
' ========================================================================================
Public Function ExecuteInputValueAndWaitByXPath(ByVal xpath As String, _
                                                ByVal valueToSet As String, _
                                                Optional ByVal contextId As String = "", _
                                                Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                                Optional ByVal waitNetworkIdle As Boolean = True, _
                                                Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String
    
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())
    
    ' 1. Ensure the SPA probe is ready if waiting is required
    If waitNetworkIdle Then InitializeSpaProbe targetCtx

    ' 2. Execute the input action using centralized JS
    Dim res As String
    res = ExecuteActionWithStaleRetry(xpath, GetActionJs("input"), valueToSet, True, targetCtx, maxRetries)

    ' 3. Synchronize with network activity if 'done' was received
    If waitNetworkIdle And IsScriptDone(res) Then
        WaitForNetworkIdleSpa minStableMs, 3000, targetCtx
    End If

    ExecuteInputValueAndWaitByXPath = res
End Function

' ========================================================================================
' ExecuteSelectValueAndWaitByXPath: Selects an option from a dropdown and waits for async tasks.
' Optimization: Injects and verifies the probe BEFORE the action to capture immediate fetches.
' ========================================================================================
Public Function ExecuteSelectValueAndWaitByXPath(ByVal xpath As String, _
                                                 ByVal valueOrText As String, _
                                                 Optional ByVal selectByText As Boolean = False, _
                                                 Optional ByVal waitNetworkIdle As Boolean = True, _
                                                 Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                                 Optional ByVal contextId As String = "", _
                                                 Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())

    ' 1. Proactive Probe Initialization
    ' Ensure the probe is intercepting fetch/XHR BEFORE the selection triggers any JS.
    If waitNetworkIdle Then
        InitializeSpaProbe targetCtx
        ProcessRecorderQueue
    End If

    ' 2. Retrieve the centralized JS snippet
    ' Logic: Choose between value-based or text-based selection
    Dim jsFunction As String
    If selectByText Then
        jsFunction = GetActionJs("select_text")
    Else
        jsFunction = GetActionJs("select")
    End If

    ' 3. Execute UI Action via the stale-retry runner
    Dim res As String
    res = ExecuteActionWithStaleRetry(xpath, jsFunction, valueOrText, True, targetCtx, maxRetries)

    ' 4. Synchronization Wait
    ' If the action was successful ('done'), enter the SPA stability loop
    If waitNetworkIdle And IsScriptDone(res) Then
        ProcessRecorderQueue
        WaitForNetworkIdleSpa minStableMs, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
    End If

    ExecuteSelectValueAndWaitByXPath = res
End Function

Public Function ExecuteGetTextByXPath(ByVal xpath As String, _
                                      Optional ByVal contextId As String = "", _
                                      Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    Dim jsFunction As String
    jsFunction = "function(element) { return element.innerText || element.textContent || ''; }"
    
    Dim strRes As String
    strRes = ExecuteActionWithStaleRetry(xpath, jsFunction, , False, contextId, maxRetries)

    ' Extract the text value directly
    ExecuteGetTextByXPath = CStr(GetJsonValue(strRes, "result.result.value"))
End Function

' ========================================================================================
' Locator Helpers
' ========================================================================================
Public Function ExecuteBrowsingContextLocateNodes(ByVal xpath As String, Optional ByVal contextId As String = "") As String
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId()

    Dim params As New Dictionary, locator As New Dictionary
    locator.Add "type", "xpath"
    locator.Add "value", xpath
    params.Add "context", targetCtx
    params.Add "locator", locator
    ExecuteBrowsingContextLocateNodes = ExecuteBiDiCommand("browsingContext.locateNodes", params)
End Function

' ========================================================================================
' GetSharedIdFromXPath: Resolves an XPath to a BiDi sharedId using centralized helpers.
' ========================================================================================
Private Function GetSharedIdFromXPath(ByVal xpath As String, Optional ByVal contextId As String = "") As String
    Dim strRes As String
    ' 1. Execute the locating command
    strRes = ExecuteBrowsingContextLocateNodes(xpath, contextId)

    ' 2. Extract the 'nodes' Collection using the helper
    Dim nodes As Collection
    Set nodes = GetJsonValue(strRes, "result.nodes")

    ' 3. Return the sharedId of the first matching element
    If Not nodes Is Nothing Then
        If nodes.Count > 0 Then
            ' Check existence of sharedId in the first node safely
            If nodes(1).Exists("sharedId") Then
                GetSharedIdFromXPath = CStr(nodes(1)("sharedId"))
                Exit Function
            End If
        End If
    End If

    GetSharedIdFromXPath = ""
End Function

' ========================================================================================
' GetCurrentUrl: Retrieves the current URL.
' Silent: Returns empty string on error to allow the polling loop to handle context swaps.
' ========================================================================================
Public Function GetCurrentUrl(Optional ByVal contextId As String = "") As String
    Dim res As String
    ' Replaced 6 lines of Dictionary construction with 1 line
    res = CallScript("function(){ return window.location.href; }", , contextId)
    
    GetCurrentUrl = CStr(GetJsonValue(res, "result.result.value"))
End Function

' ========================================================================================
' Storage & Extensions & CDP
' ========================================================================================
Public Function ExecuteStorageGetCookies(Optional ByVal filter As Dictionary = Nothing) As Collection
    Dim contextId As String: contextId = UpdateMainContextId()
    Dim params As New Dictionary, partition As New Dictionary

    partition.Add "type", "context"
    partition.Add "context", contextId
    params.Add "partition", partition
    If Not filter Is Nothing Then params.Add "filter", filter

    Dim strRes As String
    strRes = ExecuteBiDiCommand("storage.getCookies", params)

    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    Dim resultColl As New Collection
    On Error Resume Next
    If json.Exists("result") Then
        If json("result").Exists("cookies") Then Set resultColl = json("result")("cookies")
    End If
    On Error GoTo 0
    Set ExecuteStorageGetCookies = resultColl
End Function

Public Function ExecuteStorageSetCookie(ByVal cookieData As Dictionary) As String
    Dim contextId As String: contextId = UpdateMainContextId()
    Dim params As New Dictionary, partition As New Dictionary

    partition.Add "type", "context"
    partition.Add "context", contextId
    params.Add "partition", partition
    params.Add "cookie", cookieData

    ExecuteStorageSetCookie = ExecuteBiDiCommand("storage.setCookie", params)
End Function

Public Function ExecuteStorageDeleteCookies(Optional ByVal filter As Dictionary = Nothing) As String
    Dim contextId As String: contextId = UpdateMainContextId()
    Dim params As New Dictionary, partition As New Dictionary

    partition.Add "type", "context"
    partition.Add "context", contextId
    params.Add "partition", partition
    If Not filter Is Nothing Then params.Add "filter", filter

    ExecuteStorageDeleteCookies = ExecuteBiDiCommand("storage.deleteCookies", params)
End Function

Public Function ExecuteStorageDeleteAllCookies() As String
    Dim filter As New Dictionary
    ExecuteStorageDeleteAllCookies = ExecuteStorageDeleteCookies(filter)
End Function

Public Function ExecuteWebExtensionInstall(ByVal extensionPath As String) As String
    Dim params As New Dictionary, extData As New Dictionary
    extData.Add "type", "path"
    extData.Add "path", extensionPath
    params.Add "extensionData", extData
    ExecuteWebExtensionInstall = ExecuteBiDiCommand("webExtension.install", params)
End Function

' ========================================================================================
' ExecuteSetDownloadPath: Sets the directory for browser downloads via CDP.
' ========================================================================================
Public Function ExecuteSetDownloadPath(ByVal downloadPath As String) As String
    p_currentDownloadPath = downloadPath
    
    Dim cdpParams As New Dictionary
    cdpParams.Add "behavior", "allow"
    cdpParams.Add "downloadPath", downloadPath
    cdpParams.Add "eventsEnabled", True
    
    Dim resObj As Object: Set resObj = ExecuteCDPCommand("Page.setDownloadBehavior", cdpParams, True)
    
    ' Simplified error reporting
    If Not resObj("Ok") Then
        Debug.Print "BiDi Warning: Failed to set download path: " & resObj("Error")
    End If
    ExecuteSetDownloadPath = resObj("Json")
End Function

' ========================================================================================
' ExecuteWaitForDownloadCompletion: Monitors the file system for download finishing.
' Optimized: Throttles heavy Disk I/O (1s interval) while keeping BiDi events fluid (50ms).
' ========================================================================================
Public Function ExecuteWaitForDownloadCompletion(ByVal startTime As Date, _
                                                  Optional ByVal timeoutSec As Long = 30) As Boolean
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim timeoutMs As Double: timeoutMs = CDbl(timeoutSec) * 1000
    Dim lastScanTick As Double: lastScanTick = 0
    
    Dim folder As Scripting.folder
    Dim file As Scripting.file
    Dim hasTempFile As Boolean
    Dim foundNewFile As Boolean
    
    ' Throttling constant for heavy file system operations
    Const FILE_SCAN_INTERVAL_MS As Long = 1000

    ' Guard: Ensure download path is set and exists
    If p_currentDownloadPath = "" Or Not fso_.FolderExists(p_currentDownloadPath) Then
        Debug.Print "BiDi Error: Invalid or missing download path: " & p_currentDownloadPath
        Exit Function
    End If

    Debug.Print "BiDi: Monitoring download folder (Scan interval: " & FILE_SCAN_INTERVAL_MS & "ms)..."

    ' --- Optimized Hybrid Monitoring Loop ---
    Do
        ' 1. Light Processing: Always handle BiDi events to keep state synced and memory clean
        ProcessRecorderQueue
        
        ' 2. Safety Exit: Timeout Check
        If GetRawTicks() - startTick > timeoutMs Then
            Debug.Print "BiDi Warning: Download monitoring timed out after " & timeoutSec & " seconds."
            ExecuteWaitForDownloadCompletion = False
            Exit Function
        End If

        ' 3. Heavy Processing: File System Scan (executed only once per interval)
        If GetRawTicks() - lastScanTick > FILE_SCAN_INTERVAL_MS Then
            hasTempFile = False
            foundNewFile = False
            
            On Error Resume Next ' Protect against transient file access errors during writing
            Set folder = fso_.GetFolder(p_currentDownloadPath)
            
            For Each file In folder.Files
                ' Check for common browser temporary download extensions
                If InStr(LCase(file.Name), ".crdownload") > 0 Or InStr(LCase(file.Name), ".tmp") > 0 Then
                    hasTempFile = True
                Else
                    ' Check if the file was modified after the download started
                    If DateDiff("s", startTime, file.DateLastModified) >= 0 Then
                        foundNewFile = True
                    End If
                End If
            Next file
            On Error GoTo 0

            ' Completion Logic: Temporary files must be gone, and a new file must be present
            If Not hasTempFile And foundNewFile Then
                Debug.Print "BiDi: Download detected and verified as complete."
                ExecuteWaitForDownloadCompletion = True
                Exit Function
            End If
            
            ' Update the last scan timestamp
            lastScanTick = GetRawTicks()
            Set folder = Nothing ' Immediate cleanup of folder object
        End If

        ' 4. UI Maintenance: Keep Excel responsive and sleep briefly
        ThrottledDoEvents
        sleep 50 ' Short sleep for responsive background event handling
    Loop
End Function
Public Sub ExecuteClearDownloadFolder()
    If p_currentDownloadPath = "" Then Exit Sub
    If Not fso_.FolderExists(p_currentDownloadPath) Then Exit Sub

    On Error Resume Next
    fso_.DeleteFile fso_.BuildPath(p_currentDownloadPath, "*.*"), True
    On Error GoTo 0
End Sub

' ========================================================================================
' ExecuteEnableStealthMode: Injects script to mask automation footprint.
' ========================================================================================
Public Sub ExecuteEnableStealthMode()
    Dim js As String
    js = "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
    
    Dim cdpParams As New Dictionary
    cdpParams.Add "source", js
    
    Dim resObj As Object
    Set resObj = ExecuteCDPCommand("Page.addScriptToEvaluateOnNewDocument", cdpParams)
    
    If Not resObj("Ok") Then
        Debug.Print "BiDi Warning: Stealth Mode injection failed: " & resObj("Error")
    End If
End Sub

' ========================================================================================
' ExecuteShadowClickAndWait: Clicks an element inside Shadow DOM and waits for stability.
' ========================================================================================
Public Function ExecuteShadowClickAndWait(ByVal selectorsArray As Variant, _
                                          Optional ByVal waitNetworkIdle As Boolean = True, _
                                          Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                          Optional ByVal contextId As String = "", _
                                          Optional ByVal searchTimeoutMs As Long = 5000) As String
    
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())
    If waitNetworkIdle Then InitializeSpaProbe targetCtx

    ' 1. Build Arguments (Array of strings for selectors)
    Dim args As New Collection
    args.Add BuildShadowSelectorArg(selectorsArray) ' Existing helper
    Dim argTime As New Dictionary: argTime.Add "type", "number": argTime.Add "value", searchTimeoutMs
    args.Add argTime

    ' 2. Execute via CallScript
    Dim res As String: res = CallScript(GetActionJs("shadow_click"), args, targetCtx, True)

    ' 3. Sync and Wait
    If InStr(res, "Error") > 0 Then
          Err.Raise vbObjectError + 557, "BiDi:ShadowError", "Shadow DOM element not found."
    End If

    If waitNetworkIdle And IsScriptDone(res) Then
        WaitForNetworkIdleSpa minStableMs, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
    End If

    ExecuteShadowClickAndWait = res
End Function

' ========================================================================================
' ExecuteShadowClickAndNavigate: Specialized for high-speed redirect handling.
' Logic: Polling for URL change immediately after shadow-click to intercept redirects
' without waiting for 100ms recovery cycles, ensuring faster landing on login pages.
' ========================================================================================
' ========================================================================================
' ExecuteShadowClickAndNavigate: Specialized for high-speed redirect handling via Shadow DOM.
' Logic: Polls for URL change immediately after click to intercept redirects (SSO/SAML).
' ========================================================================================
Public Function ExecuteShadowClickAndNavigate(ByVal selectorsArray As Variant, _
                                              Optional ByVal waitNetworkIdle As Boolean = True, _
                                              Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                              Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                              Optional ByVal contextId As String = "", _
                                              Optional ByVal searchTimeoutMs As Long = 5000) As String
    
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId(True))
    Dim oldUrl As String: oldUrl = GetCurrentUrl(targetCtx)
    
    ' 1. Build Arguments for Shadow Click
    Dim args As New Collection
    args.Add BuildShadowSelectorArg(selectorsArray)
    Dim argTime As New Dictionary: argTime.Add "type", "number": argTime.Add "value", searchTimeoutMs
    args.Add argTime
    
    ' 2. Execute Shadow Click via CallScript
    Dim res As String
    res = CallScript(GetActionJs("shadow_click"), args, targetCtx, True)
    
    ' 3. [PROACTIVE NAVIGATION GUARD]
    ' Catch the start of redirect chains immediately (SSO, Login redirects, etc.)
    If waitNetworkIdle And IsScriptDone(res) Then
        Dim startTick As Currency: startTick = GetRawTicks()
        Do
            Dim currentUrl As String: currentUrl = GetCurrentUrl(targetCtx)
            ' Exit loop if URL changes or context blanks (start of reload)
            If currentUrl <> oldUrl Or currentUrl = "" Then Exit Do
            
            ' Cap the proactive wait at 500ms
            WaitHeartbeat 30
            If IsTimeout(startTick, 500) Then Exit Do
        Loop

        ' 4. Refresh environment for the landing page
        UpdateMainContextId True
        InitializeSpaProbe targetCtx
        
        ' 5. Wait for network/DOM to stabilize
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx
    End If
    
    ExecuteShadowClickAndNavigate = Me.GetCurrentUrl(targetCtx)
End Function

' ========================================================================================
' ExecuteBiDiCommand: Core command handler with automated recovery.
' Updated: Now passes specific tags to DebugPrintLong for better log analysis.
' ========================================================================================
Public Function ExecuteBiDiCommand(ByVal method As String, ByVal params As Dictionary) As String
    If socket_ Is Nothing Then Err.Raise 91, , "Socket object is not set."

    Dim retryCount As Integer
    Dim strRes As String, sendMsg As String
    Dim currentParams As Dictionary
    Dim dic As Dictionary
    
    Set currentParams = params

    For retryCount = 0 To MAX_CONTEXT_RECOVERY_RETRIES
        ' 1. Prepare Request
        Set dic = New Dictionary
        dic.Add "id", p_nextRequestId
        dic.Add "method", method
        dic.Add "params", currentParams
        p_nextRequestId = p_nextRequestId + 1

        sendMsg = WebJsonConverter.ConvertToJson(dic)
        
        ' --- LOGGING: Pass "[REQ]" tag for clarity ---
        If DEBUG_MODE Then DebugPrintLong method & " (Try " & retryCount & ")", sendMsg, "[REQ]"
        
        ' 2. Communicate
        strRes = socket_.SendAndReceive(sendMsg)
        
        ' --- LOGGING: Pass "[RES]" tag for clarity ---
        If DEBUG_MODE Then DebugPrintLong method & " (Try " & retryCount & ")", strRes, "[RES]"
        
        ' 3. Error Detection and Recovery
        Dim isContextDestroyed As Boolean
        isContextDestroyed = IsContextLostError(strRes)

        If isContextDestroyed And Not m_isRecovering And retryCount < MAX_CONTEXT_RECOVERY_RETRIES Then
            On Error GoTo RecoveryCleanup
            m_isRecovering = True

            ' --- Exponential Backoff (300ms base) ---
            Dim waitMs As Long
            waitMs = 300 * (2 ^ retryCount)
            If waitMs > 2400 Then waitMs = 2400
            
            ' --- LOGGING: Pass "[RECOVERY]" tag to highlight the event ---
            DebugPrintLong "Context lost. Retrying...", strRes, "[RECOVERY]"
            
            sleep waitMs
            DoEvents

            ' --- State Restoration ---
            ClearState
            Dim newCtx As String: newCtx = UpdateMainContextId(True)

            If newCtx = "" Then
                m_isRecovering = False
                CheckAndRaiseError strRes, method
                Exit For
            End If

            Dim retryParams As Dictionary
            Set retryParams = DeepCopyDictionary(params)

            If Not UpdateContextAndRealm(retryParams, newCtx) Then
                m_isRecovering = False
                CheckAndRaiseError strRes, method
                Exit For
            End If

            Set currentParams = retryParams
            m_isRecovering = False
        Else
            ' 4. Final Validation
            CheckAndRaiseError strRes, method
            Exit For
        End If
    Next retryCount

    ExecuteBiDiCommand = strRes
    Exit Function

RecoveryCleanup:
    m_isRecovering = False
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "BiDi:Recovery", "Critical error during recovery: " & Err.Description
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
' DeepCopyDictionary: Recursively clones Dictionary objects with stack overflow protection.
' Parameters:
'   - source: The Dictionary to be cloned.
'   - currentDepth: Internal tracker for recursion depth.
' Returns: A deep copy of the Dictionary up to the defined depth limit.
' ========================================================================================
Private Function DeepCopyDictionary(ByVal source As Dictionary, _
                                    Optional ByVal currentDepth As Long = 0) As Dictionary
    ' Recursion limit to prevent stack overflow on extremely deep DOM or BiDi structures.
    ' A depth of 15 is typically sufficient for standard WebDriver BiDi responses.
    Const MAX_COPY_DEPTH As Long = 15
    
    ' 1. Guard: Check recursion depth
    If currentDepth > MAX_COPY_DEPTH Then
        ' Fallback to shallow copy (reference) to protect the VBA stack when limit is reached.
        Set DeepCopyDictionary = source
        Debug.Print "BiDi Warning: DeepCopy max depth reached (" & MAX_COPY_DEPTH & "). Falling back to shallow copy."
        Exit Function
    End If

    Dim result As New Dictionary
    Dim key As Variant
    Dim value As Variant

    For Each key In source.keys
        Set value = Nothing
        ' Safely check if the value is an object to avoid Type Mismatch errors
        If IsObject(source(key)) Then
            Set value = source(key)
        End If
        
        If value Is Nothing Then
            ' Primitive types (String, Long, Boolean, etc.) are copied by value
            result.Add key, source(key)
        ElseIf TypeName(value) = "Dictionary" Then
            ' Recursive call for nested Dictionaries with incremented depth
            result.Add key, DeepCopyDictionary(value, currentDepth + 1)
        ElseIf TypeName(value) = "Collection" Then
            ' Recursive call for Collections to handle nested Dictionary arrays
            result.Add key, DeepCopyCollection(value, currentDepth + 1)
        Else
            ' Direct reference copy for other object types
            result.Add key, value
        End If
    Next key
    
    Set DeepCopyDictionary = result
End Function

' ========================================================================================
' DeepCopyCollection: Recursive helper to clone Collections containing Dictionaries.
' ========================================================================================
Private Function DeepCopyCollection(ByVal source As Collection, ByVal currentDepth As Long) As Collection
    Dim result As New Collection
    Dim Item As Variant
    
    For Each Item In source
        If IsObject(Item) Then
            ' Recursively handle Dictionaries or nested Collections within the array
            If TypeName(Item) = "Dictionary" Then
                ' FIX: Increment depth when calling DeepCopyDictionary to ensure stack protection
                result.Add DeepCopyDictionary(Item, currentDepth + 1)
            ElseIf TypeName(Item) = "Collection" Then
                ' FIX: Increment depth when calling DeepCopyCollection
                result.Add DeepCopyCollection(Item, currentDepth + 1)
            Else
                result.Add Item
            End If
        Else
            ' Copy primitive items directly
            result.Add Item
        End If
    Next Item
    Set DeepCopyCollection = result
End Function

' ========================================================================================
' CheckAndRaiseError: Standardized error reporting for BiDi and JavaScript exceptions.
' Uses GetJsonValue to extract detailed error info and prints to the Immediate Window.
' ========================================================================================
Private Sub CheckAndRaiseError(ByVal jsonString As String, ByVal methodName As String)
    ' 1. Check for Protocol-level errors
    Dim errType As String: errType = CStr(GetJsonValue(jsonString, "error"))
    Dim errMsg As String: errMsg = CStr(GetJsonValue(jsonString, "message"))

    If errType <> "" Then
        If IsIgnorableError(errType, errMsg) Then Exit Sub
        Err.Raise vbObjectError + 513, "BiDi:" & methodName, _
                  "Browser Protocol Error [" & errType & "]: " & errMsg
    End If

    ' 2. Check for JavaScript Execution Errors (Exceptions)
    ' "result.exceptionDetails" is the standard path for script failures in BiDi
    Dim exception As Object
    On Error Resume Next
    Set exception = GetJsonValue(jsonString, "result.exceptionDetails")
    On Error GoTo 0

    If Not exception Is Nothing Then
        ' Use the existing formatter to build a rich error message (stack trace, line numbers)
        Dim detailedError As String: detailedError = FormatJavaScriptError(exception)
        
        If DEBUG_MODE Then
            Debug.Print "===================================================="
            Debug.Print "BIDI JS ERROR REPORT | Method: " & methodName
            Debug.Print detailedError
            Debug.Print "===================================================="
        End If
        
        Err.Raise vbObjectError + 514, "BiDi:JS_Exception", _
                  "JS Error in " & methodName & ": " & exception("text") & _
                  " (Check Immediate Window for stack trace)"
    End If
End Sub

' ========================================================================================
' DebugPrintLong: Prints structured logs with tags and handles long strings by chunking.
' Note: The 'tag' parameter is optional to maintain compatibility with existing calls.
' ========================================================================================
Private Sub DebugPrintLong(ByVal header As String, ByVal longText As String, _
                           Optional ByVal tag As String = "[INFO]")
    Const CHUNK_SIZE As Integer = 250
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
    ' 1. Restore native APIs if the probe was previously registered
    If m_isProbeRegistered Then
        On Error Resume Next
        Dim p As New Dictionary
        p.Add "expression", GetProbeJs("restore")
        ' Use the new helper to target the main context
        p.Add "target", BuildContextTarget(UpdateMainContextId())
        p.Add "awaitPromise", False
        ExecuteBiDiCommand "script.evaluate", p
        On Error GoTo 0
    End If

    ' 2. Reset cached identification and flags
    p_mainRealmId = ""
    p_cdpSessionId = ""
    p_mainContextId = ""
    p_pendingRequestCount = 0
    m_isProbeRegistered = False
    m_isSubscribed = False
End Sub

' ========================================================================================
' UpdateMainContextId: Identifies and caches the primary browsing context ID.
' Stable: Returns the ID as soon as a context exists, even if the URL is about:blank.
' This prevents infinite loops during initial navigation or blank page states.
' ========================================================================================
Public Function UpdateMainContextId(Optional ByVal forceUpdate As Boolean = False) As String
    If Not forceUpdate And p_mainContextId <> "" Then
        UpdateMainContextId = p_mainContextId
        Exit Function
    End If

    Dim startTick As Currency: startTick = GetRawTicks()
    Dim strRes As String, resJson As Object, ctxList As Collection
    
    ' --- Reliability Loop ---
    Do
        strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        On Error Resume Next
        Set resJson = WebJsonConverter.ParseJson(strRes)
        On Error GoTo 0

        If Not resJson Is Nothing Then
            If resJson.Exists("result") Then
                Set ctxList = resJson("result")("contexts")
                
                ' SUCCESS: If at least one context exists, grab it and exit.
                If ctxList.Count > 0 Then
                    p_mainContextId = ctxList(1)("context")
                    UpdateMainContextId = p_mainContextId
                    Exit Function
                End If
            End If
        End If

        ' Safety timeout (Short: 2s)
        If GetRawTicks() - startTick > 2000 Then Exit Do
        sleep 100: DoEvents ' Increased sleep to reduce CPU pressure
    Loop

    UpdateMainContextId = p_mainContextId
End Function
Public Function GetScriptRealmId(Optional ByVal contextId As String = "", Optional ByVal forceUpdate As Boolean = False) As String
    Dim targetCtx As String
    If contextId <> "" Then
        targetCtx = contextId
    Else
        If p_mainContextId = "" Or forceUpdate Then UpdateMainContextId True
        targetCtx = p_mainContextId
    End If

    If Not forceUpdate And targetCtx = p_mainContextId And p_mainRealmId <> "" Then
        GetScriptRealmId = p_mainRealmId
        Exit Function
    End If

    Dim params As New Dictionary
    params.Add "context", targetCtx

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.getRealms", params)

    Dim realmId As String
    On Error Resume Next
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    If Not json Is Nothing Then
        If json.Exists("result") Then
            If json("result").Exists("realms") Then
                Dim realms As Collection: Set realms = json("result")("realms")
                Dim r As Dictionary
                For Each r In realms
                    If LCase(r("type")) = "window" Then
                        realmId = r("realm")
                        Exit For
                    End If
                Next r
                If realmId = "" And realms.Count > 0 Then
                    realmId = realms(1)("realm")
                End If
            End If
        End If
    End If
    On Error GoTo 0

    ' Cache the main realm ID
    If targetCtx = p_mainContextId Then p_mainRealmId = realmId
    GetScriptRealmId = realmId
End Function

Private Function GetCdpSessionId(Optional ByVal forceUpdate As Boolean = False) As String
    If p_mainContextId = "" Then UpdateMainContextId
    
    If Not forceUpdate And p_cdpSessionId <> "" Then
        GetCdpSessionId = p_cdpSessionId
        Exit Function
    End If

    Dim params As New Dictionary
    params.Add "context", p_mainContextId
    
    Dim strRes As String
    ' goog:cdp.getSession may fail if the target context is being navigated or destroyed
    strRes = ExecuteBiDiCommand("goog:cdp.getSession", params)
    
    On Error Resume Next
    Dim json As Object: Set json = WebJsonConverter.ParseJson(strRes)
    
    If Not json Is Nothing Then
        If json.Exists("result") Then
            p_cdpSessionId = json("result")("session")
            ' [FIXED] Reset failure counter on successful session acquisition
            m_consecutiveCdpFailures = 0
            GetCdpSessionId = p_cdpSessionId
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    ' [FIXED] Increment failure counter if session retrieval fails
    m_consecutiveCdpFailures = m_consecutiveCdpFailures + 1
    Debug.Print "BiDi Warning: Failed to acquire CDP Session (Count: " & m_consecutiveCdpFailures & ")"
    
    GetCdpSessionId = ""
End Function

' ========================================================================================
' ExecuteRegisterAutoClickerByXPath: Registers a background observer via CDP to click a button.
' ========================================================================================
Public Sub ExecuteRegisterAutoClickerByXPath(ByVal xpath As String, _
                                             Optional ByVal timeoutMs As Long = 30000)
    ' 1. Retrieve centralized JS
    Dim js As String: js = GetActionJs("autoclicker")
    
    ' 2. Prepare CDP injection parameters
    ' We need to wrap the xpath string safely for JS injection
    Dim sourceJs As String
    sourceJs = Replace(js, "arguments[0]", WebJsonConverter.ConvertToJson(xpath))
    sourceJs = Replace(sourceJs, "arguments[1]", timeoutMs)

    Dim cdpParams As New Dictionary
    cdpParams.Add "source", sourceJs
    
    ' 3. Register via CDP
    Dim result As Object
    Set result = Me.ExecuteCDPCommand("Page.addScriptToEvaluateOnNewDocument", cdpParams)
    
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
' ProcessRecorderQueue: Processes the incoming message queue from the WebSocket.
' ========================================================================================
Public Sub ProcessRecorderQueue()
    ' Re-entrancy guard: Prevent nested execution via DoEvents
    If m_isProcessingQueue Then Exit Sub
    
    m_isProcessingQueue = True
    
    On Error GoTo Cleanup
    
    Dim msg As String
    Dim json As Object
    Dim processedCount As Long
    Const MAX_MESSAGES_PER_BATCH As Long = 1000
    
    Do While processedCount < MAX_MESSAGES_PER_BATCH
        Set json = Nothing
        msg = socket_.GetMessage()
        
        If Len(msg) = 0 Then Exit Do
        
        ' Basic validation
        If Len(msg) >= MIN_MSG_SIZE And Len(msg) <= MAX_MSG_SIZE Then
            If InStr(msg, """method"":") > 0 Or InStr(msg, """id"":") > 0 Then
                On Error Resume Next
                Set json = WebJsonConverter.ParseJson(msg)
                If Err.Number = 0 And Not json Is Nothing Then
                    On Error GoTo Cleanup ' Restore main error handler
                    DispatchGlobalEvents json
                End If
                On Error GoTo Cleanup
            End If
        End If
        
        processedCount = processedCount + 1
        
        ' Allow OS updates every 100 messages
        If (processedCount Mod 100) = 0 Then ThrottledDoEvents
    Loop

Cleanup:
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
' QuerySpaProbe: Retrieves probe state and validates integrity via versioning.
' [FIXED] Added version verification to detect "ghost" probes caused by realm races.
' ========================================================================================
Public Function QuerySpaProbe(Optional ByVal contextId As String = "") As Dictionary
    Dim res As String
    res = CallScript(GetProbeJs("query"), , contextId, False)
    
    Dim statusStr As String: statusStr = CStr(GetJsonValue(res, "result.result.value"))
    Dim dic As Dictionary
    On Error Resume Next
    Set dic = WebJsonConverter.ParseJson(statusStr)
    On Error GoTo 0
    
    ' Validate version integrity
    If Not dic Is Nothing Then
        If Not dic("active") Or (dic("version") <> PROBE_VERSION) Then Set dic = Nothing
    End If
    Set QuerySpaProbe = dic
End Function

' ========================================================================================
' InitializeSpaProbe: Ensures the monitoring environment and event subscriptions are ready.
' Updated: Uses the extended ExecuteSessionSubscribe for cleaner targeting.
' ========================================================================================
Private Sub InitializeSpaProbe(ByVal contextId As String)
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())

    ' 1. CDP Persistent Registration: Injects probe on every new document automatically.
    If Not m_isProbeRegistered Then
        Dim cdpParams As New Dictionary
        cdpParams.Add "source", GetProbeJs("inject")
        
        Dim result As Object
        Set result = Me.ExecuteCDPCommand("Page.addScriptToEvaluateOnNewDocument", cdpParams)
        
        If result("Ok") Then
            m_isProbeRegistered = True
        End If
    End If

    ' 2. Immediate Injection: Ensure the probe is active in the current realm right now.
    StartSpaIdleProbe targetCtx
    
    Dim retry As Integer
    Dim probeStatus As Dictionary
    For retry = 1 To 10
        Set probeStatus = QuerySpaProbe(targetCtx)
        If Not probeStatus Is Nothing Then
            If probeStatus("active") Then Exit For
        End If
        WaitHeartbeat 50
    Next retry

    ' 3. Unified Event Subscription
    ' Subscribe to lifecycle and network events specifically for the target context.
    Dim events As Variant
    events = Array("network.beforeRequestSent", "network.responseCompleted", _
                   "network.fetchError", "browsingContext.load", "browsingContext.fragmentNavigated")
    
    Dim ctxs As New Collection: ctxs.Add targetCtx
    ExecuteSessionSubscribe events, ctxs
End Sub

'Recovery
Private Sub RecoverSpaProbe(ByVal contextId As String)
    ' Force update RealmID (follow new page context)
    GetScriptRealmId contextId, True
    
    ' Re-inject Probe into the new page
    StartSpaIdleProbe contextId
    
    ' Note: Navigation detected. Probe re-injected.
End Sub

' ========================================================================================
' WaitForNetworkIdleSpa: Simplified convergence waiter.
' [REFACTORED] Added a mandatory initial drain and activity-buffer to ensure VBA state
' is fully synchronized with browser bursts before starting the stability timer.
' ========================================================================================
Public Sub WaitForNetworkIdleSpa(ByVal minStableMs As Long, _
                                 ByVal maxTimeoutMs As Long, _
                                 Optional ByVal contextId As String = "")
    
    ' 1. Inject or verify the probe
    InitializeSpaProbe contextId
    
    ' 2. [CRITICAL] Initial Burst Drain
    ' Clear any backlog from the socket immediately after navigation/click
    ' to ensure VBA's pendingRequestCount is current.
    ProcessRecorderQueue
    
    ' 3. [STRATEGY] Latency Buffer
    ' Wait briefly (e.g., 100ms) to allow the first network/DOM triggers
    ' from the browser's engine to actually reach the WebSocket.
    sleep 100
    ProcessRecorderQueue ' Catch any initial triggers
    
    Dim startTick As Currency: startTick = GetRawTicks()
    Dim stableStart As Double: stableStart = GetRawTicks()
    Dim nowTick As Double

    Do
        ' 4. Continuous Drain
        ' Since MAX_MESSAGES_PER_BATCH is 1000, we stay synced even during heavy hydration.
        ProcessRecorderQueue
        
        ThrottledDoEvents
        sleep 30 ' High-frequency polling
        
        nowTick = GetRawTicks()
        If (nowTick - startTick) > CDbl(maxTimeoutMs) Then
            Debug.Print "BiDi Warning: WaitForNetworkIdleSpa timed out."
            Exit Do
        End If

        Dim probe As Dictionary: Set probe = QuerySpaProbe(contextId)
        
        If Not probe Is Nothing Then
            ' 5. Evaluate convergence
            ' This function uses the p_pendingRequestCount sync logic for extra safety.
            If EvaluateSpaIdleCondition(probe, minStableMs, stableStart, nowTick) Then
                Exit Do
            End If
        Else
            ' Probe lost due to hard navigation; recover and reset timer
            RecoverSpaProbe contextId
            stableStart = nowTick
        End If
    Loop
End Sub

' ========================================================================================
' EvaluateSpaIdleCondition: Convergent logic using unified activity timestamps.
' ========================================================================================
Private Function EvaluateSpaIdleCondition(ByVal probe As Dictionary, _
                                          ByVal minStableMs As Long, _
                                          ByRef stableStart As Double, _
                                          ByVal nowTick As Double) As Boolean
    
    Dim jsNow As Double, lastActivityTs As Double
    Dim inflFetch As Long, inflXhr As Long

    ' Extract high-res data from the injected browser probe
    jsNow = CDbl(probe("now"))
    lastActivityTs = CDbl(probe("lastMutationTs"))
    inflFetch = CLng(probe("inflightFetchCount"))
    inflXhr = CLng(probe("inflightXhrCount"))

    ' --- Idle Judgment Logic ---
    Dim netIdle As Boolean, activityStable As Boolean

    ' 1. [CRITICAL] Self-Correction Logic
    ' If the browser confirms zero inflight transfers, sync the VBA counter to 0.
    ' This prevents "zombie waiting" caused by missed Bidi network events.
    If inflFetch = 0 And inflXhr = 0 Then
        p_pendingRequestCount = 0 ' Direct sync
        netIdle = True
    Else
        netIdle = (p_pendingRequestCount = 0)
    End If

    ' 2. Unified Stability Check (DOM + Network End + WebSocket Messages)
    activityStable = (jsNow - lastActivityTs) >= CDbl(minStableMs)

    ' --- Final Convergence Check ---
    If netIdle And activityStable Then
        ' Ensure stability is maintained for the duration of the window
        If (nowTick - stableStart) >= CDbl(minStableMs) Then
            EvaluateSpaIdleCondition = True
        Else
            EvaluateSpaIdleCondition = False
        End If
    Else
        ' Activity detected; reset stability timer and continue waiting
        stableStart = nowTick
        EvaluateSpaIdleCondition = False
    End If
End Function

' ========================================================================================
' ExecuteCDPCommand: Sends a raw CDP command via BiDi tunneling.
' ========================================================================================
Public Function ExecuteCDPCommand(ByVal cdpMethod As String, _
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
' StartSpaIdleProbe: Force-injects the monitoring script into the current realm.
' ========================================================================================
Public Sub StartSpaIdleProbe(Optional ByVal contextId As String = "")
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId(True))
    Dim targetRealm As String: targetRealm = GetScriptRealmId(targetCtx)
    
    Dim p As New Dictionary
    p.Add "expression", GetProbeJs("inject")
    ' Target the specific realm to ensure proper hook placement
    p.Add "target", BuildContextTarget(targetCtx, targetRealm)
    p.Add "awaitPromise", False

    ExecuteBiDiCommand "script.evaluate", p
End Sub

' ========================================================================================
' Internal JSON Result Parsers (Strict Evaluation)
' ========================================================================================

' Extracts the actual execution value from a script.callFunction or script.evaluate response.
' Standard BiDi Path: result (CommandResponse) -> result (ScriptResult) -> value
Private Function GetScriptResultValue(ByVal jsonResponse As String) As Variant
    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonResponse)
    
    ' Validate nested structure to avoid errors with non-script responses
    If Not json Is Nothing Then
        If json.Exists("result") Then
            If json("result").Exists("result") Then
                If json("result")("result").Exists("value") Then
                    ' Found the actual value returned by JavaScript
                    GetScriptResultValue = json("result")("result")("value")
                    Exit Function
                End If
            End If
        End If
    End If
    On Error GoTo 0
    
    GetScriptResultValue = Empty
End Function

' Checks if the injected script returned exactly "done".
' This is used to confirm the completion of UI actions (click, input, etc.)
Private Function IsScriptDone(ByVal jsonResponse As String) As Boolean
    Dim resultVal As Variant
    resultVal = GetScriptResultValue(jsonResponse)
    
    ' Only return True if the value is explicitly "done"
    If Not IsEmpty(resultVal) Then
        IsScriptDone = (CStr(resultVal) = "done")
    Else
        IsScriptDone = False
    End If
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
' Internal JSON Helper: Extracts a value from a nested JSON string using a dot-notation path.
' Arguments: jsonStr (JSON string), keyPath (e.g., "result.result.value")
' ========================================================================================
Private Function GetJsonValue(ByVal jsonStr As String, ByVal keyPath As String) As Variant
    Dim json As Object
    On Error Resume Next
    
    ' 1. Parse the JSON string into an object (Dictionary or Collection)
    Set json = WebJsonConverter.ParseJson(jsonStr)
    If json Is Nothing Then Exit Function

    Dim keys() As String: keys = Split(keyPath, ".")
    Dim target As Object: Set target = json
    Dim i As Long

    ' 2. Traverse the object tree until the second-to-last key
    For i = 0 To UBound(keys) - 1
        ' Ensure the current target is a Dictionary before calling .Exists
        If TypeName(target) = "Dictionary" Then
            If Not target.Exists(keys(i)) Then Exit Function
            
            ' Move to the next nested object
            Set target = target(keys(i))
        Else
            ' Cannot traverse further if the target is not a Dictionary
            Exit Function
        End If
    Next i

    ' 3. Extract and return the final key's value
    If Not target Is Nothing Then
        Dim lastKey As String: lastKey = keys(UBound(keys))
        
        If TypeName(target) = "Dictionary" Then
            If target.Exists(lastKey) Then
                ' CRITICAL FIX: Determine if the value is an object or a primitive.
                ' VBA requires the 'Set' keyword for Objects (Dictionary/Collection).
                If IsObject(target(lastKey)) Then
                    Set GetJsonValue = target(lastKey)
                Else
                    GetJsonValue = target(lastKey)
                End If
            End If
        End If
    End If
    
    On Error GoTo 0
End Function
' ========================================================================================
' Internal JS Helper: Centralized JavaScript snippets for all actions.
' After implementing this, you can delete old GetJsFor... functions.
' ========================================================================================
Private Function GetActionJs(ByVal actionType As String) As String
    Dim js As String
    
    Select Case LCase(actionType)
        Case "click"
            ' Scrolls to the element and performs a native click.
            js = "function(e) { return new Promise((r) => { " & _
                 "  e.scrollIntoView({block:'center',inline:'center'}); " & _
                 "  e.focus(); e.click(); " & _
                 "  setTimeout(() => r('done'), 50); " & _
                 "}); }"
                 
        Case "input"
            ' Focuses, clears, and inserts text via execCommand to trigger all events.
            js = "function(e,v) { return new Promise((r) => { " & _
                 "  e.scrollIntoView({block:'center',inline:'center'}); " & _
                 "  e.click(); e.focus(); e.value=''; " & _
                 "  var s=document.execCommand('insertText',false,v); " & _
                 "  if(!s){ e.value=v; e.dispatchEvent(new Event('input',{bubbles:true})); e.dispatchEvent(new Event('change',{bubbles:true})); } " & _
                 "  e.blur(); setTimeout(() => r('done'), 50); " & _
                 "}); }"
                 
        Case "select"
            ' Selects an option by its value property.
            js = "function(s,v) { return new Promise((r) => { " & _
                 "  s.scrollIntoView({block:'center',inline:'center'}); " & _
                 "  s.focus(); s.value = v; " & _
                 "  s.dispatchEvent(new Event('input', {bubbles:true})); " & _
                 "  s.dispatchEvent(new Event('change', {bubbles:true})); " & _
                 "  s.blur(); setTimeout(() => r('done'), 50); " & _
                 "}); }"

        Case "select_text"
            ' Selects an option by its visible innerText.
            js = "function(s, textToFind) { return new Promise((r) => { " & _
                 "  const normalize = (str) => str.replace(/[\s\u00A0]+/g, ' ').trim(); " & _
                 "  var found = false; var targetText = normalize(textToFind); " & _
                 "  for (var i = 0; i < s.options.length; i++) { " & _
                 "    if (normalize(s.options[i].text) === targetText) { s.value = s.options[i].value; found = true; break; } " & _
                 "  } " & _
                 "  if (!found) { r('Option text not found: ' + textToFind); return; } " & _
                 "  s.scrollIntoView({block:'center'}); s.focus(); " & _
                 "  s.dispatchEvent(new Event('input', {bubbles:true})); s.dispatchEvent(new Event('change', {bubbles:true})); " & _
                 "  s.blur(); setTimeout(() => r('done'), 50); " & _
                 "}); }"

        Case "visibility_check"
            ' Checks if an element is rendered and not hidden by CSS.
            js = "function(el) { " & _
                 "  if (!el || !el.isConnected) return false; " & _
                 "  const rect = el.getBoundingClientRect(); " & _
                 "  if (rect.width === 0 || rect.height === 0) return false; " & _
                 "  const style = window.getComputedStyle(el); " & _
                 "  return (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0'); " & _
                 "}"
                 
        Case "ready_state"
            ' Standard check for document completion.
            js = "function(){ return document.readyState; }"
            
        Case "shadow_click"
            ' Logic: Recursively penetrates Shadow Roots to find and click an element.
            ' Expects: selectors (array), searchTimeout (ms)
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
                 "    if(!el) return 'Error: Element not found'; " & _
                 "    el.scrollIntoView({block:'center'}); el.focus(); el.click(); " & _
                 "    return 'done'; " & _
                 "  }); " & _
                 "}"
    
        Case "wait_options"
            ' Logic: Polls the select element until it has at least minCount options or times out.
            ' Args: element, minCount, timeoutMs
            js = "function(element, minCount, timeout) { " & _
                 "  return new Promise((resolve) => { " & _
                 "    const start = Date.now(); " & _
                 "    const check = () => { " & _
                 "      if (element.options && element.options.length >= minCount) { " & _
                 "        resolve(true); " & _
                 "      } else if (Date.now() - start > timeout) { " & _
                 "        resolve(false); " & _
                 "      } else { " & _
                 "        setTimeout(check, 100); " & _
                 "      } " & _
                 "    }; check(); " & _
                 "  }); " & _
                 "}"

        Case "get_options"
            ' Logic: Scans all options in a select element and returns them as a JSON string.
            js = "function(element) { " & _
                 "  var opts = []; " & _
                 "  for (var i = 0; i < element.options.length; i++) { " & _
                 "    var o = element.options[i]; " & _
                 "    opts.push({ 'text': o.innerText, 'value': o.value, 'selected': o.selected }); " & _
                 "  } " & _
                 "  return JSON.stringify(opts); " & _
                 "}"
    
        Case "click_alert"
            ' Action: Clicks an element asynchronously using setTimeout to prevent VBA freezing
            ' when a modal alert/confirm/prompt appears.
            js = "function(e) { " & _
                 "  setTimeout(() => e.click(), 0); " & _
                 "  return 'done'; " & _
                 "}"

        Case "autoclicker"
            ' Logic: Sets up a MutationObserver to click an element as soon as it appears.
            ' Args: xpath, timeoutMs
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
' Internal Helper: Simplifies script.callFunction execution by handling Dictionary boilerplate.
' Arguments: js (JS code), args (Optional Collection), contextId (Optional), awaitPromise (Optional)
' ========================================================================================
Private Function CallScript(ByVal js As String, _
                           Optional ByVal args As Collection = Nothing, _
                           Optional ByVal contextId As String = "", _
                           Optional ByVal awaitPromise As Boolean = False) As String
    Dim params As New Dictionary, target As New Dictionary
    
    ' Ensure we have a valid collection for arguments
    If args Is Nothing Then Set args = New Collection
    
    ' Target the correct realm automatically
    target.Add "realm", GetScriptRealmId(contextId)
    
    ' Pack the BiDi parameters
    params.Add "functionDeclaration", js
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", awaitPromise
    
    ' Execute command and return raw JSON response
    CallScript = ExecuteBiDiCommand("script.callFunction", params)
End Function

' ========================================================================================
' IsTimeout: Precision-safe timeout check.
' ========================================================================================
Private Function IsTimeout(ByVal startTicks As Currency, ByVal timeoutMs As Double) As Boolean
    If timeoutMs <= 0 Then Exit Function
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
' Internal JS Helper: Centralizes SPA monitoring scripts (Probe).
' Includes Injection (setup), Restore (cleanup), and Query (status check).
' ========================================================================================
Private Function GetProbeJs(ByVal probeType As String) As String
    Dim js As String
    Select Case LCase(probeType)
        Case "inject"
            ' Main monitoring logic: Wraps Fetch, XHR, and MutationObserver
            Dim selArray As String: selArray = MergeToJsArray(DEFAULT_IGNORE_SELECTORS, p_userIgnoreSelectors)
            Dim netArray As String: netArray = MergeToJsArray(DEFAULT_IGNORE_NETWORK, p_userIgnoreNetworkPatterns)
            
            js = "(function(){" & _
                 "  if(window.__vbaSpaProbeInstalled && window.__vbaSpaProbeVersion === '" & PROBE_VERSION & "') return true;" & _
                 "  window.__vbaSpaProbeInstalled = true; window.__vbaSpaProbeVersion = '" & PROBE_VERSION & "';" & _
                 "  var ignoreSelectors = " & selArray & "; var ignoreNetwork = " & netArray & ";" & _
                 "  window.__vbaOriginals = window.__vbaOriginals || { fetch: window.fetch, xhrOpen: XMLHttpRequest.prototype.open, xhrSend: XMLHttpRequest.prototype.send, ws: window.WebSocket };" & _
                 "  var p = window.__vbaIdleProbe || { inflightFetchCount: 0, inflightXhrCount: 0, wsOpenCount: 0, lastMutationTs: performance.now() };" & _
                 "  window.__vbaIdleProbe = p; var updateTs = function(){ p.lastMutationTs = performance.now(); };" & _
                 "  try { new MutationObserver(function(mutations){ var hasRealChange = mutations.some(function(m){ return !ignoreSelectors.some(function(s){ return m.target.closest && m.target.closest(s); }); }); if(hasRealChange) updateTs(); }).observe(document, {childList:true, subtree:true, attributes:true}); } catch(e) {}" & _
                 "  if(typeof window.fetch === 'function' && !window.fetch.__vbaWrapped){ var origFetch = window.fetch; window.fetch = function(u, o){ var url = (typeof u === 'string') ? u : (u.url || ''); if(ignoreNetwork.some(function(pat){ return url.indexOf(pat) !== -1; })) return origFetch.apply(this, arguments); p.inflightFetchCount++; return origFetch.apply(this, arguments).finally(function(){ p.inflightFetchCount--; updateTs(); }); }; window.fetch.__vbaWrapped = true; }" & _
                 "  if(!XMLHttpRequest.prototype.__vbaWrapped){ XMLHttpRequest.prototype.open = function(m, u){ this.__vbaUrl = u; this.__vbaHooked = true; return window.__vbaOriginals.xhrOpen.apply(this, arguments); }; XMLHttpRequest.prototype.send = function(){ var self = this; var url = this.__vbaUrl || ''; if(ignoreNetwork.some(function(pat){ return url.indexOf(pat) !== -1; })){ self.__vbaHooked = false; return window.__vbaOriginals.xhrSend.apply(this, arguments); } p.inflightXhrCount++; var done = function(){ if(self.__vbaHooked){ p.inflightXhrCount--; self.__vbaHooked=false; } updateTs(); self.removeEventListener('loadend', done); }; this.addEventListener('loadend', done); return window.__vbaOriginals.xhrSend.apply(this, arguments); }; XMLHttpRequest.prototype.__vbaWrapped = true; }" & _
                 "  return true;" & _
                 "})();"

        Case "restore"
            ' Cleanup logic: Restores native browser APIs
            js = "(function(){" & _
                 "  var o = window.__vbaOriginals; if(!o) return;" & _
                 "  if(window.fetch && window.fetch.__vbaWrapped) window.fetch = o.fetch;" & _
                 "  if(XMLHttpRequest.prototype.__vbaWrapped){ XMLHttpRequest.prototype.open = o.xhrOpen; XMLHttpRequest.prototype.send = o.xhrSend; delete XMLHttpRequest.prototype.__vbaWrapped; }" & _
                 "  delete window.__vbaOriginals; delete window.__vbaSpaProbeInstalled; delete window.__vbaSpaProbeVersion; delete window.__vbaIdleProbe;" & _
                 "})();"

        Case "query"
            ' Status check logic: Returns JSON of current activity
            js = "function(){" & _
                 "  var p = window.__vbaIdleProbe; var v = window.__vbaSpaProbeVersion || '0.0';" & _
                 "  if(!p) return JSON.stringify({active:false, version: v});" & _
                 "  return JSON.stringify({ active: true, version: v, now: performance.now(), lastMutationTs: p.lastMutationTs, inflightFetchCount: p.inflightFetchCount, inflightXhrCount: p.inflightXhrCount });" & _
                 "}"
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
    Dim curTime As Currency
    If p_timerFreq = -1 Then
        ' Fallback to VBA.Timer (converted to quasi-ticks for consistency)
        GetRawTicks = CCur(VBA.Timer * 1000)
    Else
        QueryPerformanceCounter curTime
        GetRawTicks = curTime
    End If
End Function

' ========================================================================================
' [FIXED] GetElapsedMs with wraparound protection (High-Res & Timer Fallback)
' ========================================================================================
Private Function GetElapsedMs(ByVal startTicks As Currency) As Double
    Dim currentTicks As Currency
    Dim delta As Double
    
    currentTicks = GetRawTicks()

    If p_timerFreq = -1 Then
        ' --- VBA.Timer Fallback Mode ---
        delta = CDbl(currentTicks - startTicks)
        
        ' Handle midnight wraparound (Timer resets to 0 at 00:00:00)
        ' 86,400,000 ms = 24 hours
        If delta < 0 Then
            delta = delta + 86400000
            If DEBUG_MODE Then Debug.Print "BiDi Info: Timer wraparound detected and compensated."
        End If
        
        GetElapsedMs = delta
    Else
        ' --- High-Resolution (QPC) Mode ---
        Dim tickDelta As Currency
        tickDelta = currentTicks - startTicks
        
        ' QPC wraparound is extremely rare (takes years), but delta < 0 check is for safety
        If tickDelta < 0 Then
            GetElapsedMs = 999999 ' Force immediate timeout
            Exit Function
        End If
        
        ' Calculate precision milliseconds
        GetElapsedMs = (CDbl(tickDelta) / CDbl(p_timerFreq)) * 1000
    End If
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
