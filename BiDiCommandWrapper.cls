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
Private m_lastDoEventsTime As Double
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

Private Sub Class_Terminate()
    Set p_recorderLogs = Nothing
    If Not fso_ Is Nothing Then Set fso_ = Nothing
    Set socket_ = Nothing
    If DEBUG_MODE Then Debug.Print "BiDiCommandWrapper: Class Terminated. Resources released."
End Sub

Public Property Set SetSocket(socket As BiDiSocketCommunicator)
    Set socket_ = socket
End Property
' ========================================================================================
' GetTimeMs: Returns the current system time in milliseconds.
' Uses Win32 QueryPerformanceCounter for precision, falling back to VBA.Timer if needed.
' ========================================================================================
Private Function GetTimeMs() As Double
    If p_timerFreq = -1 Then
        ' Fallback: Use VBA's standard Timer function (less precise, but safe)
        ' Timer() returns seconds since midnight; multiply by 1000 for ms.
        GetTimeMs = VBA.Timer() * 1000
    Else
        ' High-Precision: Use the Win32 performance counter
        Dim curTime As Currency
        QueryPerformanceCounter curTime
        
        ' Convert currency value to milliseconds using the cached frequency
        GetTimeMs = (curTime / p_timerFreq) * 1000
    End If
End Function

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
Public Function GetIframeContextIdByUrl(ByVal partialUrl As String, Optional ByVal timeoutMs As Long = 0) As String
    Dim startTick As Double
    Dim currentId As String
    startTick = GetTimeMs()

    Do
        Dim params As New Dictionary
        ' WebDriver limit for tree depth
        params.Add "maxDepth", 10

        Dim strRes As String
        strRes = ExecuteBiDiCommand("browsingContext.getTree", params)

        Dim json As Object
        On Error Resume Next
        Set json = WebJsonConverter.ParseJson(strRes)
        On Error GoTo 0

        If Not json Is Nothing Then
            If json.Exists("result") Then
                If json("result").Exists("contexts") Then
                    Dim contexts As Collection
                    Set contexts = json("result")("contexts")
                    currentId = RecursiveFrameSearch(contexts, partialUrl)
                End If
            End If
        End If

        If currentId <> "" Then
            GetIframeContextIdByUrl = currentId
            Exit Function
        End If

        If timeoutMs <= 0 Then Exit Do

        If GetTimeMs() - startTick > timeoutMs Then
            Debug.Print "BiDi Info: Frame search timed out for '" & partialUrl & "'"
            Exit Do
        End If

        sleep SHORT_WAIT_MS
        DoEvents
    Loop

    GetIframeContextIdByUrl = ""
End Function

' [FIXED] Recursive search with depth limit to prevent stack overflow
Private Function RecursiveFrameSearch(ByVal contextList As Collection, _
                                      ByVal partialUrl As String, _
                                      Optional ByVal currentDepth As Long = 0) As String
    
    ' Guard: Recursion depth limit (VBA stack protection)
    Const MAX_RECURSION_DEPTH As Long = 20
    If currentDepth > MAX_RECURSION_DEPTH Then
        RecursiveFrameSearch = ""
        Exit Function
    End If
    
    Dim ctx As Dictionary
    Dim childRes As String

    For Each ctx In contextList
        If InStr(ctx("url"), partialUrl) > 0 Then
            RecursiveFrameSearch = ctx("context")
            Exit Function
        End If

        If ctx.Exists("children") Then
            If IsObject(ctx("children")) Then
                If Not ctx("children") Is Nothing Then
                    ' Recursive call with incremented depth
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
' Window & Tab Management (Refactored for Strict Evaluation)
' ========================================================================================
Public Function ExecuteFindWindowContextId(ByVal matchType As String, ByVal matchValue As String, _
                                           Optional ByVal timeoutMs As Long = 5000, _
                                           Optional ByVal waitForComplete As Boolean = True) As String
    Dim startTick As Double: startTick = GetTimeMs()
    Dim contexts As Collection, ctx As Dictionary, json As Object, strRes As String

    matchType = LCase(matchType)

    Do
        ' Get the current browsing context tree
        strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Set json = WebJsonConverter.ParseJson(strRes)

        If Not json Is Nothing Then
            Set contexts = json("result")("contexts")
            For Each ctx In contexts
                Dim isMatch As Boolean
                If matchType = "url" Then
                    ' Check if URL contains the match value
                    isMatch = (InStr(ctx("url"), matchValue) > 0)
                Else
                    ' Check if Title contains the match value
                    isMatch = (InStr(ExecuteGetTitleByContextId(ctx("context")), matchValue) > 0)
                End If

                If isMatch Then
                    If Not waitForComplete Then
                        ' Return context ID immediately if wait is not required
                        ExecuteFindWindowContextId = ctx("context"): Exit Function
                    Else
                        ' REFACTORED: Use GetScriptResultValue for strict boolean check of readyState
                        Dim readyRes As String
                        readyRes = ExecuteBiDiCommand("script.callFunction", BuildReadyStateParams(ctx("context")))
                        
                        ' Explicitly check if the script returned boolean True
                        If GetScriptResultValue(readyRes) = True Then
                            ExecuteFindWindowContextId = ctx("context"): Exit Function
                        End If
                    End If
                End If
            Next ctx
        End If

        ' Timeout protection
        If GetTimeMs() - startTick > timeoutMs Then Exit Do
        sleep SHORT_WAIT_MS: DoEvents
    Loop

    ExecuteFindWindowContextId = ""
End Function

Private Function BuildReadyStateParams(ByVal ctxId As String) As Dictionary
    Dim p As New Dictionary, t As New Dictionary
    t.Add "realm", GetScriptRealmId(ctxId, True)
    p.Add "functionDeclaration", "function(){return document.readyState==='complete'}"
    p.Add "arguments", New Collection
    p.Add "target", t
    p.Add "awaitPromise", False
    Set BuildReadyStateParams = p
End Function

Public Function ExecuteCreateContext(Optional ByVal createType As String = "tab") As String
    Dim params As New Dictionary
    params.Add "type", createType

    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.create", params)

    Dim json As Object
    Set json = WebJsonConverter.ParseJson(strRes)

    On Error Resume Next
    If json.Exists("result") Then
        ExecuteCreateContext = json("result")("context")
    End If
    On Error GoTo 0
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

    On Error Resume Next
    ExecuteGetTitleByContextId = WebJsonConverter.ParseJson(strRes)("result")("result")("value")
    On Error GoTo 0
End Function

' ========================================================================================
' Anti-Freeze Alert/Prompt Handling
' ========================================================================================
Public Function ExecuteClickAndHandleAlertByXPath(ByVal xpath As String, Optional ByVal Accept As Boolean = True, Optional ByVal contextId As String = "") As String
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, contextId)
    If sharedId = "" Then ExecuteClickAndHandleAlertByXPath = "{""error"":""element not found""}": Exit Function

    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    Dim params As New Dictionary, target As New Dictionary, args As New Collection, arg1 As New Dictionary

    Dim js As String
    js = "function(e) { setTimeout(function(){ e.click(); }, 0); return 'Click scheduled'; }"

    arg1.Add "sharedId", sharedId
    args.Add arg1
    target.Add "realm", realmId

    params.Add "functionDeclaration", js
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", False

    ExecuteClickAndHandleAlertByXPath = ExecuteBiDiCommand("script.callFunction", params)

    sleep SHORT_WAIT_MS
    DoEvents

    ExecuteHandleUserPrompt Accept, "", contextId
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
' Physical Input Simulation
' ========================================================================================
Public Function ExecuteInputClick(ByVal xpath As String, Optional ByVal contextId As String = "") As String
    Dim targetCtx As String
    If contextId <> "" Then targetCtx = contextId Else targetCtx = UpdateMainContextId()

    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, targetCtx)
    If sharedId = "" Then ExecuteInputClick = "{""error"":""element not found""}": Exit Function

    Dim params As New Dictionary
    params.Add "context", targetCtx

    Dim actions As New Collection, actionSeq As New Dictionary
    actionSeq.Add "type", "pointer"
    actionSeq.Add "id", "mouse_input_vba"

    Dim steps As New Collection
    Dim stepMove As New Dictionary, stepDown As New Dictionary, stepUp As New Dictionary

    stepMove.Add "type", "pointerMove"
    stepMove.Add "origin", BuildOriginObject(sharedId)
    stepMove.Add "x", 0: stepMove.Add "y", 0

    stepDown.Add "type", "pointerDown"
    stepDown.Add "button", 0

    stepUp.Add "type", "pointerUp"
    stepUp.Add "button", 0

    steps.Add stepMove: steps.Add stepDown: steps.Add stepUp
    actionSeq.Add "actions", steps
    actions.Add actionSeq

    params.Add "actions", actions
    ExecuteInputClick = ExecuteBiDiCommand("input.performActions", params)
End Function

Private Function BuildOriginObject(ByVal sharedId As String) As Dictionary
    Dim origin As New Dictionary, element As New Dictionary
    element.Add "sharedId", sharedId
    origin.Add "type", "element"
    origin.Add "element", element
    Set BuildOriginObject = origin
End Function

' ========================================================================================
' ExecuteBiDiWaitUntilUrlContains: Polls the current URL until it matches the target string.
' Features: Overflow-safe timeout calculation and background event processing.
' ========================================================================================
Public Function ExecuteBiDiWaitUntilUrlContains(ByVal partialUrl As String, _
                                                Optional ByVal timeoutSec As Long = 300, _
                                                Optional ByVal contextId As String = "") As Boolean
    Dim startTick As Double: startTick = GetTimeMs()
    Dim lastCheckTick As Double: lastCheckTick = GetTimeMs()
    
    ' [FIXED] Use Double to prevent integer overflow during calculation
    Dim timeoutMs As Double
    timeoutMs = CDbl(timeoutSec) * 1000
    
    Me.ExecuteSessionSubscribe Array("network.responseCompleted", "browsingContext.load", "browsingContext.fragmentNavigated")

    Debug.Print "BiDi: Waiting for URL containing '" & partialUrl & "'..."

    Do
        ProcessRecorderQueue
        
        If GetTimeMs() - lastCheckTick > 500 Then
            Dim currentUrl As String
            On Error Resume Next
            currentUrl = GetCurrentUrl(contextId)
            On Error GoTo 0
            
            If Len(currentUrl) > 0 And InStr(currentUrl, partialUrl) > 0 Then
                ExecuteBiDiWaitUntilUrlContains = True
                Exit Function
            End If
            lastCheckTick = GetTimeMs()
        End If

        ' [FIXED] Safe comparison with Double precision
        If GetTimeMs() - startTick > timeoutMs Then
            Debug.Print "BiDi Warning: URL wait timed out after " & timeoutSec & " seconds."
            ExecuteBiDiWaitUntilUrlContains = False
            Exit Function
        End If

        ThrottledDoEvents
        sleep 50
    Loop
End Function
' ========================================================================================
' [MODIFIED] High-Performance Navigate & Get Status
' waitForInteractive: If True, considers navigation complete at "interactive" state (DOM ready),
'                     skipping image/resource loading. Default is False for backward compatibility.
' ========================================================================================
Public Function ExecuteNavigateAndGetStatus(ByVal targetUrl As String, _
                                            Optional ByVal waitNetworkIdle As Boolean = True, _
                                            Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                            Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                            Optional ByVal contextId As String = "", _
                                            Optional ByVal waitForInteractive As Boolean = True) As Long
    p_pendingRequestCount = 0
    ClearState
    Me.ExecuteSessionSubscribe Array("network.beforeRequestSent", "network.responseCompleted", _
                                   "browsingContext.load", "browsingContext.domContentLoaded")

    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId(True))
    
    ' Trigger navigation without waiting for the browser's internal 'complete'
    Dim params As New Dictionary: params.Add "context", targetCtx: params.Add "url", targetUrl: params.Add "wait", "none"
    ExecuteBiDiCommand "browsingContext.navigate", params

    Dim startTick As Double: startTick = GetTimeMs()
    Dim foundStatus As Long: foundStatus = 0
    Dim isLoadComplete As Boolean: isLoadComplete = False
    Dim lastReadyCheck As Double: lastReadyCheck = 0

    ' Monitoring loop for status code and readyState
    Do While GetTimeMs() - startTick < DEFAULT_NAV_TIMEOUT_MS
        ProcessRecorderQueue ' Handles BiDi events and updates pending requests
        
        ' Check readyState periodically (every 500ms)
        If Not isLoadComplete And (GetTimeMs() - lastReadyCheck > 500) Then
            Dim readyState As String: readyState = ExecuteGetReadyState(targetCtx)
            If readyState = "complete" Or (waitForInteractive And readyState = "interactive") Then
                isLoadComplete = True
                If foundStatus = 0 Then foundStatus = 200 ' Fallback
            End If
            lastReadyCheck = GetTimeMs()
        End If

        If isLoadComplete Then Exit Do
        
        ThrottledDoEvents
        sleep 50
    Loop
    
    ' Proceed to SPA-specific idle wait if required
    If waitNetworkIdle Then
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx
    End If

    ExecuteNavigateAndGetStatus = foundStatus
End Function

Private Function ExecuteGetReadyState(ByVal ctxId As String) As String
    Dim js As String: js = "function(){ return document.readyState; }"
    Dim p As New Dictionary, t As New Dictionary
    t.Add "context", ctxId
    p.Add "functionDeclaration", js
    p.Add "target", t
    p.Add "awaitPromise", False

    Dim ret As String: ret = ExecuteBiDiCommand("script.callFunction", p)
    On Error Resume Next
    ExecuteGetReadyState = WebJsonConverter.ParseJson(ret)("result")("result")("value")
    On Error GoTo 0
End Function

' ========================================================================================
' ExecuteEnableResourceBlocking: Sets URL patterns to be blocked via CDP.
' ========================================================================================
Public Function ExecuteEnableResourceBlocking(ByVal patterns As Variant) As String
    Dim params As New Dictionary
    Dim urlList As New Collection
    Dim Item As Variant

    If IsArray(patterns) Then
        For Each Item In patterns
            urlList.Add CStr(Item)
        Next Item
    Else
        urlList.Add CStr(patterns)
    End If

    params.Add "urls", urlList

    ' Handle the structural result from ExecuteCDPCommand
    Dim resObj As Object
    Set resObj = ExecuteCDPCommand("Network.setBlockedURLs", params)
    
    If resObj("Ok") Then
        ExecuteEnableResourceBlocking = resObj("Json")
    Else
        ExecuteEnableResourceBlocking = "{""error"":""CDP_FAILURE"",""message"":""" & _
                                        Replace(resObj("Error"), """", "'") & """}"
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
    If sharedId = "" Then ExecuteIsElementVisible = False: Exit Function

    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    Dim params As New Dictionary, target As New Dictionary, args As New Collection, arg1 As New Dictionary

    arg1.Add "sharedId", sharedId
    args.Add arg1
    target.Add "realm", realmId
    params.Add "functionDeclaration", GetJsForVisibilityCheck()
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", False

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.callFunction", params)

    On Error Resume Next
    Dim resJson As Object: Set resJson = WebJsonConverter.ParseJson(strRes)
    Dim resVal As Variant

    If Not resJson Is Nothing Then
        If resJson.Exists("result") Then
            If resJson("result").Exists("result") Then
                If resJson("result")("result").Exists("value") Then
                    resVal = resJson("result")("result")("value")
                    ExecuteIsElementVisible = CBool(resVal)
                End If
            End If
        End If
    End If
    On Error GoTo 0
End Function

' ========================================================================================
' React/Async Select Handling
' ========================================================================================
Public Function ExecuteWaitForSelectOptions(ByVal xpath As String, Optional ByVal minCount As Long = 1, Optional ByVal timeoutMs As Long = 5000, Optional ByVal contextId As String = "") As Boolean
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, contextId)
    If sharedId = "" Then Exit Function

    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    Dim params As New Dictionary, target As New Dictionary, args As New Collection
    Dim arg1 As New Dictionary, arg2 As New Dictionary, arg3 As New Dictionary

    arg1.Add "sharedId", sharedId
    arg2.Add "type", "number": arg2.Add "value", minCount
    arg3.Add "type", "number": arg3.Add "value", timeoutMs
    args.Add arg1: args.Add arg2: args.Add arg3

    target.Add "realm", realmId
    params.Add "functionDeclaration", GetJsForWaitOptions()
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", True

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.callFunction", params)

    On Error Resume Next
    ExecuteWaitForSelectOptions = WebJsonConverter.ParseJson(strRes)("result")("result")("value")
    On Error GoTo 0
End Function

Public Function ExecuteGetSelectOptions(ByVal xpath As String, Optional ByVal contextId As String = "") As Collection
    Dim resultColl As New Collection
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, contextId)
    If sharedId = "" Then Set ExecuteGetSelectOptions = resultColl: Exit Function

    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    Dim params As New Dictionary, target As New Dictionary, args As New Collection, arg1 As New Dictionary

    arg1.Add "sharedId", sharedId
    args.Add arg1
    target.Add "realm", realmId
    params.Add "functionDeclaration", GetJsForGetOptions()
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", False

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.callFunction", params)

    On Error Resume Next
    Dim innerJsonStr As String
    innerJsonStr = WebJsonConverter.ParseJson(strRes)("result")("result")("value")
    If innerJsonStr <> "" Then Set resultColl = WebJsonConverter.ParseJson(innerJsonStr)
    On Error GoTo 0
    Set ExecuteGetSelectOptions = resultColl
End Function

' ========================================================================================
' Standard Commands
' ========================================================================================
Public Function ExecuteSessionSubscribe(ByVal events As Variant) As String
    ' Guard: Skip if already subscribed to prevent redundant protocol overhead
    If m_isSubscribed Then
        ExecuteSessionSubscribe = "{""result"":{},""type"":""success""}"
        Exit Function
    End If

    Dim params As New Dictionary
    If IsArray(events) Then
        params.Add "events", events
    Else
        Dim arr(0) As String: arr(0) = events
        params.Add "events", arr
    End If

    ExecuteSessionSubscribe = ExecuteBiDiCommand("session.subscribe", params)
    
    ' Validate response and set subscription flag
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
' Stale Element Retry Runner
' ========================================================================================
Private Function ExecuteActionWithStaleRetry(ByVal xpath As String, _
                                             ByVal jsDeclaration As String, _
                                             Optional ByVal secondArgValue As Variant = Empty, _
                                             Optional ByVal awaitPromise As Boolean = True, _
                                             Optional ByVal contextId As String = "", _
                                             Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    Dim retryCount As Integer
    Dim sharedId As String, strRes As String, realmId As String
    Dim args As Collection, arg1 As Dictionary, arg2 As Dictionary
    Dim params As Dictionary, target As Dictionary
    Dim jsonResponse As Object
    Dim errType As String, errMsg As String

    sharedId = GetSharedIdFromXPath(xpath, contextId)

    For retryCount = 0 To maxRetries
        Set args = New Collection: Set arg1 = New Dictionary
        Set params = New Dictionary: Set target = New Dictionary

        If sharedId = "" And retryCount > 0 Then sharedId = GetSharedIdFromXPath(xpath, contextId)

        If sharedId = "" Then
            strRes = "{""error"":""element not found"", ""message"":""XPath lookup failed""}"
            If retryCount = maxRetries Then Exit For
        Else
            realmId = GetScriptRealmId(contextId)
            arg1.Add "sharedId", sharedId
            args.Add arg1

            If Not IsEmpty(secondArgValue) Then
                Set arg2 = New Dictionary
                arg2.Add "type", "string"
                arg2.Add "value", CStr(secondArgValue)
                args.Add arg2
            End If

            target.Add "realm", realmId
            params.Add "functionDeclaration", jsDeclaration
            params.Add "arguments", args
            params.Add "target", target
            params.Add "awaitPromise", awaitPromise

            On Error Resume Next
            strRes = ""
            strRes = ExecuteBiDiCommand("script.callFunction", params)
            If Err.Number <> 0 Then
                strRes = "{""error"":""VBA_Runtime_Error"",""message"":""" & Replace(Err.Description, """", "'") & """}"
                Err.Clear
            End If
            On Error GoTo 0
        End If

        Dim isStale As Boolean: isStale = False
        Set jsonResponse = Nothing

        If Len(strRes) > 0 And (Left(strRes, 1) = "{" Or Left(strRes, 1) = "[") Then
            On Error Resume Next
            Set jsonResponse = WebJsonConverter.ParseJson(strRes)
            On Error GoTo 0
        End If

        If Not jsonResponse Is Nothing Then
            If jsonResponse.Exists("error") Then
                errType = jsonResponse("error")
                If jsonResponse.Exists("message") Then errMsg = jsonResponse("message")
                If IsIgnorableError(errType, errMsg) Then isStale = True
            End If
            If Not isStale And jsonResponse.Exists("result") Then
                If jsonResponse("result").Exists("exceptionDetails") Then
                    Dim eText As String
                    On Error Resume Next
                    eText = jsonResponse("result")("exceptionDetails")("text")
                    On Error GoTo 0
                    If IsIgnorableError("", eText) Then isStale = True
                End If
            End If
        Else
            If InStr(strRes, "VBA_Runtime_Error") > 0 Then isStale = True
        End If

        If isStale Then
            If retryCount < maxRetries Then
                If DEBUG_MODE Then Debug.Print "BiDi Info: Stale detected (Try " & retryCount & "). Retrying..."
                GetScriptRealmId contextId, True
                sleep 200
                DoEvents
            Else
                Exit For
            End If
        Else
            Exit For
        End If
    Next retryCount

    If InStr(strRes, """error""") > 0 Or InStr(strRes, "exceptionDetails") > 0 Then
        Dim finalMsg As String
        finalMsg = "BiDi Action Failed on XPath: " & xpath & vbCrLf & "Response: " & strRes
        Err.Raise vbObjectError + 555, "BiDiCommandWrapper.ExecuteAction", finalMsg
    End If

    ExecuteActionWithStaleRetry = strRes
End Function

' ========================================================================================
' Enhanced Click Functions
' ========================================================================================
Public Function ExecuteClickByXPath(ByVal xpath As String, _
                                    Optional ByVal contextId As String = "", _
                                    Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    ExecuteClickByXPath = ExecuteActionWithStaleRetry(xpath, GetJsForClickAndWait(), , True, contextId, maxRetries)
End Function

' ========================================================================================
' ExecuteClickAndWaitByXPath: Clicks an element and synchronizes with page navigation.
' Balanced: Waits briefly for a URL change to avoid race conditions, but proceeds
' quickly if the navigation is handled by SPA or is instantaneous.
' ========================================================================================
Public Function ExecuteClickAndWaitByXPath(ByVal xpath As String, _
                                           Optional ByVal waitNetworkIdle As Boolean = True, _
                                           Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                           Optional ByVal contextId As String = "", _
                                           Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                           Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS) As String
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())
    Dim oldUrl As String: oldUrl = GetCurrentUrl(targetCtx)

    ' 1. Execute Click
    Dim res As String
    res = ExecuteActionWithStaleRetry(xpath, GetJsForClickAndWait(), , True, targetCtx, maxRetries)

    ' 2. Synchronize with Navigation
    If waitNetworkIdle And IsScriptDone(res) Then
        Dim startTick As Double: startTick = GetTimeMs()
        
        ' --- Proactive Transition Check ---
        ' Wait up to 500ms only if the URL hasn't changed yet.
        Do
            Dim currentUrl As String: currentUrl = GetCurrentUrl(targetCtx)
            
            ' If URL has changed OR the context is reporting blank (start of reload),
            ' we are ready to refresh our IDs.
            If currentUrl <> oldUrl Or currentUrl = "" Then Exit Do
            
            If GetTimeMs() - startTick > 500 Then Exit Do ' Cap the wait
            sleep 50: DoEvents
        Loop

        ' 3. Refresh IDs and re-inject probe
        UpdateMainContextId True
        InitializeSpaProbe targetCtx
        
        ' 4. Final network idle wait
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx
    End If

    ExecuteClickAndWaitByXPath = res
End Function

Public Function ExecuteClickAndNavigateByXPath(ByVal xpath As String, _
                                               Optional ByVal waitNetworkIdle As Boolean = True, _
                                               Optional ByVal contextId As String = "", _
                                               Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                               Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                               Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS) As String
    Dim success As Boolean
    Dim clickRes As String
    Dim attempt As Long

    ' Retry loop for initial click action
    For attempt = 0 To maxRetries
        On Error Resume Next
        clickRes = Me.ExecuteClickByXPath(xpath, contextId, 0)
        
        ' NEW: Strict validation of script return value
        If IsScriptDone(clickRes) Then
            success = True
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0
        If attempt < maxRetries Then sleep 100
    Next attempt

    If Not success Then
        Err.Raise vbObjectError + 500, "ExecuteClickAndNavigateByXPath", "Click failed or 'done' not received. XPath: " & xpath
    End If

    ' Follow navigation context and wait for idle
    Dim newCtxId As String
    newCtxId = UpdateMainContextId(True)

    If waitNetworkIdle Then
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, newCtxId
    End If

    ExecuteClickAndNavigateByXPath = Me.GetCurrentUrl(newCtxId)
End Function

' ========================================================================================
' ExecuteInputValueAndWaitByXPath: Inputs text with "Early Exit" strategy.
' Optimization: Instead of waiting for a full network idle (which takes seconds on
' ServiceNow), it starts typing as soon as the target element becomes available.
' ========================================================================================
Public Function ExecuteInputValueAndWaitByXPath(ByVal xpath As String, _
                                                ByVal valueToSet As String, _
                                                Optional ByVal contextId As String = "", _
                                                Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                                Optional ByVal waitNetworkIdle As Boolean = True, _
                                                Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String
    
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())
    
    ' 1. Start Probe BEFORE full idle wait to capture immediate state changes
    If waitNetworkIdle Then InitializeSpaProbe targetCtx

    ' 2. [EARLY EXIT GUARD]
    ' Don't wait for analytics/logs to finish. If the input field exists, use it now.
    Dim startTick As Double: startTick = GetTimeMs()
    Do
        Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, targetCtx)
        ' If element is found, we can stop waiting and proceed to input.
        If sharedId <> "" Then Exit Do
        
        ' Limit wait to 5s if element doesn't appear.
        If GetTimeMs() - startTick > 5000 Then Exit Do: ProcessRecorderQueue: sleep 50: DoEvents
    Loop

    ' 3. Perform the actual input
    Dim res As String
    res = ExecuteActionWithStaleRetry(xpath, GetJsForInputAndWait(), valueToSet, True, targetCtx, maxRetries)

    ' 4. Final short idle check after input
    If waitNetworkIdle And IsScriptDone(res) Then
        WaitForNetworkIdleSpa minStableMs, 3000, targetCtx ' Shorter timeout for input feedback
    End If

    ExecuteInputValueAndWaitByXPath = res
End Function

' ========================================================================================
' ExecuteSelectValueAndWaitByXPath: Selects an option and waits for async tasks.
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
        ' Drain any existing background events to reset the activity counter
        ProcessRecorderQueue
    End If

    ' 2. Prepare Select Logic
    Dim jsFunction As String
    If selectByText Then
        jsFunction = GetJsForSelectTextAndWait()
    Else
        jsFunction = GetJsForSelectAndWait()
    End If

    ' 3. Execute UI Action
    ' This triggers the 'triggerLongFetch()' in your page.
    Dim res As String
    res = ExecuteActionWithStaleRetry(xpath, jsFunction, valueOrText, True, targetCtx, maxRetries)

    ' 4. Synchronization Wait
    If waitNetworkIdle And IsScriptDone(res) Then
        ' Immediate drain to catch the 'beforeRequestSent' event of the long-running fetch
        ProcessRecorderQueue
        
        ' Enter the stability loop (will now see inflightFetchCount: 1)
        WaitForNetworkIdleSpa minStableMs, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
    End If

    ExecuteSelectValueAndWaitByXPath = res
End Function

Public Function ExecuteGetTextByXPath(ByVal xpath As String, Optional ByVal contextId As String = "", Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    Dim jsFunction As String
    jsFunction = "function(element) { return element.innerText || element.textContent || ''; }"
    Dim strRes As String
    strRes = ExecuteActionWithStaleRetry(xpath, jsFunction, , False, contextId, maxRetries)

    On Error Resume Next
    ExecuteGetTextByXPath = WebJsonConverter.ParseJson(strRes)("result")("result")("value")
    On Error GoTo 0
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

Private Function GetSharedIdFromXPath(ByVal xpath As String, Optional ByVal contextId As String = "") As String
    Dim strRes As String
    strRes = ExecuteBrowsingContextLocateNodes(xpath, contextId)

    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(strRes)
    On Error GoTo 0

    If Not json Is Nothing Then
        If json.Exists("result") Then
            If json("result").Exists("nodes") Then
                Dim nodes As Collection
                Set nodes = json("result")("nodes")
                If nodes.Count > 0 Then
                    If nodes(1).Exists("sharedId") Then
                        GetSharedIdFromXPath = nodes(1)("sharedId")
                        Exit Function
                    End If
                End If
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
    Dim realmId As String
    On Error Resume Next
    realmId = GetScriptRealmId(contextId)
    If realmId = "" Or Err.Number <> 0 Then
        GetCurrentUrl = "": Exit Function
    End If

    Dim params As New Dictionary, target As New Dictionary
    target.Add "realm", realmId
    params.Add "functionDeclaration", "function(){ return window.location.href; }"
    params.Add "arguments", New Collection
    params.Add "target", target
    params.Add "awaitPromise", False

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.callFunction", params)

    Dim json As Object
    Set json = WebJsonConverter.ParseJson(strRes)
    GetCurrentUrl = json("result")("result")("value")
    
    If Err.Number <> 0 Then GetCurrentUrl = ""
    On Error GoTo 0
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
' ExecuteSetDownloadPath: Sets the directory for browser downloads.
' ========================================================================================
Public Function ExecuteSetDownloadPath(ByVal downloadPath As String) As String
    p_currentDownloadPath = downloadPath
    
    Dim cdpParams As New Dictionary
    cdpParams.Add "behavior", "allow"
    cdpParams.Add "downloadPath", downloadPath
    cdpParams.Add "eventsEnabled", True
    
    Dim resObj As Object
    Set resObj = ExecuteCDPCommand("Page.setDownloadBehavior", cdpParams, True)
    
    If resObj("Ok") Then
        ExecuteSetDownloadPath = resObj("Json")
    Else
        ExecuteSetDownloadPath = "{""error"":""CDP_FAILURE"",""message"":""" & _
                                 Replace(resObj("Error"), """", "'") & """}"
    End If
End Function

' ========================================================================================
' ExecuteWaitForDownloadCompletion: Monitors the file system for download finishing.
' Optimized: Throttles heavy Disk I/O (1s interval) while keeping BiDi events fluid (50ms).
' ========================================================================================
Public Function ExecuteWaitForDownloadCompletion(ByVal startTime As Date, _
                                                  Optional ByVal timeoutSec As Long = 30) As Boolean
    Dim startTick As Double: startTick = GetTimeMs()
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
        If GetTimeMs() - startTick > timeoutMs Then
            Debug.Print "BiDi Warning: Download monitoring timed out after " & timeoutSec & " seconds."
            ExecuteWaitForDownloadCompletion = False
            Exit Function
        End If

        ' 3. Heavy Processing: File System Scan (executed only once per interval)
        If GetTimeMs() - lastScanTick > FILE_SCAN_INTERVAL_MS Then
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
            lastScanTick = GetTimeMs()
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
' Shadow DOM Clicks
' ========================================================================================
Public Function ExecuteShadowClickAndWait(ByVal selectorsArray As Variant, _
                                          Optional ByVal waitNetworkIdle As Boolean = True, _
                                          Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                          Optional ByVal contextId As String = "", _
                                          Optional ByVal searchTimeoutMs As Long = 5000) As String
    ' --- NEW: Enable monitoring BEFORE shadow interaction ---
    If waitNetworkIdle Then InitializeSpaProbe contextId

    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    Dim params As New Dictionary, target As New Dictionary, args As New Collection

    args.Add BuildShadowSelectorArg(selectorsArray)
    Dim arg2 As New Dictionary: arg2.Add "type", "number": arg2.Add "value", searchTimeoutMs
    args.Add arg2

    target.Add "realm", realmId
    params.Add "functionDeclaration", GetJsForShadowClick()
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", True

    Dim res As String
    res = ExecuteBiDiCommand("script.callFunction", params)

    If InStr(res, "Error: Element not found") > 0 Or InStr(res, """error""") > 0 Then
         Err.Raise vbObjectError + 557, "ExecuteShadowClickAndWait", "Shadow DOM element not found."
    End If

    If waitNetworkIdle And IsScriptDone(res) Then
        Dim targetCtx As String
        If contextId <> "" Then targetCtx = contextId Else targetCtx = p_mainContextId
        WaitForNetworkIdleSpa minStableMs, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
    End If

    ExecuteShadowClickAndWait = res
End Function

' ========================================================================================
' ExecuteShadowClickAndNavigate: Specialized for high-speed redirect handling.
' Logic: Polling for URL change immediately after shadow-click to intercept redirects
' without waiting for 100ms recovery cycles, ensuring faster landing on login pages.
' ========================================================================================
Public Function ExecuteShadowClickAndNavigate(ByVal selectorsArray As Variant, _
                                              Optional ByVal waitNetworkIdle As Boolean = True, _
                                              Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                              Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                              Optional ByVal contextId As String = "", _
                                              Optional ByVal searchTimeoutMs As Long = 5000) As String
    
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId(True))
    Dim oldUrl As String: oldUrl = GetCurrentUrl(targetCtx)
    Dim realmId As String: realmId = GetScriptRealmId(targetCtx)
    
    ' 1. Prepare Shadow Click Command
    Dim params As New Dictionary, target As New Dictionary, args As New Collection
    args.Add BuildShadowSelectorArg(selectorsArray)
    Dim argSearch As New Dictionary: argSearch.Add "type", "number": argSearch.Add "value", searchTimeoutMs
    args.Add argSearch
    target.Add "realm", realmId
    params.Add "functionDeclaration", GetJsForShadowClick()
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", True
    
    ' 2. Execute Click
    Dim res As String: res = ExecuteBiDiCommand("script.callFunction", params)
    
    ' 3. [PROACTIVE NAVIGATION GUARD]
    ' Catch the start of SSO/SAML redirect chains immediately.
    If waitNetworkIdle And IsScriptDone(res) Then
        Dim startTick As Double: startTick = GetTimeMs()
        Do
            Dim currentUrl As String: currentUrl = GetCurrentUrl(targetCtx)
            ' If URL changes or blanks, the redirect has started.
            If currentUrl <> oldUrl Or currentUrl = "" Then Exit Do
            If GetTimeMs() - startTick > 500 Then Exit Do
            sleep 30: DoEvents
        Loop

        ' Refresh the environment to the first landing page of the redirect.
        UpdateMainContextId True
        InitializeSpaProbe targetCtx
        
        ' Wait for network/DOM to stabilize on the login page.
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx
    End If
    
    ExecuteShadowClickAndNavigate = Me.GetCurrentUrl(targetCtx)
End Function

' ========================================================================================
' Core Command Helper: Handles outgoing BiDi commands with self-healing recovery logic
' ========================================================================================
Public Function ExecuteBiDiCommand(ByVal method As String, ByVal params As Dictionary) As String
    If socket_ Is Nothing Then Err.Raise 91, , "Socket object is not set."

    Dim retryCount As Integer
    Dim strRes As String, sendMsg As String
    Dim currentParams As Dictionary
    Dim dic As Dictionary
    
    ' Default to the original params for the first attempt
    Set currentParams = params

    For retryCount = 0 To MAX_CONTEXT_RECOVERY_RETRIES
        ' 1. Prepare the JSON request packet
        Set dic = New Dictionary
        dic.Add "id", p_nextRequestId
        dic.Add "method", method
        dic.Add "params", currentParams
        p_nextRequestId = p_nextRequestId + 1

        ' 2. Serialize and Log the request
        sendMsg = WebJsonConverter.ConvertToJson(dic)
        If DEBUG_MODE Then DebugPrintLong "Request (Try " & retryCount & ")", sendMsg

        ' 3. Execute synchronous Send and Receive
        strRes = socket_.SendAndReceive(sendMsg)
        If DEBUG_MODE Then DebugPrintLong "Response (Try " & retryCount & ")", strRes

        ' 4. Detect Context/Realm Loss
        Dim isContextDestroyed As Boolean
        isContextDestroyed = IsContextLostError(strRes)

        ' 5. Recovery Logic (Only triggered on context loss errors)
        If isContextDestroyed And Not m_isRecovering And retryCount < MAX_CONTEXT_RECOVERY_RETRIES Then
            On Error GoTo RecoveryCleanup
            m_isRecovering = True

            ' --- Exponential Backoff ---
            ' Progressively increase wait time to allow the browser to stabilize (300ms, 600ms...)
            Dim waitMs As Long
            waitMs = 100 * (2 ^ retryCount)
            If waitMs > 2000 Then waitMs = 2000
            
            Debug.Print "BiDi Info: Context lost. Recovery attempt " & retryCount + 1 & " in " & waitMs & "ms"
            sleep waitMs
            DoEvents

            ' --- State Restoration ---
            ClearState ' Clear cached IDs
            Dim newCtx As String
            newCtx = UpdateMainContextId(True) ' Re-identify the active browsing context

            If newCtx = "" Then
                Debug.Print "BiDi Error: Recovery failed - No active context found."
                m_isRecovering = False
                CheckAndRaiseError strRes, method
                Exit For
            End If

            ' --- Deep Copy & Parameter Update ---
            ' Prevent side-effects on the original 'params' object by creating a deep clone
            Dim retryParams As Dictionary
            Set retryParams = DeepCopyDictionary(params)

            ' Safely update the context and realm IDs within the cloned parameters
            If Not UpdateContextAndRealm(retryParams, newCtx) Then
                Debug.Print "BiDi Error: Failed to refresh context/realm IDs for retry."
                m_isRecovering = False
                CheckAndRaiseError strRes, method
                Exit For
            End If

            ' Switch to the updated parameters for the next iteration of the loop
            Set currentParams = retryParams
            m_isRecovering = False
            ' Loop continues to re-send the command
            
        Else
            ' 6. Standard Error Checking or Success Response
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

Private Function UpdateContextAndRealm(ByRef params As Dictionary, ByVal newCtx As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Update the root-level context ID if it exists
    If params.Exists("context") Then
        params("context") = newCtx
    End If
    
    ' Update the nested context/realm IDs inside the 'target' object
    If params.Exists("target") Then
        Dim targetObj As Dictionary
        Set targetObj = params("target")
        
        ' Update target.context
        If targetObj.Exists("context") Then
            targetObj("context") = newCtx
        End If
        
        ' Update target.realm (requires fetching a new realm ID for the new context)
        If targetObj.Exists("realm") Then
            Dim newRealm As String
            newRealm = GetScriptRealmId(newCtx, True)
            
            If newRealm = "" Then
                UpdateContextAndRealm = False
                Exit Function
            End If
            
            targetObj("realm") = newRealm
        End If
    End If
    
    UpdateContextAndRealm = True
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
                result.Add DeepCopyDictionary(Item, currentDepth)
            ElseIf TypeName(Item) = "Collection" Then
                result.Add DeepCopyCollection(Item, currentDepth)
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

Private Sub CheckAndRaiseError(ByVal jsonString As String, ByVal methodName As String)
    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonString)
    On Error GoTo 0

    If json Is Nothing Then Exit Sub
    Dim errType As String, errMsg As String

    ' A. Handle Protocol Level Errors
    If json.Exists("error") Then
        errType = json("error")
        If json.Exists("message") Then errMsg = json("message")
        If IsIgnorableError(errType, errMsg) Then Exit Sub
        Err.Raise vbObjectError + 513, "BiDi:" & methodName, "Browser Protocol Error [" & errType & "]: " & errMsg
    End If

    ' B. Handle JavaScript Execution Errors (Detailed)
    If json.Exists("result") Then
        If json("result").Exists("exceptionDetails") Then
            Dim details As Object: Set details = json("result")("exceptionDetails")
            
            ' Use the new formatter for rich error reporting
            Dim detailedError As String
            detailedError = FormatJavaScriptError(details)
            
            ' Output to Immediate Window for debugging
            If DEBUG_MODE Then
                Debug.Print "===================================================="
                Debug.Print "BIDI JS ERROR REPORT"
                Debug.Print "Method: " & methodName
                Debug.Print detailedError
                Debug.Print "===================================================="
            End If
            
            ' Raise VBA error with the formatted message
            Err.Raise vbObjectError + 514, "BiDi:JS_Exception", "JS Error in " & methodName & ": " & details("text") & " (See Immediate Window for stack trace)"
        End If
    End If
End Sub

' ========================================================================================
' Internal Helper: Log Printer (With Timestamp & High-Res Timer)
' ========================================================================================
Private Sub DebugPrintLong(ByVal header As String, ByVal longText As String)
    Const CHUNK_SIZE As Integer = 250
    Dim i As Long
    Dim cleanText As String
    Dim timePrefix As String
    Dim currentTick As Double
    
    currentTick = GetTimeMs()
    timePrefix = "[" & Format(Now, "hh:mm:ss") & " | " & Format(currentTick, "0") & "] "

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

    Debug.Print timePrefix & header & " (Length: " & Len(cleanText) & ")"

    If Len(cleanText) > 0 Then
        For i = 1 To Len(cleanText) Step CHUNK_SIZE
            Debug.Print Mid$(cleanText, i, CHUNK_SIZE);
        Next i
    End If

    Debug.Print ""
End Sub

' ========================================================================================
' ClearState: Resets the internal state and cleans up browser-side side effects.
' ========================================================================================
Public Sub ClearState()
    ' 1. [CRITICAL] Restore original browser APIs to remove side effects (‰ð“Å)
    ' This ensures that window.fetch and XMLHttpRequest are returned to their native state.
    If m_isProbeRegistered Then
        On Error Resume Next
        Dim t As New Dictionary: t.Add "context", UpdateMainContextId()
        Dim p As New Dictionary
        p.Add "expression", GetProbeRestoreExpression()
        p.Add "target", t
        p.Add "awaitPromise", False
        ExecuteBiDiCommand "script.evaluate", p
        On Error GoTo 0
    End If

    ' 2. Clear cached identification and tracking variables
    p_mainRealmId = ""
    p_cdpSessionId = ""
    p_mainContextId = ""
    p_pendingRequestCount = 0

    ' 3. Reset registration flags to allow fresh injection/subscription in the next session
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

    Dim startTick As Double: startTick = GetTimeMs()
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
        If GetTimeMs() - startTick > 2000 Then Exit Do
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
' ExecuteRegisterAutoClickerByXPath: Registers a background observer to click a button.
' [FIX - PRIORITY A] Added a lifecycle timeout to prevent CPU leaks in SPAs.
' Parameters:
'   - xpath: Target element to click.
'   - timeoutMs: Max time to keep the observer alive (default 30000ms).
' ========================================================================================
' ========================================================================================
' ExecuteRegisterAutoClickerByXPath: Registers a background observer to click a button.
' [FIXED] Uses the existing WebJsonConverter.ConvertToJson for safe string escaping.
' Parameters:
'   - xpath: Target element to click.
'   - timeoutMs: Max time to keep the observer alive (default 30000ms).
' ========================================================================================
Public Sub ExecuteRegisterAutoClickerByXPath(ByVal xpath As String, _
                                             Optional ByVal timeoutMs As Long = 30000)
    Dim js As String
    js = "(function() {"
    
    ' 1. Use the pre-declared WebJsonConverter to escape the XPath string safely.
    ' This will automatically handle quotes and backslashes.
    js = js & "  var xpath = " & WebJsonConverter.ConvertToJson(xpath) & ";"
    js = js & "  var timeout = " & timeoutMs & ";"
    js = js & "  var start = Date.now();"
    
    ' Helper to find and click the element
    js = js & "  function checkAndClick() {"
    js = js & "    try {"
    js = js & "      var res = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);"
    js = js & "      var btn = res.singleNodeValue;"
    js = js & "      if (btn && btn.offsetParent !== null) {"
    js = js & "        btn.click();"
    js = js & "        return true;"
    js = js & "      }"
    js = js & "    } catch(e) {}"
    js = js & "    return false;"
    js = js & "  };"
    
    ' 2. Immediate check
    js = js & "  if (checkAndClick()) return;"
    
    ' 3. Setup MutationObserver
    js = js & "  var observer = new MutationObserver(function(mutations, obs) {"
    js = js & "    if (checkAndClick()) {"
    js = js & "      obs.disconnect(); return;"
    js = js & "    }"
    ' [LIFECYCLE LIMIT] Auto-disconnect to prevent CPU leaks in long SPA sessions
    js = js & "    if (Date.now() - start > timeout) {"
    js = js & "      obs.disconnect();"
    js = js & "      console.log('VBA AutoClicker: Timeout reached. Observer disconnected.');"
    js = js & "    }"
    js = js & "  });"
    
    ' 4. Monitor the entire document for structural changes
    js = js & "  observer.observe(document, { childList: true, subtree: true });"
    
    ' 5. [HARD SAFETY] Secondary cleanup timer
    js = js & "  setTimeout(function() { observer.disconnect(); }, timeout + 1000);"
    
    js = js & "})();"

    ' Register via CDP to ensure persistence during SPA-style navigation
    Dim cdpParams As New Dictionary
    cdpParams.Add "source", js
    
    ' Use structural error handling (Dictionary package)
    Dim result As Object
    Set result = Me.ExecuteCDPCommand("Page.addScriptToEvaluateOnNewDocument", cdpParams)
    
    If Not result("Ok") Then
        Debug.Print "BiDi Warning: AutoClicker registration failed: " & result("Error")
    End If
End Sub

' ========================================================================================
' JavaScript Generators
' ========================================================================================
Private Function GetJsForVisibilityCheck() As String
    Dim js As String
    js = "function(el) { "
    js = js & "  if (!el || !el.isConnected) return false; "
    js = js & "  const rect = el.getBoundingClientRect(); "
    js = js & "  if (rect.width === 0 || rect.height === 0) return false; "
    js = js & "  const style = window.getComputedStyle(el); "
    js = js & "  if (style.display === 'none') return false; "
    js = js & "  if (style.visibility === 'hidden') return false; "
    js = js & "  if (style.opacity === '0') return false; "
    js = js & "  return true; "
    js = js & "}"
    GetJsForVisibilityCheck = js
End Function

Private Function GetJsForWaitOptions() As String
    Dim js As String
    js = "function(element, minCount, timeout) { "
    js = js & "  return new Promise((resolve) => { "
    js = js & "    const start = Date.now(); "
    js = js & "    function check() { "
    js = js & "      if (element.options && element.options.length >= minCount) { "
    js = js & "        resolve(true); "
    js = js & "      } else if (Date.now() - start > timeout) { "
    js = js & "        resolve(false); "
    js = js & "      } else { "
    js = js & "        setTimeout(check, 100); "
    js = js & "      } "
    js = js & "    } "
    js = js & "    check(); "
    js = js & "  }); "
    js = js & "}"
    GetJsForWaitOptions = js
End Function

Private Function GetJsForGetOptions() As String
    Dim js As String
    js = "function(element) { "
    js = js & "  var opts = []; "
    js = js & "  for (var i = 0; i < element.options.length; i++) { "
    js = js & "    var o = element.options[i]; "
    js = js & "    opts.push({ 'text': o.innerText, 'value': o.value, 'selected': o.selected }); "
    js = js & "  } "
    js = js & "  return JSON.stringify(opts); "
    js = js & "}"
    GetJsForGetOptions = js
End Function

Private Function GetJsForWaitForIdle() As String
    Dim js As String
    js = "function(timeoutMs, minStableMs) { "
    js = js & "  return new Promise((resolve) => { "
    js = js & "    const start = Date.now(); "
    js = js & "    let lastMutation = Date.now(); "
    js = js & "    const observer = new MutationObserver(() => { lastMutation = Date.now(); }); "
    js = js & "    observer.observe(document, { childList: true, subtree: true, attributes: true }); "
    js = js & "    const check = setInterval(() => { "
    js = js & "      const now = Date.now(); "
    js = js & "      if (now - start > timeoutMs) { cleanup(); resolve('timeout'); return; } "
    js = js & "      const isDomStable = (now - lastMutation) > minStableMs; "
    js = js & "      const isReady = (document.readyState === 'complete'); "
    js = js & "      if (isReady && isDomStable) { cleanup(); resolve('idle'); } "
    js = js & "    }, 100); "
    js = js & "    function cleanup() { if(check) clearInterval(check); observer.disconnect(); } "
    js = js & "  }); "
    js = js & "}"
    GetJsForWaitForIdle = js
End Function

Private Function GetJsForClickAndWait() As String
    Dim js As String
    js = "function(e) { return new Promise((r) => { "
    js = js & "  e.scrollIntoView({block:'center',inline:'center'}); "
    js = js & "  e.focus(); "
    js = js & "  e.click(); "
    js = js & "  setTimeout(() => r('done'), 50); "
    js = js & "}); }"
    GetJsForClickAndWait = js
End Function

Private Function GetJsForInputAndWait() As String
    Dim js As String
    js = "function(e,v) { return new Promise((r) => { "
    js = js & "  e.scrollIntoView({block:'center',inline:'center'}); "
    js = js & "  e.click(); e.focus(); e.value=''; "
    js = js & "  var s=document.execCommand('insertText',false,v); "
    js = js & "  if(!s){ e.value=v; e.dispatchEvent(new Event('input',{bubbles:true})); e.dispatchEvent(new Event('change',{bubbles:true})); } "
    js = js & "  e.blur(); "
    js = js & "  setTimeout(() => r('done'), 50); "
    js = js & "}); }"
    GetJsForInputAndWait = js
End Function

Private Function GetJsForSelectAndWait() As String
    Dim js As String
    js = "function(s,v) { return new Promise((r) => { "
    js = js & "  s.scrollIntoView({block:'center',inline:'center'}); "
    js = js & "  s.focus(); "
    js = js & "  s.value = v; "
    js = js & "  s.dispatchEvent(new Event('input', {bubbles:true})); "
    js = js & "  s.dispatchEvent(new Event('change', {bubbles:true})); "
    js = js & "  s.blur(); "
    js = js & "  setTimeout(() => r('done'), 50); "
    js = js & "}); }"
    GetJsForSelectAndWait = js
End Function

Private Function GetJsForSelectTextAndWait() As String
    Dim js As String
    js = "function(s, textToFind) { return new Promise((r) => { "
    js = js & "  const normalize = (str) => str.replace(/[\s\u00A0]+/g, ' ').trim(); "
    js = js & "  var found = false; "
    js = js & "  var targetText = normalize(textToFind); "
    js = js & "  for (var i = 0; i < s.options.length; i++) { "
    js = js & "    var optText = normalize(s.options[i].text); "
    js = js & "    if (optText === targetText) { s.value = s.options[i].value; found = true; break; } "
    js = js & "  } "
    js = js & "  if (!found) { r('Option text not found: ' + textToFind); return; } "
    js = js & "  s.scrollIntoView({block:'center',inline:'center'}); "
    js = js & "  s.focus(); "
    js = js & "  s.dispatchEvent(new Event('input', {bubbles:true})); "
    js = js & "  s.dispatchEvent(new Event('change', {bubbles:true})); "
    js = js & "  s.blur(); "
    js = js & "  setTimeout(() => r('done'), 50); "
    js = js & "}); }"
    GetJsForSelectTextAndWait = js
End Function

Private Function GetJsForShadowClick() As String
    Dim js As String
    js = "function(selectors, searchTimeout) { "
    js = js & GetJsCommonFindShadow()
    js = js & "  return new Promise(async (resolve) => { "
    js = js & "    const el = await findShadow(selectors, searchTimeout); "
    js = js & "    if (!el) { resolve('Error: Element not found'); return; } "
    js = js & "    el.scrollIntoView({block:'center',inline:'center'}); "
    js = js & "    el.focus(); "
    js = js & "    el.click(); "
    js = js & "    const ev = new MouseEvent('click', { bubbles: true, cancelable: true, view: window, composed: true }); "
    js = js & "    el.dispatchEvent(ev); "
    js = js & "    setTimeout(() => resolve('done'), 50); "
    js = js & "  }); "
    js = js & "}"
    GetJsForShadowClick = js
End Function

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

Private Function GetJsCommonFindShadow() As String
    Dim js As String
    js = "  function findShadow(arr, timeoutMs) { "
    js = js & "    return new Promise((res) => { "
    js = js & "      const limit = timeoutMs || 5000; "
    js = js & "      const end = Date.now() + limit; "
    js = js & "      function check() { "
    js = js & "        let el = document.querySelector(arr[0]); "
    js = js & "        let fail = false; "
    js = js & "        if(el) { "
    js = js & "          for(let i=1; i<arr.length; i++){ "
    js = js & "            if(el.shadowRoot) { el = el.shadowRoot.querySelector(arr[i]); } "
    js = js & "            else { fail = true; break; } "
    js = js & "            if(!el) { fail = true; break; } "
    js = js & "          } "
    js = js & "        } else { fail = true; } "
    js = js & "        if(!fail && el) res(el); "
    js = js & "        else if(Date.now() < end) setTimeout(check, 100); "
    js = js & "        else res(null); "
    js = js & "      } "
    js = js & "      check(); "
    js = js & "    }); "
    js = js & "  } "
    GetJsCommonFindShadow = js
End Function

' ========================================================================================
' Discovery & Recorder API
' ========================================================================================
Public Sub StartDiscoveryLog(Optional ByVal excludeImagesAndCss As Boolean = True)
    Set p_recorderLogs = New Collection
    p_isRecording = True
    p_recorderStartTime = GetTimeMs() / 1000
    p_excludeResources = excludeImagesAndCss

    Dim events(3) As String
    events(0) = "network.beforeRequestSent"
    events(1) = "network.responseCompleted"
    events(2) = "browsingContext.load"
    events(3) = "browsingContext.domContentLoaded"
    Me.ExecuteSessionSubscribe events

    Debug.Print "BiDi Recorder: Started. (Filter Resources: " & excludeImagesAndCss & ")"
End Sub

Public Sub StopAndSaveDiscoveryLog(ByVal filePath As String)
    p_isRecording = False

    If p_recorderLogs Is Nothing Then Exit Sub
    If p_recorderLogs.Count = 0 Then
        Debug.Print "BiDi Recorder: No events captured."
        Exit Sub
    End If

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
    Debug.Print "BiDi Recorder: Log saved to " & filePath & " (" & p_recorderLogs.Count & " events)"
    Set p_recorderLogs = Nothing
End Sub

Public Sub RecordEventsForSeconds(ByVal seconds As Long)
    Dim endTick As Double
    endTick = GetTimeMs() + (seconds * 1000)

    Do While GetTimeMs() < endTick
        ProcessRecorderQueue
        sleep 100
        DoEvents
    Loop
End Sub

' ========================================================================================
' ProcessRecorderQueue: Processes the incoming message queue from the WebSocket.
' [REFACTORED] Switched from time-limit to count-limit to prevent stale event accumulation.
' This ensures that during bursts (e.g., navigation or SPA hydration), all events are
' processed to keep VBA state (p_pendingRequestCount) in sync with the browser.
' ========================================================================================
Public Sub ProcessRecorderQueue()
    Dim msg As String
    Dim json As Object
    Dim processedCount As Long
    
    ' Define a threshold to clear bursts without causing a permanent hang.
    ' 1000 messages is generally sufficient to drain a navigation burst.
    Const MAX_MESSAGES_PER_BATCH As Long = 1000
    
    Do While processedCount < MAX_MESSAGES_PER_BATCH
        ' 1. Defensive Cleanup: Reset object reference for the next iteration
        Set json = Nothing
        
        ' 2. Fetch the next message from the WebSocket communicator
        msg = socket_.GetMessage()
        If Len(msg) = 0 Then Exit Do ' Success: Queue is empty
        
        ' 3. Safety Check: Filter out invalid or oversized messages
        If Len(msg) < MIN_MSG_SIZE Or Len(msg) > MAX_MSG_SIZE Then GoTo ContinueLoop
        
        ' 4. Protocol Filtering: Only attempt to parse BiDi/CDP structured strings
        If InStr(msg, """method"":") > 0 Or InStr(msg, """id"":") > 0 Then
            On Error Resume Next
            Set json = WebJsonConverter.ParseJson(msg)
            
            ' 5. Event Dispatching
            If Err.Number = 0 And Not json Is Nothing Then
                On Error GoTo 0
                DispatchGlobalEvents json
            End If
            On Error GoTo 0
        End If

ContinueLoop:
        ' 6. Memory Management: Reclaim object memory immediately
        If Not json Is Nothing Then Set json = Nothing
        
        processedCount = processedCount + 1
        
        ' 7. OS Interruption: Allow Windows to process events every 100 messages
        ' to prevent the Excel UI from appearing frozen during massive bursts.
        If (processedCount Mod 100) = 0 Then ThrottledDoEvents
    Loop

    ' Log a warning if the batch limit was reached, indicating high protocol pressure.
    If processedCount >= MAX_MESSAGES_PER_BATCH Then
        Debug.Print "BiDi Warning: Message burst reached limit (" & MAX_MESSAGES_PER_BATCH & ")."
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

    relTime = Format((GetTimeMs() / 1000) - p_recorderStartTime, "0.00") & "s"

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
' GetProbeInjectionExpression: Generates the JS payload with integrated noise filtering.
' This version merges library defaults with user-defined filters to optimize stability.
' ========================================================================================
' ========================================================================================
' GetProbeInjectionExpression: Generates the COMPLETE JS payload with noise filtering.
' Features:
'   - Integrated Noise Filtering (Mutation, Fetch, XHR)
'   - WebSocket Tracking with Zero Guard (Prevents negative counts)
'   - Double-Wrap Protection (Prevents infinite recursion if injected multiple times)
' ========================================================================================
Private Function GetProbeInjectionExpression() As String
    ' Merge library defaults and user-defined filters into JS array strings
    Dim selArray As String: selArray = MergeToJsArray(DEFAULT_IGNORE_SELECTORS, p_userIgnoreSelectors)
    Dim netArray As String: netArray = MergeToJsArray(DEFAULT_IGNORE_NETWORK, p_userIgnoreNetworkPatterns)
    
    Dim js As String
    js = "(function(){"
    
    ' 1. Skip if already installed with the correct version
    js = js & "  if(window.__vbaSpaProbeInstalled && window.__vbaSpaProbeVersion === '" & PROBE_VERSION & "') return true;"
    js = js & "  window.__vbaSpaProbeInstalled = true;"
    js = js & "  window.__vbaSpaProbeVersion = '" & PROBE_VERSION & "';"
    
    ' 2. Configuration for Noise Filtering
    js = js & "  var ignoreSelectors = " & selArray & ";"
    js = js & "  var ignoreNetwork = " & netArray & ";"

    ' 3. Backup native APIs
    js = js & "  window.__vbaOriginals = window.__vbaOriginals || {"
    js = js & "    fetch: window.fetch,"
    js = js & "    xhrOpen: XMLHttpRequest.prototype.open,"
    js = js & "    xhrSend: XMLHttpRequest.prototype.send,"
    js = js & "    ws: window.WebSocket"
    js = js & "  };"
    
    ' 4. Initialize or Reuse shared probe state
    js = js & "  var p = window.__vbaIdleProbe || {"
    js = js & "    inflightFetchCount: 0,"
    js = js & "    inflightXhrCount: 0,"
    js = js & "    wsOpenCount: 0,"
    js = js & "    lastMutationTs: performance.now()"
    js = js & "  };"
    js = js & "  window.__vbaIdleProbe = p;"
    
    ' Helper to update stability timestamp
    js = js & "  var updateTs = function(){ p.lastMutationTs = performance.now(); };"

    ' 5. Monitor DOM with Mutation Filtering
    ' Only updates lastMutationTs if the change occurs outside of 'ignoreSelectors'
    js = js & "  try {"
    js = js & "    new MutationObserver(function(mutations){"
    js = js & "      var hasRealChange = mutations.some(function(m){"
    js = js & "        return !ignoreSelectors.some(function(s){ return m.target.closest && m.target.closest(s); });"
    js = js & "      });"
    js = js & "      if(hasRealChange) updateTs();"
    js = js & "    }).observe(document, {childList:true, subtree:true, attributes:true});"
    js = js & "  } catch(e) {}"

    ' 6. Wrap Fetch with URL Filtering
    js = js & "  if(typeof window.fetch === 'function' && !window.fetch.__vbaWrapped){"
    js = js & "    var origFetch = window.fetch;"
    js = js & "    window.fetch = function(u, o){"
    js = js & "      var url = (typeof u === 'string') ? u : (u.url || '');"
    js = js & "      var skip = ignoreNetwork.some(function(pat){ return url.indexOf(pat) !== -1; });"
    js = js & "      if(skip) return origFetch.apply(this, arguments);"
    js = js & "      p.inflightFetchCount++;"
    js = js & "      return origFetch.apply(this, arguments).finally(function(){"
    js = js & "        p.inflightFetchCount--; updateTs();"
    js = js & "      });"
    js = js & "    };"
    js = js & "    window.fetch.__vbaWrapped = true;"
    js = js & "  }"

    ' 7. Wrap XHR with URL Filtering
    js = js & "  if(!XMLHttpRequest.prototype.__vbaWrapped){"
    js = js & "    XMLHttpRequest.prototype.open = function(m, u){"
    js = js & "      this.__vbaUrl = u; this.__vbaHooked = true;"
    js = js & "      return window.__vbaOriginals.xhrOpen.apply(this, arguments);"
    js = js & "    };"
    js = js & "    XMLHttpRequest.prototype.send = function(){"
    js = js & "      var self = this; var url = this.__vbaUrl || '';"
    js = js & "      var skip = ignoreNetwork.some(function(pat){ return url.indexOf(pat) !== -1; });"
    js = js & "      if(skip){ "
    js = js & "        self.__vbaHooked = false; "
    js = js & "        return window.__vbaOriginals.xhrSend.apply(this, arguments); "
    js = js & "      }"
    js = js & "      p.inflightXhrCount++;"
    js = js & "      var done = function(){ "
    js = js & "        if(self.__vbaHooked){ p.inflightXhrCount--; self.__vbaHooked=false; } "
    js = js & "        updateTs(); self.removeEventListener('loadend', done); "
    js = js & "      };"
    js = js & "      this.addEventListener('loadend', done);"
    js = js & "      return window.__vbaOriginals.xhrSend.apply(this, arguments);"
    js = js & "    };"
    js = js & "    XMLHttpRequest.prototype.__vbaWrapped = true;"
    js = js & "  }"
    
    ' 8. Wrap WebSocket with Heartbeat Tracking & Zero Guard
    js = js & "  if(typeof window.WebSocket === 'function' && !window.WebSocket.__vbaWrapped){"
    js = js & "    window.WebSocket = function(u, g){"
    js = js & "      var w = new window.__vbaOriginals.ws(u, g);"
    js = js & "      p.wsOpenCount++;"
    js = js & "      w.addEventListener('message', updateTs);"
    js = js & "      var origSend = w.send;"
    js = js & "      w.send = function(){ updateTs(); return origSend.apply(this, arguments); };"
    js = js & "      var dec = function(){ "
    js = js & "        if (p.wsOpenCount > 0) p.wsOpenCount--;"
    js = js & "        updateTs();"
    js = js & "        w.removeEventListener('close', dec); w.removeEventListener('error', dec);"
    js = js & "      };"
    js = js & "      w.addEventListener('close', dec); w.addEventListener('error', dec);"
    js = js & "      return w;"
    js = js & "    };"
    js = js & "    window.WebSocket.__vbaWrapped = true;"
    js = js & "  }"
    
    js = js & "  return true;"
    js = js & "})();"
    
    GetProbeInjectionExpression = js
End Function

Private Function GetProbeRestoreExpression() As String
    Dim js As String
    js = ""
    js = js & "(function(){"
    js = js & "  var o = window.__vbaOriginals;"
    js = js & "  if(!o) return;"
    
    ' Only restore if WE were the ones who wrapped it
    js = js & "  if(window.fetch && window.fetch.__vbaWrapped) window.fetch = o.fetch;"
    js = js & "  if(XMLHttpRequest.prototype.__vbaWrapped){"
    js = js & "    XMLHttpRequest.prototype.open = o.xhrOpen;"
    js = js & "    XMLHttpRequest.prototype.send = o.xhrSend;"
    js = js & "    delete XMLHttpRequest.prototype.__vbaWrapped;"
    js = js & "  }"
    js = js & "  if(window.WebSocket && window.WebSocket.__vbaWrapped) window.WebSocket = o.ws;"
    
    js = js & "  delete window.__vbaOriginals;"
    js = js & "  delete window.__vbaSpaProbeInstalled;"
    js = js & "  delete window.__vbaSpaProbeVersion;"
    js = js & "  delete window.__vbaIdleProbe;"
    js = js & "})();"
    GetProbeRestoreExpression = js
End Function

Private Function GetProbeQueryExpression() As String
    Dim js As String
    js = ""
    js = js & "function(){"
    js = js & "  var p = window.__vbaIdleProbe;"
    js = js & "  var v = window.__vbaSpaProbeVersion || '0.0';"
    ' Return 'active: false' if the tracking object is missing, but always include the version.
    js = js & "  if(!p) return JSON.stringify({active:false, version: v});"
    js = js & "  return JSON.stringify({"
    js = js & "    active: true,"
    js = js & "    version: v,"
    js = js & "    now: performance.now(),"
    js = js & "    lastMutationTs: p.lastMutationTs,"
    js = js & "    inflightFetchCount: p.inflightFetchCount,"
    js = js & "    inflightXhrCount: p.inflightXhrCount"
    js = js & "  });"
    js = js & "}"
    GetProbeQueryExpression = js
End Function

' ========================================================================================
' QuerySpaProbe: Retrieves probe state and validates integrity via versioning.
' [FIXED] Added version verification to detect "ghost" probes caused by realm races.
' ========================================================================================
Public Function QuerySpaProbe(Optional ByVal contextId As String = "") As Dictionary
    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    If realmId = "" Then Set QuerySpaProbe = Nothing: Exit Function

    Dim p As New Dictionary, t As New Dictionary
    t.Add "realm", realmId
    p.Add "functionDeclaration", GetProbeQueryExpression()
    p.Add "arguments", New Collection
    p.Add "target", t
    p.Add "awaitPromise", False

    Dim ret As String: ret = ExecuteBiDiCommand("script.callFunction", p)
    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(ret)
    Dim s As String: s = json("result")("result")("value")
    Dim dic As Dictionary: Set dic = WebJsonConverter.ParseJson(s)
    On Error GoTo 0
    
    ' --- Integrity Check ---
    If Not dic Is Nothing Then
        ' Verify the probe is active AND the version matches our expected constant
        If Not dic("active") Or (dic("version") <> PROBE_VERSION) Then
            Set dic = Nothing ' Mark as invalid to trigger re-injection
        End If
    End If
    
    Set QuerySpaProbe = dic
End Function

' ========================================================================================
' InitializeSpaProbe: Ensures the monitoring environment is ready.
' Fix: Added a synchronization loop to confirm the probe is active before proceeding.
' ========================================================================================
Private Sub InitializeSpaProbe(ByVal contextId As String)
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId())

    ' --- Step 1: CDP Persistent Registration ---
    ' Ensures the probe is automatically injected on any future page navigation.
    If Not m_isProbeRegistered Then
        Dim cdpParams As New Dictionary
        cdpParams.Add "source", GetProbeInjectionExpression()

        Dim result As Object
        Set result = Me.ExecuteCDPCommand("Page.addScriptToEvaluateOnNewDocument", cdpParams)

        If result("Ok") Then
            m_isProbeRegistered = True
            If DEBUG_MODE Then Debug.Print "BiDi Info: SPA Probe registered via CDP."
        Else
            Debug.Print "BiDi Warning: CDP Probe registration failed: " & result("Error")
        End If
    End If

    ' --- Step 2: Immediate Injection & Verification ---
    ' Force-inject the probe into the current realm and wait until it's verified as active.
    ' This is critical for capturing network requests that start immediately after a UI action.
    StartSpaIdleProbe targetCtx
    
    Dim retry As Integer
    Dim probeStatus As Dictionary
    For retry = 1 To 10 ' Max 500ms wait for JS execution to reflect
        Set probeStatus = QuerySpaProbe(targetCtx)
        If Not probeStatus Is Nothing Then
            If probeStatus("active") Then Exit For
        End If
        sleep 50: DoEvents
    Next retry

    ' --- Step 3: Event Subscription ---
    ' Subscribe to BiDi network events if not already done.
    Dim events As Variant
    events = Array("network.beforeRequestSent", "network.responseCompleted", _
                   "network.fetchError", "browsingContext.load", "browsingContext.fragmentNavigated")
    
    Dim bidiParams As New Dictionary
    bidiParams.Add "events", events
    
    Dim contexts As New Collection: contexts.Add targetCtx
    bidiParams.Add "contexts", contexts

    On Error Resume Next
    ExecuteBiDiCommand "session.subscribe", bidiParams
    On Error GoTo 0
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
    
    Dim startTick As Double: startTick = GetTimeMs()
    Dim stableStart As Double: stableStart = GetTimeMs()
    Dim nowTick As Double

    Do
        ' 4. Continuous Drain
        ' Since MAX_MESSAGES_PER_BATCH is 1000, we stay synced even during heavy hydration.
        ProcessRecorderQueue
        
        ThrottledDoEvents
        sleep 30 ' High-frequency polling
        
        nowTick = GetTimeMs()
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
' [REFACTORED] Returns a Dictionary to provide structural error handling.
' Dictionary Keys:
'   - "Ok"    : Boolean (True if command succeeded at both protocol and execution level)
'   - "Json"  : String (The raw result JSON if Ok is True)
'   - "Error" : String (Detailed error message if Ok is False)
' ========================================================================================
Public Function ExecuteCDPCommand(ByVal cdpMethod As String, _
                                  ByVal cdpParams As Dictionary, _
                                  Optional ByVal useSession As Boolean = True) As Object
    
    ' Initialize the result package
    Dim res As New Dictionary
    res.Add "Ok", False
    res.Add "Json", ""
    res.Add "Error", ""

    ' 1. Connectivity Check
    If socket_ Is Nothing Then
        res("Error") = "Socket object is not set."
        Set ExecuteCDPCommand = res
        Exit Function
    End If

    ' 2. Health Check: Prevent execution if the session is in a fatal state
    If m_consecutiveCdpFailures >= MAX_CONSECUTIVE_CDP_FAILURES Then
        PerformFatalReset
        res("Error") = "Fatal: Multiple CDP session failures. Class reset triggered."
        Set ExecuteCDPCommand = res
        Exit Function
    End If

    ' 3. Parameter Preparation
    Dim params As New Dictionary
    params.Add "method", cdpMethod
    params.Add "params", IIf(cdpParams Is Nothing, New Dictionary, cdpParams)

    ' 4. Session Attachment
    If useSession Then
        Dim sessionId As String: sessionId = GetCdpSessionId()
        If sessionId <> "" Then
            params.Add "session", sessionId
        Else
            ' [CRITICAL FIX] Structural failure reporting for missing sessions
            res("Error") = "CDP session unavailable for context."
            Set ExecuteCDPCommand = res
            Exit Function
        End If
    End If

    ' 5. Execution via BiDi Tunnel
    Dim strRes As String
    On Error Resume Next
    strRes = ExecuteBiDiCommand("goog:cdp.sendCommand", params)
    
    If Err.Number <> 0 Then
        ' Protocol/Socket Level Error
        res("Error") = "BiDi Protocol Error: " & Err.Description
        Err.Clear
    ElseIf InStr(strRes, """error""") > 0 Then
        ' Browser/CDP Level Error (Command rejected by the renderer)
        res("Error") = "CDP Execution Error: " & strRes
    Else
        ' SUCCESS: Command executed and result received
        res("Ok") = True
        res("Json") = strRes
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
' StartSpaIdleProbe: Forces injection of the monitoring script into the current realm.
' ========================================================================================
Public Sub StartSpaIdleProbe(Optional ByVal contextId As String = "")
    Dim targetCtx As String: targetCtx = IIf(contextId <> "", contextId, UpdateMainContextId(True))
    Dim realmId As String: realmId = GetScriptRealmId(targetCtx)
    
    Dim p As New Dictionary, t As New Dictionary

    ' Prioritize Realm for execution stability
    If realmId <> "" Then
        t.Add "realm", realmId
    Else
        t.Add "context", targetCtx
    End If

    p.Add "expression", GetProbeInjectionExpression()
    p.Add "target", t
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
    Dim nowMs As Double
    nowMs = GetTimeMs()
    
    ' Only allow DoEvents every 100ms to maintain stability and UI responsiveness
    If (nowMs - m_lastDoEventsTime) > 100 Then
        DoEvents
        m_lastDoEventsTime = nowMs
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
