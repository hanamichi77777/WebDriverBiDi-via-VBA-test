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

' [Configuration Constants]
Private Const DEFAULT_NAV_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_IDLE_STABLE_MS As Long = 500
Private Const DEFAULT_IDLE_TIMEOUT_MS As Long = 10000
Private Const DEFAULT_STALE_RETRIES As Long = 2
Private Const FRAME_SEARCH_TIMEOUT_MS As Long = 3000
Private Const SHORT_WAIT_MS As Long = 500
Private Const MAX_LOG_SIZE As Long = 500

' File System Object for Download Monitoring
Private fso_ As New Scripting.FileSystemObject
Private p_currentDownloadPath As String

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
' Initialization & Properties
' ========================================================================================
Private Sub Class_Initialize()
    Dim socket As New BiDiSocketCommunicator
    Set socket_ = socket
    
    p_nextRequestId = 1
    ' Initialize high-resolution timer frequency
    QueryPerformanceFrequency p_timerFreq
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
' High-Resolution Timer Helper (Prevents GetTickCount 24.9-day overflow)
' ========================================================================================
Private Function GetTimeMs() As Double
    Dim curTime As Currency
    If p_timerFreq = 0 Then QueryPerformanceFrequency p_timerFreq
    QueryPerformanceCounter curTime
    GetTimeMs = (curTime / p_timerFreq) * 1000
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
Public Function GetIframeContextIdByUrl(ByVal partialUrl As String, Optional ByVal timeOutms As Long = 0) As String
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

        If timeOutms <= 0 Then Exit Do

        If GetTimeMs() - startTick > timeOutms Then
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
                Err.Raise Err.Number, Err.Source, errDesc
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
' Window & Tab Management
' ========================================================================================
Public Function ExecuteFindWindowContextId(ByVal matchType As String, ByVal matchValue As String, _
                                           Optional ByVal timeOutms As Long = 5000, _
                                           Optional ByVal waitForComplete As Boolean = True) As String
    Dim startTick As Double: startTick = GetTimeMs()
    Dim contexts As Collection, ctx As Dictionary, json As Object, strRes As String

    matchType = LCase(matchType)

    Do
        strRes = ExecuteBiDiCommand("browsingContext.getTree", New Dictionary)
        Set json = WebJsonConverter.ParseJson(strRes)

        If Not json Is Nothing Then
            Set contexts = json("result")("contexts")
            For Each ctx In contexts
                Dim isMatch As Boolean
                If matchType = "url" Then
                    isMatch = (InStr(ctx("url"), matchValue) > 0)
                Else
                    isMatch = (InStr(ExecuteGetTitleByContextId(ctx("context")), matchValue) > 0)
                End If

                If isMatch Then
                    If Not waitForComplete Then
                        ExecuteFindWindowContextId = ctx("context"): Exit Function
                    Else
                        If InStr(ExecuteBiDiCommand("script.callFunction", BuildReadyStateParams(ctx("context"))), """value"":true") > 0 Then
                            ExecuteFindWindowContextId = ctx("context"): Exit Function
                        End If
                    End If
                End If
            Next ctx
        End If

        If GetTimeMs() - startTick > timeOutms Then Exit Do
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
' Navigation & Waiting Logic
' ========================================================================================
Public Function ExecuteBiDiWaitUntilUrlContains(ByVal partialUrl As String, Optional ByVal timeoutSec As Long = 300) As Boolean
    Dim startTick As Double: startTick = GetTimeMs()
    Dim lastCheckTick As Double: lastCheckTick = GetTimeMs()
    Dim timeOutms As Long: timeOutms = timeoutSec * 1000
    Dim receivedMsg As String, json As Object

    Dim events(2) As String
    events(0) = "network.responseCompleted"
    events(1) = "browsingContext.load"
    events(2) = "browsingContext.fragmentNavigated"
    Me.ExecuteSessionSubscribe events

    Debug.Print "BiDi: Waiting for URL containing '" & partialUrl & "'..."

    Do
        receivedMsg = socket_.GetMessage()
        If Len(receivedMsg) > 0 Then
            Set json = WebJsonConverter.ParseJson(receivedMsg)
            DispatchGlobalEvents json

            If json.Exists("method") Then
                Select Case json("method")
                    Case "browsingContext.load", "browsingContext.fragmentNavigated"
                        If json("params")("url") Like "*" & partialUrl & "*" Then
                            ExecuteBiDiWaitUntilUrlContains = True: Exit Function
                        End If
                    Case "network.responseCompleted"
                        If json("params").Exists("response") Then
                            Dim res As Object: Set res = json("params")("response")
                            If InStr(res("url"), partialUrl) > 0 Then
                                If InStr(LCase(res("mimeType")), "text/html") > 0 Then
                                    ExecuteBiDiWaitUntilUrlContains = True: Exit Function
                                End If
                            End If
                        End If
                End Select
            End If
        Else
            sleep 50
        End If

        If GetTimeMs() - lastCheckTick > 500 Then
            Dim currentUrl As String
            currentUrl = GetCurrentUrl()
            If InStr(currentUrl, partialUrl) > 0 Then
                ExecuteBiDiWaitUntilUrlContains = True: Exit Function
            End If
            lastCheckTick = GetTimeMs()
        End If

        If GetTimeMs() - startTick > timeOutms Then
            Debug.Print "BiDi: Wait Timed out."
            ExecuteBiDiWaitUntilUrlContains = False: Exit Function
        End If
        DoEvents
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

    ' Subscribe to standard events plus domContentLoaded for faster "Interactive" detection
    Me.ExecuteSessionSubscribe Array( _
        "network.beforeRequestSent", _
        "network.responseCompleted", _
        "browsingContext.load", _
        "browsingContext.domContentLoaded", _
        "browsingContext.fragmentNavigated" _
    )

    Dim targetCtx As String
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId(True)
    End If

    Dim params As New Dictionary
    params.Add "context", targetCtx
    params.Add "url", targetUrl
    ' "none" returns control immediately, allowing us to handle the wait logic manually in the loop below.
    params.Add "wait", "none"
    ExecuteBiDiCommand "browsingContext.navigate", params

    Dim startTick As Double: startTick = GetTimeMs()
    Dim foundStatus As Long: foundStatus = 0
    Dim isLoadComplete As Boolean: isLoadComplete = False
    Dim loadEventReceived As Boolean: loadEventReceived = False
    Dim lastReadyStateCheck As Double: lastReadyStateCheck = 0
    Dim firstActivityTick As Double: firstActivityTick = 0

    Do While GetTimeMs() - startTick < DEFAULT_NAV_TIMEOUT_MS
        Dim msg As String: msg = socket_.GetMessage()

        If Len(msg) > 0 Then
            Dim json As Object: Set json = WebJsonConverter.ParseJson(msg)
            DispatchGlobalEvents json

            If json.Exists("method") Then
                Select Case json("method")
                    Case "network.responseCompleted"
                        If firstActivityTick = 0 Then firstActivityTick = GetTimeMs()
                        If json("params").Exists("response") Then
                            Dim res As Object: Set res = json("params")("response")
                            ' Capture status code if the URL matches the target (or if targetUrl is empty)
                            If InStr(res("url"), targetUrl) > 0 Or targetUrl = "" Then
                                foundStatus = res("status")
                            End If
                        End If

                    Case "browsingContext.load"
                        isLoadComplete = True
                        loadEventReceived = True
                    
                    ' [NEW] If waitForInteractive is True, treat domContentLoaded as completion
                    Case "browsingContext.domContentLoaded"
                        If waitForInteractive Then
                            isLoadComplete = True
                            loadEventReceived = True
                        End If

                    Case "browsingContext.fragmentNavigated"
                        isLoadComplete = True
                End Select
            End If
        Else
            sleep 50
        End If

        ' Poll readyState if no completion event has been received yet
        ' Reduced check interval (1000ms -> 500ms) for better responsiveness
        If Not isLoadComplete And (GetTimeMs() - lastReadyStateCheck > 500) Then
            Dim readyState As String
            readyState = ExecuteGetReadyState(targetCtx)
            
            ' [MODIFIED] Logic: Accept "interactive" state if configured to do so
            If readyState = "complete" Or (waitForInteractive And readyState = "interactive") Then
                isLoadComplete = True
                If foundStatus = 0 Then foundStatus = 200 ' Fallback if status code was missed
            End If
            lastReadyStateCheck = GetTimeMs()
        End If

        ' Break loop if we have the status code and the page load condition is met
        If foundStatus > 0 And isLoadComplete Then Exit Do
        DoEvents
    Loop
    
    ' [OPTIMIZATION] If the caller does NOT want to wait for network idle, return immediately.
    ' This is the key to speed when you only need the status code or basic DOM access.
    If Not waitNetworkIdle Then
        ExecuteNavigateAndGetStatus = foundStatus
        Exit Function
    End If

    ' --- Wait for Network Idle Logic (Only executed if waitNetworkIdle = True) ---
    
    Dim useSpaIdle As Boolean: useSpaIdle = False

    If Not loadEventReceived Then
        Dim nowTick As Double: nowTick = GetTimeMs()
        If (firstActivityTick > 0 And (nowTick - firstActivityTick) >= SPA_LOAD_GRACE_MS) Then
            useSpaIdle = True
        ElseIf isLoadComplete And foundStatus > 0 Then
            useSpaIdle = True
        End If
    End If

    WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx, domDeltaThreshold:=20, allowWsAlwaysOn:=True

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
' Resource Blocking via CDP Tunneling (Fetch Domain)
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
    ExecuteEnableResourceBlocking = ExecuteCDPCommand("Network.setBlockedURLs", params)
End Function

Public Function ExecuteDisableResourceBlocking() As String
    Dim params As New Dictionary
    Dim emptyList As New Collection
    params.Add "urls", emptyList
    ExecuteDisableResourceBlocking = ExecuteCDPCommand("Network.setBlockedURLs", params)
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
Public Function ExecuteWaitForSelectOptions(ByVal xpath As String, Optional ByVal minCount As Long = 1, Optional ByVal timeOutms As Long = 5000, Optional ByVal contextId As String = "") As Boolean
    Dim sharedId As String: sharedId = GetSharedIdFromXPath(xpath, contextId)
    If sharedId = "" Then Exit Function

    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    Dim params As New Dictionary, target As New Dictionary, args As New Collection
    Dim arg1 As New Dictionary, arg2 As New Dictionary, arg3 As New Dictionary

    arg1.Add "sharedId", sharedId
    arg2.Add "type", "number": arg2.Add "value", minCount
    arg3.Add "type", "number": arg3.Add "value", timeOutms
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
    Dim params As New Dictionary
    If IsArray(events) Then
        params.Add "events", events
    Else
        Dim arr(0) As String: arr(0) = events
        params.Add "events", arr
    End If
    ExecuteSessionSubscribe = ExecuteBiDiCommand("session.subscribe", params)
End Function

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
    strRes = ExecuteBiDiCommand("browsingContext.navigate", params)

    If InStr(strRes, """type"":""success""") > 0 Then
        WaitForNetworkIdleSpa DEFAULT_IDLE_STABLE_MS, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
        ExecuteBrowsingContextNavigate = strRes
    Else
        ExecuteBrowsingContextNavigate = strRes
    End If
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

Public Function ExecuteClickAndWaitByXPath(ByVal xpath As String, _
                                           Optional ByVal waitNetworkIdle As Boolean = True, _
                                           Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                           Optional ByVal contextId As String = "", _
                                           Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                           Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS) As String
    ' --- NEW: Start monitoring BEFORE the action to eliminate the blind spot ---
    If waitNetworkIdle Then InitializeSpaProbe contextId

    Dim res As String
    res = ExecuteClickByXPath(xpath, contextId, maxRetries)

    ' Validate the script result using the helper
    If waitNetworkIdle And IsScriptDone(res) Then
        Dim targetCtx As String
        If contextId <> "" Then targetCtx = contextId Else targetCtx = p_mainContextId
        ' Proceed to the idle loop (InitializeSpaProbe inside is idempotent)
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx, 30, True
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

    For attempt = 0 To maxRetries
        On Error Resume Next
        clickRes = Me.ExecuteClickByXPath(xpath, contextId, 0)
        If InStr(clickRes, "done") > 0 Or InStr(clickRes, "success") > 0 Then
            success = True
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0
        If attempt < maxRetries Then sleep 100
    Next attempt

    If Not success Then
        Err.Raise vbObjectError + 500, "ExecuteClickAndNavigateByXPath", "Failed to click element after " & maxRetries & " retries. XPath: " & xpath
    End If

    Dim newCtxId As String
    newCtxId = UpdateMainContextId(True)

    If waitNetworkIdle Then
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, newCtxId, 30, True
    End If

    ExecuteClickAndNavigateByXPath = Me.GetCurrentUrl(newCtxId)
End Function

Public Function ExecuteInputValueAndWaitByXPath(ByVal xpath As String, _
                                                ByVal valueToSet As String, _
                                                Optional ByVal contextId As String = "", _
                                                Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES, _
                                                Optional ByVal waitNetworkIdle As Boolean = True, _
                                                Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS) As String
    ' --- NEW: Start monitoring BEFORE the action to capture immediate network/DOM changes ---
    If waitNetworkIdle Then InitializeSpaProbe contextId

    Dim res As String
    res = ExecuteActionWithStaleRetry(xpath, GetJsForInputAndWait(), valueToSet, True, contextId, maxRetries)

    If waitNetworkIdle And IsScriptDone(res) Then
        Dim targetCtx As String
        If contextId <> "" Then targetCtx = contextId Else targetCtx = p_mainContextId
        WaitForNetworkIdleSpa minStableMs, DEFAULT_IDLE_TIMEOUT_MS, targetCtx
    End If

    ExecuteInputValueAndWaitByXPath = res
End Function

Public Function ExecuteSelectValueAndWaitByXPath(ByVal xpath As String, _
                                                 ByVal valueOrText As String, _
                                                 Optional ByVal selectByText As Boolean = False, _
                                                 Optional ByVal waitNetworkIdle As Boolean = True, _
                                                 Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                                 Optional ByVal contextId As String = "", _
                                                 Optional ByVal maxRetries As Long = DEFAULT_STALE_RETRIES) As String
    ' --- NEW: Proactively subscribe to events BEFORE selection occurs ---
    If waitNetworkIdle Then InitializeSpaProbe contextId

    Dim jsFunction As String
    If selectByText Then jsFunction = GetJsForSelectTextAndWait() Else jsFunction = GetJsForSelectAndWait()

    Dim res As String
    res = ExecuteActionWithStaleRetry(xpath, jsFunction, valueOrText, True, contextId, maxRetries)

    If waitNetworkIdle And IsScriptDone(res) Then
        Dim targetCtx As String
        If contextId <> "" Then targetCtx = contextId Else targetCtx = p_mainContextId
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
' Get Current URL
' ========================================================================================
Public Function GetCurrentUrl(Optional ByVal contextId As String = "") As String
    Dim js As String
    js = "function(){ return window.location.href; }"

    Dim realmId As String
    realmId = GetScriptRealmId(contextId)
    If realmId = "" Then Exit Function

    Dim params As New Dictionary, target As New Dictionary
    target.Add "realm", realmId
    params.Add "functionDeclaration", js
    params.Add "arguments", New Collection
    params.Add "target", target
    params.Add "awaitPromise", False

    Dim strRes As String
    strRes = ExecuteBiDiCommand("script.callFunction", params)

    On Error Resume Next
    Dim json As Object
    Set json = WebJsonConverter.ParseJson(strRes)

    If Not json Is Nothing Then
        If json.Exists("result") Then
            If json("result").Exists("result") Then
                GetCurrentUrl = json("result")("result")("value")
            End If
        End If
    End If
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

Public Function ExecuteSetDownloadPath(ByVal downloadPath As String) As String
    p_currentDownloadPath = downloadPath
    Dim cdpParams As New Dictionary
    cdpParams.Add "behavior", "allow"
    cdpParams.Add "downloadPath", downloadPath
    cdpParams.Add "eventsEnabled", True
    ExecuteSetDownloadPath = ExecuteCDPCommand("Page.setDownloadBehavior", cdpParams, True)
End Function

Public Function ExecuteWaitForDownloadCompletion(ByVal startTime As Date, Optional ByVal timeoutSec As Long = 30) As Boolean
    Dim startTick As Double: startTick = GetTimeMs()
    Dim timeOutms As Long: timeOutms = timeoutSec * 1000
    Dim folder As Scripting.folder
    Dim file As Scripting.file
    Dim hasTempFile As Boolean
    Dim foundNewFile As Boolean

    If p_currentDownloadPath = "" Then
        Debug.Print "BiDi Error: Download path not set. Call ExecuteSetDownloadPath first."
        Exit Function
    End If

    If Not fso_.FolderExists(p_currentDownloadPath) Then
        Debug.Print "BiDi Error: Download folder does not exist."
        Exit Function
    End If

    Debug.Print "BiDi: Monitoring file system for changes after " & startTime & "..."

    Do
        ProcessRecorderQueue

        If GetTimeMs() - startTick > timeOutms Then
            Debug.Print "BiDi: Wait download timed out."
            ExecuteWaitForDownloadCompletion = False
            Exit Function
        End If

        hasTempFile = False
        foundNewFile = False

        Set folder = fso_.GetFolder(p_currentDownloadPath)

        For Each file In folder.Files
            If InStr(LCase(file.Name), ".crdownload") > 0 Or InStr(LCase(file.Name), ".tmp") > 0 Then
                hasTempFile = True
            Else
                If DateDiff("s", startTime, file.DateLastModified) >= 0 Then
                    foundNewFile = True
                End If
            End If
        Next file

        If Not hasTempFile And foundNewFile Then
            Debug.Print "BiDi: Download detected (New or Updated file found)."
            ExecuteWaitForDownloadCompletion = True
            Exit Function
        End If

        sleep SHORT_WAIT_MS
        DoEvents
    Loop
End Function

Public Sub ExecuteClearDownloadFolder()
    If p_currentDownloadPath = "" Then Exit Sub
    If Not fso_.FolderExists(p_currentDownloadPath) Then Exit Sub

    On Error Resume Next
    fso_.DeleteFile fso_.BuildPath(p_currentDownloadPath, "*.*"), True
    On Error GoTo 0
End Sub

Public Sub ExecuteEnableStealthMode()
    Dim js As String
    js = "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
    Dim cdpParams As New Dictionary
    cdpParams.Add "source", js
    ExecuteCDPCommand "Page.addScriptToEvaluateOnNewDocument", cdpParams
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
Public Function ExecuteShadowClickAndNavigate(ByVal selectorsArray As Variant, _
                                              Optional ByVal waitNetworkIdle As Boolean = True, _
                                              Optional ByVal minStableMs As Long = DEFAULT_IDLE_STABLE_MS, _
                                              Optional ByVal maxTimeoutMs As Long = DEFAULT_IDLE_TIMEOUT_MS, _
                                              Optional ByVal contextId As String = "", _
                                              Optional ByVal searchTimeoutMs As Long = 5000) As String
    
    ' 1. Prepare Context and Realm
    Dim targetCtx As String
    If contextId <> "" Then
        targetCtx = contextId
    Else
        targetCtx = UpdateMainContextId(True)
    End If
    
    Dim realmId As String: realmId = GetScriptRealmId(targetCtx)
    
    ' 2. Build Arguments
    Dim params As New Dictionary, target As New Dictionary, args As New Collection
    args.Add BuildShadowSelectorArg(selectorsArray)
    
    Dim argSearch As New Dictionary
    argSearch.Add "type", "number"
    argSearch.Add "value", searchTimeoutMs
    args.Add argSearch
    
    target.Add "realm", realmId
    params.Add "functionDeclaration", GetJsForShadowClick()
    params.Add "arguments", args
    params.Add "target", target
    params.Add "awaitPromise", True ' Wait for the JS click promise to resolve
    
    ' 3. Execute Click
    Dim res As String
    res = ExecuteBiDiCommand("script.callFunction", params)
    
    ' Error Check
    If InStr(res, "Error: Element not found") > 0 Or InStr(res, """error""") > 0 Then
         ' Shadow DOM element not found or click failed
         ExecuteShadowClickAndNavigate = res
         Exit Function
    End If
    
    ' 4. Wait Strategy (Use SPA Logic)
    If waitNetworkIdle Then
        ' Since the click triggers navigation or DOM updates, invoke the SPA-optimized wait logic.
        ' It internally handles "Probe Loss (Navigation) -> Recovery -> Stabilization",
        ' so the manual event loop (from legacy code) is no longer required.
        WaitForNetworkIdleSpa minStableMs, maxTimeoutMs, targetCtx, 30, True
        
        ' Refresh Context ID after wait completes (ID changes upon navigation)
        targetCtx = UpdateMainContextId(True)
    Else
        ' Even if not waiting, pause briefly and update context
        sleep 100
        targetCtx = UpdateMainContextId(True)
    End If
    
    ' 5. Return Current URL
    ExecuteShadowClickAndNavigate = Me.GetCurrentUrl(targetCtx)
End Function

' ========================================================================================
' Core Command Helper: Handles outgoing BiDi commands and context recovery
' ========================================================================================
Public Function ExecuteBiDiCommand(ByVal method As String, ByVal params As Dictionary) As String
    If socket_ Is Nothing Then Err.Raise 91, , "Socket object is not set."

    Dim retryCount As Integer, maxRetries As Integer: maxRetries = 1
    Dim strRes As String, sendMsg As String, dic As New Dictionary

    For retryCount = 0 To maxRetries
        ' 1. Prepare the JSON request packet
        dic.RemoveAll
        dic.Add "id", p_nextRequestId
        dic.Add "method", method
        dic.Add "params", params
        p_nextRequestId = p_nextRequestId + 1

        ' 2. Serialize and Log the request
        sendMsg = WebJsonConverter.ConvertToJson(dic)
        If DEBUG_MODE Then DebugPrintLong "Request (Try " & retryCount & ")", sendMsg

        ' 3. Execute synchronous Send and Receive
        strRes = socket_.SendAndReceive(sendMsg)

        ' [Performance Guard] Log a warning for oversized responses.
        ' Large strings incur high latency during JSON parsing ($O(n^2)$ complexity in some parsers).
        ' We do not skip these as they are explicit requests, but we alert the developer.
        If Len(strRes) > MAX_MSG_SIZE And DEBUG_MODE Then
            Debug.Print "BiDi Warning: Large response received for method '" & method & "' (" & Len(strRes) & " chars)."
        End If

        If DEBUG_MODE Then DebugPrintLong "Response (Try " & retryCount & ")", strRes

        ' 4. Inspect for "Context Destroyed" or "Stale" errors requiring recovery
        Dim isContextDestroyed As Boolean
        isContextDestroyed = (InStr(strRes, "Execution context was destroyed") > 0) Or _
                             (InStr(strRes, "Inspected target navigated or closed") > 0) Or _
                             (InStr(strRes, "Generic failure: No node with given id found") > 0) Or _
                             (InStr(strRes, "no such frame") > 0)

        If isContextDestroyed Then
            If retryCount < maxRetries Then
                ' Execute recovery: Clear cached IDs and re-identify active contexts
                Debug.Print "BiDi Info: Context destroyed. Attempting recovery (Retry " & retryCount + 1 & ")..."
                ClearState

                ' Refresh context/realm IDs in the parameter dictionary
                If Not params.Exists("target") Then
                    UpdateMainContextId True
                Else
                    Dim newRealm As String
                    newRealm = GetScriptRealmId("", True)
                    If params("target").Exists("realm") Then params("target")("realm") = newRealm
                End If

                sleep SHORT_WAIT_MS
                DoEvents
            Else
                ' Max retries reached, escalate the error
                CheckAndRaiseError strRes, method
            End If
        Else
            ' Successful response or generic BiDi error
            CheckAndRaiseError strRes, method
            Exit For
        End If
    Next retryCount

    ExecuteBiDiCommand = strRes
End Function

Private Sub CheckAndRaiseError(ByVal jsonString As String, ByVal methodName As String)
    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonString)
    On Error GoTo 0

    If json Is Nothing Then Exit Sub
    Dim errType As String, errMsg As String

    If json.Exists("error") Then
        errType = json("error")
        If json.Exists("message") Then errMsg = json("message")
        If IsIgnorableError(errType, errMsg) Then Exit Sub
        Err.Raise vbObjectError + 513, "BiDi:" & methodName, "Browser Error [" & errType & "]: " & errMsg
    End If

    If json.Exists("result") Then
        If json("result").Exists("exceptionDetails") Then
            Dim details As Object: Set details = json("result")("exceptionDetails")
            If details.Exists("text") Then errMsg = details("text")
            If details.Exists("exception") Then
                If details("exception").Exists("description") Then
                    errMsg = errMsg & " - " & details("exception")("description")
                End If
            End If
            If IsIgnorableError("", errMsg) Then Exit Sub
            Err.Raise vbObjectError + 514, "BiDi:JS_Exception", "JavaScript Error in " & methodName & ": " & errMsg
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
' State Management
' ========================================================================================
Public Sub ClearState()
    p_mainRealmId = ""
    p_cdpSessionId = ""
    p_mainContextId = ""
    p_pendingRequestCount = 0
End Sub

Public Function UpdateMainContextId(Optional ByVal forceUpdate As Boolean = False) As String
    If Not forceUpdate And p_mainContextId <> "" Then
        UpdateMainContextId = p_mainContextId
        Exit Function
    End If

    Dim params As New Dictionary
    Dim strRes As String
    strRes = ExecuteBiDiCommand("browsingContext.getTree", params)

    Dim resJson As Object
    On Error Resume Next
    Set resJson = WebJsonConverter.ParseJson(strRes)
    On Error GoTo 0

    If resJson Is Nothing Then Exit Function
    If Not resJson.Exists("result") Then Exit Function
    If Not resJson("result").Exists("contexts") Then Exit Function

    Dim ctxList As Collection
    Set ctxList = resJson("result")("contexts")

    Dim ctx As Dictionary
    For Each ctx In ctxList
        If ctx.Exists("context") Then
            p_mainContextId = ctx("context")
            Exit For
        End If
    Next ctx

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
    strRes = ExecuteBiDiCommand("goog:cdp.getSession", params)
    On Error Resume Next
    p_cdpSessionId = WebJsonConverter.ParseJson(strRes)("result")("session")
    On Error GoTo 0
    GetCdpSessionId = p_cdpSessionId
End Function

' Helper to determine if the injected JavaScript returned 'done'.
' This prevents false positives from the BiDi protocol's top-level "success" message.
Private Function IsScriptDone(ByVal jsonResponse As String) As Boolean
    Dim json As Object
    On Error Resume Next
    Set json = WebJsonConverter.ParseJson(jsonResponse)
    
    ' Extract the actual script return value: result.result.value
    If Not json Is Nothing Then
        IsScriptDone = (json("result")("result")("value") = "done")
    End If
    On Error GoTo 0
End Function

' ========================================================================================
' Auto-Clicker Injection (XPath Version)
' ========================================================================================
Public Sub ExecuteRegisterAutoClickerByXPath(ByVal targetXPath As String)
    Dim js As String
    Dim safeXPath As String
    safeXPath = Replace(targetXPath, "'", "\'")

    js = "(function() {"
    js = js & "  var xpath = '" & safeXPath & "';"
    js = js & "  function checkAndClick() {"
    js = js & "    try {"
    js = js & "      var result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);"
    js = js & "      var btn = result.singleNodeValue;"
    js = js & "      if (btn && btn.offsetParent !== null) { btn.click(); return true; }"
    js = js & "    } catch(e) {}"
    js = js & "    return false;"
    js = js & "  }"
    js = js & "  if (!checkAndClick()) {"
    js = js & "    var observer = new MutationObserver(function(mutations) {"
    js = js & "      if (checkAndClick()) observer.disconnect();"
    js = js & "    });"
    js = js & "    observer.observe(document, { childList: true, subtree: true });"
    js = js & "  }"
    js = js & "})();"

    Dim params As New Dictionary
    params.Add "source", js
    Me.ExecuteCDPCommand "Page.addScriptToEvaluateOnNewDocument", params
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

Public Sub ProcessRecorderQueue()
    Dim msg As String, json As Object, processedCount As Long, startTick As Double
    startTick = GetTimeMs()
    
    Do
        msg = socket_.GetMessage()
        If Len(msg) = 0 Then Exit Do

        ' --- High-speed filtering by message size ---
        ' Skip noise (extremely small) and massive data (oversized) to protect VBA main thread
        If Len(msg) < MIN_MSG_SIZE Then GoTo NextMessage
        
        If Len(msg) > MAX_MSG_SIZE Then
            If DEBUG_MODE Then
                Debug.Print "BiDi Info: Skipped oversized event (" & Len(msg) & " chars). Hint: " & Left(msg, 150)
            End If
            GoTo NextMessage
        End If

        ' --- Pre-parsing optimization ---
        ' Only parse if the message contains essential BiDi keys.
        If InStr(msg, """method"":") > 0 Or InStr(msg, """id"":") > 0 Then
            On Error Resume Next
            Set json = WebJsonConverter.ParseJson(msg)
            On Error GoTo 0

            If Not json Is Nothing Then
                DispatchGlobalEvents json
            End If
        End If

NextMessage:
        processedCount = processedCount + 1
        ' Pump UI every 25 messages to keep Excel responsive during bursts
        If processedCount Mod 25 = 0 Then DoEvents

        ' Safety break: Return control to VBA if processing takes more than 200ms
        If GetTimeMs() - startTick > 200 Then Exit Do
    Loop
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
' SPA Idle Probe (Injection & Query) - JS without comments
' ========================================================================================
Private Function GetProbeInjectionExpression() As String
    Dim js As String
    js = ""
    js = js & "(function(){"
    js = js & "if(window.__vbaIdleProbe&&window.__vbaIdleProbe.__wired2)return;"
    js = js & "var p=window.__vbaIdleProbe||{startTs:Date.now(),lastDomMutationTs:Date.now(),nodeCountDelta:0,inflightFetchCount:0,inflightXhrCount:0,wsOpenCount:0};"
    js = js & "try{new MutationObserver(function(m){p.lastDomMutationTs=Date.now();var d=0;for(var i=0;i<m.length;i++){var a=m[i].addedNodes,r=m[i].removedNodes;if(a)d+=a.length;if(r)d+=r.length}p.nodeCountDelta+=d}).observe(document,{childList:true,subtree:true});}catch(e){}"
    js = js & "try{if(typeof window.fetch==='function'&&!p.__origFetch){p.__origFetch=window.fetch.bind(window);Object.defineProperty(window,'fetch',{configurable:true,writable:true,value:function(){p.inflightFetchCount++;var pr;try{pr=p.__origFetch.apply(this,arguments)}catch(e){p.inflightFetchCount--;throw e}return Promise.resolve(pr).finally(function(){if(p.inflightFetchCount>0)p.inflightFetchCount--;});}});}}catch(e){}"
    js = js & "try{var X=window.XMLHttpRequest;if(typeof X==='function'&&!p.__xhrPatched){var origOpen=X.prototype.open;var origSend=X.prototype.send;var dec=function(){if(p.inflightXhrCount>0)p.inflightXhrCount--;};X.prototype.open=function(){this.__vbaIdleHooked=false;return origOpen.apply(this,arguments)};X.prototype.send=function(){try{p.inflightXhrCount++;if(!this.__vbaIdleHooked){try{this.addEventListener('loadend',dec)}catch(e){}try{this.addEventListener('error',dec)}catch(e){}try{this.addEventListener('abort',dec)}catch(e){}this.__vbaIdleHooked=true}}catch(e){}try{return origSend.apply(this,arguments)}catch(e){dec();throw e}};p.__xhrPatched=true;}}catch(e){}"
    js = js & "try{var OW=window.WebSocket;if(typeof OW==='function'&&!p.__origWS){p.__origWS=OW;window.WebSocket=function(u,g){var w=new p.__origWS(u,g);p.wsOpenCount++;var dec=function(){p.wsOpenCount--;try{w.removeEventListener('close',dec)}catch(e){}try{w.removeEventListener('error',dec)}catch(e){}};try{w.addEventListener('close',dec)}catch(e){}try{w.addEventListener('error',dec)}catch(e){}return w};}}catch(e){}"
    js = js & "window.__vbaIdleProbe=p;"
    js = js & "p.__wired2=true;"
    js = js & "return true;"
    js = js & "})();"
    GetProbeInjectionExpression = js
End Function


Private Function GetProbeQueryExpression() As String
    Dim js As String
    js = ""
    js = js & "function(){"
    js = js & "var p=window.__vbaIdleProbe;"
    js = js & "if(!p)return JSON.stringify({active:false});"
    js = js & "return JSON.stringify({active:true,now:Date.now(),lastDomMutationTs:p.lastDomMutationTs,nodeCountDelta:p.nodeCountDelta,inflightFetchCount:p.inflightFetchCount,inflightXhrCount:p.inflightXhrCount,wsOpenCount:p.wsOpenCount});"
    js = js & "}"
    GetProbeQueryExpression = js
End Function

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
    Set QuerySpaProbe = dic
End Function

' ========================================================================================
' SPA Idle (Hybrid Conditions) - REFACTORED & CLEAN
' ========================================================================================
Public Sub WaitForNetworkIdleSpa(ByVal minStableMs As Long, _
                                  ByVal maxTimeoutMs As Long, _
                                  Optional ByVal contextId As String = "", _
                                  Optional ByVal domDeltaThreshold As Long = 50, _
                                  Optional ByVal allowWsAlwaysOn As Boolean = True)

    ' --- 1. Initialization and Probe Injection ---
    InitializeSpaProbe contextId, False

    ' --- 2. Variable Definitions ---
    Dim startTick As Double: startTick = GetTimeMs()
    Dim stableStart As Double: stableStart = GetTimeMs()
    Dim nowTick As Double
    
    Dim baselineDelta As Long
    Dim baselineSet As Boolean: baselineSet = False
    Dim isIdle As Boolean

    ' --- 3. Monitoring Loop ---
    Do
        ProcessRecorderQueue
        DoEvents
        sleep 50
        nowTick = GetTimeMs()

        ' Check for timeout
        If (nowTick - startTick) > maxTimeoutMs Then Exit Do

        ' Query Probe status
        Dim probe As Dictionary
        Set probe = QuerySpaProbe(contextId)
        
        Dim isProbeActive As Boolean: isProbeActive = False
        If Not probe Is Nothing Then
            If probe.Exists("active") Then isProbeActive = probe("active")
        End If

        If isProbeActive Then
            ' --- A. Probe Active: Evaluate idle condition ---
            isIdle = EvaluateSpaIdleCondition(probe, minStableMs, domDeltaThreshold, allowWsAlwaysOn, _
                                              stableStart, baselineDelta, baselineSet, nowTick)
            
            If isIdle Then Exit Do
        Else
            ' --- B. Probe Lost (Navigation Detected): Execute recovery ---
            RecoverSpaProbe contextId
            
            ' Reset state (rewind stability timer)
            stableStart = nowTick
            baselineSet = False
        End If
    Loop
End Sub

' Modified: Added reset parameter to prevent wiping out active counters after an action.
Private Sub InitializeSpaProbe(ByVal contextId As String, Optional ByVal reset As Boolean = True)
    ' Subscribe to necessary events
    Me.ExecuteSessionSubscribe Array( _
        "network.beforeRequestSent", _
        "network.responseCompleted", _
        "network.fetchError", _
        "browsingContext.load", _
        "browsingContext.fragmentNavigated" _
    )

    ' Initial Probe injection (JS guard prevents duplicate patches)
    StartSpaIdleProbe contextId

    ' Short buffer (wait out post-click transient state)
    Dim startTick As Double: startTick = GetTimeMs()
    Do
        ProcessRecorderQueue
        sleep 10
        If p_pendingRequestCount > 0 Then Exit Do
    Loop While (GetTimeMs() - startTick) < 200

    ' --- Only reset counters if explicitly requested ---
    ' This prevents clearing the '1' count triggered by the preceding action.
    If reset Then ResetSpaProbeCounters contextId
End Sub

'Recovery
Private Sub RecoverSpaProbe(ByVal contextId As String)
    ' Force update RealmID (follow new page context)
    GetScriptRealmId contextId, True
    
    ' Re-inject Probe into the new page
    StartSpaIdleProbe contextId
    
    ' Note: Navigation detected. Probe re-injected.
End Sub

'Condition Evaluation
Private Function EvaluateSpaIdleCondition(ByVal probe As Dictionary, _
                                          ByVal minStableMs As Long, _
                                          ByVal domDeltaThreshold As Long, _
                                          ByVal allowWsAlwaysOn As Boolean, _
                                          ByRef stableStart As Double, _
                                          ByRef baselineDelta As Long, _
                                          ByRef baselineSet As Boolean, _
                                          ByVal nowTick As Double) As Boolean
    
    Dim ts As Double, lastTs As Double
    Dim inflFetch As Long, inflXhr As Long, wsOpen As Long, domDelta As Long

    ' Extract values from Probe
    ts = CDbl(probe("now"))
    lastTs = CDbl(probe("lastDomMutationTs"))
    inflFetch = CLng(probe("inflightFetchCount"))
    inflXhr = CLng(probe("inflightXhrCount"))
    wsOpen = CLng(probe("wsOpenCount"))
    domDelta = CLng(probe("nodeCountDelta"))

    ' Set baseline (measure DOM changes from this point forward)
    If Not baselineSet Then
        baselineDelta = domDelta
        baselineSet = True
    End If

    ' Calculate DOM change delta
    Dim deltaGrowth As Long: deltaGrowth = domDelta - baselineDelta
    If deltaGrowth < 0 Then deltaGrowth = 0

    ' --- Judgment Logic (Prioritize Internal Counters) ---
    Dim netIdle As Boolean, xferIdle As Boolean, domStable As Boolean, wsOk As Boolean

    ' 1. Network Idle Check
    xferIdle = (inflFetch = 0 And inflXhr = 0)
    
    If xferIdle Then
        ' If internal (JS) is quiet, ignore external (VBA) counter drift
        netIdle = True
        p_pendingRequestCount = 0 ' Correct external counter for next run
    Else
        ' If internal is busy, check external as well (safety fallback)
        netIdle = (p_pendingRequestCount = 0)
    End If

    ' 2. DOM Stability & WebSocket Check
    domStable = ((ts - lastTs) >= minStableMs And deltaGrowth <= domDeltaThreshold)
    wsOk = (allowWsAlwaysOn Or wsOpen = 0)

    ' --- Final Decision ---
    If netIdle And xferIdle And domStable And wsOk Then
        ' Has the condition been met for minStableMs?
        If (nowTick - stableStart) >= minStableMs Then
            EvaluateSpaIdleCondition = True
        Else
            EvaluateSpaIdleCondition = False
        End If
    Else
        ' If conditions break, reset timer & update baseline
        stableStart = nowTick
        baselineDelta = domDelta ' Set current state as the new baseline
        EvaluateSpaIdleCondition = False
    End If

End Function

' Send a raw CDP command via BiDi's goog:cdp.sendCommand.
' If useSession=True, the current CDP session tied to the main context will be included.
Public Function ExecuteCDPCommand(ByVal cdpMethod As String, _
                                  ByVal cdpParams As Dictionary, _
                                  Optional ByVal useSession As Boolean = True) As String
    If socket_ Is Nothing Then Err.Raise 91, , "Socket object is not set."

    Dim params As New Dictionary
    params.Add "method", cdpMethod

    If cdpParams Is Nothing Then
        params.Add "params", New Dictionary
    Else
        params.Add "params", cdpParams
    End If

    If useSession Then
        Dim sessionId As String
        sessionId = GetCdpSessionId()
        If sessionId <> "" Then
            params.Add "session", sessionId
        Else
            Debug.Print "BiDi Warning: CDP session requested but not found."
        End If
    End If

    ExecuteCDPCommand = ExecuteBiDiCommand("goog:cdp.sendCommand", params)
End Function

' Inject the SPA idle probe into the current page using script.evaluate (idempotent).
' Prefer "window" realm for stability; falls back to context if realm not available.
Public Sub StartSpaIdleProbe(Optional ByVal contextId As String = "")
    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    Dim p As New Dictionary, t As New Dictionary

    If realmId <> "" Then
        t.Add "realm", realmId
    Else
        t.Add "context", UpdateMainContextId(True)
    End If

    p.Add "expression", GetProbeInjectionExpression()
    p.Add "target", t
    p.Add "awaitPromise", False

    ExecuteBiDiCommand "script.evaluate", p
End Sub

' Build a small function to reset probe counters (especially nodeCountDelta).
Private Function GetProbeResetCountersExpression() As String
    Dim js As String
    js = ""
    js = js & "function(){"
    js = js & "var p=window.__vbaIdleProbe;"
    js = js & "if(!p)return false;"
    js = js & "p.nodeCountDelta=0;"
    js = js & "return true;"
    js = js & "}"
    GetProbeResetCountersExpression = js
End Function

' Reset SPA probe counters so the stability window starts clean.
Private Sub ResetSpaProbeCounters(Optional ByVal contextId As String = "")
    Dim realmId As String: realmId = GetScriptRealmId(contextId)
    If realmId = "" Then Exit Sub

    Dim p As New Dictionary, t As New Dictionary
    t.Add "realm", realmId
    p.Add "functionDeclaration", GetProbeResetCountersExpression()
    p.Add "arguments", New Collection
    p.Add "target", t
    p.Add "awaitPromise", False

    ExecuteBiDiCommand "script.callFunction", p
End Sub
