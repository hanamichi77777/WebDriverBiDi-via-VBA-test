Attribute VB_Name = "A_01_BiDi_Sample"
Option Explicit

' Message box that is always displayed in the foreground
Public Declare PtrSafe Function MESSAGEbox Lib "user32.dll" Alias "MessageBoxA" _
                                (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long
Public Const MB_OK = &H0                         ' OK button flag
Public Const MB_ForeFront = &H40000              ' Topmost flag
Public Const MB_ICONINFORMATION As Long = &H40
Public Const MB_ICONERROR As Long = &H10
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

' [Select Box (Wait for completion if an event occurs)]

Public Sub Main01()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
  ' Start
  .StartChrome
    
  ' Browser startup settings (for both Chrome and Edge)
  Dim caps As SeleniumVBA.WebCapabilities: Set caps = .CreateCapabilities
  ' /Open maximized
  caps.AddArguments "--start-maximized"
  ' /Do not show intrusive guidance messages from Chrome
  caps.AddArguments "--propagate-iph-for-testing"
    
  ' Required to enable Chrome extensions
  caps.AddArguments "--remote-debugging-pipe"
  caps.AddArguments "--enable-unsafe-extension-debugging"
  ' ==========================================
  ' Enable BiDi (True is mandatory for this program)
  caps.EnableBiDiMode
  ' ==========================================
      
  ' Open
  .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' ==========================================
          
' --- 1. Execute BiDi Commands ---

    ' Enable Chrome extension(Please install the Google Translate extension in advance.)
    Dim extensionPath As String
    extensionPath = Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\Default\Extensions\aapbdbdomjkkjkaonfhkkikfgjllcleb\2.0.16_0"
    bidi.ExecuteWebExtensionInstall extensionPath
    
      
    ' Navigate to test.htmlÅiPlease place "test.html" in the same directory as this file.Åj
    Dim url As String: url = "file:///" & Replace(driver.ResolvePath(".\") & "\test.html", "\", "/")
    bidi.ExecuteNavigateAndGetStatus url
    
    Dim msgText As String, msgCaption As String
    msgText = "Google Translate extension installed."
    msgCaption = "Success"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
      
' --- 2. Search for XPath element and execute click ---
    ' Search for XPath element and execute click (Argument is the Value of the Option tag)
    bidi.ExecuteSelectValueByXPath "//select[@id='userSelector']", "1", False, , True
      
' --- 3. Verification and Termination ---
    Dim str As String
    ' Check if the calendar switched as expected
    str = .FindElementByXPath("//input[@id='nameField']").GetProperty("value")
      
    If str = "Success: Response Received!" Then
        msgText = "Successfully waited until the document switched."
        msgCaption = "Verification Complete"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    Else
        msgText = "The document has not switched. Retrieved value: " & str
        msgCaption = "Verification Failed"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    End If
      
    ' End
    Set bidi = Nothing
    .CloseBrowser
    .Shutdown
      
  End With
End Sub

Public Sub Main02()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
   
  ' Start
  .StartChrome
   
  ' Browser startup settings (for both Chrome and Edge)
  Dim caps As SeleniumVBA.WebCapabilities: Set caps = .CreateCapabilities
  ' /Open maximized
  caps.AddArguments "--start-maximized"
  ' ==========================================
  ' Enable BiDi (True is mandatory for this program)
  caps.EnableBiDiMode
  ' ==========================================
    
  ' Open
  .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' ==========================================
             
   bidi.ExecuteRegisterAutoClickerByXPath ("//input[@type='submit'][contains(@value,'ï\é¶')]")
    
    ' Navigate to page
    Dim url As String: url = "https://note.com/topic/novel"
    Dim statusCode As Long
    statusCode = bidi.ExecuteNavigateAndGetStatus(url, True)
    
' --- 2. Wait process verification ---
    Dim elms_title1 As WebElements ' List of article elements
    Dim elms_title2 As WebElements ' List of article elements (after waiting)
    
    ' [1st time] Search article count with FindElements
    Set elms_title1 = .FindElements(By.cssSelector, ".a-link.m-largeNoteWrapper__link.fn")
      
    ' Wait 4 seconds
    .Wait 4000
      
    ' [2nd time] Search article count with FindElements
    Set elms_title2 = .FindElements(By.cssSelector, ".a-link.m-largeNoteWrapper__link.fn")
    
    ' [Verification of page load completion]
    Dim msgText As String, msgCaption As String
    If elms_title1.Count <> elms_title2.Count Then
      msgText = "Waited, but" & Chr(10) & "statusCode: " & statusCode & Chr(10) & " - Initial article count: " & elms_title1.Count & Chr(10) & " - Article count after 4 sec: " & elms_title2.Count & Chr(10) & " therefore the wait time is insufficient."
      msgCaption = "Wait Insufficient statusCode: " & statusCode
      MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    Else
      msgText = "Waited, but" & Chr(10) & "statusCode: " & statusCode & Chr(10) & " - Initial article count: " & elms_title1.Count & Chr(10) & " - Article count after 4 sec: " & elms_title2.Count & Chr(10) & " therefore it waited as expected."
      msgCaption = "Wait Complete"
      MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    End If
      
      
  ' End
   Set bidi = Nothing
  .CloseBrowser
  .Shutdown
End With

End Sub

' [Text box input (Wait for completion if an event occurs)]
Public Sub Main03()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
  ' Start
  .StartChrome
    
  ' Browser startup settings (for both Chrome and Edge)
  Dim caps As SeleniumVBA.WebCapabilities: Set caps = .CreateCapabilities
  ' /Open maximized
  caps.AddArguments "--start-maximized"
  ' ==========================================
  ' Enable BiDi (True is mandatory for this program)
  caps.EnableBiDiMode
  ' ==========================================
    
  ' Open
  .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' ==========================================
    
'  ' Setting to block images and ads
  Dim blockList As Variant
  ' Example patterns to block common image formats and ad services
  blockList = Array( _
    "*.png", "*.jpg", "*.jpeg", "*.gif", "*.svg", "*.woff2", _
    "*ad_service*", "*analytics*", "*googletagmanager*", _
    "*doubleclick*", "*googlesyndication*", "*amazon-adsystem*", _
    "*criteo*", "*adnxs*", "*teads*", "*popin*", "*logly*", _
    "*microad*", "*fout*", "*yads*", "*yjt*", _
    "*facebook.net*", "*scorecardresearch*", _
    "*/collect*", "*/beacon*")

  ' Method name changed to match CDP implementation
  bidi.ExecuteEnableResourceBlocking blockList

  ' Registering this prior to navigation ensures the button is clicked automatically the instant it appears in the DOM.
  bidi.ExecuteRegisterAutoClickerByXPath "//button[@id='search_button_main']", 10000

  ' Page transition
  Dim url As String: url = "https://world.jorudan.co.jp/mln/en/"
  .NavigateTo url

  ' Departure: Tokyo
  bidi.ExecuteInputValueByXPath "//input[@id='from_value']", "Tokyo"
  ' Arrival: Shinjuku
  bidi.ExecuteInputValueByXPath "//input[@id='to_value']", "Shinjuku"
  ' Click search button
  bidi.ExecuteClickByXPath "//button[@id='search_button_main1']", , True
   
  ' End
  Set bidi = Nothing
  .CloseBrowser
  .Shutdown
    
  ' Completion
  MsgBox "Completed"
    
End With
End Sub

' [Login Wait (True BiDi Implementation)]
Public Sub Main04()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
     
    .StartChrome
    
    ' Browser startup settings
    Dim caps As SeleniumVBA.WebCapabilities: Set caps = .CreateCapabilities
    caps.AddArguments "--start-maximized"
    ' ==========================================
    ' Enable BiDi (Mandatory)
    caps.EnableBiDiMode
    ' ==========================================
      
    ' Open
    .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' ==========================================
        
    ' --- Execute BiDi Wait Logic ---
    ' Navigate to Login Page
    Dim loginUrl As String: loginUrl = "https://hotel-example-site.takeyaqa.dev/ja/login.html"
    'userName = "ichiro@example.com"
    'pw = "password"
    bidi.ExecuteNavigateAndGetStatus loginUrl, True
      
    Dim isLoginSuccess As Boolean
    isLoginSuccess = bidi.ExecuteIsUrlContains("https://hotel-example-site.takeyaqa.dev/ja/mypage.html", True, , 30000)
      
    ' Verification
    Dim msgText As String, msgCaption As String
      
    If isLoginSuccess Then
        msgText = "BiDi Event Received!" & vbCrLf & "Login (Navigation) Confirmed."
        msgCaption = "Success"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    Else
        msgText = "Timed out while waiting for login event."
        msgCaption = "Failed"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    End If
      
    ' End
    Set bidi = Nothing
    .CloseBrowser
    .Shutdown
      
  End With
End Sub

' Frame Piercing
Public Sub Main06()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
  ' Start
  .StartChrome
    
  ' Browser startup settings
  Dim caps As SeleniumVBA.WebCapabilities: Set caps = .CreateCapabilities
  caps.AddArguments "--start-maximized"
  ' ==========================================
  ' Enable BiDi (Mandatory)
  caps.EnableBiDiMode
  ' ==========================================
    
  ' Open
  .OpenBrowser caps
  ' ==========================================
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
  ' =========================================
   .NavigateTo "https://www.customs.go.jp/toukei/srch/index.htm?M=01&P=0"
   
   ' Frame Piercing
   Dim conID As String
   conID = bidi.GetIframeContextIdByUrl("jccht00d")
   bidi.ExecuteClickByXPath "//input[@id='la_imp']", , , , , conID
   
  ' End
  Set bidi = Nothing
  .CloseBrowser
  .Shutdown
   
  End With
End Sub

' Shadow DOM Interaction (Click, Input) with WAF Evasion Retry Loop
' Input: selectorsArray (Array of Strings) -> e.g. Array("user-card", "paper-button")
Public Sub Main07()
    Dim driver As New WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    Dim bidi As BiDiCommandWrapper
    Dim targetUrl As String: targetUrl = "https://developer.servicenow.com/"
    
    Dim retryCount As Integer: retryCount = 0
    Dim maxRetries As Integer: maxRetries = 3
    Dim isSuccess As Boolean: isSuccess = False
    
    With driver
        Do While retryCount < maxRetries And Not isSuccess
            .StartChrome

            Debug.Print "--- Attempt " & retryCount + 1 & " Start ---"
            
            ' 1. Create clean Capabilities for each attempt
            Set caps = .CreateCapabilities
            caps.EnableBiDiMode
            caps.AddArguments "--disable-blink-features=AutomationControlled"
            caps.AddExcludeSwitches "enable-automation"
            
            ' [CRUCIAL] Recovery Key: Rotate User-Agent per attempt
            ' Modify the minor version with the retry index to offset the browser fingerprint
            caps.AddArguments "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) " & _
                             "AppleWebKit/537.36 (KHTML, like Gecko) " & _
                             "Chrome/131.0.0." & retryCount & " Safari/537.36"

            ' 2. Browser Initialization (OpenBrowser is mandatory [cite: 2025-10-02])
            .OpenBrowser caps
            Set bidi = New BiDiCommandWrapper: bidi.ConnectTo driver.GetWebSocketUrl
            
            ' 3. Navigation Execution
            ' Register auto-clicker for the consent banner before navigation
            bidi.ExecuteRegisterAutoClickerByXPath "//button[@id='truste-consent-button']"
            ' wait="none" is acceptable as the subsequent health check handles synchronization
            bidi.ExecuteNavigateAndGetStatus targetUrl, False
            
            ' 4. Health Check: Verify if dps-app has the [component-id] attribute
            ' This confirms the JS has successfully executed and initialized the SPA
            Dim checkStart As Double: checkStart = Timer
            Dim isInitialized As Boolean: isInitialized = False
            
            Do While Timer < checkStart + 7 ' Sufficient duration for heavy JS initialization
                ' waitForIdle = False ensures a snappy check without unnecessary network wait
                If bidi.ExecuteIsElementVisible("//dps-app[@component-id]", 1000, False) = True Then
                    isInitialized = True
                    Exit Do
                End If
                .Wait 500
            Loop
            
            ' 5. Validation and Branching logic
            If isInitialized Then
                Debug.Print "[WIN] dps-app[component-id] detected. Proceeding with operation."
                isSuccess = True
            Else
                Debug.Print "[FAIL] WAF detection caught (attribute not injected). Destroying session for retry."
                .CloseBrowser
                .Shutdown
                retryCount = retryCount + 1
                ' Wait interval to allow the WAF's short-term request threshold to cool down
                .Wait 3000
            End If
        Loop

        ' --- Main Operations: Only executed if health check passed ---
        If isSuccess Then
            ' Target Path: Shadow DOM hierarchy for Sign-In button
            Dim path(3) As String
            path(0) = "dps-app"
            path(1) = "dps-navigation-header"
            path(2) = "sn-cx-navigation"
            path(3) = "#utility-sign-in button"  ' Maps to: //*[@id='utility-sign-in']//button
            
            ' Execute Click in Shadow DOM
            bidi.ExecuteShadowClick path, 20000, True
            
            ' Execute Input on Sign-In page (Light DOM)
            bidi.ExecuteInputValueByXPath "//input[@id='username']", "aaa", 20000, True
            
            ' Cleanup
            Set bidi = Nothing
            .CloseBrowser
            .Shutdown
        
        Else
            MsgBox "Failed to bypass WAF after maximum retry attempts.", vbCritical
        End If
           
    End With
End Sub

' Recorder
Sub Main09()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
    .StartChrome
    
    ' Browser startup settings
    Dim caps As SeleniumVBA.WebCapabilities: Set caps = .CreateCapabilities
    caps.AddArguments "--start-maximized"
    ' ==========================================
    ' Enable BiDi (Mandatory)
    caps.EnableBiDiMode
    ' ==========================================
      
    ' Open
    .OpenBrowser caps
    ' ==========================================
    Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
    ' ==========================================
        
    ' Navigate to Page
    Dim url As String: url = "https://note.com/"
    .NavigateTo url
    
    ' ==========================================================
    ' Start Recording & Wait
    Const RECORDING_SECONDS As Long = 20
    
    ' Show Message (Blocks execution until OK is clicked)
    Dim msgText As String, msgCaption As String
    msgText = "Please prepare the browser for recording." & vbCrLf & vbCrLf & _
              "Click [OK] to start recording." & vbCrLf & _
              "Duration: " & RECORDING_SECONDS & " seconds." & vbCrLf & _
              "Please manually interact with the page immediately after clicking OK."
    msgCaption = "Ready to Record"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    
    'Start logging (True = exclude images/css noise)
    bidi.StartDiscoveryLog excludeImagesAndCss:=True
    ' Wait and process events for the specified duration
    bidi.RecordEventsForSeconds RECORDING_SECONDS
    
    ' Save Log
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath
    
    MsgBox "Discovery Log Saved!" & vbCrLf & logPath
    ' ==========================================================
    
    ' End
    Set bidi = Nothing
    .CloseBrowser
    .Shutdown
    
End With
End Sub

Public Sub test_Classic_wait()
    Dim driver As New WebDriver
    Dim elem As WebElement

    driver.StartChrome
    driver.OpenBrowser
    driver.NavigateTo "file:///" & Replace(driver.ResolvePath(".\") & "\test.html", "\", "/")

    InstallNetworkHooks driver

    Set elem = driver.FindElementByXPath("//select[@id='userSelector']")
    elem.SelectByValue "1"

    WaitForIdleStateAsync driver

    Debug.Assert driver.FindElementByXPath("//input[@id='nameField']").GetProperty("value") = "Success: Response Received!"

    driver.CloseBrowser
    driver.Shutdown
End Sub

Private Sub InstallNetworkHooks(oDriver As WebDriver)
    Dim js As String
    js = vbNullString
    js = js & "if(!window.__niw){(function(){" & vbCrLf
    js = js & "var s={req:0,of:window.fetch,oo:XMLHttpRequest.prototype.open,os:XMLHttpRequest.prototype.send,onchange:null};" & vbCrLf
    js = js & "window.fetch=function(){s.req++;if(s.onchange)s.onchange();var p=s.of.apply(this,arguments);return p.finally(function(){s.req--;if(s.onchange)s.onchange();});};" & vbCrLf
    js = js & "XMLHttpRequest.prototype.open=function(){return s.oo.apply(this,arguments);};" & vbCrLf
    js = js & "XMLHttpRequest.prototype.send=function(){s.req++;if(s.onchange)s.onchange();this.addEventListener('loadend',function(){s.req--;if(s.onchange)s.onchange();});return s.os.apply(this,arguments);};" & vbCrLf
    js = js & "window.__niw=s;" & vbCrLf
    js = js & "})();}" & vbCrLf
    oDriver.ExecuteScript js
End Sub

Private Sub WaitForIdleStateAsync(oDriver As WebDriver, Optional idleTimeout As Long = 500, Optional maxTimeToWait As Long = 30000)
    Dim js As String, ret As Variant
    js = vbNullString
    js = js & "var idleTimeout=arguments[0],maxTimeout=arguments[1],cb=arguments[2];" & vbCrLf
    js = js & "var s=window.__niw;if(!s){cb('not-armed');return;}" & vbCrLf
    js = js & "var tm=null,st=false,safety=setTimeout(function(){done('timeout');},maxTimeout);" & vbCrLf
    js = js & "function done(v){if(st)return;st=true;if(tm)clearTimeout(tm);clearTimeout(safety);cb(v);}" & vbCrLf
    js = js & "function schedule(){if(tm)clearTimeout(tm);if(s.req===0){tm=setTimeout(function(){done('ok');},idleTimeout);}}" & vbCrLf
    js = js & "s.onchange=schedule;" & vbCrLf
    js = js & "schedule();" & vbCrLf
    ret = oDriver.ExecuteScriptAsync(js, idleTimeout, maxTimeToWait)
    If ret <> "ok" Then Err.Raise 404, , "Maximum time exceeded or context lost: " & ret
End Sub

