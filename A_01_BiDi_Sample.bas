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
       
    ' Navigate to test.html
    Dim html As String
    html = "<!DOCTYPE html><html lang=""en""><head><meta charset=""UTF-8""><title>BiDi Tester</title><style>body{font-family:sans-serif;padding:20px;line-height:1.6}.container{border:1px solid #333;padding:20px;width:450px;background:#fff}label{font-weight:bold;display:block;margin-bottom:5px}select{width:100%;padding:8px;margin-bottom:20px}input{width:100%;padding:10px;border:1px solid #ccc;font-size:1em;color:#0056b3;background:#f0f8ff}.note{font-size:.85em;color:#666;margin-top:15px;border-top:1px dashed #ccc;padding-top:10px}</style></head>"
    html = html & "<body><div class=""container""><h3>Server-Side Delay Tester (5s)</h3><p>Proves BiDi can track long-running network requests.</p><label for=""userSelector"">Select Action:</label><select id=""userSelector"" onchange=""triggerLongFetch()""><option value="""">-- Choose to Trigger --</option><option value=""1"">Start 5-Second Server Request</option></select>"
    html = html & "<label for=""nameField"">Network/DOM Status:</label><input type=""text"" id=""nameField"" readonly placeholder=""Idle...""><div class=""note""><b>Mechanism:</b><br>1. Immediate DOM change (0ms).<br>2. Fetch starts immediately and stays open for <b>5s</b> using httpbin.org.<br>3. Final update occurs only after the server responds.</div></div>"
    html = html & "<script>function triggerLongFetch(){const s=document.getElementById('userSelector'),n=document.getElementById('nameField');if(s.value==="""")return;n.value=""Requesting... (Connection Open for 5s)"";fetch('https://httpbin.org/delay/5').then(r=>{if(!r.ok)throw new Error('Network error');return r.json()}).then(d=>{n.value=""Success: Response Received!""}).catch(e=>{n.value=""Error: Server unreachable or Timeout."";console.error(e)})}</script></body></html>"
    
    .NavigateToString html
    
    Dim msgText As String, msgCaption As String
    msgText = "Google Translate extension installed."
    msgCaption = "Success"
    MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront


' --- 2. Search for XPath element and execute click ---

    ' Start Recording BEFORE the action
    bidi.StartDiscoveryLog

    ' Search for XPath element and execute click (Argument is the Value of the Option tag)
    bidi.ExecuteSelectValueByXPath "//select[@id='userSelector']", "1", False, , True

    ' Stop and Save AFTER the wait is finished
    Dim logPath As String
    logPath = .ResolvePath(".\") & "\discovery_log.txt"
    bidi.StopAndSaveDiscoveryLog logPath

' --- 3. Verification and Termination ---
    Dim str As String
    ' Check if the calendar switched as expected
    str = .FindElementByXPath("//input[@id='nameField']").GetProperty("value")

    If str = "Success: Response Received!" Then
        msgText = "Successfully waited until the document switched." & Chr(10) & "See the discovery_log.txt"
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
             
   bidi.ExecuteRegisterAutoClickerByXPath ("//input[@type='submit'][contains(@value,'•\Ž¦')]")
    
    ' Navigate to page
    Dim url As String: url = "https://note.com/topic/novel"
    Dim statusCode As Long
    statusCode = bidi.ExecuteNavigateAndGetStatus(url, True)
    
' --- 2. Wait process verification ---
    Dim elms_title1 As WebElements ' List of article elements
    Dim elms_title2 As WebElements ' List of article elements (after waiting)
    
    ' [1st time] Search article count with FindElements
    Set elms_title1 = .FindElements(By.xpath, "//div[contains(@class,'flex w-full rounded-lg bg-surface-normal')]")
    ' Wait 4 seconds
    .Wait 4000
      
    ' [2nd time] Search article count with FindElements
    Set elms_title2 = .FindElements(By.xpath, "//div[contains(@class,'flex w-full rounded-lg bg-surface-normal')]")
    
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
  bidi.ExecuteClickByXPath "//button[@id='search_button_main1']"
   
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

' Wait for async event completion
Public Sub Main05()
    Dim driver As New WebDriver
    Dim elem As WebElement

    driver.StartChrome
    
   ' Browser startup settings
    Dim caps As SeleniumVBA.WebCapabilities: Set caps = driver.CreateCapabilities
    caps.EnableBiDiMode
    
    ' Open
    driver.OpenBrowser caps
    Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo driver.GetWebSocketUrl

    driver.NavigateTo "https://www.selenium.dev/selenium/web/ajaxy_page.html"

    driver.FindElement(By.xpath, "//input[@name='typer']").SendKeys "aaa"
    driver.FindElement(By.xpath, "//input[@id='red']").Click
    
    bidi.ExecuteClickByXPath "//input[@value='Add Label']"

    Debug.Assert driver.FindElement(By.xpath, "//div[@id='update_butter']").GetText = "Done!"

    Set bidi = Nothing
    driver.CloseBrowser
    driver.Shutdown
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
Public Sub Main07()
    Dim driver As New WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    Dim bidi As BiDiCommandWrapper
    Dim targetUrl As String: targetUrl = "https://developer.servicenow.com/"
        
    With driver
    .StartChrome
    Set caps = .CreateCapabilities
    
    caps.EnableBiDiMode
    .OpenBrowser caps
    Set bidi = New BiDiCommandWrapper: bidi.ConnectTo driver.GetWebSocketUrl
        
    ' Register auto-clicker for the consent banner before navigation
    bidi.ExecuteRegisterAutoClickerByXPath "//button[@id='truste-consent-button']"
       
    ' NavigateTo Page
    bidi.ExecuteNavigateAndGetStatus targetUrl, False
   
    ' Execute Click in Shadow DOM
    bidi.ExecuteShadowClick Array("dps-app", "dps-navigation-header", "sn-cx-navigation", "#utility-sign-in button"), 10000
            
    ' Execute Input on Sign-In page
    bidi.ExecuteInputValueByXPath "//input[@id='username']", "aaa", 10000
            
    ' Cleanup
    Set bidi = Nothing
    .CloseBrowser
    .Shutdown
           
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
    
    Dim html As String
    html = "<!DOCTYPE html><html lang=""en""><head><meta charset=""UTF-8""><title>BiDi Tester</title><style>body{font-family:sans-serif;padding:20px;line-height:1.6}.container{border:1px solid #333;padding:20px;width:450px;background:#fff}label{font-weight:bold;display:block;margin-bottom:5px}select{width:100%;padding:8px;margin-bottom:20px}input{width:100%;padding:10px;border:1px solid #ccc;font-size:1em;color:#0056b3;background:#f0f8ff}.note{font-size:.85em;color:#666;margin-top:15px;border-top:1px dashed #ccc;padding-top:10px}</style></head>"
    html = html & "<body><div class=""container""><h3>Server-Side Delay Tester (5s)</h3><p>Proves BiDi can track long-running network requests.</p><label for=""userSelector"">Select Action:</label><select id=""userSelector"" onchange=""triggerLongFetch()""><option value="""">-- Choose to Trigger --</option><option value=""1"">Start 5-Second Server Request</option></select>"
    html = html & "<label for=""nameField"">Network/DOM Status:</label><input type=""text"" id=""nameField"" readonly placeholder=""Idle...""><div class=""note""><b>Mechanism:</b><br>1. Immediate DOM change (0ms).<br>2. Fetch starts immediately and stays open for <b>5s</b> using httpbin.org.<br>3. Final update occurs only after the server responds.</div></div>"
    html = html & "<script>function triggerLongFetch(){const s=document.getElementById('userSelector'),n=document.getElementById('nameField');if(s.value==="""")return;n.value=""Requesting... (Connection Open for 5s)"";fetch('https://httpbin.org/delay/5').then(r=>{if(!r.ok)throw new Error('Network error');return r.json()}).then(d=>{n.value=""Success: Response Received!""}).catch(e=>{n.value=""Error: Server unreachable or Timeout."";console.error(e)})}</script></body></html>"
    
    driver.NavigateToString html

    InstallNetworkAndDomHooks driver
    driver.FindElementByXPath("//select[@id='userSelector']").SelectByValue "1"
    WaitForIdleStateAsync driver

    Set elem = driver.FindElementByXPath("//input[@id='nameField']")
    Debug.Assert elem.GetProperty("value") = "Success: Response Received!"

    driver.CloseBrowser
    driver.Shutdown
End Sub

Private Sub InstallNetworkAndDomHooks(oDriver As WebDriver)
    Dim js As String
    js = ""

    ' Guard & state
    js = js & "if(!window.__niw){(function(){"
    js = js & "var s={req:0,lastMut:Date.now(),of:window.fetch,oo:XMLHttpRequest.prototype.open,os:XMLHttpRequest.prototype.send,onchange:null};"
    js = js & "function markChanged(){ s.lastMut=Date.now(); if(s.onchange) s.onchange(); }"

    ' fetch hook
    js = js & "try{window.fetch=function(){ s.req++; if(s.onchange) s.onchange(); return s.of.apply(this,arguments).finally(function(){ s.req--; if(s.onchange) s.onchange(); }); };}catch(e){}"

    ' XHR hook
    js = js & "try{XMLHttpRequest.prototype.open=function(){ return s.oo.apply(this,arguments); }; XMLHttpRequest.prototype.send=function(){ s.req++; if(s.onchange) s.onchange(); this.addEventListener('loadend',function(){ s.req--; if(s.onchange) s.onchange(); }); return s.os.apply(this,arguments); };}catch(e){}"

    ' MutationObserver
    js = js & "try{ new MutationObserver(function(){ markChanged(); }).observe(document,{childList:true,subtree:true,attributes:true}); }catch(e){}"

    ' value setter: input
    js = js & "try{ var d1=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value'); if(d1&&d1.configurable){ Object.defineProperty(HTMLInputElement.prototype,'value',{ get:d1.get, set:function(v){ markChanged(); return d1.set.call(this,v); } }); } }catch(e){}"

    ' value setter: textarea
    js = js & "try{ var d2=Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype,'value'); if(d2&&d2.configurable){ Object.defineProperty(HTMLTextAreaElement.prototype,'value',{ get:d2.get, set:function(v){ markChanged(); return d2.set.call(this,v); } }); } }catch(e){}"

    ' expose
    js = js & "window.__niw=s;})();}"

    oDriver.ExecuteScript js
End Sub

Private Sub WaitForIdleStateAsync( _
    oDriver As WebDriver, _
    Optional idleTimeout As Long = 1000, _
    Optional maxTimeToWait As Long = 30000)

    Dim js As String, ret As Variant
    js = ""

    ' Receive arguments: idle quiet-window (ms), max overall timeout (ms), async callback
    js = js & "var idleTimeout=arguments[0], maxTimeout=arguments[1], cb=arguments[2];"

    ' Guard: hooks must be installed (__niw) or exit early
    js = js & "var s=window.__niw; if(!s){ cb('not-armed'); return; }"

    ' Timers: tm=quiet-window timer, safety=global timeout; settled prevents double completion
    js = js & "var tm=null, settled=false, safety=setTimeout(function(){ done('timeout'); }, maxTimeout);"

    ' Finish: ensure single completion, cleanup timers, then callback
    js = js & "function done(v){ if(settled) return; settled=true; if(tm) clearTimeout(tm); clearTimeout(safety); cb(v); }"

    ' Final re-check: verify conditions right before resolving to avoid races
    js = js & "function tryResolve(){ var quietFor=Date.now()-s.lastMut; if(s.req===0 && quietFor>=idleTimeout){ done('ok'); } else { schedule(); } }"

    ' Scheduler: when no network is in-flight (req==0), wait remaining quiet time then re-check
    js = js & "function schedule(){ if(tm) clearTimeout(tm); var quietFor=Date.now()-s.lastMut; if(s.req===0){ var remain=Math.max(0,idleTimeout-quietFor); var delay=Math.max(50,remain); tm=setTimeout(tryResolve,delay); } }"

    ' Subscribe to changes; defer the initial check by one tick to avoid s.req++ races
    js = js & "s.onchange=schedule; setTimeout(schedule,0);"

    ret = oDriver.ExecuteScriptAsync(js, idleTimeout, maxTimeToWait)
    If ret <> "ok" Then Err.Raise 404, , "Wait failed or timed out: " & ret
End Sub
