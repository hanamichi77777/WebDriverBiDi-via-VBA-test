Attribute VB_Name = "BiDi_Sample"
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
  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
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

    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
      
  End With
End Sub

Public Sub Main02()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
   
  ' Start
  .StartEdge
   
  ' Browser startup settings (for both Chrome and Edge)
  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
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
    ' Navigate to page
    Dim url As String: url = "https://note.com/topic/novel"
    Dim statusCode As String
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
    If elms_title1.count <> elms_title2.count Then
      msgText = "Waited, but" & Chr(10) & "statusCode: " & statusCode & Chr(10) & " - Initial article count: " & elms_title1.count & Chr(10) & " - Article count after 4 sec: " & elms_title2.count & Chr(10) & " therefore the wait time is insufficient."
      msgCaption = "Wait Insufficient statusCode: " & statusCode
      MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    Else
      msgText = "Waited, but" & Chr(10) & "statusCode: " & statusCode & Chr(10) & " - Initial article count: " & elms_title1.count & Chr(10) & " - Article count after 4 sec: " & elms_title2.count & Chr(10) & " therefore it waited as expected."
      msgCaption = "Wait Complete"
      MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    End If
      
      
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
End With

End Sub

' [Text box input (Wait for completion if an event occurs)]
Public Sub Main03()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
  ' Start
  .StartEdge
    
  ' Browser startup settings (for both Chrome and Edge)
  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
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

  ' Page transition
  Dim url As String: url = "https://world.jorudan.co.jp/mln/en/"
  bidi.ExecuteNavigateAndGetStatus url

  ' Departure: Tokyo
  bidi.ExecuteInputValueByXPath "//input[@id='from_value']", "Tokyo"
  ' Arrival: Shinjuku
  bidi.ExecuteInputValueByXPath "//input[@id='to_value']", "Shinjuku"
  ' Click search button
  bidi.ExecuteClickByXPath "//button[starts-with(@id, 'search_button_main')]"
  ' Click search button
  bidi.ExecuteClickByXPath "//button[@id='search_button_main']"
   
  ' Cleanup
  bidi.Shutdown: Set bidi = Nothing
  .CloseBrowser: .Shutdown
    
  ' Completion
  MsgBox "Completed"
    
End With
End Sub

' [Login Wait (True BiDi Implementation)]
Public Sub Main04()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
     
    .StartEdge
    
    ' Browser startup settings
    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
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
      
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
      
  End With
End Sub

' Wait for async event completion
Public Sub Main05()
    Dim driver As WebDriver: Set driver = New WebDriver
    With driver

    .StartEdge
    
   ' Browser startup settings
    Dim caps As WebCapabilities: Set caps = driver.CreateCapabilities
    caps.EnableBiDiMode
    
    ' Open
    .OpenBrowser caps
    Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl

    .NavigateTo "https://www.selenium.dev/selenium/web/ajaxy_page.html"

    'Specify False if waiting for the completion of the asynchronous event is not required
    bidi.ExecuteInputValueByXPath "//input[@name='typer']", "aaa", , False
    bidi.ExecuteClickByXPath "//input[@id='red']", , False
    
    'Wait for the asynchronous event that occurs after clicking the AddLabel button
    bidi.ExecuteClickByXPath "//input[@value='Add Label']", , , 1000

    Debug.Assert driver.FindElement(By.xpath, "//div[@id='update_butter']").GetText = "Done!"

    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
   End With
End Sub

' Frame Piercing
Public Sub Main06()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
  ' Start
  .StartEdge
    
  ' Browser startup settings
  Dim caps As WebCapabilities: Set caps = .CreateCapabilities
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
   bidi.ExecuteNavigateAndGetStatus "https://www.customs.go.jp/toukei/srch/index.htm?M=01&P=0", False
   
   ' Frame Piercing
   Dim conID As String
   conID = bidi.GetIframeContextIdByUrl("jccht00d")
   bidi.ExecuteClickByXPath "//input[@id='la_imp']", , , , , conID
   
   ' Cleanup
   bidi.Shutdown: Set bidi = Nothing
   .CloseBrowser: .Shutdown
   
  End With
End Sub

' Shadow DOM Interaction (Click, Input) with WAF Evasion Retry Loop
Public Sub Main07()
    Dim driver As New WebDriver
    Dim caps As WebCapabilities
    Dim bidi As BiDiCommandWrapper
    Dim targetUrl As String: targetUrl = "https://developer.servicenow.com/"
        
    With driver
    .StartEdge
    Set caps = .CreateCapabilities
    
    caps.EnableBiDiMode
    .OpenBrowser caps
    Set bidi = New BiDiCommandWrapper: bidi.ConnectTo driver.GetWebSocketUrl
        
    ' Register auto-clicker for the consent banner before navigation
    bidi.ExecuteRegisterAutoClickerByXPath "//button[@id='truste-consent-button']"
       
    ' NavigateTo Page
    bidi.ExecuteNavigateAndGetStatus targetUrl
   
    ' Execute Click in Shadow DOM
    bidi.ExecuteShadowClick "#utility-sign-in button"
            
    ' Execute Input on Sign-In page
    bidi.ExecuteInputValueByXPath "//input[@id='username']", "aaa"
            
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
           
    End With
End Sub

Public Sub Main08()
' ========================================================================================
' Google Flights Test Program (Final)
'
' PURPOSE :
'   Stress test for BiDiCommandWrapper on a heavy, reactive SPA (Google Flights).
'   Exercises:
'     - React/Wiz-controlled combobox inputs with dynamic suggestions
'     - High-frequency background network traffic (XHR/fetch)
'     - SPA transitions and DOM mutation storms
' TARGET  :
'   https://www.google.com/travel/flights
' SCENARIO :
'   Sapporo Å® Paris, round trip, then click Search
' NOTE :
'   Google Flights uses Wiz/Lit-driven comboboxes that spawn multiple <input>
'   elements on activation ? a hidden store input, a surface trigger input, and
'   the actual editable input (always the last one in DOM order).
'   The input_keys action handles this transparently:
'     Phase 0 : Click the resolved element, then poll document.activeElement
'               with a stability check (same element observed on 2 consecutive
'               polls) to skip transient intermediate inputs and lock onto the
'               final replacement.  Works identically for SPAs that replace
'               the input (Google Flights) and traditional forms that don't
'               (Jorudan, ServiceNow).
'     Phase 1 : Clear existing value (select Å® delete Å® forwardDelete Å® fallback)
'     Phase 2 : Per-character insertText with activeElement tracking
'               (each character is confirmed via selectionStart before next)
'     Phase 3 : Final validation ? throws if value length < expected
'   This guarantees:
'     - No "first character swallowed" issue
'     - Trusted InputEvent firing for each character
'     - Suggestion dropdown reacts correctly
'     - SPA IdleProbe remains stable (CallScript bypasses ExecuteBaseAction)
' XPATH STRATEGY :
'   - Use [last()] for inputs that may exist as multiple copies in DOM
'     (e.g. "(//input[contains(@aria-label, 'Where from')])[last()]")
'     This is the first line of defense; Phase 0 is the safety net.
'   - aria-label may contain trailing whitespace Å® use contains()
'   - Suggestion order varies Å® select by text content, not index
' LESSONS LEARNED :
'   - Google background telemetry (/log?, /gen_204, ogs.google.com) must be ignored
'     or SPA Idle consensus never stabilizes
'   - input_keys mode is the only reliable method for React/Wiz comboboxes
' ========================================================================================

    Dim driver As WebDriver: Set driver = New WebDriver
    With driver
        ' ==========================================
        ' Browser Setup
        ' ==========================================
        .StartEdge
        
        Dim caps As WebCapabilities: Set caps = .CreateCapabilities
        caps.AddArguments "--start-maximized"
        caps.AddArguments "--lang=en"
        caps.EnableBiDiMode
        
        .OpenBrowser caps
        
        ' ==========================================
        ' BiDi Connection
        ' ==========================================
        Dim bidi As New BiDiCommandWrapper
        bidi.ConnectTo .GetWebSocketUrl
        
        ' ==========================================
        ' Resource Blocking (Images, Ads, Analytics)
        ' ==========================================
        Dim blockList As Variant
        blockList = Array( _
            "*.png", "*.jpg", "*.jpeg", "*.gif", "*.svg", "*.woff2", _
            "*googletagmanager*", "*doubleclick*", "*googlesyndication*", _
            "*google-analytics*", _
            "*/collect*", "*/beacon*", "*pagead*")
        bidi.ExecuteEnableResourceBlocking blockList
        
        ' ==========================================
        ' Noise Filter: Ignore Google's telemetry
        ' ==========================================
        ' Google sends constant background telemetry.
        ' Without filtering, SPA Idle consensus never stabilizes.
        bidi.AddIdleIgnoreNetworkPattern "/log?"
        bidi.AddIdleIgnoreNetworkPattern "play.google.com"
        bidi.AddIdleIgnoreNetworkPattern "ogs.google.com"
        bidi.AddIdleIgnoreNetworkPattern "/gen_204"
        bidi.AddIdleIgnoreNetworkPattern "gstatic.com"
        
        ' ==========================================
        ' Navigation
        ' ==========================================
        Dim url As String: url = "https://www.google.com/travel/flights"
        bidi.ExecuteNavigateAndGetStatus url
        
        ' ==========================================
        ' STEP 0: Set Ticket Type (Custom Dropdown)
        ' ==========================================
        ' Bypasses obfuscated class names by targeting W3C ARIA semantics.
        ' Automatically handles pop-up mutation spikes and framework re-renders.
        Dim ticketTypeTrigger As String
        ticketTypeTrigger = "(//div[@role='combobox' and @aria-haspopup='listbox'])[1]"
        
        ' Switch from "Round trip" to "One way"
        bidi.ExecuteSelectValueByXPath "(//div[@role='combobox'])[1]", "One way"
                
        ' ==========================================
        ' STEP 1: Set Departure City - "Sapporo"
        ' ==========================================
        ' [last()] ensures we resolve the final (real) input when Wiz spawns
        ' multiple copies.  Phase 0's stability check provides a second layer
        ' of protection inside the JS execution.
        
        Dim depXPath As String
        depXPath = "(//input[contains(@aria-label, 'Where from')])[last()]"
        
        ' Clear pre-populated value and type new city
        bidi.ExecuteInputValueByXPath depXPath, "Sapporo"
        
        ' Targets the visible listbox only (role=listbox + aria-hidden=false) generated by the SPA.
        ' Selects the option element containing "Sapporo" after the suggestion list is fully rendered.
        Dim depSuggestXPath As String
        depSuggestXPath = "//*[@role='listbox' and not(@aria-hidden='true')]//li[@role='option' and contains(@aria-label, 'Sapporo')][1]"
        bidi.ExecuteClickByXPath depSuggestXPath, , , , 300
        
        ' ==========================================
        ' STEP 2: Set Destination City - "Paris"
        ' ==========================================
        ' After selecting departure, Google Flights auto-activates
        ' the destination combobox.
        
        Dim destXPath As String
        destXPath = "(//input[contains(@aria-label, 'Where to')])[last()]"
        
        ' Type destination using the same Phase 0-2 logic
        bidi.ExecuteInputValueByXPath destXPath, "Paris", 5000
        
        ' Click matching suggestion
        Dim destSuggestXPath As String
        destSuggestXPath = "//*[@role='listbox' and not(@aria-hidden='true')]//li[@role='option' and contains(@aria-label, 'Paris')][1]"
        bidi.ExecuteClickByXPath destSuggestXPath, 5000
        
        ' ==========================================
        ' STEP 3: Select Dates
        ' ==========================================
        ' Calendar opens automatically after destination is selected.
        ' Strategy:
        '   - Pick first available date as departure
        '   - Pick 7th available date as return
        
        Dim depDateFieldXPath As String
        depDateFieldXPath = "//input[@aria-label='Departure']"
        bidi.ExecuteClickByXPath depDateFieldXPath, 5000
        
        Dim depDateXPath As String
        depDateXPath = "(//div[@role='gridcell' and @aria-hidden='false'])[1]//div[@role='button']"
        bidi.ExecuteClickByXPath depDateXPath, 5000
        
        Dim retDateXPath As String
        retDateXPath = "(//div[@role='gridcell' and @aria-hidden='false'])[8]//div[@role='button']"
        bidi.ExecuteClickByXPath retDateXPath, 5000
        
        Dim doneXPath As String
        doneXPath = "//button[contains(., 'Done')]"
        bidi.ExecuteClickByXPath doneXPath, 10000
        
        ' ==========================================
        ' STEP 4: Click Search Button
        ' ==========================================
        Dim searchXPath As String
        searchXPath = "//button[@aria-label='Search']"
        bidi.ExecuteClickByXPath searchXPath, 10000
        
        ' ==========================================
        ' Cleanup
        ' ==========================================
        bidi.Shutdown: Set bidi = Nothing
        .CloseBrowser: .Shutdown
        
        ' ==========================================
        ' Completion
        ' ==========================================
        MsgBox "Google Flights Test Completed"
        
    End With
End Sub

' Recorder
Sub Main09()
  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
    .StartEdge
    
    ' Browser startup settings
    Dim caps As WebCapabilities: Set caps = .CreateCapabilities
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
    bidi.ExecuteNavigateAndGetStatus url
    
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
    
    ' Cleanup
    bidi.Shutdown: Set bidi = Nothing
    .CloseBrowser: .Shutdown
    
End With
End Sub
