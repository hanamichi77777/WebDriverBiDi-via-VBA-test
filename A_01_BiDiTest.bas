Attribute VB_Name = "A_01_BiDiTest"
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
    ' Enable Chrome extension
    Dim extensionPath As String
    extensionPath = Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\Default\Extensions\aapbdbdomjkkjkaonfhkkikfgjllcleb\2.0.16_0"
    bidi.ExecuteWebExtensionInstall (extensionPath)
      
    ' Navigate to page
    Dim url As String: url = "http://keylopment.com/faq/2357"
    bidi.ExecuteBrowsingContextNavigate url
      
' --- 2. Search for XPath element and execute click ---
    ' XPath of the option element displaying March 2026
    Dim strClickXpath As String: strClickXpath = "//select[@name='calselect']"
      
    ' Search for XPath element and execute click (Argument is the Value of the Option tag)
    bidi.ExecuteSelectValueByXPath strClickXpath, "2026年03月", True
      
' --- 3. Verification and Termination ---
    Dim str As String
    ' Check if the calendar switched as expected
    str = .FindElement(By.xpath, "//h3[@class='title-level03']").GetText
      
    Dim msgText As String, msgCaption As String
    ' Note: Keeping "2026年03月" in Japanese as it compares against site content
    If str = "2026年03月" Then
        msgText = "Successfully waited until the calendar switched."
        msgCaption = "Verification Complete"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    Else
        msgText = "The calendar has not switched. Retrieved value: " & str
        msgCaption = "Verification Failed"
        MESSAGEbox 0, msgText, msgCaption, MB_OK Or MB_ForeFront
    End If
      
    ' End
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
             
   bidi.ExecuteRegisterAutoClickerByXPath ("//input[@type='submit'][contains(@value,'表示')]")
    
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
    
' [All processing is written in BiDi] Here, save Yahoo! JAPAN Transit Guide
'  ' Setting to block images and ads
  Dim blockList As Variant
'  ' Example patterns to block common image formats and ad services
  blockList = Array( _
    "*.png", "*.jpg", "*.jpeg", "*.gif", "*.svg", _
    "*ad_service*", "*analytics*", _
    "*doubleclick*", "*googlesyndication*", _
    "*criteo*", _
    "*yads*", _
    "*yjt*" _
  ) '

  ' Method name changed to match CDP implementation
  bidi.ExecuteEnableResourceBlocking blockList

  ' Page transition
  Dim url As String: url = "https://transit.yahoo.co.jp/"
  bidi.ExecuteNavigateAndGetStatus url

  ' Departure: Tokyo
  bidi.ExecuteInputValueByXPath "//input[@name='from']", "東京"
  ' Arrival: Shinjuku
  bidi.ExecuteInputValueByXPath "//input[@name='to']", "新宿"
  ' Date/Time: Last train (using Japanese text in XPath)
  bidi.ExecuteClickByXPath "//label[text()='終電']"
  ' Display order: Lowest price
   bidi.ExecuteSelectValueByXPath "//select[@id='s']", "料金が安い順", True
  ' Click search button
  bidi.ExecuteClickByXPath "//input[@id='searchModuleSubmit']"
   
   
  ' End
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
    bidi.ExecuteBrowsingContextNavigate loginUrl, "none"
      
    Dim isLoginSuccess As Boolean
    isLoginSuccess = bidi.ExecuteBiDiWaitUntilUrlContains("https://hotel-example-site.takeyaqa.dev/ja/mypage.html")
      
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
    .CloseBrowser
    .Shutdown
      
  End With
End Sub

' [Network Blocking Example - Updated for CDP Tunneling]
Public Sub Main05()

  Dim driver As WebDriver: Set driver = New WebDriver
  With driver
    
    .StartChrome
    Dim caps As SeleniumVBA.WebCapabilities: Set caps = .CreateCapabilities
    caps.EnableBiDiMode
    
    .OpenBrowser caps
   Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo .GetWebSocketUrl
        
    ' -----------------------------------------------------------
    ' 1. Setting to block images and ads via WebDriver BiDi (CDP Tunneling)
    ' -----------------------------------------------------------
    Dim blockList As Variant
    ' Example patterns to block common image formats and ad services
    blockList = Array( _
    "*.png", "*.jpg", "*.jpeg", "*.gif", "*.svg", _
    "*doubleclick*", _
    "*googleads*", _
    "*googlesyndication*", _
    "*yads.c.yimg.jp*", _
    "*yads.yahoo.co.jp*", _
    "*ad_service*", _
    "*analytics*" _
)
    
    ' [UPDATED] Use the new CDP-based method
    bidi.ExecuteEnableResourceBlocking blockList
    
    ' -----------------------------------------------------------
    ' 2. Page Navigation (Display speed is faster without images)
    ' -----------------------------------------------------------
    ' Navigate to Yahoo! Japan
    bidi.ExecuteNavigateAndGetStatus "https://www.yahoo.co.jp/"
    
    ' Use API MESSAGEbox (Topmost)
    Dim msgText As String
    msgText = "Images are currently blocked via CDP Tunneling." & vbCrLf & "Press OK to continue (Blocking will be cleared)."
    MESSAGEbox 0, msgText, "Blocking Active", MB_OK Or MB_ForeFront
    
    ' -----------------------------------------------------------
    ' 3. Clear blocking and reload
    ' -----------------------------------------------------------
    ' [UPDATED] Use the new CDP-based disable method
    bidi.ExecuteDisableResourceBlocking
    
    .Refresh ' Reload using standard Selenium command
    
    ' Use API MESSAGEbox (Topmost)
    msgText = "Are images displayed now?"
    MESSAGEbox 0, msgText, "Verification", MB_OK Or MB_ForeFront
    
    ' End
    .CloseBrowser
    .Shutdown
    
  End With
End Sub

' Shadow DOM Interaction (Click, Input, GetText)
' Input: selectorsArray (Array of Strings) -> e.g. Array("user-card", "paper-button")
Public Sub Main07()

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
        
  ' Example: Handling the "truste" consent button
  bidi.ExecuteRegisterAutoClickerByXPath "//button[@id='truste-consent-button']"
        
  Dim targetUrl As String: targetUrl = "https://developer.servicenow.com/"
  .NavigateTo targetUrl

  '/html/body/dps-app//div/header/dps-navigation-header//sn-cx-navigation//header/nav[1]/div[2]/ul/li[3]/button
  Dim path(3) As String
  path(0) = "dps-app"
  path(1) = "dps-navigation-header"
  path(2) = "sn-cx-navigation"
  path(3) = "#utility-sign-in button"  'XPath:[@id='utility-sign-in']//button
   
  ' Click
  bidi.ExecuteShadowClick path, True

  ' Wait
  If .IsPresent(By.xpath, "//input[@id='username']", 5000) = True Then
      ' Input
      bidi.ExecuteInputValueByXPath "//input[@id='username']", "aaa", False
  End If
  
  ' End
  .CloseBrowser
  .Shutdown
    
  End With
End Sub

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
    ' ==========================================================
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
    
    ' End
    .CloseBrowser
    .Shutdown
    
End With
End Sub

Public Sub test_BiDi()

    Dim driver As New WebDriver
    Dim elem As WebElement

    driver.StartEdge
    Dim caps As SeleniumVBA.WebCapabilities: Set caps = driver.CreateCapabilities
    caps.EnableBiDiMode
    driver.OpenBrowser caps
    Dim bidi As New BiDiCommandWrapper: bidi.ConnectTo driver.GetWebSocketUrl

    driver.NavigateTo "file:///" & Replace(driver.ResolvePath(".\") & "\test.html", "\", "/")
    
    bidi.ExecuteSelectValueByXPath "//select[@id='userSelector']", "1", False, , , 0
    
    Debug.Assert driver.FindElementByXPath("//input[@id='nameField']").GetProperty("value") = "Success: Response Received!"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Public Sub test_Classic()
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

Private Sub WaitForIdleStateAsync(oDriver As WebDriver, Optional idleTimeout As Long = 1500, Optional maxTimeToWait As Long = 30000)
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

