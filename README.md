【6/26/2025】
I have uploaded an experimental file that shows WebDriverBiDi working with SeleniumVBA6.5 test1.

This VBAprogram has developed based on "ZeroInstall BrowserDriver for VBA" (@kabkabkab) and changed the connection from CDP to WebDriverBidi with a WebSocket connection. I created this in hopes of making it possible to detect Events using SeleniumVBA6.5 test1(@GCuser99).

[Modifications]
・Cross-browser compatible (Edge, Chrome)
・Added standard module "A_00_BIDI".Can be enabled or disabled in BiDI
・Class module WebDriver modification
・Detected by specifying the event name as the second argument of "SendAndReceive".
・Move common processing for BiDI related to a class module that starts with "BiDi"
[Operation explanation]
1. Press the above button and you will see it in the download folder.
Connect to an existing WebDriver and perform WebSocket communication.
2. Enable log notifications limited to browsingContext and script
(Shows in the Immediate window)
3. Note Novel Categories for Dynamic Page Loading with WebDriverBIDI
Transition to the URL.
https://note.com/topic/novel
4. Wait for browsingContext.load event to occur
5. Obtain the number of elements for each article using FindElements using normal operation not BiDi
(In this case, the event will not be detected.)
6. After 4 seconds, the same process as 4 is performed, and after completion, the process is suspended for a while to confirm.
However, the number of elements for 4 and 5 is different, so the waiting is insufficient.why?
7. Close the browser and exit WebDriver
8. You can check the contents of various events in the Immediate window
*Among these events that do not start with Receive with ID will be an event.
[Reference materials]
1. WebSocket communication related with VBA
・ZeroInstall BrowserDriver for VBA (@kabkabkab)
https://qiita.com/kabkabkab/items/d187fd1622fede038cce
2. BIDI Implementation Notes
・Puppeteerのコードを見つつ、BiDiを手でさわってみる(@Yusuke Iwaki)
https://zenn.dev/yusukeiwaki/scraps/00fd022cb857e0
![image](https://github.com/user-attachments/assets/1cd2413c-645e-4108-9b08-0d3fa3854c6d)


