**【9/27/2025 Updated】**

I have uploaded an experimental file that shows WebDriverBiDi working with SeleniumVBA6.6 (@GCuser99).
To prevent the file from being deleted due to a false positive by Defender, a password [pass] is set for the file.

This VBAprogram has developed based on "ZeroInstall BrowserDriver for VBA" (@kabkabkab) and changed the connection from CDP to WebDriverBidi with a WebSocket connection. I created this in hopes of making it possible to detect Events using SeleniumVBA(@GCuser99).

**[Modifications]**

・Cross-browser compatible (Edge, Chrome)
・Added standard module "A_00_BIDI".Can be enabled or disabled in BiDI
・Class module WebDriver modification
・Detected by specifying the event name as the second argument of "SendAndReceive".
・Move common processing for BiDI related to a class module that starts with "BiDi"

**[Operation explanation]**

① Press the above button to connect to an existing WebDriver in the download folder and perform WebSocket communication.

② Enable log notifications limited to network.beforeRequestSent and network.responseCompleted (displayed in the immediate window)

**③ Move to the URL of the note novel category where pages are dynamically loaded using WebDriverBIDI.**

**④ Wait until the network.responseCompleted event occurs.
However, the waiting will continue until all requests in network.beforeRequestSent have completed."**

**⑤ Even if all requests are processed, a new network.beforeRequestSent may occur, so after waiting for the specified number of seconds, it continues if an event occurs, and ends if it does not occur.**

⑥ Obtain the number of elements for each article using FindElements using normal operation (in this case, event will not be detected)

⑦ After 4 seconds, the same process as 6 is performed, and the process is temporarily suspended for confirmation after completion.
As a result, the number of elements ⑥ and ⑦ coincides with the same number, so it can be confirmed that they are waiting.
If no match is made, increase the number of seconds specified in ⑤."


**[My article (Only Japanese)]**

[https://qiita.com/sele_chan/items/b6bdc321cf440fe5ac83](https://qiita.com/sele_chan/items/b6bdc321cf440fe5ac83)


**[Reference materials]**
1. WebSocket communication related with VBA
・ZeroInstall BrowserDriver for VBA (@kabkabkab)
https://qiita.com/kabkabkab/items/d187fd1622fede038cce
2. BIDI Implementation Notes
・Puppeteerのコードを見つつ、BiDiを手でさわってみる(@Yusuke Iwaki)
https://zenn.dev/yusukeiwaki/scraps/00fd022cb857e0
