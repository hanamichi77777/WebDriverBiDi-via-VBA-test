**【11/23/2025 Updated】**

I have uploaded an experimental file that shows WebDriverBiDi working with SeleniumVBA7.0 (@GCuser99).
To prevent the file from being deleted due to a false positive by Defender, a password "123" is set for the file.

This VBAprogram has developed based on "ZeroInstall BrowserDriver for VBA" (@kabkabkab) and changed the connection from CDP to WebDriverBidi with a WebSocket connection. I created this in hopes of making it possible to detect Events using SeleniumVBA(@GCuser99).

**[Operation explanation]**

Left side of button

〇 Loading Chrome extensions that are only supported by WebDriver BiDi 

*ExtensionPath needs to be changed as appropriate.

〇 On a web page where an asynchronous event occurs when a select box is selected, wait until the event completes.

Button center

〇 Wait until the dynamic page has finished loading

〇 Get status code even when logged in

Right side of button

〇  On a web page where an asynchronous event occurs when characters in the TextBox is enterd, wait until the event completes.

Other（inside the module）

〇 Waiting during login operations

〇 Network blocking processing

**[My article (Only Japanese)]**

[https://qiita.com/sele_chan/items/b6bdc321cf440fe5ac83](https://qiita.com/sele_chan/items/b6bdc321cf440fe5ac83)

[https://note.com/sele_chan/n/n2b21c7c26ef8](https://note.com/sele_chan/n/n2b21c7c26ef8)

**[Reference materials]**
1. WebSocket communication related with VBA
・ZeroInstall BrowserDriver for VBA (@kabkabkab)
https://qiita.com/kabkabkab/items/d187fd1622fede038cce
2. BIDI Implementation Notes
・Puppeteerのコードを見つつ、BiDiを手でさわってみる(@Yusuke Iwaki)
https://zenn.dev/yusukeiwaki/scraps/00fd022cb857e0
