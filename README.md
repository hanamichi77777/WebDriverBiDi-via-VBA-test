【Operation explanation】
①When you press the above button, it will connect to the existing WebDriver in the download folder
 and perform WebSocket communication.

②Launching the SeleniumVBA homepage using WebDriverBiDi

③Detect the event "browsingContext.domContentLoaded", 
display a message box, and close the browser.

④Open the Immediate Window to see the details of various events.
※Anything not beginning with an ID is received due to event detection.

【Reference materials】
〇WebSocket communication using VBA					
ZeroInstall BrowserDriver for VBA(@kabkabkab)					
https://qiita.com/kabkabkab/items/d187fd1622fede038cce					
					
〇BiDi implementation related					
Puppeteerのコードを見つつ、BiDiを手でさわってみる(@Yusuke Iwaki)					
https://zenn.dev/yusukeiwaki/scraps/00fd022cb857e0					
					
〇regular expression					
vba-regex fork (@sihlfall @GCuser99)					
https://github.com/GCuser99/SeleniumVBA/discussions/158					
![image](https://github.com/user-attachments/assets/ac875cd7-5919-4487-ad70-2e84346a3c6f)

