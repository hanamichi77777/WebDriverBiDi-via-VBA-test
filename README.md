# WebDriver BiDi Automation for SeleniumVBA

### 【1/31/2026 Updated】 Improved robustness for SPA environments.
I have uploaded an experimental file that shows WebDriver BiDi working with **[SeleniumVBA ver.7.2](https://github.com/GCuser99/SeleniumVBA)** (@GCuser99). To prevent the file from being deleted due to a false positive by Windows Defender, a password **"123"** is set for the file.

This VBA program was developed based on **"ZeroInstall BrowserDriver for VBA"** (@kabkabkab) and changed the connection from CDP to WebDriver BiDi with a WebSocket communication. I created this in hopes of making it possible to detect Events using SeleniumVBA (@GCuser99).

To overcome the flakiness arising from DOM updates and async requests in modern SPAs like React and Vue.js, I’m challenging the boundaries of what’s possible with VBA.

---

## [Supported Browsers]
* **Edge / Chrome**
* *※ Firefox is not supported due to functional limitations.*

---

## 📂 Procedure Overview (Sample Module: `A_01BiDi_Sample`)

#### 1. Main01: Enhanced Select Box（Use test.html） & Extension（Use Google Translate extension） Injection
This procedure focuses on handling elements that trigger complex JavaScript state changes. 
* **Dynamic Extension** Injection: Utilizes the WebDriver BiDi webExtension.install command to load extensions directly into the browser session from a local path. This enables the runtime "bypass injection" of extensions—such as ad-blockers or custom tools—without cluttering the system registry or permanent configuration files.
* **Smart Selection:** Utilizes `ExecuteSelectValueByXPath`. Unlike standard Selenium, this command can be configured to wait for the browser's "Idle" state immediately after the selection, ensuring that any subsequent calendar or UI updates are fully rendered before the script proceeds.

#### 2. Main02: SPA Auto-Clicking & Dynamic Synchronization
Designed for high-activity SPA environments like *note.com*, this procedure ensures interaction with elements that are dynamically added to the DOM.
* **Full-Stack Idleness Monitoring:** Once the main navigation (`browsingContext.navigate`) completes, the script injects the `window.__vbaIdleProbe`. 
* **Real-time Traffic Tracking:** The logs demonstrate the probe tracking `inflightXhrCount` and `inflightFetchCount`. The VBA code effectively "waits" for these counts to spike and then return to zero, combined with a stable `lastMutationTs`, ensuring the dynamic article feed has finished streaming before proceeding.

#### 3. Main03: Performance Optimization via CDP-over-BiDi Bridge
This procedure demonstrates how to make automation up to 5x faster by controlling the network layer using a hybrid protocol approach.
* **Hybrid Protocol Bridge:** Utilizes `goog:cdp.sendCommand` to access low-level domains. It enables `Network.setBlockedURLs` to filter out images and ad scripts before the navigation command is sent.
* **Post-Navigation Idleness Probe:** Injects `window.__vbaIdleProbe` to ensure the environment is quiescent before entering search data.

#### 4. Main04: Event-Driven URL Monitoring
Bypasses the "flaky" nature of login redirections.
* **Event vs. Polling:** Uses `ExecuteIsUrlContains` to hook into the browser's internal navigation events. The VBA script wakes up instantly the millisecond the URL matches the target, ensuring no time is wasted.

#### 5. Main06: Iframe Context Piercing & Hierarchical Mapping
Solves the "nested frame" problem found in legacy portals.
* **Hierarchical Tree Mapping:** Executes `browsingContext.getTree` to map the entire context hierarchy, identifying deeply nested sub-frames.
* **Direct Context Targeting:** Instead of using context switching, the script retrieves the unique `context` ID for the specific frame and passes it directly to interaction commands.

#### 6. Main07: SPA Idleness Detection & Shadow DOM Traversal
Targeting heavy JavaScript platforms (e.g., ServiceNow), this procedure implements a sophisticated **"BiDi Probe"** system for SPA synchronization.
* **Advanced SPA Synchronization:** Uses a triple-layer check (XHR, Fetch, and Mutation timestamps) to ensure the framework is fully hydrated.
* **Robust Context Recovery:** Automatically detects "Context Lost" errors during SPA redirects and waits (500ms) for context recovery.

#### 7. Main09: Discovery Log & Diagnostic Recording
A specialized tool for reverse-engineering and debugging.
* **Event Stream:** Uses `StartDiscoveryLog` to capture a raw feed of every browser event, including network requests, console logs, and DOM changes.
* **Analysis:** Records activity for 20 seconds and saves it to `discovery_log.txt`.

---

### 🔗 External Links
* [Qiita - Article by sele_chan](https://qiita.com/sele_chan/items/6475a1f7ae8a21435d6c)
* [note - Article by sele_chan](https://note.com/sele_chan/n/n2b21c7c26ef8)

### [Reference Materials]
* **WebSocket communication related with VBA:** [ZeroInstall BrowserDriver for VBA (@kabkabkab)](https://qiita.com/kabkabkab/items/d187fd1622fede038cce)
* **BiDi Implementation Notes:** [Puppeteerのコードを見つつ、BiDiを手でさわってみる (@Yusuke Iwaki)](https://zenn.dev/yusukeiwaki/scraps/00fd022cb857e0)
