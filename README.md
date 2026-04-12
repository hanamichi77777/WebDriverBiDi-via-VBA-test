# WebDriver BiDi Automation for SeleniumVBA

### 【4/12/2026 Updated] 　Improved robustness for SPA environments.
I have uploaded an experimental file that shows WebDriver BiDi working with **[SeleniumVBA](https://github.com/GCuser99/SeleniumVBA)** (@GCuser99). To prevent the file from being deleted due to a false positive by Windows Defender, a password **"123"** is set for the file.

This VBA program was developed based on **"ZeroInstall BrowserDriver for VBA"** (@kabkabkab) and changed the connection from CDP to WebDriver BiDi with a WebSocket communication. I created this in hopes of making it possible to detect Events using SeleniumVBA (@GCuser99).

To overcome the flakiness arising from DOM updates and async requests in modern SPAs like React and Vue.js, I’m challenging the boundaries of what’s possible with VBA.Since we assume the concurrent use of both Classic and BiDi, the BiDi methods are kept to a minimum.

---

## [Supported Browsers]
* **Edge / Chrome**
* *Firefox is not supported due to functional limitations.*

---

## 📂 Procedure Overview (Sample Module: `A_01_BiDi_Sample`)

#### 1. Main01: Enhanced Select Box & Extension Injection & Recording
This procedure focuses on handling elements that trigger complex JavaScript state changes.

Dynamic Extension Injection: Utilizes the WebDriver BiDi webExtension.install command to load extensions directly into the browser session from a local path. This enables the runtime "bypass injection" of extensions without cluttering the system registry.

Smart Selection: Utilizes ExecuteSelectValueByXPath. This command can be configured to wait for the browser's "Idle" state immediately after selection, ensuring subsequent UI updates are fully rendered before proceeding.

#### 2. Main02: SPA Auto-Clicking & Dynamic Synchronization
Designed for high-activity SPA environments like note.com, this procedure ensures interaction with elements that are dynamically added to the DOM.

Full-Stack Idleness Monitoring: Once navigation completes, the script injects window.__vbaIdleProbe to monitor internal browser states.

Real-time Traffic Tracking: The probe tracks inflightXhrCount and inflightFetchCount. The VBA code waits for these counts to return to zero combined with a stable lastMutationTs, ensuring the dynamic feed has finished streaming.

#### 3. Main03: Performance Optimization via CDP-over-BiDi Bridge
This procedure demonstrates how to make automation up to 5x faster by controlling the network layer using a hybrid protocol approach.

Hybrid Protocol Bridge: Utilizes goog:cdp.sendCommand to access low-level domains, enabling Network.setBlockedURLs to filter out images and ad scripts before navigation.

Post-Navigation Idleness Probe: Injects window.__vbaIdleProbe to ensure the environment is quiescent before entering data, maximizing execution speed and reliability.

#### 4. Main04: Event-Driven URL Monitoring
Bypasses the "flaky" nature of login redirections by moving away from polling.

Event vs. Polling: Uses ExecuteIsUrlContains to hook into the browser's internal navigation events. The script wakes up instantly the millisecond the URL matches the target, ensuring no time is wasted waiting for fixed intervals.

#### 5. Main05: Asynchronous DOM Mutation & State Validation
Focuses on synchronizing with elements that are delayed or generated via AJAX, ensuring the script does not outpace the UI updates.

Smart Async Interaction: Utilizes ExecuteClickByXPath to interact with AJAX-driven content. The command internally monitors BiDi events to ensure the action is processed during a stable browser state.

Instant State Verification: Demonstrates how to validate dynamic DOM insertions (e.g., the "Done!" label) immediately after an action, eliminating the need for manual polling loops.

#### 6. Main06: Iframe Context Piercing & Hierarchical Mapping
Solves the "nested frame" problem found in legacy portals.

Hierarchical Tree Mapping: Executes browsingContext.getTree to map the entire context hierarchy, identifying deeply nested sub-frames.

Direct Context Targeting: Instead of using traditional context switching, the script retrieves a unique context ID for the specific frame and passes it directly to interaction commands.

#### 7. Main07: SPA Idleness Detection & Shadow DOM Traversal
Targeting heavy JavaScript platforms (e.g., ServiceNow), this procedure implements a sophisticated "BiDi Probe" system.

Advanced SPA Synchronization: Uses a triple-layer check (XHR, Fetch, and Mutation timestamps) to ensure the framework is fully hydrated.

Robust Context Recovery: Automatically detects "Context Lost" errors during SPA redirects and waits for context recovery before proceeding.

#### 8. Main09: Discovery Log & Diagnostic Recording
A specialized tool for reverse-engineering and debugging complex automation scenarios.

Event Stream: Uses StartDiscoveryLog to capture a raw feed of every browser event, including network requests, console logs, and DOM changes.

Analysis: Records activity for a specified duration (e.g., 20 seconds) and saves it to discovery_log.txt for post-mortem analysis.

---

### 🔗 External Links
* [Qiita - Article by sele_chan](https://qiita.com/sele_chan/items/6475a1f7ae8a21435d6c)
* [note - Article by sele_chan](https://note.com/sele_chan/n/n2b21c7c26ef8)

### [Reference Materials]
* **WebSocket communication related with VBA:** [ZeroInstall BrowserDriver for VBA (@kabkabkab)](https://qiita.com/kabkabkab/items/d187fd1622fede038cce)
* **BiDi Implementation Notes:** [Puppeteerのコードを見つつ、BiDiを手でさわってみる (@Yusuke Iwaki)](https://zenn.dev/yusukeiwaki/scraps/00fd022cb857e0)
