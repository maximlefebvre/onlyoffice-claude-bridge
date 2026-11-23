// Claude Bridge menu bar app (macOS 13+).
//
// Starts and stops the shared Claude Bridge server (`server.mjs --http`, bundled
// in the app's Resources) and shows its state in the menu bar: connected
// OnlyOffice documents, Claude sessions, and whether the plugin port is free.
// Launching the app starts the server; quitting the app stops it.
//
// Build: macos-app/build.sh (no Xcode project needed).

import AppKit
import ServiceManagement

private let mcpPort = 47616
private let configDir = FileManager.default.homeDirectoryForCurrentUser
    .appendingPathComponent(".config/onlyoffice-claude-bridge")
private let logURL = FileManager.default.homeDirectoryForCurrentUser
    .appendingPathComponent("Library/Logs/onlyoffice-claude-bridge.log")

struct ServerStatus: Decodable {
    struct Document: Decodable { let id: Int?; let title: String; let editor: String }
    struct Holder: Decodable { let pid: Int?; let bridge: Bool?; let version: String? }
    let version: String
    let pid: Int32
    let sessions: Int
    let port: Int?
    let listening: Bool?
    let portHolder: Holder?
    let documents: [Document]?
}

@MainActor
final class AppDelegate: NSObject, NSApplicationDelegate, NSMenuDelegate {
    private let statusItem = NSStatusBar.system.statusItem(withLength: NSStatusItem.variableLength)
    private let menu = NSMenu()
    private var process: Process?
    private var status: ServerStatus?
    private var lastExit: String?
    private var stopping = false
    private var timer: Timer?

    func applicationDidFinishLaunching(_ notification: Notification) {
        menu.delegate = self
        statusItem.menu = menu
        updateIcon()
        // Check first whether a shared server already runs, then start ours if not.
        Task {
            await refresh()
            startServer()
        }
        timer = Timer.scheduledTimer(withTimeInterval: 3, repeats: true) { [weak self] _ in
            Task { @MainActor in await self?.refresh() }
        }
    }

    func applicationWillTerminate(_ notification: Notification) {
        stopServer(wait: true)
    }

    // MARK: Server process

    private var ownsServer: Bool { process?.isRunning == true }

    private func findNode() -> String? {
        let candidates = ["/opt/homebrew/bin/node", "/usr/local/bin/node", "/usr/bin/node"]
        if let found = candidates.first(where: { FileManager.default.isExecutableFile(atPath: $0) }) { return found }
        // Fall back to the login shell's PATH (nvm, asdf, volta…).
        let shell = Process()
        shell.executableURL = URL(fileURLWithPath: "/bin/zsh")
        shell.arguments = ["-lc", "command -v node"]
        let pipe = Pipe()
        shell.standardOutput = pipe
        shell.standardError = FileHandle.nullDevice
        do { try shell.run() } catch { return nil }
        shell.waitUntilExit()
        let out = String(data: pipe.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8)?
            .trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
        return out.isEmpty ? nil : out
    }

    private func logHandle() -> FileHandle? {
        let fm = FileManager.default
        try? fm.createDirectory(at: logURL.deletingLastPathComponent(), withIntermediateDirectories: true)
        if !fm.fileExists(atPath: logURL.path) { fm.createFile(atPath: logURL.path, contents: nil) }
        let handle = try? FileHandle(forWritingTo: logURL)
        handle?.seekToEndOfFile()
        return handle
    }

    func startServer() {
        guard !ownsServer else { return }
        if status != nil { lastExit = nil; return } // already served by another process
        guard let node = findNode() else {
            lastExit = "Node.js not found (install Node.js 20+)"
            return
        }
        guard let server = Bundle.main.path(forResource: "server", ofType: "mjs") else {
            lastExit = "server.mjs missing from the app bundle"
            return
        }
        let p = Process()
        p.executableURL = URL(fileURLWithPath: node)
        p.arguments = [server, "--http"]
        let log = logHandle()
        p.standardOutput = log
        p.standardError = log
        p.terminationHandler = { [weak self] proc in
            let code = proc.terminationStatus
            Task { @MainActor in
                guard let self else { return }
                if !self.stopping { self.lastExit = "Server stopped unexpectedly (exit code \(code)), see the log" }
                self.stopping = false
                self.process = nil
                await self.refresh()
            }
        }
        do {
            try p.run()
            process = p
            lastExit = nil
        } catch {
            lastExit = "Could not start the server: \(error.localizedDescription)"
        }
        Task {
            try? await Task.sleep(nanoseconds: 1_500_000_000)
            await refresh()
        }
    }

    func stopServer(wait: Bool = false) {
        if let p = process, p.isRunning {
            stopping = true
            p.terminate() // SIGTERM: the server closes its sockets and exits
            if wait { p.waitUntilExit() }
        } else if let pid = status?.pid {
            // Started outside this app (e.g. by the launchd service, which would restart it).
            stopping = true
            kill(pid, SIGTERM)
        }
    }

    // MARK: Status

    private func readToken() -> String? {
        guard let data = try? Data(contentsOf: configDir.appendingPathComponent("mcp-token")) else { return nil }
        return String(data: data, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    func refresh() async {
        var next: ServerStatus?
        if let token = readToken(), let url = URL(string: "http://127.0.0.1:\(mcpPort)/status") {
            var req = URLRequest(url: url, timeoutInterval: 1.5)
            req.setValue("Bearer \(token)", forHTTPHeaderField: "Authorization")
            if let (data, resp) = try? await URLSession.shared.data(for: req),
               (resp as? HTTPURLResponse)?.statusCode == 200 {
                next = try? JSONDecoder().decode(ServerStatus.self, from: data)
            }
        }
        status = next
        updateIcon()
    }

    private func updateIcon() {
        guard let button = statusItem.button else { return }
        let running = status != nil
        let docs = status?.documents?.count ?? 0
        let blocked = status?.listening == false
        let symbol = !running ? "link.circle" : (blocked ? "exclamationmark.circle" : (docs > 0 ? "link.circle.fill" : "link.circle"))
        let image = NSImage(systemSymbolName: symbol, accessibilityDescription: "Claude Bridge")
        image?.isTemplate = true
        button.image = image
        button.appearsDisabled = !running
        button.toolTip = running ? "Claude Bridge: \(docs) document(s) connected" : "Claude Bridge: stopped"
    }

    // MARK: Menu

    func menuNeedsUpdate(_ menu: NSMenu) {
        menu.removeAllItems()
        let running = status != nil

        func info(_ title: String, symbol: String? = nil) {
            let item = NSMenuItem(title: title, action: nil, keyEquivalent: "")
            item.isEnabled = false
            if let symbol { item.image = NSImage(systemSymbolName: symbol, accessibilityDescription: nil) }
            menu.addItem(item)
        }

        if let s = status {
            info("Claude Bridge \(s.version): running" + (ownsServer ? "" : " (started outside this app)"))
            if s.listening == false {
                let holder = s.portHolder?.pid.map { "pid \($0)" } ?? "another process"
                info("Port \(s.port ?? 47615) held by \(holder): waiting for it", symbol: "exclamationmark.triangle")
            }
            info("\(s.sessions) Claude session(s) connected")
            let docs = s.documents ?? []
            if docs.isEmpty {
                info("No OnlyOffice document connected")
            } else {
                info("Documents:")
                for d in docs {
                    let symbol = ["word": "doc.text", "cell": "tablecells", "slide": "rectangle.on.rectangle"][d.editor] ?? "doc"
                    info("   \(d.title)", symbol: symbol)
                }
            }
        } else {
            info("Claude Bridge: stopped")
        }
        if let lastExit { info(lastExit, symbol: "exclamationmark.triangle") }

        menu.addItem(.separator())
        if running {
            menu.addItem(withTitle: "Stop Server", action: #selector(stopClicked), keyEquivalent: "").target = self
        } else {
            menu.addItem(withTitle: "Start Server", action: #selector(startClicked), keyEquivalent: "").target = self
        }
        let login = NSMenuItem(title: "Open at Login", action: #selector(toggleLogin), keyEquivalent: "")
        login.target = self
        login.state = SMAppService.mainApp.status == .enabled ? .on : .off
        menu.addItem(login)
        menu.addItem(withTitle: "Show Log", action: #selector(showLog), keyEquivalent: "").target = self
        menu.addItem(.separator())
        menu.addItem(withTitle: "Quit (stops the server)", action: #selector(quit), keyEquivalent: "q").target = self
    }

    @objc private func startClicked() { lastExit = nil; startServer() }
    @objc private func stopClicked() { stopServer(); Task { try? await Task.sleep(nanoseconds: 800_000_000); await refresh() } }
    @objc private func showLog() { NSWorkspace.shared.open(logURL) }
    @objc private func quit() { NSApp.terminate(nil) }

    @objc private func toggleLogin() {
        do {
            if SMAppService.mainApp.status == .enabled {
                try SMAppService.mainApp.unregister()
            } else {
                try SMAppService.mainApp.register()
            }
        } catch {
            lastExit = "Open at Login: \(error.localizedDescription)"
        }
    }
}

MainActor.assumeIsolated {
    let app = NSApplication.shared
    let delegate = AppDelegate()
    app.delegate = delegate
    app.setActivationPolicy(.accessory) // menu bar only, no Dock icon
    app.run()
}
