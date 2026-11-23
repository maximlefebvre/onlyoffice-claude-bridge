// Renders the app icon (AppIcon.iconset, to be packed by `iconutil`) from the
// menu bar glyph: the white glyph on a rounded terracotta square.
// Usage: make-icon <output .iconset folder>

import AppKit

@main
struct MakeIcon {
    static func render(size: Int) -> Data {
        let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: size, pixelsHigh: size,
                                   bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true, isPlanar: false,
                                   colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
        NSGraphicsContext.saveGraphicsState()
        NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
        let s = CGFloat(size) / 1024

        // Rounded square with the macOS icon grid margins (824 pt body in 1024).
        let body = NSRect(x: 100 * s, y: 100 * s, width: 824 * s, height: 824 * s)
        let shape = NSBezierPath(roundedRect: body, xRadius: 185 * s, yRadius: 185 * s)
        NSGraphicsContext.saveGraphicsState()
        let shadow = NSShadow()
        shadow.shadowColor = NSColor.black.withAlphaComponent(0.25)
        shadow.shadowOffset = NSSize(width: 0, height: -10 * s)
        shadow.shadowBlurRadius = 24 * s
        shadow.set()
        NSColor(red: 0.85, green: 0.47, blue: 0.34, alpha: 1).setFill()
        shape.fill()
        NSGraphicsContext.restoreGraphicsState()
        NSGradient(starting: NSColor(red: 0.89, green: 0.55, blue: 0.42, alpha: 1),
                   ending: NSColor(red: 0.78, green: 0.39, blue: 0.26, alpha: 1))!
            .draw(in: shape, angle: -90)

        // White glyph, tinted from the template image.
        let glyph = MenuBarIcon.make(filled: true, warning: false)
        let g = NSRect(x: 222 * s, y: 222 * s, width: 580 * s, height: 580 * s)
        let white = NSImage(size: glyph.size, flipped: false) { rect in
            glyph.draw(in: rect)
            NSColor.white.set()
            rect.fill(using: .sourceAtop)
            return true
        }
        white.draw(in: g)

        NSGraphicsContext.restoreGraphicsState()
        return rep.representation(using: .png, properties: [:])!
    }

    static func main() throws {
        guard CommandLine.arguments.count == 2 else {
            FileHandle.standardError.write("usage: make-icon <output .iconset folder>\n".data(using: .utf8)!)
            exit(1)
        }
        let out = URL(fileURLWithPath: CommandLine.arguments[1])
        try FileManager.default.createDirectory(at: out, withIntermediateDirectories: true)
        for base in [16, 32, 128, 256, 512] {
            try render(size: base).write(to: out.appendingPathComponent("icon_\(base)x\(base).png"))
            try render(size: base * 2).write(to: out.appendingPathComponent("icon_\(base)x\(base)@2x.png"))
        }
    }
}
