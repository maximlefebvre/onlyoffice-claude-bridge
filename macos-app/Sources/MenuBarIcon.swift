import AppKit

// Monochrome menu bar glyph (template image, so macOS tints it for light/dark
// menu bars): stacked document layers for the office suite, and a small
// radiating spark for Claude. Original drawing, not either brand's logo.
//   filled:  a document is connected (top layer filled)
//   warning: the plugin port is taken (the spark becomes "!")
enum MenuBarIcon {
    static func make(filled: Bool, warning: Bool) -> NSImage {
        let image = NSImage(size: NSSize(width: 18, height: 18), flipped: false) { _ in
            NSColor.black.setStroke()
            NSColor.black.setFill()

            // Layers: a rhombus on top of two chevrons.
            let cx: CGFloat = 7.5, hw: CGFloat = 6, hh: CGFloat = 2.6
            let top = NSBezierPath()
            top.move(to: NSPoint(x: cx - hw, y: 11.0))
            top.line(to: NSPoint(x: cx, y: 11.0 + hh))
            top.line(to: NSPoint(x: cx + hw, y: 11.0))
            top.line(to: NSPoint(x: cx, y: 11.0 - hh))
            top.close()
            top.lineWidth = 1.3
            top.lineJoinStyle = .round
            if filled { top.fill() }
            top.stroke()
            for y: CGFloat in [7.9, 4.8] {
                let chevron = NSBezierPath()
                chevron.move(to: NSPoint(x: cx - hw, y: y))
                chevron.line(to: NSPoint(x: cx, y: y - hh))
                chevron.line(to: NSPoint(x: cx + hw, y: y))
                chevron.lineWidth = 1.3
                chevron.lineJoinStyle = .round
                chevron.lineCapStyle = .round
                chevron.stroke()
            }

            let c = NSPoint(x: 15.3, y: 14.6)
            if warning {
                let bar = NSBezierPath()
                bar.move(to: NSPoint(x: c.x, y: 17.0))
                bar.line(to: NSPoint(x: c.x, y: 13.9))
                bar.lineWidth = 1.6
                bar.lineCapStyle = .round
                bar.stroke()
                NSBezierPath(ovalIn: NSRect(x: c.x - 0.95, y: 11.6, width: 1.9, height: 1.9)).fill()
            } else {
                // Spark: 8 rays, long and short alternating.
                for i in 0..<8 {
                    let a = CGFloat(i) * .pi / 4 + .pi / 8
                    let r: CGFloat = i % 2 == 0 ? 2.5 : 1.6
                    let ray = NSBezierPath()
                    ray.move(to: NSPoint(x: c.x + 0.5 * cos(a), y: c.y + 0.5 * sin(a)))
                    ray.line(to: NSPoint(x: c.x + r * cos(a), y: c.y + r * sin(a)))
                    ray.lineWidth = 1.1
                    ray.lineCapStyle = .round
                    ray.stroke()
                }
            }
            return true
        }
        image.isTemplate = true
        image.accessibilityDescription = "Claude Bridge"
        return image
    }
}
