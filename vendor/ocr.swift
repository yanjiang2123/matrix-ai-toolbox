
import Foundation
import Vision
import AppKit

let args = CommandLine.arguments
guard args.count > 1, let img = NSImage(contentsOfFile: args[1]),
      let tiff = img.tiffRepresentation,
      let bmp = NSBitmapImageRep(data: tiff),
      let cg = bmp.cgImage else {
    FileHandle.standardError.write("无法读取图片\n".data(using: .utf8)!)
    exit(2)
}
let req = VNRecognizeTextRequest()
req.recognitionLevel = .accurate
req.recognitionLanguages = ["zh-Hans", "en-US"]
req.usesLanguageCorrection = false
let handler = VNImageRequestHandler(cgImage: cg, options: [:])
do { try handler.perform([req]) } catch {
    FileHandle.standardError.write("OCR 失败: \(error)\n".data(using: .utf8)!)
    exit(3)
}
var out: [[String: Any]] = []
for ob in (req.results ?? []) {
    guard let top = ob.topCandidates(1).first else { continue }
    let b = ob.boundingBox            // 归一化坐标，原点在左下
    out.append(["text": top.string, "conf": top.confidence,
                "x": b.minX, "y": 1 - b.maxY, "w": b.width, "h": b.height])
}
print(String(data: try! JSONSerialization.data(withJSONObject: out),
             encoding: .utf8)!)
