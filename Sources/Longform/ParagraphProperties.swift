//
//  ParagraphProperties.swift
//  Longform
//
//  Created by Eric Summers on 12/30/21.
//

import Foundation
import AppKit

public enum TextAlignment: String {
    case auto
    case baseline
    case bottom
    case center
    case top
}

/// Properties to apply to a paragraph or run.
public struct ParagraphProperties {
    public init() { }
    
    public var style: String? = nil
    public var italicize: Bool = false
    public var italicizeComplexGlyphs: Bool = false
    public var bold: Bool = false
    public var boldComplexGlyphs: Bool = false
    public var textAlignment: TextAlignment? = nil
    
    func ooxml() -> String {
        var output: String = "<w:pPr>"
        if let textAlignment = textAlignment {
            output.append(
                """
                <w:textAlignment w:val="\(textAlignment.rawValue)"/>
                """)
        }
        output.append("</w:pPr>")
        return output
    }
    
}
