//
//  File.swift
//  
//
//  Created by Eric Summers on 10/31/21.
//

import Foundation
import OOXML

/// WordprocessingML text
///
/// - SeeAlso: [Wordprocessing Paragraphs](http://officeopenxml.com/WPparagraph.php)
public struct Text: RunContent {
    init(_ text: String) {
        self.text = text
    }
    
    // TODO: attributed string
    
    var text: String
}
    
extension Text: OOXMLConvertible {
    public func ooxml() -> String {
        // TODO: Remove space preserve when not needed.
        // TODO: Merge adjacent text runs.
        return "<w:t xml:space=\"preserve\">\(text)</w:t>"
    }
}

extension Text: HTMLConvertible {
    public func html() -> String {
        return text
    }
}

