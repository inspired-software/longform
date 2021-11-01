//
//  Paragraph.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation
import OOXML

// Example OOXML:
// <w:p><!-- paragraph -->
//     <w:pPr><!-- paragraph properties -->
//         <w:pStyle> w:val="NormalWeb"/>
//         <w:spacing w:before="120" w:after="120"/>
//     </w:pPr>
//     <w:r><!-- text run -->
//         <w:t xml:space="preserve">Hello World!</w:t><!-- text -->
//     </w:r>
// </w:p>

// MARK: - Paragraph

/// A WordprocessingML paragraph
///
/// - SeeAlso: [Wordprocessing Paragraphs](http://officeopenxml.com/WPparagraph.php)
public struct Paragraph: PrimativeBodyElement {
    public var properties: [ParagraphProperty] = [] // TODO: Maybe make this a dictionary keyed on class type.
    public var runElements: [PrimativeRunElement] = []
}
    
extension Paragraph: OOXMLConvertible {
    public func ooxml() -> String {
        let propertiesXML = properties.count > 0 ? "<w:pPr>\(properties.map { $0.ooxml() }.joined())</w:pPr>" : ""
        return "<w:p>\(propertiesXML)<w:r>\(runElements.map { $0.ooxml() }.joined())</w:r></w:p>"
    }
}

extension Paragraph: HTMLConvertible {
    public func html() -> String {
        return "<p>\(runElements.map { $0.html() }.joined())</p>"
    }
}

// MARK: -

// Allow strings to be used as a BodyElement
extension String: BodyElement {
    // TODO: Change this to a markdown style attributed string.
    public var primativeBodyElement: PrimativeBodyElement { var p = Paragraph(); p.runElements = [ Text(self) ]; return p } // TODO: embed text with builder
}

// MARK: -

public protocol ParagraphProperty: OOXMLExportable { }

// TODO: Should this be a enum or protocol? Look at making similar to AttributedString properties
/*
public enum ParagraphProperties {
    case frameProperties:
}
 */

// MARK: -

/// A paragraph element that may be converted to a primative OOXML element.
public protocol RunElement {
    var primativeRunElement: PrimativeRunElement { get }
}

/// A primative OOXML paragraph element.
public protocol PrimativeRunElement: RunElement, OOXMLExportable { }

extension PrimativeRunElement {
    public var primativeRunElement: PrimativeRunElement { self }
}
