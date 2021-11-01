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
public struct Paragraph: SectionElement {
    
    public init(@ElementsBuilder<RunElement> _ runElements: () -> [RunElement]) {
        self.properties = []
        self.runElements = runElements()
    }
    
    var properties: [ParagraphProperty] // TODO: Maybe make this a dictionary keyed on class type.
    var runElements: [RunElement]
}
    
extension Paragraph: OOXMLConvertible {
    func ooxml(sectionProperties: SectionProperties?) -> String {
        let props = sectionProperties.map { properties + [ $0 ] } ?? properties
        let propertiesXML = props.count > 0 ? "<w:pPr>\(props.map { $0.ooxml() }.joined())</w:pPr>" : ""
        return "<w:p>\(propertiesXML)<w:r>\(runElements.map { $0.ooxml() }.joined())</w:r></w:p>"
    }
    
    public func ooxml() -> String { ooxml(sectionProperties: nil) }
}

extension Paragraph: HTMLConvertible {
    public func html() -> String {
        return "<p>\(runElements.map { $0.html() }.joined())</p>"
    }
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

/// A primative OOXML run element.
public protocol RunElement: OOXMLExportable { }

extension ElementsBuilder where Element == RunElement {
    // TODO: Change this to a markdown style attributed string or run through the Longform Markdown processor.
    public static func buildExpression(_ text: String) -> [Element] { [ Text(text) ] }
}
