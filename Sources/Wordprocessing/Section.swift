//
//  File.swift
//  
//
//  Created by Eric Summers on 10/31/21.
//

import Foundation
import OOXML

/// A WordprocessingML section
///
/// - SeeAlso: [Wordprocessing Sections](http://officeopenxml.com/WPsection.php)
public struct Section {
 
    public init(@ElementsBuilder<SectionElement> _ sectionElements: () -> [SectionElement]) {
        self.sectionElements = sectionElements()
    }
    
    init(sectionElements: [SectionElement]) {
        self.sectionElements = sectionElements
    }
    
    var sectionElements: [SectionElement]
    
    var sectionProperties: SectionProperties = SectionProperties()
    
    public var properties: [SectionProperty] {
        get { sectionProperties.properties }
        set { sectionProperties.properties = newValue }
    }

    public func ooxml(lastSection: Bool) -> String {
        guard !sectionElements.isEmpty || !properties.isEmpty else { return "" }
        if lastSection {
            // The last section is handled differently
            return sectionElements.map { $0.ooxml() }.joined() + sectionProperties.ooxml()
        } else {
            var elements: [SectionElement]
            var lastParagraphXML: String
            if properties.isEmpty {
                // There are no section properties
                elements = sectionElements
                lastParagraphXML = ""
            } else if !sectionElements.isEmpty, let paragraph = sectionElements.last as? Paragraph {
                // The last section element is a paragraph
                elements = sectionElements.dropLast()
                lastParagraphXML = paragraph.ooxml(sectionProperties: sectionProperties)
            } else {
                // The last section element is not a paragraph
                elements = sectionElements
                lastParagraphXML = Paragraph { }.ooxml(sectionProperties: sectionProperties)
            }
            return elements.map { $0.ooxml() }.joined() + lastParagraphXML
        }
    }

    public func html() -> String {
        return sectionElements.map { $0.html() }.joined()
    }
}

// MARK: -

public protocol SectionElement: OOXMLExportable { }

extension ElementsBuilder where Element == SectionElement {
    // TODO: Change this to a markdown style attributed string or run through the Longform Markdown processor.
    public static func buildExpression(_ text: String) -> [Element] { [ Paragraph() { Text(text) } ] }
}

// MARK: -

public protocol SectionProperty: OOXMLExportable { }

// MARK: -

/// The secction properties for the last paragraph in a section.
struct SectionProperties: ParagraphProperty {
    init() { }
    
    var properties: [SectionProperty] = []
}

extension SectionProperties: OOXMLConvertible {
    public func ooxml() -> String {
        return !properties.isEmpty ? "<w:sectPr>\(properties.map { $0.ooxml() }.joined())</w:sectPr>" : ""
    }
}

extension SectionProperties: HTMLConvertible {
    public func html() -> String {
        return ""
    }
}
