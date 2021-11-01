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
 
    public init(@ContentBuilder<Component> _ components: () -> [Component]) {
        self.components = components()
    }
    
    init(components: [Component]) {
        self.components = components
    }
    
    var components: [Component]
    
    func content() -> [SectionContent] { components.flatMap { $0.body } }
    
    var sectionProperties: SectionProperties = SectionProperties()
    
    public var properties: [SectionProperty] {
        get { sectionProperties.properties }
        set { sectionProperties.properties = newValue }
    }

    public func ooxml(lastSection: Bool) -> String {
        let content = content()
        guard !content.isEmpty || !properties.isEmpty else { return "" }
        if lastSection {
            // The last section is handled differently
            return content.map { $0.ooxml() }.joined() + sectionProperties.ooxml()
        } else {
            var sectContent: [SectionContent]
            var lastParagraphXML: String
            if properties.isEmpty {
                // There are no section properties
                sectContent = content
                lastParagraphXML = ""
            } else if !content.isEmpty, let paragraph = content.last as? Paragraph {
                // The last section element is a paragraph
                sectContent = content.dropLast()
                lastParagraphXML = paragraph.ooxml(sectionProperties: sectionProperties)
            } else {
                // The last section element is not a paragraph
                sectContent = content
                lastParagraphXML = Paragraph { }.ooxml(sectionProperties: sectionProperties)
            }
            return sectContent.map { $0.ooxml() }.joined() + lastParagraphXML
        }
    }

    public func html() -> String {
        let content = content()
        return content.map { $0.html() }.joined()
    }
    
    // MARK: Modifiers
    
    public enum SectionOrientation: String {
        case portrait
        case landscape
    }
    
    /*
    @inlinable public func pageSize(width: Int, height: Int, orientation: SectionOrientation) -> Paragraph {
        return self
    }
     */
}

// MARK: -

public protocol SectionContent: OOXMLExportable, Component { }

extension ContentBuilder where Content == SectionContent {
    // TODO: Change this to a markdown style attributed string or run through the Longform Markdown processor.
    public static func buildExpression(_ text: String) -> [Content] { [ Paragraph() { Text(text) } ] }
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
