//
//  Document.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation
import OOXML

/*
 Example OOXML:
 <?xml version="1.0" encoding="UTF-8"?>
 <w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:w10="urn:schemas-microsoft-com:office:word"
     xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
     xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
     mc:Ignorable="w14">
    <w:body>
        ...
    </w:body>
 </w:document>
 */

/// An OOXML WordprocessingML document
///
/// - SeeAlso: [Wordprocessing Document](http://officeopenxml.com/WPdocument.php)
public struct Document {
    
    public init(conformance: Conformance = .transitional, @DocumentBuilder _ bodyElements: () -> [PrimativeBodyElement]) {
        self.conformance = conformance
        self.bodyElements = bodyElements()
    }
    
    // MARK: Types
    
    /// Specifies the conformance class to which the WordprocessingML document conforms.
    public enum Conformance: String {
        /// The document conforms to Office Open XML Strict
        case strict
        /// The document conforms to Office Open XML Transitional.
        case transitional
    }
    
    // MARK: Attributes
    
    /// Specifies the conformance class to which the WordprocessingML document conforms.
    ///
    /// Possible values are:
    /// - `.strict` - The document conforms to Office Open XML Strict
    /// - `.transitional` - The document conforms to Office Open XML Transitional. This is the default value.
    public var conformance: Conformance
    
    /// Additional attributes to add
    public var additionalAttributes: [String] = [
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"",
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
        "xmlns:w10=\"urn:schemas-microsoft-com:office:word\"",
        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"",
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"",
        "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\"",
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",
        "xmlns:v=\"urn:schemas-microsoft-com:vml\"",
        "xmlns:o=\"urn:schemas-microsoft-com:office:office\"",
        "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"",
        "mc:Ignorable=\"w14\"",
    ]
    
    // MARK: Elements
    
    /// Specifies the contents of the body of the document
    ///
    /// Common elements include:
    /// - `Paragraph` - Specifies a paragraph of content
    /// - `SectionProperties` - Specifies the section properties for the final section
    /// - `Table` - Specifies a table
    public var bodyElements: [PrimativeBodyElement] = []
    
    /// Specifies the background for every page of the document
    public var background: Background? = nil
}

extension Document: OOXMLConvertible {
    public func ooxml() -> String {
        let backgroundXML = background?.ooxml() ?? ""
        let conformanceAttr = conformance != .transitional ? " conformance=\(conformance)" : ""
        let additionalAttrs = additionalAttributes.count > 0 ? " " + additionalAttributes.joined(separator: " ") : ""
        return """
        <?xml version="1.0" encoding="UTF-8"?>\
        <w:document\(conformanceAttr)\(additionalAttrs)>\
        \(backgroundXML)\
        <w:body>\(bodyElements.map { $0.ooxml() }.joined())</w:body>\
        </w:document>
        """
    }
}

extension Document: HTMLConvertible {
    public func html() -> String {
        return """
        <?xml version="1.0" encoding="UTF-8"?>\
        <html>\
        <head></head>\
        <body>\(bodyElements.map { $0.html() }.joined())</body>\
        </html>
        """
    }
}

// MARK: -

@resultBuilder
public struct DocumentBuilder {
    static public func buildBlock() -> [PrimativeBodyElement] { [] }
    static public func buildBlock(_ elements: BodyElement...) -> [PrimativeBodyElement] { elements.map { $0.primativeBodyElement } }
    // TODO: Conditionals
}

// MARK: -

public protocol BodyElement {
    var primativeBodyElement: PrimativeBodyElement { get }
}

public protocol PrimativeBodyElement: BodyElement, OOXMLExportable { }

extension PrimativeBodyElement {
    public var primativeBodyElement: PrimativeBodyElement { self }
}
