//
//  OOXMLGenerator.swift
//  Longform
//
//  Created by Eric Summers on 12/27/21.
//

import Foundation
import Markdown

// TODO: May need variations of this for the different OOXML document types.

public struct OOXMLGenerator: MarkupWalker {
    public private(set) var output: String = ""
    
    /// Specifies the conformance class to which the WordprocessingML document conforms.
    ///
    /// Possible values are:
    /// - `.strict` - The document conforms to Office Open XML Strict
    /// - `.transitional` - The document conforms to Office Open XML Transitional. This is the default value.
    public var conformance: Conformance
    
    /// Document level attributes
    public var documentAttributes: [String] = [
        #"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
        #"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""#,
        #"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
        #"xmlns:w10="urn:schemas-microsoft-com:office:word""#,
        #"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#,
        #"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#,
        #"xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup""#,
        #"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
        #"xmlns:v="urn:schemas-microsoft-com:vml""#,
        #"xmlns:o="urn:schemas-microsoft-com:office:office""#,
        #"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math""#,
        #"mc:Ignorable="w14""#,
    ]
    
    public init(conformance: Conformance = .transitional) {
        self.conformance = conformance
    }
    
    // MARK: Types
    
    /// Specifies the conformance class to which the WordprocessingML document conforms.
    public enum Conformance: String {
        /// The document conforms to Office Open XML Strict
        case strict
        /// The document conforms to Office Open XML Transitional.
        case transitional
    }
    
    // MARK: Document Visitors
    
    public mutating func visitDocument(_ document: Document) {
        let backgroundXML = "" // TODO: Implement this. Probably need to make sure we decend in to a metadata block first.
        let conformanceAttr = conformance != .transitional ? " conformance=\(conformance)" : ""
        let additionalAttrs = documentAttributes.count > 0 ? " " + documentAttributes.joined(separator: " ") : ""
        output.append(
            """
            <?xml version="1.0" encoding="UTF-8"?>
            <w:document\(conformanceAttr)\(additionalAttrs)>\
            \(backgroundXML)\
            <w:body>
            """)
        descendInto(document)
        output.append("</w:body></w:document>")
    }
    
    // MARK: Inline Container Visitors
    
    public mutating func visitEmphasis(_ emphasizedText: Emphasis) {
        // Turn on italics for both standard text (w:i) and complex glyphs (w:iCs).
        output.append(
            """
            <w:r><w:rPr><w:i w:val="1"/><w:iCs w:val="1"/></w:rPr><w:t>\(emphasizedText.plainText)</w:t></w:r>
            """)
    }
    
    public mutating func visitImage(_ image: Image) {
        print(image.title!) //
    }

    public mutating func visitLink(_ link: Link) {
        print(link.destination!) //
    }
    
    public mutating func visitStrikethrough(_ strikethroughText: Strikethrough) {
        print(strikethroughText.plainText) //
    }
    
    public mutating func visitStrong(_ strongText: Strong) {
        // Turn on bold for both standard text (w:b) and complex glyphs (w:bCs).
        output.append(
            """
            <w:r><w:rPr><w:b w:val="1"/><w:bCs w:val="1"/></w:rPr><w:t>\(strongText.plainText)</w:t></w:r>
            """)
    }
    
    // MARK: Inline Leaf Visitors
    
    public mutating func visitCustomInline(_ customText: CustomInline) {
        print(customText.plainText) //
    }
    
    public mutating func visitInlineCode(_ inlineCode: InlineCode) {
        print(inlineCode.plainText) //
    }
    
    public mutating func visitInlineHTML(_ inlineHTML: InlineHTML) {
        print(inlineHTML.plainText) //
    }
    
    public mutating func visitLineBreak(_ lineBreak: LineBreak) {
        print(lineBreak.plainText) // "\n"
    }
    
    public mutating func visitSoftBreak(_ softBreak: SoftBreak) {
        print(softBreak.plainText) // " "
    }
    
    public mutating func visitSymbolLink(_ symbolLink: SymbolLink) {
        // TODO: Maybe don't implement this or treat it like inline code?
        print(symbolLink.plainText) //
    }
    
    public mutating func visitText(_ text: Text) {
        print(text.plainText) //
        output.append("<w:r><w:t>\(text.plainText)</w:t></w:r>")
    }
    
    // MARK: Block Container Block Visitors
    
    public mutating func visitBlockDirective(_ block: BlockDirective) {
        print(block.name) //
        descendInto(block)
    }
    
    public mutating func visitBlockQuote(_ blockQuote: BlockQuote) {
        print("...Block Quote...") //
        descendInto(blockQuote)
    }
    
    public mutating func visitCustomBlock(_ block: CustomBlock) {
        print("...Custom Block...") //
        descendInto(block)
    }
    
    public mutating func visitListItem(_ list: ListItem) {
        // TODO: Handle optional checkbox
        print("...List Item...") //
        descendInto(list)
    }
    
    public mutating func visitOrderedList(_ list: OrderedList) {
        // TODO: Handle optional checkbox
        print("...Ordered List Item...") //
        descendInto(list)
    }
    
    public mutating func visitUnorderedList(_ list: UnorderedList) {
        // TODO: Handle optional checkbox
        print("...Unordered List Item...") //
        descendInto(list)
    }
    
    // MARK: Inline Container Blocks
    
    public mutating func visitParagraph(_ paragraph: Paragraph) {
        output.append(
            """
            <w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr>
            """)
        // TODO: Add Properties
        descendInto(paragraph)
        output.append("</w:p>")
    }
    
    // MARK: Leaf Block Visitors
    
    public mutating func visitCodeBlock(_ block: CodeBlock) {
        print(block.code) //
        descendInto(block)
    }
    
    public mutating func visitHTMLBlock(_ block: HTMLBlock) {
        print(block.rawHTML) //
        descendInto(block)
    }
    
    // TODO: This is putting a paragraph inside a paragraph. I need to abstract paragraph styles in other struct so I can apply it later. In fact run rendering in general probably should be abstracted.
    public mutating func visitHeading(_ heading: Heading) {
        output.append(
            """
            <w:p><w:pPr><w:pStyle w:val="Heading \(heading.level)"/></w:pPr><w:r><w:t>\(heading.plainText)</w:t></w:r></w:p>
            """)
        // TODO: Add Properties, Can headings have other styles on top of them?
    }
    
    public mutating func visitThematicBreak(_ thematicBreak: ThematicBreak) {
        print("...Thematic Break...") //
        descendInto(thematicBreak)
    }
    
    // MARK: Table Visitors
    
    public mutating func visitTable(_ table: Table) {
        print("...Table...") //
        descendInto(table)
    }
    
    public mutating func visitTableBody(_ tableBody: Table.Body) {
        print("...Body...") //
        descendInto(tableBody)
    }
    
    public mutating func visitTableCell(_ cell: Table.Cell) {
        print("...Cell...") //
        descendInto(cell)
    }
    
    public mutating func visitTableHead(_ tableHeader: Table.Head) {
        print("...Header...") //
        descendInto(tableHeader)
    }
    
    public mutating func visitTableRow(_ row: Table.Row) {
        print("...Row...") //
        descendInto(row)
    }
    
}
