//
//  ParagraphProperties.swift
//  Longform
//
//  Created by Eric Summers on 12/30/21.
//

import Foundation
import AppKit

/// Specifies the alignment of characters on each line when characters are of varying size.
public enum VerticalTextAlignment: String {
    case auto
    case baseline
    case bottom
    case center
    case top
    
    func ooxml() -> String {
        return #"<w:textAlignment w:val="\#(rawValue)"/>"#
    }
}

/// Properties to apply to a paragraph or run.
public struct ParagraphProperties {
    public init() { }
    
    /// Defines the indentation of a paragraph.
    ///
    /// - Note: Some less common indentation options are not implemented.
    public struct Indentation {
        /// Specifies the indentation to be placed at the left (for paragraphs going left to right). Value is in twentieths of a point.
        public var leading: Int? = nil
        
        /// Specifies the indentation to be placed at the right (for paragraphs going left to right). Value is in twentieths of a point.
        public var trailing: Int? = nil
        
        /// Specifies indentation to be removed from the first line. Value is in twentieths of a point.
        ///
        /// This attribute and `firstLine` are mutually exclusive. This attribute controls when both are specified.
        public var hanging: Int? = nil
        
        /// Specifies additional indentation to be applied to the first line. Value is in twentieths of a point.
        ///
        /// This attribute and `hanging` are mutually exclusive. This attribute is ignored if `hanging` is specified.
        public var firstLine: Int? = nil
        
        func ooxml() -> String {
            if leading == nil && trailing == nil && hanging == nil && firstLine == nil {
                return ""
            }
            var output: String = "<w:ind "
            if let leading = leading { output.append(#"w:start="\#(leading)" "#) }
            if let trailing = trailing { output.append(#"w:end="\#(trailing)" "#) }
            if let hanging = hanging { output.append(#"w:hanging="\#(hanging)" "#) }
            if let firstLine = firstLine, hanging == nil { output.append(#"w:firstLine="\#(firstLine)" "#) }
            output.append("/>")
            return output
        }
    }
    
//    /// Defines the spacing between paragraphs and between lines of a paragaph.
//    ///
//    /// - Note: Some less common spacing options are not implemented.
//    ///
//    /// - SeeAlso:
//    /// * [SpacingBetweenLines](https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.spacingbetweenlines)
//    /// * [Wordprocessing Spacing](http://officeopenxml.com/WPspacing.php)
//    public enum SpacingBetweenLines {
//        /// The value is interpreted as 240th of a line.
//        case line(before: LineRuleValue?, after: LineRuleValue?)
//        /// The line height is interpreted as 240th of a point.
//        case exact(before: LineRuleValue?, after: LineRuleValue?)
//        /// The minimum line height is interpreted as 240th of a point.
//        case atLeast(before: LineRuleValue?, after: LineRuleValue?)
//
//        public func lineRule: String {
//            switch self {
//            case .atLeast(before: _, after: _): return "atLeast"
//            case .exact(before: _, after: _): return "exact"
//            case .atLeast(before: _, after: _): return "atLeast"
//            }
//        }
//
//        func ooxml() -> String {
//
//        }
//
//    }


//    public struct Spacing {
//
//
//        public enum LineRuleValue {
//            /// Spacing should be determined by the wordprocessor
//            case auto
//
//            case value(Int)
//        }
//
//        /// Specifies the spacing that should be added before the first line of the paragraph. Value is in twentieths of a point.
//        public var before: Int? = nil
//
//        /// Specifies the spacing that should be added after the last line of the paragraph. Value is in twentieths of a point.
//        public var after: Int? = nil
//
//        /// Specifies the amount of vertical spacing between lines of text within the paragraph.
//        ///
//        /// If the value of the `lineRule` attribute is `.atLeast` or `.exactly`, then the value of the line attribute is interpreted as 240th of a point.
//        /// If the value of `lineRule` is `.auto`, then the value of line is interpreted as 240th of a line.
//        public var line: Int? = nil
//
//        public var lineRule: LineRule? = nil
//
//        /// Specifies whether spacing before the paragraph should be determined by the wordprocessor.
//        public var autospacing: Autospacing? = nil
//
//        func ooxml() -> String {
//            if leading == nil && trailing == nil && hanging == nil && firstLine == nil {
//                return ""
//            }
//            var output: String = "<w:spacing "
//            if let before = before {
//                output.append(#"w:before="\#(leading)" "#)
//            }
//            if let after = after {
//                output.append(#"w:after="\#(trailing)" "#)
//            }
//            if let lineRule = lineRule {
//                output.append(#"w:lineRule="\#(lineRule.rawValue)" "#)
//            }
//            if let autospacing = autospacing {
//                if autospacing == .before || autospacing == .beforeAndAfter {
//                    output.append(#"w:beforeAutospacing="1" "#)
//                }
//                if autospacing == .before || autospacing == .beforeAndAfter {
//                    output.append(#"w:afterAutospacing="1" "#)
//                }
//            }
//            output.append("/>")
//            return output
//        }
//    }
    
    public var style: String? = nil // TODO: Maybe this should refer to a Style class instead.
    public var italicize: Bool = false
    public var italicizeComplexGlyphs: Bool = false
    public var bold: Bool = false
    public var boldComplexGlyphs: Bool = false
    
    /// Specifies the alignment of characters on each line when characters are of varying size.
    public var verticalTextAlignment: VerticalTextAlignment? = nil
    
    /// Defines the indentation for the paragraph.
    public var indentation: Indentation = Indentation()
    
    /// Defines the spacing between paragraphs and between lines of a paragaph.
    //public var spacing: Spacing = Spacing()
    
    func ooxml() -> String {
        var output: String = "<w:pPr>"
        if let style = style {
            output.append(#"<w:pStyle w:val="\#(style)"/>"#)
        }
        if italicize { output.append(#"<w:i w:val="1"/>"#) }
        if italicizeComplexGlyphs { output.append(#"<w:iCs w:val="1"/>"#) }
        if bold { output.append(#"<w:b w:val="1"/>"#) }
        if boldComplexGlyphs { output.append(#"<w:bCs w:val="1"/>"#) }
        if let alignment = verticalTextAlignment { output.append(alignment.ooxml()) }
        
        output.append("</w:pPr>")
        return output
    }
    
}
