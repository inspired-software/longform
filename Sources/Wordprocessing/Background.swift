//
//  Background.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation
import OOXML

/// The background can be either an DrawingML object or a sold color. If it is a DrawingML object,
/// then the background element contains a drawing element. If a solid color is used, then background is an empty element, with the color specified in the
/// attributes.
///
/// - SeeAlso: [WordprocessingML Document](http://officeopenxml.com/WPdocument.php)
public struct Background {
        
    // MARK: Attributes
    
    /// Specifies the color
    ///
    /// Possible values are either hex-encoded RGB values (in RRGGBB format) or `auto`.
    ///
    /// Example: `"2C34FF"`
    var color: String? = nil
    
    /// Specifies the base theme color (which is specified in the Theme part)
    ///
    /// Example:  `"accent5"`
    var themeColor: String? = nil
    
    /// Specifies the shade value applied to the theme color (in hex encoding of values 0-255)
    ///
    /// Example: `"BF"`
    var themeShade: String? = nil
    
    /// Specifies the tint value applied to the theme color (in hex encoding of values 0-255)
    ///
    /// Example: `"99"`
    var themeTint: String? = nil
    
    // TODO: Maybe implement these attributes as an enum.
    
    // MARK: Elements
    
    // TODO: Add elements here
}

extension Background: OOXMLConvertible {
    public func ooxml() -> String {
        // TODO: elements
        let attributes = """
            \(color.map { " color=\"\($0)\"" } ?? "")\
            \(themeColor.map { " themeColor=\"\($0)\"" } ?? "")\
            \(themeShade.map { " themeShade=\"\($0)\"" } ?? "")\
            \(themeTint.map { " themeTint=\"\($0)\"" } ?? "")
            """
        return "<w:background\(attributes) />"
    }
}

extension Background: HTMLConvertible {
    public func html() -> String {
        return ""
    }
}
