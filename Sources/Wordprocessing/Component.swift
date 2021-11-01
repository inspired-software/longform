//
//  Component.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation

// TODO: Should I allow a component to have multiple sections? That might require allowing sections to be nested as a sectioncomponent. I'm undecided if this makes sense.

/// A component that may be reused.
public protocol Component {
    typealias Content = [SectionContent]
    
    @ContentBuilder<SectionContent> var body: Content { get }
}

extension SectionContent {
    // Allow `SectionContent` to look like a `Component`.
    public var body: Content { [self] }
}

extension ContentBuilder where Content == Component {
    // TODO: Change this to a markdown style attributed string or run through the Longform Markdown processor.
    public static func buildExpression(_ text: String) -> [Content] { [ Paragraph() { Text(text) } ] }
}
