//
//  ContentBuilder.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation
import OOXML

// TODO: Should attempt to not retain the whole tree or only keep serialized data when possible to reduce allocations.

@resultBuilder
public struct ContentBuilder<Content> {
    public static func buildExpression(_ content: Content) -> [Content] { [content] }
    public static func buildBlock() -> [Content] { [] }
    public static func buildBlock(_ contents: [Content]...) -> [Content] { contents.flatMap { $0 } }
    public static func buildOptional(_ contents: [Content]?) -> [Content] { contents ?? [] }
    public static func buildEither(first contents: [Content]) -> [Content] { contents }
    public static func buildEither(second contents: [Content]) -> [Content] { contents }
    public static func buildArray(_ contents: [[Content]]) -> [Content] { contents.flatMap { $0 } }
}
