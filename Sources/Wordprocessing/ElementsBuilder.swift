//
//  ElementsBuilder.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation

@resultBuilder
public struct ElementsBuilder<Element> {
    public static func buildExpression(_ element: Element) -> [Element] { [element] }
    static public func buildBlock() -> [Element] { [] }
    public static func buildBlock(_ elements: [Element]...) -> [Element] { elements.flatMap { $0 } }
    public static func buildOptional(_ elements: [Element]?) -> [Element] { elements ?? [] }
    public static func buildEither(first elements: [Element]) -> [Element] { elements }
    public static func buildEither(second elements: [Element]) -> [Element] { elements }
    public static func buildArray(_ elements: [[Element]]) -> [Element] { elements.flatMap { $0 } }
}
