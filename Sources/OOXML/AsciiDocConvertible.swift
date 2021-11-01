//
//  AsciiDocConvertible.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation

public protocol AsciiDocConvertible {
    func asciidoc() -> String
}
