//
//  OOXMLConvertible.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation

// TODO: Also implement HTMLConvertable with modifiers similar to: https://github.com/JohnSundell/Ink/blob/master/Sources/Ink/Internal/HTMLConvertible.swift

public protocol OOXMLConvertible {
    func ooxml() -> String
}
