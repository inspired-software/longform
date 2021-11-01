//
//  LongformTests.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import XCTest
@testable import Longform
@testable import Wordprocessing

final class LongformTests: XCTestCase {
    func testOOXMLGeneration() throws {
        let doc = Wordprocessing.Document {
            Paragraph {
                "A String Run!"
                Text("A Text Run!")
            }
            "A String Paragraph!"
            for i in 1...5 {
                if i != 2 {
                    "Number \(i)"
                } else {
                    "Number Two"
                }
            }
        }
        print(doc.ooxml())
        //XCTAssert(doc.ooxml().contains("<w:p></w:p>"))
    }
}
