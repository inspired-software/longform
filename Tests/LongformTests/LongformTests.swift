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
            Paragraph()
            "Hello World!"
        }
        print(doc.ooxml())
        //XCTAssert(doc.ooxml().contains("<w:p></w:p>"))
    }
}
