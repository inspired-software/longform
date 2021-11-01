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
            Section { }
            Section {
                Paragraph {
                    "A String Run!"
                }
                Paragraph {
                    Text("A Text Run!")
                }
            }
            Section {
                "A String Paragraph!"
                for i in 1...5 {
                    if i != 2 {
                        "Paragraph \(i)"
                    } else {
                        "Paragraph Two"
                    }
                }
                Paragraph {
                    for i in 1...5 {
                        "Run \(i)"
                    }
                }
            }
        }
        print(doc.ooxml())
        //XCTAssert(doc.ooxml().contains("<w:p></w:p>"))
    }
}
