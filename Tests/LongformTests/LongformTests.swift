//
//  LongformTests.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import XCTest
@testable import Longform
@testable import Wordprocessing
import Markdown

struct MyComponent: Component {
    var body: Content {
        Paragraph {
            "Hello World!"
        }
        "Hello Universe!"
    }
}

final class LongformTests: XCTestCase {
    
    func testMarkdownParser() throws {
        let source =
            """
            # Heading 1
            
            ## Heading 2
            
            ## Heading With **Bold** Text
            
            ## **A Bold Heading**
            
            This is a markup *document*.
            
            - SeeAlso: [Anatomy of a WordProcessingML File](http://officeopenxml.com/anatomyofOOXML.php)
            
            """
        let longform = Longform(source: source)
        longform.save(to: URL(fileURLWithPath: "/Users/esummers/Desktop/test.docx"))
    }
    
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
                MyComponent()
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
