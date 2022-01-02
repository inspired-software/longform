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
            @TOC
            
            # Heading 1
            
            ## Heading 2
            
            ## Heading With **Bold** Text
            
            ## **A Bold Heading**
            
            @Style(plain, keepNext: true) {
               This is a markup *document*.
            }
            
            This is markup with a *<style color="blue">**custom inline**</style>* attribute.
            
            <style name="Alt Body Text">
            This markup has *style* overridden
            </style>
            
            Example link: [Anatomy of a WordProcessingML File](http://officeopenxml.com/anatomyofOOXML.php)
            
            Another example link: [WWDC Notes][wwdc-notes]
            
            Possible way for [inline styles][some-style] maybe?
            
            [wwdc-notes]: https://www.wwdcnotes.com/notes/wwdc21/10109/
            [some-style]: ^(color:blue,bold:true)
            
            """
        // Note: this doesn't work "This is markup with a ^[custom inline](color: blue, bold: true) attribute."
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
