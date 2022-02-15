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
            @Metadata(
                previousSection: "foo",
                nextSection: "bar"
                )
            
            @TOC
            
            # Heading 1
            
            @Discrete
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
            
            @Table {
                # First Name, Last Name
                "John", "Smith"
                "**Hello**", "*World*"
            }
            
            <!-- A comment -->
            
            @Comment {
                Various terms to add to a glossary or index.
            
                Docc uses the following syntax, but that is redundant in this case.
            
                - term Some Term: A description for the term.
            
                Should have something analogous for an index.
            
                Glossary might be the default type.
            
                What about some descriptive text at the beginning of a section. Maybe section titles shouldn't be included in DefineTerms...
            }

            @DefineTerms(type: "glossary", section: "glossarySection") {
                # Section Name
            
                Description of this section.
            
                - Some Term: A description for the term.
            }
            
            <!-- If this is not defined then all terms referencing the section would be ignored. -->
            @DefineTermsSection(name: "glossarySection") {
                # Section Name
                        
                A description of this section
            }
            
            @DefineTemplate(name: "someSection") {
                # Section Name
            
                A description of this section
            }
            
            <!-- alternative... -->
            <!-- TODO: Maybe something like the following for table rows. -->
            <!-- Add to a set of templates. By default sort by section name. -->
            @DefineTemplateSet(name: "glossarySection") {
                # Section Name
                        
                A description of this section
                        
                @Terms(type: "glossary")
            }
            
            # Glossary
            
            @Terms(type: "glossary")
            
            # Sites in Alphabetical Order
            
            @Terms {
                - [Swift](https://www.swift.org):
                    A general purpose programming language.
            }
            
            # Ordered Sites
            
            @Terms {
                1. [Swift](https://www.swift.org):
                    A general purpose programming language.
            }
            
            Example link: [Anatomy of a WordProcessingML File](http://officeopenxml.com/anatomyofOOXML.php)
            
            Example Image: ![An Image](image.png "An Image Title")
            
            <!--
            Example Image 2: ^[![An Image](image.png "An Image Title")](width: 100)
            -->
            
            @ImageStyle(name: "image.png", width: 100)
            
            - List Item 1
            
                Testing
            
                <!-- Looks like swift-markdown needs to support cmark attributes -->
                Image: <!-- ^[](image: imageName) -->
                                
                + List Item 2
                        
                + List Item 3
            
            <!-- Maybe allow linking to glossary or hover-over tip. -->
            <!-- ^[Some Term](term: glossary) -->
            
            <!-- It would be nice to move this to the bottom like links, but might not be able to with cmark. -->
            <!-- ^[See footnote.](footnote: "This is a footnote.") -->
            
            Another example link: [WWDC Notes][1]
            
            <!-- Possible way for [inline styles][some-style] maybe? -->
            
            <!--
            [MyStyle]: Style (color: blue)
            [MyEnvVariable]: String "This is a value"
            [2]: http://path.to/image "Image title" width=40px height=400px
            -->
            
            [1]: https://www.wwdcnotes.com/notes/wwdc21/10109/
            
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
