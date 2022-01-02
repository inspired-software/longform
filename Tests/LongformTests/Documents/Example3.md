@TOC
            
# Heading 1
            
## Heading 2
            
## Heading With **Bold** Text
            
## **A Bold Heading**
            
@Style(name: Alt Body Text, keepNext: true) {
    This is a markup *document*.
}
            
This is markup with a *<style color="blue">**custom inline**</style>* attribute.

<style name="Alt Body Text">
This markup has style overridden
</style>

@Override(style: body, font: Helvetica, fontSize: 14)

This is markup with a ^[custom inline](color: blue, bold: true) attribute. Another option would be ^[custom inline 2](style: color-text)

Maybe allow dropping the carrot. This is markup with a [custom inline](color: blue, bold: true) attribute.

### Variables

This is a possible syntax for custom variables: \(myEnvironmentVariable)

Another idea: ^[placeholder value](variable: myEnvironmentVariable)

Maybe use Yaml syntax in environment block?

@Environment {
    myEnvironmentVariable: Hello World.
}

@DocumentProperties {
    Author: John Smith
}

### Comments

[This is a comment that will be hidden.]: # 

@Comment { 
    Alternate comment block.
}

### Image

[Maybe extend cmark to add parameters?]: #
![My Image](https://example.com/someimg.png, scale: fit)

### Tables

@Table(format: csv, includeHeader: true, style: myTableStyle) {
    First Name, Last Name
    John, Smith
    Mary, Jane
}

### Links
   
@Comment { Maybe extend cmark to add parameters? }            
Example link: [Anatomy of a WordProcessingML File](http://officeopenxml.com/anatomyofOOXML.php, color: blue)
            
Another example link: [WWDC Notes][wwdc-notes]
          
[wwdc-notes]: https://www.wwdcnotes.com/notes/wwdc21/10109/ "WWDC Notes"

[MyStyle]: style (color: blue)

[myEnvironmentVariable]: Hello World.
