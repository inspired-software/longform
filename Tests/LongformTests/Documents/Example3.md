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

### Variables

This is a possible syntax for custom variables: \(myVariable)

Another idea: ^[placeholder value](variable: myVariable)

### Tables

@Table(format: csv, includeHeader: true, style: myTableStyle) {
    First Name, Last Name
    John, Smith
    Mary, Jane
}

### Links
            
Example link: [Anatomy of a WordProcessingML File](http://officeopenxml.com/anatomyofOOXML.php)
            
Another example link: [WWDC Notes][wwdc-notes]
                        
[wwdc-notes]: https://www.wwdcnotes.com/notes/wwdc21/10109/
