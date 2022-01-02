//
//  Longform.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import Foundation
import ZIPFoundation
import Markdown

// TODO: Rename to Package
/// - SeeAlso: [Anatomy of a WordProcessingML File](http://officeopenxml.com/anatomyofOOXML.php)
public struct Longform {
    var content: String = ""
    
    let contentRelsXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>\
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>\
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>\
    <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>\
    </Relationships>
    """
    
    let header1XML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <w:hdr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" mc:Ignorable="w14"><w:p><w:r/></w:p></w:hdr>
    """
    
    let footer1XML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <w:ftr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" mc:Ignorable="w14"><w:p><w:r/></w:p></w:ftr>
    """
    
    let contentTypesXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
    <Default Extension="xml" ContentType="application/xml"/>\
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
    <Default Extension="jpeg" ContentType="image/jpg"/>\
    <Default Extension="png" ContentType="image/png"/>\
    <Default Extension="bmp" ContentType="image/bmp"/>\
    <Default Extension="gif" ContentType="image/gif"/>\
    <Default Extension="tif" ContentType="image/tif"/>\
    <Default Extension="pdf" ContentType="application/pdf"/>\
    <Default Extension="mov" ContentType="application/movie"/>\
    <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>\
    <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>\
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>\
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\
    <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\
    <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>\
    <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\
    </Types>
    """

    let relsXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\
    </Relationships>
    """
    
    let docPropsAppXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
    """
    
    let docPropsCoreXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>
    """
    
    let fontTableXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">\
    <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="E0002AFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font>\
    <w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font>\
    <w:font w:name="Calibri Light"><w:panose1 w:val="020F0302020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="E0002AFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font>\
    </w:fonts>
    """
    
    let settingsXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\
    <w:view w:val="print"/>\
    <w:mirrorMargins w:val="0"/>\
    <w:bordersDoNotSurroundHeader w:val="0"/>\
    <w:bordersDoNotSurroundFooter w:val="0"/>\
    <w:displayBackgroundShape/>\
    <w:revisionView w:markup="1" w:comments="1" w:insDel="1" w:formatting="0"/>\
    <w:defaultTabStop w:val="720"/>\
    <w:autoHyphenation w:val="0"/>\
    <w:evenAndOddHeaders w:val="0"/>\
    <w:bookFoldPrinting w:val="0"/>\
    <w:noLineBreaksAfter w:lang="English" w:val="‘“(〔[{〈《「『【⦅〘〖«〝︵︷︹︻︽︿﹁﹃﹇﹙﹛﹝｢"/>\
    <w:noLineBreaksBefore w:lang="English" w:val="’”)〕]}〉"/>\
    <w:compat>\
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>\
    </w:compat>\
    <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>\
    </w:settings>
    """
    /*
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh"><w:zoom w:percent="141"/><w:proofState w:spelling="clean" w:grammar="clean"/>\
        <w:defaultTabStop w:val="720"/>\
        <w:characterSpacingControl w:val="doNotCompress"/>\
        <w:compat>\
        <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>\
        <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>\
        <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>\
        <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>\
        <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>\
        <w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>\
        </w:compat>\
        <w:rsids>\
        <w:rsidRoot w:val="00120ED4"/>\
        <w:rsid w:val="00120ED4"/>\
        <w:rsid w:val="00186729"/>\
        <w:rsid w:val="00593A65"/>\
        <w:rsid w:val="00EE1094"/>\
        <w:rsid w:val="00FD2D82"/>\
        </w:rsids>\
        <m:mathPr>\
        <m:mathFont m:val="Cambria Math"/>\
        <m:brkBin m:val="before"/>\
        <m:brkBinSub m:val="--"/>\
        <m:smallFrac m:val="0"/>\
        <m:dispDef/>\
        <m:lMargin m:val="0"/>\
        <m:rMargin m:val="0"/>\
        <m:defJc m:val="centerGroup"/>\
        <m:wrapIndent m:val="1440"/>\
        <m:intLim m:val="subSup"/>\
        <m:naryLim m:val="undOvr"/>\
        </m:mathPr>\
        <w:themeFontLang w:val="en-US"/>\
        <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>\
        <w:decimalSymbol w:val="."/>\
        <w:listSeparator w:val=","/>\
        <w14:docId w14:val="21873B23"/>\
        <w15:chartTrackingRefBased/>\
        <w15:docId w15:val="{B91AA399-ACC0-B744-A018-956FE4C199F3}"/>\
        </w:settings>
        """
    */
    
    let webSettingsXML = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">\
    <w:optimizeForBrowser/>\
    <w:allowPNG/>\
    </w:webSettings>
    """
    
    let stylesXML = """
    <?xml version="1.0" encoding="UTF-8"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Arial Unicode MS"/><w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="0"/><w:iCs w:val="0"/><w:caps w:val="0"/><w:smallCaps w:val="0"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:outline w:val="0"/><w:emboss w:val="0"/><w:imprint w:val="0"/><w:vanish w:val="0"/><w:color w:val="auto"/><w:spacing w:val="0"/><w:w w:val="100"/><w:kern w:val="0"/><w:position w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:u w:val="none" w:color="auto"/><w:bdr w:val="nil"/><w:vertAlign w:val="baseline"/><w:lang/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/><w:framePr w:anchorLock="0" w:w="0" w:h="0" w:vSpace="0" w:hSpace="0" w:xAlign="left" w:y="0" w:hRule="exact" w:vAnchor="margin"/><w:widowControl w:val="1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr><w:suppressLineNumbers w:val="0"/><w:pBdr><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:between w:val="nil"/><w:bar w:val="nil"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:suppressAutoHyphens w:val="0"/><w:spacing w:before="0" w:beforeAutospacing="0" w:after="0" w:afterAutospacing="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="9"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:next w:val="Normal"/><w:pPr/><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="Default Paragraph Font"><w:name w:val="Default Paragraph Font"/><w:next w:val="Default Paragraph Font"/></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style><w:style w:type="table" w:default="1" w:styleId="Table Normal"><w:name w:val="Table Normal"/><w:next w:val="Table Normal"/><w:pPr/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/></w:tblPr><w:trPr/><w:tcPr/><w:tblStylePr w:type="firstRow"/><w:tblStylePr w:type="lastRow"/><w:tblStylePr w:type="firstCol"/><w:tblStylePr w:type="lastCol"/><w:tblStylePr w:type="band1Vert"/><w:tblStylePr w:type="band2Vert"/><w:tblStylePr w:type="band1Horz"/><w:tblStylePr w:type="band2Horz"/><w:tblStylePr w:type="neCell"/><w:tblStylePr w:type="nwCell"/><w:tblStylePr w:type="seCell"/><w:tblStylePr w:type="swCell"/></w:style><w:style w:type="numbering" w:default="1" w:styleId="No List"><w:name w:val="No List"/><w:next w:val="No List"/><w:pPr/></w:style><w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/><w:next w:val="Body"/><w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/><w:widowControl w:val="1"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:suppressAutoHyphens w:val="0"/><w:bidi w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Helvetica Neue" w:cs="Arial Unicode MS" w:hAnsi="Helvetica Neue" w:eastAsia="Arial Unicode MS"/><w:b w:val="1"/><w:bCs w:val="1"/><w:i w:val="0"/><w:iCs w:val="0"/><w:caps w:val="0"/><w:smallCaps w:val="0"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:outline w:val="0"/><w:color w:val="000000"/><w:spacing w:val="0"/><w:kern w:val="0"/><w:position w:val="0"/><w:sz w:val="36"/><w:szCs w:val="36"/><w:u w:val="none"/><w:shd w:val="nil" w:color="auto" w:fill="auto"/><w:vertAlign w:val="baseline"/><w:lang w:val="en-US"/><w14:textOutline><w14:noFill/></w14:textOutline><w14:textFill><w14:solidFill><w14:srgbClr w14:val="000000"/></w14:solidFill></w14:textFill></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:next w:val="Body"/><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/><w:widowControl w:val="1"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:suppressAutoHyphens w:val="0"/><w:bidi w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="9"/></w:pPr><w:rPr><w:rFonts w:ascii="Helvetica Neue" w:cs="Arial Unicode MS" w:hAnsi="Helvetica Neue" w:eastAsia="Arial Unicode MS"/><w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="0"/><w:iCs w:val="0"/><w:caps w:val="0"/><w:smallCaps w:val="0"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:outline w:val="0"/><w:color w:val="000000"/><w:spacing w:val="0"/><w:kern w:val="0"/><w:position w:val="0"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:u w:val="none"/><w:shd w:val="nil" w:color="auto" w:fill="auto"/><w:vertAlign w:val="baseline"/><w:lang w:val="en-US"/><w14:textOutline><w14:noFill/></w14:textOutline><w14:textFill><w14:solidFill><w14:srgbClr w14:val="000000"/></w14:solidFill></w14:textFill></w:rPr></w:style></w:styles>
    """
    /*
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">\
        <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>\
        <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376">\
        <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>\
        <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>\
        <w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="index 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 6" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 7" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 8" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index 9" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 1" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 2" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 3" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 4" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 5" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 6" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 7" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 8" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toc 9" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Normal Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="footnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="annotation text" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="header" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="footer" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="index heading" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="caption" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="table of figures" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="envelope address" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="envelope return" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="footnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="annotation reference" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="line number" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="page number" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="endnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="endnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="table of authorities" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="macro" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="toa heading" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Bullet" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Number" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Bullet 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Bullet 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Bullet 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Bullet 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Number 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Number 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Number 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Number 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>\
        <w:lsdException w:name="Closing" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Continue" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Continue 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Continue 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Continue 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="List Continue 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Message Header" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>\
        <w:lsdException w:name="Salutation" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Date" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text First Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text First Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Note Heading" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Body Text Indent 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Block Text" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>\
        <w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>\
        <w:lsdException w:name="Document Map" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Plain Text" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="E-mail Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Top of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Bottom of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Normal (Web)" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Acronym" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Address" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Cite" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Code" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Definition" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Keyboard" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Preformatted" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Sample" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Typewriter" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="HTML Variable" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Normal Table" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="annotation subject" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="No List" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Outline List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Outline List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Outline List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Simple 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Simple 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Simple 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Classic 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Classic 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Classic 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Classic 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Colorful 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Colorful 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Colorful 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Columns 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Columns 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Columns 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Columns 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Columns 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 6" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 7" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid 8" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 6" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 7" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table List 8" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table 3D effects 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table 3D effects 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table 3D effects 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Contemporary" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Elegant" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Professional" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Subtle 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Subtle 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Web 1" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Web 2" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Web 3" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Balloon Text" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Table Grid" w:uiPriority="39"/>\
        <w:lsdException w:name="Table Theme" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Placeholder Text" w:semiHidden="1"/>\
        <w:lsdException w:name="No Spacing" w:uiPriority="1" w:qFormat="1"/>\
        <w:lsdException w:name="Light Shading" w:uiPriority="60"/>\
        <w:lsdException w:name="Light List" w:uiPriority="61"/>\
        <w:lsdException w:name="Light Grid" w:uiPriority="62"/>\
        <w:lsdException w:name="Medium Shading 1" w:uiPriority="63"/>\
        <w:lsdException w:name="Medium Shading 2" w:uiPriority="64"/>\
        <w:lsdException w:name="Medium List 1" w:uiPriority="65"/>\
        <w:lsdException w:name="Medium List 2" w:uiPriority="66"/>\
        <w:lsdException w:name="Medium Grid 1" w:uiPriority="67"/>\
        <w:lsdException w:name="Medium Grid 2" w:uiPriority="68"/>\
        <w:lsdException w:name="Medium Grid 3" w:uiPriority="69"/>\
        <w:lsdException w:name="Dark List" w:uiPriority="70"/>\
        <w:lsdException w:name="Colorful Shading" w:uiPriority="71"/>\
        <w:lsdException w:name="Colorful List" w:uiPriority="72"/>\
        <w:lsdException w:name="Colorful Grid" w:uiPriority="73"/>\
        <w:lsdException w:name="Light Shading Accent 1" w:uiPriority="60"/>\
        <w:lsdException w:name="Light List Accent 1" w:uiPriority="61"/>\
        <w:lsdException w:name="Light Grid Accent 1" w:uiPriority="62"/>\
        <w:lsdException w:name="Medium Shading 1 Accent 1" w:uiPriority="63"/>\
        <w:lsdException w:name="Medium Shading 2 Accent 1" w:uiPriority="64"/>\
        <w:lsdException w:name="Medium List 1 Accent 1" w:uiPriority="65"/>\
        <w:lsdException w:name="Revision" w:semiHidden="1"/>\
        <w:lsdException w:name="List Paragraph" w:uiPriority="34" w:qFormat="1"/>\
        <w:lsdException w:name="Quote" w:uiPriority="29" w:qFormat="1"/>\
        <w:lsdException w:name="Intense Quote" w:uiPriority="30" w:qFormat="1"/>\
        <w:lsdException w:name="Medium List 2 Accent 1" w:uiPriority="66"/>\
        <w:lsdException w:name="Medium Grid 1 Accent 1" w:uiPriority="67"/>\
        <w:lsdException w:name="Medium Grid 2 Accent 1" w:uiPriority="68"/>\
        <w:lsdException w:name="Medium Grid 3 Accent 1" w:uiPriority="69"/>\
        <w:lsdException w:name="Dark List Accent 1" w:uiPriority="70"/>\
        <w:lsdException w:name="Colorful Shading Accent 1" w:uiPriority="71"/>\
        <w:lsdException w:name="Colorful List Accent 1" w:uiPriority="72"/>\
        <w:lsdException w:name="Colorful Grid Accent 1" w:uiPriority="73"/>\
        <w:lsdException w:name="Light Shading Accent 2" w:uiPriority="60"/>\
        <w:lsdException w:name="Light List Accent 2" w:uiPriority="61"/>\
        <w:lsdException w:name="Light Grid Accent 2" w:uiPriority="62"/>\
        <w:lsdException w:name="Medium Shading 1 Accent 2" w:uiPriority="63"/>\
        <w:lsdException w:name="Medium Shading 2 Accent 2" w:uiPriority="64"/>\
        <w:lsdException w:name="Medium List 1 Accent 2" w:uiPriority="65"/>\
        <w:lsdException w:name="Medium List 2 Accent 2" w:uiPriority="66"/>\
        <w:lsdException w:name="Medium Grid 1 Accent 2" w:uiPriority="67"/>\
        <w:lsdException w:name="Medium Grid 2 Accent 2" w:uiPriority="68"/>\
        <w:lsdException w:name="Medium Grid 3 Accent 2" w:uiPriority="69"/>\
        <w:lsdException w:name="Dark List Accent 2" w:uiPriority="70"/>\
        <w:lsdException w:name="Colorful Shading Accent 2" w:uiPriority="71"/>\
        <w:lsdException w:name="Colorful List Accent 2" w:uiPriority="72"/>\
        <w:lsdException w:name="Colorful Grid Accent 2" w:uiPriority="73"/>\
        <w:lsdException w:name="Light Shading Accent 3" w:uiPriority="60"/>\
        <w:lsdException w:name="Light List Accent 3" w:uiPriority="61"/>\
        <w:lsdException w:name="Light Grid Accent 3" w:uiPriority="62"/>\
        <w:lsdException w:name="Medium Shading 1 Accent 3" w:uiPriority="63"/>\
        <w:lsdException w:name="Medium Shading 2 Accent 3" w:uiPriority="64"/>\
        <w:lsdException w:name="Medium List 1 Accent 3" w:uiPriority="65"/>\
        <w:lsdException w:name="Medium List 2 Accent 3" w:uiPriority="66"/>\
        <w:lsdException w:name="Medium Grid 1 Accent 3" w:uiPriority="67"/>\
        <w:lsdException w:name="Medium Grid 2 Accent 3" w:uiPriority="68"/>\
        <w:lsdException w:name="Medium Grid 3 Accent 3" w:uiPriority="69"/>\
        <w:lsdException w:name="Dark List Accent 3" w:uiPriority="70"/>\
        <w:lsdException w:name="Colorful Shading Accent 3" w:uiPriority="71"/>\
        <w:lsdException w:name="Colorful List Accent 3" w:uiPriority="72"/>\
        <w:lsdException w:name="Colorful Grid Accent 3" w:uiPriority="73"/>\
        <w:lsdException w:name="Light Shading Accent 4" w:uiPriority="60"/>\
        <w:lsdException w:name="Light List Accent 4" w:uiPriority="61"/>\
        <w:lsdException w:name="Light Grid Accent 4" w:uiPriority="62"/>\
        <w:lsdException w:name="Medium Shading 1 Accent 4" w:uiPriority="63"/>\
        <w:lsdException w:name="Medium Shading 2 Accent 4" w:uiPriority="64"/>\
        <w:lsdException w:name="Medium List 1 Accent 4" w:uiPriority="65"/>\
        <w:lsdException w:name="Medium List 2 Accent 4" w:uiPriority="66"/>\
        <w:lsdException w:name="Medium Grid 1 Accent 4" w:uiPriority="67"/>\
        <w:lsdException w:name="Medium Grid 2 Accent 4" w:uiPriority="68"/>\
        <w:lsdException w:name="Medium Grid 3 Accent 4" w:uiPriority="69"/>\
        <w:lsdException w:name="Dark List Accent 4" w:uiPriority="70"/>\
        <w:lsdException w:name="Colorful Shading Accent 4" w:uiPriority="71"/>\
        <w:lsdException w:name="Colorful List Accent 4" w:uiPriority="72"/>\
        <w:lsdException w:name="Colorful Grid Accent 4" w:uiPriority="73"/>\
        <w:lsdException w:name="Light Shading Accent 5" w:uiPriority="60"/>\
        <w:lsdException w:name="Light List Accent 5" w:uiPriority="61"/>\
        <w:lsdException w:name="Light Grid Accent 5" w:uiPriority="62"/>\
        <w:lsdException w:name="Medium Shading 1 Accent 5" w:uiPriority="63"/>\
        <w:lsdException w:name="Medium Shading 2 Accent 5" w:uiPriority="64"/>\
        <w:lsdException w:name="Medium List 1 Accent 5" w:uiPriority="65"/>\
        <w:lsdException w:name="Medium List 2 Accent 5" w:uiPriority="66"/>\
        <w:lsdException w:name="Medium Grid 1 Accent 5" w:uiPriority="67"/>\
        <w:lsdException w:name="Medium Grid 2 Accent 5" w:uiPriority="68"/>\
        <w:lsdException w:name="Medium Grid 3 Accent 5" w:uiPriority="69"/>\
        <w:lsdException w:name="Dark List Accent 5" w:uiPriority="70"/>\
        <w:lsdException w:name="Colorful Shading Accent 5" w:uiPriority="71"/>\
        <w:lsdException w:name="Colorful List Accent 5" w:uiPriority="72"/>\
        <w:lsdException w:name="Colorful Grid Accent 5" w:uiPriority="73"/>\
        <w:lsdException w:name="Light Shading Accent 6" w:uiPriority="60"/>\
        <w:lsdException w:name="Light List Accent 6" w:uiPriority="61"/>\
        <w:lsdException w:name="Light Grid Accent 6" w:uiPriority="62"/>\
        <w:lsdException w:name="Medium Shading 1 Accent 6" w:uiPriority="63"/>\
        <w:lsdException w:name="Medium Shading 2 Accent 6" w:uiPriority="64"/>\
        <w:lsdException w:name="Medium List 1 Accent 6" w:uiPriority="65"/>\
        <w:lsdException w:name="Medium List 2 Accent 6" w:uiPriority="66"/>\
        <w:lsdException w:name="Medium Grid 1 Accent 6" w:uiPriority="67"/>\
        <w:lsdException w:name="Medium Grid 2 Accent 6" w:uiPriority="68"/>\
        <w:lsdException w:name="Medium Grid 3 Accent 6" w:uiPriority="69"/>\
        <w:lsdException w:name="Dark List Accent 6" w:uiPriority="70"/>\
        <w:lsdException w:name="Colorful Shading Accent 6" w:uiPriority="71"/>\
        <w:lsdException w:name="Colorful List Accent 6" w:uiPriority="72"/>\
        <w:lsdException w:name="Colorful Grid Accent 6" w:uiPriority="73"/>\
        <w:lsdException w:name="Subtle Emphasis" w:uiPriority="19" w:qFormat="1"/>\
        <w:lsdException w:name="Intense Emphasis" w:uiPriority="21" w:qFormat="1"/>\
        <w:lsdException w:name="Subtle Reference" w:uiPriority="31" w:qFormat="1"/>\
        <w:lsdException w:name="Intense Reference" w:uiPriority="32" w:qFormat="1"/>\
        <w:lsdException w:name="Book Title" w:uiPriority="33" w:qFormat="1"/>\
        <w:lsdException w:name="Bibliography" w:semiHidden="1" w:uiPriority="37" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="TOC Heading" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1" w:qFormat="1"/>\
        <w:lsdException w:name="Plain Table 1" w:uiPriority="41"/>\
        <w:lsdException w:name="Plain Table 2" w:uiPriority="42"/>\
        <w:lsdException w:name="Plain Table 3" w:uiPriority="43"/>\
        <w:lsdException w:name="Plain Table 4" w:uiPriority="44"/>\
        <w:lsdException w:name="Plain Table 5" w:uiPriority="45"/>\
        <w:lsdException w:name="Grid Table Light" w:uiPriority="40"/>\
        <w:lsdException w:name="Grid Table 1 Light" w:uiPriority="46"/>\
        <w:lsdException w:name="Grid Table 2" w:uiPriority="47"/>\
        <w:lsdException w:name="Grid Table 3" w:uiPriority="48"/>\
        <w:lsdException w:name="Grid Table 4" w:uiPriority="49"/>\
        <w:lsdException w:name="Grid Table 5 Dark" w:uiPriority="50"/>\
        <w:lsdException w:name="Grid Table 6 Colorful" w:uiPriority="51"/>\
        <w:lsdException w:name="Grid Table 7 Colorful" w:uiPriority="52"/>\
        <w:lsdException w:name="Grid Table 1 Light Accent 1" w:uiPriority="46"/>\
        <w:lsdException w:name="Grid Table 2 Accent 1" w:uiPriority="47"/>\
        <w:lsdException w:name="Grid Table 3 Accent 1" w:uiPriority="48"/>\
        <w:lsdException w:name="Grid Table 4 Accent 1" w:uiPriority="49"/>\
        <w:lsdException w:name="Grid Table 5 Dark Accent 1" w:uiPriority="50"/>\
        <w:lsdException w:name="Grid Table 6 Colorful Accent 1" w:uiPriority="51"/>\
        <w:lsdException w:name="Grid Table 7 Colorful Accent 1" w:uiPriority="52"/>\
        <w:lsdException w:name="Grid Table 1 Light Accent 2" w:uiPriority="46"/>\
        <w:lsdException w:name="Grid Table 2 Accent 2" w:uiPriority="47"/>\
        <w:lsdException w:name="Grid Table 3 Accent 2" w:uiPriority="48"/>\
        <w:lsdException w:name="Grid Table 4 Accent 2" w:uiPriority="49"/>\
        <w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>\
        <w:lsdException w:name="Grid Table 6 Colorful Accent 2" w:uiPriority="51"/>\
        <w:lsdException w:name="Grid Table 7 Colorful Accent 2" w:uiPriority="52"/>\
        <w:lsdException w:name="Grid Table 1 Light Accent 3" w:uiPriority="46"/>\
        <w:lsdException w:name="Grid Table 2 Accent 3" w:uiPriority="47"/>\
        <w:lsdException w:name="Grid Table 3 Accent 3" w:uiPriority="48"/>\
        <w:lsdException w:name="Grid Table 4 Accent 3" w:uiPriority="49"/>\
        <w:lsdException w:name="Grid Table 5 Dark Accent 3" w:uiPriority="50"/>\
        <w:lsdException w:name="Grid Table 6 Colorful Accent 3" w:uiPriority="51"/>\
        <w:lsdException w:name="Grid Table 7 Colorful Accent 3" w:uiPriority="52"/>\
        <w:lsdException w:name="Grid Table 1 Light Accent 4" w:uiPriority="46"/>\
        <w:lsdException w:name="Grid Table 2 Accent 4" w:uiPriority="47"/>\
        <w:lsdException w:name="Grid Table 3 Accent 4" w:uiPriority="48"/>\
        <w:lsdException w:name="Grid Table 4 Accent 4" w:uiPriority="49"/>\
        <w:lsdException w:name="Grid Table 5 Dark Accent 4" w:uiPriority="50"/>\
        <w:lsdException w:name="Grid Table 6 Colorful Accent 4" w:uiPriority="51"/>\
        <w:lsdException w:name="Grid Table 7 Colorful Accent 4" w:uiPriority="52"/>\
        <w:lsdException w:name="Grid Table 1 Light Accent 5" w:uiPriority="46"/>\
        <w:lsdException w:name="Grid Table 2 Accent 5" w:uiPriority="47"/>\
        <w:lsdException w:name="Grid Table 3 Accent 5" w:uiPriority="48"/>\
        <w:lsdException w:name="Grid Table 4 Accent 5" w:uiPriority="49"/>\
        <w:lsdException w:name="Grid Table 5 Dark Accent 5" w:uiPriority="50"/>\
        <w:lsdException w:name="Grid Table 6 Colorful Accent 5" w:uiPriority="51"/>\
        <w:lsdException w:name="Grid Table 7 Colorful Accent 5" w:uiPriority="52"/>\
        <w:lsdException w:name="Grid Table 1 Light Accent 6" w:uiPriority="46"/>\
        <w:lsdException w:name="Grid Table 2 Accent 6" w:uiPriority="47"/>\
        <w:lsdException w:name="Grid Table 3 Accent 6" w:uiPriority="48"/>\
        <w:lsdException w:name="Grid Table 4 Accent 6" w:uiPriority="49"/>\
        <w:lsdException w:name="Grid Table 5 Dark Accent 6" w:uiPriority="50"/>\
        <w:lsdException w:name="Grid Table 6 Colorful Accent 6" w:uiPriority="51"/>\
        <w:lsdException w:name="Grid Table 7 Colorful Accent 6" w:uiPriority="52"/>\
        <w:lsdException w:name="List Table 1 Light" w:uiPriority="46"/>\
        <w:lsdException w:name="List Table 2" w:uiPriority="47"/>\
        <w:lsdException w:name="List Table 3" w:uiPriority="48"/>\
        <w:lsdException w:name="List Table 4" w:uiPriority="49"/>\
        <w:lsdException w:name="List Table 5 Dark" w:uiPriority="50"/>\
        <w:lsdException w:name="List Table 6 Colorful" w:uiPriority="51"/>\
        <w:lsdException w:name="List Table 7 Colorful" w:uiPriority="52"/>\
        <w:lsdException w:name="List Table 1 Light Accent 1" w:uiPriority="46"/>\
        <w:lsdException w:name="List Table 2 Accent 1" w:uiPriority="47"/>\
        <w:lsdException w:name="List Table 3 Accent 1" w:uiPriority="48"/>\
        <w:lsdException w:name="List Table 4 Accent 1" w:uiPriority="49"/>\
        <w:lsdException w:name="List Table 5 Dark Accent 1" w:uiPriority="50"/>\
        <w:lsdException w:name="List Table 6 Colorful Accent 1" w:uiPriority="51"/>\
        <w:lsdException w:name="List Table 7 Colorful Accent 1" w:uiPriority="52"/>\
        <w:lsdException w:name="List Table 1 Light Accent 2" w:uiPriority="46"/>\
        <w:lsdException w:name="List Table 2 Accent 2" w:uiPriority="47"/>\
        <w:lsdException w:name="List Table 3 Accent 2" w:uiPriority="48"/>\
        <w:lsdException w:name="List Table 4 Accent 2" w:uiPriority="49"/>\
        <w:lsdException w:name="List Table 5 Dark Accent 2" w:uiPriority="50"/>\
        <w:lsdException w:name="List Table 6 Colorful Accent 2" w:uiPriority="51"/>\
        <w:lsdException w:name="List Table 7 Colorful Accent 2" w:uiPriority="52"/>\
        <w:lsdException w:name="List Table 1 Light Accent 3" w:uiPriority="46"/>\
        <w:lsdException w:name="List Table 2 Accent 3" w:uiPriority="47"/>\
        <w:lsdException w:name="List Table 3 Accent 3" w:uiPriority="48"/>\
        <w:lsdException w:name="List Table 4 Accent 3" w:uiPriority="49"/>\
        <w:lsdException w:name="List Table 5 Dark Accent 3" w:uiPriority="50"/>\
        <w:lsdException w:name="List Table 6 Colorful Accent 3" w:uiPriority="51"/>\
        <w:lsdException w:name="List Table 7 Colorful Accent 3" w:uiPriority="52"/>\
        <w:lsdException w:name="List Table 1 Light Accent 4" w:uiPriority="46"/>\
        <w:lsdException w:name="List Table 2 Accent 4" w:uiPriority="47"/>\
        <w:lsdException w:name="List Table 3 Accent 4" w:uiPriority="48"/>\
        <w:lsdException w:name="List Table 4 Accent 4" w:uiPriority="49"/>\
        <w:lsdException w:name="List Table 5 Dark Accent 4" w:uiPriority="50"/>\
        <w:lsdException w:name="List Table 6 Colorful Accent 4" w:uiPriority="51"/>\
        <w:lsdException w:name="List Table 7 Colorful Accent 4" w:uiPriority="52"/>\
        <w:lsdException w:name="List Table 1 Light Accent 5" w:uiPriority="46"/>\
        <w:lsdException w:name="List Table 2 Accent 5" w:uiPriority="47"/>\
        <w:lsdException w:name="List Table 3 Accent 5" w:uiPriority="48"/>\
        <w:lsdException w:name="List Table 4 Accent 5" w:uiPriority="49"/>\
        <w:lsdException w:name="List Table 5 Dark Accent 5" w:uiPriority="50"/>\
        <w:lsdException w:name="List Table 6 Colorful Accent 5" w:uiPriority="51"/>\
        <w:lsdException w:name="List Table 7 Colorful Accent 5" w:uiPriority="52"/>\
        <w:lsdException w:name="List Table 1 Light Accent 6" w:uiPriority="46"/>\
        <w:lsdException w:name="List Table 2 Accent 6" w:uiPriority="47"/>\
        <w:lsdException w:name="List Table 3 Accent 6" w:uiPriority="48"/>\
        <w:lsdException w:name="List Table 4 Accent 6" w:uiPriority="49"/>\
        <w:lsdException w:name="List Table 5 Dark Accent 6" w:uiPriority="50"/>\
        <w:lsdException w:name="List Table 6 Colorful Accent 6" w:uiPriority="51"/>\
        <w:lsdException w:name="List Table 7 Colorful Accent 6" w:uiPriority="52"/>\
        <w:lsdException w:name="Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Smart Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Hashtag" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Unresolved Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        <w:lsdException w:name="Smart Link" w:semiHidden="1" w:unhideWhenUsed="1"/>\
        </w:latentStyles>\
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>\
        <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="Heading1Char"/><w:uiPriority w:val="9"/><w:qFormat/><w:rsid w:val="00120ED4"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>\
        <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>\
        <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>\
        <w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>\
        <w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char"><w:name w:val="Heading 1 Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Heading1"/><w:uiPriority w:val="9"/><w:rsid w:val="00120ED4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>\
        <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="TitleChar"/><w:uiPriority w:val="10"/><w:qFormat/><w:rsid w:val="00120ED4"/><w:pPr><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>\
        <w:style w:type="character" w:customStyle="1" w:styleId="TitleChar"><w:name w:val="Title Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Title"/><w:uiPriority w:val="10"/><w:rsid w:val="00120ED4"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>\
        </w:styles>
        """
     */
    
    let theme1XML = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">\
    <a:themeElements>\
    <a:clrScheme name="Office">\
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>\
    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\
    <a:dk2><a:srgbClr val="44546A"/></a:dk2>\
    <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>\
    <a:accent1><a:srgbClr val="4472C4"/></a:accent1>\
    <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>\
    <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>\
    <a:accent4><a:srgbClr val="FFC000"/></a:accent4>\
    <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>\
    <a:accent6><a:srgbClr val="70AD47"/></a:accent6>\
    <a:hlink><a:srgbClr val="0563C1"/></a:hlink>\
    <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>\
    </a:clrScheme>\
    <a:fontScheme name="Office">
    <a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont>\
    <a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游明朝"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont>\
    </a:fontScheme>\
    <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>\
    </a:themeElements>\
    <a:objectDefaults/>\
    <a:extraClrSchemeLst/>\
    <a:extLst>\
    <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">\
    <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>\
    </a:ext>\
    </a:extLst>\
    </a:theme>
    """
    
    let tempDoc = """
    <?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" mc:Ignorable="w14">
    <w:body>\
    <w:p><w:pPr><w:pStyle w:val="Heading"/><w:bidi w:val="0"/></w:pPr><w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:t>Heading</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Body"/><w:bidi w:val="0"/></w:pPr><w:r><w:rPr><w:b w:val="1"/><w:bCs w:val="1"/><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:i w:val="1"/><w:iCs w:val="1"/><w:rtl w:val="0"/><w:lang w:val="en-US"/></w:rPr><w:t>Italic</w:t></w:r></w:p>\
    <w:p><w:pPr><w:pStyle w:val="Body"/><w:bidi w:val="0"/></w:pPr><w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:t>Testing 123</w:t></w:r></w:p>\
    <w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId5"/><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="864"/><w:bidi w:val="0"/></w:sectPr>\
    </w:body>\
    </w:document>
    """
    
    public init(source: String) {
        let document = Document(parsing: source, options: [.parseBlockDirectives])
        print(document.debugDescription()) //
        var generator = OOXMLGenerator()
        generator.visit(document)
        print(generator.output) //
        content = generator.output
    }
    
    private func addFile(file: String, content fileContent: String, toArchive archive: Archive) {
        guard let data = fileContent.data(using: .utf8) else { return }
        try? archive.addEntry(with: file, type: .file, uncompressedSize: Int64(data.count), modificationDate: Date(), permissions: nil, compressionMethod: .deflate, bufferSize: defaultWriteChunkSize, progress: nil) {
            (position, size) -> Data in
            // This will be called until `data` is exhausted.
            return data.subdata(in: Int(position) ..< Int(position)+size)
        }
    }
    
    public func save(to outputFileURL: URL) {
        guard let archive = Archive(url: outputFileURL, accessMode: .create) else  {
            print("Unable to save file.") //
            return
        }
        
        addFile(file: "[Content_Types].xml", content: contentTypesXML, toArchive: archive)
        addFile(file: "_rels/.rels", content: relsXML, toArchive: archive)
        addFile(file: "docProps/app.xml", content: docPropsAppXML, toArchive: archive)
        addFile(file: "docProps/core.xml", content: docPropsCoreXML, toArchive: archive)
        addFile(file: "word/document.xml", content: content, toArchive: archive) // content or tempDoc
        addFile(file: "word/header1.xml", content: header1XML, toArchive: archive)
        addFile(file: "word/footer1.xml", content: footer1XML, toArchive: archive)
        addFile(file: "word/_rels/document.xml.rels", content: contentRelsXML, toArchive: archive)
        addFile(file: "word/fontTable.xml", content: fontTableXML, toArchive: archive)
        addFile(file: "word/settings.xml", content: settingsXML, toArchive: archive)
        addFile(file: "word/webSettings.xml", content: webSettingsXML, toArchive: archive)
        addFile(file: "word/styles.xml", content: stylesXML, toArchive: archive)
        //addFile(file: "word/theme/theme1.xml", content: theme1XML, toArchive: archive)
    }
}
