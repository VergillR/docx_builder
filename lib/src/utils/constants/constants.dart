import 'package:docx_builder/src/utils/utils.dart';

const String version = '1.0';

/// For example: timestamp is 2020-07-04T22:32:56Z
final String timestamp = getUtcTimestampFromDateTime(DateTime.now());

const String blankPageA4 =
    '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16839" w:orient="portrait" /></w:sectPr></w:pPr></w:p>';

const String dotRels =
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';

const String commentsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">';

const String reservedTarget = '%RESERVED%';

/// Replace %RESERVED% with the generated content types when the document is created.
const String contentTypes =
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>$reservedTarget<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/_rels/document.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';

const String startDoc =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:body>';

const String docPropsAppXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template></Template><TotalTime>0</TotalTime><Application>NovaPrisma DocXBuilder by Vergill Lemmert $version</Application><Company></Company><AppVersion>$version</AppVersion></Properties>';

final String docPropsCoreXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">$timestamp</dcterms:created><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:modified xsi:type="dcterms:W3CDTF">$timestamp</dcterms:modified><cp:revision>1</cp:revision>';

const String footerXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">';

const String headerXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">';

String numberingXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs><w:ind w:left="420" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="1" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="◦"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="840"/></w:tabs><w:ind w:left="840" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="2" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="▪"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="1260"/></w:tabs><w:ind w:left="1260" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="3" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="▫"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="1680"/></w:tabs><w:ind w:left="1680" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="4" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="2100"/></w:tabs><w:ind w:left="2100" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="5" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="◦"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="2520"/></w:tabs><w:ind w:left="2520" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="6" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="▪"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="2940"/></w:tabs><w:ind w:left="2940" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="7" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="▫"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="3360"/></w:tabs><w:ind w:left="3360" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl><w:lvl w:ilvl="8" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="3780"/></w:tabs><w:ind w:left="3780" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="425"/></w:tabs><w:ind w:left="425" w:leftChars="0" w:hanging="425" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="1" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="840"/></w:tabs><w:ind w:left="840" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="2" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="1260"/></w:tabs><w:ind w:left="1260" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="3" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%4."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="1680"/></w:tabs><w:ind w:left="1680" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="4" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%5)"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="2100"/></w:tabs><w:ind w:left="2100" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="5" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%6."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="2520"/></w:tabs><w:ind w:left="2520" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="6" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%7."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="2940"/></w:tabs><w:ind w:left="2940" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="7" w:tentative="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%8)"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="3360"/></w:tabs><w:ind w:left="3360" w:leftChars="0" w:hanging="420" w:firstLineChars="0"/></w:pPr><w:rPr><w:rFonts w:hint="default"/></w:rPr></w:lvl></w:abstractNum>';

const String wordRelsDocumentXmlRels =
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';

const String wordFontTableXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="Times New Roman"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font><w:font w:name="Symbol"><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font><w:font w:name="Arial"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font></w:fonts>';

const String wordSettingsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14"><w:zoom w:percent="100"/><w:displayBackgroundShape w:val="1"/><w:documentProtection w:enforcement="0"/><w:defaultTabStop w:val="710"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/></w:compat><w:themeFontLang w:val="" w:eastAsia="" w:bidi=""/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/></w:settings>';

const String wordStylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="WenQuanYi Micro Hei" w:cs="Arial"/><w:kern w:val="2"/><w:sz w:val="36"/><w:szCs w:val="36"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="hi-IN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr></w:pPr></w:pPrDefault></w:docDefaults>';
