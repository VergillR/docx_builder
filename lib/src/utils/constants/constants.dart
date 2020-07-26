import 'package:docx_builder/src/utils/utils.dart';

const String version = '1.0';

/// For example: timestamp is 2020-07-04T22:32:56Z
final String timestamp = getUtcTimestampFromDateTime(DateTime.now());

const String blankPageA4 =
    '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11907" w:h="16839" w:orient="portrait" /></w:sectPr></w:pPr></w:p>';

const String dotRels =
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';

const String contentTypes =
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/_rels/document.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';

const String startDoc =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14"><w:body>';

const String docPropsAppXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template></Template><TotalTime>1</TotalTime><Application>NovaPrisma DocXBuilder Flutter $version</Application><Company></Company><AppVersion>$version</AppVersion></Properties>';

final String docPropsCoreXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">$timestamp</dcterms:created><dc:creator></dc:creator><dc:description></dc:description><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:modified xsi:type="dcterms:W3CDTF">$timestamp</dcterms:modified><cp:revision>1</cp:revision><dc:subject></dc:subject><dc:title></dc:title></cp:coreProperties>';

// close with e.g. '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Cello Sans" w:hAnsi="Cello Sans"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>This is a fancy footer or voetnoot.</w:t></w:r></w:p></w:ftr>'
const String footerXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14">';

// close with '<w:p><w:pPr><w:pStyle w:val="Header"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr></w:r></w:p></w:hdr>'
const String headerXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14">';

const String wordRelsDocumentXmlRels =
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';

const String wordFontTableXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="Times New Roman"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font><w:font w:name="Symbol"><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font><w:font w:name="Arial"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font></w:fonts>';

const String wordSettingsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:displayBackgroundShape/><w:defaultTabStop w:val="710"/><w:compat></w:compat><w:themeFontLang w:val="" w:eastAsia="" w:bidi=""/></w:settings>';

const String wordStylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="WenQuanYi Micro Hei" w:cs="Arial"/><w:kern w:val="2"/><w:sz w:val="36"/><w:szCs w:val="36"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="hi-IN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="000080"/><w:u w:val="single"/><w:lang w:val="zxx" w:eastAsia="zxx" w:bidi="zxx"/></w:rPr></w:style></w:styles>';
