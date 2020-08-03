import 'package:docx_builder/src/utils/constants/constants.dart';

class WordStylesXml {
  WordStylesXml();

  String getWordStylesXml({String hyperlinkStylingXml}) {
    final String style = hyperlinkStylingXml?.isNotEmpty ?? false
        ? hyperlinkStylingXml
        : '<w:rPr><w:color w:val="000080"/><w:u w:val="single"/></w:rPr>';
    const String openTag =
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>';
    const String closeTag = '</w:style></w:styles>';
    return '$wordStylesXml$openTag$style$closeTag';
  }
}
