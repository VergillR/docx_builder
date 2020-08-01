import 'package:docx_builder/src/utils/constants/constants.dart';

class Numbering {
  Numbering();

  String getNumberingXml({String customNumberingXml = ''}) {
    final String customTag = customNumberingXml.isNotEmpty
        ? '<w:num w:numId="3"><w:abstractNumId w:val="92"/></w:num>'
        : '';
    final String closeTag =
        '<w:num w:numId="1"><w:abstractNumId w:val="90"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="91"/></w:num>$customTag</w:numbering>';
    return '$numberingXml$customNumberingXml$closeTag';
  }
}
