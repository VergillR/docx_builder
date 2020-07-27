import 'package:docx_builder/src/utils/constants/constants.dart';

class DocumentXmlRels {
  DocumentXmlRels();

  String getDocumentXmlRels(Map<String, String> references) {
    const String closeTag = '</Relationships>';
    if (references.isNotEmpty) {
      final StringBuffer mediaRels = StringBuffer();
      references.forEach((k, v) {
        mediaRels.write('<Relationship Id="$k" $v/>');
      });
      return '$wordRelsDocumentXmlRels${mediaRels.toString()}$closeTag';
    } else {
      return '$wordRelsDocumentXmlRels$closeTag';
    }
  }
}
