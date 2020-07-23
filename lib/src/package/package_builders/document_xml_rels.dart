import 'package:docx_builder/src/utils/constants/constants.dart';

class DocumentXmlRels {
  DocumentXmlRels();

  // TODO: Add more references when needed, for example: media and hyperlinks.
  /* For example:
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.png"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image4.png"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image5.png"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image6.png"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  */
  String getDocumentXmlRels(Map<String, String> references) {
    const String closeTag = '</Relationships>';
    if (references.isNotEmpty) {
      final StringBuffer mediaRels = StringBuffer();
      references.forEach((k, v) {
        mediaRels.write(
            '<Relationship Id="$k" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="$v"/>');
      });
      return '$wordRelsDocumentXmlRels${mediaRels.toString()}$closeTag';
    } else {
      return '$wordRelsDocumentXmlRels$closeTag';
    }
  }
}
