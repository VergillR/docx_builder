/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

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
