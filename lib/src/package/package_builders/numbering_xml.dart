/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/utils/constants/constants.dart';

class NumberingXml {
  NumberingXml();

  String getNumberingXml({String customNumberingXml = ''}) {
    final String customTag = customNumberingXml.isNotEmpty
        ? '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
        : '';
    final String closeTag =
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>$customTag</w:numbering>';
    return '$numberingXml$customNumberingXml$closeTag';
  }
}
