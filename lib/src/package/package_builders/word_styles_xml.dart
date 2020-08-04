/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/utils/constants/constants.dart';

class WordStylesXml {
  WordStylesXml();

  String getWordStylesXml({String hyperlinkStylingXml = ''}) {
    final String style = hyperlinkStylingXml.isNotEmpty
        ? hyperlinkStylingXml
        : '<w:rPr><w:color w:val="000080"/><w:u w:val="single"/></w:rPr>';
    const String openTag =
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>';
    const String closeTag = '</w:style></w:styles>';
    return '$wordStylesXml$openTag$style$closeTag';
  }
}
