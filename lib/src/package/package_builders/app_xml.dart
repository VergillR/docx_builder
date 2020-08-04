/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/utils/constants/constants.dart';

class AppXml {
  // final int chars;
  // final int charsWithSpaces;
  // final int paragraphs;

  // AppXml({this.chars = 0, this.charsWithSpaces = 0, this.paragraphs = 0});

  AppXml();

  String getAppXml() => docPropsAppXml;

  // String getAppXml() => docPropsAppXml.replaceAll('</Properties>',
  //     '<Characters>$chars</Characters><CharactersWithSpaces>$charsWithSpaces</CharactersWithSpaces><Paragraphs>$paragraphs</Paragraphs></Properties>');
}
