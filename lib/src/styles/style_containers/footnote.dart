/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../style_classes/index.dart';
import '../style_enums.dart';
import './index.dart';

class Footnote {
  final String id;
  final List<String> text;
  final List<TextStyle> textStyles;
  final bool addTab;
  final List<ComplexField> complexFields;
  final TextAlignment textAlignment;
  final String setBookmarkName;
  final NumberingList numberingList;
  final int numberLevelInList;

  Footnote({
    this.id,
    this.text,
    this.textStyles,
    this.addTab = true,
    this.complexFields,
    this.textAlignment,
    this.setBookmarkName,
    this.numberingList,
    this.numberLevelInList,
  });
}
