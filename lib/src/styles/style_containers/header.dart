/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/styles/style_enums.dart';

abstract class Header {
  String rId;
  final HeaderType headerType;
  String taggedText;

  Header({
    this.rId,
    this.headerType,
    this.taggedText,
  });
}

class EvenPageHeader extends Header {
  EvenPageHeader({
    String taggedText,
    String rId,
  }) : super(headerType: HeaderType.evenPage, taggedText: taggedText, rId: rId);
}

/// Also called default header.
class OddPageHeader extends Header {
  OddPageHeader({String taggedText, String rId})
      : super(headerType: HeaderType.oddPage, taggedText: taggedText, rId: rId);
}

/// Also called titlePageHeader.
class FirstPageHeader extends Header {
  FirstPageHeader({String taggedText, String rId})
      : super(
            headerType: HeaderType.firstPage, taggedText: taggedText, rId: rId);
}
