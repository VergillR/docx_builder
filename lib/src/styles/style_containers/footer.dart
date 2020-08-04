/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/styles/style_enums.dart';

abstract class Footer {
  String rId;
  final FooterType footerType;
  String taggedText;

  Footer({this.rId, this.footerType, this.taggedText});
}

class EvenPageFooter extends Footer {
  EvenPageFooter({String taggedText, String rId})
      : super(
            footerType: FooterType.evenPage, taggedText: taggedText, rId: rId);
}

/// Also called default Footer.
class OddPageFooter extends Footer {
  OddPageFooter({String taggedText, String rId})
      : super(footerType: FooterType.oddPage, taggedText: taggedText, rId: rId);
}

/// Also called titlePageFooter.
class FirstPageFooter extends Footer {
  FirstPageFooter({String taggedText, String rId})
      : super(
            footerType: FooterType.firstPage, taggedText: taggedText, rId: rId);
}
