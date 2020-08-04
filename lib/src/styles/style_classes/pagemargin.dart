/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../style.dart';

class PageMargin extends Style {
  final int bottom;
  final int footer;
  final int gutter;
  final int header;
  final int left;
  final int right;
  final int top;

  PageMargin({
    this.bottom = 1417,
    this.footer = 708,
    this.gutter = 0,
    this.header = 708,
    this.left = 1417,
    this.right = 1417,
    this.top = 1417,
  });

  @override
  String getXml() =>
      'w:pgMar w:header="$header" w:footer="$footer" w:gutter="$gutter" w:left="$left" w:right="$right" w:top="$top" w:bottom="$bottom" />';
}
