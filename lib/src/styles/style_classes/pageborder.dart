/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class PageBorder extends Style {
  final PageBorderSide pageBorderSide;
  final int size;
  final int space;
  final String color;
  final ParagraphBorderStyle pbrStyle;
  final bool shadow;

  PageBorder({
    this.pageBorderSide = PageBorderSide.bottom,
    this.size = 8,
    this.space = 24,
    this.color = '000000',
    this.pbrStyle = ParagraphBorderStyle.single,
    this.shadow = false,
  });

  @override
  String getXml() {
    final String borderColor = isValidColor(color) ? color : '000000';
    return '<w:${getValueFromEnum(pageBorderSide)} w:val="${getValueFromEnum(pbrStyle)}" w:color="$borderColor" w:sz="$size" w:space="$space" w:shadow="$shadow" />';
  }
}
