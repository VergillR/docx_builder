/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import 'package:docx_builder/src/styles/style_enums.dart';
import '../../utils/utils.dart';
import '../style.dart';

class TableBorder extends Style {
  final TableBorderSide borderSide;
  final int width;
  final int space;
  final String color;
  final ParagraphBorderStyle pbrStyle;
  final bool shadow;

  TableBorder({
    this.borderSide = TableBorderSide.bottom,
    this.width = 24,
    this.space = 1,
    this.color = "000000",
    this.pbrStyle = ParagraphBorderStyle.single,
    this.shadow = false,
  });

  @override
  String getXml() {
    final String borderColor = isValidColor(color) ? color : '000000';
    return '<w:${getValueFromEnum(borderSide)} w:val="${getValueFromEnum(pbrStyle)}" w:sz="$width" w:space="$space" w:color="$borderColor" w:shadow="$shadow" />';
  }
}
