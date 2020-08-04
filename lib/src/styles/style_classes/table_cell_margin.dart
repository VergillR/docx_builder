/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class TableCellMargin extends Style {
  final TableCellSide tableCellSide;
  final int marginWidth;
  final bool isNil;

  TableCellMargin({
    this.tableCellSide = TableCellSide.bottom,
    this.marginWidth = 0,
    this.isNil = false,
  });

  @override
  String getXml() =>
      '<w:${getValueFromEnum(tableCellSide)} w:w="$marginWidth" w:type="${isNil ? "nil" : "dxa"}"/>';
}
