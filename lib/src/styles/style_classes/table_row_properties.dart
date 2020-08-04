/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class TableRowProperties extends Style {
  final bool cantSplit;
  final bool hidden;
  final TableTextAlignment textAlignment;
  final String tableCellSpacing;
  final PreferredWidthType tableCellSpacingWidthType;
  final bool isTableHeader;
  final HRule rowHeightSetting;
  final int rowHeight;

  TableRowProperties({
    this.cantSplit,
    this.hidden,
    this.textAlignment,
    this.tableCellSpacing,
    this.tableCellSpacingWidthType,
    this.isTableHeader,
    this.rowHeightSetting,
    this.rowHeight,
  });

  @override
  String getXml() {
    final StringBuffer trp = StringBuffer()..write('<w:trPr>');
    if (cantSplit ?? false) {
      trp.write('<w:cantSplit/>');
    }
    if (hidden ?? false) {
      trp.write('<w:hidden/>');
    }
    if (textAlignment != null) {
      trp.write('<w:jc w:val="${getValueFromEnum(textAlignment)}"/>');
    }
    if (tableCellSpacing != null || tableCellSpacingWidthType != null) {
      trp.write(
          '<w:tblCellSpacing w:w="${tableCellSpacing ?? 0}" w:type="${getValueFromEnum(tableCellSpacingWidthType ?? PreferredWidthType.nil)}"/>');
    }
    if (isTableHeader ?? false) {
      trp.write('<w:tblHeader/>');
    }

    if (rowHeight != null) {
      trp.write(
          '<w:trHeight w:hRule="${getValueFromEnum(rowHeightSetting ?? HRule.auto)}" w: val="$rowHeight"/>');
    }

    trp.write('</w:trPr>');
    return trp.toString();
  }
}
