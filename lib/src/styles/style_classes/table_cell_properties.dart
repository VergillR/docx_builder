/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';
import 'index.dart';

class TableCellProperties extends Style {
  final int gridSpan;
  final bool hideMark;
  final bool noWrap;
  final Shading shading;
  final List<TableCellBorder> tableCellBorders;
  final bool fitText;
  final List<TableCellMargin> tableCellMargins;
  final String preferredWidth;
  final PreferredWidthType preferredWidthType;
  final TableCellVerticalAlignment verticalAlignment;
  final bool restartVMerge;
  final bool continueVMerge;

  TableCellProperties({
    this.gridSpan,
    this.hideMark,
    this.noWrap,
    this.shading,
    this.tableCellBorders,
    this.fitText,
    this.tableCellMargins,
    this.preferredWidth,
    this.preferredWidthType,
    this.verticalAlignment,
    this.restartVMerge,
    this.continueVMerge,
  });

  @override
  String getXml() {
    final StringBuffer tcp = StringBuffer()..write('<w:tcPr>');

    if (gridSpan != null) {
      tcp.write('<w:gridSpan w:val="$gridSpan"/>');
    }
    if (hideMark ?? false) {
      tcp.write('<w:hideMark/>');
    }
    if (noWrap ?? false) {
      tcp.write('<w:noWrap/>');
    }
    if (shading != null) {
      tcp.write(shading.getXml());
    }
    if (tableCellBorders != null &&
        tableCellBorders.isNotEmpty &&
        tableCellBorders.length <= TableCellBorderSide.values.length) {
      tcp.write('<w:tcBorders>');
      for (int i = 0; i < tableCellBorders.length; i++) {
        tcp.write(tableCellBorders[i].getXml());
      }
      tcp.write('</w:tcBorders>');
    }
    if (fitText ?? false) {
      tcp.write('<w:tcFitText/>');
    }

    if (tableCellMargins != null &&
        tableCellMargins.isNotEmpty &&
        tableCellMargins.length <= TableCellSide.values.length) {
      tcp.write('<w:tcMar>');
      for (int i = 0; i < tableCellMargins.length; i++) {
        tcp.write(tableCellMargins[i].getXml());
      }
      tcp.write('</w:tcMar>');
    }

    if (preferredWidth != null) {
      tcp.write(
          '<w:tcW w:type="${getValueFromEnum(preferredWidthType ?? PreferredWidthType.dxa)}" w:w="$preferredWidth"/>');
    }

    if (verticalAlignment != null) {
      tcp.write('<w:vAlign w:val="${getValueFromEnum(verticalAlignment)}"/>');
    }

    if (restartVMerge ?? false) {
      tcp.write('<w:vMerge w:val="restart"/>');
    } else if (continueVMerge ?? false) {
      tcp.write('<w:vMerge w:val="continue"/>');
    }

    tcp.write('</w:tcPr>');
    return tcp.toString();
  }
}
