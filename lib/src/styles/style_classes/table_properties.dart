/*
  DocxBuilder 2020. Vergill Lemmert (NovaPrisma).
  This Source Code Form is subject to the terms of the Mozilla Public
  License, v. 2.0. If a copy of the MPL was not distributed with this
  file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/

import '../../utils/utils.dart';
import '../style.dart';
import '../style_classes/index.dart';
import '../style_enums.dart';
import 'index.dart';

class TableProperties extends Style {
  final TableTextAlignment tableTextAlignment;
  final Shading shading;
  final List<TableBorder> tableBorders;
  final String tableCaption;
  final List<TableCellMargin> tableCellMargins;
  final String tableCellSpacing;
  final PreferredWidthType tableCellSpacingType;
  final String tableIndentation;
  final PreferredWidthType tableIndentationType;
  final bool tableLayoutUsesFixedWidth;
  final List<TableConditionalFormatting> tableLook;
  final bool allowFloatingTableOverlapping;
  final FloatingTable floatingTable;
  final String preferredTableWidth;
  final PreferredWidthType preferredTableWidthType;

  TableProperties({
    this.tableTextAlignment,
    this.shading,
    this.tableBorders,
    this.tableCaption,
    this.tableCellMargins,
    this.tableCellSpacing,
    this.tableCellSpacingType = PreferredWidthType.dxa,
    this.tableIndentation,
    this.tableIndentationType = PreferredWidthType.dxa,
    this.tableLayoutUsesFixedWidth,
    this.tableLook,
    this.allowFloatingTableOverlapping,
    this.floatingTable,
    this.preferredTableWidth,
    this.preferredTableWidthType = PreferredWidthType.dxa,
  });

  @override
  String getXml() {
    final StringBuffer tp = StringBuffer();
    tp.write('<w:tblPr>');
    if (tableTextAlignment != null) {
      tp.write('<w:jc w:val="${getValueFromEnum(tableTextAlignment)}"/>');
    }

    if (shading != null) {
      tp.write(shading.getXml());
    }

    if (tableBorders != null &&
        tableBorders.isNotEmpty &&
        tableBorders.length <= TableBorderSide.values.length) {
      tp.write('<w:tblBorders>');
      for (int i = 0; i < tableBorders.length; i++) {
        final TableBorder border = tableBorders[i];
        tp.write(border.getXml());
      }
      tp.write('</w:tblBorders>');
    }

    if (tableCaption != null) {
      tp.write('<w:tblCaption w:val="$tableCaption"/>');
    }

    if (tableCellMargins != null &&
        tableCellMargins.length <= TableCellSide.values.length) {
      tp.write('<w:tblCellMar>');
      for (int i = 0; i < tableCellMargins.length; i++) {
        final TableCellMargin target = tableCellMargins[i];
        tp.write(target.getXml());
      }
      tp.write('</w:tblCellMar>');
    }

    if (tableCellSpacing != null) {
      tp.write(
          '<w:tblCellSpacing w:w="$tableCellSpacing" w:type="${getValueFromEnum(tableCellSpacingType)}"/>');
    }

    if (tableIndentation != null) {
      tp.write(
          '<w:tblInd w:w="$tableIndentation" w:type="${getValueFromEnum(tableIndentationType)}"/>');
    }

    if (tableLayoutUsesFixedWidth ?? false) {
      tp.write('<w:tblLayout w:type="fixed"/>');
    }

    if (tableLook != null &&
        tableLook.isNotEmpty &&
        tableLook.length <= TableConditionalFormatting.values.length) {
      final StringBuffer conditions = StringBuffer();
      for (int i = 0; i < tableLook.length; i++) {
        conditions.write('w:${getValueFromEnum(tableLook[i])}="true" ');
      }
      tp.write('<w:tblLook ${conditions.toString()}/>');
    }

    if (allowFloatingTableOverlapping != null) {
      tp.write(
          '<w:tblOverlap val="${allowFloatingTableOverlapping ? "overlap" : "never"}"/>');
    }

    if (floatingTable != null) {
      tp.write(floatingTable.getXml());
    }

    if (preferredTableWidth != null) {
      tp.write(
          '<w:tblW w:w="$preferredTableWidth" w:type="${getValueFromEnum(preferredTableWidthType)}"/>');
    }

    tp.write('</w:tblPr>');
    return tp.toString();
  }
}
