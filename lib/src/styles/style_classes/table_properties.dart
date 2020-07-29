import '../../utils/utils.dart';
import '../style_classes/index.dart';
import '../style_enums.dart';
import 'index.dart';

class TableProperties {
  final TableTextAlignment tableTextAlignment;
  final Shading shading;
  final List<TableBorder> tableBorders;
  final String tableCaption;
  final List<TableCellMargin> tableCellMargins;
  final int tableCellSpacing;
  final int tableIndentation;
  final bool tableLayoutUsesFixedWidth;
  final List<TableConditionalFormatting> tableLook;
  final bool allowFloatingTableOverlapping;
  final FloatingTable floatingTable;
  final int preferredWidth;
  final PreferredWidthType preferredWidthType;

  TableProperties({
    this.tableTextAlignment,
    this.shading,
    this.tableBorders,
    this.tableCaption,
    this.tableCellMargins,
    this.tableCellSpacing,
    this.tableIndentation,
    this.tableLayoutUsesFixedWidth = false,
    this.tableLook,
    this.allowFloatingTableOverlapping,
    this.floatingTable,
    this.preferredWidth,
    this.preferredWidthType,
  });

  String getTableProperties() {
    final StringBuffer tp = StringBuffer();
    tp.write('<w:tblPr>');
    if (tableTextAlignment != null) {
      tp.write('<w:jc w:val="${getValueFromEnum(tableTextAlignment)}"/>');
    }
    if (shading != null) {
      final String shadingColor =
          isValidColor(shading.shadingColor) ? shading.shadingColor : 'FFFFFF';
      final String shadingPatternColor =
          isValidColor(shading.shadingPatternColor)
              ? shading.shadingPatternColor
              : 'FFFFFF';
      tp.write(
          '<w:shd w:val="${getValueFromEnum(shading.shadingPattern ?? ShadingPatternStyle.nil)}" w:fill="$shadingPatternColor" w:color="$shadingColor" />');
    }

    if (tableBorders != null &&
        tableBorders.isNotEmpty &&
        tableBorders.length < 7) {
      tp.write('<w:tblBorders>');
      for (int i = 0; i < tableBorders.length; i++) {
        final TableBorder border = tableBorders[i];
        final String borderColor =
            isValidColor(border.color) ? border.color : '000000';
        tp.write(
            '<w:${getValueFromEnum(border.borderSide)} w:val="${getValueFromEnum(border.pbrStyle)}" w:sz="${border.width}" w:space="${border.space}" w:color="$borderColor" w:shadow="${border.shadow}" />');
      }
      tp.write('</w:tblBorders>');
    }

    tp.write('</w:tblPr>');
    return tp.toString();
  }
}
