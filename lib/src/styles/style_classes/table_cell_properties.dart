import '../style_enums.dart';
import 'index.dart';

class TableCellProperties {
  final int gridSpan;
  final bool hideMark;
  final bool noWrap;
  final Shading shading;
  final List<TableCellBorder> tableCellBorders;
  final bool fitText;
  final List<TableCellMargin> tableCellMargins;
  final int preferredWidth;
  final PreferredWidthType preferredWidthType;
  final TableCellVerticalAlignment verticalAlignment;
  final bool restartVMerge;
  final int vMergeCellSpan;

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
    this.vMergeCellSpan,
  });
}
