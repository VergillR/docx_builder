import '../style_enums.dart';

class TableCellMargin {
  final TableCellSide tableCellSide;
  final int marginWidth;
  final bool isNil;

  TableCellMargin({
    this.tableCellSide = TableCellSide.bottom,
    this.marginWidth = 0,
    this.isNil = false,
  });
}
