import '../style_enums.dart';

class TableRowProperties {
  final bool cantSplit;
  final bool hidden;
  final TableTextAlignment textAlignment;
  final int tableCellSpacing;
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
}
