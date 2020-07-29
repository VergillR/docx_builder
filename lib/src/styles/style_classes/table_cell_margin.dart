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
