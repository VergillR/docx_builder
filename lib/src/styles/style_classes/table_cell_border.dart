import '../../utils/utils.dart';
import '../style.dart';
import '../style_enums.dart';

class TableCellBorder extends Style {
  final TableCellBorderSide borderSide;
  final int width;
  final int space;
  final String color;
  final ParagraphBorderStyle pbrStyle;
  final bool shadow;

  TableCellBorder({
    this.borderSide = TableCellBorderSide.bottom,
    this.width = 24,
    this.space = 1,
    this.color = "000000",
    this.pbrStyle = ParagraphBorderStyle.single,
    this.shadow = false,
  });

  @override
  String getXml() {
    final String borderColor = isValidColor(color) ? color : '000000';
    return '<w:${getValueFromEnum(borderSide)} w:val="${getValueFromEnum(pbrStyle)}" w:color="$borderColor" w:sz="$width" w:space="$space" w:shadow="$shadow" />';
  }
}
