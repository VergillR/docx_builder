import 'package:docx_builder/src/styles/style_enums.dart';

class TableCellBorder {
  final TableCellBorderSide borderSide;
  final int width;
  final int space;
  final String color;
  final ParagraphBorderStyle pbrStyle;
  final bool shadow;

  const TableCellBorder({
    this.borderSide = TableCellBorderSide.bottom,
    this.width = 24,
    this.space = 1,
    this.color = "000000",
    this.pbrStyle = ParagraphBorderStyle.single,
    this.shadow = false,
  });
}
