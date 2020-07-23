import 'package:docx_builder/src/styles/styles.dart';

class ParagraphBorder {
  final ParagraphBorderSide borderSide;
  final int width;
  final int space;
  final String color;
  final ParagraphBorderStyle pbrStyle;
  final bool shadow;

  const ParagraphBorder({
    this.borderSide = ParagraphBorderSide.bottom,
    this.width = 24,
    this.space = 1,
    this.color = "000000",
    this.pbrStyle = ParagraphBorderStyle.single,
    this.shadow = false,
  });
}
