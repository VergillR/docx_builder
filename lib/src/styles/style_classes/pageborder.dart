import '../style_enums.dart';

class PageBorder {
  final PageBorderSide pageBorderSide;
  final int size;
  final int space;
  final String color;
  final ParagraphBorderStyle pbrStyle;
  final bool shadow;

  PageBorder({
    this.pageBorderSide = PageBorderSide.bottom,
    this.size = 8,
    this.space = 24,
    this.color = '000000',
    this.pbrStyle = ParagraphBorderStyle.single,
    this.shadow = false,
  });
}
